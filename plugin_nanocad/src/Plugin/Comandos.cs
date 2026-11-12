// Plugin/Comandos.cs — comandos do plugin no nanoCAD. Tudo que mexe no desenho passa por aqui (contexto de comando).
//   ORCAMENTO      tela inicial: Construir do zero / Ler projeto / Continuar obra salva
//   ORC_LER        lê o desenho inteiro pelas regras (layer + cor)       ORC_RELER     o mesmo, do zero (descarta ajustes)
//   ORC_ATUALIZAR  relê mantendo os ajustes                              ORC_CONTINUAR abre a obra salva (DWG / .obra.json)
//   ORC_CONSTRUIR  obra vazia para montar à mão                          ORC_POSTE / ORC_VAO  apontar poste/vão + anotação
//   ORC_VINCULAR   refaz os vínculos automáticos                         ORC_EDITAR    pop-up do vão/poste (número ou clique)
//   ORC_VINCULOS   tabela de vínculos                                    ORC_ATIVOS    ativos com parâmetros
//   ORC_SALVAR     grava no DWG e em <desenho>.obra.json                 ORC_PAINEL    reabre o painel
//   ORC_MARCAS     liga/desliga números e vínculos na tela               ORC_ZOOM      localiza um elemento
//   ORC_REGRAS     recarrega as regras (pasta "regras" ao lado da DLL)   ORC_AJUDA     esta lista
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using HostMgd.ApplicationServices;
using HostMgd.EditorInput;
using OrcamentoNanoCAD.Nucleo;
using OrcamentoNanoCAD.UI;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;
using Application = HostMgd.ApplicationServices.Application;

namespace OrcamentoNanoCAD
{
    public class Comandos
    {
        // ── utilidades ──
        private static Sessao Sessao() { return Aplicacao.Sessao(Aplicacao.Doc(), true); }

        private static void Rodar(string nome, Action<Sessao> acao)
        {
            Sessao s = null;
            Aplicacao.Log("comando " + nome);
            try
            {
                s = Sessao();
                if (s == null) return;
                acao(s);
            }
            catch (System.Exception e)
            {
                Aplicacao.Log("[" + nome + "] " + e);
                Aplicacao.Msg("[" + nome + "] erro: " + e.Message + " (detalhes em " + Aplicacao.ArquivoLog + ")");
            }
        }

        private static bool PrecisaObra(Sessao s)
        {
            if (s.Obra != null) return true;
            Aplicacao.Msg("Nenhuma obra aberta neste desenho. Use ORCAMENTO (ou ORC_LER / ORC_CONSTRUIR / ORC_CONTINUAR).");
            return false;
        }

        private static bool Confirmar(Editor ed, string pergunta)
        {
            PromptKeywordOptions o = new PromptKeywordOptions("\n" + pergunta + " [Sim/Não]", "Sim Não");
            o.Keywords.Default = "Não";
            o.AllowNone = true;
            PromptResult r = ed.GetKeywords(o);
            return r.Status == PromptStatus.OK && r.StringResult == "Sim";
        }

        private static void Mudou(Sessao s, bool modificou)
        {
            if (modificou) s.Modificada = true;
            s.Revisao++;                      // o Sincronizador confere de novo com a base da aplicação
            s.Atualizar();
        }

        private static string Agora() { return DateTime.Now.ToString("dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture); }

        // ── início ──
        [CommandMethod("ORCAMENTO")]
        public void Orcamento()
        {
            Rodar("ORCAMENTO", s =>
            {
                PersistenciaDwg.ObraSalva salva = PersistenciaDwg.Procurar(s);
                using (TelaInicio t = new TelaInicio(s, salva))
                {
                    if (Application.ShowModalDialog(t) != DialogResult.OK || t.Escolha == Escolha.Nada) return;
                    if (t.Escolha == Escolha.Painel) { PainelObra.Abrir(s); return; }
                    if (t.Regras != null) s.Regras = RegrasFrescas(t.Regras, t.Projeto);
                    if (t.Escolha == Escolha.Leitura) { ConfigurarLeitura(s); return; }
                    if (t.Escolha == Escolha.Ler) Ler(s, false, t.Projeto);
                    else if (t.Escolha == Escolha.Construir) Construir(s, t.Projeto);
                    else if (t.Escolha == Escolha.Continuar) Continuar(s, salva);
                }
            });
        }

        /// <summary>Regras da aplicação escolhidas: baixa a versão atual (se conectado); sem conexão, fica a do cache.</summary>
        private static RegrasProjeto RegrasFrescas(RegrasProjeto escolhidas, string projeto)
        {
            if (!escolhidas.DaAplicacao) return escolhidas;
            string codigo = string.IsNullOrEmpty(projeto) ? escolhidas.Codigo : projeto;
            if (Conexao.Logado)
            {
                try
                {
                    RegrasProjeto r = Conexao.BaixarRegras(codigo);
                    Aplicacao.Msg("Regras do projeto " + r.Nome + " baixadas da aplicação" + (r.LeituraDaAplicacao ? " (com a configuração de leitura do projeto)." : "."));
                    return r;
                }
                catch (ErroAplicacao e) { Aplicacao.Msg("Não deu para baixar as regras do projeto " + codigo + ": " + e.Message); }
            }
            if (escolhidas.Processamento.Count > 0 || escolhidas.Classificacao.Count > 0)
            {
                Aplicacao.Msg("Usando a última cópia das regras da aplicação para " + escolhidas.Nome + ".");
                return escolhidas;
            }
            // projeto nunca baixado e sem a aplicação: regras embutidas (Pernambuco para 081, senão o padrão)
            RegrasProjeto emb = Aplicacao.Regras.FirstOrDefault(x => x.Codigo == codigo) ?? Aplicacao.Regras.FirstOrDefault();
            Aplicacao.Msg("Sem as regras da aplicação para " + codigo + ": usando as regras embutidas " + (emb != null ? emb.Nome : "") + ".");
            return emb ?? escolhidas;
        }

        private static void Ler(Sessao s, bool manterAjustes, string projeto)
        {
            Obra antiga = s.Obra;
            DesenhoBruto bruto = ExtratorDesenho.Ler(s.Doc);
            s.UltimoBruto = bruto;
            Desenho d = MontadorDesenho.Montar(bruto, s.Regras.Leitura);
            s.UltimoDesenho = d;
            Obra nova = LeitorProjeto.Ler(d, s.Regras);
            nova.Projeto = !string.IsNullOrEmpty(projeto) ? projeto : (antiga != null && !string.IsNullOrEmpty(antiga.Projeto) ? antiga.Projeto : s.Regras.Codigo);
            if (manterAjustes && antiga != null) { Construtor.ReaplicarAjustes(nova, antiga, s.Regras); nova.SalvoEm = antiga.SalvoEm; }
            s.Obra = nova;
            Aplicacao.Msg("ORC_LER " + Relatorios.Resumo(nova) + " (" + d.Textos.Count + " textos, " + d.Postes.Count + " postes no desenho, " + d.Linhas.Count
                          + " linhas; " + nova.Ignorados.Count + " textos ignorados pelas regras)");
            foreach (string sug in Diagnostico.SugestoesDeRede(d, nova, s.Regras.Leitura).Concat(Diagnostico.SugestoesDeBlocos(bruto, nova, s.Regras.Leitura)))
                Aplicacao.Msg("💡 " + sug);
            if (d.Postes.Count == 0 || nova.Vaos.Count == 0)
                Aplicacao.Msg("⚠ " + (d.Postes.Count == 0 ? "Nenhum bloco de poste achado. " : "") + (nova.Vaos.Count == 0 ? "Nenhuma rede achada. " : "")
                              + "Confira em ⚙ Leitura (ORC_LEITURA) as layers da rede e o nome do bloco de poste deste desenho; ORC_DIAGNOSTICO mostra o que foi lido.");
            foreach (string a in bruto.Avisos.Take(5)) Aplicacao.Log("leitura: " + a);
            Mudou(s, true);
            PainelObra.Abrir(s);
        }

        private static void Construir(Sessao s, string projeto)
        {
            s.Obra = Construtor.NovaObra(s.Doc.Name, s.Regras);
            if (!string.IsNullOrEmpty(projeto)) s.Obra.Projeto = projeto;
            Mudou(s, true);
            PainelObra.Abrir(s);
            Aplicacao.Msg("Construir do zero: use + Poste (ORC_POSTE) e + Vão (ORC_VAO) no painel.");
        }

        private static void Continuar(Sessao s, PersistenciaDwg.ObraSalva salva)
        {
            if (salva == null) { Aplicacao.Msg("Não há obra salva neste desenho."); return; }
            Obra o = ObraArquivo.Ler(salva.Json, s.Regras);
            RegrasProjeto r = Aplicacao.RegrasParaEscolher().FirstOrDefault(x => x.Nome == o.NomeRegras);
            if (r != null && r != s.Regras) { s.Regras = r; o = ObraArquivo.Ler(salva.Json, r); }
            s.Obra = o;
            s.Modificada = false;
            Aplicacao.Msg("Obra salva aberta (" + salva.Onde + "): " + Relatorios.Resumo(o));
            Mudou(s, false);
            PainelObra.Abrir(s);
        }

        [CommandMethod("ORC_LER")]
        public void CmdLer() { Rodar("ORC_LER", s => Ler(s, false, null)); }

        [CommandMethod("ORC_RELER")]
        public void CmdReler()
        {
            Rodar("ORC_RELER", s =>
            {
                if (s.Obra != null && s.Obra.Elementos.Any(e => e.OperacaoAjustada || e.Manual) && !Confirmar(s.Doc.Editor, "Reler do zero descarta os ajustes feitos. Continuar?")) return;
                Ler(s, false, null);
            });
        }

        [CommandMethod("ORC_ATUALIZAR")]
        public void CmdAtualizar() { Rodar("ORC_ATUALIZAR", s => Ler(s, true, null)); }

        [CommandMethod("ORC_CONTINUAR")]
        public void CmdContinuar() { Rodar("ORC_CONTINUAR", s => Continuar(s, PersistenciaDwg.Procurar(s))); }

        [CommandMethod("ORC_CONSTRUIR")]
        public void CmdConstruir() { Rodar("ORC_CONSTRUIR", s => Construir(s, null)); }

        [CommandMethod("ORC_PAINEL")]
        public void CmdPainel() { Rodar("ORC_PAINEL", s => { if (PrecisaObra(s)) PainelObra.Abrir(s); }); }

        // ── construir do zero ──
        /// <summary>Entidades → desenho montado com a configuração de leitura das regras (textos já com cor e parágrafos).</summary>
        private static Desenho Montar(Sessao s, IEnumerable<ObjectId> ids)
        {
            DesenhoBruto b = ExtratorDesenho.LerSelecao(s.Doc, ids);
            foreach (string a in b.Avisos) Aplicacao.Log("seleção: " + a);
            return MontadorDesenho.Montar(b, s.Regras.Leitura);
        }

        private static List<TextoDesenho> SelecionarTextos(Sessao s, string mensagem)
        {
            Editor ed = s.Doc.Editor;
            PromptSelectionOptions o = new PromptSelectionOptions { MessageForAdding = mensagem };
            // TEXT, MTEXT e linha de chamada múltipla (MLEADER = "MULTILEADER" no filtro)
            SelectionFilter f = new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "TEXT,MTEXT,MULTILEADER") });
            PromptSelectionResult r = ed.GetSelection(o, f);
            if (r.Status != PromptStatus.OK) return new List<TextoDesenho>();
            List<TextoDesenho> textos = Montar(s, r.Value.GetObjectIds()).Textos.Where(t => t.Tipo != "BLOCO" && t.Tipo != "BLOCO_DEF").ToList();
            Aplicacao.Msg(textos.Count == 0 ? "Nenhum texto selecionado." : "Anotação: " + string.Join(" / ", textos.Select(t => t.Texto).ToArray()));
            return textos;
        }

        private static bool PrepararConstrucao(Sessao s)
        {
            if (s.Obra == null) { s.Obra = Construtor.NovaObra(s.Doc.Name, s.Regras); PainelObra.Abrir(s); }
            else if (s.Obra.Modo != "construir") Aplicacao.Msg("Acrescentando à obra lida (os elementos novos ficam marcados como manuais).");
            return true;
        }

        [CommandMethod("ORC_POSTE")]
        public void CmdPoste()
        {
            Rodar("ORC_POSTE", s =>
            {
                PrepararConstrucao(s);
                Editor ed = s.Doc.Editor;
                while (true)
                {
                    // o poste pode ser bloco OU círculo (projetos sem padrão); clicar direto num texto = poste só com a anotação
                    PromptEntityOptions o = new PromptEntityOptions("\nSelecione o POSTE — bloco ou círculo (Enter = sem símbolo, Esc = terminar): ");
                    o.AllowNone = true;
                    PromptEntityResult r = ed.GetEntity(o);
                    if (r.Status == PromptStatus.Cancel || r.Status == PromptStatus.Error) break;
                    PosteDesenho bloco = null;
                    List<TextoDesenho> textos = new List<TextoDesenho>();
                    if (r.Status == PromptStatus.OK)
                    {
                        DesenhoBruto b = ExtratorDesenho.LerSelecao(s.Doc, new[] { r.ObjectId });
                        EntidadeBruta e = b.Entidades.FirstOrDefault();
                        if (e != null && (e.Tipo == "INSERT" || e.Tipo == "CIRCLE"))
                            bloco = new PosteDesenho { Handle = e.Handle, X = e.X, Y = e.Y, Bloco = e.Tipo == "CIRCLE" ? "(círculo)" : (e.Bloco ?? ""), Layer = e.Layer };
                        else if (e != null && (e.Tipo == "TEXT" || e.Tipo == "MTEXT" || e.Tipo == "MLEADER"))
                            textos = MontadorDesenho.Montar(b, s.Regras.Leitura).Textos;
                        else { Aplicacao.Msg("Isso não é bloco, círculo nem texto. Tente de novo."); continue; }
                    }
                    if (textos.Count == 0) textos = SelecionarTextos(s, "\nSelecione a(s) anotação(ões) do poste e Enter: ");
                    if (bloco == null && textos.Count == 0) break;
                    Poste p = Construtor.AdicionarPoste(s.Obra, bloco, textos, s.Regras);
                    Aplicacao.Msg(p.Numero + ": " + (p.TemAnotacao ? p.AtivoLido : "(sem anotação)") + " [" + p.Operacao + "]"
                                  + (p.Estado != Estados.Ok ? " — " + string.Join("; ", p.Mensagens.ToArray()) : ""));
                    Mudou(s, true);
                }
            });
        }

        private static int PedirPoste(Editor ed, string msg, Obra obra, Poste sugerido, out Poste escolhido)
        {
            escolhido = sugerido;
            PromptStringOptions o = new PromptStringOptions(msg + (sugerido != null ? " <" + sugerido.Numero + ">: " : " (Enter = sem poste): "));
            o.AllowSpaces = false;
            PromptResult r = ed.GetString(o);
            if (r.Status != PromptStatus.OK) return -1;
            string t = (r.StringResult ?? "").Trim().ToUpperInvariant();
            if (t.Length == 0) return 0;
            if (t == "-") { escolhido = null; return 0; }
            escolhido = obra.PostePorNumero(t.StartsWith("P") ? t : "P" + t);
            if (escolhido == null) { Aplicacao.Msg("Poste " + t + " não existe na obra."); return -1; }
            return 0;
        }

        [CommandMethod("ORC_VAO")]
        public void CmdVao()
        {
            Rodar("ORC_VAO", s =>
            {
                PrepararConstrucao(s);
                Editor ed = s.Doc.Editor;
                while (true)
                {
                    PromptEntityOptions o = new PromptEntityOptions("\nSelecione a LINHA do vão — linha ou polilinha (Esc/Enter = terminar): ");
                    o.AllowNone = true;
                    PromptEntityResult r = ed.GetEntity(o);
                    if (r.Status != PromptStatus.OK) break;
                    LinhaDesenho linha = Montar(s, new[] { r.ObjectId }).Linhas.FirstOrDefault(l => !l.Leader);
                    if (linha == null || linha.Pontos.Count < 2) { Aplicacao.Msg("Isso não é linha nem polilinha. Tente de novo."); continue; }
                    List<TextoDesenho> textos = SelecionarTextos(s, "\nSelecione a anotação do vão e Enter: ");
                    Poste a, b;
                    Construtor.SugerirPostes(s.Obra, linha.Pontos, 4 * s.Obra.AlturaTexto, out a, out b);
                    // o plugin sugere os postes das pontas; o usuário confirma (Enter) ou troca (número, ou "-" para nenhum)
                    if (PedirPoste(ed, "\nPoste da ponta A", s.Obra, a, out a) < 0) break;
                    if (PedirPoste(ed, "\nPoste da ponta B", s.Obra, b, out b) < 0) break;
                    Vao v = Construtor.AdicionarVao(s.Obra, linha, textos, a, b, s.Regras);
                    Aplicacao.Msg(v.Numero + ": " + (v.TemAnotacao ? v.Ativo : "(sem anotação)") + " " + (v.PosteA ?? "—") + " → " + (v.PosteB ?? "—"));
                    Mudou(s, true);
                }
            });
        }

        [CommandMethod("ORC_VINCULAR")]
        public void CmdVincular()
        {
            Rodar("ORC_VINCULAR", s =>
            {
                if (!PrecisaObra(s)) return;
                Construtor.Recalcular(s.Obra, s.Regras);
                Aplicacao.Msg("Vínculos refeitos (os manuais foram mantidos). " + Relatorios.Resumo(s.Obra));
                Mudou(s, true);
            });
        }

        // ── localizar / pop-up ──
        private static Elemento PedirElemento(Sessao s, string prompt)
        {
            Editor ed = s.Doc.Editor;
            PromptStringOptions o = new PromptStringOptions("\n" + prompt + " (P1, V2…) ou Enter para clicar no desenho: ");
            o.AllowSpaces = false;
            PromptResult r = ed.GetString(o);
            if (r.Status != PromptStatus.OK) return null;
            string n = (r.StringResult ?? "").Trim().ToUpperInvariant();
            if (n.StartsWith("H:"))
            {
                // linha de "Ignorados"/diagnóstico: um texto pelo handle
                string hh = n.Substring(2);
                Elemento porH = s.Obra.PorHandle(hh);
                if (porH != null) return porH;
                TextoIgnorado t = s.Obra.Ignorados.FirstOrDefault(x => string.Equals(x.Handle, hh, StringComparison.OrdinalIgnoreCase));
                TextoDesenho td = t == null && s.UltimoDesenho != null ? s.UltimoDesenho.Textos.FirstOrDefault(x => string.Equals(x.Handle, hh, StringComparison.OrdinalIgnoreCase)) : null;
                if (t != null) ZoomPontos(s, new List<Ponto> { new Ponto(t.X, t.Y) });
                else if (td != null) ZoomPontos(s, new List<Ponto> { new Ponto(td.X, td.Y) });
                else Aplicacao.Msg("Texto " + hh + " não encontrado.");
                return null;
            }
            if (n.Length > 0)
            {
                Elemento e = s.Obra.PorNumero(n);
                if (e == null) Aplicacao.Msg(n + " não existe na obra.");
                return e;
            }
            PromptEntityResult pe = ed.GetEntity("\nClique no poste, no vão ou na anotação: ");
            if (pe.Status != PromptStatus.OK) return null;
            string h;
            using (Transaction tr = s.Doc.Database.TransactionManager.StartTransaction())
            {
                h = tr.GetObject(pe.ObjectId, OpenMode.ForRead).Handle.ToString();
                tr.Commit();
            }
            // a mesma linha pode ter vários vãos: vale o trecho mais perto do clique
            Ponto clique = new Ponto(pe.PickedPoint.X, pe.PickedPoint.Y);
            Vao vao = s.Obra.Vaos.Where(v => v.LinhaHandle == h && v.Pontos.Count > 1)
                .OrderBy(v => Geometria.Projetar(clique, v.Pontos).Distancia).FirstOrDefault();
            Elemento el = (Elemento)vao ?? s.Obra.PorHandle(h);
            if (el == null) Aplicacao.Msg("Esse objeto não faz parte da obra.");
            return el;
        }

        private static void Zoom(Sessao s, Elemento e)
        {
            List<Ponto> pts = new List<Ponto>();
            Poste p = e as Poste;
            Vao v = e as Vao;
            if (p != null) pts.Add(p.Ponto);
            if (v != null) pts.AddRange(v.Pontos);
            if (e.TextoX.HasValue) pts.Add(new Ponto(e.TextoX.Value, e.TextoY.Value));
            if (e.GuiaUsada != null) pts.AddRange(e.GuiaUsada);
            if (pts.Count == 0) return;
            ZoomPontos(s, pts);
            if (s.MarcasVisiveis) { s.Marcas.Mostrar(s.Obra, s.Regras); s.Marcas.Destacar(s.Obra, e); }
        }

        private static void ZoomPontos(Sessao s, List<Ponto> pts)
        {
            double h = Math.Max(s.Obra != null ? s.Obra.AlturaTexto : 1, 1e-3);
            double minX = pts.Min(q => q.X), maxX = pts.Max(q => q.X), minY = pts.Min(q => q.Y), maxY = pts.Max(q => q.Y);
            double margem = h * 12;
            if (Compat.Tentar("zoom pela API", () => ZoomVista(s.Doc.Editor, minX - margem, minY - margem, maxX + margem, maxY + margem))) return;
            // sem a vista pela API (nanoCAD 5): o próprio comando ZOOM com janela
            System.Globalization.CultureInfo ci = System.Globalization.CultureInfo.InvariantCulture;
            s.Doc.SendStringToExecute("_.ZOOM _W " + (minX - margem).ToString(ci) + "," + (minY - margem).ToString(ci) + " "
                                      + (maxX + margem).ToString(ci) + "," + (maxY + margem).ToString(ci) + "\n", true, false, false);
        }

        private static void ZoomVista(Editor ed, double x1, double y1, double x2, double y2)
        {
            double minX = x1, minY = y1, maxX = x2, maxY = y2, margem = 0;
            ViewTableRecord vista = ed.GetCurrentView();
            double largura = maxX - minX + 2 * margem, altura = maxY - minY + 2 * margem;
            double razao = vista.Width / Math.Max(vista.Height, 1e-9);
            if (largura / altura > razao) altura = largura / razao; else largura = altura * razao;
            vista.CenterPoint = new Point2d((minX + maxX) / 2, (minY + maxY) / 2);
            vista.Width = largura;
            vista.Height = altura;
            ed.SetCurrentView(vista);
        }

        [CommandMethod("ORC_ZOOM")]
        public void CmdZoom()
        {
            Rodar("ORC_ZOOM", s =>
            {
                if (!PrecisaObra(s)) return;
                Elemento e = PedirElemento(s, "Localizar");
                if (e != null) Zoom(s, e);
            });
        }

        [CommandMethod("ORC_EDITAR")]
        public void CmdEditar()
        {
            Rodar("ORC_EDITAR", s =>
            {
                if (!PrecisaObra(s)) return;
                Elemento e = PedirElemento(s, "Ajustar");
                if (e == null) return;
                Zoom(s, e);
                using (TelaElemento t = new TelaElemento(s.Obra, e, s.Regras, s.Obra.Modo == "construir" || e.Manual))
                {
                    if (Application.ShowModalDialog(t) != DialogResult.OK || t.Acao == AcaoPopup.Nada) { s.Atualizar(); return; }
                    AplicarPopup(s, e, t);
                }
            });
        }

        private static void AplicarPopup(Sessao s, Elemento e, TelaElemento t)
        {
            Vao v = e as Vao;
            Poste p = e as Poste;
            if (t.Acao == AcaoPopup.Excluir)
            {
                bool manual = e.Manual;
                Construtor.Excluir(s.Obra, e, s.Regras);
                Aplicacao.Msg(e.Numero + " excluído da obra (o desenho não foi alterado)."
                              + (manual ? "" : " Fica em Ignorados: duplo clique → Restaurar à obra."));
                Mudou(s, true);
                return;
            }
            if (t.Acao == AcaoPopup.Restaurar)
            {
                Construtor.Restaurar(s.Obra, e, s.Regras);
                Aplicacao.Msg(e.Numero + " voltou para a obra.");
                Mudou(s, true);
                return;
            }
            if (t.Acao == AcaoPopup.VoltarAoLido)
            {
                string opAntes = e.Operacao;
                Construtor.VoltarAoLido(s.Obra, e, s.Regras);
                if (opAntes != e.Operacao) TrocarLayerDoTexto(s, e, true);
                Mudou(s, true);
                return;
            }
            string antes = e.Operacao;
            if (t.AtivoNovo != null)
            {
                Construtor.TrocarAtivo(s.Obra, e, t.AtivoNovo, s.Regras);
                Aplicacao.Msg(e.Numero + ": ativo corrigido para " + e.AtivoLido + (e.AtivoEditado ? " (o desenho continua com " + e.AtivoDoDesenho + ")" : ""));
            }
            if (v != null)
            {
                Construtor.AjustarVao(s.Obra, v, t.Operacao, t.Metros, s.Regras);
                if (t.VinculoA != -2) AjustarVinculo(s, v, true, t.VinculoA);
                if (t.VinculoB != -2) AjustarVinculo(s, v, false, t.VinculoB);
            }
            else if (p != null) Construtor.AjustarPoste(s.Obra, p, t.Operacao, s.Regras);
            else { e.Operacao = t.Operacao; Construtor.Recalcular(s.Obra, s.Regras); }
            if (t.AlterarLayer && antes != e.Operacao) TrocarLayerDoTexto(s, e, false);
            Mudou(s, true);
        }

        private static void AjustarVinculo(Sessao s, Vao v, bool pontaA, int k)
        {
            if (k < 0)
            {
                if (pontaA) v.VinculoA = null; else v.VinculoB = null;
                Construtor.Recalcular(s.Obra, s.Regras);
                return;
            }
            string motivo = Construtor.VincularManual(s.Obra, v, pontaA, k, s.Regras);
            if (motivo != null) Aplicacao.Msg("Vínculo não feito: " + motivo);
        }

        /// <summary>Muda a layer da anotação para a layer da operação (pelas regras). Com `paraALida`, volta à layer lida.</summary>
        private static void TrocarLayerDoTexto(Sessao s, Elemento e, bool paraALida)
        {
            if (string.IsNullOrEmpty(e.TextoHandle)) return;
            string layer = paraALida ? e.TextoLayer : s.Regras.LayerDaOperacao(e.Operacao);
            if (string.IsNullOrEmpty(layer)) { Aplicacao.Msg("As regras não têm layer para a operação " + e.Operacao + "; o texto ficou na layer atual."); return; }
            Database db = s.Doc.Database;
            using (DocumentLock trava = s.Doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                ObjectId id = db.GetObjectId(false, new Handle(Convert.ToInt64(e.TextoHandle, 16)), 0);
                if (id.IsNull || id.IsErased) { Aplicacao.Msg("A anotação de " + e.Numero + " não está mais no desenho."); return; }
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (!lt.Has(layer))
                {
                    lt.UpgradeOpen();
                    LayerTableRecord nova = new LayerTableRecord { Name = layer };
                    lt.Add(nova);
                    tr.AddNewlyCreatedDBObject(nova, true);
                }
                Entity ent = (Entity)tr.GetObject(id, OpenMode.ForWrite);
                string de = ent.Layer;
                ent.Layer = layer;
                tr.Commit();
                if (!paraALida) Aplicacao.Msg("Anotação de " + e.Numero + ": layer " + de + " → " + layer + ".");
            }
        }

        // ── tabelas ──
        private static TelaTabela _vinculos, _ativos;

        private static void AbrirTabela(ref TelaTabela tela, Func<TelaTabela> criar)
        {
            if (tela != null && !tela.IsDisposed) tela.Close();
            tela = criar();
            Visual.MostrarNaoModal(tela);
        }

        [CommandMethod("ORC_VINCULOS")]
        public void CmdVinculos()
        {
            Rodar("ORC_VINCULOS", s =>
            {
                if (!PrecisaObra(s)) return;
                AbrirTabela(ref _vinculos, () => new TelaTabela("🔗 Vínculos — vãos ⇄ postes ⇄ estruturas", Relatorios.ColunasVinculos,
                    () => Relatorios.Vinculos(s.Obra, s.Regras),
                    "Clique numa linha para localizar o vão; duplo clique abre o pop-up. ✋ = vínculo feito à mão. R e *R não entram no vínculo."));
            });
        }

        [CommandMethod("ORC_ATIVOS")]
        public void CmdAtivos()
        {
            Rodar("ORC_ATIVOS", s =>
            {
                if (!PrecisaObra(s)) return;
                AbrirTabela(ref _ativos, () => new TelaTabela("📦 Ativos gerados — com parâmetros preenchidos", Relatorios.ColunasAtivos,
                    () => Relatorios.Ativos(s.Obra),
                    "Amarelo = parâmetro pendente (\"?\") ou fora do padrão. O cálculo do orçamento continua na aplicação (Carregar Obra / modo autônomo)."));
            });
        }

        // ── salvar ──
        [CommandMethod("ORC_SALVAR")]
        public void CmdSalvar()
        {
            Rodar("ORC_SALVAR", s =>
            {
                if (!PrecisaObra(s)) return;
                s.Obra.SalvoEm = Agora();
                if (string.IsNullOrEmpty(s.Obra.Desenho) || s.DesenhoSalvo) s.Obra.Desenho = s.Doc.Name;
                string json = ObraArquivo.Gerar(s.Obra, DateTime.UtcNow);
                PersistenciaDwg.GravarNoDwg(s.Doc, json);
                string arq = s.ArquivoObra;
                if (arq != null)
                {
                    PersistenciaDwg.GravarArquivo(arq, json);
                    Aplicacao.Msg("Obra salva no DWG e em " + arq + " (salve o DWG para guardar a cópia interna).");
                }
                else Aplicacao.Msg("Obra salva no DWG. O desenho ainda não tem arquivo: salve o DWG e rode ORC_SALVAR de novo para gerar o .obra.json.");
                s.Modificada = false;
                s.Atualizar();
            });
        }

        // ── outros ──
        [CommandMethod("ORC_MARCAS")]
        public void CmdMarcas()
        {
            Rodar("ORC_MARCAS", s =>
            {
                s.MarcasVisiveis = !s.MarcasVisiveis;
                if (!s.MarcasVisiveis) s.Marcas.Limpar();
                s.Atualizar();
                Aplicacao.Msg("Números e vínculos na tela: " + (s.MarcasVisiveis ? "ligados" : "desligados") + ".");
            });
        }

        [CommandMethod("ORC_REGRAS")]
        public void CmdRegras()
        {
            Rodar("ORC_REGRAS", s =>
            {
                Aplicacao.RecarregarRegras();
                foreach (RegrasProjeto r in Aplicacao.Regras)
                    Aplicacao.Msg("  " + r.Nome + " [" + (r.Codigo.Length > 0 ? r.Codigo : "sem código") + "] — " + r.Processamento.Count + "+" + r.Classificacao.Count
                                  + " regras do leitor, " + r.Vinculo.Cabos.Count + "+" + r.Vinculo.Estruturas.Count + " de vínculo (" + r.Origem + ")");
                RegrasProjeto atual = Aplicacao.Regras.FirstOrDefault(x => s.Regras != null && x.Nome == s.Regras.Nome);
                s.Regras = atual ?? Aplicacao.Regras.FirstOrDefault();
                Aplicacao.Msg("Pasta para regras próprias: " + Path.Combine(Aplicacao.PastaDaDll, "regras", "<codigo>") + " (processamento.json, classificacao.json, vinculacao.json, nome.txt)");
            });
        }

        // ── ligação com a aplicação ──
        [CommandMethod("ORC_CONECTAR")]
        public void CmdConectar()
        {
            Rodar("ORC_CONECTAR", s =>
            {
                using (TelaConexao t = new TelaConexao()) Application.ShowModalDialog(t);
                Sincronizador.Agora();
                UI.PainelObra.AtualizarStatusConexao();
            });
        }

        // ── TASK-017: linha nova sem apontar no desenho ──
        [CommandMethod("ORC_NOVO_CABO")]
        public void CmdNovoCabo() { Rodar("ORC_NOVO_CABO", s => NovaLinha(s, true)); }

        [CommandMethod("ORC_NOVO_OUTROS")]
        public void CmdNovoOutros() { Rodar("ORC_NOVO_OUTROS", s => NovaLinha(s, false)); }

        private static void NovaLinha(Sessao s, bool cabo)
        {
            if (s.Obra == null) PrepararConstrucao(s);
            using (TelaNovaLinha t = new TelaNovaLinha(s.Obra.Projeto, s.Regras, cabo))
            {
                if (Application.ShowModalDialog(t) != DialogResult.OK) return;
                Avulso a = Construtor.NovaLinha(s.Obra, cabo, t.Ativo, t.Operacao, s.Regras);
                Aplicacao.Msg(a.Numero + " adicionado à Tabela de " + (cabo ? "Cabos" : "Outros") + ": " + a.AtivoLido + " [" + a.Operacao + "]");
                Mudou(s, true);
            }
        }

        [CommandMethod("ORC_ATUALIZAR_PLUGIN")]
        public void CmdAtualizarPlugin()
        {
            Rodar("ORC_ATUALIZAR_PLUGIN", s => AtualizacaoPlugin.Conferir(true));
        }

        [CommandMethod("ORC_VOLTAR_VERSAO")]
        public void CmdVoltarVersao()
        {
            Rodar("ORC_VOLTAR_VERSAO", s => AtualizacaoPlugin.Voltar());
        }

        [CommandMethod("ORC_SINCRONIZAR")]
        public void CmdSincronizar()
        {
            Rodar("ORC_SINCRONIZAR", s =>
            {
                if (!Conexao.Logado) { Aplicacao.Msg("Sem ligação com a aplicação. Use ORC_CONECTAR."); return; }
                string projeto = s.Obra != null && !string.IsNullOrEmpty(s.Obra.Projeto) ? s.Obra.Projeto : (s.Regras != null ? s.Regras.Codigo : "");
                if (string.IsNullOrEmpty(projeto)) { Aplicacao.Msg("Informe o código do projeto (ORCAMENTO)."); return; }
                try
                {
                    RegrasProjeto novas = Conexao.BaixarRegras(projeto);
                    if (s.Obra != null) Sincronizador.TrocarRegras(s, novas); else s.Regras = novas;
                    Aplicacao.Msg("Regras do projeto " + projeto + " baixadas da aplicação: " + novas.Processamento.Count + "+" + novas.Classificacao.Count
                                  + " do leitor, " + novas.Vinculo.Cabos.Count + "+" + novas.Vinculo.Estruturas.Count + " de vínculo.");
                    if (s.Obra != null)
                    {
                        ConferenciaBase c = Conexao.Cliente.Conferir(projeto, ObraArquivo.Dados(s.Obra));
                        c.Aplicar(s.Obra, s.Regras, c.Versao.ProjetoNome);
                        s.RevisaoConferida = s.Revisao;
                        s.VersaoConferida = c.Versao;
                        s.UltimaConferencia = DateTime.Now;
                        s.ConferenciaPendencias = c.ComPendencia;
                        Aplicacao.Msg(c.NaoEncontrados.Count == 0 ? "Base de orçamento: todos os ativos da obra estão cadastrados."
                            : "Fora da base de orçamento (" + c.NaoEncontrados.Count + "): " + string.Join(", ", c.NaoEncontrados.Take(30).ToArray())
                              + (c.NaoEncontrados.Count > 30 ? "…" : "") + ". Veja Conferir no painel.");
                        s.Atualizar();
                    }
                }
                catch (ErroAplicacao e) { Aplicacao.Msg("Não sincronizou: " + e.Message); }
                Sincronizador.Agora();
            });
        }

        // ── navegação, leitura, diagnóstico, abrir na aplicação (TASK-004/006/008/010) ──
        [CommandMethod("ORC_FECHAR")]
        public void CmdFechar()
        {
            Rodar("ORC_FECHAR", s =>
            {
                if (s.Obra == null) { PainelObra.Esconder(); return; }
                if (s.Modificada && !Confirmar(s.Doc.Editor, "A obra tem mudanças não salvas. Fechar mesmo assim?")) return;
                s.Obra = null;
                s.Modificada = false;
                s.Marcas.Limpar();
                PainelObra.Esconder();
                Aplicacao.Msg("Obra fechada (o desenho não foi alterado). ORCAMENTO abre a tela inicial.");
            });
        }

        private static void ListasDoDesenho(Sessao s, out List<string> layers, out List<string> blocos)
        {
            DesenhoBruto b = s.UltimoBruto;
            if (b == null) { try { b = ExtratorDesenho.Ler(s.Doc); s.UltimoBruto = b; } catch (System.Exception) { b = new DesenhoBruto(); } }
            layers = b.Layers.Keys.ToList();
            blocos = b.Entidades.Where(e => e.Tipo == "INSERT").Select(e => e.Bloco ?? "").Distinct().ToList();
        }

        private static void ConfigurarLeitura(Sessao s)
        {
            List<string> layers, blocos;
            ListasDoDesenho(s, out layers, out blocos);
            using (TelaLeitura t = new TelaLeitura(s.Regras, layers, blocos))
            {
                if (Application.ShowModalDialog(t) != DialogResult.OK || t.Resultado == null) return;
                s.Regras.Leitura = t.Resultado;
                ConfigsLocais.Salvar(s.Regras.Codigo, t.Resultado);         // cópia neste computador (vale sem a aplicação)
                string onde = "neste computador";
                if (s.Regras.DaAplicacao && Conexao.Logado)
                {
                    try
                    {
                        string aviso = Conexao.Cliente.SalvarLeitura(s.Regras.Codigo, t.Resultado);
                        s.Regras.LeituraDaAplicacao = true;
                        onde = "na aplicação (vale em todos os computadores)" + (string.IsNullOrEmpty(aviso) ? "" : " — " + aviso);
                    }
                    catch (ErroAplicacao e) { onde = "só neste computador (a aplicação recusou: " + e.Message + ")"; }
                }
                Aplicacao.Msg("Configuração de leitura de " + s.Regras.Nome + " salva " + onde + ".");
            }
            if (s.Obra != null && s.Obra.Modo == "ler") Ler(s, true, null);
            else if (s.Obra == null && Confirmar(s.Doc.Editor, "Ler o projeto agora com essa configuração?")) Ler(s, false, null);
        }

        [CommandMethod("ORC_LEITURA")]
        public void CmdLeitura() { Rodar("ORC_LEITURA", ConfigurarLeitura); }

        private static TelaTabela _diag;

        [CommandMethod("ORC_DIAGNOSTICO")]
        public void CmdDiagnostico()
        {
            Rodar("ORC_DIAGNOSTICO", s =>
            {
                DesenhoBruto b = ExtratorDesenho.Ler(s.Doc);
                s.UltimoBruto = b;
                Desenho d = MontadorDesenho.Montar(b, s.Regras.Leitura);
                Obra o = LeitorProjeto.Ler(d, s.Regras);
                JObj diag = Diagnostico.Gerar(b, d, o, s.Regras, DateTime.Now);
                string arq = s.DesenhoSalvo ? Path.ChangeExtension(s.Doc.Name, null) + ".diagnostico.json" : Path.Combine(Conexao.Pasta, "diagnostico.json");
                Directory.CreateDirectory(Path.GetDirectoryName(arq));
                File.WriteAllText(arq, Json.Escrever(diag, true), new System.Text.UTF8Encoding(false));
                foreach (string l in Diagnostico.Linhas(diag)) Aplicacao.Msg(l);
                Aplicacao.Msg("Diagnóstico gravado em " + arq + " — mande esse arquivo para conferir a leitura deste desenho.");
                AbrirTabela(ref _diag, () => new TelaTabela("🩺 Diagnóstico — o que o plugin leu em cada texto", Relatorios.ColunasDiagnostico,
                    () => Relatorios.TextosDoDiagnostico(diag),
                    string.Join("  ·  ", Diagnostico.Linhas(diag).ToArray()) + "\nArquivo: " + arq));
            });
        }

        [CommandMethod("ORC_ABRIR_APP")]
        public void CmdAbrirApp()
        {
            Rodar("ORC_ABRIR_APP", s =>
            {
                if (!PrecisaObra(s)) return;
                if (string.IsNullOrEmpty(s.Obra.Projeto)) { Aplicacao.Msg("A obra não tem o código do projeto (ORCAMENTO → Projeto na aplicação)."); return; }
                string chaveObra = string.IsNullOrEmpty(s.Doc.Name) ? s.Obra.Nome : s.Doc.Name;
                // TASK-016: o GridReader instalado neste computador (aberto: entrega a obra; fechado: abre já com ela)
                if (ProgramaNoComputador.Abrir(chaveObra, ObraArquivo.Envelope(s.Obra, DateTime.UtcNow))) return;
                if (!Conexao.Logado)
                {
                    using (TelaConexao tc = new TelaConexao()) Application.ShowModalDialog(tc);
                    if (!Conexao.Logado) { Aplicacao.Msg("Sem ligação com a aplicação: conecte em 🔗 Aplicação e tente de novo."); return; }
                    Sincronizador.Agora();
                }
                try
                {
                    string chave = string.IsNullOrEmpty(s.Doc.Name) ? s.Obra.Nome : s.Doc.Name;
                    string url = Conexao.Cliente.EnviarObra(chave, ObraArquivo.Envelope(s.Obra, DateTime.UtcNow));
                    Aplicacao.Msg("GridReader não está instalado neste computador: obra enviada para " + Conexao.Dados.Base + "; abrindo no navegador.");
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(url) { UseShellExecute = true });
                }
                catch (ErroAplicacao e) { Aplicacao.Msg("Não abriu na aplicação: " + e.Message); }
            });
        }

        [CommandMethod("ORC_AJUDA")]
        public void CmdAjuda()
        {
            Aplicacao.Msg("GridReader para nanoCAD " + ObraArquivo.VersaoPlugin + " — desenvolvido por Gabriel Sales — comandos:\n"
                + "  ORCAMENTO      início (Construir do zero / Ler projeto / Continuar obra salva)\n"
                + "  ORC_LER  ORC_RELER  ORC_ATUALIZAR  ORC_CONTINUAR  ORC_CONSTRUIR\n"
                + "  ORC_POSTE  ORC_VAO  ORC_VINCULAR  ORC_EDITAR  ORC_ZOOM\n"
                + "  ORC_VINCULOS  ORC_ATIVOS  ORC_SALVAR  ORC_PAINEL  ORC_MARCAS  ORC_REGRAS\n"
                + "  ORC_CONECTAR   ligar à aplicação (localhost/online): regras e base de orçamento ao vivo\n"
                + "  ORC_SINCRONIZAR baixar agora as regras da aplicação e conferir a obra com a base\n"
                + "  ORC_NOVO_CABO  ORC_NOVO_OUTROS  linha nova sem apontar no desenho (com o autocompletar da base)\n"
                + "  ORC_ABRIR_APP  abrir a obra no GridReader deste computador (aberto ou fechado); sem ele, na aplicação do ORC_CONECTAR\n"
                + "  ORC_LEITURA    configurar layers da rede, bloco de poste e guias · ORC_DIAGNOSTICO o que foi lido\n"
                + "  ORC_FECHAR     fechar a obra · log: " + Aplicacao.ArquivoLog + "\n"
                + "  ORC_ATUALIZAR_PLUGIN conferir agora se há versão nova · ORC_VOLTAR_VERSAO voltar para a anterior\n"
                + "  Plugin " + ObraArquivo.VersaoPlugin + ": " + AtualizacaoPlugin.Status + "\n"
                + "  " + Sincronizador.Status);
        }
    }
}
