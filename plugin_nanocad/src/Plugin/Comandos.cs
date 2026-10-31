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
            try
            {
                s = Sessao();
                if (s == null) return;
                acao(s);
            }
            catch (System.Exception e)
            {
                Aplicacao.Msg("[" + nome + "] erro: " + e.Message);
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
                    if (t.Regras != null) s.Regras = RegrasFrescas(t.Regras, t.Projeto);
                    if (t.Escolha == Escolha.Ler) Ler(s, false, t.Projeto);
                    else if (t.Escolha == Escolha.Construir) Construir(s, t.Projeto);
                    else if (t.Escolha == Escolha.Continuar) Continuar(s, salva);
                }
            });
        }

        /// <summary>Regras da aplicação escolhidas: baixa a versão atual (se conectado); sem conexão, fica a do cache.</summary>
        private static RegrasProjeto RegrasFrescas(RegrasProjeto escolhidas, string projeto)
        {
            if (!escolhidas.DaAplicacao || !Conexao.Logado) return escolhidas;
            try { return Conexao.BaixarRegras(string.IsNullOrEmpty(projeto) ? escolhidas.Codigo : projeto); }
            catch (ErroAplicacao e) { Aplicacao.Msg("Usando as regras guardadas da aplicação (" + e.Message + ")."); return escolhidas; }
        }

        private static void Ler(Sessao s, bool manterAjustes, string projeto)
        {
            Obra antiga = s.Obra;
            Desenho d = ExtratorDesenho.Ler(s.Doc, s.Regras);
            s.UltimoDesenho = d;
            Obra nova = LeitorProjeto.Ler(d, s.Regras);
            nova.Projeto = !string.IsNullOrEmpty(projeto) ? projeto : (antiga != null && !string.IsNullOrEmpty(antiga.Projeto) ? antiga.Projeto : s.Regras.Codigo);
            if (manterAjustes && antiga != null) { Construtor.ReaplicarAjustes(nova, antiga, s.Regras); nova.SalvoEm = antiga.SalvoEm; }
            s.Obra = nova;
            Aplicacao.Msg("ORC_LER " + Relatorios.Resumo(nova) + " (" + d.Textos.Count + " textos, " + d.Postes.Count + " blocos de poste, " + d.Linhas.Count + " linhas)");
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
        private static List<TextoDesenho> SelecionarTextos(Sessao s, string mensagem)
        {
            Editor ed = s.Doc.Editor;
            PromptSelectionOptions o = new PromptSelectionOptions { MessageForAdding = mensagem };
            SelectionFilter f = new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "TEXT,MTEXT") });
            PromptSelectionResult r = ed.GetSelection(o, f);
            List<TextoDesenho> textos = new List<TextoDesenho>();
            if (r.Status != PromptStatus.OK) return textos;
            using (Transaction tr = s.Doc.Database.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in r.Value.GetObjectIds())
                {
                    Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                    if (ent != null) textos.AddRange(ExtratorDesenho.Textos(ent, tr, null));
                }
                tr.Commit();
            }
            return textos;
        }

        private static bool PrepararConstrucao(Sessao s)
        {
            if (s.Obra == null) { s.Obra = Construtor.NovaObra(s.Doc.Name, s.Regras); PainelObra.Abrir(s); }
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
                    PromptEntityOptions o = new PromptEntityOptions("\nSelecione o BLOCO do poste (Enter = sem bloco, Esc = terminar): ");
                    o.SetRejectMessage("\nPrecisa ser um bloco.");
                    o.AddAllowedClass(typeof(BlockReference), true);
                    o.AllowNone = true;
                    PromptEntityResult r = ed.GetEntity(o);
                    if (r.Status == PromptStatus.Cancel || r.Status == PromptStatus.Error) break;
                    PosteDesenho bloco = null;
                    if (r.Status == PromptStatus.OK)
                        using (Transaction tr = s.Doc.Database.TransactionManager.StartTransaction())
                        {
                            BlockReference br = (BlockReference)tr.GetObject(r.ObjectId, OpenMode.ForRead);
                            bloco = ExtratorDesenho.Poste(br, tr);
                            tr.Commit();
                        }
                    List<TextoDesenho> textos = SelecionarTextos(s, "\nSelecione a(s) anotação(ões) do poste e Enter: ");
                    if (bloco == null && textos.Count == 0) break;
                    Poste p = Construtor.AdicionarPoste(s.Obra, bloco, textos, s.Regras);
                    Aplicacao.Msg(p.Numero + ": " + (p.TemAnotacao ? p.AtivoLido : "(sem anotação)") + " [" + p.Operacao + "]");
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
                    PromptEntityOptions o = new PromptEntityOptions("\nSelecione a LINHA do vão (Esc/Enter = terminar): ");
                    o.SetRejectMessage("\nPrecisa ser linha ou polilinha.");
                    o.AddAllowedClass(typeof(Line), true);
                    o.AddAllowedClass(typeof(Polyline), true);
                    o.AddAllowedClass(typeof(Polyline2d), true);
                    o.AllowNone = true;
                    PromptEntityResult r = ed.GetEntity(o);
                    if (r.Status != PromptStatus.OK) break;
                    LinhaDesenho linha;
                    using (Transaction tr = s.Doc.Database.TransactionManager.StartTransaction())
                    {
                        linha = ExtratorDesenho.Linha((Entity)tr.GetObject(r.ObjectId, OpenMode.ForRead), tr);
                        tr.Commit();
                    }
                    if (linha == null || linha.Pontos.Count < 2) continue;
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
            double h = Math.Max(s.Obra.AlturaTexto, 1e-3);
            List<Ponto> pts = new List<Ponto>();
            Poste p = e as Poste;
            Vao v = e as Vao;
            if (p != null) pts.Add(p.Ponto);
            if (v != null) pts.AddRange(v.Pontos);
            if (e.TextoX.HasValue) pts.Add(new Ponto(e.TextoX.Value, e.TextoY.Value));
            if (pts.Count == 0) return;
            double minX = pts.Min(q => q.X), maxX = pts.Max(q => q.X), minY = pts.Min(q => q.Y), maxY = pts.Max(q => q.Y);
            double margem = h * 12;
            Editor ed = s.Doc.Editor;
            ViewTableRecord vista = ed.GetCurrentView();
            double largura = maxX - minX + 2 * margem, altura = maxY - minY + 2 * margem;
            double razao = vista.Width / Math.Max(vista.Height, 1e-9);
            if (largura / altura > razao) altura = largura / razao; else largura = altura * razao;
            vista.CenterPoint = new Point2d((minX + maxX) / 2, (minY + maxY) / 2);
            vista.Width = largura;
            vista.Height = altura;
            ed.SetCurrentView(vista);
            if (s.MarcasVisiveis) { s.Marcas.Mostrar(s.Obra, s.Regras); s.Marcas.Destacar(s.Obra, e); }
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
                Construtor.Remover(s.Obra, e, s.Regras);
                Aplicacao.Msg(e.Numero + " excluído da obra (o desenho não foi alterado).");
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
            Application.ShowModelessDialog(tela);
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

        [CommandMethod("ORC_AJUDA")]
        public void CmdAjuda()
        {
            Aplicacao.Msg("Orçamento (Leitor de Projetos) " + ObraArquivo.VersaoPlugin + " — comandos:\n"
                + "  ORCAMENTO      início (Construir do zero / Ler projeto / Continuar obra salva)\n"
                + "  ORC_LER  ORC_RELER  ORC_ATUALIZAR  ORC_CONTINUAR  ORC_CONSTRUIR\n"
                + "  ORC_POSTE  ORC_VAO  ORC_VINCULAR  ORC_EDITAR  ORC_ZOOM\n"
                + "  ORC_VINCULOS  ORC_ATIVOS  ORC_SALVAR  ORC_PAINEL  ORC_MARCAS  ORC_REGRAS\n"
                + "  ORC_CONECTAR   ligar à aplicação (localhost/online): regras e base de orçamento ao vivo\n"
                + "  ORC_SINCRONIZAR baixar agora as regras da aplicação e conferir a obra com a base\n"
                + "  " + Sincronizador.Status);
        }
    }
}
