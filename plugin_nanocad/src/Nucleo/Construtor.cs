// Nucleo/Construtor.cs — "Construir do zero" (o usuário aponta poste + anotação e vão + anotação) e os ajustes do pop-up
// (operação, metros, voltar ao lido) que sobrevivem a um "Atualizar" (reler mantendo ajustes). "Reler" descarta tudo.
using System;
using System.Collections.Generic;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    public static class Construtor
    {
        public static Obra NovaObra(string arquivo, RegrasProjeto regras)
        {
            Obra o = new Obra();
            o.Modo = "construir";
            o.Desenho = arquivo ?? "";
            o.Nome = LeitorProjeto.NomeDaObra(arquivo);
            o.Projeto = regras.Codigo;
            o.NomeRegras = regras.Nome;
            return o;
        }

        private static int Proximo(IEnumerable<Elemento> l, char prefixo)
        {
            int max = 0, n;
            foreach (Elemento e in l)
                if (e.Numero != null && e.Numero.Length > 1 && e.Numero[0] == prefixo && int.TryParse(e.Numero.Substring(1), out n)) max = Math.Max(max, n);
            return max + 1;
        }

        /// <summary>Lê os textos escolhidos com o motor do leitor. Sem regra que entenda (projeto sem padrão): o próprio texto, com a entidade pedida.</summary>
        public static TextoLido LerAnotacao(List<TextoDesenho> textos, RegrasProjeto regras, string entidadePadrao)
        {
            return LerAnotacao(textos, regras, entidadePadrao, null);
        }

        /// <summary>TASK-021: com entidade CABO, lê como cabo e completa o comprimento (mesma linha, escolhidos ou vizinhança do desenho).</summary>
        public static TextoLido LerAnotacao(List<TextoDesenho> textos, RegrasProjeto regras, string entidadePadrao, Func<IEnumerable<TextoDesenho>> desenho)
        {
            if (textos == null || textos.Count == 0) return null;
            if (entidadePadrao == "CABO")
            {
                LeituraDeTexto c = LerCaboDeTexto(textos, regras, desenho);
                TextoLido tc = new TextoLido();
                tc.T = c.Texto ?? textos[0];
                tc.R = new ResultadoLeitor();
                tc.R.Entidade = "CABO";
                tc.R.Ativo = c.Ativo;
                tc.R.Operacao = c.Operacao;
                return tc;
            }
            List<TextoLido> lidos = LeitorProjeto.ClassificarTextos(textos, regras);
            TextoLido bom = lidos.FirstOrDefault(t => t.R.Entidade == entidadePadrao)
                ?? lidos.FirstOrDefault(t => entidadePadrao == "POSTE" && t.R.Entidade == "ESTRUTURA");
            if (bom != null) return bom;
            TextoLido primeiro = lidos[0];
            TextoLido r = new TextoLido();
            r.T = primeiro.T;
            r.R = new ResultadoLeitor();
            r.R.Entidade = entidadePadrao;
            r.R.Ativo = AtivoDosTextos(lidos);
            r.R.Operacao = primeiro.R.Operacao == "0" || string.IsNullOrEmpty(primeiro.R.Operacao) ? "I" : primeiro.R.Operacao;
            return r;
        }

        /// <summary>TASK-019: o ativo dos textos escolhidos pelo PROCESSAMENTO das regras do leitor, mesmo quando nenhuma regra de classificação
        /// os reconhece (ex.: layer que não é de operação). Assim o comprimento da linha de baixo vai para a linha do cabo ("3XA35P… 35 m",
        /// como a leitura de Pernambuco faz) e a linha só de comprimento, já usada, não entra separada por vírgula. Textos diferentes que o
        /// processamento não juntou continuam separados por ", " (vários itens de Outros).</summary>
        public static string AtivoDosTextos(List<TextoLido> lidos)
        {
            List<string> partes = new List<string>();
            foreach (TextoLido t in lidos)
            {
                string a = (t.R != null ? t.R.Ativo : null) ?? "";
                if (a.Trim().Length == 0) continue;                 // usado por outra linha (ex.: comprimento) ou descartado pelas regras
                if (!partes.Contains(a)) partes.Add(a);
            }
            if (partes.Count == 0) partes = lidos.Select(t => t.T.Texto).Where(x => !string.IsNullOrEmpty(x)).ToList();
            return MotorLeitor.LimparTexto(string.Join(", ", partes.ToArray()));
        }

        public static Poste AdicionarPoste(Obra obra, PosteDesenho bloco, List<TextoDesenho> textos, RegrasProjeto regras)
        {
            Poste p = bloco != null && !string.IsNullOrEmpty(bloco.Handle) ? obra.Postes.FirstOrDefault(x => x.Handle == bloco.Handle) : null;
            if (p == null)
            {
                p = new Poste();
                p.Numero = "P" + Proximo(obra.Postes.Cast<Elemento>(), 'P');
                obra.Postes.Add(p);
            }
            p.Manual = true;
            if (bloco != null) { p.Handle = bloco.Handle; p.X = bloco.X; p.Y = bloco.Y; p.Bloco = bloco.Bloco ?? ""; }
            TextoLido tl = LerAnotacao(textos, regras, "POSTE");
            if (tl != null && tl.T.Altura > 0 && obra.Postes.Count == 1) obra.AlturaTexto = tl.T.Altura;
            if (tl != null)
            {
                LeitorProjeto.Anotar(p, tl);
                if (bloco == null) { p.X = tl.T.X; p.Y = tl.T.Y; p.Handle = ""; }
            }
            Recalcular(obra, regras);
            return p;
        }

        /// <summary>Postes da obra mais próximos das pontas da linha (até `tol`), como sugestão para o vão.</summary>
        public static void SugerirPostes(Obra obra, List<Ponto> pts, double tol, out Poste a, out Poste b)
        {
            a = MaisPerto(obra, pts[0], tol, null);
            b = MaisPerto(obra, pts[pts.Count - 1], tol, a);
        }

        private static Poste MaisPerto(Obra obra, Ponto q, double tol, Poste exceto)
        {
            Poste melhor = null;
            double md = double.MaxValue;
            foreach (Poste p in obra.Postes)
            {
                if (p == exceto) continue;
                double d = q.Dist(p.Ponto);
                if (d <= tol && d < md) { md = d; melhor = p; }
            }
            return melhor;
        }

        public static Vao AdicionarVao(Obra obra, LinhaDesenho linha, List<TextoDesenho> textos, Poste a, Poste b, RegrasProjeto regras)
        {
            return AdicionarVao(obra, linha, textos, a, b, regras, null);
        }

        public static Vao AdicionarVao(Obra obra, LinhaDesenho linha, List<TextoDesenho> textos, Poste a, Poste b, RegrasProjeto regras, Func<IEnumerable<TextoDesenho>> desenho)
        {
            return AdicionarVaos(obra, linha, textos, a, b, regras, desenho)[0];
        }

        public static List<Vao> AdicionarVaos(Obra obra, LinhaDesenho linha, List<TextoDesenho> textos, Poste a, Poste b, RegrasProjeto regras, Func<IEnumerable<TextoDesenho>> desenho)
        {
            List<LeituraDeTexto> cabos = LerCabosDeTextos(textos, regras, desenho);
            var grupo = obra.Vaos.Where(v => v.LinhaHandle == linha.Handle && v.Trecho == 0).ToList();
            Vao raiz = grupo.FirstOrDefault(v => v.IrmaoDe == null);
            var incluidos = new List<Vao>();
            foreach (LeituraDeTexto cabo in cabos.Count == 0 ? new List<LeituraDeTexto> { null } : cabos)
            {
                Vao vao = cabo == null ? raiz : grupo.FirstOrDefault(v => !string.IsNullOrEmpty(cabo.Texto.Handle) &&
                    v.TextoHandle == cabo.Texto.Handle && v.ParteTexto == cabo.Texto.Parte);
                if (vao == null && raiz != null && (!raiz.TemAnotacao || (cabos.Count == 1 && grupo.Count == 1))) vao = raiz;
                if (vao == null)
                {
                    vao = new Vao { Numero = "V" + Proximo(obra.Vaos.Cast<Elemento>(), 'V'), IrmaoDe = raiz == null ? null : raiz.Numero };
                    obra.Vaos.Add(vao);
                }
                if (raiz == null) raiz = vao;
                PreencherVao(vao, linha, a, b, regras, cabo);
                incluidos.Add(vao);
            }
            Recalcular(obra, regras);
            return incluidos;
        }

        private static void PreencherVao(Vao v, LinhaDesenho linha, Poste a, Poste b, RegrasProjeto regras, LeituraDeTexto cabo)
        {
            v.Manual = true;
            v.LinhaHandle = linha.Handle;
            v.Pontos = new List<Ponto>(linha.Pontos);
            v.LinhaLayer = linha.Layer;
            v.LinhaCor = linha.Cor;
            v.Tipo = regras.Leitura.ClassificarRede(linha.Layer, linha.Cor) ?? (regras.Leitura.Regra == "SO_LAYER" ? null : regras.Leitura.CorDoProjeto(linha.Cor)) ?? "";
            v.PosteA = a != null ? a.Numero : null;
            v.PosteB = b != null ? b.Numero : null;
            if (cabo != null)
                LeitorProjeto.Anotar(v, new TextoLido { T = cabo.Texto, Comprimento = cabo.TextoComprimento,
                    R = new ResultadoLeitor { Entidade = "CABO", Ativo = cabo.Ativo, Operacao = cabo.Operacao } });
        }

        public static void Remover(Obra obra, Elemento e, RegrasProjeto regras)
        {
            Poste p = e as Poste;
            if (p != null)
            {
                obra.Postes.Remove(p);
                foreach (Vao v in obra.Vaos)
                {
                    if (v.PosteA == p.Numero) { v.PosteA = null; v.VinculoA = null; }
                    if (v.PosteB == p.Numero) { v.PosteB = null; v.VinculoB = null; }
                }
            }
            if (e is Vao) obra.Vaos.Remove((Vao)e);
            if (e is Avulso) obra.Avulsos.Remove((Avulso)e);
            Recalcular(obra, regras);
        }

        /// <summary>Depois de uma mudança: vínculos automáticos de novo (mantendo os manuais), parâmetros e conferência.</summary>
        public static void Recalcular(Obra obra, RegrasProjeto regras)
        {
            Vinculador.Vincular(obra, regras, true);
            Conferencia.Avaliar(obra, regras);
        }

        // ── ajustes do pop-up ──
        public static void AjustarVao(Obra obra, Vao v, string operacao, double? metros, RegrasProjeto regras)
        {
            if (!string.IsNullOrEmpty(operacao)) v.Operacao = operacao;
            if (metros.HasValue) v.Metros = metros;
            if (v.Cabo != null) v.InfoCabo = null;
            Recalcular(obra, regras);
        }

        public static void AjustarPoste(Obra obra, Poste p, string operacao, RegrasProjeto regras)
        {
            if (!string.IsNullOrEmpty(operacao)) p.Operacao = operacao;
            Recalcular(obra, regras);
        }

        // ── TASK-017: corrigir o ativo, tirar/devolver à obra, linha nova sem desenho ──

        /// <summary>Troca o ativo (o texto da anotação na obra). O do desenho fica guardado (voltar ao lido); o vão relê cabo e metros do texto novo.</summary>
        public static void TrocarAtivo(Obra obra, Elemento e, string novo, RegrasProjeto regras)
        {
            string t = (novo ?? "").Trim();
            if (t.Length == 0 || t == e.AtivoLido) return;
            if (e.AtivoDoDesenho == null) e.AtivoDoDesenho = e.AtivoLido ?? "";
            if (t == e.AtivoDoDesenho) e.AtivoDoDesenho = null;
            AplicarAtivo(e, t);
            Recalcular(obra, regras);
        }

        private static void AplicarAtivo(Elemento e, string ativo)
        {
            e.AtivoLido = ativo;
            Vao v = e as Vao;
            if (v != null)
            {
                v.Cabo = Anotacoes.LerCabo(ativo);
                v.MetrosLidos = v.Cabo.Comprimento;
                v.Metros = v.Cabo.Comprimento;
                v.InfoCabo = null;
            }
            Poste p = e as Poste;
            if (p != null) p.Lido = Anotacoes.LerPoste(ativo);
        }

        /// <summary>Tira da obra. O que foi montado à mão sai de vez; o que veio do desenho fica marcado (não volta ao Atualizar; Restaurar devolve).</summary>
        public static void Excluir(Obra obra, Elemento e, RegrasProjeto regras)
        {
            if (e.Manual) { Remover(obra, e, regras); return; }
            e.Excluido = true;
            Vao v = e as Vao;
            if (v != null) { v.VinculoA = null; v.VinculoB = null; }
            Recalcular(obra, regras);
        }

        public static void Restaurar(Obra obra, Elemento e, RegrasProjeto regras)
        {
            e.Excluido = false;
            Recalcular(obra, regras);
        }

        /// <summary>Linha nova sem apontar no desenho: vai para a Tabela de Cabos ("CABO") ou de Outros (demais), como as anotações soltas.</summary>
        public static Avulso NovaLinha(Obra obra, bool cabo, string ativo, string operacao, RegrasProjeto regras)
        {
            Avulso a = new Avulso();
            a.Numero = "L" + Proximo(obra.Avulsos.Cast<Elemento>(), 'L');
            a.Manual = true;
            a.Entidade = cabo ? "CABO" : "ESTRUTURA";
            a.AtivoLido = (ativo ?? "").Trim();
            a.OperacaoLida = string.IsNullOrEmpty(operacao) ? "I" : operacao;
            a.Operacao = a.OperacaoLida;
            a.TextoOriginal = a.AtivoLido;
            obra.Avulsos.Add(a);
            Recalcular(obra, regras);
            return a;
        }

        /// <summary>TASK-018: o que um texto (TEXT/MTEXT) escolhido no desenho vira na linha nova — pelas regras do leitor.</summary>
        public class LeituraDeTexto
        {
            public string Ativo = "", Operacao = "I", Entidade;
            public bool Reconhecido;          // alguma regra do leitor entendeu o texto
            public bool? Cabo;                // pelas regras: CABO → Tabela de Cabos; outra entidade → Outros; null = não se sabe
            public TextoDesenho Texto;        // o texto usado (ligação com o desenho)
            public string Lido = "";          // o que foi selecionado, como está no desenho
            public TextoDesenho TextoComprimento;
            public string Comprimento;        // TASK-021: comprimento achado ("35 m"), null = não achou
            public bool ComprimentoDaVizinhanca;   // veio de um texto do desenho que não foi selecionado
        }

        public static LeituraDeTexto LerTextoParaLinha(List<TextoDesenho> textos, RegrasProjeto regras)
        {
            LeituraDeTexto r = new LeituraDeTexto();
            if (textos == null || textos.Count == 0) return r;
            List<TextoLido> lidos = LeitorProjeto.ClassificarTextos(textos, regras);
            r.Lido = string.Join(" / ", textos.Select(t => t.Texto).ToArray());
            TextoLido bom = lidos.FirstOrDefault(t => t.R.Entidade != null && t.R.Entidade != "0" && !string.IsNullOrEmpty(t.R.Ativo));
            if (bom != null)
            {
                r.Reconhecido = true;
                r.Ativo = bom.R.Ativo;
                r.Entidade = bom.R.Entidade;
                r.Cabo = bom.R.Entidade == "CABO";
                if (!string.IsNullOrEmpty(bom.R.Operacao) && bom.R.Operacao != "0") r.Operacao = bom.R.Operacao;
                r.Texto = bom.T;
                return r;
            }
            TextoLido primeiro = lidos.Count > 0 ? lidos[0] : null;
            r.Ativo = lidos.Count > 0 ? AtivoDosTextos(lidos) : MotorLeitor.LimparTexto(string.Join(", ", textos.Select(t => t.Texto).ToArray()));
            r.Texto = primeiro != null ? primeiro.T : textos[0];
            // sem classificação: se o texto montado é um cabo com comprimento (como a regra "Cabo (label + comprimento)"), vai para Cabos
            CaboLido cab = Anotacoes.LerCabo(r.Ativo);
            if (cab != null && cab.Erro == null && !string.IsNullOrEmpty(cab.Codigo) && cab.Comprimento.HasValue) r.Cabo = true;
            if (primeiro != null && !string.IsNullOrEmpty(primeiro.R.Operacao) && primeiro.R.Operacao != "0") r.Operacao = primeiro.R.Operacao;
            return r;
        }

        // ── TASK-021: "+ Cabo" (e "+ Vão") — o texto escolhido é SEMPRE lido como cabo; o comprimento vem da mesma linha, dos textos
        // escolhidos ou da VIZINHANÇA no desenho (como a regra de Pernambuco "comprimento na linha de baixo", mas pela posição no desenho:
        // o texto "35 m" ao lado ou embaixo do rótulo, mesmo sem ter sido selecionado). ──
        private static readonly System.Text.RegularExpressions.Regex SoComprimento =
            new System.Text.RegularExpressions.Regex(@"^\s*(\d+(?:[.,]\d+)?)\s*m(?:etros?)?\s*\.?\s*$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        private static readonly System.Text.RegularExpressions.Regex ComComprimento =
            new System.Text.RegularExpressions.Regex(@"\s(\d+(?:[.,]\d+)?)\s*m(?:etros?)?\s*$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);

        /// <summary>O texto é só um comprimento ("35 m", "35,5M", "40 metros")?</summary>
        public static bool EhComprimento(string texto) { return SoComprimento.IsMatch(texto ?? ""); }

        /// <summary>"3XA35P(7,94)-13,8KV 35 m" já tem o comprimento no fim?</summary>
        public static bool TemComprimento(string ativo) { return ComComprimento.IsMatch(ativo ?? ""); }

        /// <summary>O texto de comprimento mais perto do rótulo (até 8 alturas; ao lado ou embaixo vale mais que em cima), fora dos já escolhidos.</summary>
        public static TextoDesenho ComprimentoVizinho(TextoDesenho rotulo, IEnumerable<TextoDesenho> escolhidos, IEnumerable<TextoDesenho> desenho)
        {
            if (rotulo == null || desenho == null) return null;
            double h = rotulo.Altura > 0 ? rotulo.Altura : 1;
            Ponto c = rotulo.Centro;
            HashSet<string> usados = new HashSet<string>((escolhidos ?? Enumerable.Empty<TextoDesenho>()).Select(t => t.Handle + "|" + t.Parte));
            TextoDesenho melhor = null;
            double nota = double.MaxValue;
            foreach (TextoDesenho t in desenho)
            {
                if (t == null || t.Tipo == "BLOCO" || t.Tipo == "BLOCO_DEF" || usados.Contains(t.Handle + "|" + t.Parte) || !EhComprimento(t.Texto)) continue;
                double d = c.Dist(t.Centro);
                if (d > 8 * h) continue;
                double n = d + (t.Centro.Y > c.Y + 0.5 * h ? 4 * h : 0);      // em cima do rótulo: só se não houver ao lado/embaixo
                if (n < nota) { nota = n; melhor = t; }
            }
            return melhor;
        }

        /// <summary>Um cabo por linha selecionada. Um texto que só contém metros pertence a um único cabo.</summary>
        public static List<LeituraDeTexto> LerCabosDeTextos(List<TextoDesenho> textos, RegrasProjeto regras, Func<IEnumerable<TextoDesenho>> desenho)
        {
            var selecionados = (textos ?? new List<TextoDesenho>()).Where(t => t.Texto != null && t.Texto.Trim().Length != 0)
                .OrderByDescending(t => t.Y).ThenBy(t => t.X).ThenBy(t => t.Parte).ToList();
            var rotulos = selecionados.Where(t => !EhComprimento(t.Texto)).ToList();
            var lidos = rotulos.Select(t => LerCaboDeTexto(new List<TextoDesenho> { t }, regras, null)).ToList();
            var comprimentos = selecionados.Where(t => EhComprimento(t.Texto)).ToList();
            if (desenho != null && lidos.Where(l => l.Comprimento == null).Any(l =>
                !comprimentos.Any(m => DonoDoComprimento(m, rotulos, true) == l.Texto)))
            {
                var idsSelecionados = new HashSet<string>(selecionados.Select(IdTexto));
                comprimentos.AddRange(desenho().Where(t => t.Tipo != "BLOCO" && t.Tipo != "BLOCO_DEF"
                    && EhComprimento(t.Texto) && !idsSelecionados.Contains(IdTexto(t))));
            }
            var porCabo = comprimentos.Select(t => new { Texto = t, Escolhido = selecionados.Contains(t) })
                .ToLookup(m => DonoDoComprimento(m.Texto, rotulos, m.Escolhido));
            for (int i = 0; i < lidos.Count; i++)
            {
                LeituraDeTexto lido = lidos[i];
                if (lido.Comprimento != null) continue;
                var comprimento = porCabo[lido.Texto].OrderByDescending(m => m.Escolhido)
                    .ThenByDescending(m => MesmoMText(lido.Texto, m.Texto))
                    .ThenBy(m => lido.Texto.Centro.Dist(m.Texto.Centro)).FirstOrDefault();
                if (comprimento == null) continue;
                var completo = LerCaboDeTexto(new List<TextoDesenho> { lido.Texto, comprimento.Texto }, regras, null);
                completo.TextoComprimento = comprimento.Texto;
                completo.ComprimentoDaVizinhanca = !comprimento.Escolhido;
                lidos[i] = completo;
            }
            return lidos;
        }

        private static string IdTexto(TextoDesenho t) { return t.Handle + "|" + t.Parte; }

        private static bool MesmoMText(TextoDesenho a, TextoDesenho b)
        {
            return !string.IsNullOrEmpty(a.Handle) && a.Handle == b.Handle;
        }

        private static TextoDesenho DonoDoComprimento(TextoDesenho metros, List<TextoDesenho> rotulos, bool escolhido)
        {
            var doMesmoTexto = rotulos.Where(t => MesmoMText(t, metros)).ToList();
            if (doMesmoTexto.Count > 0)
                return doMesmoTexto.Where(t => t.Parte < metros.Parte).OrderByDescending(t => t.Parte).FirstOrDefault();
            return rotulos.Select(t => new { Texto = t, Altura = t.Altura > 0 ? t.Altura : 1, Distancia = t.Centro.Dist(metros.Centro) })
                .Where(x => escolhido || x.Distancia <= 8 * x.Altura)
                .OrderBy(x => x.Distancia + (metros.Centro.Y > x.Texto.Centro.Y + 0.5 * x.Altura ? 4 * x.Altura : 0))
                .Select(x => x.Texto).FirstOrDefault();
        }

        /// <summary>Lê os textos escolhidos COMO CABO: rótulo pelas regras do leitor (o que elas reconhecem como CABO, ou o processamento) e o
        /// comprimento da mesma linha, das linhas escolhidas ou da vizinhança no desenho (`desenho`: todos os textos; chamado só se faltar).</summary>
        public static LeituraDeTexto LerCaboDeTexto(List<TextoDesenho> textos, RegrasProjeto regras, Func<IEnumerable<TextoDesenho>> desenho)
        {
            LeituraDeTexto r = new LeituraDeTexto();
            r.Cabo = true;
            if (textos == null || textos.Count == 0) return r;
            List<TextoLido> lidos = LeitorProjeto.ClassificarTextos(textos, regras);
            r.Lido = string.Join(" / ", textos.Select(t => t.Texto).ToArray());
            TextoLido cabo = lidos.FirstOrDefault(t => t.R.Entidade == "CABO" && !string.IsNullOrEmpty(t.R.Ativo));
            string ativo;
            TextoDesenho rotulo;
            if (cabo != null)
            {
                r.Reconhecido = true;
                r.Entidade = "CABO";
                ativo = cabo.R.Ativo;
                rotulo = cabo.T;
                if (!string.IsNullOrEmpty(cabo.R.Operacao) && cabo.R.Operacao != "0") r.Operacao = cabo.R.Operacao;
            }
            else
            {
                // nenhuma regra classificou como cabo: o rótulo é o que o processamento montou (sem as linhas que são só comprimento)
                List<TextoLido> partes = lidos.Where(t => !EhComprimento(t.T.Texto)).ToList();
                List<string> l = partes.Select(t => !string.IsNullOrEmpty(t.R.Ativo) ? t.R.Ativo : t.T.Texto).Where(x => !string.IsNullOrEmpty(x) && !EhComprimento(x)).Distinct().ToList();
                ativo = string.Join(" ", l.ToArray());
                rotulo = partes.Count > 0 ? partes[0].T : textos[0];
                TextoLido comOp = lidos.FirstOrDefault(t => !string.IsNullOrEmpty(t.R.Operacao) && t.R.Operacao != "0");
                if (comOp != null) r.Operacao = comOp.R.Operacao;
            }
            if (!TemComprimento(ativo))
            {
                TextoDesenho comp = textos.FirstOrDefault(t => EhComprimento(t.Texto));       // escolhido junto
                if (comp == null && desenho != null) { comp = ComprimentoVizinho(rotulo, textos, desenho()); if (comp != null) r.ComprimentoDaVizinhanca = true; }
                if (comp != null)
                {
                    string num = SoComprimento.Match(comp.Texto).Groups[1].Value;
                    ativo = ativo.TrimEnd() + " " + num + " m";
                    r.Comprimento = num + " m";
                }
            }
            else r.Comprimento = ComComprimento.Match(ativo).Groups[1].Value + " m";
            r.Ativo = MotorLeitor.LimparTexto(ativo);
            r.Texto = rotulo;
            return r;
        }

        /// <summary>Liga a linha ao texto do desenho (localizar, marcas, layer). Se o ativo foi mudado antes de adicionar, o do desenho fica guardado.</summary>
        public static void LigarTexto(Elemento e, LeituraDeTexto lido)
        {
            if (lido == null || lido.Texto == null) return;
            TextoDesenho t = lido.Texto;
            e.TextoHandle = t.Handle;
            e.ParteTexto = t.Parte;
            e.OrigemMetros = lido.TextoComprimento;
            e.TextoOriginal = t.Texto;
            e.TextoLayer = t.Layer;
            e.TextoCor = t.Cor;
            e.TextoX = t.X;
            e.TextoY = t.Y;
            if (!string.IsNullOrEmpty(lido.Ativo) && lido.Ativo != e.AtivoLido) e.AtivoDoDesenho = lido.Ativo;
            if (!string.IsNullOrEmpty(lido.Operacao)) e.OperacaoLida = lido.Operacao;
        }

        public static void VoltarAoLido(Obra obra, Elemento e, RegrasProjeto regras)
        {
            if (e.AtivoDoDesenho != null) { AplicarAtivo(e, e.AtivoDoDesenho); e.AtivoDoDesenho = null; }
            e.Operacao = e.OperacaoLida;
            Vao v = e as Vao;
            if (v != null)
            {
                v.Metros = v.MetrosLidos;
                if (v.VinculoA != null && v.VinculoA.Manual) v.VinculoA = null;
                if (v.VinculoB != null && v.VinculoB.Manual) v.VinculoB = null;
            }
            Recalcular(obra, regras);
        }

        /// <summary>Liga à mão a ponta (A ou B) do vão à estrutura k do poste daquela ponta. Devolve o motivo se as regras não deixam.</summary>
        public static string VincularManual(Obra obra, Vao v, bool pontaA, int k, RegrasProjeto regras)
        {
            Poste p = obra.PostePorNumero(pontaA ? v.PosteA : v.PosteB);
            if (p == null) return "essa ponta do vão não tem poste";
            if (k < 0 || k >= p.Lido.Estruturas.Count) return "o poste " + p.Numero + " não tem essa estrutura";
            Vinculo vi = new Vinculo();
            vi.Poste = p.Numero; vi.Estrutura = p.Lido.Estruturas[k].Nome; vi.K = k; vi.Manual = true;
            if (pontaA) v.VinculoA = vi; else v.VinculoB = vi;
            Recalcular(obra, regras);
            return null;
        }

        /// <summary>"Atualizar": a obra relida recebe de volta os ajustes da anterior (casando por handle da linha/bloco ou do texto).</summary>
        public static void ReaplicarAjustes(Obra nova, Obra antiga, RegrasProjeto regras)
        {
            if (antiga == null) return;
            nova.Conferidos = antiga.Conferidos;
            nova.LeituraAnterior = Auditoria.Fotografar(antiga);
            foreach (Poste pa in antiga.Postes)
            {
                Poste pn = nova.Postes.FirstOrDefault(x => !string.IsNullOrEmpty(pa.TextoHandle) && x.TextoHandle == pa.TextoHandle && x.ParteTexto == pa.ParteTexto)
                        ?? nova.Postes.FirstOrDefault(x => !string.IsNullOrEmpty(pa.Handle) && x.Handle == pa.Handle && x.IrmaoDe == null);
                if (pn == null) continue;
                if (pa.Excluido) pn.Excluido = true;
                if (pn.TextoHandle != pa.TextoHandle) continue;
                LevarAtivo(pa, pn);
                if (pa.OperacaoAjustada) pn.Operacao = pa.Operacao;
            }
            foreach (Vao va in antiga.Vaos)
            {
                // pela anotação primeiro (vários vãos "irmãos" dividem a mesma linha), depois pela linha
                Vao vn = nova.Vaos.FirstOrDefault(x => !string.IsNullOrEmpty(va.TextoHandle) && x.TextoHandle == va.TextoHandle && x.ParteTexto == va.ParteTexto)
                      ?? nova.Vaos.FirstOrDefault(x => x.LinhaHandle == va.LinhaHandle && x.Trecho == va.Trecho && x.IrmaoDe == null);
                if (vn != null && va.Excluido) vn.Excluido = true;
                if (vn == null || vn.TextoHandle != va.TextoHandle) continue;
                LevarAtivo(va, vn);
                if (va.OperacaoAjustada) vn.Operacao = va.Operacao;
                if (va.Metros.HasValue && va.Metros != va.MetrosLidos) vn.Metros = va.Metros;
                foreach (bool a in new[] { true, false })
                {
                    Vinculo vi = a ? va.VinculoA : va.VinculoB;
                    if (vi == null || !vi.Manual) continue;
                    Poste pAnt = antiga.PostePorNumero(vi.Poste);
                    Poste pNovo = pAnt == null ? null : nova.Postes.FirstOrDefault(x => x.Handle == pAnt.Handle && x.TextoHandle == pAnt.TextoHandle);
                    if (pNovo == null || (a ? vn.PosteA : vn.PosteB) != pNovo.Numero) continue;
                    Vinculo nv = new Vinculo();
                    nv.Poste = pNovo.Numero; nv.Estrutura = vi.Estrutura; nv.K = vi.K; nv.Manual = true;
                    if (a) vn.VinculoA = nv; else vn.VinculoB = nv;
                }
            }
            foreach (Avulso aa in antiga.Avulsos)
            {
                if (aa.Manual)
                {
                    // linha adicionada à mão (digitada ou lida de um texto escolhido): volta como estava — a não ser que a leitura
                    // agora já pegue aquele texto (aí vale a leitura, para não ficar em dobro)
                    if (!string.IsNullOrEmpty(aa.TextoHandle) && nova.Elementos.Any(x => x.TextoHandle == aa.TextoHandle && x.ParteTexto == aa.ParteTexto)) continue;
                    Avulso c = new Avulso();
                    c.Numero = aa.Numero; c.Manual = true; c.Entidade = aa.Entidade; c.AtivoLido = aa.AtivoLido; c.TextoOriginal = aa.TextoOriginal;
                    c.OperacaoLida = aa.OperacaoLida; c.Operacao = aa.Operacao; c.AtivoDoDesenho = aa.AtivoDoDesenho; c.Excluido = aa.Excluido;
                    c.TextoHandle = aa.TextoHandle; c.ParteTexto = aa.ParteTexto; c.OrigemMetros = aa.OrigemMetros; c.TextoLayer = aa.TextoLayer; c.TextoCor = aa.TextoCor; c.TextoX = aa.TextoX; c.TextoY = aa.TextoY;
                    nova.Avulsos.Add(c);
                    continue;
                }
                Avulso an = nova.Avulsos.FirstOrDefault(x => !string.IsNullOrEmpty(aa.TextoHandle) && x.TextoHandle == aa.TextoHandle && x.ParteTexto == aa.ParteTexto);
                if (an == null) continue;
                if (aa.Excluido) an.Excluido = true;
                LevarAtivo(aa, an);
                if (aa.OperacaoAjustada) an.Operacao = aa.Operacao;
            }
            Recalcular(nova, regras);
        }

        /// <summary>Atualizar: a correção do ativo vale enquanto o texto do desenho continuar o mesmo (se mudou no desenho, vale o desenho).</summary>
        private static void LevarAtivo(Elemento antigo, Elemento novo)
        {
            if (antigo.AtivoDoDesenho == null || novo.AtivoLido != antigo.AtivoDoDesenho) return;
            novo.AtivoDoDesenho = novo.AtivoLido;
            AplicarAtivo(novo, antigo.AtivoLido);
        }
    }
}
