// Nucleo/LeitorProjeto.cs — "Ler projeto": do desenho (textos, postes, redes) para a obra (postes P1…, vãos V1…, vínculos).
//   1) cada texto passa pelo motor de regras do leitor (layer + cor → operação, ativo, entidade), na ordem do leitor de DXF;
//   2) redes = linhas nas layers REDE_MT/REDE_BT cuja cor confirma a layer (services/dxf_geometria.py); cada rede é cortada
//      nos postes que estão sobre ela: cada pedaço entre dois postes é um VÃO (vão = linha + anotação);
//   3) anotação de cabo → vão mais próximo; anotação de poste → poste mais próximo (até 4 × a altura do texto, seguindo a
//      linha-guia quando houver), como a associação do DXF da aplicação;
//   4) numeração seguindo a rede (P1, V1, P2…), vínculo automático pelas regras de vínculo e conferência (Ok/Conferir/Erro).
using System;
using System.Collections.Generic;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    public class TextoLido
    {
        public TextoDesenho T;
        public ResultadoLeitor R;
    }

    public static class LeitorProjeto
    {
        // ── 1) motor do leitor ──
        public static List<TextoLido> ClassificarTextos(IEnumerable<TextoDesenho> textos, RegrasProjeto regras)
        {
            List<TextoDesenho> ordem = textos.Where(t => !string.IsNullOrEmpty(t.Texto))
                .OrderByDescending(t => t.Y).ThenBy(t => t.X).ToList();       // OrderBy é estável: parágrafos do mesmo MTEXT ficam em ordem
            List<ItemLeitor> itens = ordem.Select(t =>
            {
                ItemLeitor i = new ItemLeitor();
                i.Texto = t.Texto; i.Cor = string.IsNullOrEmpty(t.Cor) ? "#000000" : t.Cor; i.Layer = t.Layer ?? ""; i.X = t.X; i.Y = t.Y;
                return i;
            }).ToList();
            List<TextoLido> saida = new List<TextoLido>();
            for (int k = 0; k < itens.Count; k++)
            {
                TextoLido tl = new TextoLido();
                tl.T = ordem[k];
                tl.R = MotorLeitor.ProcessarEClassificar(itens, k, regras.Processamento, regras.Classificacao);
                saida.Add(tl);
            }
            return saida;
        }

        // ── 2) redes ──
        public static string TipoPelaCor(string hex)
        {
            if (string.IsNullOrEmpty(hex) || hex.Length < 7) return null;
            int r, g, b;
            try
            {
                r = Convert.ToInt32(hex.Substring(1, 2), 16); g = Convert.ToInt32(hex.Substring(3, 2), 16); b = Convert.ToInt32(hex.Substring(5, 2), 16);
            }
            catch (FormatException) { return null; }
            if (r >= 170 && g <= 100 && b <= 100) return "MT";
            if (g >= 140 && r <= 120 && b <= 120) return "BT";
            return null;
        }

        public static string TipoPelaLayer(string layer)
        {
            string u = (layer ?? "").ToUpperInvariant();
            if (u.Contains("REDE_MT")) return "MT";
            if (u.Contains("REDE_BT")) return "BT";
            return null;
        }

        /// <summary>A cor decide o tipo e a layer confirma. Cor fora de vermelho/verde: vale a layer. Divergência: null (descarta).</summary>
        public static string ClassificarRede(string layer, string hex)
        {
            string daLayer = TipoPelaLayer(layer);
            if (daLayer == null) return null;
            string daCor = TipoPelaCor(hex);
            return daCor == null || daCor == daLayer ? daLayer : null;
        }

        private class Rede { public LinhaDesenho L; public string Tipo; }

        private class Contexto
        {
            public RegrasProjeto Regras;
            public double H;
            public List<List<Ponto>> Guias = new List<List<Ponto>>();
            public List<Poste> Postes = new List<Poste>();
            public List<Vao> Vaos = new List<Vao>();
        }

        public static double AlturaTipica(IEnumerable<TextoDesenho> textos)
        {
            double h = Geometria.Mediana(textos.Where(t => t.Altura > 0 && t.Tipo != "BLOCO").Select(t => t.Altura).ToList());
            return h > 0 ? h : 1.0;
        }

        public static Obra Ler(Desenho d, RegrasProjeto regras)
        {
            Contexto ctx = new Contexto();
            ctx.Regras = regras;
            ctx.H = AlturaTipica(d.Textos);

            List<Rede> redes = new List<Rede>();
            foreach (LinhaDesenho l in d.Linhas)
            {
                if (l.Pontos.Count < 2) continue;
                string tipo = l.Leader ? null : ClassificarRede(l.Layer, l.Cor);
                if (tipo != null) { Rede r = new Rede(); r.L = l; r.Tipo = tipo; redes.Add(r); }
                else if (TipoPelaLayer(l.Layer) == null && Geometria.Comprimento(l.Pontos) <= regras.GuiaMaxAlturas * ctx.H) ctx.Guias.Add(l.Pontos);
            }

            foreach (PosteDesenho pd in d.Postes)
            {
                Poste p = new Poste();
                p.Handle = pd.Handle; p.X = pd.X; p.Y = pd.Y; p.Bloco = pd.Bloco ?? "";
                ctx.Postes.Add(p);
            }

            double tol = regras.TolPosteAlturas * ctx.H;
            foreach (Rede r in redes) CortarEmVaos(r, ctx, tol);

            List<TextoLido> lidos = ClassificarTextos(d.Textos, regras);
            List<Avulso> avulsos = new List<Avulso>();
            AssociarTextos(lidos, ctx, avulsos);

            Obra obra = new Obra();
            obra.Projeto = regras.Codigo;
            obra.NomeRegras = regras.Nome;
            obra.Desenho = d.Arquivo ?? "";
            obra.Nome = NomeDaObra(d.Arquivo);
            obra.Modo = "ler";
            obra.AlturaTexto = ctx.H;
            obra.Postes = ctx.Postes;
            obra.Vaos = ctx.Vaos;
            obra.Avulsos = avulsos;
            Numerar(obra);
            Vinculador.Vincular(obra, regras, false);
            Conferencia.Avaliar(obra, regras);
            return obra;
        }

        public static string NomeDaObra(string arquivo)
        {
            if (string.IsNullOrEmpty(arquivo)) return "obra";
            string n = arquivo.Replace('\\', '/');
            n = n.Substring(n.LastIndexOf('/') + 1);
            int p = n.LastIndexOf('.');
            return p > 0 ? n.Substring(0, p) : n;
        }

        private class NaRede { public Poste P; public double S; public double D; }

        private static void CortarEmVaos(Rede r, Contexto ctx, double tol)
        {
            List<Ponto> pts = r.L.Pontos;
            double comp = Geometria.Comprimento(pts);
            List<NaRede> nela = new List<NaRede>();
            foreach (Poste p in ctx.Postes)
            {
                Projecao pr = Geometria.Projetar(p.Ponto, pts);
                if (pr.Distancia <= tol) { NaRede n = new NaRede(); n.P = p; n.S = pr.Estacao; n.D = pr.Distancia; nela.Add(n); }
            }
            nela = nela.OrderBy(n => n.S).ToList();
            // dois postes "no mesmo ponto" da rede: fica o mais próximo da linha
            List<NaRede> limpos = new List<NaRede>();
            foreach (NaRede n in nela)
            {
                if (limpos.Count > 0 && n.S - limpos[limpos.Count - 1].S < tol)
                {
                    if (n.D < limpos[limpos.Count - 1].D) limpos[limpos.Count - 1] = n;
                    continue;
                }
                limpos.Add(n);
            }
            List<double> cortes = new List<double>();
            List<Poste> postes = new List<Poste>();
            if (limpos.Count == 0 || limpos[0].S > tol) { cortes.Add(0); postes.Add(null); }
            foreach (NaRede n in limpos) { cortes.Add(n.S); postes.Add(n.P); }
            if (limpos.Count == 0 || limpos[limpos.Count - 1].S < comp - tol) { cortes.Add(comp); postes.Add(null); }
            int trecho = 0;
            for (int i = 0; i + 1 < cortes.Count; i++)
            {
                if (cortes[i + 1] - cortes[i] <= 1e-9) continue;
                Vao v = new Vao();
                v.LinhaHandle = r.L.Handle;
                v.Trecho = trecho++;
                v.Pontos = Geometria.Trecho(pts, cortes[i], cortes[i + 1]);
                v.Tipo = r.Tipo;
                v.LinhaLayer = r.L.Layer;
                v.LinhaCor = r.L.Cor;
                v.PosteObjA = postes[i];
                v.PosteObjB = postes[i + 1];
                ctx.Vaos.Add(v);
            }
        }

        // ── 3) associação texto → vão / poste ──
        private static List<Ponto> Ancoras(Ponto p, double raio, Contexto ctx)
        {
            List<KeyValuePair<double, Ponto>> achados = new List<KeyValuePair<double, Ponto>>();
            foreach (List<Ponto> g in ctx.Guias)
            {
                double d0 = p.Dist(g[0]), d1 = p.Dist(g[g.Count - 1]);
                if (d0 <= raio) achados.Add(new KeyValuePair<double, Ponto>(d0, g[g.Count - 1]));
                if (d1 <= raio) achados.Add(new KeyValuePair<double, Ponto>(d1, g[0]));
            }
            List<Ponto> saida = achados.OrderBy(a => a.Key).Select(a => a.Value).ToList();
            saida.Add(p);
            return saida;
        }

        private static Vao VaoMaisProximo(List<Ponto> busca, double raio, Contexto ctx, out double dist)
        {
            foreach (Ponto q in busca)
            {
                Vao melhor = null;
                double md = double.MaxValue;
                foreach (Vao v in ctx.Vaos)
                {
                    double dd = Geometria.Projetar(q, v.Pontos).Distancia;
                    if (dd <= raio && dd < md) { md = dd; melhor = v; }
                }
                if (melhor != null) { dist = md; return melhor; }
            }
            dist = double.MaxValue;
            return null;
        }

        private static Poste PosteMaisProximo(List<Ponto> busca, double raio, Contexto ctx, out double dist)
        {
            foreach (Ponto q in busca)
            {
                Poste melhor = null;
                double md = double.MaxValue;
                foreach (Poste p in ctx.Postes)
                {
                    double dd = q.Dist(p.Ponto);
                    if (dd <= raio && dd < md) { md = dd; melhor = p; }
                }
                if (melhor != null) { dist = md; return melhor; }
            }
            dist = double.MaxValue;
            return null;
        }

        public static void Anotar(Elemento e, TextoLido tl)
        {
            e.TextoHandle = tl.T.Handle;
            e.TextoOriginal = tl.T.Texto;
            e.TextoLayer = tl.T.Layer;
            e.TextoCor = tl.T.Cor;
            e.TextoX = tl.T.X;
            e.TextoY = tl.T.Y;
            e.Entidade = tl.R.Entidade;
            e.AtivoLido = tl.R.Ativo;
            e.OperacaoLida = tl.R.Operacao;
            e.Operacao = tl.R.Operacao;
            Vao v = e as Vao;
            if (v != null)
            {
                v.Cabo = Anotacoes.LerCabo(tl.R.Ativo);
                v.MetrosLidos = v.Cabo.Comprimento;
                v.Metros = v.Cabo.Comprimento;
            }
            Poste p = e as Poste;
            if (p != null) p.Lido = Anotacoes.LerPoste(tl.R.Ativo);
        }

        private static bool EhCabo(string ent) { return ent == "CABO"; }
        private static bool EhPoste(string ent) { return ent == "POSTE" || ent == "ESTRUTURA"; }
        private static bool VaiParaOutros(string ent) { return ent != "0" && ent != "CABO" && ent != "RAMAIS" && ent != "IP"; }

        private static void AssociarTextos(List<TextoLido> lidos, Contexto ctx, List<Avulso> avulsos)
        {
            Dictionary<Vao, List<KeyValuePair<double, TextoLido>>> porVao = new Dictionary<Vao, List<KeyValuePair<double, TextoLido>>>();
            Dictionary<Poste, List<KeyValuePair<double, TextoLido>>> porPoste = new Dictionary<Poste, List<KeyValuePair<double, TextoLido>>>();
            foreach (TextoLido tl in lidos)
            {
                string ent = tl.R.Entidade ?? "0";
                if (ent == "0" || tl.T.Tipo == "BLOCO") continue;
                double h = tl.T.Altura > 0 ? tl.T.Altura : ctx.H;
                double raio = ctx.Regras.RaioAlturas * h;
                Ponto p = new Ponto(tl.T.X, tl.T.Y);
                List<Ponto> busca = Ancoras(p, raio, ctx);
                double dist;
                if (EhCabo(ent))
                {
                    Vao v = VaoMaisProximo(busca, raio, ctx, out dist);
                    if (v != null) { Adicionar(porVao, v, dist, tl); continue; }
                    avulsos.Add(NovoAvulso(tl, Estados.Conferir, "anotação de cabo sem vão por perto (rede REDE_MT/REDE_BT)"));
                    continue;
                }
                if (EhPoste(ent))
                {
                    Poste ps = PosteMaisProximo(busca, raio, ctx, out dist);
                    if (ps != null) { Adicionar(porPoste, ps, dist, tl); continue; }
                    // sem bloco de poste: vira um poste só da anotação (fica sem vãos)
                    Poste novo = new Poste();
                    novo.Handle = ""; novo.X = tl.T.X; novo.Y = tl.T.Y;
                    Anotar(novo, tl);
                    novo.MensagensFixas.Add("anotação de poste sem bloco de poste por perto");
                    ctx.Postes.Add(novo);
                    continue;
                }
                if (VaiParaOutros(ent)) avulsos.Add(NovoAvulso(tl, Estados.Ok, null));
            }
            foreach (KeyValuePair<Vao, List<KeyValuePair<double, TextoLido>>> kv in porVao)
            {
                List<KeyValuePair<double, TextoLido>> l = kv.Value.OrderBy(x => x.Key).ToList();
                Anotar(kv.Key, l[0].Value);
                for (int i = 1; i < l.Count; i++)
                {
                    Avulso a = NovoAvulso(l[i].Value, Estados.Conferir, null);
                    a.PertoDe = kv.Key;
                    avulsos.Add(a);
                }
            }
            foreach (KeyValuePair<Poste, List<KeyValuePair<double, TextoLido>>> kv in porPoste)
            {
                // a anotação do poste é a de entidade POSTE mais próxima; senão a ESTRUTURA mais próxima
                List<KeyValuePair<double, TextoLido>> l = kv.Value.OrderBy(x => x.Value.R.Entidade == "POSTE" ? 0 : 1).ThenBy(x => x.Key).ToList();
                Anotar(kv.Key, l[0].Value);
                for (int i = 1; i < l.Count; i++)
                {
                    Avulso a = NovoAvulso(l[i].Value, Estados.Conferir, null);
                    a.PertoDe = kv.Key;
                    avulsos.Add(a);
                }
            }
        }

        private static void Adicionar<T>(Dictionary<T, List<KeyValuePair<double, TextoLido>>> d, T chave, double dist, TextoLido tl)
        {
            List<KeyValuePair<double, TextoLido>> l;
            if (!d.TryGetValue(chave, out l)) { l = new List<KeyValuePair<double, TextoLido>>(); d[chave] = l; }
            l.Add(new KeyValuePair<double, TextoLido>(dist, tl));
        }

        private static Avulso NovoAvulso(TextoLido tl, string estado, string msg)
        {
            Avulso a = new Avulso();
            Anotar(a, tl);
            if (msg != null) a.MensagensFixas.Add(msg);
            return a;
        }

        // ── 4) numeração seguindo a rede ──
        private class Aresta { public Vao V; public Poste Outro; public double Dx, Dy; }

        public static void Numerar(Obra obra)
        {
            Dictionary<Poste, List<Aresta>> adj = new Dictionary<Poste, List<Aresta>>();
            foreach (Poste p in obra.Postes) adj[p] = new List<Aresta>();
            foreach (Vao v in obra.Vaos)
            {
                if (v.PosteObjA != null) AdicionarAresta(adj, v.PosteObjA, v.PosteObjB, v, v.Pontos[v.Pontos.Count - 1]);
                if (v.PosteObjB != null) AdicionarAresta(adj, v.PosteObjB, v.PosteObjA, v, v.Pontos[0]);
            }
            HashSet<Poste> vistos = new HashSet<Poste>();
            HashSet<Vao> vaosVistos = new HashSet<Vao>();
            List<Poste> ordemP = new List<Poste>();
            List<Vao> ordemV = new List<Vao>();
            foreach (Poste ini in obra.Postes.OrderBy(p => adj[p].Count == 0 ? 1 : 0).ThenBy(p => p.X).ThenByDescending(p => p.Y))
            {
                if (vistos.Contains(ini)) continue;
                // DFS iterativa: em cada poste segue primeiro o vão que mais "continua" a direção de chegada
                Stack<KeyValuePair<Poste, IEnumerator<Aresta>>> pilha = new Stack<KeyValuePair<Poste, IEnumerator<Aresta>>>();
                vistos.Add(ini);
                ordemP.Add(ini);
                pilha.Push(new KeyValuePair<Poste, IEnumerator<Aresta>>(ini, Ordenadas(adj[ini], 1, 0).GetEnumerator()));
                while (pilha.Count > 0)
                {
                    IEnumerator<Aresta> it = pilha.Peek().Value;
                    if (!it.MoveNext()) { pilha.Pop(); continue; }
                    Aresta a = it.Current;
                    if (vaosVistos.Contains(a.V)) continue;
                    vaosVistos.Add(a.V);
                    ordemV.Add(a.V);
                    if (a.Outro != null && !vistos.Contains(a.Outro))
                    {
                        vistos.Add(a.Outro);
                        ordemP.Add(a.Outro);
                        pilha.Push(new KeyValuePair<Poste, IEnumerator<Aresta>>(a.Outro, Ordenadas(adj[a.Outro], a.Dx, a.Dy).GetEnumerator()));
                    }
                }
            }
            foreach (Vao v in obra.Vaos.Where(x => !vaosVistos.Contains(x)).OrderByDescending(x => Geometria.Meio(x.Pontos).Y).ThenBy(x => Geometria.Meio(x.Pontos).X))
                ordemV.Add(v);
            for (int i = 0; i < ordemP.Count; i++) ordemP[i].Numero = "P" + (i + 1);
            for (int i = 0; i < ordemV.Count; i++)
            {
                Vao v = ordemV[i];
                v.Numero = "V" + (i + 1);
                // A = a ponta por onde a numeração chegou (poste de número menor)
                if (v.PosteObjA != null && v.PosteObjB != null && NumeroDe(v.PosteObjB) < NumeroDe(v.PosteObjA))
                {
                    Poste t = v.PosteObjA; v.PosteObjA = v.PosteObjB; v.PosteObjB = t;
                    v.Pontos.Reverse();
                }
                else if (v.PosteObjA == null && v.PosteObjB != null)
                {
                    v.PosteObjA = v.PosteObjB; v.PosteObjB = null;
                    v.Pontos.Reverse();
                }
                v.PosteA = v.PosteObjA != null ? v.PosteObjA.Numero : null;
                v.PosteB = v.PosteObjB != null ? v.PosteObjB.Numero : null;
            }
            obra.Postes = ordemP;
            obra.Vaos = ordemV;
            for (int i = 0; i < obra.Avulsos.Count; i++) obra.Avulsos[i].Numero = "A" + (i + 1);
        }

        private static int NumeroDe(Poste p)
        {
            int n;
            return p.Numero != null && int.TryParse(p.Numero.Substring(1), out n) ? n : int.MaxValue;
        }

        private static void AdicionarAresta(Dictionary<Poste, List<Aresta>> adj, Poste de, Poste para, Vao v, Ponto pontaLonge)
        {
            Aresta a = new Aresta();
            a.V = v;
            a.Outro = para;
            Ponto alvo = para != null ? para.Ponto : pontaLonge;
            a.Dx = alvo.X - de.X;
            a.Dy = alvo.Y - de.Y;
            adj[de].Add(a);
        }

        private static IEnumerable<Aresta> Ordenadas(List<Aresta> l, double dx, double dy)
        {
            return l.OrderBy(a => Geometria.Desvio(dx, dy, a.Dx, a.Dy)).ToList();
        }
    }
}
