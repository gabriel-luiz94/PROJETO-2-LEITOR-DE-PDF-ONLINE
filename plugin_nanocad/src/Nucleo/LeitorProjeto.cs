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

        // ── 2) redes (critérios em ConfigLeitura; estes atalhos usam o padrão de Pernambuco) ──
        private static readonly ConfigLeitura _padrao = new ConfigLeitura();
        public static string TipoPelaCor(string hex) { return ConfigLeitura.TipoPelaCor(hex); }
        public static string TipoPelaLayer(string layer) { return _padrao.TipoPelaLayer(layer); }
        public static string ClassificarRede(string layer, string hex) { return _padrao.ClassificarRede(layer, hex); }

        private class Rede { public LinhaDesenho L; public string Tipo; }

        private class Contexto
        {
            public RegrasProjeto Regras;
            public ConfigLeitura Cfg;
            public double H;
            public List<List<Ponto>> Guias = new List<List<Ponto>>();
            public List<Poste> Postes = new List<Poste>();
            public List<Vao> Vaos = new List<Vao>();
            public List<TextoIgnorado> Ignorados = new List<TextoIgnorado>();
            public List<TextoLido> Lidos;
            public int RedesSemPoste;
        }

        public static double AlturaTipica(IEnumerable<TextoDesenho> textos)
        {
            double h = Geometria.Mediana(textos.Where(t => t.Altura > 0 && t.Tipo != "BLOCO" && t.Tipo != "BLOCO_DEF").Select(t => t.Altura).ToList());
            return h > 0 ? h : 1.0;
        }

        public static Obra Ler(Desenho d, RegrasProjeto regras)
        {
            Contexto ctx = new Contexto();
            ctx.Regras = regras;
            ctx.Cfg = regras.Leitura ?? _padrao;
            ctx.H = AlturaTipica(d.Textos);
            ConfigLeitura cfg = ctx.Cfg;
            double tol = cfg.TolPosteAlturas * ctx.H;

            List<Rede> redes = new List<Rede>();
            foreach (LinhaDesenho l in d.Linhas)
            {
                if (l.Pontos.Count < 2) continue;
                string tipo = l.Leader ? null : cfg.ClassificarRede(l.Layer, l.Cor);
                if (tipo != null) { Rede r = new Rede(); r.L = l; r.Tipo = tipo; redes.Add(r); }
                else if (l.Leader) { ctx.Guias.Add(l.Pontos); _layerDaGuia[l.Pontos] = l.Layer; }      // LEADER/MLEADER: sempre guia, de qualquer tamanho
                else if (cfg.PodeSerGuia(l.Layer) && !Fechada(l.Pontos) && Geometria.Comprimento(l.Pontos) <= cfg.GuiaMaxAlturas * ctx.H)
                {
                    ctx.Guias.Add(l.Pontos);
                    _layerDaGuia[l.Pontos] = l.Layer;
                }
            }

            AgruparPostes(d.Postes, ctx, tol);
            redes = JuntarRedes(redes, ctx, tol);
            foreach (Rede r in redes) CortarEmVaos(r, ctx, tol);

            ctx.Lidos = ClassificarTextos(d.Textos, regras);
            List<Avulso> avulsos = new List<Avulso>();
            Dictionary<Elemento, List<TextoLido>> extras = AssociarTextos(ctx.Lidos, ctx, avulsos);

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
            obra.Ignorados = ctx.Ignorados;
            obra.RedesSemPoste = ctx.RedesSemPoste;
            Numerar(obra);
            CriarIrmaos(obra, extras);
            LerPlacas(obra, ctx);
            _layerDaGuia.Clear();
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

        private static bool Fechada(List<Ponto> pts)
        {
            return pts.Count > 2 && pts[0].Dist(pts[pts.Count - 1]) < 1e-6;
        }

        /// <summary>Blocos de poste no MESMO lugar (ex.: troca de poste: o que sai + o que entra) viram um poste só.</summary>
        private static void AgruparPostes(List<PosteDesenho> postes, Contexto ctx, double tol)
        {
            foreach (PosteDesenho pd in postes)
            {
                Poste junto = ctx.Postes.FirstOrDefault(x => x.Ponto.Dist(new Ponto(pd.X, pd.Y)) <= tol);
                if (junto != null)
                {
                    junto.OutrosHandles.Add(pd.Handle);
                    if (!string.IsNullOrEmpty(pd.Bloco) && junto.Bloco.IndexOf(pd.Bloco, StringComparison.Ordinal) < 0) junto.Bloco += " + " + pd.Bloco;
                    continue;
                }
                Poste p = new Poste();
                p.Handle = pd.Handle; p.X = pd.X; p.Y = pd.Y; p.Bloco = pd.Bloco ?? "";
                ctx.Postes.Add(p);
            }
        }

        private static bool TemPoste(Ponto q, Contexto ctx, double tol)
        {
            return ctx.Postes.Any(p => p.Ponto.Dist(q) <= tol);
        }

        /// <summary>Redes do mesmo tipo desenhadas em pedaços (pontas que se encontram FORA de poste, só duas no nó) viram uma linha só.</summary>
        private static List<Rede> JuntarRedes(List<Rede> redes, Contexto ctx, double tol)
        {
            double junta = Math.Max(ctx.H, 1e-6);
            List<Rede> l = redes.Select(r => new Rede { L = new LinhaDesenho { Handle = r.L.Handle, Layer = r.L.Layer, Cor = r.L.Cor, Pontos = new List<Ponto>(r.L.Pontos) }, Tipo = r.Tipo }).ToList();
            bool mudou = true;
            while (mudou)
            {
                mudou = false;
                for (int i = 0; i < l.Count && !mudou; i++)
                    for (int fi = 0; fi < 2 && !mudou; fi++)
                    {
                        List<Ponto> a = l[i].L.Pontos;
                        Ponto pa = fi == 0 ? a[0] : a[a.Count - 1];
                        if (TemPoste(pa, ctx, tol)) continue;
                        List<int> outros = new List<int>();
                        List<int> lados = new List<int>();
                        for (int j = 0; j < l.Count; j++)
                        {
                            if (j == i || l[j].Tipo != l[i].Tipo) continue;
                            List<Ponto> b = l[j].L.Pontos;
                            if (b[0].Dist(pa) <= junta) { outros.Add(j); lados.Add(0); }
                            else if (b[b.Count - 1].Dist(pa) <= junta) { outros.Add(j); lados.Add(1); }
                        }
                        if (outros.Count != 1) continue;                         // nó com 3+ linhas = derivação: não junta
                        List<Ponto> c = new List<Ponto>(l[outros[0]].L.Pontos);
                        List<Ponto> pi = new List<Ponto>(a);
                        if (fi == 0) pi.Reverse();                                // pi termina no nó
                        if (lados[0] == 1) c.Reverse();                           // c começa no nó
                        pi.AddRange(c.Skip(1));
                        l[i].L.Pontos = pi;
                        l[i].L.Handle = l[i].L.Handle + "+" + l[outros[0]].L.Handle;
                        l.RemoveAt(outros[0]);
                        mudou = true;
                    }
            }
            return l;
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
            // linha "de rede" que não passa por poste nenhum não é rede (marca, nuvem, desenho de apoio na mesma layer): fica fora
            if (nela.Count == 0 && ctx.Postes.Any(p => !string.IsNullOrEmpty(p.Handle))) { ctx.RedesSemPoste++; return; }
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
                if (postes[i] == null && postes[i + 1] == null && limpos.Count > 0) continue;
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
        private class Ancora { public Ponto P; public List<Ponto> Guia; }

        /// <summary>Segue uma guia a partir da ponta que toca o texto; se a outra ponta encosta em outra guia (guia em pedaços:
        /// sublinhado + chamada), continua por ela. Devolve a ponta final e o caminho percorrido.</summary>
        private static Ancora Seguir(List<Ponto> g, bool comecaNoInicio, Contexto ctx)
        {
            double junta = ctx.H * 0.5, limite = ctx.Cfg.GuiaMaxAlturas * ctx.H;
            string layer;
            _layerDaGuia.TryGetValue(g, out layer);
            List<Ponto> caminho = new List<Ponto>(comecaNoInicio ? g : Enumerable.Reverse(g));
            HashSet<List<Ponto>> usadas = new HashSet<List<Ponto>> { g };
            for (int passo = 0; passo < 4 && Geometria.Comprimento(caminho) < limite; passo++)
            {
                Ponto fim = caminho[caminho.Count - 1];
                List<Ponto> prox = null;
                bool inicio = true;
                foreach (List<Ponto> o in ctx.Guias)
                {
                    string lo;
                    // só continua por guia da MESMA layer (sublinhado + chamada); não passeia por ruas e outras linhas
                    if (usadas.Contains(o) || !_layerDaGuia.TryGetValue(o, out lo) || lo != layer) continue;
                    if (o[0].Dist(fim) <= junta) { prox = o; inicio = true; break; }
                    if (o[o.Count - 1].Dist(fim) <= junta) { prox = o; inicio = false; break; }
                }
                if (prox == null) break;
                usadas.Add(prox);
                caminho.AddRange((inicio ? prox : Enumerable.Reverse(prox)).Skip(1));
            }
            return new Ancora { P = caminho[caminho.Count - 1], Guia = caminho };
        }

        [ThreadStatic] private static Dictionary<List<Ponto>, string> _layerDaGuiaTs;
        private static Dictionary<List<Ponto>, string> _layerDaGuia
        {
            get { if (_layerDaGuiaTs == null) _layerDaGuiaTs = new Dictionary<List<Ponto>, string>(); return _layerDaGuiaTs; }
        }

        /// <summary>De onde a busca parte: a outra ponta das linhas de chamada do próprio texto (MLEADER), depois a outra ponta de cada
        /// guia que toca o CONTORNO do texto (da mais perto à mais longe, seguindo guias em pedaços), depois o próprio texto.</summary>
        private static List<Ancora> Ancoras(TextoDesenho t, Contexto ctx)
        {
            List<Ancora> saida = new List<Ancora>();
            foreach (List<Ponto> g in t.Guias)
            {
                Ponto a = g[0], b = g[g.Count - 1];
                saida.Add(new Ancora { P = t.Distancia(a) >= t.Distancia(b) ? a : b, Guia = g });
            }
            double h = t.Altura > 0 ? t.Altura : ctx.H;
            List<KeyValuePair<double, Ancora>> achados = new List<KeyValuePair<double, Ancora>>();
            foreach (List<Ponto> g in ctx.Guias)
            {
                string lg;
                _layerDaGuia.TryGetValue(g, out lg);
                // guia na MESMA layer do texto (o normal num projeto organizado): pode começar mais longe; outra layer: só encostada
                double toque = (lg != null && string.Equals(lg, t.Layer, StringComparison.OrdinalIgnoreCase) ? ctx.Cfg.GuiaToqueAlturas : ctx.Cfg.RaioAlturas) * h;
                double d0 = t.Distancia(g[0]), d1 = t.Distancia(g[g.Count - 1]);
                if (d0 <= toque && d0 <= d1) achados.Add(new KeyValuePair<double, Ancora>(d0, Seguir(g, true, ctx)));
                else if (d1 <= toque) achados.Add(new KeyValuePair<double, Ancora>(d1, Seguir(g, false, ctx)));
            }
            saida.AddRange(achados.OrderBy(a => a.Key).Select(a => a.Value));
            saida.Add(new Ancora { P = t.Centro });
            if (t.Caixa != null) saida.Add(new Ancora { P = new Ponto(t.X, t.Y) });
            return saida;
        }

        private class Candidato
        {
            public TextoLido T;
            public Elemento E;
            public double Nota;          // menor = melhor
            public double Dist;
            public List<Ponto> Guia;
            public bool Reserva;         // sem guia, por proximidade ampliada
        }

        /// <summary>Nível que a anotação do cabo indica: MT com tensão em kV; BT para multiplex (PB/RAMA); null = não diz.</summary>
        private static string NivelDoTexto(TextoLido tl)
        {
            CaboLido c = Anotacoes.LerCabo(tl.R.Ativo);
            if (!string.IsNullOrEmpty(c.Tensao)) return c.Nivel;
            string cod = (c.Codigo ?? "").ToUpperInvariant();
            if (cod.StartsWith("PB") || cod.StartsWith("RAMA")) return "BT";
            return null;
        }

        private static List<Candidato> Candidatos(TextoLido tl, Contexto ctx, bool cabo, double raio, bool reserva)
        {
            List<Candidato> l = new List<Candidato>();
            string nivel = cabo ? NivelDoTexto(tl) : null;
            List<Ancora> ancoras = Ancoras(tl.T, ctx);
            for (int k = 0; k < ancoras.Count; k++)
            {
                Ancora q = ancoras[k];
                bool viaGuia = q.Guia != null;
                // pela guia, a ponta precisa CAIR no elemento; pelo próprio texto, vale o raio (na reserva, o raio ampliado vale para os dois)
                double alcance = viaGuia && !reserva ? Math.Max(ctx.Cfg.TolPosteAlturas * ctx.H, ctx.H * 2) : raio;
                IEnumerable<Elemento> elems = cabo ? ctx.Vaos.Cast<Elemento>() : ctx.Postes.Where(p => !string.IsNullOrEmpty(p.Handle)).Cast<Elemento>();
                foreach (Elemento e in elems)
                {
                    double dd = cabo ? Geometria.Projetar(q.P, ((Vao)e).Pontos).Distancia : q.P.Dist(((Poste)e).Ponto);
                    if (dd > alcance) continue;
                    double nota = (viaGuia ? 0 : 1000) + k + dd / Math.Max(alcance, 1e-9);
                    if (cabo && nivel != null && ((Vao)e).Tipo != nivel) continue;           // anotação de MT nunca vai para vão de BT (e vice-versa)
                    l.Add(new Candidato { T = tl, E = e, Nota = nota, Dist = dd, Guia = q.Guia, Reserva = reserva });
                }
            }
            return l.OrderBy(c => c.Nota).ToList();
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

        private static void Ignorar(Contexto ctx, TextoLido tl, string motivo)
        {
            ctx.Ignorados.Add(new TextoIgnorado { Handle = tl.T.Handle, Texto = tl.T.Texto, Layer = tl.T.Layer, Cor = tl.T.Cor, Tipo = tl.T.Tipo, X = tl.T.X, Y = tl.T.Y, Motivo = motivo });
        }

        /// <summary>Liga cada anotação a um poste/vão. Devolve as anotações a MAIS de cada elemento (viram "irmãos": V7b, P3b).</summary>
        private static Dictionary<Elemento, List<TextoLido>> AssociarTextos(List<TextoLido> lidos, Contexto ctx, List<Avulso> avulsos)
        {
            ConfigLeitura cfg = ctx.Cfg;
            List<TextoLido> paraLigar = new List<TextoLido>();
            foreach (TextoLido tl in lidos)
            {
                if (tl.T.Tipo == "BLOCO") continue;
                string ent = tl.R.Entidade ?? "0";
                if (ent == "0")
                {
                    Ignorar(ctx, tl, tl.R.Operacao == "0" ? "regra do leitor manda ignorar (operação 0)" : "nenhuma regra de classificação reconheceu o texto");
                    continue;
                }
                if (ent == "RAMAIS" || ent == "IP") { Ignorar(ctx, tl, "ramal/IP (vai para Ramais na aplicação, fora da obra do plugin)"); continue; }
                if (tl.T.Tipo == "BLOCO_DEF") { if (VaiParaOutros(ent) || EhCabo(ent)) avulsos.Add(NovoAvulso(tl, null)); continue; }
                if (EhCabo(ent) || EhPoste(ent)) paraLigar.Add(tl);
                else if (VaiParaOutros(ent)) avulsos.Add(NovoAvulso(tl, null));
            }

            // 1ª passada: cada anotação no melhor candidato (guia primeiro, depois o raio normal)
            Dictionary<Elemento, List<Candidato>> por = new Dictionary<Elemento, List<Candidato>>();
            Dictionary<TextoLido, List<Candidato>> cands = new Dictionary<TextoLido, List<Candidato>>();
            List<TextoLido> sobra = new List<TextoLido>();
            foreach (TextoLido tl in paraLigar)
            {
                double h = tl.T.Altura > 0 ? tl.T.Altura : ctx.H;
                List<Candidato> c = Candidatos(tl, ctx, EhCabo(tl.R.Entidade), cfg.RaioAlturas * h, false);
                cands[tl] = c;
                if (c.Count == 0) { sobra.Add(tl); continue; }
                Lista(por, c[0].E).Add(c[0]);
            }
            // rebalanceia: anotação a mais num elemento vai para um candidato SEM anotação, se ele for quase tão bom
            foreach (Elemento e in por.Keys.ToList())
            {
                List<Candidato> l = por[e].OrderBy(x => x.Nota).ToList();
                for (int i = 1; i < l.Count; i++)
                {
                    Candidato alt = cands[l[i].T].FirstOrDefault(x => x.E != e && !por.ContainsKey(x.E) && x.Nota <= l[i].Nota + 1);
                    if (alt == null) continue;
                    por[e].Remove(l[i]);
                    Lista(por, alt.E).Add(alt);
                }
            }
            // 2ª passada (sem guia e longe): elemento SEM anotação mais perto, até o raio de reserva; vai para Conferir
            foreach (TextoLido tl in sobra)
            {
                double h = tl.T.Altura > 0 ? tl.T.Altura : ctx.H;
                Candidato c = Candidatos(tl, ctx, EhCabo(tl.R.Entidade), cfg.RaioReservaAlturas * h, true).FirstOrDefault(x => !por.ContainsKey(x.E));
                if (c != null) { Lista(por, c.E).Add(c); continue; }
                if (EhCabo(tl.R.Entidade))
                {
                    avulsos.Add(NovoAvulso(tl, "anotação de cabo sem vão por perto (rede nas layers " + string.Join("/", cfg.LayersRedeMT.Concat(cfg.LayersRedeBT).ToArray()) + ")"));
                    continue;
                }
                Poste novo = new Poste();                                       // sem bloco: poste só da anotação (fica sem vãos)
                novo.Handle = ""; novo.X = tl.T.X; novo.Y = tl.T.Y;
                Anotar(novo, tl);
                novo.MensagensFixas.Add("anotação de poste sem bloco de poste por perto (blocos: " + string.Join(", ", cfg.BlocosPoste.ToArray())
                                        + (cfg.LayersPoste.Count > 0 ? "; layers: " + string.Join(", ", cfg.LayersPoste.ToArray()) : "") + ")");
                ctx.Postes.Add(novo);
            }

            Dictionary<Elemento, List<TextoLido>> extras = new Dictionary<Elemento, List<TextoLido>>();
            foreach (KeyValuePair<Elemento, List<Candidato>> kv in por)
            {
                // poste: a anotação de entidade POSTE vem antes da de ESTRUTURA
                List<Candidato> l = kv.Value.OrderBy(x => x.T.R.Entidade == "POSTE" ? 0 : 1).ThenBy(x => x.Nota).ToList();
                Anotar(kv.Key, l[0].T);
                kv.Key.GuiaUsada = l[0].Guia;
                if (l[0].Reserva) kv.Key.MensagensFixas.Add("anotação ligada por proximidade (sem linha guia): confira");
                if (l.Count > 1) extras[kv.Key] = l.Skip(1).Select(x => x.T).ToList();
                for (int i = 1; i < l.Count; i++) _guiaExtra[l[i].T] = l[i].Guia;
            }
            return extras;
        }

        [ThreadStatic] private static Dictionary<TextoLido, List<Ponto>> _guiaExtraTs;
        private static Dictionary<TextoLido, List<Ponto>> _guiaExtra
        {
            get { if (_guiaExtraTs == null) _guiaExtraTs = new Dictionary<TextoLido, List<Ponto>>(); return _guiaExtraTs; }
        }

        private static List<Candidato> Lista(Dictionary<Elemento, List<Candidato>> d, Elemento e)
        {
            List<Candidato> l;
            if (!d.TryGetValue(e, out l)) { l = new List<Candidato>(); d[e] = l; }
            return l;
        }

        private static Avulso NovoAvulso(TextoLido tl, string msg)
        {
            Avulso a = new Avulso();
            Anotar(a, tl);
            if (msg != null) a.MensagensFixas.Add(msg);
            return a;
        }

        /// <summary>Mais de uma anotação no mesmo vão/poste (recondutagem: instala + retira + BT; troca de poste): cada uma vira um
        /// elemento "irmão" na mesma geometria (V7b, P3b) — vai para a obra como uma linha a mais, com vínculo próprio.</summary>
        private static void CriarIrmaos(Obra obra, Dictionary<Elemento, List<TextoLido>> extras)
        {
            foreach (KeyValuePair<Elemento, List<TextoLido>> kv in extras)
            {
                int k = 2;
                foreach (TextoLido tl in kv.Value)
                {
                    Vao v = kv.Key as Vao;
                    Poste p = kv.Key as Poste;
                    Elemento novo;
                    if (v != null)
                    {
                        Vao n = new Vao { LinhaHandle = v.LinhaHandle, Trecho = v.Trecho, Pontos = v.Pontos, Tipo = v.Tipo, LinhaLayer = v.LinhaLayer, LinhaCor = v.LinhaCor,
                                          PosteA = v.PosteA, PosteB = v.PosteB, PosteObjA = v.PosteObjA, PosteObjB = v.PosteObjB };
                        obra.Vaos.Insert(obra.Vaos.IndexOf(v) + k - 1, n);
                        novo = n;
                    }
                    else
                    {
                        Poste n = new Poste { Handle = p.Handle, X = p.X, Y = p.Y, Bloco = p.Bloco };
                        n.OutrosHandles.AddRange(p.OutrosHandles);
                        obra.Postes.Insert(obra.Postes.IndexOf(p) + k - 1, n);
                        novo = n;
                    }
                    Anotar(novo, tl);
                    List<Ponto> g;
                    if (_guiaExtra.TryGetValue(tl, out g)) novo.GuiaUsada = g;
                    novo.Numero = kv.Key.Numero + (char)('a' + k - 1);              // V7b, V7c… (só ASCII: vai na linha de comando)
                    novo.IrmaoDe = kv.Key.Numero;
                    k++;
                }
            }
            _guiaExtra.Clear();
            // anotação repetida (mesmo texto e operação no mesmo vão/poste) é quase sempre desenho duplicado
            foreach (IGrouping<string, Elemento> g in obra.Postes.Cast<Elemento>().Concat(obra.Vaos.Cast<Elemento>())
                         .Where(e => e.TemAnotacao).GroupBy(e => (e.IrmaoDe ?? e.Numero) + "|" + e.Operacao + "|" + MotorLeitor.LimparTexto(e.AtivoLido).ToUpperInvariant()))
                foreach (Elemento e in g.Skip(1)) e.MensagensFixas.Add("anotação repetida no mesmo " + (e is Vao ? "vão" : "poste") + " (" + g.First().Numero + "): desenho duplicado?");
        }

        /// <summary>Placa do poste (texto como B039571, que as regras ignoram) perto do poste ou ligada a ele por guia.</summary>
        private static void LerPlacas(Obra obra, Contexto ctx)
        {
            System.Text.RegularExpressions.Regex re = MotorLeitor.RegexJs(ctx.Cfg.PlacaRegex);
            if (re == null) return;
            List<Poste> comBloco = obra.Postes.Where(p => !string.IsNullOrEmpty(p.Handle) && p.IrmaoDe == null).ToList();
            foreach (TextoIgnorado t in ctx.Ignorados.ToList())
            {
                if (!re.IsMatch((t.Texto ?? "").Trim())) continue;
                TextoLido tl = ctx.Lidos.FirstOrDefault(x => x.T.Handle == t.Handle && x.T.Texto == t.Texto);
                if (tl == null) continue;
                double h = tl.T.Altura > 0 ? tl.T.Altura : ctx.H;
                List<Candidato> c = Candidatos(tl, ctx, false, ctx.Cfg.RaioAlturas * h, false);
                Candidato melhor = c.FirstOrDefault(x => comBloco.Contains((Poste)x.E) && string.IsNullOrEmpty(((Poste)x.E).Placa));
                if (melhor == null) continue;
                ((Poste)melhor.E).Placa = t.Texto.Trim();
                t.Motivo = "placa do poste " + melhor.E.Numero;
            }
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
