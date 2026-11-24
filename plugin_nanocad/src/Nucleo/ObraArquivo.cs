// Nucleo/ObraArquivo.cs — o arquivo `<desenho>.obra.json` (CONTRATO com a aplicação) e a obra guardada no DWG.
//
// Formato = o de exportar/importar obra da aplicação (services/obras_arquivo.py, "leitor-obra" versão 1):
//   {"formato": "leitor-obra", "versao": 1, "exportado_em", "tipo": "obra", "nome", "projeto",
//    "dados": {"cabos": [...], "outros": [...], "totalizadora": []},
//    "cad": {...}}                       ← seção do plugin (a aplicação ignora chaves que não conhece)
// Cada linha de Cabos/Outros tem os campos que a tela de Resumo grava (entidade, operacao, ativo, textoOriginal, _x/_y,
// _geo/_geoTipo/_poste e vinculoEstruturas = ids das unidades de estrutura em Outros: linha + k/100 para a 2ª+ estrutura).
// A seção "cad" guarda a obra do plugin inteira (números, handles, ajustes) para "Continuar obra salva".
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    public static class ObraArquivo
    {
        public const string Formato = "leitor-obra";
        public const int Versao = 1;
        public const int VersaoCad = 1;
        public const string VersaoPlugin = "0.16.4";
        private const int MaxPontosGeo = 60;

        public static string NomeDoArquivo(string desenho)
        {
            if (string.IsNullOrEmpty(desenho)) return null;
            int p = desenho.LastIndexOf('.');
            int barra = Math.Max(desenho.LastIndexOf('\\'), desenho.LastIndexOf('/'));
            string sem = p > barra ? desenho.Substring(0, p) : desenho;
            return sem + ".obra.json";
        }

        private static double R2(double v) { return Math.Round(v, 2); }

        private static List<object> Ponto(double x, double y) { return new List<object> { R2(x), R2(y) }; }

        private static List<object> Geo(List<Ponto> pts)
        {
            List<Ponto> p = pts;
            if (p.Count > MaxPontosGeo)
            {
                double passo = (p.Count - 1) / (double)(MaxPontosGeo - 1);
                p = Enumerable.Range(0, MaxPontosGeo).Select(i => pts[(int)Math.Round(i * passo)]).ToList();
            }
            return p.Select(q => (object)Ponto(q.X, q.Y)).ToList();
        }

        public static double IdUnidade(int linha, int k)
        {
            return k == 0 ? linha : Math.Round(linha + k / 100.0, 2);
        }

        private static JObj LinhaBase(Elemento e)
        {
            JObj l = new JObj();
            l.Set("entidade", string.IsNullOrEmpty(e.Entidade) ? "0" : e.Entidade);
            l.Set("operacao", e.Operacao ?? "");
            MarcasDaLinha(l, e);
            return l;
        }

        private static void MarcasDaLinha(JObj o, Elemento e)
        {
            if (e.Comentario) o.Set("comentario", true);
            if (e.Congelada || e.Comentario) o.Set("congelada", true);
            if (e.Bloqueada) o.Set("bloqueada", true);
        }

        private static void Coordenadas(JObj l, Elemento e)
        {
            if (e.TextoX.HasValue && e.TextoY.HasValue) { l.Set("_x", R2(e.TextoX.Value)); l.Set("_y", R2(e.TextoY.Value)); }
        }

        /// <summary>Os elementos de cada linha de Cabos e de Outros, na ordem em que Dados() escreve (para casar respostas da aplicação).</summary>
        public static void ElementosDasLinhas(Obra obra, out List<Elemento> cabos, out List<Elemento> outros)
        {
            cabos = new List<Elemento>();
            outros = new List<Elemento>();
            foreach (Poste p in obra.Postes) if (p.TemAnotacao && !p.Excluido) outros.Add(p);
            foreach (Vao v in obra.Vaos) if (v.TemAnotacao && !v.Excluido) cabos.Add(v);
            foreach (Avulso a in obra.Avulsos) if (!a.Excluido) (a.Entidade == "CABO" ? cabos : outros).Add(a);
            cabos = cabos.OrderBy(x => x.OrdemLeitor ?? int.MaxValue).ToList();
            outros = outros.OrderBy(x => x.OrdemLeitor ?? int.MaxValue).ToList();
        }

        private static JObj LinhaAnotacao(Elemento e)
        {
            JObj linha = LinhaBase(e);
            Vao v = e as Vao;
            linha.Set("ativo", v == null ? e.AtivoLido : v.Ativo);
            linha.Set("textoOriginal", e.TextoOriginal);
            Coordenadas(linha, e);
            Poste p = e as Poste;
            if (p != null && !string.IsNullOrEmpty(p.Handle)) linha.Set("_poste", Ponto(p.X, p.Y));
            else if (p != null && (p.X != 0 || p.Y != 0)) linha.Set("_posteAprox", Ponto(p.X, p.Y));     // sem bloco: a posição da anotação, só para o mapa da aplicação
            return linha;
        }

        public static JObj Dados(Obra obra)
        {
            List<Elemento> elementosCabos, elementosOutros;
            ElementosDasLinhas(obra, out elementosCabos, out elementosOutros);
            Dictionary<string, int> linhas = new Dictionary<string, int>();
            for (int i = 0; i < elementosOutros.Count; i++)
                if (elementosOutros[i] is Poste) linhas[elementosOutros[i].Numero] = i;

            List<object> cabos = new List<object>();
            foreach (Elemento e in elementosCabos)
            {
                JObj l = LinhaAnotacao(e);
                cabos.Add(l);
                Vao v = e as Vao;
                if (v == null) continue;
                if (v.Pontos.Count >= 2) { l.Set("_geo", Geo(v.Pontos)); if (!string.IsNullOrEmpty(v.Tipo)) l.Set("_geoTipo", v.Tipo); }
                if (obra.Modo == "construir")
                {
                    var pontas = new[] { v.PosteA, v.PosteB }.Where(p => !string.IsNullOrEmpty(p) && linhas.ContainsKey(p))
                        .Select(p => (object)linhas[p]).Distinct().ToList();
                    if (pontas.Count > 0) l.Set("vinculoPostes", pontas);
                }
                List<object> ids = new List<object>();
                foreach (Vinculo vi in v.Vinculos)
                    if (linhas.ContainsKey(vi.Poste)) ids.Add(IdUnidade(linhas[vi.Poste], vi.K));
                if (v.ParticipaVinculo)
                {
                    l.Set("vinculoEstruturas", ids);
                    if (ids.Count > 0 && v.Vinculos.Any(x => x.Manual)) l.Set("vinculoOrigem", "manual");
                }
            }
            JObj d = new JObj();
            d.Set("cabos", cabos);
            d.Set("outros", elementosOutros.Select(e => (object)LinhaAnotacao(e)).ToList());
            // TASK-038: pontos e cabos montados no formato do mapa (vão depois das linhas lidas)
            ((List<object>)d["outros"]).AddRange(obra.PontosMapa.SelectMany(MapaManual.LinhasDoPonto));
            cabos.AddRange(obra.CabosMapa.SelectMany(MapaManual.LinhasDoCabo));
            d.Set("totalizadora", new List<object>());
            return d;
        }

        public static JObj Envelope(Obra obra, DateTime agoraUtc)
        {
            JObj env = new JObj();
            env.Set("formato", Formato);
            env.Set("versao", Versao);
            env.Set("exportado_em", agoraUtc.ToString("yyyy-MM-dd'T'HH:mm:ss'+00:00'", CultureInfo.InvariantCulture));
            env.Set("tipo", "obra");
            env.Set("nome", string.IsNullOrEmpty(obra.Nome) ? "obra" : (obra.Nome.Length > 120 ? obra.Nome.Substring(0, 120) : obra.Nome));
            env.Set("projeto", obra.Projeto ?? "");
            env.Set("dados", Dados(obra));
            env.Set("cad", Cad(obra));
            return env;
        }

        public static string Gerar(Obra obra, DateTime agoraUtc)
        {
            return Json.Escrever(Envelope(obra, agoraUtc), true);
        }

        // ── seção "cad": a obra do plugin inteira ──
        private static JObj ElementoParaJson(Elemento e)
        {
            JObj o = new JObj();
            o.Set("numero", e.Numero);
            MarcasDaLinha(o, e);
            if (e.FontesMontagem.Count > 0) o.Set("fontesMontagem", e.FontesMontagem.Cast<object>().ToList());
            if (e.OrigemMetros != null)
            {
                TextoDesenho t = e.OrigemMetros;
                o.Set("origemMetros", new JObj().Set("h",t.Handle).Set("texto",t.Texto).Set("layer",t.Layer).Set("x",t.X).Set("y",t.Y));
            }
            if (e.GuiaUsada != null && e.GuiaUsada.Count >= 2) o.Set("guia", Geo(e.GuiaUsada));
            if (e.OrdemLeitor.HasValue) o.Set("ordemLeitor", e.OrdemLeitor.Value);
            if (e.ParteTexto != 0) o.Set("parteTexto", e.ParteTexto);
            if (e.Candidatos.Count > 1) o.Set("candidatos", e.Candidatos.Cast<object>().ToList());
            o.Set("texto", e.TextoHandle);
            o.Set("textoOriginal", e.TextoOriginal);
            o.Set("textoLayer", e.TextoLayer);
            if (e.LayerAplicada != null) o.Set("layerAplicada",e.LayerAplicada);
            o.Set("textoCor", e.TextoCor);
            if (e.TextoX.HasValue) { o.Set("textoX", e.TextoX.Value); o.Set("textoY", e.TextoY.Value); }
            o.Set("entidade", e.Entidade);
            o.Set("ativoLido", e.AtivoLido);
            o.Set("operacaoLida", e.OperacaoLida);
            o.Set("operacao", e.Operacao);
            if (e.Manual) o.Set("manual", true);
            if (e.AtivoDoDesenho != null) o.Set("ativoDoDesenho", e.AtivoDoDesenho);
            if (e.Excluido) o.Set("excluido", true);
            if (e.IrmaoDe != null) o.Set("irmaoDe", e.IrmaoDe);
            if (e.MensagensFixas.Count > 0) o.Set("avisosLeitura", e.MensagensFixas.Cast<object>().ToList());
            o.Set("estado", e.Estado);
            if (e.Mensagens.Count > 0) o.Set("mensagens", e.Mensagens.Cast<object>().ToList());
            return o;
        }

        private static JObj VinculoParaJson(Vinculo v)
        {
            if (v == null) return null;
            JObj o = new JObj();
            o.Set("poste", v.Poste); o.Set("estrutura", v.Estrutura); o.Set("k", v.K);
            if (v.Manual) o.Set("manual", true);
            return o;
        }

        public static JObj Cad(Obra obra)
        {
            JObj c = new JObj();
            c.Set("versao", VersaoCad);
            if (obra.Conferidos.Count > 0) c.Set("conferidos", obra.Conferidos);
            if (obra.Associacoes.Count > 0) c.Set("associacoes", obra.Associacoes);
            c.Set("origem", "plugin-nanocad");
            c.Set("plugin", VersaoPlugin);
            c.Set("desenho", obra.Desenho ?? "");
            c.Set("regras", obra.NomeRegras ?? "");
            if (obra.FonteDxf != null) c.Set("fonte_dxf", obra.FonteDxf);
            if (obra.HashLeitor != null) c.Set("regras_leitor", obra.HashLeitor);
            c.Set("modo", obra.Modo);
            c.Set("altura_texto", obra.AlturaTexto);
            if (obra.SalvoEm != null) c.Set("salvo_em", obra.SalvoEm);
            List<object> postes = new List<object>();
            foreach (Poste p in obra.Postes)
            {
                JObj o = ElementoParaJson(p);
                o.Set("handle", p.Handle ?? "");
                o.Set("x", p.X); o.Set("y", p.Y);
                o.Set("bloco", p.Bloco ?? "");
                if (p.OutrosHandles.Count > 0) o.Set("outrosBlocos", p.OutrosHandles.Cast<object>().ToList());
                if (!string.IsNullOrEmpty(p.Placa)) o.Set("placa", p.Placa);
                // só informativo (recalculado ao abrir): o que cada estrutura recebeu do vínculo — conferido contra a aplicação no teste de contrato
                List<object> estr = new List<object>();
                for (int k = 0; k < p.Lido.Estruturas.Count; k++)
                {
                    JObj e = new JObj();
                    e.Set("nome", p.Lido.Estruturas[k].Nome);
                    e.Set("cabos", k < p.CabosPorEstrutura.Count ? p.CabosPorEstrutura[k] : 0);
                    JObj pars = new JObj();
                    if (k < p.ParametrosEstruturas.Count)
                        foreach (KeyValuePair<string, string> kv in p.ParametrosEstruturas[k])
                        {
                            List<object> l = pars[kv.Key] as List<object>;
                            if (l == null) { l = new List<object>(); pars.Set(kv.Key, l); }
                            l.Add(kv.Value);
                        }
                    e.Set("parametros", pars);
                    estr.Add(e);
                }
                o.Set("estruturas", estr);
                postes.Add(o);
            }
            List<object> vaos = new List<object>();
            foreach (Vao v in obra.Vaos)
            {
                JObj o = ElementoParaJson(v);
                o.Set("linha", v.LinhaHandle ?? "");
                o.Set("trecho", v.Trecho);
                o.Set("pontos", v.Pontos.Select(q => (object)new List<object> { q.X, q.Y }).ToList());
                o.Set("tipo", v.Tipo ?? "");
                o.Set("linhaLayer", v.LinhaLayer ?? "");
                o.Set("linhaCor", v.LinhaCor ?? "");
                o.Set("posteA", v.PosteA);
                o.Set("posteB", v.PosteB);
                if (v.MetrosLidos.HasValue) o.Set("metrosLidos", v.MetrosLidos.Value);
                if (v.Metros.HasValue) o.Set("metros", v.Metros.Value);
                o.Set("vinculoA", VinculoParaJson(v.VinculoA));
                o.Set("vinculoB", VinculoParaJson(v.VinculoB));
                if (v.MotivoVinculo != null) o.Set("motivoVinculo", v.MotivoVinculo);
                vaos.Add(o);
            }
            List<object> avulsos = obra.Avulsos.Select(a =>
            {
                JObj o = ElementoParaJson(a);
                if (a.PertoDe != null) o.Set("pertoDe", a.PertoDe.Numero);
                return (object)o;
            }).ToList();
            c.Set("postes", postes);
            c.Set("vaos", vaos);
            c.Set("avulsos", avulsos);
            if (obra.PontosMapa.Count > 0) c.Set("pontosMapa", MapaManual.PontosParaJson(obra.PontosMapa));
            if (obra.CabosMapa.Count > 0) c.Set("cabosMapa", MapaManual.CabosParaJson(obra.CabosMapa));
            if (obra.SemRegras.HasValue) c.Set("semRegras", obra.SemRegras.Value);
            return c;
        }

        // ── leitura (Continuar obra salva) ──
        public class ErroArquivo : Exception { public ErroArquivo(string m) : base(m) { } }

        private static void ElementoDeJson(Elemento e, JObj o)
        {
            JObj metros = o.Obj("origemMetros");
            if (metros != null) e.OrigemMetros = new TextoDesenho { Handle = metros.Str("h"), Texto = metros.Str("texto"),
                Layer = metros.Str("layer"), X = metros.Num("x") ?? 0, Y = metros.Num("y") ?? 0 };
            if (o.Lista("guia") != null)
                e.GuiaUsada = o.Lista("guia").OfType<List<object>>().Where(x => x.Count >= 2)
                    .Select(x => new Ponto(Json.ComoNumero(x[0]) ?? 0, Json.ComoNumero(x[1]) ?? 0)).ToList();
            e.OrdemLeitor = (int?)o.Num("ordemLeitor");
            e.ParteTexto = (int)(o.Num("parteTexto") ?? 0);
            e.Candidatos = Json.ListaDeTextos(o["candidatos"]);
            e.Numero = o.Str("numero");
            e.Comentario = o.Bool("comentario", false);
            e.Congelada = o.Bool("congelada", false);
            e.Bloqueada = o.Bool("bloqueada", false);
            e.FontesMontagem = Json.ListaDeTextos(o["fontesMontagem"]);
            e.TextoHandle = o.Str("texto");
            e.TextoOriginal = o.Str("textoOriginal") ?? "";
            e.LayerAplicada = o.Str("layerAplicada");
            e.TextoLayer = o.Str("textoLayer") ?? "";
            e.TextoCor = o.Str("textoCor") ?? "";
            e.TextoX = o.Num("textoX");
            e.TextoY = o.Num("textoY");
            e.Entidade = o.Str("entidade") ?? "0";
            e.AtivoLido = o.Str("ativoLido") ?? "";
            e.OperacaoLida = o.Str("operacaoLida") ?? "";
            e.Operacao = o.Str("operacao") ?? e.OperacaoLida;
            e.Manual = o.Bool("manual", false);
            e.AtivoDoDesenho = o.Str("ativoDoDesenho");
            e.Excluido = o.Bool("excluido", false);
            e.IrmaoDe = o.Str("irmaoDe");
            e.MensagensFixas = Json.ListaDeTextos(o["avisosLeitura"]);
        }

        private static Vinculo VinculoDeJson(object v)
        {
            JObj o = v as JObj;
            if (o == null) return null;
            Vinculo r = new Vinculo();
            r.Poste = o.Str("poste"); r.Estrutura = o.Str("estrutura"); r.K = (int)(o.Num("k") ?? 0); r.Manual = o.Bool("manual", false);
            return r;
        }

        /// <summary>Envelope (texto do .obra.json ou do DWG) → obra do plugin. Recalcula parâmetros e conferência com as regras dadas.</summary>
        public static Obra Ler(string json, RegrasProjeto regras)
        {
            JObj env = Json.Ler(json) as JObj;
            if (env == null || env.Str("formato") != Formato) throw new ErroArquivo("não é um arquivo de obra (formato desconhecido)");
            JObj cad = env.Obj("cad");
            if (cad == null) throw new ErroArquivo("a obra não foi gerada pelo plugin (sem a seção 'cad')");
            if ((cad.Num("versao") ?? 0) > VersaoCad) throw new ErroArquivo("obra de uma versão mais nova do plugin; atualize o plugin");
            Obra obra = new Obra();
            obra.Conferidos = cad.Obj("conferidos") ?? new JObj();
            obra.Associacoes = cad.Obj("associacoes") ?? new JObj();
            obra.Nome = env.Str("nome") ?? "obra";
            obra.Projeto = env.Str("projeto") ?? "";
            obra.Desenho = cad.Str("desenho") ?? "";
            obra.NomeRegras = cad.Str("regras") ?? "";
            obra.FonteDxf = cad.Str("fonte_dxf");
            obra.HashLeitor = cad.Str("regras_leitor");
            obra.Modo = cad.Str("modo") ?? "ler";
            obra.PontosMapa = MapaManual.PontosDeJson(cad["pontosMapa"]);
            obra.CabosMapa = MapaManual.CabosDeJson(cad["cabosMapa"]);
            if (cad["semRegras"] is bool) obra.SemRegras = (bool)cad["semRegras"];
            obra.SalvoEm = cad.Str("salvo_em");
            obra.AlturaTexto = cad.Num("altura_texto") ?? 1;
            ProjetoArquivo.LerTrabalho(env, obra);           // TASK-020: nome e número de obra do projeto de trabalho (.obra)
            foreach (JObj o in (cad.Lista("postes") ?? new List<object>()).OfType<JObj>())
            {
                Poste p = new Poste();
                ElementoDeJson(p, o);
                p.Handle = o.Str("handle") ?? "";
                p.X = o.Num("x") ?? 0; p.Y = o.Num("y") ?? 0;
                p.Bloco = o.Str("bloco") ?? "";
                p.OutrosHandles = Json.ListaDeTextos(o["outrosBlocos"]);
                p.Placa = o.Str("placa");
                p.Lido = Anotacoes.LerPoste(p.AtivoLido);
                obra.Postes.Add(p);
            }
            foreach (JObj o in (cad.Lista("vaos") ?? new List<object>()).OfType<JObj>())
            {
                Vao v = new Vao();
                ElementoDeJson(v, o);
                v.LinhaHandle = o.Str("linha") ?? "";
                v.Trecho = (int)(o.Num("trecho") ?? 0);
                foreach (object pt in o.Lista("pontos") ?? new List<object>())
                {
                    List<object> xy = pt as List<object>;
                    if (xy != null && xy.Count >= 2) v.Pontos.Add(new Ponto(Json.ComoNumero(xy[0]) ?? 0, Json.ComoNumero(xy[1]) ?? 0));
                }
                v.Tipo = o.Str("tipo") ?? "";
                v.LinhaLayer = o.Str("linhaLayer") ?? "";
                v.LinhaCor = o.Str("linhaCor") ?? "";
                v.PosteA = o.Str("posteA");
                v.PosteB = o.Str("posteB");
                v.MetrosLidos = o.Num("metrosLidos");
                v.Metros = o.Num("metros");
                v.Cabo = v.TemAnotacao ? Anotacoes.LerCabo(v.AtivoLido) : null;
                v.VinculoA = VinculoDeJson(o["vinculoA"]);
                v.VinculoB = VinculoDeJson(o["vinculoB"]);
                v.MotivoVinculo = o.Str("motivoVinculo");
                obra.Vaos.Add(v);
            }
            foreach (JObj o in (cad.Lista("avulsos") ?? new List<object>()).OfType<JObj>())
            {
                Avulso a = new Avulso();
                ElementoDeJson(a, o);
                a.PertoDe = obra.PorNumero(o.Str("pertoDe"));
                obra.Avulsos.Add(a);
            }
            Vinculador.AtualizarParametros(obra, regras);
            Conferencia.Avaliar(obra, regras);
            return obra;
        }
    }
}
