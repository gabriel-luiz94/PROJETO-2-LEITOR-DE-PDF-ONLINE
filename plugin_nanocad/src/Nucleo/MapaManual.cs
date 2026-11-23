// Nucleo/MapaManual.cs — TASK-038: montagem manual no formato do mapa do OrçaCore (TASK-235/236 da aplicação).
//
// O usuário seleciona o POSTE (bloco/círculo ou nenhum) e as anotações (vários TEXT/MTEXT de uma vez); cada parágrafo vira uma linha L e,
// em cada linha, um ou mais CAMPOS com a operação (F, I, *I, M, *M, R, *R) — como a janela do «anota poste». Com regras de projeto, o ativo e a
// operação vêm das regras do Leitor; texto selecionado que nenhuma regra reconhece entra como F (fictício). Sem regras, o texto vai como está.
// O CABO é a linha inteira ou um trecho dela (dois pontos sobre a linha); o comprimento é o anotado ou, sem anotação, o medido.
//
// Saída para a aplicação (mesmas linhas que o Resumo grava): o ponto = grupo «@nome - comentário» em Outros, um membro por campo
// (operação F = congelada, `_congeladaPorF`), com `_poste` (símbolo do poste) ou `_posteAprox` (sem símbolo: posição da anotação);
// o cabo = uma linha por campo em Cabos, com `_geo`/`_geoTipo` e a metragem. O vínculo cabo ↔ estrutura é feito depois, no OrçaCore.
// C# 5 (o mesmo código compila para o nanoCAD 5, .NET 3.5).
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace OrcamentoNanoCAD.Nucleo
{
    public class CampoMapa
    {
        public int L = 1;                 // linha do texto (como no «anota poste»)
        public string Texto = "";         // o ativo desta linha da tabela
        public string Operacao = "I";     // F | I | *I | M | *M | R | *R
        public string Origem = "";        // regra | sem regra | manual
        public string Entidade;           // da regra (POSTE, ESTRUTURA…); vazio = decide pelo texto
    }

    public class PontoMapa
    {
        public string Numero = "";        // P1, P2… (nome automático)
        public string Nome = "";          // nome do ponto (vazio = o número)
        public string Comentario = "";
        public string Handle = "";        // bloco/círculo do poste (vazio = sem símbolo)
        public string Bloco = "";
        public double X, Y;               // símbolo do poste; sem símbolo, a posição da anotação
        public List<CampoMapa> Campos = new List<CampoMapa>();
        public List<string> Textos = new List<string>();   // handles das anotações usadas
        public bool Enviado;
        public string NomeFinal { get { return string.IsNullOrEmpty((Nome ?? "").Trim()) ? Numero : Nome.Trim(); } }
    }

    public class CaboMapa
    {
        public string Numero = "";        // C1, C2…
        public string LinhaHandle = "";
        public bool Trecho;
        public List<Ponto> Pontos = new List<Ponto>();
        public string Tipo = "";          // MT | BT (pela layer/cor da linha)
        public double Medido;             // comprimento da linha/trecho no desenho
        public double? Metros;            // o que vai para o orçamento (anotado ou medido)
        public List<CampoMapa> Campos = new List<CampoMapa>();
        public List<string> Textos = new List<string>();
        public bool Enviado;
    }

    public static class MapaManual
    {
        public static readonly string[] Operacoes = { "F", "I", "*I", "M", "*M", "R", "*R" };
        private static readonly Regex Metragem = new Regex(@"\s+(\d+(?:[.,]\d+)?)\s*[mM]?\s*$");
        private const int MaxPontosGeo = 60;

        public static string NormalizarOperacao(string op)
        {
            string o = (op ?? "").Trim().ToUpperInvariant();
            return Operacoes.Contains(o) ? o : null;
        }

        /// <summary>Ordem de leitura das anotações (de cima para baixo, da esquerda para a direita): a linha L de cada uma.</summary>
        public static List<TextoDesenho> OrdemDeLeitura(IEnumerable<TextoDesenho> textos)
        {
            return (textos ?? Enumerable.Empty<TextoDesenho>()).Where(t => t != null && !string.IsNullOrEmpty((t.Texto ?? "").Trim()))
                .OrderByDescending(t => Math.Round(t.Y, 3)).ThenBy(t => t.X).ThenBy(t => t.Parte).ToList();
        }

        /// <summary>
        /// Campos iniciais da janela: um por texto (parágrafo), na linha L dele. `regras` = null → sem regras (texto como está, operação `opPadrao`).
        /// Com regras: ativo/operação da regra; texto selecionado sem regra (ou que a regra manda ignorar) → F.
        /// </summary>
        public static List<CampoMapa> CamposDosTextos(IEnumerable<TextoDesenho> textos, RegrasProjeto regras, string opPadrao, bool cabo)
        {
            List<TextoDesenho> ordem = OrdemDeLeitura(textos);
            var saida = new List<CampoMapa>();
            if (ordem.Count == 0) return saida;
            string padrao = NormalizarOperacao(opPadrao) ?? "I";
            var porTexto = new Dictionary<TextoDesenho, TextoLido>();
            if (regras != null)
                foreach (TextoLido tl in LeitorProjeto.ClassificarTextos(ordem, regras))
                    if (tl != null && tl.T != null && !porTexto.ContainsKey(tl.T)) porTexto[tl.T] = tl;
            for (int i = 0; i < ordem.Count; i++)
            {
                TextoDesenho t = ordem[i];
                var campo = new CampoMapa { L = i + 1 };
                TextoLido tl;
                string ativo = null, op = null;
                if (regras != null && porTexto.TryGetValue(t, out tl) && tl.R != null)
                {
                    ativo = (tl.R.Ativo ?? "").Trim();
                    op = NormalizarOperacao(tl.R.Operacao);
                    string ent = (tl.R.Entidade ?? "").Trim();
                    if (ent.Length > 0 && ent != "0") campo.Entidade = ent;
                }
                if (regras != null && (string.IsNullOrEmpty(ativo) || op == null || op == "F"))
                {
                    campo.Texto = Limpar(t.Texto); campo.Operacao = "F"; campo.Origem = "sem regra";
                }
                else if (regras != null)
                {
                    campo.Texto = ativo; campo.Operacao = op; campo.Origem = "regra";
                }
                else
                {
                    campo.Texto = Limpar(t.Texto); campo.Operacao = padrao; campo.Origem = "manual";
                }
                if (cabo) campo.Texto = SemMetragem(campo.Texto);
                if (campo.Texto.Length > 0) saida.Add(campo);
            }
            return saida;
        }

        public static string Limpar(string texto)
        {
            return Regex.Replace((texto ?? "").Replace("\\P", " "), @"\s+", " ").Trim().ToUpperInvariant();
        }

        /// <summary>A metragem escrita no fim do rótulo do cabo («… 38 m» → 38), ou null.</summary>
        public static double? MetrosDoTexto(string texto)
        {
            Match m = Metragem.Match(texto ?? "");
            if (!m.Success) return null;
            double v;
            return double.TryParse(m.Groups[1].Value.Replace(',', '.'), NumberStyles.Float, CultureInfo.InvariantCulture, out v) && v > 0 ? (double?)v : null;
        }

        public static string SemMetragem(string texto)
        {
            string t = (texto ?? "").Trim();
            return MetrosDoTexto(t).HasValue ? Metragem.Replace(t, "").Trim() : t;
        }

        public static string FormatarMetros(double m)
        {
            double r = Math.Round(m, 1);
            return Math.Abs(r - Math.Round(r)) < 1e-9 ? ((long)Math.Round(r)).ToString(CultureInfo.InvariantCulture)
                : r.ToString("0.0", CultureInfo.InvariantCulture).Replace('.', ',');
        }

        public static double Comprimento(IList<Ponto> pts)
        {
            double s = 0;
            for (int i = 1; i < (pts == null ? 0 : pts.Count); i++) s += pts[i - 1].Dist(pts[i]);
            return s;
        }

        /// <summary>Posição (comprimento percorrido) do ponto mais perto sobre a polilinha.</summary>
        public static double Estacao(IList<Ponto> pts, Ponto p)
        {
            double melhor = double.MaxValue, est = 0, acum = 0;
            for (int i = 1; i < pts.Count; i++)
            {
                Ponto a = pts[i - 1], b = pts[i];
                double dx = b.X - a.X, dy = b.Y - a.Y, l2 = dx * dx + dy * dy, seg = Math.Sqrt(l2);
                double t = l2 <= 0 ? 0 : Math.Max(0, Math.Min(1, ((p.X - a.X) * dx + (p.Y - a.Y) * dy) / l2));
                double d = new Ponto(a.X + t * dx, a.Y + t * dy).Dist(p);
                if (d < melhor) { melhor = d; est = acum + t * seg; }
                acum += seg;
            }
            return est;
        }

        private static Ponto NaEstacao(IList<Ponto> pts, double est)
        {
            double acum = 0;
            for (int i = 1; i < pts.Count; i++)
            {
                double seg = pts[i - 1].Dist(pts[i]);
                if (acum + seg >= est || i == pts.Count - 1)
                {
                    double t = seg <= 0 ? 0 : Math.Max(0, Math.Min(1, (est - acum) / seg));
                    return new Ponto(pts[i - 1].X + t * (pts[i].X - pts[i - 1].X), pts[i - 1].Y + t * (pts[i].Y - pts[i - 1].Y));
                }
                acum += seg;
            }
            return pts[pts.Count - 1];
        }

        /// <summary>O trecho da polilinha entre os dois pontos (projetados sobre ela), no sentido da linha.</summary>
        public static List<Ponto> Trecho(IList<Ponto> pts, Ponto a, Ponto b)
        {
            if (pts == null || pts.Count < 2) return new List<Ponto>();
            double ea = Estacao(pts, a), eb = Estacao(pts, b);
            if (ea > eb) { double x = ea; ea = eb; eb = x; }
            var saida = new List<Ponto> { NaEstacao(pts, ea) };
            double acum = 0;
            for (int i = 1; i < pts.Count; i++)
            {
                acum += pts[i - 1].Dist(pts[i]);
                if (acum > ea + 1e-9 && acum < eb - 1e-9) saida.Add(pts[i]);
            }
            saida.Add(NaEstacao(pts, eb));
            return saida;
        }

        public static string ProximoNumero(IEnumerable<string> usados, string prefixo)
        {
            int max = 0;
            foreach (string u in usados ?? Enumerable.Empty<string>())
            {
                int n;
                if (u != null && u.StartsWith(prefixo) && int.TryParse(u.Substring(prefixo.Length), out n)) max = Math.Max(max, n);
            }
            return prefixo + (max + 1);
        }

        private static List<object> Par(double x, double y) { return new List<object> { Math.Round(x, 2), Math.Round(y, 2) }; }

        private static List<object> Geo(List<Ponto> pts)
        {
            List<Ponto> p = pts;
            if (p.Count > MaxPontosGeo)
            {
                double passo = (p.Count - 1) / (double)(MaxPontosGeo - 1);
                p = Enumerable.Range(0, MaxPontosGeo).Select(i => pts[(int)Math.Round(i * passo)]).ToList();
            }
            return p.Select(q => (object)Par(q.X, q.Y)).ToList();
        }

        private static void Operacao(JObj l, string op)
        {
            l.Set("operacao", op);
            if (op == "F") { l.Set("congelada", true); l.Set("_congeladaPorF", true); }
        }

        /// <summary>Linhas de Outros do ponto: o título «@nome - comentário» e um membro por campo (`gid` «plug-NOME» só liga o grupo; a aplicação dá outro ao receber).</summary>
        public static List<object> LinhasDoPonto(PontoMapa p)
        {
            var linhas = new List<object>();
            var campos = p.Campos.Where(c => !string.IsNullOrEmpty((c.Texto ?? "").Trim())).ToList();
            if (campos.Count == 0) return linhas;
            string gid = "plug-" + Regex.Replace(p.NomeFinal ?? p.Numero ?? "", @"[^A-Za-z0-9_-]", "").ToUpperInvariant();
            if (gid.Length > 60) gid = gid.Substring(0, 60);
            string comentario = Regex.Replace(p.Comentario ?? "", @"\s+", " ").Trim();
            string op0 = campos[0].Operacao == "F" ? "I" : campos[0].Operacao;
            linhas.Add(new JObj().Set("entidade", "0").Set("operacao", op0).Set("ativo", "@" + p.NomeFinal.ToUpperInvariant() + (comentario.Length > 0 ? " - " + comentario : ""))
                .Set("grupoDecl", true).Set("congelada", true).Set("gid", gid));
            foreach (CampoMapa c in campos)
            {
                string texto = c.Texto.Trim().ToUpperInvariant();
                string ent = !string.IsNullOrEmpty(c.Entidade) && c.Entidade != "CABO" ? c.Entidade
                    : (!string.IsNullOrEmpty(Anotacoes.LerPoste(texto).CodigoPoste) ? "POSTE" : "ESTRUTURA");
                JObj l = new JObj().Set("entidade", ent).Set("ativo", texto).Set("textoOriginal", texto).Set("entidadeManual", true).Set("grupoDe", gid);
                Operacao(l, c.Operacao);
                l.Set(string.IsNullOrEmpty(p.Handle) ? "_posteAprox" : "_poste", Par(p.X, p.Y));
                linhas.Add(l);
            }
            return linhas;
        }

        /// <summary>Linhas de Cabos: uma por campo, todas com a mesma geometria (a linha/trecho) e a metragem do cabo.</summary>
        public static List<object> LinhasDoCabo(CaboMapa c)
        {
            var linhas = new List<object>();
            double? m = c.Metros ?? (c.Medido > 0 ? (double?)c.Medido : null);
            foreach (CampoMapa campo in c.Campos.Where(x => !string.IsNullOrEmpty((x.Texto ?? "").Trim())))
            {
                string rotulo = SemMetragem(campo.Texto.Trim());
                string ativo = m.HasValue ? rotulo + " " + FormatarMetros(m.Value) + " m" : rotulo;
                JObj l = new JObj().Set("entidade", "CABO").Set("ativo", ativo).Set("textoOriginal", ativo).Set("entidadeManual", true);
                Operacao(l, campo.Operacao);
                if (c.Pontos.Count >= 2) { l.Set("_geo", Geo(c.Pontos)); if (!string.IsNullOrEmpty(c.Tipo)) l.Set("_geoTipo", c.Tipo); }
                linhas.Add(l);
            }
            return linhas;
        }

        // ── seção "cad" (Continuar obra salva) ──
        private static JObj CampoParaJson(CampoMapa c)
        {
            JObj o = new JObj().Set("l", c.L).Set("texto", c.Texto).Set("operacao", c.Operacao).Set("origem", c.Origem);
            if (!string.IsNullOrEmpty(c.Entidade)) o.Set("entidade", c.Entidade);
            return o;
        }

        private static List<CampoMapa> CamposDeJson(object v)
        {
            var saida = new List<CampoMapa>();
            List<object> l = v as List<object>;
            if (l == null) return saida;
            foreach (JObj o in l.OfType<JObj>())
                saida.Add(new CampoMapa { L = (int)(o.Num("l") ?? 1), Texto = o.Str("texto") ?? "", Operacao = NormalizarOperacao(o.Str("operacao")) ?? "I",
                    Origem = o.Str("origem") ?? "", Entidade = o.Str("entidade") });
            return saida;
        }

        public static List<object> PontosParaJson(IEnumerable<PontoMapa> pontos)
        {
            return pontos.Select(p => (object)new JObj().Set("numero", p.Numero).Set("nome", p.Nome).Set("comentario", p.Comentario).Set("handle", p.Handle)
                .Set("bloco", p.Bloco).Set("x", p.X).Set("y", p.Y).Set("enviado", p.Enviado)
                .Set("textos", p.Textos.Cast<object>().ToList()).Set("campos", p.Campos.Select(c => (object)CampoParaJson(c)).ToList())).ToList();
        }

        public static List<object> CabosParaJson(IEnumerable<CaboMapa> cabos)
        {
            return cabos.Select(c =>
            {
                JObj o = new JObj().Set("numero", c.Numero).Set("linha", c.LinhaHandle).Set("trecho", c.Trecho).Set("tipo", c.Tipo).Set("medido", c.Medido)
                    .Set("enviado", c.Enviado).Set("pontos", c.Pontos.Select(q => (object)new List<object> { q.X, q.Y }).ToList())
                    .Set("textos", c.Textos.Cast<object>().ToList()).Set("campos", c.Campos.Select(x => (object)CampoParaJson(x)).ToList());
                if (c.Metros.HasValue) o.Set("metros", c.Metros.Value);
                return (object)o;
            }).ToList();
        }

        public static List<PontoMapa> PontosDeJson(object v)
        {
            var saida = new List<PontoMapa>();
            List<object> l = v as List<object>;
            if (l == null) return saida;
            foreach (JObj o in l.OfType<JObj>())
                saida.Add(new PontoMapa { Numero = o.Str("numero") ?? "", Nome = o.Str("nome") ?? "", Comentario = o.Str("comentario") ?? "", Handle = o.Str("handle") ?? "",
                    Bloco = o.Str("bloco") ?? "", X = o.Num("x") ?? 0, Y = o.Num("y") ?? 0, Enviado = o.Bool("enviado", false),
                    Textos = Json.ListaDeTextos(o["textos"]), Campos = CamposDeJson(o["campos"]) });
            return saida;
        }

        public static List<CaboMapa> CabosDeJson(object v)
        {
            var saida = new List<CaboMapa>();
            List<object> l = v as List<object>;
            if (l == null) return saida;
            foreach (JObj o in l.OfType<JObj>())
            {
                var pts = (o.Lista("pontos") ?? new List<object>()).OfType<List<object>>().Where(x => x.Count >= 2)
                    .Select(x => new Ponto(Json.ComoNumero(x[0]) ?? 0, Json.ComoNumero(x[1]) ?? 0)).ToList();
                saida.Add(new CaboMapa { Numero = o.Str("numero") ?? "", LinhaHandle = o.Str("linha") ?? "", Trecho = o.Bool("trecho", false), Tipo = o.Str("tipo") ?? "",
                    Medido = o.Num("medido") ?? 0, Metros = o.Num("metros"), Enviado = o.Bool("enviado", false), Pontos = pts,
                    Textos = Json.ListaDeTextos(o["textos"]), Campos = CamposDeJson(o["campos"]) });
            }
            return saida;
        }
    }
}
