// Nucleo/Autocompletar.cs — o autocompletar do ativo, como na aplicação (TASK-017 do plugin / TASK-144):
//   • sugestões de nomes da Base de Orçamento do projeto (com descrição; apelidos levam ao nome oficial) para a palavra que está sendo digitada;
//   • dentro de "(": a dica dos parâmetros do ativo (na ordem, com padrão e valores) e os valores do parâmetro atual;
//   • ")" automático ao confirmar (porte de completarParenteses / contextoDoParentese de static/ativo_dica.js).
// O catálogo vem de GET /api/plugin/catalogo (services/ativo_parametros.py, a mesma lógica da dica da aplicação). Sem dependência do CAD.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OrcamentoNanoCAD.Nucleo
{
    public class ItemCatalogo
    {
        public string Nome, Desc, Origem;
        public string ApelidoDe;           // preenchido quando a sugestão é um apelido (ex.: 11300D → DT11/300)
        public override string ToString()
        {
            string n = ApelidoDe != null ? Nome + " → " + ApelidoDe : Nome;
            return string.IsNullOrEmpty(Desc) ? n : n + "  —  " + Desc;
        }
    }

    public class ParametroCatalogo
    {
        public string Nome, Padrao;
        public List<string> Valores = new List<string>(), Mascaras = new List<string>();
        public bool Aberto = true, PorNome;
    }

    public class CatalogoBase
    {
        public List<ItemCatalogo> Ativos = new List<ItemCatalogo>();
        public Dictionary<string, string> Apelidos = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, List<ParametroCatalogo>> Parametros = new Dictionary<string, List<ParametroCatalogo>>(StringComparer.OrdinalIgnoreCase);
        public string Versao;
        private HashSet<string> _nomes;

        public bool Vazio { get { return Ativos.Count == 0; } }

        public bool Existe(string nome)
        {
            if (_nomes == null) _nomes = new HashSet<string>(Ativos.Select(a => a.Nome), StringComparer.OrdinalIgnoreCase);
            return _nomes.Contains(nome ?? "") || Apelidos.ContainsKey(nome ?? "");
        }

        public static CatalogoBase De(JObj o)
        {
            CatalogoBase c = new CatalogoBase();
            if (o == null) return c;
            c.Versao = o.Str("versao");
            foreach (JObj a in (o.Lista("ativos") ?? new List<object>()).OfType<JObj>())
                if (!string.IsNullOrEmpty(a.Str("nome"))) c.Ativos.Add(new ItemCatalogo { Nome = a.Str("nome"), Desc = a.Str("desc") ?? "", Origem = a.Str("origem") });
            JObj ap = o.Obj("apelidos");
            if (ap != null) foreach (string k in ap.Keys) c.Apelidos[k] = ap.Str(k) ?? "";
            JObj ps = o.Obj("parametros");
            if (ps != null)
                foreach (string k in ps.Keys)
                {
                    List<ParametroCatalogo> l = new List<ParametroCatalogo>();
                    foreach (JObj p in (ps.Lista(k) ?? new List<object>()).OfType<JObj>())
                        l.Add(new ParametroCatalogo
                        {
                            Nome = p.Str("nome") ?? "", Padrao = p["padrao"] == null ? null : Json.ComoTexto(p["padrao"]),
                            Valores = Json.ListaDeTextos(p["valores"]), Mascaras = Json.ListaDeTextos(p["mascaras"]),
                            Aberto = p.Bool("aberto", true), PorNome = p.Bool("porNome", false),
                        });
                    c.Parametros[k] = l;
                }
            return c;
        }
    }

    /// <summary>O parêntese ainda aberto até o cursor (como contextoDoParentese do JS).</summary>
    public class ContextoParentese
    {
        public string Nome = "";
        public int Indice;                 // posição do parâmetro (0-based)
        public string NomeParam;           // "nome=" já digitado
        public string Parcial = "";
        public int InicioParcial;
    }

    /// <summary>A palavra (nome de ativo) que está sendo digitada antes do cursor.</summary>
    public class PalavraAtual
    {
        public int Inicio;                 // onde começa o NOME (depois do prefixo de quantidade)
        public string Prefixo = "";        // "2-", "3x"…
        public string Parcial = "";
    }

    public class DicaAtivo
    {
        public string Tipo;                // parametros | sem_parametros | nao_encontrado
        public string Ativo, Mensagem;
        public List<ParametroCatalogo> Params = new List<ParametroCatalogo>();
    }

    public static class Autocompletar
    {
        private static readonly Regex Quantidade = new Regex(@"^\*?\d+(?:[.,]\d+)?[-xX×]");
        private static readonly Regex NomeAntesDoParentese = new Regex(@"(\[[^\]]*\]|[^\s,(]+)$");

        /// <summary>Fecha os parênteses que faltam (N1(D, X → N1(D, X)); vírgula/espaço pendentes saem antes; [literal] não conta.</summary>
        public static string CompletarParenteses(string texto)
        {
            string t = texto ?? "";
            int abertos = 0;
            bool colchete = false;
            foreach (char c in t)
            {
                if (colchete) { if (c == ']') colchete = false; continue; }
                if (c == '[') colchete = true;
                else if (c == '(') abertos++;
                else if (c == ')' && abertos > 0) abertos--;
            }
            if (abertos == 0 || colchete) return t;
            return t.TrimEnd(',', ' ', '\t') + new string(')', abertos);
        }

        public static ContextoParentese ContextoDoParentese(string texto, int cursor)
        {
            string t = (texto ?? "");
            if (cursor >= 0 && cursor < t.Length) t = t.Substring(0, cursor);
            Stack<int> pilha = new Stack<int>();
            bool colchete = false;
            for (int i = 0; i < t.Length; i++)
            {
                char c = t[i];
                if (c == '[') colchete = true;
                else if (c == ']') colchete = false;
                else if (!colchete && c == '(') pilha.Push(i);
                else if (!colchete && c == ')' && pilha.Count > 0) pilha.Pop();
            }
            if (pilha.Count == 0) return null;
            int abre = pilha.Peek();
            ContextoParentese r = new ContextoParentese();
            Match m = NomeAntesDoParentese.Match(t.Substring(0, abre));
            if (m.Success)
            {
                string nome = Regex.Replace(m.Groups[1].Value, @"^\*?\d+(?:[.,]\d+)?-(?=\S)", "");
                if (nome.StartsWith("[") && nome.EndsWith("]")) nome = nome.Substring(1, nome.Length - 2);
                r.Nome = nome.Trim().ToUpperInvariant();
            }
            string dentro = t.Substring(abre + 1);
            string[] partes = dentro.Split(',');
            string seg = partes[partes.Length - 1];
            r.Indice = partes.Length - 1;
            string parcial = seg;
            int eq = seg.IndexOf('=');
            if (eq >= 0) { r.NomeParam = seg.Substring(0, eq).Trim().ToUpperInvariant(); if (r.NomeParam.Length == 0) r.NomeParam = null; parcial = seg.Substring(eq + 1); }
            string sem = parcial.TrimStart();
            r.Parcial = sem;
            r.InicioParcial = t.Length - sem.Length;
            return r;
        }

        public static PalavraAtual Palavra(string texto, int cursor)
        {
            string t = texto ?? "";
            if (cursor < 0 || cursor > t.Length) cursor = t.Length;
            int i = cursor;
            while (i > 0 && !char.IsWhiteSpace(t[i - 1]) && t[i - 1] != ',' && t[i - 1] != '(' && t[i - 1] != ')') i--;
            string w = t.Substring(i, cursor - i);
            PalavraAtual p = new PalavraAtual();
            Match m = Quantidade.Match(w);
            if (m.Success) { p.Prefixo = m.Value; w = w.Substring(m.Length); i += m.Length; }
            p.Inicio = i;
            p.Parcial = w;
            return p;
        }

        /// <summary>Nomes que começam com o que foi digitado (depois os que contêm), apelidos levando ao oficial. `cabo`: os de origem CABO primeiro.</summary>
        public static List<ItemCatalogo> Sugerir(CatalogoBase c, string parcial, int max, bool cabo)
        {
            List<ItemCatalogo> r = new List<ItemCatalogo>();
            string q = (parcial ?? "").Trim();
            if (c == null || q.Length == 0) return r;
            Func<ItemCatalogo, int> origem = a => cabo ? (a.Origem == "CABO" ? 0 : 1) : (a.Origem == "CABO" ? 1 : 0);
            IEnumerable<ItemCatalogo> comeca = c.Ativos.Where(a => a.Nome.StartsWith(q, StringComparison.OrdinalIgnoreCase)).OrderBy(origem).ThenBy(a => a.Nome.Length).ThenBy(a => a.Nome);
            IEnumerable<ItemCatalogo> apel = c.Apelidos.Where(k => k.Key.StartsWith(q, StringComparison.OrdinalIgnoreCase))
                .Select(k => new ItemCatalogo { Nome = k.Key, ApelidoDe = k.Value, Desc = (c.Ativos.FirstOrDefault(a => a.Nome == k.Value) ?? new ItemCatalogo()).Desc });
            IEnumerable<ItemCatalogo> contem = q.Length < 2 ? Enumerable.Empty<ItemCatalogo>()
                : c.Ativos.Where(a => !a.Nome.StartsWith(q, StringComparison.OrdinalIgnoreCase) && a.Nome.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0).OrderBy(origem).ThenBy(a => a.Nome);
            foreach (ItemCatalogo a in comeca.Concat(apel).Concat(contem))
            {
                if (r.Count >= max) break;
                if (a.Nome.Equals(q, StringComparison.OrdinalIgnoreCase) && a.ApelidoDe == null && r.Count == 0 && !c.Ativos.Any(x => x != a && x.Nome.StartsWith(q, StringComparison.OrdinalIgnoreCase))) continue;   // já completo
                r.Add(a);
            }
            return r;
        }

        /// <summary>O que mostrar para o nome antes do "(" (como `dica` do JS).</summary>
        public static DicaAtivo Dica(CatalogoBase c, string nome)
        {
            string n = (nome ?? "").Trim().ToUpperInvariant();
            string b;
            string baseNome = c != null && c.Apelidos.TryGetValue(n, out b) ? b : n;
            List<ParametroCatalogo> ps;
            if (c != null && c.Parametros.TryGetValue(baseNome, out ps) && ps.Count > 0)
                return new DicaAtivo { Tipo = "parametros", Ativo = baseNome, Params = ps };
            if (c != null && (c.Existe(n) || c.Existe(baseNome)))
                return new DicaAtivo { Tipo = "sem_parametros", Ativo = baseNome, Mensagem = baseNome + ": ativo sem parâmetros declarados" };
            return new DicaAtivo { Tipo = "nao_encontrado", Ativo = n, Mensagem = n + ": ativo não encontrado na base" };
        }

        /// <summary>Texto da dica numa linha: "N1 — 1. tipocabo [P, M] (padrão P) · 2. poste…"; o parâmetro atual entre » «.</summary>
        public static string TextoDaDica(DicaAtivo d, ContextoParentese ctx)
        {
            if (d.Tipo != "parametros") return d.Mensagem;
            List<string> l = new List<string>();
            for (int i = 0; i < d.Params.Count; i++)
            {
                ParametroCatalogo p = d.Params[i];
                bool atual = ctx != null && (ctx.NomeParam != null ? p.Nome.ToUpperInvariant() == ctx.NomeParam : (!p.PorNome && i == ctx.Indice));
                string s = (p.PorNome ? "· " : (i + 1) + ". ") + p.Nome.ToLowerInvariant();
                if (p.Valores.Count > 0) s += " [" + string.Join(", ", p.Valores.Take(8).ToArray()) + (p.Valores.Count > 8 ? ", …" : "") + "]";
                if (p.Padrao != null) s += " (padrão " + p.Padrao + ")";
                if (p.PorNome) s += " (por nome: " + p.Nome.ToLowerInvariant() + "=…)";
                l.Add(atual ? "»" + s + "«" : s);
            }
            return d.Ativo + " — " + string.Join("  ·  ", l.ToArray());
        }

        /// <summary>Valores do parâmetro atual que começam com o que foi digitado.</summary>
        public static List<string> ValoresDoParametro(DicaAtivo d, ContextoParentese ctx)
        {
            if (d == null || d.Tipo != "parametros" || ctx == null) return new List<string>();
            ParametroCatalogo p = ctx.NomeParam != null
                ? d.Params.FirstOrDefault(x => x.Nome.ToUpperInvariant() == ctx.NomeParam)
                : (ctx.Indice < d.Params.Count && !d.Params[ctx.Indice].PorNome ? d.Params[ctx.Indice] : null);
            if (p == null) return new List<string>();
            return p.Valores.Where(v => v.StartsWith(ctx.Parcial ?? "", StringComparison.OrdinalIgnoreCase) && !v.Equals(ctx.Parcial ?? "", StringComparison.OrdinalIgnoreCase)).ToList();
        }

        /// <summary>Troca o trecho [inicio, cursor) por `valor`; devolve o texto novo e a posição do cursor.</summary>
        public static string Substituir(string texto, int inicio, int cursor, string valor, out int novoCursor)
        {
            string t = texto ?? "";
            if (cursor < inicio || cursor > t.Length) cursor = t.Length;
            string r = t.Substring(0, inicio) + valor + t.Substring(cursor);
            novoCursor = inicio + valor.Length;
            return r;
        }
    }
}
