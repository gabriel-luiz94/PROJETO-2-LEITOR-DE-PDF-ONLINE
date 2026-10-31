// Nucleo/Anotacoes.cs — leitura das anotações de cabo (LABEL) e de poste (notação compacta), no mesmo critério da aplicação:
//   cabo:  services/anotacao_cabos.py (interpretar / prefixo_e_comprimento) e vinculacao_regras.py (derivar_cabo, descricao_cabo);
//   poste: services/anotacao_outros.py (itens "[qtd ]ativo" separados por ", "; poste compacto 11300D = DT11/300).
// Só a parte que o plugin precisa para numerar, vincular e mostrar os parâmetros; o cálculo continua na aplicação.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OrcamentoNanoCAD.Nucleo
{
    public class CaboLido
    {
        public string Codigo;          // código do cabo principal (A35P, S04, C16, RAMA25…), null se não entendeu
        public double? Comprimento;
        public int Fases = 1;
        public string Tensao = "";     // "13,8kV"
        public string Nivel = "BT";    // MT quando a tensão passa de 1 kV
        public string Rotulo = "";     // anotação sem o comprimento ("3xA35P(7,94)-13,8kV")
        public string Descricao = "";  // descrição completa no formato de parâmetro (3XA35P_7.94_-13.8KV)
        public string Erro;
    }

    public class ItemPoste
    {
        public string Qtd;
        public string Nome;            // nome-base, sem parâmetros (N1, N4, 9400D…)
        public string Bruto;
    }

    public class PosteLido
    {
        public string CodigoPoste;     // 9400D (como escrito)
        public List<ItemPoste> Estruturas = new List<ItemPoste>();
        public List<ItemPoste> Itens = new List<ItemPoste>();
    }

    public static class Anotacoes
    {
        private const RegexOptions A = RegexOptions.CultureInvariant;
        private static readonly Regex _separa = new Regex(@"^(.*\S)\s+([0-9]+(?:[.,][0-9]+)?)\s*[mM]?$", A | RegexOptions.Singleline);
        private static readonly Regex _tensao = new Regex(@"\s*[-+]\s*([0-9]+(?:[.,][0-9]+)?)\s*KV\s*$", A | RegexOptions.IgnoreCase);
        private static readonly Regex _solto = new Regex(@"^[(\[]\s*([A-Z][A-Za-z0-9_]*)\s*[)\]]$", A);
        private static readonly Regex _potencia = new Regex(@"^C\s*[0-9][xX]", A);
        private static readonly Regex _codigo = new Regex(@"^(?:([0-9])[xX])?([A-Z]{0,4}[0-9][0-9A-Z/,.]*)(?:\(([^)]*)\))?$", A);
        private static readonly Regex _curto = new Regex(@"^(?:RAMA|PB)[0-9]+$", A);
        private static readonly Regex _numeroSozinho = new Regex(@"^[0-9]{2,3}$", A);
        private static readonly Regex _marca = new Regex(@"^[0-9][xX]|^C\s*[0-9][xX]|[(\[]|KV", A | RegexOptions.IgnoreCase);
        private static readonly Regex _espacos = new Regex(@"\s+", A);
        private static readonly Regex _fases = new Regex(@"(?<![0-9A-Za-z])([1-9])\s*[xX]\s*(?=[A-Za-z])", A);
        private static readonly Regex _kv = new Regex(@"(\d+(?:[.,]\d+)?)\s*kV", A | RegexOptions.IgnoreCase);
        private static readonly Regex _comprimentoFinal = new Regex(@"\s+\d+(?:[.,]\d+)?\s*m?\s*$", A | RegexOptions.IgnoreCase);

        public static double? Numero(string t)
        {
            double d;
            if (t != null && double.TryParse(t.Replace(",", "."), NumberStyles.Float, CultureInfo.InvariantCulture, out d)) return d;
            return null;
        }

        /// <summary>"45" / "45,5" — vírgula decimal, como o projeto escreve.</summary>
        public static string FormatarMetros(double m)
        {
            return Json.Numero(Math.Round(m, 2)).Replace(".", ",");
        }

        /// <summary>Anotação do cabo sem o comprimento, e o comprimento (null se não houver).</summary>
        public static void Separar(string texto, out string rotulo, out string comprimento)
        {
            string t = _espacos.Replace(texto ?? "", " ").Trim();
            Match m = _separa.Match(t);
            if (m.Success) { rotulo = m.Groups[1].Value.Trim(); comprimento = m.Groups[2].Value; }
            else { rotulo = t; comprimento = null; }
        }

        public static bool EhNotacao(string texto)
        {
            string rotulo, comp;
            Separar(texto, out rotulo, out comp);
            string r = rotulo.ToUpperInvariant();
            if (r.Length == 0) return false;
            return _marca.IsMatch(r) || _curto.IsMatch(r) || _numeroSozinho.IsMatch(r);
        }

        /// <summary>Descrição do cabo no formato aceito num parâmetro (descricao_cabo): maiúsculas, sem espaços, vírgula → ponto, parênteses → "_".</summary>
        public static string DescricaoCabo(string texto, string codigo)
        {
            string d = _espacos.Replace(_comprimentoFinal.Replace((texto ?? "").Trim(), ""), "").ToUpperInvariant().Replace(",", ".");
            d = Regex.Replace(d, @"[()\[\]]", "_");
            return d.Length > 0 ? d : (codigo ?? "").Trim().ToUpperInvariant();
        }

        public static CaboLido LerCabo(string texto)
        {
            CaboLido c = new CaboLido();
            string t = (texto ?? "").Trim();
            Match mf = _fases.Match(t), mt = _kv.Match(t);
            if (mf.Success) c.Fases = int.Parse(mf.Groups[1].Value, CultureInfo.InvariantCulture);
            if (mt.Success)
            {
                c.Tensao = mt.Groups[1].Value.Replace(".", ",") + "kV";
                double kv = Numero(mt.Groups[1].Value) ?? 0;
                c.Nivel = kv > 1 ? "MT" : "BT";
            }
            c.Descricao = DescricaoCabo(t, "");
            string rotulo, comp;
            Separar(t, out rotulo, out comp);
            c.Rotulo = rotulo;
            c.Comprimento = comp != null ? Numero(comp) : null;
            if (!EhNotacao(t))
            {
                // formato antigo "CAA2 ABC 35 m": código = 1º token
                string[] tokens = rotulo.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                c.Codigo = tokens.Length > 0 ? tokens[0].ToUpperInvariant() : null;
                if (c.Codigo == null) c.Erro = "anotação de cabo vazia";
                return c;
            }
            string r = rotulo.ToUpperInvariant();
            Match mk = _tensao.Match(r);
            if (mk.Success) r = r.Substring(0, mk.Index).Trim();
            Match solto = _solto.Match(r);
            if (solto.Success) c.Codigo = solto.Groups[1].Value;
            else if (_potencia.IsMatch(r)) c.Codigo = _espacos.Replace(r, "");
            else if (_numeroSozinho.IsMatch(r)) c.Codigo = "C" + r;
            else
            {
                Match m = _codigo.Match(r.Replace(" ", ""));
                if (!m.Success) { c.Erro = "rótulo '" + r + "' não reconhecido (ex.: 3xA35P(7,94)-13,8kV 45 m)"; return c; }
                c.Codigo = m.Groups[2].Value;      // fases continuam as de derivar_cabo (é o que o vínculo usa)
            }
            if (c.Comprimento == null) c.Erro = "falta o comprimento em '" + r + "'";
            return c;
        }

        /// <summary>Troca o comprimento da anotação ("3xA35P… 45 m" → "3xA35P… 50 m").</summary>
        public static string ComMetros(string texto, double metros)
        {
            string rotulo, comp;
            Separar(texto, out rotulo, out comp);
            return rotulo + " " + FormatarMetros(metros) + " m";
        }

        // ── poste ──
        private static readonly Regex _posteCompacto = new Regex(@"^(\d+)(FV|D|R|F)$", A | RegexOptions.IgnoreCase);
        private static readonly Regex _posteNome = new Regex(@"^(DT|RC|CV|FV)(\d+)/(\d+)$", A | RegexOptions.IgnoreCase);
        private static readonly Regex _qtd = new Regex(@"^\*?\d+(?:[.,]\d+)?$", A);

        public static bool EhPoste(string nome)
        {
            string n = (nome ?? "").Trim();
            return _posteCompacto.IsMatch(n) || _posteNome.IsMatch(n);
        }

        /// <summary>11300D → DT11/300; 7150FV → FV7/150 (altura 7..16 m). null se não for poste compacto.</summary>
        public static string PosteOficial(string compacto)
        {
            Match m = _posteCompacto.Match((compacto ?? "").Trim());
            if (!m.Success) return _posteNome.IsMatch((compacto ?? "").Trim()) ? compacto.Trim().ToUpperInvariant() : null;
            string d = m.Groups[1].Value, suf = m.Groups[2].Value.ToUpperInvariant();
            string tipo = suf == "D" ? "DT" : suf == "R" ? "RC" : suf == "F" ? "CV" : "FV";
            foreach (int n in new[] { 2, 1 })
            {
                if (d.Length <= n) continue;
                int altura = int.Parse(d.Substring(0, n), CultureInfo.InvariantCulture);
                if (altura >= 7 && altura <= 16) return tipo + altura + "/" + d.Substring(n);
            }
            return null;
        }

        /// <summary>Divide por vírgula de topo (fora de parênteses), ignorando a vírgula decimal (sem espaço depois entre dígitos).</summary>
        public static List<string> DividirItens(string texto)
        {
            List<string> r = new List<string>();
            string t = texto ?? "";
            int nivel = 0, ini = 0;
            for (int i = 0; i < t.Length; i++)
            {
                char c = t[i];
                if (c == '(') nivel++;
                else if (c == ')') nivel = Math.Max(0, nivel - 1);
                else if (c == ',' && nivel == 0)
                {
                    bool decimalNum = i > 0 && i + 1 < t.Length && char.IsDigit(t[i - 1]) && char.IsDigit(t[i + 1]);
                    if (decimalNum) continue;
                    r.Add(t.Substring(ini, i - ini).Trim());
                    ini = i + 1;
                }
            }
            r.Add(t.Substring(ini).Trim());
            r.RemoveAll(s => s.Length == 0);
            return r;
        }

        public static string NomeBase(string item)
        {
            string s = (item ?? "").Trim();
            int p = s.IndexOf('(');
            return (p >= 0 ? s.Substring(0, p) : s).Trim().ToUpperInvariant();
        }

        public static PosteLido LerPoste(string ativo)
        {
            PosteLido p = new PosteLido();
            foreach (string bruto in DividirItens(ativo))
            {
                ItemPoste it = new ItemPoste();
                it.Bruto = bruto;
                string corpo = bruto;
                int esp = bruto.IndexOf(' ');
                if (esp > 0 && _qtd.IsMatch(bruto.Substring(0, esp))) { it.Qtd = bruto.Substring(0, esp); corpo = bruto.Substring(esp + 1).Trim(); }
                it.Nome = NomeBase(corpo);
                p.Itens.Add(it);
                if (p.CodigoPoste == null && EhPoste(it.Nome)) p.CodigoPoste = it.Nome;
                else if (!EhPoste(it.Nome)) p.Estruturas.Add(it);
            }
            return p;
        }
    }
}
