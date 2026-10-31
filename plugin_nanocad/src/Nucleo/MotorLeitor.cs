// Nucleo/MotorLeitor.cs — PORTE do motor de regras do leitor (static/regras_leitor_engine.js, TASK-006/082 do projeto principal).
// As REGRAS são as mesmas (JSON do projeto: data/regras_leitor_*_seed.json); só o motor é repetido aqui para o leitor rodar
// dentro do CAD sem servidor. Paridade conferida contra o próprio JS (tests/oraculo/oraculo_motor.js + MotorLeitorParidadeTests).
// Qualquer mudança no JS precisa ser repetida aqui (ver .ai/RULES.md, Regra 3).
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace OrcamentoNanoCAD.Nucleo
{
    /// <summary>Um texto do desenho, na ordem de leitura (mesmo papel do item do /upload).</summary>
    public class ItemLeitor
    {
        public string Texto;
        public string Cor = "#000000";
        public string Layer = "";
        public int Pagina = 1;
        public double? X;
        public double? Y;
    }

    public class ResultadoLeitor
    {
        public string Operacao;
        public string Ativo;
        public string Entidade;
    }

    public static class MotorLeitor
    {
        private static readonly Dictionary<string, Regex> _cache = new Dictionary<string, Regex>();
        private static readonly object _trava = new object();

        /// <summary>Regex com a flag "i" do JS. null quando a regex é inválida (o JS também ignora a regra).</summary>
        public static Regex RegexJs(string fonte)
        {
            if (fonte == null) return null;
            lock (_trava)
            {
                Regex r;
                if (_cache.TryGetValue(fonte, out r)) return r;
                try { r = new Regex(fonte, RegexOptions.ECMAScript | RegexOptions.IgnoreCase); }
                catch (ArgumentException)
                {
                    try { r = new Regex(fonte, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant); }
                    catch (ArgumentException) { r = null; }
                }
                _cache[fonte] = r;
                return r;
            }
        }

        private static int Hex2(string hex, int ini)
        {
            int v;
            if (hex.Length >= ini + 2 && int.TryParse(hex.Substring(ini, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out v)) return v;
            return -1;
        }

        private static bool EhCinza(string hex)
        {
            if (string.IsNullOrEmpty(hex) || hex.Length < 7) return false;
            int r = Hex2(hex, 1), g = Hex2(hex, 3), b = Hex2(hex, 5);
            if (r < 0 || g < 0 || b < 0) return false;
            return Math.Abs(r - g) < 5 && Math.Abs(g - b) < 5 && r > 20 && r < 230;
        }

        public static bool EhAzul(string hex)
        {
            if (string.IsNullOrEmpty(hex) || hex.Length < 7) return false;
            int r = Hex2(hex, 1), g = Hex2(hex, 3), b = Hex2(hex, 5);
            if (r < 0 || g < 0 || b < 0) return false;
            return b > 80 && b > r * 1.5 && b > g * 1.5;
        }

        public static string ClassificarCor(string hex)
        {
            string h = (string.IsNullOrEmpty(hex) ? "#000000" : hex).ToUpperInvariant();
            if (h == "#FF0000") return "VERMELHO";
            if (EhCinza(h)) return "CINZA";
            if (EhAzul(h)) return "AZUL";
            return "OUTRA";
        }

        public static bool CorCasa(List<string> corEm, string hex, List<string> corNaoEm)
        {
            string c = ClassificarCor(hex);
            if (corNaoEm != null && corNaoEm.Contains(c)) return false;
            if (corEm == null || corEm.Count == 0) return true;
            foreach (string regraCor in corEm)
            {
                if (regraCor == "PRETO") { if (c != "VERMELHO" && c != "CINZA") return true; }
                else if (c == regraCor) return true;
            }
            return false;
        }

        public static bool CorHexCasa(List<string> hexEm, List<string> hexNaoEm, string hex)
        {
            string h = (string.IsNullOrEmpty(hex) ? "#000000" : hex).ToLowerInvariant();
            if (hexNaoEm != null && hexNaoEm.Any(x => x.Trim().ToLowerInvariant() == h)) return false;
            if (hexEm == null || hexEm.Count == 0) return true;
            return hexEm.Any(x => x.Trim().ToLowerInvariant() == h);
        }

        public static bool LayerCasa(List<string> layerEm, string layerReal)
        {
            if (layerEm == null || layerEm.Count == 0) return true;
            string lr = (layerReal ?? "").Trim().ToUpperInvariant();
            return layerEm.Any(l => l.ToUpperInvariant() == lr);
        }

        private static readonly Regex _hifens = new Regex("[­‐-―−]");
        private static readonly Regex _hifenEspacos = new Regex(@"\s*-\s*");
        private static readonly Regex _espacos = new Regex(@"\s+");

        public static string LimparTexto(string texto)
        {
            string n = _hifens.Replace(texto ?? "", "-").Replace(' ', ' ');
            return _espacos.Replace(_hifenEspacos.Replace(n, "-").Trim(), " ");
        }

        public static string OperacaoPelaCor(string hex)
        {
            string c = ClassificarCor(hex);
            if (c == "VERMELHO") return "I";
            if (c == "CINZA") return "R";
            return "M";
        }

        private static bool MesmoMtext(ItemLeitor a, ItemLeitor b)
        {
            return a != null && b != null && a.X.HasValue && b.X.HasValue && a.X.Value == b.X.Value
                && a.Y.HasValue && b.Y.HasValue && a.Y.Value == b.Y.Value
                && (a.Layer ?? "") == (b.Layer ?? "") && a.Pagina == b.Pagina;
        }

        private static string Grupo1OuTudo(Match m)
        {
            return m.Groups.Count > 1 && m.Groups[1].Success ? m.Groups[1].Value : m.Value;
        }

        private static string Template(string template, Match match, string textoLimpo, string vizValor, string textoBruto, string vizJuntas)
        {
            if (template == null) return null;
            string o = template;
            if (match != null)
                for (int i = 1; i < match.Groups.Count; i++)
                    o = o.Replace("{" + i + "}", match.Groups[i].Success ? match.Groups[i].Value : "");
            o = o.Replace("{TEXTO_LIMPO}", textoLimpo ?? "");
            o = o.Replace("{vizinhancas}", vizJuntas ?? "");
            o = o.Replace("{vizinhanca}", vizValor ?? "");
            if (o.IndexOf("{TEXTO_SEM_ULTIMOS5}", StringComparison.Ordinal) >= 0)
            {
                string arrume = _espacos.Replace((textoBruto ?? "").Trim(), " ");
                o = o.Replace("{TEXTO_SEM_ULTIMOS5}", arrume.Substring(0, Math.Max(0, arrume.Length - 5)));
            }
            return o;
        }

        private class Estado { public string Operacao; public string Ativo = ""; }

        private static List<string> L(JObj r, string k) { return r.Has(k) && r[k] is List<object> ? Json.ListaDeTextos(r[k]) : null; }

        private static void RodarFase(IList<ItemLeitor> todos, int index, List<JObj> regras, Estado estado, string textoLimpo)
        {
            ItemLeitor item = todos[index];
            foreach (JObj r in regras)
            {
                if (!CorCasa(L(r, "cor_em"), item.Cor, L(r, "cor_nao_em"))) continue;
                if (!CorHexCasa(L(r, "cor_hex_em"), L(r, "cor_hex_nao_em"), item.Cor)) continue;
                if (!LayerCasa(L(r, "layer_em"), item.Layer)) continue;

                string modo = r.Str("modo");
                if (string.IsNullOrEmpty(modo)) modo = "DEFINIR";
                bool casou = false;
                string resultado = null;

                if (modo == "DEFINIR")
                {
                    Match match = null;
                    string textoRegex = r.Str("texto_regex");
                    if (!string.IsNullOrEmpty(textoRegex))
                    {
                        Regex re = RegexJs(textoRegex);
                        if (re == null) continue;
                        match = string.IsNullOrEmpty(item.Texto) ? null : re.Match(item.Texto);
                        if (match == null || !match.Success) continue;
                    }
                    casou = true;
                    JObj vz = r.Obj("vizinhanca");
                    string vzRegex = vz != null ? vz.Str("regex") : null;
                    if (vz != null && !string.IsNullOrEmpty(vzRegex) && vz.Bool("mesmo_mtext", false))
                    {
                        Regex re = RegexJs(vzRegex);
                        int passo = vz.Str("direcao") == "ANTES" ? -1 : 1;
                        int janela = (int)(vz.Num("janela") ?? 0);
                        if (janela == 0) janela = 2;
                        bool juntar = vz.Bool("juntar", false);
                        List<string> achados = new List<string>();
                        for (int k = 1; re != null && k <= janela; k++)
                        {
                            int j = index + passo * k;
                            ItemLeitor viz = j >= 0 && j < todos.Count ? todos[j] : null;
                            if (viz == null || string.IsNullOrEmpty(viz.Texto) || !MesmoMtext(item, viz)) break;
                            Match m = re.Match(viz.Texto);
                            if (!m.Success) { if (juntar) break; continue; }
                            achados.Add(Grupo1OuTudo(m));
                            if (!juntar) break;
                        }
                        if (vz.Bool("obrigatoria", false) && achados.Count == 0) continue;
                        if (passo == -1) achados.Reverse();
                        string vizValor = achados.Count > 0 ? achados[0] : null;
                        string vizJuntas = achados.Count > 0 ? ", " + string.Join(", ", achados.ToArray()) : "";
                        resultado = (juntar || vizValor != null)
                            ? Template(r.Str("ativo_template"), match, textoLimpo, vizValor, item.Texto, vizJuntas) : null;
                    }
                    else if (vz != null && !string.IsNullOrEmpty(vzRegex))
                    {
                        int janela = (int)(vz.Num("janela") ?? 0);
                        if (janela == 0) janela = 10;
                        int ini = vz.Str("direcao") == "DEPOIS" ? Math.Min(index + 1, todos.Count) : Math.Max(0, index - janela);
                        int fim = Math.Min(todos.Count - 1, index + janela);
                        Regex re = RegexJs(vzRegex);
                        string vizValor = null;
                        bool mesmaPagina = !(vz["mesma_pagina"] is bool) || (bool)vz["mesma_pagina"];
                        if (re != null)
                        {
                            for (int i = ini; i <= fim; i++)
                            {
                                if (i == index) continue;
                                ItemLeitor viz = todos[i];
                                if (viz == null || string.IsNullOrEmpty(viz.Texto)) continue;
                                if (mesmaPagina && viz.Pagina != item.Pagina) continue;
                                Match m = re.Match(viz.Texto);
                                if (m.Success) { vizValor = Grupo1OuTudo(m); break; }
                            }
                        }
                        resultado = vizValor == null ? null : Template(r.Str("ativo_template"), match, textoLimpo, vizValor, item.Texto, null);
                    }
                    else
                    {
                        resultado = Template(r.Str("ativo_template"), match, textoLimpo, null, item.Texto, null);
                    }
                }
                else if (modo == "SUBSTITUIR" || modo == "SUBSTITUIR_TOTAL")
                {
                    string ativoRegex = r.Str("ativo_regex");
                    if (string.IsNullOrEmpty(ativoRegex)) continue;
                    Regex re = RegexJs(ativoRegex);
                    if (re == null) continue;
                    Match match = re.Match(estado.Ativo);
                    if (!match.Success) continue;
                    casou = true;
                    if (modo == "SUBSTITUIR_TOTAL") resultado = Template(r.Str("ativo_template"), match, textoLimpo, null, item.Texto, null);
                    else resultado = re.Replace(estado.Ativo, r.Str("ativo_template") ?? "", 1);
                }

                if (!casou) continue;

                string op = r.Str("operacao");
                if (op == "PELA_COR") estado.Operacao = OperacaoPelaCor(item.Cor);
                else if (op == "0") estado.Operacao = "0";
                else if (!string.IsNullOrEmpty(op)) estado.Operacao = op;

                if (resultado != null) estado.Ativo = resultado;
                if (r.Bool("parar", false) || modo == "SUBSTITUIR_TOTAL") break;
            }
        }

        private static List<JObj> Ativas(IEnumerable<JObj> regras)
        {
            return (regras ?? new JObj[0]).Where(r => r != null && !(r["ativa"] is bool && (bool)r["ativa"] == false)).ToList();
        }

        private static double Ordem(JObj r) { return r.Num("ordem") ?? 0; }

        /// <summary>`r.fase || 3` do JS: sem fase (ou 0) = fase 3.</summary>
        private static int FaseDe(JObj r) { int f = (int)(r.Num("fase") ?? 0); return f == 0 ? 3 : f; }

        public static ResultadoLeitor Processar(IList<ItemLeitor> todos, int index, IEnumerable<JObj> regras)
        {
            ItemLeitor item = todos[index];
            string textoLimpo = LimparTexto(item.Texto);
            Estado estado = new Estado();
            List<JObj> ativas = Ativas(regras);
            for (int fase = 1; fase <= 3; fase++)
            {
                int f = fase;
                List<JObj> doFase = ativas.Where(r => FaseDe(r) == f).OrderBy(r => Ordem(r)).ToList();
                RodarFase(todos, index, doFase, estado, textoLimpo);
            }
            string operacao = estado.Operacao ?? OperacaoPelaCor(item.Cor);
            string layer = (item.Layer ?? "").Trim().ToUpperInvariant();
            if (layer == "01_LV" && !string.IsNullOrEmpty(operacao) && operacao[0] != '*') operacao = "*" + operacao;
            if (EhAzul(item.Cor) && layer != "01_LV" && layer != "01_RETENS" && layer != "01_RETENS_LV") operacao = "0";
            if (layer == "01_PARTICULAR") operacao = "0";
            ResultadoLeitor res = new ResultadoLeitor();
            res.Operacao = operacao;
            res.Ativo = (estado.Ativo ?? "").Trim();
            return res;
        }

        public static ResultadoLeitor Classificar(string operacao, string ativo, string cor, string layer, string texto, IEnumerable<JObj> regras)
        {
            string entidade = "0";
            string opFinal = operacao;
            foreach (JObj r in Ativas(regras).OrderBy(r => Ordem(r)))
            {
                List<string> opEm = L(r, "operacao_em");
                if (opEm != null && opEm.Count > 0 && !opEm.Contains(opFinal)) continue;
                if (!CorCasa(L(r, "cor_em"), cor, L(r, "cor_nao_em"))) continue;
                if (!CorHexCasa(L(r, "cor_hex_em"), L(r, "cor_hex_nao_em"), cor)) continue;
                if (!LayerCasa(L(r, "layer_em"), layer)) continue;
                string ar = r.Str("ativo_regex");
                if (!string.IsNullOrEmpty(ar)) { Regex re = RegexJs(ar); if (re == null || !re.IsMatch(ativo ?? "")) continue; }
                string tr = r.Str("texto_regex");
                if (!string.IsNullOrEmpty(tr)) { Regex re = RegexJs(tr); if (re == null || !re.IsMatch(texto ?? "")) continue; }
                string ent = r.Str("entidade");
                if (!string.IsNullOrEmpty(ent)) entidade = ent;
                string aj = r.Str("operacao_ajustada");
                if (!string.IsNullOrEmpty(aj)) opFinal = aj;
                if (r.Bool("parar", false)) break;
            }
            ResultadoLeitor res = new ResultadoLeitor();
            res.Operacao = opFinal;
            res.Entidade = entidade;
            return res;
        }

        public static ResultadoLeitor ProcessarEClassificar(IList<ItemLeitor> todos, int index, IEnumerable<JObj> proc, IEnumerable<JObj> cls)
        {
            ResultadoLeitor p = Processar(todos, index, proc);
            if (p.Operacao == "0") { p.Entidade = "0"; return p; }
            ItemLeitor item = todos[index];
            ResultadoLeitor c = Classificar(p.Operacao, p.Ativo, item.Cor, item.Layer, item.Texto, cls);
            c.Ativo = p.Ativo;
            return c;
        }
    }
}
