// Nucleo/RegrasVinculo.cs — regras de vinculação cabo ⇄ estrutura (PORTE de services/vinculacao_regras.py, TASK-084/093/095/096/100
// do projeto principal): quem aceita quem, capacidade, lados iguais e os parâmetros que a estrutura recebe (tipocabo, cabo, cabo2).
// A configuração é o MESMO JSON do projeto (data/vinculacao_pernambuco_seed.json ou o exportado da tela de vínculo).
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OrcamentoNanoCAD.Nucleo
{
    public class InfoCabo
    {
        public string Codigo;
        public int? Regra;
        public string ValorCabo;
        public int Pontas;
        public bool Vincula = true;
        public object Estruturas;
        public List<string> Operacoes;
        public string Op;
        public string Base;
        public string Descricao;
        public string Grupo;
    }

    public class InfoEstrutura
    {
        public string Codigo;
        public int? Regra;
        public int? Qtd;
        public int? QtdMin;
        public int? MaxReais;
        public string Lados = "LIVRE";
        public bool Amarracao;
        public bool AceitaI = true;
        public object Cabos;
        public string Parametro = "CABO";
        public List<ParametroEstrutura> Parametros;
        public string SeDiferem = "MAIOR";
    }

    public class ParametroEstrutura
    {
        public string Nome;
        public string Fonte;   // descricao | valor_cabo | ativo
    }

    public class Ocupante
    {
        public InfoCabo Info;
        public bool Ficticio;
    }

    public class RegrasVinculo
    {
        public List<JObj> Cabos = new List<JObj>();
        public List<JObj> Estruturas = new List<JObj>();
        public List<string> OrdemMaior = new List<string> { "C", "P", "A" };

        public static RegrasVinculo De(JObj cfg)
        {
            RegrasVinculo r = new RegrasVinculo();
            if (cfg == null) return r;
            List<object> c = cfg.Lista("cabos"), e = cfg.Lista("estruturas"), om = cfg.Lista("ordem_maior");
            if (c != null) r.Cabos = c.OfType<JObj>().ToList();
            if (e != null) r.Estruturas = e.OfType<JObj>().ToList();
            if (om != null && om.Count > 0) r.OrdemMaior = om.Select(x => (Json.ComoTexto(x) ?? "").Trim().ToUpperInvariant()).ToList();
            return r;
        }

        private static readonly Regex _ehRegex = new Regex(@"[\\^$()\[\]|+*?{}]");

        /// <summary>Vazio casa tudo; lista compara por igualdade (sem caixa); "A, B" sem caracteres de regex = lista; senão regex sem caixa.</summary>
        public static bool Casa(object padrao, string texto)
        {
            if (padrao == null) return true;
            string t = (texto ?? "").Trim().ToUpperInvariant();
            List<object> lista = padrao as List<object>;
            if (lista != null)
                return lista.Count == 0 || lista.Any(x => (Json.ComoTexto(x) ?? "").Trim().ToUpperInvariant() == t);
            string p = (Json.ComoTexto(padrao) ?? "").Trim();
            if (p.Length == 0) return true;
            if (!_ehRegex.IsMatch(p))
                return p.Split(',').Select(x => x.Trim().ToUpperInvariant()).Where(x => x.Length > 0).Contains(t);
            Regex re = MotorLeitor.RegexJs(p);
            return re != null && re.IsMatch((texto ?? "").Trim());
        }

        public static string OpBase(string operacao)
        {
            string o = (operacao ?? "").Trim().ToUpperInvariant();
            return o == "M" || o == "*M" ? "M" : "I";
        }

        public static bool ParticipaVinculo(string operacao)
        {
            string o = (operacao ?? "").Trim().ToUpperInvariant();
            return o != "R" && o != "*R";
        }

        private static int? Inteiro(object v)
        {
            double? d = Json.ComoNumero(v);
            if (d.HasValue && d.Value == Math.Floor(d.Value)) return (int)d.Value;
            return null;
        }

        private bool RegraCaboCasa(JObj r, string cod, CaboLido extra)
        {
            if (!Casa(r["padrao"], cod)) return false;
            if (extra == null) return true;
            object fases = r["fases"];
            if (fases != null && !(fases is string && ((string)fases).Length == 0))
            {
                List<string> lista = fases is List<object> ? Json.ListaDeTextos(fases) : (Json.ComoTexto(fases) ?? "").Split(',').ToList();
                bool algum = false;
                foreach (string x in lista)
                {
                    int f;
                    if (!int.TryParse(x.Trim(), out f)) return false;
                    if (f == extra.Fases) algum = true;
                }
                if (!algum) return false;
            }
            string nivel = (r.Str("nivel") ?? "").Trim();
            if (nivel.Length > 0 && nivel.ToUpperInvariant() != extra.Nivel) return false;
            string pt = (r.Str("padrao_texto") ?? "").Trim();
            if (pt.Length > 0)
            {
                Regex re = MotorLeitor.RegexJs(pt);
                string texto = extra.Rotulo ?? "";
                if (re == null || texto.Length == 0 || !re.IsMatch(texto)) return false;
            }
            return true;
        }

        /// <summary>info_cabo: sem regra na tabela aceita tudo, pontas 2 (I) / 1 (M).</summary>
        public InfoCabo InfoCabo(string codigo, string operacao, CaboLido extra)
        {
            InfoCabo c = new InfoCabo();
            c.Codigo = (codigo ?? "").Trim().ToUpperInvariant();
            c.Op = (operacao ?? "").Trim().ToUpperInvariant();
            c.Base = OpBase(operacao);
            c.Pontas = c.Base == "M" ? 1 : 2;
            JObj regra = null;
            if (c.Codigo.Length > 0)
                for (int i = 0; i < Cabos.Count; i++)
                {
                    if (Cabos[i]["ativa"] is bool && !(bool)Cabos[i]["ativa"]) continue;
                    if (RegraCaboCasa(Cabos[i], c.Codigo, extra)) { regra = Cabos[i]; c.Regra = i; break; }
                }
            if (regra != null)
            {
                JObj pontas = regra.Obj("pontas");
                int? p = pontas != null ? Inteiro(pontas[c.Base]) : null;
                if (p.HasValue) c.Pontas = p.Value;
                string vc = (regra.Str("valor_cabo") ?? "").Trim().ToUpperInvariant();
                c.ValorCabo = vc.Length > 0 ? vc : null;
                c.Vincula = !(regra["vincula"] is bool && !(bool)regra["vincula"]);
                c.Estruturas = regra["estruturas"];
                c.Operacoes = regra["operacoes"] is List<object> ? Json.ListaDeTextos(regra["operacoes"]) : null;
                c.Grupo = regra.Str("grupo");
            }
            c.Descricao = extra != null && !string.IsNullOrEmpty(extra.Descricao) ? extra.Descricao : c.Codigo;
            return c;
        }

        public static List<ParametroEstrutura> NormalizarParametros(object v)
        {
            if (v == null || (v is string && ((string)v).Length == 0)) return null;
            List<object> itens = v as List<object> ?? (Json.ComoTexto(v) ?? "").Split(',').Cast<object>().ToList();
            List<ParametroEstrutura> lista = new List<ParametroEstrutura>();
            foreach (object x in itens)
            {
                string nome, fonte;
                JObj o = x as JObj;
                if (o != null)
                {
                    nome = (o.Str("nome") ?? "").Trim().ToUpperInvariant();
                    fonte = (o.Str("fonte") ?? "").Trim().ToLowerInvariant();
                }
                else
                {
                    string s = Json.ComoTexto(x) ?? "";
                    int eq = s.IndexOf('=');
                    nome = (eq >= 0 ? s.Substring(0, eq) : s).Trim().ToUpperInvariant();
                    fonte = (eq >= 0 ? s.Substring(eq + 1) : "").Trim().ToLowerInvariant();
                }
                if (fonte.Length == 0) fonte = FonteDoNome(nome);
                if (nome.Length > 0) { ParametroEstrutura p = new ParametroEstrutura(); p.Nome = nome; p.Fonte = fonte; lista.Add(p); }
            }
            return lista.Count > 0 ? lista : null;
        }

        public static string FonteDoNome(string nome)
        {
            if (Regex.IsMatch(nome, @"^CABO\d*$")) return "descricao";
            if (Regex.IsMatch(nome, @"^TIPOCABO\d*$")) return "valor_cabo";
            return "ativo";
        }

        public InfoEstrutura InfoEstrutura(string codigo)
        {
            InfoEstrutura d = new InfoEstrutura();
            d.Codigo = (codigo ?? "").Trim().ToUpperInvariant();
            JObj r = null;
            if (d.Codigo.Length > 0)
                for (int i = 0; i < Estruturas.Count; i++)
                {
                    if (Estruturas[i]["ativa"] is bool && !(bool)Estruturas[i]["ativa"]) continue;
                    if (Casa(Estruturas[i]["padrao"], d.Codigo)) { r = Estruturas[i]; d.Regra = i; break; }
                }
            if (r == null) return d;
            d.Qtd = Inteiro(r["qtd_max"]) ?? Inteiro(r["qtd_cabos"]);
            d.QtdMin = Inteiro(r["qtd_min"]) ?? d.Qtd;
            d.MaxReais = Inteiro(r["max_reais"]);
            d.Lados = r.Str("lados") == "IGUAIS" ? "IGUAIS" : "LIVRE";
            d.Amarracao = r.Bool("amarracao", false);
            d.AceitaI = !(r["aceita_i"] is bool && !(bool)r["aceita_i"]);
            d.Cabos = r["cabos"];
            string par = r.Str("parametro");
            if (r.Has("parametro") && par == "") d.Parametro = "";
            else d.Parametro = string.IsNullOrEmpty((par ?? "").Trim()) ? "CABO" : par.Trim().ToUpperInvariant();
            d.Parametros = NormalizarParametros(r["parametros"]);
            if (d.Parametros != null) d.Parametro = "";
            string sd = r.Str("se_diferem");
            d.SeDiferem = sd == "BLOQUEAR" || sd == "MAIOR" || sd == "AVISAR" ? sd : "MAIOR";
            return d;
        }

        /// <summary>(ok, motivo): o cabo aceita a estrutura E a estrutura aceita o cabo (e a operação).</summary>
        public static string PodeVincular(InfoCabo c, InfoEstrutura e)
        {
            if (!c.Vincula) return "o cabo " + c.Codigo + " não vincula (acompanha o cabo principal)";
            if (c.Operacoes != null && c.Operacoes.Count > 0)
            {
                List<string> aceitas = c.Operacoes.Select(x => x.Trim().ToUpperInvariant().Replace("*", "")).ToList();
                string opn = c.Op.Replace("*", "");
                if (!aceitas.Contains(opn) && !(opn == "R" && aceitas.Contains("I")))
                    return "a regra do cabo " + c.Codigo + " não aceita a operação " + c.Op;
            }
            if (!Casa(c.Estruturas, e.Codigo)) return "o cabo " + c.Codigo + " não aceita a estrutura " + e.Codigo;
            if (!Casa(e.Cabos, c.Codigo)) return "a estrutura " + e.Codigo + " não aceita o cabo " + c.Codigo;
            if (c.Base == "M" && e.Regra.HasValue && !e.Amarracao) return "cabo M só se liga a estrutura de amarração (" + e.Codigo + " não é)";
            if (c.Base == "I" && e.Regra.HasValue && !e.AceitaI) return "a estrutura " + e.Codigo + " só aceita cabo M";
            return null;
        }

        /// <summary>pode_com_ocupantes: PodeVincular + capacidade + lados iguais. null = pode.</summary>
        public static string PodeComOcupantes(InfoCabo c, InfoEstrutura e, List<Ocupante> ocupantes, bool ficticio)
        {
            string motivo = PodeVincular(c, e);
            if (motivo != null) return motivo;
            if (e.Qtd.HasValue && ocupantes.Count + 1 > e.Qtd.Value) return "a estrutura " + e.Codigo + " já tem os " + e.Qtd.Value + " cabo(s) que aceita";
            if (!ficticio && e.MaxReais.HasValue && ocupantes.Count(o => !o.Ficticio) + 1 > e.MaxReais.Value)
                return "a estrutura " + e.Codigo + " aceita só " + e.MaxReais.Value + " cabo(s) real(is); o excedente deve ser fictício";
            if (e.Lados == "IGUAIS" && ocupantes.Count > 0)
            {
                Ocupante refe = ocupantes[0];
                bool mesmo = refe.Info.Codigo == c.Codigo && (refe.Ficticio || refe.Info.Op == c.Op);
                if (!mesmo) return e.Codigo + " exige cabos iguais: " + refe.Info.Codigo + " × " + c.Codigo;
            }
            return null;
        }

        /// <summary>Valor consolidado do tipo de cabo (se_diferem MAIOR: C &gt; P &gt; A) e avisos.</summary>
        public string ValorConsolidado(InfoEstrutura e, List<Ocupante> ocupantes, List<string> avisos)
        {
            List<string> distintos = ocupantes.Where(o => o.Info.ValorCabo != null).Select(o => o.Info.ValorCabo).Distinct().OrderBy(v => v, StringComparer.Ordinal).ToList();
            if (distintos.Count == 1) return distintos[0];
            if (distintos.Count == 0) return null;
            if (e.SeDiferem == "MAIOR")
            {
                string valor = distintos.OrderBy(v => OrdemMaior.Contains(v) ? OrdemMaior.IndexOf(v) : OrdemMaior.Count).First();
                avisos.Add(e.Codigo + ": cabos diferentes (" + string.Join(", ", distintos.ToArray()) + "); usado o maior (" + valor + ")");
                return valor;
            }
            avisos.Add(e.Codigo + ": cabos diferentes (" + string.Join(", ", distintos.ToArray()) + "); parâmetro não preenchido");
            return null;
        }

        /// <summary>Parâmetros da estrutura (nome → valores, null = pendente "?"), na ordem cadastrada. Mesma regra de _analisar_base.</summary>
        public List<KeyValuePair<string, string>> Parametros(InfoEstrutura e, List<Ocupante> ocupantes, List<string> avisos)
        {
            List<KeyValuePair<string, string>> saida = new List<KeyValuePair<string, string>>();
            string valor = ValorConsolidado(e, ocupantes, avisos);
            if (e.Parametros != null)
            {
                Dictionary<string, int> cont = new Dictionary<string, int>(), vezes = new Dictionary<string, int>();
                foreach (ParametroEstrutura p in e.Parametros) vezes[p.Nome] = (vezes.ContainsKey(p.Nome) ? vezes[p.Nome] : 0) + 1;
                foreach (ParametroEstrutura p in e.Parametros)
                {
                    Match suf = Regex.Match(p.Nome, @"^(?:CABO|TIPOCABO)(\d+)$");
                    int k = suf.Success ? int.Parse(suf.Groups[1].Value) - 1 : (cont.ContainsKey(p.Nome) ? cont[p.Nome] : 0);
                    cont[p.Nome] = (cont.ContainsKey(p.Nome) ? cont[p.Nome] : 0) + 1;
                    Ocupante o = k >= 0 && k < ocupantes.Count ? ocupantes[k] : null;
                    string v = o == null ? null : (p.Fonte == "valor_cabo" ? o.Info.ValorCabo : p.Fonte == "descricao" ? o.Info.Descricao : o.Info.Codigo);
                    if (p.Fonte == "valor_cabo" && vezes[p.Nome] == 1 && ocupantes.Count > 1) v = valor;
                    saida.Add(new KeyValuePair<string, string>(p.Nome, string.IsNullOrEmpty(v) ? null : v));
                }
            }
            else if (!string.IsNullOrEmpty(e.Parametro) && (e.Regra.HasValue || valor != null))   // sem regra e sem valor: nada a pedir (evita "?" à toa)
            {
                saida.Add(new KeyValuePair<string, string>(e.Parametro, valor));
            }
            return saida;
        }
    }
}
