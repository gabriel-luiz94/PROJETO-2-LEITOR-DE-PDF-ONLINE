// Nucleo/Json.cs — JSON mínimo, sem dependências (o nanoCAD 5 roda em .NET Framework 4.0, sem System.Text.Json).
// Objetos viram JObj (mantém a ordem das chaves), listas viram List<object>, números viram double.
// C# 5: este arquivo também é compilado pelo csc do .NET Framework (build_nanocad5.bat).
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace OrcamentoNanoCAD.Nucleo
{
    /// <summary>Objeto JSON com ordem de inserção preservada.</summary>
    public class JObj : IEnumerable<KeyValuePair<string, object>>
    {
        private readonly List<string> _chaves = new List<string>();
        private readonly Dictionary<string, object> _valores = new Dictionary<string, object>();

        public object this[string chave]
        {
            get { object v; return _valores.TryGetValue(chave, out v) ? v : null; }
            set { Set(chave, value); }
        }

        public JObj Set(string chave, object valor)
        {
            if (!_valores.ContainsKey(chave)) _chaves.Add(chave);
            _valores[chave] = valor;
            return this;
        }

        public bool Has(string chave) { return _valores.ContainsKey(chave); }
        public void Remove(string chave) { if (_valores.Remove(chave)) _chaves.Remove(chave); }
        public IList<string> Keys { get { return _chaves.AsReadOnly(); } }
        public int Count { get { return _chaves.Count; } }

        public string Str(string chave) { return Json.ComoTexto(this[chave]); }
        public double? Num(string chave) { return Json.ComoNumero(this[chave]); }
        public bool Bool(string chave, bool padrao) { object v = this[chave]; return v is bool ? (bool)v : padrao; }
        public JObj Obj(string chave) { return this[chave] as JObj; }
        public List<object> Lista(string chave) { return this[chave] as List<object>; }

        public IEnumerator<KeyValuePair<string, object>> GetEnumerator()
        {
            foreach (string k in _chaves) yield return new KeyValuePair<string, object>(k, _valores[k]);
        }
        IEnumerator IEnumerable.GetEnumerator() { return GetEnumerator(); }
    }

    public class ErroJson : Exception
    {
        public ErroJson(string msg) : base(msg) { }
    }

    public static class Json
    {
        // ── leitura ──
        public static object Ler(string texto)
        {
            if (texto == null) throw new ErroJson("JSON vazio");
            int i = 0;
            if (texto.Length > 0 && texto[0] == '﻿') i = 1;
            object v = LerValor(texto, ref i);
            Pular(texto, ref i);
            if (i != texto.Length) throw new ErroJson("conteúdo extra depois do JSON na posição " + i);
            return v;
        }

        private static void Pular(string s, ref int i)
        {
            while (i < s.Length && (s[i] == ' ' || s[i] == '\t' || s[i] == '\r' || s[i] == '\n')) i++;
        }

        private static object LerValor(string s, ref int i)
        {
            Pular(s, ref i);
            if (i >= s.Length) throw new ErroJson("fim inesperado do JSON");
            char c = s[i];
            if (c == '{') return LerObjeto(s, ref i);
            if (c == '[') return LerLista(s, ref i);
            if (c == '"') return LerTexto(s, ref i);
            if (c == 't' && Prox(s, i, "true")) { i += 4; return true; }
            if (c == 'f' && Prox(s, i, "false")) { i += 5; return false; }
            if (c == 'n' && Prox(s, i, "null")) { i += 4; return null; }
            if (c == '-' || (c >= '0' && c <= '9')) return LerNumero(s, ref i);
            throw new ErroJson("caractere inesperado '" + c + "' na posição " + i);
        }

        private static bool Prox(string s, int i, string palavra)
        {
            return string.CompareOrdinal(s, i, palavra, 0, palavra.Length) == 0;
        }

        private static JObj LerObjeto(string s, ref int i)
        {
            JObj o = new JObj();
            i++;
            Pular(s, ref i);
            if (i < s.Length && s[i] == '}') { i++; return o; }
            while (true)
            {
                Pular(s, ref i);
                if (i >= s.Length || s[i] != '"') throw new ErroJson("chave esperada na posição " + i);
                string k = LerTexto(s, ref i);
                Pular(s, ref i);
                if (i >= s.Length || s[i] != ':') throw new ErroJson("':' esperado na posição " + i);
                i++;
                o.Set(k, LerValor(s, ref i));
                Pular(s, ref i);
                if (i < s.Length && s[i] == ',') { i++; continue; }
                if (i < s.Length && s[i] == '}') { i++; return o; }
                throw new ErroJson("',' ou '}' esperado na posição " + i);
            }
        }

        private static List<object> LerLista(string s, ref int i)
        {
            List<object> l = new List<object>();
            i++;
            Pular(s, ref i);
            if (i < s.Length && s[i] == ']') { i++; return l; }
            while (true)
            {
                l.Add(LerValor(s, ref i));
                Pular(s, ref i);
                if (i < s.Length && s[i] == ',') { i++; continue; }
                if (i < s.Length && s[i] == ']') { i++; return l; }
                throw new ErroJson("',' ou ']' esperado na posição " + i);
            }
        }

        private static string LerTexto(string s, ref int i)
        {
            StringBuilder sb = new StringBuilder();
            i++;
            while (i < s.Length)
            {
                char c = s[i++];
                if (c == '"') return sb.ToString();
                if (c != '\\') { sb.Append(c); continue; }
                if (i >= s.Length) break;
                char e = s[i++];
                switch (e)
                {
                    case '"': sb.Append('"'); break;
                    case '\\': sb.Append('\\'); break;
                    case '/': sb.Append('/'); break;
                    case 'b': sb.Append('\b'); break;
                    case 'f': sb.Append('\f'); break;
                    case 'n': sb.Append('\n'); break;
                    case 'r': sb.Append('\r'); break;
                    case 't': sb.Append('\t'); break;
                    case 'u':
                        if (i + 4 > s.Length) throw new ErroJson("escape \\u incompleto");
                        sb.Append((char)int.Parse(s.Substring(i, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture));
                        i += 4;
                        break;
                    default: throw new ErroJson("escape inválido \\" + e);
                }
            }
            throw new ErroJson("texto sem aspas de fechamento");
        }

        private static double LerNumero(string s, ref int i)
        {
            int ini = i;
            while (i < s.Length && "+-0123456789.eE".IndexOf(s[i]) >= 0) i++;
            double d;
            if (!double.TryParse(s.Substring(ini, i - ini), NumberStyles.Float, CultureInfo.InvariantCulture, out d))
                throw new ErroJson("número inválido na posição " + ini);
            return d;
        }

        // ── escrita ──
        public static string Escrever(object v, bool indentar)
        {
            StringBuilder sb = new StringBuilder();
            Escrever(sb, v, indentar, 0);
            return sb.ToString();
        }

        private static void Quebra(StringBuilder sb, bool indentar, int nivel)
        {
            if (!indentar) return;
            sb.Append('\n');
            sb.Append(' ', nivel * 2);
        }

        private static void Escrever(StringBuilder sb, object v, bool ind, int nivel)
        {
            if (v == null) { sb.Append("null"); return; }
            if (v is string) { EscreverTexto(sb, (string)v); return; }
            if (v is bool) { sb.Append((bool)v ? "true" : "false"); return; }
            if (v is double || v is float || v is decimal || v is int || v is long || v is short || v is byte)
            {
                sb.Append(Numero(Convert.ToDouble(v, CultureInfo.InvariantCulture)));
                return;
            }
            JObj o = v as JObj;
            if (o != null)
            {
                if (o.Count == 0) { sb.Append("{}"); return; }
                sb.Append('{');
                bool primeiro = true;
                foreach (KeyValuePair<string, object> kv in o)
                {
                    if (!primeiro) sb.Append(',');
                    primeiro = false;
                    Quebra(sb, ind, nivel + 1);
                    EscreverTexto(sb, kv.Key);
                    sb.Append(ind ? ": " : ":");
                    Escrever(sb, kv.Value, ind, nivel + 1);
                }
                Quebra(sb, ind, nivel);
                sb.Append('}');
                return;
            }
            IDictionary dic = v as IDictionary;
            if (dic != null)
            {
                JObj copia = new JObj();
                foreach (DictionaryEntry de in dic) copia.Set(Convert.ToString(de.Key, CultureInfo.InvariantCulture), de.Value);
                Escrever(sb, copia, ind, nivel);
                return;
            }
            IEnumerable lista = v as IEnumerable;
            if (lista != null)
            {
                List<object> itens = new List<object>();
                foreach (object x in lista) itens.Add(x);
                if (itens.Count == 0) { sb.Append("[]"); return; }
                bool simples = true;
                foreach (object x in itens) if (x is JObj || (x is IEnumerable && !(x is string))) simples = false;
                sb.Append('[');
                for (int k = 0; k < itens.Count; k++)
                {
                    if (k > 0) sb.Append(simples && ind ? ", " : ",");
                    if (!simples) Quebra(sb, ind, nivel + 1);
                    Escrever(sb, itens[k], ind, nivel + 1);
                }
                if (!simples) Quebra(sb, ind, nivel);
                sb.Append(']');
                return;
            }
            EscreverTexto(sb, Convert.ToString(v, CultureInfo.InvariantCulture));
        }

        public static string Numero(double d)
        {
            if (double.IsNaN(d) || double.IsInfinity(d)) return "null";
            if (d == Math.Floor(d) && Math.Abs(d) < 1e15) return ((long)d).ToString(CultureInfo.InvariantCulture);
            return d.ToString("R", CultureInfo.InvariantCulture);
        }

        private static void EscreverTexto(StringBuilder sb, string s)
        {
            sb.Append('"');
            foreach (char c in s)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    case '\b': sb.Append("\\b"); break;
                    case '\f': sb.Append("\\f"); break;
                    default:
                        if (c < 0x20) sb.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture));
                        else sb.Append(c);
                        break;
                }
            }
            sb.Append('"');
        }

        // ── conversões tolerantes ──
        public static string ComoTexto(object v)
        {
            if (v == null) return null;
            if (v is string) return (string)v;
            if (v is double) return Numero((double)v);
            if (v is bool) return (bool)v ? "true" : "false";
            return Convert.ToString(v, CultureInfo.InvariantCulture);
        }

        public static double? ComoNumero(object v)
        {
            if (v is double) return (double)v;
            if (v is int) return (int)v;
            if (v is long) return (long)v;
            return null;
        }

        public static List<string> ListaDeTextos(object v)
        {
            List<string> r = new List<string>();
            List<object> l = v as List<object>;
            if (l == null) return r;
            foreach (object x in l) if (x != null) r.Add(ComoTexto(x));
            return r;
        }
    }
}
