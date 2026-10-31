// Nucleo/Cores.cs — cor e texto como o leitor de DXF da aplicação resolve (services/dxf_service.py), para o plugin ler o DWG igual.
// Hierarquia da cor: true color da entidade → ACI da entidade → layer TXT_RRGGBB → true color da layer → ACI da layer → preto.
// ACI 0 (BYBLOCK) e 7 (branco) = preto. MTEXT: quebra por parágrafo (\P), cor herdada entre parágrafos e local entre { }.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace OrcamentoNanoCAD.Nucleo
{
    public class ParteTexto
    {
        public string Texto;
        public string Cor;
    }

    public static class Cores
    {
        // Tabela ACI → RGB (mesma do ezdxf.colors.aci2rgb usada pela aplicação).
        private static readonly int[] _aci = {
            0x000000, 0xFF0000, 0xFFFF00, 0x00FF00, 0x00FFFF, 0x0000FF, 0xFF00FF, 0xFFFFFF,
            0x808080, 0xC0C0C0, 0xFF0000, 0xFF7F7F, 0xA50000, 0xA55252, 0x7F0000, 0x7F3F3F,
            0x4C0000, 0x4C2626, 0x260000, 0x261313, 0xFF3F00, 0xFF9F7F, 0xA52900, 0xA56752,
            0x7F1F00, 0x7F4F3F, 0x4C1300, 0x4C2F26, 0x260900, 0x261713, 0xFF7F00, 0xFFBF7F,
            0xA55200, 0xA57C52, 0x7F3F00, 0x7F5F3F, 0x4C2600, 0x4C3926, 0x261300, 0x261C13,
            0xFFBF00, 0xFFDF7F, 0xA57C00, 0xA59152, 0x7F5F00, 0x7F6F3F, 0x4C3900, 0x4C4226,
            0x261C00, 0x262113, 0xFFFF00, 0xFFFF7F, 0xA5A500, 0xA5A552, 0x7F7F00, 0x7F7F3F,
            0x4C4C00, 0x4C4C26, 0x262600, 0x262613, 0xBFFF00, 0xDFFF7F, 0x7CA500, 0x91A552,
            0x5F7F00, 0x6F7F3F, 0x394C00, 0x424C26, 0x1C2600, 0x212613, 0x7FFF00, 0xBFFF7F,
            0x52A500, 0x7CA552, 0x3F7F00, 0x5F7F3F, 0x264C00, 0x394C26, 0x132600, 0x1C2613,
            0x3FFF00, 0x9FFF7F, 0x29A500, 0x67A552, 0x1F7F00, 0x4F7F3F, 0x134C00, 0x2F4C26,
            0x092600, 0x172613, 0x00FF00, 0x7FFF7F, 0x00A500, 0x52A552, 0x007F00, 0x3F7F3F,
            0x004C00, 0x264C26, 0x002600, 0x132613, 0x00FF3F, 0x7FFF9F, 0x00A529, 0x52A567,
            0x007F1F, 0x3F7F4F, 0x004C13, 0x264C2F, 0x002609, 0x135817, 0x00FF7F, 0x7FFFBF,
            0x00A552, 0x52A57C, 0x007F3F, 0x3F7F5F, 0x004C26, 0x264C39, 0x002613, 0x13581C,
            0x00FFBF, 0x7FFFDF, 0x00A57C, 0x52A591, 0x007F5F, 0x3F7F6F, 0x004C39, 0x264C42,
            0x00261C, 0x135858, 0x00FFFF, 0x7FFFFF, 0x00A5A5, 0x52A5A5, 0x007F7F, 0x3F7F7F,
            0x004C4C, 0x264C4C, 0x002626, 0x135858, 0x00BFFF, 0x7FDFFF, 0x007CA5, 0x5291A5,
            0x005F7F, 0x3F6F7F, 0x00394C, 0x26427E, 0x001C26, 0x135858, 0x007FFF, 0x7FBFFF,
            0x0052A5, 0x527CA5, 0x003F7F, 0x3F5F7F, 0x00264C, 0x26397E, 0x001326, 0x131C58,
            0x003FFF, 0x7F9FFF, 0x0029A5, 0x5267A5, 0x001F7F, 0x3F4F7F, 0x00134C, 0x262F7E,
            0x000926, 0x131758, 0x0000FF, 0x7F7FFF, 0x0000A5, 0x5252A5, 0x00007F, 0x3F3F7F,
            0x00004C, 0x26267E, 0x000026, 0x131358, 0x3F00FF, 0x9F7FFF, 0x2900A5, 0x6752A5,
            0x1F007F, 0x4F3F7F, 0x13004C, 0x2F267E, 0x090026, 0x171358, 0x7F00FF, 0xBF7FFF,
            0x5200A5, 0x7C52A5, 0x3F007F, 0x5F3F7F, 0x26004C, 0x39267E, 0x130026, 0x1C1358,
            0xBF00FF, 0xDF7FFF, 0x7C00A5, 0x9152A5, 0x5F007F, 0x6F3F7F, 0x39004C, 0x42264C,
            0x1C0026, 0x581358, 0xFF00FF, 0xFF7FFF, 0xA500A5, 0xA552A5, 0x7F007F, 0x7F3F7F,
            0x4C004C, 0x4C264C, 0x260026, 0x581358, 0xFF00BF, 0xFF7FDF, 0xA5007C, 0xA55291,
            0x7F005F, 0x7F3F6F, 0x4C0039, 0x4C2642, 0x26001C, 0x581358, 0xFF007F, 0xFF7FBF,
            0xA50052, 0xA5527C, 0x7F003F, 0x7F3F5F, 0x4C0026, 0x4C2639, 0x260013, 0x58131C,
            0xFF003F, 0xFF7F9F, 0xA50029, 0xA55267, 0x7F001F, 0x7F3F4F, 0x4C0013, 0x4C262F,
            0x260009, 0x581317, 0x000000, 0x656565, 0x666666, 0x999999, 0xCCCCCC, 0xFFFFFF,
        };

        public static string Hex(int r, int g, int b)
        {
            return "#" + r.ToString("x2", CultureInfo.InvariantCulture) + g.ToString("x2", CultureInfo.InvariantCulture) + b.ToString("x2", CultureInfo.InvariantCulture);
        }

        /// <summary>ACI → "#rrggbb". 0 e 7 (e fora da faixa) = preto, como _aci_to_hex.</summary>
        public static string AciParaHex(int aci)
        {
            if (aci == 0 || aci == 7 || aci < 0 || aci > 255) return "#000000";
            int v = _aci[aci];
            return Hex((v >> 16) & 0xFF, (v >> 8) & 0xFF, v & 0xFF);
        }

        private static readonly Regex _layerHex = new Regex("^.*_([0-9A-Fa-f]{6})$");

        /// <summary>Layer no padrão TXT_RRGGBB → "#RRGGBB" (maiúsculas, como a aplicação); senão null.</summary>
        public static string CorDoNomeDaLayer(string layer)
        {
            Match m = _layerHex.Match(layer ?? "");
            return m.Success ? "#" + m.Groups[1].Value.ToUpperInvariant() : null;
        }

        private static readonly Regex _paraFmt = new Regex(@"\\p[^;\\\\]*;");
        private static readonly Regex _divide = new Regex(@"\\P|[\r\n]+|\^M|\^J|(?=\{\\[cC]\d+;)|(?<!\{)(?=\\[cC]\d+;)|(?<=\})");
        private static readonly Regex _codCor = new Regex(@"\\([cC])(\d+)");

        private static string CorDaParte(string parte, string herdada)
        {
            MatchCollection ms = _codCor.Matches(parte);
            if (ms.Count == 0) return herdada;
            Match m = ms[ms.Count - 1];
            int valor;
            if (!int.TryParse(m.Groups[2].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out valor)) return herdada;
            if (m.Groups[1].Value == "c") return Hex((valor >> 16) & 0xFF, (valor >> 8) & 0xFF, valor & 0xFF);
            if (valor == 0 || valor == 256) return herdada;
            return AciParaHex(valor);
        }

        /// <summary>Conteúdo bruto do MTEXT (com códigos) → parágrafos limpos com a cor de cada um.</summary>
        public static List<ParteTexto> DividirMText(string bruto, string corBase)
        {
            List<ParteTexto> saida = new List<ParteTexto>();
            string pre = _paraFmt.Replace(bruto ?? "", "");
            string atual = corBase;
            foreach (string parte in _divide.Split(pre))
            {
                if (string.IsNullOrEmpty(parte) || parte.Trim().Length == 0) continue;
                string p = parte.Trim();
                bool bloco = p.StartsWith("{") && p.EndsWith("}");
                string cor = CorDaParte(parte, atual);
                string txt = LimparMText(parte);
                if (txt.Length == 0) continue;
                ParteTexto pt = new ParteTexto();
                pt.Texto = txt;
                pt.Cor = cor;
                saida.Add(pt);
                if (!bloco) atual = cor;
            }
            return saida;
        }

        private static readonly Regex _empilhado = new Regex(@"\\S([^;]*?)[\^/#]([^;]*);");
        private static readonly Regex _codigoComArg = new Regex(@"\\[fFcCHhWwQqTtAap][^;]*;");
        private static readonly Regex _codigoSimples = new Regex(@"\\[LlOoKkNX]");
        private static readonly Regex _residuos = new Regex(@"\\[a-zA-Z][^\;\s]*;?");
        private static readonly Regex _espacos = new Regex(@"[ \t]+");

        /// <summary>Equivalente a clean_dxf_text (plain_mtext + limpeza de resíduos).</summary>
        public static string LimparMText(string bruto)
        {
            if (string.IsNullOrEmpty(bruto)) return "";
            string t = _paraFmt.Replace(bruto, "");
            t = t.Replace("\\\\", "\u0001").Replace("\\{", "\u0002").Replace("\\}", "\u0003");
            t = _empilhado.Replace(t, "$1 $2");
            t = _codigoComArg.Replace(t, "");
            t = _codigoSimples.Replace(t, "");
            t = t.Replace("\\~", " ").Replace("\\P", "\n");
            t = t.Replace("{", "").Replace("}", "");
            t = t.Replace("%%d", "°").Replace("%%D", "°").Replace("%%c", "Ø").Replace("%%C", "Ø").Replace("%%p", "±").Replace("%%P", "±");
            t = _residuos.Replace(t, "");
            t = t.Replace("\u0001", "\\").Replace("\u0002", "{").Replace("\u0003", "}");
            t = Regex.Replace(t, @"\\{2,}", "");
            t = _espacos.Replace(t, " ");
            return t.Trim();
        }

        /// <summary>TEXT/ATTRIB: só colapsa espaços (o texto já vem sem códigos de formatação).</summary>
        public static string LimparTextoSimples(string bruto)
        {
            return _espacos.Replace(bruto ?? "", " ").Trim();
        }
    }
}
