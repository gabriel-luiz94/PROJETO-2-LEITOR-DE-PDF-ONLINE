using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    public static partial class DesenhoJsonModelo
    {
        private const string AppAnotaposte = "ANOTAPOSTE";

        private sealed class NotaPostePatch
        {
            public string Handle, Grupo;
            public List<IList<object>> Tags;
            public List<string> Pecas;
            public readonly List<string> Alvos = new List<string>();
        }

        private static bool TemAnotaposte(JObj registro)
        {
            return registro.Lista("tags").Cast<IList<object>>().Any(t => TagIgual(t, 1001, AppAnotaposte));
        }

        private static bool TagIgual(IList<object> tag, int codigo, string valor)
        {
            return Json.ComoNumero(tag[0]) == codigo && string.Equals(tag[1] as string, valor, StringComparison.Ordinal);
        }

        private static NotaPostePatch LerAnotaposte(string handle, JObj registro)
        {
            var erro = new InvalidDataException("ANOTAPOSTE inválida ou com metadados não suportados no handle " + handle + ".");
            string tipo = registro.Str("type");
            var tags = registro.Lista("tags").Cast<IList<object>>().ToList();
            if ((tipo != "LINE" && tipo != "MTEXT" && tipo != "LWPOLYLINE") ||
                tags.Any(t => { double? c = Json.ComoNumero(t[0]); return c == 102 || (c >= 320 && c <= 369 && c != 330)
                    || (c >= 390 && c <= 399) || c == 480 || c == 481; }))
                throw erro;
            var xd = tags.Where(t => Json.ComoNumero(t[0]) >= 1000).ToList();
            if (xd.Count < 4 || !TagIgual(xd[0], 1001, AppAnotaposte) || Json.ComoNumero(xd[1][0]) != 1000 ||
                ((string)xd[1][1]).Trim().Length == 0)
                throw erro;
            var nota = new NotaPostePatch { Handle = handle, Grupo = (string)xd[1][1], Tags = xd };
            if (TagIgual(xd[2], 1000, "PECA"))
            {
                string grupo = (string)xd[3][1];
                if (Json.ComoNumero(xd[3][0]) != 1000 ||
                    !new[] { "G=m", "G=i", "G=a", "G=r", "G=ax", "G=c" }.Contains(grupo) ||
                    (xd.Count != 4 && !(xd.Count == 5 && grupo == "G=ax" && TagIgual(xd[4], 1000, "REAPX"))))
                    throw erro;
                return nota;
            }
            if (tipo != "LWPOLYLINE" || xd.Count < 10 || !TagIgual(xd[2], 1000, "GUIA") ||
                Json.ComoNumero(xd[3][0]) != 1000 || Json.ComoNumero(xd[4][0]) != 1000 ||
                NormalizarHandle((string)xd[4][1]) != handle)
                throw erro;
            for (int i = 0; i < 3; i++)
            {
                double coord;
                if (Json.ComoNumero(xd[5 + i][0]) != 1010 + 10 * i ||
                    !double.TryParse((string)xd[5 + i][1], NumberStyles.Float, CultureInfo.InvariantCulture, out coord)
                    || double.IsNaN(coord) || double.IsInfinity(coord))
                    throw erro;
            }
            int pos = 8;
            nota.Pecas = LerHandlesNota(xd, ref pos, erro);
            nota.Alvos.AddRange(LerHandlesNota(xd, ref pos, erro));
            string mascara = (string)xd[3][1];
            if (nota.Pecas.Count == 0 || mascara.Length != nota.Pecas.Count ||
                mascara.Any(c => "MIARC".IndexOf(c) < 0) || nota.Alvos.Count > 1 ||
                (pos != xd.Count && !(pos + 1 == xd.Count && TagIgual(xd[pos], 1000, "GCEN=1"))))
                throw erro;
            return nota;
        }

        private static List<string> LerHandlesNota(List<IList<object>> tags, ref int pos, InvalidDataException erro)
        {
            short total;
            if (pos >= tags.Count || Json.ComoNumero(tags[pos][0]) != 1070 ||
                !short.TryParse((string)tags[pos][1], NumberStyles.Integer, CultureInfo.InvariantCulture, out total)
                || total < 0 || total >= tags.Count - pos)
                throw erro;
            pos++;
            var handles = new List<string>();
            var vistos = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < total; i++)
            {
                var tag = tags[pos++];
                string h = NormalizarHandle((string)tag[1]);
                if (Json.ComoNumero(tag[0]) != 1005 || h.Length == 0 || !vistos.Add(h)) throw erro;
                handles.Add(h);
            }
            return handles;
        }

        private static List<NotaPostePatch> LerGruposAnotaposte(IEnumerable<string> handles, Dictionary<string, JObj> entidades)
        {
            var notas = handles.Where(h => TemAnotaposte(entidades[h])).Select(h => LerAnotaposte(h, entidades[h])).ToList();
            foreach (var grupo in notas.GroupBy(n => n.Grupo, StringComparer.Ordinal))
            {
                var guias = grupo.Where(n => n.Pecas != null).ToList();
                var pecas = new HashSet<string>(grupo.Where(n => n.Pecas == null).Select(n => n.Handle), StringComparer.OrdinalIgnoreCase);
                if (guias.Count != 1 || !pecas.SetEquals(guias[0].Pecas))
                    throw new InvalidDataException("Envie a guia e todas as peças do grupo ANOTAPOSTE " + grupo.Key + " na mesma operação.");
            }
            return notas;
        }

        private static string GrupoAnotaposte(JObj registro)
        {
            var tags = registro.Lista("tags").Cast<IList<object>>().ToList();
            int i = tags.FindIndex(t => TagIgual(t, 1001, AppAnotaposte));
            return i >= 0 && i + 1 < tags.Count && Json.ComoNumero(tags[i + 1][0]) == 1000 ? (string)tags[i + 1][1] : null;
        }

        public static List<IList<object>> RemapearXDataPatch(IList<IList<object>> tags, string handleProprio,
            Func<string, string> resolver)
        {
            bool guiaPoste = tags.Count > 4 && TagIgual(tags[0], 1001, AppAnotaposte) && TagIgual(tags[2], 1000, "GUIA");
            var resultado = new List<IList<object>>();
            for (int i = 0; i < tags.Count; i++)
            {
                var tag = tags[i];
                if (guiaPoste && i == 4) resultado.Add(new object[] { tag[0], handleProprio });
                else if (Json.ComoNumero(tag[0]) == 1005)
                {
                    string handle = resolver(NormalizarHandle((string)tag[1]));
                    if (string.IsNullOrEmpty(handle)) throw new InvalidDataException("Referência não encontrada ao aplicar o patch: " + tag[1]);
                    resultado.Add(new object[] { tag[0], handle });
                }
                else resultado.Add(tag);
            }
            return resultado;
        }
    }
}
