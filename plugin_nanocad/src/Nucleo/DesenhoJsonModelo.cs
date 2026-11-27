using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace OrcamentoNanoCAD.Nucleo
{
    public static partial class DesenhoJsonModelo
    {
        private static readonly Dictionary<string, HashSet<int>> Geometria = new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase)
        {
            { "LINE", new HashSet<int> { 10, 20, 30, 11, 21, 31 } },
            { "CIRCLE", new HashSet<int> { 10, 20, 30, 40 } },
            { "ARC", new HashSet<int> { 10, 20, 30, 40, 50, 51 } },
            { "LWPOLYLINE", new HashSet<int> { 10, 20, 38, 40, 41, 42 } }
        };

        public static void RestringirAoModelo(JObj desenho)
        {
            foreach (JObj secao in desenho.Lista("sections").OfType<JObj>())
            {
                var registros = secao.Lista("records");
                if (secao.Str("name") == "ENTITIES")
                {
                    bool modelo = true;
                    registros.RemoveAll(obj =>
                    {
                        var r = (JObj)obj;
                        string tipo = r.Str("type");
                        if (tipo == "VERTEX" || tipo == "ATTRIB" || tipo == "SEQEND")
                        {
                            bool remover = !modelo || Papel(r);
                            if (tipo == "SEQEND") modelo = true;
                            return remover;
                        }
                        modelo = !Papel(r);
                        return !modelo;
                    });
                }
                else if (secao.Str("name") == "BLOCKS")
                {
                    bool papel = false;
                    registros.RemoveAll(obj =>
                    {
                        var r = (JObj)obj;
                        if (r.Str("type") == "BLOCK")
                        {
                            papel = Valor(r, 2).StartsWith("*Paper_Space", StringComparison.OrdinalIgnoreCase);
                            return false;
                        }
                        if (r.Str("type") == "ENDBLK") { papel = false; return false; }
                        return papel;
                    });
                }
            }
        }

        public static void MarcarOrigem(JObj desenho)
        {
            desenho.Set("sourceModel", Assinatura(Registros(desenho)));
            desenho.Set("sourceDependencies", Assinatura(Dependencias(desenho)));
        }

        public static void GarantirSubstituicaoSemReferencias(JObj integral, JObj modelo)
        {
            var registros = Registros(modelo);
            var handles = new HashSet<string>(registros.OfType<JObj>().Select(r => NormalizarHandle(Valor(r, 5))), StringComparer.OrdinalIgnoreCase);
            handles.Remove("");
            var registrosDoModelo = new HashSet<string>(registros.Select(r => Json.Escrever(r, false)));
            foreach (JObj secao in integral.Lista("sections").OfType<JObj>())
            {
                foreach (JObj registro in secao.Lista("records").OfType<JObj>())
                {
                    if (secao.Str("name") == "ENTITIES" && registrosDoModelo.Contains(Json.Escrever(registro, false))) continue;
                    foreach (IList<object> tag in registro.Lista("tags").Cast<IList<object>>())
                    {
                        double? codigo = Json.ComoNumero(tag[0]);
                        string valor = NormalizarHandle((string)tag[1]);
                        if ((codigo == 1005 || (codigo >= 320 && codigo <= 369) || (codigo >= 390 && codigo <= 399)) &&
                            handles.Contains(valor))
                            throw new InvalidDataException("Há uma referência externa a uma entidade do modelo (" + valor +
                                "). Use ORC_IMPORTAR_JSON para manter os vínculos existentes.");
                    }
                }
            }
        }

        public static void GarantirSemProxies(JObj desenho)
        {
            if (TemProxies(desenho))
                throw new NotSupportedException("O JSON ainda contém objetos proxy sem conversão. Abra o desenho no aplicativo que os criou.");
        }

        public static bool TemProxies(JObj desenho)
        {
            return desenho.Lista("sections").OfType<JObj>().SelectMany(s => s.Lista("records").OfType<JObj>()).Any(r =>
                r.Str("type") == "ACAD_PROXY_ENTITY" || r.Str("type") == "ACAD_PROXY_OBJECT" || r.Str("type") == "SPDSNOTEPOSITION");
        }

        public static Dictionary<string, JObj> Entidades(JObj desenho)
        {
            int total;
            return IndexarEntidades(desenho, false, out total);
        }

        public static int ContarParaImportacao(JObj desenho)
        {
            int total;
            IndexarEntidades(desenho, VersaoDxf(desenho) == "AC1009", out total);
            return total;
        }

        private static Dictionary<string, JObj> IndexarEntidades(JObj desenho, bool aceitarSemHandle, out int total)
        {
            var resultado = new Dictionary<string, JObj>(StringComparer.OrdinalIgnoreCase);
            total = 0;
            int posicao = 0;
            foreach (JObj r in Registros(desenho).OfType<JObj>())
            {
                posicao++;
                string tipo = r.Str("type");
                if (tipo == "VERTEX" || tipo == "ATTRIB" || tipo == "SEQEND") continue;
                if (tipo == "data" && r.Lista("tags").Cast<IList<object>>().All(t => Json.ComoNumero(t[0]) == 999)) continue;
                total++;
                string bruto = Valor(r, 5, null);
                if (bruto == null && aceitarSemHandle) continue;
                string local = tipo + " no registro " + posicao + " de ENTITIES.";
                if (bruto == null)
                    throw new InvalidDataException("Handle ausente em " + local);
                string handle = NormalizarHandle(bruto);
                if (handle.Length == 0)
                    throw new InvalidDataException("Handle inválido em " + local);
                if (resultado.ContainsKey(handle))
                    throw new InvalidDataException("Handle repetido (" + handle + ") em " + local);
                resultado.Add(handle, r);
            }
            return resultado;
        }

        private static string VersaoDxf(JObj desenho)
        {
            var cabecalho = desenho.Lista("sections").OfType<JObj>().FirstOrDefault(s => s.Str("name") == "HEADER");
            if (cabecalho == null) return "";
            bool versao = false;
            foreach (var tag in cabecalho.Lista("records").OfType<JObj>().SelectMany(r => r.Lista("tags")).Cast<IList<object>>())
            {
                double? codigo = Json.ComoNumero(tag[0]);
                if (codigo == 999) continue;
                string valor = ((string)tag[1]).Trim();
                if (versao) return codigo == 1 ? valor : "";
                versao = codigo == 9 && valor == "$ACADVER";
            }
            return "";
        }

        private static string NormalizarHandle(string valor)
        {
            long numero;
            return long.TryParse(valor.Trim(), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out numero) && numero > 0
                ? numero.ToString("X", CultureInfo.InvariantCulture) : "";
        }

        public static List<string> PlanejarAtualizacao(JObj atual, JObj recebido)
        {
            if (recebido.Str("sourceModel") != Assinatura(Registros(atual)))
                throw new InvalidDataException("O desenho ativo não corresponde à exportação. Exporte novamente antes de atualizar.");
            if (recebido.Str("sourceDependencies") != Assinatura(Dependencias(recebido)) ||
                recebido.Str("sourceDependencies") != Assinatura(Dependencias(atual)))
                throw new InvalidDataException("As dependências DXF mudaram. Use importação ou substituição para alterar definições.");

            var antes = Entidades(atual);
            var depois = Entidades(recebido);
            if (antes.Count != depois.Count || antes.Keys.Any(h => !depois.ContainsKey(h)))
                throw new InvalidDataException("A atualização requer as mesmas entidades. Use importação ou substituição para mudar a lista.");

            // Registros subordinados a INSERT/POLYLINE não são editáveis isoladamente.
            var filhos = Assinatura(Registros(atual).Except(antes.Values.Cast<object>()));
            if (filhos != Assinatura(Registros(recebido).Except(depois.Values.Cast<object>())))
                throw new InvalidDataException("Os registros subordinados mudaram.");
            var alterados = new List<string>();
            foreach (var item in antes)
            {
                JObj novo = depois[item.Key];
                if (Json.Escrever(item.Value, false) == Json.Escrever(novo, false)) continue;
                string tipo = item.Value.Str("type");
                HashSet<int> codigos;
                if (!Geometria.TryGetValue(tipo, out codigos) || tipo != novo.Str("type"))
                    throw new InvalidDataException("Geometria sem atualização no lugar: " + tipo + " (" + item.Key + ").");
                if (Assinatura(TagsNaoGeometricas(item.Value, codigos)) != Assinatura(TagsNaoGeometricas(novo, codigos)))
                    throw new InvalidDataException("Propriedades, dependências ou metadados mudaram no handle " + item.Key + ".");
                alterados.Add(item.Key);
            }
            return alterados;
        }

        private static List<object> Registros(JObj desenho)
        {
            var secao = desenho.Lista("sections").OfType<JObj>().FirstOrDefault(s => s.Str("name") == "ENTITIES");
            if (secao == null) throw new InvalidDataException("Seção ENTITIES ausente.");
            return secao.Lista("records");
        }

        private static IEnumerable<object> Dependencias(JObj desenho)
        {
            return desenho.Lista("sections").OfType<JObj>().Where(s =>
                s.Str("name") == "TABLES" || s.Str("name") == "BLOCKS" || s.Str("name") == "OBJECTS").Cast<object>();
        }

        private static string Assinatura(IEnumerable<object> dados)
        {
            string json = Json.Escrever(new JObj().Set("registros", dados.ToList()), false);
            using (var sha = SHA256.Create())
                return BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(json))).Replace("-", "");
        }

        private static bool Papel(JObj registro)
        {
            string layout = Valor(registro, 410).Trim();
            return Valor(registro, 67).Trim() == "1" ||
                (layout.Length > 0 && !string.Equals(layout, "Model", StringComparison.OrdinalIgnoreCase));
        }

        private static string Valor(JObj registro, int codigo, string padrao = "")
        {
            int nivel = 0;
            foreach (var par in registro.Lista("tags"))
            {
                IList<object> tag = (IList<object>)par;
                double? n = Json.ComoNumero(tag[0]);
                string v = (string)tag[1];
                if (n == 102 && v.TrimStart().StartsWith("{", StringComparison.Ordinal)) nivel++;
                else if (n == 102 && v.Trim() == "}" && nivel > 0) nivel--;
                else if (nivel == 0 && n == codigo) return v;
            }
            return padrao;
        }

        private static IEnumerable<object> TagsNaoGeometricas(JObj registro, HashSet<int> codigos)
        {
            int extensao = 0;
            bool xdata = false;
            foreach (object obj in registro.Lista("tags"))
            {
                var tag = (IList<object>)obj;
                int codigo = (int)Json.ComoNumero(tag[0]).Value;
                string valor = ((string)tag[1]).Trim();
                if (codigo == 102 && valor.StartsWith("{", StringComparison.Ordinal)) extensao++;
                if (codigo >= 1000) xdata = true;
                if (extensao > 0 || xdata || !codigos.Contains(codigo)) yield return obj;
                if (codigo == 102 && valor == "}" && extensao > 0) extensao--;
            }
        }
    }
}
