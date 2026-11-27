using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    public static partial class DesenhoJsonModelo
    {
        private static string IdentidadeDwg(JObj desenho)
        {
            return NormalizarGuid(VariavelDxf(desenho, "$FINGERPRINTGUID", 2));
        }

        private static string NormalizarGuid(string texto)
        {
            if (string.IsNullOrEmpty(texto)) return "";
            try
            {
                var guid = new Guid(texto);
                return guid == Guid.Empty ? "" : guid.ToString("D");
            }
            catch (FormatException) { return ""; }
            catch (OverflowException) { return ""; }
        }

        private static bool ConferirOrigemPatch(JObj patch, JObj atual, bool semantico, PlanoPatch plano)
        {
            if (semantico)
            {
                string id = NormalizarGuid(patch.Str("sourceDrawingId"));
                if (id.Length == 0)
                    throw new InvalidDataException("Patch v2 sem sourceDrawingId válido. Exporte o DWG de origem novamente.");
                if (id != IdentidadeDwg(atual))
                    throw new InvalidDataException("O identificador persistente do patch difere do desenho ativo ou está indisponível no DXF.");
            }

            var divergencias = new List<string>();
            bool compararModelo = !semantico || patch.Has("sourceModel");
            bool modeloConfere = compararModelo && patch.Str("sourceModel") == Assinatura(Registros(atual));
            if (compararModelo && !modeloConfere)
                divergencias.Add("sourceModel (Model Space)");
            if ((!semantico || patch.Has("sourceDependencies")) &&
                patch.Str("sourceDependencies") != Assinatura(Dependencias(atual)))
                divergencias.Add("sourceDependencies (tabelas, blocos e objetos)");
            if (divergencias.Count == 0) return modeloConfere;

            string detalhe = string.Join(", ", divergencias.ToArray());
            if (!semantico)
                throw new InvalidDataException("Origem incremental v1 divergente em " + detalhe
                    + ". Reexporte o desenho ou gere um patch v2 com identidade persistente.");
            plano.Diagnosticos.Add("Assinaturas integrais divergentes em " + detalhe + ". Validação individual em uso.");
            return modeloConfere;
        }

        private static void ConferirReferenciasPatch(JObj patch, PlanoPatch plano,
            Dictionary<string, JObj> antes, Dictionary<string, JObj> depois, List<NotaPostePatch> novasNotas, bool modeloConfere)
        {
            var requeridas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string h in plano.Novos)
                foreach (var tag in XDataObservacao(depois[h]))
                    if (Json.ComoNumero(tag[0]) == 1005) requeridas.Add(NormalizarHandle((string)tag[1]));
            var alvosNota = new HashSet<string>(novasNotas.SelectMany(n => n.Alvos), StringComparer.OrdinalIgnoreCase);
            foreach (string h in alvosNota)
                if (!antes.ContainsKey(h) || plano.Excluidos.Contains(h))
                    throw new InvalidDataException("Alvo ANOTAPOSTE " + h + " ausente ou marcado para exclusão.");

            var referencias = patch.Lista("references");
            if (referencias == null)
                throw new InvalidDataException("O patch v2 precisa de um registro em references para cada alvo distinto de guia.");

            var conferidas = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (object item in referencias)
            {
                var registro = item as JObj;
                if (registro == null || registro.Lista("tags") == null)
                    throw new InvalidDataException("Registro de referência inválido no patch v2.");
                string h = NormalizarHandle(Valor(registro, 5));
                JObj entidade;
                if (!conferidas.Add(h) || !(requeridas.Contains(h) || alvosNota.Contains(h)) ||
                    !antes.TryGetValue(h, out entidade) || !RegistroEquivalente(registro, entidade))
                    throw new InvalidDataException("Conflito no alvo da guia " + h + ": a entidade de referência mudou, foi removida ou é inválida.");
            }
            if (!requeridas.IsSubsetOf(conferidas))
                throw new InvalidDataException("O patch v2 precisa de um registro em references para cada alvo distinto de guia.");
            alvosNota.ExceptWith(conferidas);
            if (alvosNota.Count == 0) return;
            // O exportador anterior omitia os alvos de ANOTAPOSTE. A assinatura do modelo inteiro
            // só substitui os snapshots se comprovar que o modelo ainda corresponde à base.
            if (!modeloConfere)
                throw new InvalidDataException("ANOTAPOSTE sem registro de referência em references: " +
                    string.Join(", ", alvosNota.ToArray()) + ". Reexporte o patch com os registros originais dos alvos.");
            plano.Diagnosticos.Add("Alvos ANOTAPOSTE conferidos pela assinatura original do Model Space.");
        }

        private static bool RegistroEquivalente(JObj a, JObj b)
        {
            return string.Equals(a.Str("type"), b.Str("type"), StringComparison.OrdinalIgnoreCase)
                && a.Lista("tags") != null && TagsEquivalentes(a.Lista("tags"), b.Lista("tags"));
        }

        private static bool TagsEquivalentes(IEnumerable<object> a, IEnumerable<object> b)
        {
            return a.Select(TagSemantica).SequenceEqual(b.Select(TagSemantica));
        }

        private static KeyValuePair<int, string> TagSemantica(object obj)
        {
            var tag = (IList<object>)obj;
            int codigo = (int)Json.ComoNumero(tag[0]).Value;
            string valor = (string)tag[1];
            if (codigo == 5 || codigo == 105 || (codigo >= 320 && codigo <= 369)
                || (codigo >= 390 && codigo <= 399) || codigo == 480 || codigo == 481 || codigo == 1005)
            {
                string handle = NormalizarHandle(valor);
                if (handle.Length > 0) valor = handle;
            }
            else if ((codigo >= 10 && codigo < 60) || (codigo >= 110 && codigo < 150)
                || (codigo >= 210 && codigo < 240) || (codigo >= 460 && codigo < 470)
                || (codigo >= 1010 && codigo <= 1059))
            {
                double numero;
                if (double.TryParse(valor, NumberStyles.Float, CultureInfo.InvariantCulture, out numero)
                    && !double.IsNaN(numero) && !double.IsInfinity(numero))
                    valor = numero == 0 ? "0" : numero.ToString("R", CultureInfo.InvariantCulture);
            }
            else if ((codigo >= 60 && codigo < 100) || (codigo >= 160 && codigo < 180)
                || (codigo >= 270 && codigo < 300)
                || (codigo >= 370 && codigo < 390) || (codigo >= 400 && codigo <= 409)
                || (codigo >= 420 && codigo <= 429) || (codigo >= 440 && codigo <= 459)
                || codigo == 1070 || codigo == 1071)
            {
                long numero;
                if (long.TryParse(valor, NumberStyles.Integer, CultureInfo.InvariantCulture, out numero))
                    valor = numero.ToString(CultureInfo.InvariantCulture);
            }
            return new KeyValuePair<int, string>(codigo, valor);
        }

        private static void ConferirDependenciasPatch(JObj atual, JObj recebido, IEnumerable<JObj> entidades)
        {
            var tabelasAtuais = RegistrosDeTabelas(atual);
            var tabelasRecebidas = RegistrosDeTabelas(recebido);
            var cache = new Dictionary<string, JObj>(StringComparer.OrdinalIgnoreCase);
            foreach (var r in entidades)
            {
                string layer = Valor(r, 8);
                var camada = ConferirDefinicao(tabelasAtuais, tabelasRecebidas, cache, "LAYER", layer.Length == 0 ? "0" : layer);
                string linha = Valor(r, 6);
                if (linha.Length == 0 || linha.Equals("BYLAYER", StringComparison.OrdinalIgnoreCase))
                    linha = Valor(camada, 6);
                if (linha.Length != 0 && !linha.Equals("BYLAYER", StringComparison.OrdinalIgnoreCase)
                    && !linha.Equals("BYBLOCK", StringComparison.OrdinalIgnoreCase))
                {
                    var tipoLinha = ConferirDefinicao(tabelasAtuais, tabelasRecebidas, cache, "LTYPE", linha);
                    foreach (var tag in tipoLinha.Lista("tags").Cast<IList<object>>())
                    {
                        if (Json.ComoNumero(tag[0]) != 340) continue;
                        string handle = NormalizarHandle((string)tag[1]);
                        if (handle.Length == 0) continue;
                        var estilos = tabelasAtuais.Where(t => t.Str("type") == "STYLE"
                            && NormalizarHandle(Valor(t, 5)) == handle).ToList();
                        if (estilos.Count != 1)
                            throw new InvalidDataException("Estilo de texto " + handle + " usado pelo tipo de linha " + linha + " indisponível.");
                        ConferirDefinicao(tabelasAtuais, tabelasRecebidas, cache, "STYLE", Valor(estilos[0], 2));
                    }
                }
                if (r.Str("type") == "MTEXT")
                {
                    string estilo = Valor(r, 7);
                    ConferirDefinicao(tabelasAtuais, tabelasRecebidas, cache, "STYLE", estilo.Length == 0 ? "STANDARD" : estilo);
                }
            }
        }

        private static List<JObj> RegistrosDeTabelas(JObj desenho)
        {
            return desenho.Lista("sections").OfType<JObj>().Where(s => s.Str("name") == "TABLES")
                .SelectMany(s => s.Lista("records").OfType<JObj>()).ToList();
        }

        private static JObj ConferirDefinicao(List<JObj> a, List<JObj> b, Dictionary<string, JObj> cache, string tipo, string nome)
        {
            string chave = tipo + "/" + nome;
            JObj validada;
            if (cache.TryGetValue(chave, out validada)) return validada;
            var originais = a.Where(r => r.Str("type") == tipo && Valor(r, 2).Equals(nome, StringComparison.OrdinalIgnoreCase)).ToList();
            var recebidas = b.Where(r => r.Str("type") == tipo && Valor(r, 2).Equals(nome, StringComparison.OrdinalIgnoreCase)).ToList();
            if (originais.Count != 1 || recebidas.Count != 1 || !RegistroEquivalente(originais[0], recebidas[0]))
                throw new InvalidDataException("Conflito na dependência " + tipo + " " + nome + ": definição ausente, ambígua ou alterada.");
            cache.Add(chave, originais[0]);
            return originais[0];
        }
    }
}
