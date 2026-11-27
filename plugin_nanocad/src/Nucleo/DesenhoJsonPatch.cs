using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    public sealed class PlanoPatch
    {
        public readonly List<string> Novos = new List<string>();
        public readonly List<string> Modificados = new List<string>();
        public readonly List<string> Excluidos = new List<string>();
    }

    public static partial class DesenhoJsonModelo
    {
        private static List<string> ListaPatch(JObj p, string nome)
        {
            var lista = p.Lista(nome);
            if (lista == null) throw new InvalidDataException("Lista incremental ausente: " + nome);
            var resultado = new List<string>();
            foreach (object o in lista)
            {
                string s = o as string;
                string h = s == null ? "" : NormalizarHandle(s);
                if (h.Length == 0 || resultado.Contains(h))
                    throw new InvalidDataException("Handle inválido ou repetido em " + nome);
                resultado.Add(h);
            }
            return resultado;
        }

        private static bool SimplesSemMetadados(JObj r)
        {
            if (!Geometria.ContainsKey(r.Str("type"))) return false;
            return !r.Lista("tags").Cast<IList<object>>().Any(t => {
                double? c = Json.ComoNumero(t[0]);
                return c == 102 || c >= 1000 || (c >= 340 && c <= 369) || (c >= 390 && c <= 399);
            });
        }

        public static PlanoPatch PlanejarPatch(JObj integral, JObj recebido)
        {
            var p = recebido.Obj("cadsuitePatch");
            if (p == null || p.Num("version") != 1) throw new InvalidDataException("Formato incremental não suportado.");
            var atual = (JObj)Json.Ler(Json.Escrever(integral, false));
            RestringirAoModelo(atual);
            if (p.Str("sourceModel") != Assinatura(Registros(atual)) ||
                p.Str("sourceDependencies") != Assinatura(Dependencias(atual)))
                throw new InvalidDataException("A origem incremental não corresponde ao desenho ativo. Exporte novamente; nenhuma alteração aplicada.");
            var antes = Entidades(atual);
            var depois = Entidades(recebido);
            var plano = new PlanoPatch();
            plano.Novos.AddRange(ListaPatch(p, "added"));
            plano.Modificados.AddRange(ListaPatch(p, "modified"));
            plano.Excluidos.AddRange(ListaPatch(p, "deleted"));
            var todos = plano.Novos.Concat(plano.Modificados).Concat(plano.Excluidos).ToList();
            if (todos.Distinct(StringComparer.OrdinalIgnoreCase).Count() != todos.Count)
                throw new InvalidDataException("Operações sobrepostas no arquivo incremental.");
            if (!new HashSet<string>(depois.Keys, StringComparer.OrdinalIgnoreCase).SetEquals(plano.Novos.Concat(plano.Modificados)) ||
                Registros(recebido).Count != depois.Count || depois.Values.Any(Papel))
                throw new InvalidDataException("Registros incrementais incompatíveis ou subordinados não suportados.");
            var anteriores = p.Lista("before");
            if (anteriores == null || anteriores.Count != plano.Modificados.Count + plano.Excluidos.Count)
                throw new InvalidDataException("Registros anteriores incompletos.");
            var conferidos = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (object item in anteriores)
            {
                var r = item as JObj;
                if (r == null) throw new InvalidDataException("Registro anterior inválido.");
                string h = NormalizarHandle(Valor(r, 5));
                JObj original;
                if (!conferidos.Add(h) || !antes.TryGetValue(h, out original) ||
                    !plano.Modificados.Concat(plano.Excluidos).Contains(h) ||
                    Json.Escrever(r, false) != Json.Escrever(original, false))
                    throw new InvalidDataException("Conflito no registro anterior: " + h);
            }
            foreach (string h in plano.Novos)
            {
                if (antes.ContainsKey(h)) throw new InvalidDataException("Handle novo já existe: " + h);
                var r = depois[h];
                if (!SimplesSemMetadados(r) && !ObservacaoSimples(r, antes))
                    throw new InvalidDataException("Acréscimo com dependências ou metadados não suportados: " + h);
                foreach (var tag in XDataObservacao(r))
                    if (Json.ComoNumero(tag[0]) == 1005 && plano.Excluidos.Contains(NormalizarHandle((string)tag[1])))
                        throw new InvalidDataException("A guia da observação " + h + " aponta para uma entidade marcada para exclusão.");
            }
            foreach (string h in plano.Modificados)
            {
                JObj a, b = depois[h];
                HashSet<int> codigos;
                if (!antes.TryGetValue(h, out a) || a.Str("type") != b.Str("type") ||
                    !Geometria.TryGetValue(a.Str("type"), out codigos) ||
                    Assinatura(TagsNaoGeometricas(a, codigos)) != Assinatura(TagsNaoGeometricas(b, codigos)))
                    throw new InvalidDataException("Modificação não geométrica ou tipo não suportado: " + h);
            }
            foreach (string h in plano.Excluidos)
            {
                JObj r;
                if (!antes.TryGetValue(h, out r) || !SimplesSemMetadados(r))
                    throw new InvalidDataException("Exclusão com metadados não suportada: " + h);
                foreach (JObj s in integral.Lista("sections").OfType<JObj>())
                    foreach (JObj obj in s.Lista("records").OfType<JObj>())
                    {
                        if (s.Str("name") == "HEADER") continue;
                        if (obj.Lista("tags").Cast<IList<object>>().Any(t => {
                            double? c = Json.ComoNumero(t[0]);
                            return (c == 1005 || (c >= 320 && c <= 369) || (c >= 390 && c <= 399) || c == 480 || c == 481) &&
                                NormalizarHandle((string)t[1]) == h;
                        })) throw new InvalidDataException("Exclusão bloqueada por referência a " + h);
                    }
            }
            return plano;
        }

        private static bool ObservacaoSimples(JObj r, Dictionary<string,JObj> antes)
        {
            if (r.Str("type") != "LINE" && r.Str("type") != "MTEXT") return false;
            var tags = r.Lista("tags").Cast<IList<object>>().ToList();
            if (tags.Any(t => { double? c = Json.ComoNumero(t[0]); return c == 102 || (c >= 340 && c <= 369) || (c >= 390 && c <= 399); })) return false;
            var xd = tags.Where(t => Json.ComoNumero(t[0]) >= 1000).ToList();
            if (xd.Count != (r.Str("type") == "LINE" ? 4 : 3)) return false;
            if (Json.ComoNumero(xd[0][0]) != 1001 || (string)xd[0][1] != "CADSUITE_OBS" ||
                Json.ComoNumero(xd[1][0]) != 1000 || Json.ComoNumero(xd[2][0]) != 1000 ||
                string.IsNullOrEmpty((string)xd[1][1])) return false;
            if (r.Str("type") == "MTEXT") return (string)xd[2][1] == "TEXTO";
            return (string)xd[2][1] == "GUIA" && Json.ComoNumero(xd[3][0]) == 1005 &&
                antes.ContainsKey(NormalizarHandle((string)xd[3][1]));
        }

        public static List<IList<object>> XDataObservacao(JObj r)
        {
            var tags = r.Lista("tags").Cast<IList<object>>().Where(t => Json.ComoNumero(t[0]) >= 1000).ToList();
            return tags.Count > 0 && (string)tags[0][1] == "CADSUITE_OBS" ? tags : new List<IList<object>>();
        }
    }
}
