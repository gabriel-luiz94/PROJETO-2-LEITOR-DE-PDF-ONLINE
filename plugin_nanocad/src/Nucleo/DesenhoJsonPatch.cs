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
        public readonly List<string> Diagnosticos = new List<string>();
        public readonly Dictionary<string, List<IList<object>>> XDataNovos = new Dictionary<string, List<IList<object>>>(StringComparer.OrdinalIgnoreCase);
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

        private static bool SimplesSemMetadados(JObj r, bool aceitarTexto = false)
        {
            if (!Geometria.ContainsKey(r.Str("type")) && !(aceitarTexto && r.Str("type") == "MTEXT")) return false;
            return !r.Lista("tags").Cast<IList<object>>().Any(t => {
                double? c = Json.ComoNumero(t[0]);
                return c == 102 || c >= 1000 || (c >= 340 && c <= 369) || (c >= 390 && c <= 399);
            });
        }

        public static PlanoPatch PlanejarPatch(JObj integral, JObj recebido)
        {
            var p = recebido.Obj("cadsuitePatch");
            if (p == null || (p.Num("version") != 1 && p.Num("version") != 2))
                throw new InvalidDataException("Formato incremental não suportado.");
            bool semantico = p.Num("version") == 2;
            var plano = new PlanoPatch();
            // O filtro altera somente as listas de registros de ENTITIES e BLOCKS.
            var secoes = new List<object>();
            foreach (JObj secao in integral.Lista("sections").OfType<JObj>())
            {
                string nome = secao.Str("name");
                if (nome == "ENTITIES" || nome == "BLOCKS")
                {
                    var copia = new JObj();
                    foreach (var campo in secao) copia.Set(campo.Key, campo.Value);
                    secoes.Add(copia.Set("records", new List<object>(secao.Lista("records"))));
                }
                else secoes.Add(secao);
            }
            var atual = new JObj().Set("sections", secoes);
            RestringirAoModelo(atual);
            bool modeloConfere = ConferirOrigemPatch(p, atual, semantico, plano);
            var antes = Entidades(atual);
            var depois = Entidades(recebido);
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
                    !(semantico ? RegistroEquivalente(r, original) : Json.Escrever(r, false) == Json.Escrever(original, false)))
                    throw new InvalidDataException("Conflito na entidade " + h + ": registro anterior mudou, foi removido ou é inválido.");
            }
            var excluidos = new HashSet<string>(plano.Excluidos, StringComparer.OrdinalIgnoreCase);
            var novosPoste = semantico ? LerGruposAnotaposte(plano.Novos, depois) : new List<NotaPostePatch>();
            var excluidosPoste = semantico ? LerGruposAnotaposte(plano.Excluidos, antes) : new List<NotaPostePatch>();
            var gruposPoste = new HashSet<string>(novosPoste.Concat(excluidosPoste).Select(n => n.Grupo), StringComparer.Ordinal);
            if (gruposPoste.Count > 0 && antes.Any(e => !excluidos.Contains(e.Key) && gruposPoste.Contains(GrupoAnotaposte(e.Value))))
                throw new InvalidDataException("Já existe uma peça do grupo ANOTAPOSTE que não foi incluída na substituição.");
            foreach (var nota in novosPoste) plano.XDataNovos.Add(nota.Handle, nota.Tags);
            if (semantico) ConferirReferenciasPatch(p, plano, antes, depois, novosPoste, modeloConfere);
            foreach (string h in plano.Novos)
            {
                if (antes.ContainsKey(h)) throw new InvalidDataException("Handle novo já existe: " + h);
                var r = depois[h];
                if (plano.XDataNovos.ContainsKey(h)) continue;
                if (!SimplesSemMetadados(r) && !ObservacaoSimples(r, antes))
                    throw new InvalidDataException("Acréscimo com dependências ou metadados não suportados: " + h);
                var tags = XDataObservacao(r);
                if (tags.Count > 0) plano.XDataNovos.Add(h, tags);
                foreach (var tag in tags)
                    if (Json.ComoNumero(tag[0]) == 1005 && plano.Excluidos.Contains(NormalizarHandle((string)tag[1])))
                        throw new InvalidDataException("A guia da observação " + h + " aponta para uma entidade marcada para exclusão.");
            }
            foreach (string h in plano.Modificados)
            {
                JObj a, b = depois[h];
                HashSet<int> codigos;
                if (semantico && TemAnotaposte(b))
                    throw new InvalidDataException("Envie ANOTAPOSTE como substituição do grupo completo, não como alteração individual: " + h);
                if (!antes.TryGetValue(h, out a) || a.Str("type") != b.Str("type") ||
                    !Geometria.TryGetValue(a.Str("type"), out codigos) ||
                    (semantico ? !TagsEquivalentes(TagsNaoGeometricas(a, codigos), TagsNaoGeometricas(b, codigos))
                        : Assinatura(TagsNaoGeometricas(a, codigos)) != Assinatura(TagsNaoGeometricas(b, codigos))))
                    throw new InvalidDataException("Modificação não geométrica ou tipo não suportado: " + h);
            }
            foreach (string h in plano.Excluidos)
            {
                JObj r;
                if (!antes.TryGetValue(h, out r) ||
                    !(SimplesSemMetadados(r, novosPoste.Count > 0) || excluidosPoste.Any(n => n.Handle == h)))
                    throw new InvalidDataException("Exclusão com metadados não suportada: " + h);
            }
            if (excluidos.Count > 0) ConferirExclusoesPatch(integral, excluidos, semantico);
            if (semantico) ConferirDependenciasPatch(atual, recebido, depois.Values);
            return plano;
        }

        private static void ConferirExclusoesPatch(JObj integral, HashSet<string> excluidos, bool semantico)
        {
            foreach (JObj secao in integral.Lista("sections").OfType<JObj>())
            {
                string nome = secao.Str("name");
                if (nome == "HEADER") continue;
                foreach (JObj registro in secao.Lista("records").OfType<JObj>())
                {
                    if (semantico && nome == "ENTITIES" && excluidos.Contains(NormalizarHandle(Valor(registro, 5)))) continue;
                    foreach (var tag in registro.Lista("tags").Cast<IList<object>>())
                    {
                        double? c = Json.ComoNumero(tag[0]);
                        if (!(c == 1005 || (c >= 320 && c <= 369) || (c >= 390 && c <= 399) || c == 480 || c == 481)) continue;
                        string handle = NormalizarHandle((string)tag[1]);
                        if (excluidos.Contains(handle)) throw new InvalidDataException("Exclusão bloqueada por referência a " + handle);
                    }
                }
            }
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
