// Conferência da obra já classificada. Os avisos não modificam dados nem regras do orçamento.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    public static class Auditoria
    {
        public static string Origem(string handle, int parte, string texto)
        {
            return Json.Escrever(new List<object> { handle, parte, texto }, false);
        }

        public static string Alvo(Elemento e)
        {
            Poste p = e as Poste;
            Vao v = e as Vao;
            if (p != null) return "P:" + p.Handle;
            if (v != null) return "V:" + v.LinhaHandle + "/" + v.Trecho;
            return e.Numero;
        }

        public static string Chave(Elemento e)
        {
            return string.IsNullOrEmpty(e.TextoHandle) ? Alvo(e) + "/" + e.Numero : e.TextoHandle + "/" + e.ParteTexto;
        }

        private static string Metros(double? m)
        {
            return m.HasValue ? m.Value.ToString("0.##", CultureInfo.InvariantCulture) + " m" : "—";
        }

        public static string Assinatura(Elemento e)
        {
            Vao v = e as Vao;
            List<object> dados = new List<object> { Alvo(e), e.TextoOriginal, e.LayerAtual, e.TextoCor, e.Entidade, e.AtivoLido,
                e.OperacaoLida, e.Operacao, e.Excluido, e.Estado, e.Mensagens.Cast<object>().ToList(), e.Candidatos.Cast<object>().ToList() };
            if (v != null)
            {
                dados.Add(v.Metros); dados.Add(v.MetrosLidos); dados.Add(v.PosteA); dados.Add(v.PosteB);
                dados.Add(v.Pontos.Select(p => (object)new List<object> { p.X, p.Y }).ToList());
                foreach (Vinculo vi in new[] { v.VinculoA, v.VinculoB })
                    dados.Add(vi == null ? null : new List<object> { vi.Poste, vi.K, vi.Estrutura });
            }
            return Json.Escrever(dados, false);
        }

        public static JObj Fotografar(Obra o)
        {
            JObj r = new JObj();
            foreach (Elemento e in o.Elementos) r.Set(Chave(e), Assinatura(e));
            return r;
        }

        public static IEnumerable<string> Motivos(Obra o, Elemento e, RegrasProjeto regras, IDictionary<string,string> layers)
        {
            foreach (string m in e.Mensagens) yield return m;
            if (e.Candidatos.Count > 1 && o.Associacoes.Str(Origem(e.TextoHandle, e.ParteTexto, e.TextoOriginal)) != Alvo(e))
                yield return "Mais de um destino possível para a anotação";
            string opLayer;
            if (layers.TryGetValue(e.LayerAtual ?? "",out opLayer) && opLayer != e.Operacao)
                yield return "Layer indica " + opLayer + " · obra " + e.Operacao;
            else if (e.OperacaoAjustada && e.LayerAplicada == null)
                yield return "Operação lida " + e.OperacaoLida + " · operação da obra " + e.Operacao;
            Vao v = e as Vao;
            if (v != null && v.TemAnotacao && v.Pontos.Count >= 2 && v.MetrosLidos.HasValue && regras.Leitura.ToleranciaMetros > 0
                && Math.Abs(Geometria.Comprimento(v.Pontos) - v.MetrosLidos.Value) > regras.Leitura.ToleranciaMetros)
                yield return "Metragem escrita difere da geometria além da tolerância do projeto";
        }

        public static string Detalhe(Obra o, Elemento e)
        {
            Vao v = e as Vao;
            Poste p = e as Poste;
            string saida = e.Excluido || (!e.TemAnotacao && !(e is Avulso)) ? "Fora da exportação" :
                Relatorios.PainelDe(e) + "  ·  " + e.Entidade + "  ·  " + e.Operacao + "  ·  " + (v == null ? e.AtivoLido : v.Ativo);
            List<string> linhas = new List<string> { e.Numero + "   " + e.TextoOriginal,
                "Layer " + e.LayerAtual + "  ·  cor " + e.TextoCor + "  ·  texto " + e.TextoHandle + "  ·  operação lida " + e.OperacaoLida,
                "OrçaCore   " + saida };
            if (v != null)
            {
                linhas.Add("Postes " + v.PosteA + " → " + v.PosteB + "  ·  anotado " + Metros(v.MetrosLidos) + "  ·  geometria " +
                    Metros(v.Pontos.Count >= 2 ? (double?)Geometria.Comprimento(v.Pontos) : null) + "  ·  na obra " + Metros(v.Metros ?? v.MetrosLidos));
                linhas.Add("Vínculos   " + string.Join(", ", v.Vinculos.Select(x => x.Poste + " · " + x.Estrutura).ToArray()));
            }
            if (p != null)
            {
                string raiz = p.IrmaoDe ?? p.Numero;
                linhas.Add("No poste " + raiz + "   " + string.Join(" · ", o.GrupoDoPoste(p).Select(x => x.Numero + " " + x.AtivoLido).ToArray()));
                linhas.Add("Vãos   " + string.Join(", ", o.Vaos.Where(x => x.PosteA == raiz || x.PosteB == raiz).Select(x => x.Numero).ToArray()));
            }
            if (e.OrigemMetros != null)
                linhas.Add("Texto da metragem   " + e.OrigemMetros.Texto + " · layer " + e.OrigemMetros.Layer + " · handle " + e.OrigemMetros.Handle);
            if (e.Candidatos.Count > 1)
                linhas.Add("Candidatos   " + string.Join(", ", o.Elementos.Where(x => x.IrmaoDe == null && e.Candidatos.Contains(Alvo(x))).Select(x => x.Numero).ToArray()));
            linhas.AddRange(e.Mensagens);
            return string.Join(Environment.NewLine, linhas.ToArray());
        }

        public static bool Suspeito(Obra o, TextoIgnorado t, RegrasProjeto regras, HashSet<string> textosLidos, IDictionary<string,string> layers)
        {
            if (t.Tipo == "BLOCO_DEF" || t.Tipo == "BLOCO" || string.IsNullOrEmpty(t.Texto) ||
                (t.Motivo ?? "").StartsWith("placa do poste") || t.Motivo == "parte de outra anotação classificada" || (t.Motivo ?? "").StartsWith("ramal/IP") || (t.Layer ?? "").IndexOf("LEGENDA", StringComparison.OrdinalIgnoreCase) >= 0 ||
                (t.Layer ?? "").IndexOf("CARIMBO", StringComparison.OrdinalIgnoreCase) >= 0) return false;
            string texto = t.Texto.Trim();
            if (!textosLidos.Contains(texto) && !Construtor.EhComprimento(texto) &&
                !layers.ContainsKey(t.Layer ?? "")) return false;
            Ponto pt = new Ponto(t.X, t.Y);
            double raio = o.AlturaTexto * regras.Leitura.RaioAlturas;
            return o.Postes.Any(p => p.Ponto.Dist(pt) <= raio) || o.Vaos.Any(v => v.Pontos.Count >= 2 && Geometria.Projetar(pt, v.Pontos).Distancia <= raio)
                || o.Elementos.Any(e => e.GuiaUsada != null && e.GuiaUsada.Count >= 2 && Geometria.Projetar(pt, e.GuiaUsada).Distancia <= raio);
        }
    }
}
