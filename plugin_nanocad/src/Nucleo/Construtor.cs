// Nucleo/Construtor.cs — "Construir do zero" (o usuário aponta poste + anotação e vão + anotação) e os ajustes do pop-up
// (operação, metros, voltar ao lido) que sobrevivem a um "Atualizar" (reler mantendo ajustes). "Reler" descarta tudo.
using System;
using System.Collections.Generic;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    public static class Construtor
    {
        public static Obra NovaObra(string arquivo, RegrasProjeto regras)
        {
            Obra o = new Obra();
            o.Modo = "construir";
            o.Desenho = arquivo ?? "";
            o.Nome = LeitorProjeto.NomeDaObra(arquivo);
            o.Projeto = regras.Codigo;
            o.NomeRegras = regras.Nome;
            return o;
        }

        private static int Proximo(IEnumerable<Elemento> l, char prefixo)
        {
            int max = 0, n;
            foreach (Elemento e in l)
                if (e.Numero != null && e.Numero.Length > 1 && e.Numero[0] == prefixo && int.TryParse(e.Numero.Substring(1), out n)) max = Math.Max(max, n);
            return max + 1;
        }

        /// <summary>Lê os textos escolhidos com o motor do leitor. Sem regra que entenda (projeto sem padrão): o próprio texto, com a entidade pedida.</summary>
        public static TextoLido LerAnotacao(List<TextoDesenho> textos, RegrasProjeto regras, string entidadePadrao)
        {
            if (textos == null || textos.Count == 0) return null;
            List<TextoLido> lidos = LeitorProjeto.ClassificarTextos(textos, regras);
            TextoLido bom = lidos.FirstOrDefault(t => t.R.Entidade == entidadePadrao)
                ?? lidos.FirstOrDefault(t => entidadePadrao == "POSTE" && t.R.Entidade == "ESTRUTURA");
            if (bom != null) return bom;
            TextoLido primeiro = lidos[0];
            TextoLido r = new TextoLido();
            r.T = primeiro.T;
            r.R = new ResultadoLeitor();
            r.R.Entidade = entidadePadrao;
            r.R.Ativo = MotorLeitor.LimparTexto(string.Join(", ", lidos.Select(t => t.T.Texto).ToArray()));
            r.R.Operacao = primeiro.R.Operacao == "0" || string.IsNullOrEmpty(primeiro.R.Operacao) ? "I" : primeiro.R.Operacao;
            return r;
        }

        public static Poste AdicionarPoste(Obra obra, PosteDesenho bloco, List<TextoDesenho> textos, RegrasProjeto regras)
        {
            Poste p = bloco != null && !string.IsNullOrEmpty(bloco.Handle) ? obra.Postes.FirstOrDefault(x => x.Handle == bloco.Handle) : null;
            if (p == null)
            {
                p = new Poste();
                p.Numero = "P" + Proximo(obra.Postes.Cast<Elemento>(), 'P');
                obra.Postes.Add(p);
            }
            p.Manual = true;
            if (bloco != null) { p.Handle = bloco.Handle; p.X = bloco.X; p.Y = bloco.Y; p.Bloco = bloco.Bloco ?? ""; }
            TextoLido tl = LerAnotacao(textos, regras, "POSTE");
            if (tl != null && tl.T.Altura > 0 && obra.Postes.Count == 1) obra.AlturaTexto = tl.T.Altura;
            if (tl != null)
            {
                LeitorProjeto.Anotar(p, tl);
                if (bloco == null) { p.X = tl.T.X; p.Y = tl.T.Y; p.Handle = ""; }
            }
            Recalcular(obra, regras);
            return p;
        }

        /// <summary>Postes da obra mais próximos das pontas da linha (até `tol`), como sugestão para o vão.</summary>
        public static void SugerirPostes(Obra obra, List<Ponto> pts, double tol, out Poste a, out Poste b)
        {
            a = MaisPerto(obra, pts[0], tol, null);
            b = MaisPerto(obra, pts[pts.Count - 1], tol, a);
        }

        private static Poste MaisPerto(Obra obra, Ponto q, double tol, Poste exceto)
        {
            Poste melhor = null;
            double md = double.MaxValue;
            foreach (Poste p in obra.Postes)
            {
                if (p == exceto) continue;
                double d = q.Dist(p.Ponto);
                if (d <= tol && d < md) { md = d; melhor = p; }
            }
            return melhor;
        }

        public static Vao AdicionarVao(Obra obra, LinhaDesenho linha, List<TextoDesenho> textos, Poste a, Poste b, RegrasProjeto regras)
        {
            Vao v = obra.Vaos.FirstOrDefault(x => x.LinhaHandle == linha.Handle && x.Trecho == 0);
            if (v == null)
            {
                v = new Vao();
                v.Numero = "V" + Proximo(obra.Vaos.Cast<Elemento>(), 'V');
                obra.Vaos.Add(v);
            }
            v.Manual = true;
            v.LinhaHandle = linha.Handle;
            v.Pontos = new List<Ponto>(linha.Pontos);
            v.LinhaLayer = linha.Layer;
            v.LinhaCor = linha.Cor;
            v.Tipo = LeitorProjeto.ClassificarRede(linha.Layer, linha.Cor) ?? LeitorProjeto.TipoPelaCor(linha.Cor) ?? "";
            v.PosteA = a != null ? a.Numero : null;
            v.PosteB = b != null ? b.Numero : null;
            TextoLido tl = LerAnotacao(textos, regras, "CABO");
            if (tl != null) LeitorProjeto.Anotar(v, tl);
            Recalcular(obra, regras);
            return v;
        }

        public static void Remover(Obra obra, Elemento e, RegrasProjeto regras)
        {
            Poste p = e as Poste;
            if (p != null)
            {
                obra.Postes.Remove(p);
                foreach (Vao v in obra.Vaos)
                {
                    if (v.PosteA == p.Numero) { v.PosteA = null; v.VinculoA = null; }
                    if (v.PosteB == p.Numero) { v.PosteB = null; v.VinculoB = null; }
                }
            }
            if (e is Vao) obra.Vaos.Remove((Vao)e);
            if (e is Avulso) obra.Avulsos.Remove((Avulso)e);
            Recalcular(obra, regras);
        }

        /// <summary>Depois de uma mudança: vínculos automáticos de novo (mantendo os manuais), parâmetros e conferência.</summary>
        public static void Recalcular(Obra obra, RegrasProjeto regras)
        {
            Vinculador.Vincular(obra, regras, true);
            Conferencia.Avaliar(obra, regras);
        }

        // ── ajustes do pop-up ──
        public static void AjustarVao(Obra obra, Vao v, string operacao, double? metros, RegrasProjeto regras)
        {
            if (!string.IsNullOrEmpty(operacao)) v.Operacao = operacao;
            if (metros.HasValue) v.Metros = metros;
            if (v.Cabo != null) v.InfoCabo = null;
            Recalcular(obra, regras);
        }

        public static void AjustarPoste(Obra obra, Poste p, string operacao, RegrasProjeto regras)
        {
            if (!string.IsNullOrEmpty(operacao)) p.Operacao = operacao;
            Recalcular(obra, regras);
        }

        public static void VoltarAoLido(Obra obra, Elemento e, RegrasProjeto regras)
        {
            e.Operacao = e.OperacaoLida;
            Vao v = e as Vao;
            if (v != null)
            {
                v.Metros = v.MetrosLidos;
                if (v.VinculoA != null && v.VinculoA.Manual) v.VinculoA = null;
                if (v.VinculoB != null && v.VinculoB.Manual) v.VinculoB = null;
            }
            Recalcular(obra, regras);
        }

        /// <summary>Liga à mão a ponta (A ou B) do vão à estrutura k do poste daquela ponta. Devolve o motivo se as regras não deixam.</summary>
        public static string VincularManual(Obra obra, Vao v, bool pontaA, int k, RegrasProjeto regras)
        {
            Poste p = obra.PostePorNumero(pontaA ? v.PosteA : v.PosteB);
            if (p == null) return "essa ponta do vão não tem poste";
            if (k < 0 || k >= p.Lido.Estruturas.Count) return "o poste " + p.Numero + " não tem essa estrutura";
            Vinculo vi = new Vinculo();
            vi.Poste = p.Numero; vi.Estrutura = p.Lido.Estruturas[k].Nome; vi.K = k; vi.Manual = true;
            if (pontaA) v.VinculoA = vi; else v.VinculoB = vi;
            Recalcular(obra, regras);
            return null;
        }

        /// <summary>"Atualizar": a obra relida recebe de volta os ajustes da anterior (casando por handle da linha/bloco ou do texto).</summary>
        public static void ReaplicarAjustes(Obra nova, Obra antiga, RegrasProjeto regras)
        {
            if (antiga == null) return;
            foreach (Poste pa in antiga.Postes)
            {
                Poste pn = nova.Postes.FirstOrDefault(x => (!string.IsNullOrEmpty(pa.Handle) && x.Handle == pa.Handle) || (!string.IsNullOrEmpty(pa.TextoHandle) && x.TextoHandle == pa.TextoHandle));
                if (pn != null && pa.OperacaoAjustada && pn.TextoHandle == pa.TextoHandle) pn.Operacao = pa.Operacao;
            }
            foreach (Vao va in antiga.Vaos)
            {
                Vao vn = nova.Vaos.FirstOrDefault(x => x.LinhaHandle == va.LinhaHandle && x.Trecho == va.Trecho)
                      ?? nova.Vaos.FirstOrDefault(x => !string.IsNullOrEmpty(va.TextoHandle) && x.TextoHandle == va.TextoHandle);
                if (vn == null || vn.TextoHandle != va.TextoHandle) continue;
                if (va.OperacaoAjustada) vn.Operacao = va.Operacao;
                if (va.Metros.HasValue && va.Metros != va.MetrosLidos) vn.Metros = va.Metros;
                foreach (bool a in new[] { true, false })
                {
                    Vinculo vi = a ? va.VinculoA : va.VinculoB;
                    if (vi == null || !vi.Manual) continue;
                    Poste pAnt = antiga.PostePorNumero(vi.Poste);
                    Poste pNovo = pAnt == null ? null : nova.Postes.FirstOrDefault(x => x.Handle == pAnt.Handle && x.TextoHandle == pAnt.TextoHandle);
                    if (pNovo == null || (a ? vn.PosteA : vn.PosteB) != pNovo.Numero) continue;
                    Vinculo nv = new Vinculo();
                    nv.Poste = pNovo.Numero; nv.Estrutura = vi.Estrutura; nv.K = vi.K; nv.Manual = true;
                    if (a) vn.VinculoA = nv; else vn.VinculoB = nv;
                }
            }
            Recalcular(nova, regras);
        }
    }
}
