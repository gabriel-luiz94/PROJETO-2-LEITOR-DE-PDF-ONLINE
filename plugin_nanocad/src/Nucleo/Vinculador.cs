// Nucleo/Vinculador.cs — vínculo automático vão ⇄ estrutura do poste, parâmetros de cada estrutura e a conferência (Ok/Conferir/Erro).
// O vínculo usa as regras de vínculo do projeto (RegrasVinculo); só entram cabos e estruturas de operação I/M (e *I/*M) — R e *R ficam fora.
using System;
using System.Collections.Generic;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    public static class Vinculador
    {
        public static string Chave(string poste, int k) { return poste + "#" + k; }

        public static bool PosteParticipa(Poste p)
        {
            return p != null && p.TemAnotacao && RegrasVinculo.ParticipaVinculo(p.Operacao);
        }

        private static void PrepararCabos(Obra obra, RegrasProjeto regras)
        {
            foreach (Vao v in obra.Vaos)
            {
                v.PosteObjA = obra.PostePorNumero(v.PosteA);
                v.PosteObjB = obra.PostePorNumero(v.PosteB);
                v.InfoCabo = v.TemAnotacao && v.Cabo != null && !string.IsNullOrEmpty(v.Cabo.Codigo)
                    ? regras.Vinculo.InfoCabo(v.Cabo.Codigo, v.Operacao, v.Cabo) : null;
            }
        }

        /// <summary>Refaz os vínculos automáticos (os manuais ficam, se `manterManuais`) e recalcula os parâmetros.</summary>
        public static void Vincular(Obra obra, RegrasProjeto regras, bool manterManuais)
        {
            PrepararCabos(obra, regras);
            Dictionary<string, List<Ocupante>> ocup = new Dictionary<string, List<Ocupante>>();
            foreach (Vao v in obra.Vaos)
            {
                v.MotivoVinculo = null;
                if (!manterManuais || v.VinculoA == null || !v.VinculoA.Manual) v.VinculoA = null;
                if (!manterManuais || v.VinculoB == null || !v.VinculoB.Manual) v.VinculoB = null;
                if (!v.ParticipaVinculo || v.InfoCabo == null) { v.VinculoA = null; v.VinculoB = null; continue; }
                foreach (Vinculo vi in v.Vinculos) Ocupar(ocup, vi, v.InfoCabo);
            }
            foreach (Vao v in obra.Vaos)
            {
                if (!v.ParticipaVinculo || v.InfoCabo == null || !v.InfoCabo.Vincula) continue;
                int tem = v.Vinculos.Count();
                for (int ponta = 0; ponta < 2 && tem < v.InfoCabo.Pontas; ponta++)
                {
                    if ((ponta == 0 ? v.VinculoA : v.VinculoB) != null) continue;
                    Poste p = ponta == 0 ? v.PosteObjA : v.PosteObjB;
                    if (!PosteParticipa(p)) continue;
                    for (int k = 0; k < p.Lido.Estruturas.Count; k++)
                    {
                        InfoEstrutura e = regras.Vinculo.InfoEstrutura(p.Lido.Estruturas[k].Nome);
                        List<Ocupante> l = Lista(ocup, Chave(p.Numero, k));
                        string motivo = RegrasVinculo.PodeComOcupantes(v.InfoCabo, e, l, false);
                        if (motivo != null) { v.MotivoVinculo = motivo; continue; }
                        Vinculo vi = new Vinculo();
                        vi.Poste = p.Numero; vi.Estrutura = e.Codigo; vi.K = k;
                        if (ponta == 0) v.VinculoA = vi; else v.VinculoB = vi;
                        Ocupar(ocup, vi, v.InfoCabo);
                        tem++;
                        break;
                    }
                }
            }
            AtualizarParametros(obra, regras);
        }

        private static List<Ocupante> Lista(Dictionary<string, List<Ocupante>> ocup, string chave)
        {
            List<Ocupante> l;
            if (!ocup.TryGetValue(chave, out l)) { l = new List<Ocupante>(); ocup[chave] = l; }
            return l;
        }

        private static void Ocupar(Dictionary<string, List<Ocupante>> ocup, Vinculo vi, InfoCabo c)
        {
            Ocupante o = new Ocupante();
            o.Info = c;
            Lista(ocup, Chave(vi.Poste, vi.K)).Add(o);
        }

        /// <summary>Ocupantes de cada estrutura (na ordem dos vãos) → parâmetros. Não mexe nos vínculos.</summary>
        public static void AtualizarParametros(Obra obra, RegrasProjeto regras)
        {
            PrepararCabos(obra, regras);
            Dictionary<string, List<Ocupante>> ocup = new Dictionary<string, List<Ocupante>>();
            foreach (Vao v in obra.Vaos)
                if (v.InfoCabo != null) foreach (Vinculo vi in v.Vinculos) Ocupar(ocup, vi, v.InfoCabo);
            foreach (Poste p in obra.Postes)
            {
                p.ParametrosEstruturas.Clear();
                p.CabosPorEstrutura.Clear();
                p.AvisosParametros.Clear();
                for (int k = 0; k < p.Lido.Estruturas.Count; k++)
                {
                    List<Ocupante> l = Lista(ocup, Chave(p.Numero, k));
                    InfoEstrutura e = regras.Vinculo.InfoEstrutura(p.Lido.Estruturas[k].Nome);
                    p.ParametrosEstruturas.Add(regras.Vinculo.Parametros(e, l, p.AvisosParametros));
                    p.CabosPorEstrutura.Add(l.Count);
                }
            }
        }

        /// <summary>Posição do vão entre os cabos da estrutura (1-based) e a capacidade, para "N1 (slot 1/2)".</summary>
        public static string Slot(Obra obra, Vao vao, Vinculo vi, RegrasProjeto regras)
        {
            int pos = 0, n = 0;
            foreach (Vao v in obra.Vaos)
                foreach (Vinculo x in v.Vinculos)
                    if (x.Poste == vi.Poste && x.K == vi.K) { n++; if (x == vi) pos = n; }
            InfoEstrutura e = regras.Vinculo.InfoEstrutura(vi.Estrutura);
            int cap = e.Qtd ?? n;
            return cap > 1 ? "slot " + pos + "/" + cap : "";
        }
    }

    public static class Conferencia
    {
        public static void Avaliar(Obra obra, RegrasProjeto regras)
        {
            foreach (Elemento el in obra.Elementos)
            {
                el.Estado = Estados.Ok;
                el.Mensagens.Clear();
                foreach (string m in el.MensagensFixas) el.Marcar(Estados.Conferir, m);
            }
            foreach (Poste p in obra.Postes)
            {
                if (string.IsNullOrEmpty(p.Handle)) p.Marcar(Estados.Conferir, "anotação sem bloco de poste");
                if (!p.TemAnotacao) { p.Marcar(Estados.Conferir, "poste sem anotação"); continue; }
                if (!RegrasVinculo.ParticipaVinculo(p.Operacao)) { p.Mensagens.Add("R não vincula"); continue; }
                for (int k = 0; k < p.Lido.Estruturas.Count && k < p.CabosPorEstrutura.Count; k++)
                {
                    InfoEstrutura e = regras.Vinculo.InfoEstrutura(p.Lido.Estruturas[k].Nome);
                    if (!e.Regra.HasValue) continue;
                    int tem = p.CabosPorEstrutura[k];
                    int precisa = e.QtdMin ?? 0;
                    if (tem < precisa) p.Marcar(Estados.Conferir, e.Codigo + " " + tem + "/" + precisa + " cabos");
                }
                foreach (string a in p.AvisosParametros) p.Marcar(Estados.Conferir, a);
            }
            foreach (Vao v in obra.Vaos)
            {
                if (v.PosteA == null || v.PosteB == null) v.Marcar(Estados.Erro, "sem poste na ponta");
                if (!v.TemAnotacao) { v.Marcar(Estados.Conferir, "vão sem anotação"); continue; }
                if (v.Cabo != null && v.Cabo.Erro != null) v.Marcar(Estados.Conferir, v.Cabo.Erro);
                if (!v.ParticipaVinculo) { v.Mensagens.Add("R não vincula"); continue; }
                if (v.InfoCabo == null) continue;
                if (!v.InfoCabo.Regra.HasValue) { v.Marcar(Estados.Conferir, "tipo de cabo sem regra de vínculo (" + v.InfoCabo.Codigo + ")"); continue; }
                if (!v.InfoCabo.Vincula) continue;
                int tem = v.Vinculos.Count();
                if (tem < v.InfoCabo.Pontas)
                    v.Marcar(Estados.Conferir, "vínculo " + tem + "/" + v.InfoCabo.Pontas + (v.MotivoVinculo != null ? " — " + v.MotivoVinculo : ""));
            }
            foreach (Avulso a in obra.Avulsos)
                if (a.PertoDe != null && a.PertoDe.Numero != null)
                    a.Marcar(Estados.Conferir, "outra anotação perto de " + a.PertoDe.Numero + ", que já tem a sua");
        }

        public static int Contar(Obra obra, string estado)
        {
            return obra.Elementos.Count(e => estado == null || e.Estado == estado);
        }
    }
}
