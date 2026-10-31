// Nucleo/Relatorios.cs — as tabelas que o plugin mostra: conferência, "Ver vínculos" e "Ver ativos com parâmetros".
// Ficam no núcleo (e não na tela) para serem testadas sem o CAD.
using System;
using System.Collections.Generic;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    public class LinhaTabela
    {
        public string Chave;          // número do elemento (P1, V2…) para zoom/pop-up
        public string[] Celulas;
        public string Estado = Estados.Ok;
    }

    public static class Relatorios
    {
        public static string Resumir(string texto, int max)
        {
            string t = texto ?? "";
            return t.Length <= max ? t : t.Substring(0, max - 1) + "…";
        }

        public static string TextoEstado(Elemento e)
        {
            string marca = e.Estado == Estados.Erro ? "✖ " : e.Estado == Estados.Conferir ? "⚠ " : "✔ ";
            return marca + (e.Mensagens.Count > 0 ? string.Join("; ", e.Mensagens.ToArray()) : "");
        }

        public static readonly string[] ColunasConferencia = { "Nº", "Anotação / texto", "Op", "Estado" };

        public static List<LinhaTabela> Conferencia(Obra obra, string filtro)
        {
            List<LinhaTabela> r = new List<LinhaTabela>();
            foreach (Elemento e in obra.Elementos)
            {
                if (filtro != null && e.Estado != filtro) continue;
                LinhaTabela l = new LinhaTabela();
                l.Chave = e.Numero;
                Vao v = e as Vao;
                string anot = v != null ? v.Ativo : e.AtivoLido;
                l.Celulas = new[] { e.Numero, string.IsNullOrEmpty(anot) ? "(sem anotação)" : anot, e.Operacao ?? "", TextoEstado(e) };
                l.Estado = e.Estado;
                r.Add(l);
            }
            return r;
        }

        public static readonly string[] ColunasVinculos = { "Vão", "Anotação", "Op", "Poste A", "Estrutura em A", "Poste B", "Estrutura em B", "Estado" };

        private static string PosteTxt(Obra obra, string numero)
        {
            Poste p = obra.PostePorNumero(numero);
            if (p == null) return "—";
            return p.Numero + (p.Lido.CodigoPoste != null ? " · " + p.Lido.CodigoPoste : "");
        }

        private static string EstruturaTxt(Obra obra, Vao v, Vinculo vi, RegrasProjeto regras, bool temPoste)
        {
            if (vi == null) return temPoste ? "— sem vínculo" : "— sem poste";
            string slot = Vinculador.Slot(obra, v, vi, regras);
            return vi.Estrutura + (slot.Length > 0 ? " (" + slot + ")" : "") + (vi.Manual ? " ✋" : "");
        }

        public static List<LinhaTabela> Vinculos(Obra obra, RegrasProjeto regras)
        {
            List<LinhaTabela> r = new List<LinhaTabela>();
            foreach (Vao v in obra.Vaos)
            {
                LinhaTabela l = new LinhaTabela();
                l.Chave = v.Numero;
                l.Estado = v.Estado;
                string anot = v.TemAnotacao ? v.Ativo : "(sem anotação)";
                if (v.TemAnotacao && !v.ParticipaVinculo)
                    l.Celulas = new[] { v.Numero, anot, v.Operacao, "Operação " + v.Operacao + " não entra no vínculo", "", "", "", TextoEstado(v) };
                else
                    l.Celulas = new[] { v.Numero, anot, v.Operacao ?? "", PosteTxt(obra, v.PosteA), EstruturaTxt(obra, v, v.VinculoA, regras, v.PosteA != null),
                                        PosteTxt(obra, v.PosteB), EstruturaTxt(obra, v, v.VinculoB, regras, v.PosteB != null), TextoEstado(v) };
                r.Add(l);
            }
            return r;
        }

        public static readonly string[] ColunasAtivos = { "Origem", "Ativo gerado (pedido)", "Parâmetros" };

        /// <summary>Ativo de uma estrutura com os parâmetros preenchidos: N4(9400D, P, 3XA35P_7.94_-13.8KV, ?) — "?" = pendente.</summary>
        public static string AtivoDaEstrutura(Poste p, int k, out string parametros, out bool pendente)
        {
            List<string> valores = new List<string>(), nomes = new List<string>();
            pendente = false;
            if (p.Lido.CodigoPoste != null) { valores.Add(p.Lido.CodigoPoste); nomes.Add("poste=" + p.Lido.CodigoPoste); }
            List<KeyValuePair<string, string>> ps = k < p.ParametrosEstruturas.Count ? p.ParametrosEstruturas[k] : new List<KeyValuePair<string, string>>();
            foreach (KeyValuePair<string, string> kv in ps)
            {
                string nome = kv.Key.ToLowerInvariant();
                if (kv.Value == null) { valores.Add("?"); nomes.Add(nome + " faltando"); pendente = true; }
                else { valores.Add(kv.Value); nomes.Add(nome + "=" + kv.Value); }
            }
            parametros = nomes.Count > 0 ? string.Join(" · ", nomes.ToArray()) : "—";
            string nomeEstr = p.Lido.Estruturas[k].Nome;
            return valores.Count > 0 ? nomeEstr + "(" + string.Join(", ", valores.ToArray()) + ")" : nomeEstr;
        }

        public static List<LinhaTabela> Ativos(Obra obra)
        {
            List<LinhaTabela> r = new List<LinhaTabela>();
            foreach (Vao v in obra.Vaos)
            {
                if (!v.TemAnotacao) continue;
                LinhaTabela l = new LinhaTabela();
                l.Chave = v.Numero;
                string rotulo = v.Cabo != null && !string.IsNullOrEmpty(v.Cabo.Rotulo) ? v.Cabo.Rotulo : v.Ativo;
                l.Celulas = new[] { v.Numero, rotulo, v.Cabo != null && v.Cabo.Erro != null ? v.Cabo.Erro : "—" };
                l.Estado = v.Cabo != null && v.Cabo.Erro != null ? Estados.Conferir : Estados.Ok;
                r.Add(l);
            }
            foreach (Poste p in obra.Postes)
            {
                for (int k = 0; k < p.Lido.Estruturas.Count; k++)
                {
                    string pars;
                    bool pendente;
                    LinhaTabela l = new LinhaTabela();
                    l.Chave = p.Numero;
                    l.Celulas = new[] { p.Numero + " · " + p.Lido.Estruturas[k].Nome, AtivoDaEstrutura(p, k, out pars, out pendente), pars };
                    l.Estado = pendente ? Estados.Conferir : Estados.Ok;
                    r.Add(l);
                }
            }
            foreach (Poste p in obra.Postes)
            {
                if (p.Lido.CodigoPoste == null) continue;
                LinhaTabela l = new LinhaTabela();
                l.Chave = p.Numero;
                string oficial = Anotacoes.PosteOficial(p.Lido.CodigoPoste);
                l.Celulas = new[] { p.Numero + " · poste", oficial ?? p.Lido.CodigoPoste, oficial == null ? "poste fora da notação compacta" : "—" };
                l.Estado = oficial == null ? Estados.Conferir : Estados.Ok;
                r.Add(l);
            }
            return r;
        }

        /// <summary>Texto do comando no rodapé: "15 elementos · 11 ok · 3 a conferir · 1 erro · regras: Pernambuco (081)".</summary>
        public static string Resumo(Obra obra)
        {
            int total = Conferencia_(obra, null), ok = Conferencia_(obra, Estados.Ok), conf = Conferencia_(obra, Estados.Conferir), erro = Conferencia_(obra, Estados.Erro);
            return total + " elementos · " + ok + " ok · " + conf + " a conferir · " + erro + " erro" + (erro == 1 ? "" : "s") + " · regras: " + obra.NomeRegras;
        }

        private static int Conferencia_(Obra obra, string estado) { return Nucleo.Conferencia.Contar(obra, estado); }
    }
}
