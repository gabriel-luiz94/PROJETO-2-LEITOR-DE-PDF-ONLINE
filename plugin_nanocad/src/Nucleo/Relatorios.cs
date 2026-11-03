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
        public string[] Icones;       // ícone desenhado em cada coluna (null = só texto): "estado:OK", "rede:MT", "op:I", "poste", "avulso"
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

        // ── painel em duas listas, como o Resumo da aplicação (Tabela de Cabos · Tabela de Outros) ──
        public const string PainelCabos = "CABOS", PainelOutros = "OUTROS";

        /// <summary>O que aparece no painel/na tela. Ler: só o que as regras do leitor classificaram (elemento com anotação lida) —
        /// rede existente e poste sem anotação ficam ocultos. Construir (e o que foi montado à mão): tudo, sem filtro das regras.</summary>
        public static bool Visivel(Obra obra, Elemento e)
        {
            return obra.Modo == "construir" || e.Manual || e.TemAnotacao;
        }

        public static string PainelDe(Elemento e)
        {
            if (e is Vao) return PainelCabos;
            if (e is Poste) return PainelOutros;
            return e.Entidade == "CABO" ? PainelCabos : PainelOutros;
        }

        public static IEnumerable<Elemento> DoPainel(Obra obra, string painel)
        {
            return obra.Elementos.Where(e => Visivel(obra, e) && (painel == null || PainelDe(e) == painel));
        }

        public static int Contar(Obra obra, string painel, string estado)
        {
            return DoPainel(obra, painel).Count(e => estado == null || e.Estado == estado);
        }

        /// <summary>Elementos que a leitura achou mas o painel não mostra (sem anotação classificada pelo leitor).</summary>
        public static int Ocultos(Obra obra) { return obra.Elementos.Count(e => !Visivel(obra, e)); }

        public static readonly string[] ColunasCabos = { "", "Nº", "Rede", "Op", "Anotação", "m", "Postes", "Observação" };
        public static readonly string[] ColunasOutros = { "", "Nº", "Tipo", "Op", "Anotação", "Placa", "Observação" };

        private static string OpCurta(string op) { return (op ?? "").Trim(); }

        public static List<LinhaTabela> Cabos(Obra obra, string filtro)
        {
            List<LinhaTabela> r = new List<LinhaTabela>();
            foreach (Elemento e in DoPainel(obra, PainelCabos))
            {
                if (filtro != null && e.Estado != filtro) continue;
                Vao v = e as Vao;
                string rede = v != null ? (v.Tipo ?? "") : "";
                string anot = v != null ? v.Ativo : e.AtivoLido;
                double? m = v != null ? (v.Metros ?? v.MetrosLidos) : null;
                string metros = m.HasValue ? m.Value.ToString("0.#", System.Globalization.CultureInfo.InvariantCulture) : "";
                string postes = v != null ? (v.PosteA ?? "—") + " → " + (v.PosteB ?? "—") : "solto";
                r.Add(new LinhaTabela
                {
                    Chave = e.Numero, Estado = e.Estado,
                    Celulas = new[] { "", e.Numero, rede, OpCurta(e.Operacao), string.IsNullOrEmpty(anot) ? "(sem anotação)" : anot, metros, postes, Observacao(e) },
                    Icones = new[] { "estado:" + e.Estado, null, v != null ? "rede:" + rede : "avulso", "op:" + OpCurta(e.Operacao), null, null, null, null },
                });
            }
            return r;
        }

        public static List<LinhaTabela> Outros(Obra obra, string filtro)
        {
            List<LinhaTabela> r = new List<LinhaTabela>();
            foreach (Elemento e in DoPainel(obra, PainelOutros))
            {
                if (filtro != null && e.Estado != filtro) continue;
                Poste p = e as Poste;
                string tipo = p != null ? (p.Lido.CodigoPoste != null ? "Poste" : p.Lido.Estruturas.Count > 0 ? "Estrutura" : "Poste") : "Solto";
                r.Add(new LinhaTabela
                {
                    Chave = e.Numero, Estado = e.Estado,
                    Celulas = new[] { "", e.Numero, tipo, OpCurta(e.Operacao), string.IsNullOrEmpty(e.AtivoLido) ? "(sem anotação)" : e.AtivoLido,
                                      p != null ? (p.Placa ?? "") : "", Observacao(e) },
                    Icones = new[] { "estado:" + e.Estado, null, p != null ? "poste" : "avulso", "op:" + OpCurta(e.Operacao), null, null, null },
                });
            }
            return r;
        }

        private static string Observacao(Elemento e)
        {
            List<string> l = new List<string>(e.Mensagens);
            if (e.IrmaoDe != null) l.Insert(0, "junto com " + e.IrmaoDe);
            if (e.Manual) l.Insert(0, "✋ manual");
            return string.Join("; ", l.ToArray());
        }

        public static List<LinhaTabela> Conferencia(Obra obra, string filtro)
        {
            List<LinhaTabela> r = new List<LinhaTabela>();
            foreach (Elemento e in obra.Elementos)
            {
                if (filtro != null && e.Estado != filtro) continue;
                LinhaTabela l = new LinhaTabela();
                l.Chave = e.Numero;
                Vao v = e as Vao;
                Poste p = e as Poste;
                string anot = v != null ? v.Ativo : e.AtivoLido;
                if (p != null && !string.IsNullOrEmpty(p.Placa)) anot = (string.IsNullOrEmpty(anot) ? "" : anot + " · ") + "placa " + p.Placa;
                l.Celulas = new[] { e.Numero, string.IsNullOrEmpty(anot) ? "(sem anotação)" : anot, e.Operacao ?? "", TextoEstado(e) };
                l.Estado = e.Estado;
                r.Add(l);
            }
            return r;
        }

        public static readonly string[] ColunasIgnorados = { "Texto", "Layer", "Cor", "Motivo" };

        /// <summary>Textos que as regras descartaram (chave "H:&lt;handle&gt;" para localizar no desenho).</summary>
        public static List<LinhaTabela> TextosIgnorados(Obra obra)
        {
            return obra.Ignorados.Select(t => new LinhaTabela { Chave = "H:" + t.Handle, Celulas = new[] { t.Texto, t.Layer, t.Cor, t.Motivo }, Estado = Estados.Ok }).ToList();
        }

        public static readonly string[] ColunasDiagnostico = { "Texto", "Tipo", "Layer", "Cor", "Entidade", "Op", "Ativo", "Destino" };

        /// <summary>A lista de textos do arquivo de diagnóstico, como tabela.</summary>
        public static List<LinhaTabela> TextosDoDiagnostico(JObj diag)
        {
            List<LinhaTabela> r = new List<LinhaTabela>();
            foreach (JObj t in (diag.Lista("textos") ?? new List<object>()).OfType<JObj>())
            {
                string destino = t.Str("destino") ?? "";
                r.Add(new LinhaTabela
                {
                    Chave = "H:" + t.Str("h"),
                    Celulas = new[] { t.Str("texto"), t.Str("tipo"), t.Str("layer"), t.Str("cor"), t.Str("entidade"), t.Str("operacao"), t.Str("ativo"), destino },
                    Estado = destino.StartsWith("ignorado") || destino.StartsWith("texto da defini") ? Estados.Ok : destino.Contains(" — ") || destino.StartsWith("A") ? Estados.Conferir : Estados.Ok,
                });
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
                if (!Visivel(obra, v)) continue;
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
                string obs = v.Cabo != null && v.Cabo.Erro != null ? v.Cabo.Erro : "—";
                if (v.ForaDaBase.Count > 0) obs = (obs == "—" ? "" : obs + " · ") + ForaDaBase(v.ForaDaBase);
                l.Celulas = new[] { v.Numero, rotulo, obs };
                l.Estado = obs != "—" ? Estados.Conferir : Estados.Ok;
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
                    string ativo = AtivoDaEstrutura(p, k, out pars, out pendente);
                    bool fora = p.ForaDaBase.Contains(p.Lido.Estruturas[k].Nome);
                    if (fora) pars += " · não está na base";
                    l.Celulas = new[] { p.Numero + " · " + p.Lido.Estruturas[k].Nome, ativo, pars };
                    l.Estado = pendente || fora ? Estados.Conferir : Estados.Ok;
                    r.Add(l);
                }
            }
            foreach (Poste p in obra.Postes)
            {
                if (p.Lido.CodigoPoste == null) continue;
                LinhaTabela l = new LinhaTabela();
                l.Chave = p.Numero;
                string oficial = Anotacoes.PosteOficial(p.Lido.CodigoPoste);
                bool fora = oficial != null && p.ForaDaBase.Contains(oficial);
                l.Celulas = new[] { p.Numero + " · poste", oficial ?? p.Lido.CodigoPoste, oficial == null ? "poste fora da notação compacta" : fora ? "não está na base" : "—" };
                l.Estado = oficial == null || fora ? Estados.Conferir : Estados.Ok;
                r.Add(l);
            }
            return r;
        }

        private static string ForaDaBase(List<string> l) { return "não está na base: " + string.Join(", ", l.ToArray()); }

        /// <summary>Texto do comando no rodapé: "15 elementos · 11 ok · 3 a conferir · 1 erro · regras: Pernambuco (081)".</summary>
        public static string Resumo(Obra obra)
        {
            int total = Contar(obra, null, null), ok = Contar(obra, null, Estados.Ok), conf = Contar(obra, null, Estados.Conferir), erro = Contar(obra, null, Estados.Erro);
            int cabos = Contar(obra, PainelCabos, null), ocultos = Ocultos(obra);
            return total + " elementos (" + cabos + " cabos · " + (total - cabos) + " postes/outros) · " + ok + " ok · " + conf + " a conferir · " + erro + " erro"
                   + (erro == 1 ? "" : "s") + (ocultos > 0 ? " · " + ocultos + " sem anotação ocultos" : "") + " · regras: " + obra.NomeRegras;
        }
    }
}
