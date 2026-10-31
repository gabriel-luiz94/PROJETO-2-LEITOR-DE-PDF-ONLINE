// Nucleo/RegrasProjeto.cs — as regras de um projeto (leitor: processamento + classificação; vínculo) e as opções de geometria.
// As regras são DADOS do projeto principal (data/*.json), embutidos na DLL na compilação — uma só fonte de regras.
// Também podem vir de uma pasta "regras\<codigo>\" ao lado da DLL (exportadas da aplicação), que tem prioridade.
using System;
using System.Collections.Generic;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    public class RegrasProjeto
    {
        public string Codigo = "";
        public string Nome = "";
        public string Origem = "";              // "embutida" ou a pasta de onde veio
        public List<JObj> Processamento = new List<JObj>();
        public List<JObj> Classificacao = new List<JObj>();
        public RegrasVinculo Vinculo = new RegrasVinculo();

        // geometria (mesmos números de services/dxf_geometria.py)
        public double RaioAlturas = 4;          // texto → rede/poste: até 4 × a altura do texto
        public double GuiaMaxAlturas = 40;      // linha fora da rede vale como guia até 40 × a altura típica
        public double TolPosteAlturas = 1.5;    // poste "na" rede: até 1,5 × a altura típica da linha
        public string FiltroBlocoPoste = "poste";

        public override string ToString() { return Nome; }

        public static List<JObj> ListaDeRegras(string json)
        {
            if (string.IsNullOrEmpty(json)) return new List<JObj>();
            object v = Json.Ler(json);
            List<object> l = v as List<object>;
            if (l == null && v is JObj) l = ((JObj)v).Lista("regras");
            return l == null ? new List<JObj>() : l.OfType<JObj>().ToList();
        }

        public static RegrasProjeto Montar(string codigo, string nome, string origem, string procJson, string clsJson, string vincJson)
        {
            RegrasProjeto r = new RegrasProjeto();
            r.Codigo = codigo;
            r.Nome = nome;
            r.Origem = origem;
            r.Processamento = ListaDeRegras(procJson);
            r.Classificacao = ListaDeRegras(clsJson);
            r.Vinculo = RegrasVinculo.De(string.IsNullOrEmpty(vincJson) ? null : Json.Ler(vincJson) as JObj);
            return r;
        }

        /// <summary>Layer que representa a operação (tirada das próprias regras: "layer_em: [01_REMOVER] → operacao R").</summary>
        public string LayerDaOperacao(string operacao)
        {
            string op = (operacao ?? "").Trim().ToUpperInvariant();
            foreach (JObj r in Processamento.Where(x => (int)(x.Num("fase") ?? 3) == 3).OrderBy(x => x.Num("ordem") ?? 0))
            {
                List<string> layers = Json.ListaDeTextos(r["layer_em"]);
                if (layers.Count == 1 && (r.Str("operacao") ?? "").ToUpperInvariant() == op) return layers[0];
            }
            return null;
        }

        /// <summary>Operações que o pop-up oferece (as das regras + as padrão).</summary>
        public List<string> Operacoes()
        {
            List<string> ops = new List<string> { "I", "R", "M", "*I", "*R", "*M" };
            return ops;
        }
    }

    /// <summary>Lista de regras disponíveis. `lerRecurso(nome)` devolve o texto de um recurso embutido (ou null).</summary>
    public static class CatalogoRegras
    {
        public class Fonte
        {
            public string Codigo, Nome, Processamento, Classificacao, Vinculacao;
            public Fonte(string codigo, string nome, string proc, string cls, string vinc)
            {
                Codigo = codigo; Nome = nome; Processamento = proc; Classificacao = cls; Vinculacao = vinc;
            }
        }

        // recursos embutidos (LogicalName no .csproj / build_nanocad5.bat) — os arquivos são os de ../data do projeto principal
        public static readonly Fonte[] Embutidas = {
            new Fonte("081", "Pernambuco (081) — Leitor", "regras.081.processamento.json", "regras.081.classificacao.json", "regras.081.vinculacao.json"),
            new Fonte("", "Padrão (genérico)", "regras.padrao.processamento.json", "regras.padrao.classificacao.json", null),
        };

        public static List<RegrasProjeto> Carregar(Func<string, string> lerRecurso, IEnumerable<RegrasProjeto> externas)
        {
            List<RegrasProjeto> lista = new List<RegrasProjeto>();
            if (externas != null) lista.AddRange(externas);
            foreach (Fonte f in Embutidas)
            {
                if (lista.Any(r => r.Codigo == f.Codigo && f.Codigo != "")) continue;
                string proc = lerRecurso(f.Processamento);
                if (proc == null) continue;
                lista.Add(RegrasProjeto.Montar(f.Codigo, f.Nome, "embutida", proc, lerRecurso(f.Classificacao), f.Vinculacao == null ? null : lerRecurso(f.Vinculacao)));
            }
            return lista;
        }
    }
}
