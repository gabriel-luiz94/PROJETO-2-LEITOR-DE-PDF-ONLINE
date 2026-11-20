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
        public string Versao;                   // impressão digital das regras na aplicação (só quando vêm dela)
        public bool DaAplicacao;                // baixadas da aplicação (atualizam sozinhas quando ela muda)

        // onde estão rede, postes e guias no desenho (padrão = Pernambuco, como services/dxf_geometria.py); editável por projeto
        public ConfigLeitura Leitura = new ConfigLeitura();
        public bool LeituraDaAplicacao;         // TASK-015: a configuração de leitura veio da aplicação (vale em qualquer computador)

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

        public static RegrasProjeto MontarDeObjetos(string codigo, string nome, string origem, List<JObj> proc, List<JObj> cls, JObj vinc)
        {
            RegrasProjeto r = new RegrasProjeto();
            r.Codigo = codigo;
            r.Nome = nome;
            r.Origem = origem;
            r.Processamento = proc ?? new List<JObj>();
            r.Classificacao = cls ?? new List<JObj>();
            r.Vinculo = RegrasVinculo.De(vinc);
            return r;
        }

        /// <summary>Layer que representa a operação (tirada das próprias regras: "layer_em: [01_REMOVER] → operacao R").</summary>
        public string LayerDaOperacao(string operacao)
        {
            string op = (operacao ?? "").Trim().ToUpperInvariant();
            return OperacoesPorLayer().Where(x => x.Value == op).Select(x => x.Key).FirstOrDefault();
        }

        public Dictionary<string,string> OperacoesPorLayer()
        {
            var mapa = new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
            List<string> operacoes = Operacoes();
            foreach (JObj r in Processamento.OrderBy(x => x.Num("ordem") ?? 0))
            {
                string op = r.Str("operacao");
                if (!r.Bool("ativa",true) || (r.Num("fase") ?? 3) != 3 || !operacoes.Contains(op) ||
                    !string.IsNullOrEmpty(r.Str("texto_regex")) || !string.IsNullOrEmpty(r.Str("ativo_regex")) || r.Obj("vizinhanca") != null ||
                    new[] { "cor_em","cor_nao_em","cor_hex_em","cor_hex_nao_em" }.Any(k => Json.ListaDeTextos(r[k]).Count > 0)) continue;
                foreach (string layer in Json.ListaDeTextos(r["layer_em"]))
                    if (!string.IsNullOrEmpty(layer) && layer.IndexOfAny(new[] { '*','?' }) < 0 && !mapa.ContainsKey(layer)) mapa.Add(layer,op);
            }
            return mapa;
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
        public const string Paraiba = "027";

        public static readonly Fonte[] Embutidas = {
            new Fonte("081", "Pernambuco (081) — Leitor", "regras.081.processamento.json", "regras.081.classificacao.json", "regras.081.vinculacao.json"),
            new Fonte(Paraiba, "Paraíba (027) — Leitor de fábrica", "regras.padrao.processamento.json", "regras.padrao.classificacao.json", null),
            new Fonte("", "Padrão (genérico)", "regras.padrao.processamento.json", "regras.padrao.classificacao.json", null),
        };

        /// <summary>Resolve somente pacotes do projeto pedido. A base embutida de 027 é a usada no seed do OrçaCore.</summary>
        public static RegrasProjeto Resolver(string codigo, RegrasProjeto escolhidas, Func<string, RegrasProjeto> baixar,
            Func<IEnumerable<RegrasProjeto>> disponiveis)
        {
            bool mesmaEscolha = escolhidas != null && escolhidas.Codigo == codigo;
            if (mesmaEscolha && !escolhidas.DaAplicacao && (codigo != Paraiba || escolhidas.Origem != "embutida"))
                return escolhidas;

            if (baixar != null && !string.IsNullOrEmpty(codigo))
            {
                RegrasProjeto remotas = baixar(codigo);
                if (remotas != null && remotas.Codigo == codigo) return remotas;
            }

            if (mesmaEscolha && escolhidas.DaAplicacao && TemRegrasBaixadas(escolhidas)) return escolhidas;
            return disponiveis().FirstOrDefault(r => r.Codigo == codigo && (!r.DaAplicacao || TemRegrasBaixadas(r)));
        }

        private static bool TemRegrasBaixadas(RegrasProjeto regras)
        {
            return !string.IsNullOrEmpty(regras.Versao) || regras.Processamento.Count > 0 || regras.Classificacao.Count > 0;
        }

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
