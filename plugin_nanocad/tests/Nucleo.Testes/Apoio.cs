using System;
using System.IO;
using OrcamentoNanoCAD.Nucleo;

namespace OrcamentoNanoCAD.Testes
{
    /// <summary>Caminhos e regras reais do projeto principal (data/*.json) — as mesmas que a DLL embute.</summary>
    public static class Apoio
    {
        public static string RaizRepo
        {
            get
            {
                string d = AppContext.BaseDirectory;
                while (d != null && !File.Exists(Path.Combine(d, "data", "regras_leitor_processamento_pernambuco_seed.json")))
                    d = Path.GetDirectoryName(d);
                if (d == null) throw new InvalidOperationException("raiz do repositório não encontrada");
                return d;
            }
        }

        public static string Plugin { get { return Path.Combine(RaizRepo, "plugin_nanocad"); } }

        public static string Dados(string arquivo) { return File.ReadAllText(Path.Combine(RaizRepo, "data", arquivo)); }

        private static RegrasProjeto _pe, _padrao;

        public static RegrasProjeto Pernambuco
        {
            get
            {
                return _pe ?? (_pe = RegrasProjeto.Montar("081", "Pernambuco (081) — Leitor", "teste",
                    Dados("regras_leitor_processamento_pernambuco_seed.json"), Dados("regras_leitor_classificacao_pernambuco_seed.json"),
                    Dados("vinculacao_pernambuco_seed.json")));
            }
        }

        public static RegrasProjeto Padrao
        {
            get
            {
                return _padrao ?? (_padrao = RegrasProjeto.Montar("", "Padrão", "teste",
                    Dados("regras_leitor_processamento_seed.json"), Dados("regras_leitor_classificacao_seed.json"), null));
            }
        }
    }
}
