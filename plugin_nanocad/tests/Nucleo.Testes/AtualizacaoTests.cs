// Atualização automática do plugin (TASK-013): consulta, download conferido e troca da DLL.
using System;
using System.IO;
using System.Text;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class BaixadorFalso : IBaixador
    {
        public byte[] Dados;
        public string UltimaUrl;
        public byte[] Baixar(string url, int timeoutMs, out int status, out string erro)
        {
            UltimaUrl = url; erro = null; status = Dados == null ? 404 : 200;
            return Dados;
        }
    }

    public class AtualizacaoTests
    {
        private static readonly byte[] Dll = Encoding.ASCII.GetBytes("MZ" + new string('x', 100));

        private static TransporteFalso Servidor(string versao, byte[] dados, string url = "/api/plugin/atualizacao/arquivo?alvo=nanocad26")
        {
            TransporteFalso t = new TransporteFalso();
            t.Rotas["GET /api/plugin/atualizacao"] = (c, tok) => TransporteFalso.Ok(Json.Escrever(new JObj().Set("alvo", "nanocad26").Set("versao", versao)
                .Set("sha256", Atualizacao.Sha256(dados)).Set("tamanho", dados.Length).Set("url", url), false));
            return t;
        }

        [Fact]
        public void Compara_versoes_numero_a_numero()
        {
            Assert.True(Atualizacao.Comparar("0.10.0", "0.9.9") > 0);
            Assert.True(Atualizacao.Comparar("0.6.0", "0.6") == 0);
            Assert.True(Atualizacao.Comparar("v1.0.0-beta", "0.99.0") > 0);
            Assert.True(Atualizacao.Comparar("0.5.0", "0.6.0") < 0);
        }

        [Fact]
        public void So_baixa_de_https_ou_do_proprio_computador()
        {
            Assert.True(Atualizacao.OrigemPermitida("https://app.exemplo.com"));
            Assert.True(Atualizacao.OrigemPermitida("http://127.0.0.1:8000"));
            Assert.True(Atualizacao.OrigemPermitida("http://localhost:8000"));
            Assert.False(Atualizacao.OrigemPermitida("http://app.exemplo.com"));
            Assert.False(Atualizacao.OrigemPermitida("ftp://x"));
            string motivo;
            Assert.Null(Atualizacao.Consultar(Servidor("9.0.0", Dll), "http://192.168.0.5:8000", "nanocad26", "0.6.0", null, out motivo));
            Assert.Contains("segurança", motivo);
        }

        [Fact]
        public void Versao_mais_nova_e_baixada_e_conferida_pelo_sha256()
        {
            string motivo;
            PacoteAtualizacao p = Atualizacao.Consultar(Servidor("0.7.0", Dll), "http://127.0.0.1:8000", "nanocad26", "0.6.0", null, out motivo);
            Assert.NotNull(p);
            BaixadorFalso b = new BaixadorFalso { Dados = Dll };
            Assert.Equal(Dll, Atualizacao.Baixar(b, "http://127.0.0.1:8000", p, out motivo));
            Assert.Equal("http://127.0.0.1:8000/api/plugin/atualizacao/arquivo?alvo=nanocad26", b.UltimaUrl);

            byte[] adulterado = (byte[])Dll.Clone();
            adulterado[10] = (byte)'y';
            Assert.Null(Atualizacao.Baixar(new BaixadorFalso { Dados = adulterado }, "http://127.0.0.1:8000", p, out motivo));
            Assert.Contains("SHA-256", motivo);
        }

        [Fact]
        public void Mesma_versao_versao_desfeita_e_arquivo_fora_da_aplicacao_nao_atualizam()
        {
            string motivo;
            Assert.Null(Atualizacao.Consultar(Servidor("0.6.0", Dll), "http://127.0.0.1:8000", "nanocad26", "0.6.0", null, out motivo));
            Assert.Contains("mais nova", motivo);
            Assert.Null(Atualizacao.Consultar(Servidor("0.7.0", Dll), "http://127.0.0.1:8000", "nanocad26", "0.6.0", "0.7.0", out motivo));
            Assert.Contains("desfeita", motivo);
            Assert.NotNull(Atualizacao.Consultar(Servidor("0.8.0", Dll), "http://127.0.0.1:8000", "nanocad26", "0.6.0", "0.7.0", out motivo));
            Assert.Null(Atualizacao.Consultar(new TransporteFalso(), "http://127.0.0.1:8000", "nanocad5", "0.6.0", null, out motivo));
            Assert.Contains("não publica", motivo);

            PacoteAtualizacao fora = Atualizacao.Consultar(Servidor("0.7.0", Dll, "https://outro.site/x.dll"), "http://127.0.0.1:8000", "nanocad26", "0.6.0", null, out motivo);
            Assert.Null(Atualizacao.Baixar(new BaixadorFalso { Dados = Dll }, "http://127.0.0.1:8000", fora, out motivo));
            Assert.Contains("fora da aplicação", motivo);
        }

        [Fact]
        public void Instalar_troca_a_dll_guarda_a_anterior_e_voltar_desfaz()
        {
            string pasta = Path.Combine(Path.GetTempPath(), "orc_atualizacao_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(pasta);
            try
            {
                string dll = Path.Combine(pasta, Atualizacao.ArquivoDll);
                File.WriteAllText(dll, "velha");
                string motivo;
                Assert.Equal(ResultadoTroca.Trocada, Atualizacao.Instalar(dll, Encoding.ASCII.GetBytes("nova"), out motivo));
                Assert.Equal("nova", File.ReadAllText(dll));
                Assert.Equal("velha", File.ReadAllText(dll + ".anterior"));
                Assert.False(File.Exists(dll + ".nova"));

                Assert.True(Atualizacao.Reverter(dll, out motivo));
                Assert.Equal("velha", File.ReadAllText(dll));
                Assert.Equal("nova", File.ReadAllText(dll + ".desfeita"));
                Atualizacao.Limpar(dll);
                Assert.False(File.Exists(dll + ".desfeita"));
                Assert.False(Atualizacao.Reverter(dll, out motivo));
                Assert.Contains("não há versão anterior", motivo);
            }
            finally { Directory.Delete(pasta, true); }
        }
    }
}
