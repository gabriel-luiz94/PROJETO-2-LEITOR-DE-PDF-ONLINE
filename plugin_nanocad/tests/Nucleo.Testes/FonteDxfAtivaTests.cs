using System;
using System.IO;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class FonteDxfAtivaTests
    {
        [Theory]
        [InlineData(true)]
        [InlineData(false)]
        public void Dxf_aberto_reutiliza_original_apenas_sem_alteracoes(bool arquivoAtual)
        {
            string pasta = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(pasta);
            string original = Path.Combine(pasta, "Desenho.DXF");
            File.WriteAllText(original, "original");
            string exportado = null;
            try
            {
                string lido = FonteDxfAtiva.Usar(original, () => arquivoAtual, destino =>
                {
                    exportado = destino;
                    File.WriteAllText(destino, "alterado");
                }, (caminho, temporario) =>
                {
                    Assert.Equal(!arquivoAtual, temporario);
                    Assert.Equal(arquivoAtual ? original : exportado, caminho);
                    return File.ReadAllText(caminho);
                }, erro => throw erro);

                Assert.Equal(arquivoAtual ? "original" : "alterado", lido);
                Assert.Equal("original", File.ReadAllText(original));
                if (arquivoAtual) Assert.Null(exportado);
                else Assert.False(Directory.Exists(Path.GetDirectoryName(exportado)));
            }
            finally { Directory.Delete(pasta, true); }
        }

        [Fact]
        public void Dwg_sem_salvar_exporta_em_cada_leitura_e_limpa_apos_falha()
        {
            string primeiro = null;
            string segundo = null;
            Assert.Throws<InvalidOperationException>(() => FonteDxfAtiva.Usar<string>("Drawing1.dwg",
                () => throw new Exception("DWG não depende de DBMOD."),
                destino =>
                {
                    primeiro = destino;
                    File.WriteAllText(destino, "primeira revisão");
                },
                (caminho, temporario) =>
                {
                    Assert.True(temporario);
                    Assert.Equal("primeira revisão", File.ReadAllText(caminho));
                    throw new InvalidOperationException("Falha na leitura.");
                },
                erro => throw erro));
            Assert.NotNull(primeiro);
            Assert.False(Directory.Exists(Path.GetDirectoryName(primeiro)));

            string lido = FonteDxfAtiva.Usar("Drawing1.dwg", () => true, destino =>
            {
                segundo = destino;
                File.WriteAllText(destino, "segunda revisão");
            }, (caminho, temporario) => File.ReadAllText(caminho), erro => throw erro);
            Assert.Equal("segunda revisão", lido);
            Assert.NotEqual(primeiro, segundo);
            Assert.False(Directory.Exists(Path.GetDirectoryName(segundo)));
        }

        [Fact]
        public void Exportacao_vazia_nao_envia_arquivo_e_limpa_a_pasta()
        {
            string exportado = null;
            Assert.Throws<IOException>(() => FonteDxfAtiva.Usar<bool>("Drawing1.dwg", () => false, destino =>
            {
                exportado = destino;
                File.WriteAllText(destino, "");
            }, (caminho, temporario) => throw new InvalidOperationException("Envio indevido."), erro => throw erro));
            Assert.NotNull(exportado);
            Assert.False(Directory.Exists(Path.GetDirectoryName(exportado)));
        }
    }
}
