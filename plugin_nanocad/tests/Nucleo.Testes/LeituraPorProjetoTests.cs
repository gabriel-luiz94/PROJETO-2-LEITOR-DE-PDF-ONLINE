using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class LeituraPorProjetoTests
    {
        [Theory]
        [InlineData("COR_MANDA", "media_1", "#0000ff", "MT")]
        [InlineData("COR_MANDA", "media_1", "#ffff00", null)]
        [InlineData("COR_MANDA", "outra", "#0000ff", null)]
        [InlineData("LAYER_MANDA", "outra", "#0000ff", "MT")]
        [InlineData("LAYER_MANDA", "media_12", "#ffff00", "BT")]
        [InlineData("LAYER_MANDA", "media_1", "#ffff00", "MT")]
        [InlineData("SO_LAYER", "outra", "#ffff00", null)]
        [InlineData("SO_LAYER", "media_1", "#ffff00", "MT")]
        public void Usa_o_perfil_mesmo_depois_de_reabrir(string regra, string layer, string cor, string esperado)
        {
            var dados = (JObj)Json.Ler(@"{""layers_rede_mt"":[""MEDIA_?""],""cor_mt"":[5,""#123456""],""cor_bt"":{""r"":[255,255],""g"":[254,255],""b"":[0,0]}}");
            var leitura = ConfigLeitura.De(dados.Set("regra", regra)).Copia();
            Assert.Equal(esperado, leitura.ClassificarRede(layer, cor));
        }

        [Fact]
        public void Perfil_antigo_sem_confirmacao_usa_so_layer()
        {
            var leitura = ConfigLeitura.De(new JObj().Set("cor_confirma", false));
            Assert.Equal("MT", leitura.ClassificarRede("REDE_MT", "#00ff00"));
        }
    }
}
