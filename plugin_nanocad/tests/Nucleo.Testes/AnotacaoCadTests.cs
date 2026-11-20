using System.Collections.Generic;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class AnotacaoCadTests
    {
        [Theory]
        [InlineData("POSTE_INSTALANDO", "I")]
        [InlineData("POSTE_INSTALANDO_LV", "*I")]
        [InlineData("POSTE_REMOVENDO", "R")]
        [InlineData("POSTE_REMOVENDO_LV", "*R")]
        [InlineData("POSTE_REAPLICANDO", "M")]
        [InlineData("POSTE_REAPLICANDO_LV", "*M")]
        public void Layer_de_anotacao_tem_a_mesma_operacao_no_leitor(string layer, string operacao)
        {
            RegrasProjeto regras = Apoio.Pernambuco;
            var texto = new ItemLeitor { Layer = layer, Cor = "#000000", Texto = "9200D, N1" };
            var resultado = MotorLeitor.ProcessarEClassificar(new[] { texto }, 0, regras.Processamento, regras.Classificacao);
            Assert.Equal(operacao, resultado.Operacao);
            Assert.Equal(operacao, regras.OperacoesPorLayer()[layer]);
            Assert.Equal(operacao, AnotacaoCad.AlternarLinhaViva(AnotacaoCad.AlternarLinhaViva(operacao)));
        }

        [Fact]
        public void Guia_vinculada_e_exclusiva_do_mtext_de_origem()
        {
            var bruto = new DesenhoBruto();
            bruto.Entidades.Add(new EntidadeBruta { Tipo = "MTEXT", Handle = "A", Texto = "N1", Altura = 1 });
            bruto.Entidades.Add(new EntidadeBruta { Tipo = "MTEXT", Handle = "B", Texto = "N3", Altura = 1 });
            var guia = new EntidadeBruta { Tipo = "LINE", Handle = "C", NotaDaGuia = "B", Pontos = new List<Ponto> { new Ponto(0,0), new Ponto(20,0) } };
            bruto.Entidades.Add(EntidadeBruta.De(guia.ParaJson()));
            Desenho desenho = MontadorDesenho.Montar(bruto, new ConfigLeitura());
            Assert.Empty(desenho.Linhas);
            Assert.Empty(desenho.Textos.Single(t => t.Handle == "A").Guias);
            Assert.Single(desenho.Textos.Single(t => t.Handle == "B").Guias);
        }

        [Fact]
        public void Saida_da_guia_termina_na_borda_da_nota()
        {
            Ponto borda = AnotacaoCad.CentroDaBorda(new Ponto(0,0), new Ponto(12,4), new Ponto(6,20));
            Assert.Equal(6, borda.X);
            Assert.Equal(4, borda.Y);
        }
    }
}
