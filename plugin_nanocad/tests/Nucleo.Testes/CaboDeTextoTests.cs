// TASK-021: "+ Cabo" — o texto escolhido é lido COMO CABO e o comprimento vem da mesma linha, dos escolhidos ou da vizinhança no desenho.
using System;
using System.Collections.Generic;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class CaboDeTextoTests
    {
        private const string Rotulo = "3XA35P(7,94)-13,8KV";

        private static TextoDesenho T(string h, string texto, double x, double y, string layer = "0")
        {
            return new TextoDesenho { Handle = h, Texto = texto, Layer = layer, Cor = "#FFFFFF", X = x, Y = y, Altura = 1 };
        }

        private static Func<IEnumerable<TextoDesenho>> Desenho(params TextoDesenho[] t) { return () => t; }

        [Theory]
        [InlineData("0")]
        [InlineData("01_INSTALANDO")]        // layer de operação: a operação vem da regra
        public void Comprimento_embaixo_no_desenho_mesmo_sem_selecionar(string layer)
        {
            TextoDesenho rot = T("R1", Rotulo, 0, 10, layer);
            Construtor.LeituraDeTexto r = Construtor.LerCaboDeTexto(new List<TextoDesenho> { rot }, Apoio.Pernambuco,
                Desenho(rot, T("C1", "35 M", 0.5, 8.5, layer), T("X", "RUA DAS FLORES", 0, 4)));
            Assert.Equal(Rotulo + " 35 m", r.Ativo);
            Assert.True(r.Cabo.Value);
            Assert.True(r.ComprimentoDaVizinhanca);
            Assert.Equal("35 m", r.Comprimento);
        }

        [Fact]
        public void Comprimento_na_mesma_linha_nao_procura_no_desenho()
        {
            Construtor.LeituraDeTexto r = Construtor.LerCaboDeTexto(new List<TextoDesenho> { T("R1", Rotulo + " 40 m", 0, 10, "01_INSTALANDO") }, Apoio.Pernambuco,
                () => { throw new InvalidOperationException("não devia ler o desenho"); });
            Assert.Equal(Rotulo + " 40 m", r.Ativo);
            Assert.False(r.ComprimentoDaVizinhanca);
            Assert.Equal("40 m", r.Comprimento);
        }

        [Fact]
        public void Comprimento_ao_lado_na_mesma_linha_em_outro_texto()
        {
            TextoDesenho rot = T("R1", Rotulo, 0, 10);
            Construtor.LeituraDeTexto r = Construtor.LerCaboDeTexto(new List<TextoDesenho> { rot }, Apoio.Pernambuco, Desenho(rot, T("C1", "22,5m", 6, 10)));
            Assert.Equal(Rotulo + " 22,5 m", r.Ativo);
        }

        [Fact]
        public void Comprimento_escolhido_junto_vale_sem_ler_o_desenho()
        {
            Construtor.LeituraDeTexto r = Construtor.LerCaboDeTexto(new List<TextoDesenho> { T("R1", Rotulo, 0, 10), T("C1", "18 metros", 30, 30) }, Apoio.Pernambuco,
                () => { throw new InvalidOperationException("não devia ler o desenho"); });
            Assert.Equal(Rotulo + " 18 m", r.Ativo);
            Assert.False(r.ComprimentoDaVizinhanca);
        }

        [Fact]
        public void Longe_demais_nao_entra_e_embaixo_ganha_de_em_cima()
        {
            TextoDesenho rot = T("R1", Rotulo, 0, 10);
            Construtor.LeituraDeTexto longe = Construtor.LerCaboDeTexto(new List<TextoDesenho> { rot }, Apoio.Pernambuco, Desenho(rot, T("C1", "35 m", 0, -5)));
            Assert.Equal(Rotulo, longe.Ativo);
            Assert.Null(longe.Comprimento);
            Construtor.LeituraDeTexto dois = Construtor.LerCaboDeTexto(new List<TextoDesenho> { rot }, Apoio.Pernambuco,
                Desenho(rot, T("CIMA", "50 m", 0, 11.5), T("BAIXO", "35 m", 0, 8)));
            Assert.Equal(Rotulo + " 35 m", dois.Ativo);
        }

        [Fact]
        public void Mais_vao_tambem_le_como_cabo_com_vizinhanca()
        {
            TextoDesenho rot = T("R1", Rotulo, 0, 10);
            TextoLido tl = Construtor.LerAnotacao(new List<TextoDesenho> { rot }, Apoio.Pernambuco, "CABO", Desenho(rot, T("C1", "35 m", 0, 8.5)));
            Assert.Equal("CABO", tl.R.Entidade);
            Assert.Equal(Rotulo + " 35 m", tl.R.Ativo);
        }

        [Fact]
        public void MText_com_dois_cabos_respeita_as_partes_e_a_metragem_de_cada_um()
        {
            const string outro = "3XA185P(7,94)-13,8KV";
            var textos = new List<TextoDesenho> {
                T("A", Rotulo, 0, 10, "01_INSTALANDO"),
                T("A", "35 m", 0, 10, "01_INSTALANDO"),
                T("A", outro, 0, 10, "01_INSTALANDO"),
                T("A", "22 m", 0, 10, "01_INSTALANDO")
            };
            for (int i = 0; i < textos.Count; i++) textos[i].Parte = i;
            var leitura = Construtor.LerCabosDeTextos(textos, Apoio.Pernambuco, null);
            Assert.Equal(new[] { Rotulo + " 35 m", outro + " 22 m" }, leitura.Select(x => x.Ativo));
            Assert.Equal(new[] { 0, 2 }, leitura.Select(x => x.Texto.Parte));

            var obra = Construtor.NovaObra("", Apoio.Pernambuco);
            var linha = new LinhaDesenho { Handle = "L", Pontos = new List<Ponto> { new Ponto(0, 0), new Ponto(10, 0) } };
            Construtor.AdicionarVaos(obra, linha, textos, null, null, Apoio.Pernambuco, null);
            Construtor.AdicionarVaos(obra, linha, textos, null, null, Apoio.Pernambuco, null);
            Assert.Equal(2, obra.Vaos.Count);
            Assert.Equal(obra.Vaos[0].Numero, obra.Vaos[1].IrmaoDe);
            JObj saida = ObraArquivo.Dados(obra);
            Assert.Equal(2, saida.Lista("cabos").Count);
            Assert.Equal(new[] { 0, 2 }, obra.Vaos.Select(x => x.ParteTexto));
        }

        [Fact]
        public void A_mesma_metragem_vizinha_nao_e_copiada_para_outro_cabo()
        {
            var primeiro = T("A", Rotulo, 0, 10, "01_INSTALANDO");
            var segundo = T("B", Rotulo, 2, 10, "01_REMOVER");
            var comprimento = T("C", "30 m", 0, 9);
            var leitura = Construtor.LerCabosDeTextos(new List<TextoDesenho> { primeiro, segundo }, Apoio.Pernambuco, Desenho(comprimento));
            Assert.Equal(new[] { Rotulo + " 30 m", Rotulo }, leitura.Select(x => x.Ativo));
            Assert.Equal(new[] { "I", "R" }, leitura.Select(x => x.Operacao));
        }

        [Theory]
        [InlineData("35 m", true)]
        [InlineData("35,5 M", true)]
        [InlineData("35m", true)]
        [InlineData("40 metros", true)]
        [InlineData("RUA 35", false)]
        [InlineData("N1", false)]
        public void Reconhece_texto_de_comprimento(string t, bool esperado) { Assert.Equal(esperado, Construtor.EhComprimento(t)); }
    }
}
