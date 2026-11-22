using System.Collections.Generic;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class MontagemManualTests
    {
        private static Poste PosteLido(string numero, string irmao, string fonte, string layer, string cor, string op, string ativo)
        {
            return new Poste { Numero = numero, IrmaoDe = irmao, Handle = "P-DWG", TextoHandle = fonte,
                Entidade = "ESTRUTURA", AtivoLido = ativo, TextoOriginal = ativo, TextoLayer = layer,
                TextoCor = cor, Operacao = op, OperacaoLida = op };
        }

        [Fact]
        public void Agrupa_poste_por_layer_cor_e_operacao_e_rele_sem_duplicar()
        {
            Obra obra = Construtor.NovaObra("teste.dwg", Apoio.Pernambuco);
            System.Func<Obra> ler = () => new Obra { Postes = new List<Poste> {
                PosteLido("P1", null, "10", "INST", "#FF0000", "I", "N1"),
                PosteLido("P1b", "P1", "11", "INST", "#FF0000", "I", "U4"),
                PosteLido("P1c", "P1", "12", "INST", "#FF0000", "R", "N3"),
                PosteLido("P1d", "P1", "13", "INST", "#00FF00", "I", "U5")
            }};
            MontagemManual.Incorporar(obra, ler(), Apoio.Pernambuco);
            Assert.Equal(new[] { "N1 U4", "N3", "U5" }, obra.Postes.Select(p => p.AtivoLido));
            Assert.Equal(3, obra.GrupoDoPoste(obra.Postes[0]).Count);
            MontagemManual.Incorporar(obra, ler(), Apoio.Pernambuco);
            Assert.Equal(new[] { "N1 U4", "N3", "U5" }, obra.Postes.Select(p => p.AtivoLido));
        }

        [Fact]
        public void Linhas_que_cruzam_regiao_sao_inteiras_e_distintas()
        {
            var poligono = new[] { new Ponto(0, 0), new Ponto(10, 0), new Ponto(10, 10), new Ponto(0, 10) };
            var d = new Desenho();
            d.Linhas.Add(new LinhaDesenho { Handle = "A", Pontos = new List<Ponto> { new Ponto(-10, 5), new Ponto(7, 5) } });
            d.Linhas.Add(new LinhaDesenho { Handle = "B", Pontos = new List<Ponto> { new Ponto(7, 5), new Ponto(20, 5) } });
            d.Linhas.Add(new LinhaDesenho { Handle = "C", Pontos = new List<Ponto> { new Ponto(-2, -1), new Ponto(15, -1) } });
            MontagemManual.Recortar(d, poligono);
            Assert.Equal(new[] { "A", "B" }, d.Linhas.Select(l => l.Handle));
            Assert.Equal(17.0, Geometria.Comprimento(d.Linhas[0].Pontos));
            Assert.Equal(13.0, Geometria.Comprimento(d.Linhas[1].Pontos));
        }

        [Fact]
        public void A_regiao_nao_junta_segmentos_contiguos_da_mesma_layer()
        {
            var d = new Desenho();
            d.Linhas.Add(new LinhaDesenho { Handle = "A", Layer = "REDE_MT", Cor = "#FF0000",
                Pontos = new List<Ponto> { new Ponto(0, 0), new Ponto(40, 0) } });
            d.Linhas.Add(new LinhaDesenho { Handle = "B", Layer = "REDE_MT", Cor = "#FF0000",
                Pontos = new List<Ponto> { new Ponto(40, 0), new Ponto(80, 0) } });
            var obra = LeitorProjeto.Ler(d, Apoio.Pernambuco, null, new List<TextoLido>(), true);
            Assert.Equal(new[] { "A", "B" }, obra.Vaos.Select(v => v.LinhaHandle).OrderBy(h => h));
            Assert.All(obra.Vaos, v => Assert.Equal(40, Geometria.Comprimento(v.Pontos)));
        }

        [Fact]
        public void Depois_do_envio_nova_anotacao_do_mesmo_grupo_fica_em_um_novo_lote()
        {
            var obra = Construtor.NovaObra("teste.dwg", Apoio.Pernambuco);
            var a = PosteLido("P1", null, "10", "INST", "#FF0000", "I", "N1");
            MontagemManual.Incorporar(obra, new Obra { Postes = new List<Poste> { a } }, Apoio.Pernambuco);
            obra.Postes[0].EnviadaNaMontagem = true;
            var b = PosteLido("P1", null, "11", "INST", "#FF0000", "I", "U4");
            MontagemManual.Incorporar(obra, new Obra { Postes = new List<Poste> { b } }, Apoio.Pernambuco);
            Assert.Equal(new[] { "N1", "U4" }, obra.Postes.Select(p => p.AtivoLido));
            Assert.Equal("U4", Assert.Single(obra.Postes.Where(p => !p.EnviadaNaMontagem)).AtivoLido);
        }

        [Fact]
        public void Editor_e_exportacao_respeitam_comentario_e_bloqueio()
        {
            var obra = Construtor.NovaObra("teste.dwg", Apoio.Pernambuco);
            var l = Construtor.NovaLinha(obra, false, "N3", "I", Apoio.Pernambuco);
            MontagemManual.Editar(obra, l, "Conferir o material", "*M", true, false, true, Apoio.Pernambuco);
            MontagemManual.Editar(obra, l, "Não deve aplicar", "I", false, false, true, Apoio.Pernambuco);
            var serial = ObraArquivo.Gerar(obra, System.DateTime.UtcNow);
            var recarregada = ObraArquivo.Ler(serial, Apoio.Pernambuco);
            Assert.True(recarregada.Avulsos[0].Comentario && recarregada.Avulsos[0].Congelada && recarregada.Avulsos[0].Bloqueada);
            var linha = (JObj)ObraArquivo.Dados(recarregada).Lista("outros")[0];
            Assert.Equal("*M", linha.Str("operacao"));
            Assert.Equal("Conferir o material", linha.Str("ativo"));
            Assert.True(linha.Bool("comentario", false));
        }
    }
}
