using System.Collections.Generic;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    /// <summary>Anotação de poste longe do bloco (sem guia, além do raio de reserva): casa 1:1 com o bloco LIVRE mais perto
    /// (até RaioBlocoLivreAlturas × a altura) e vai para Conferir; a que sobra vira poste "só anotação" com a posição aproximada
    /// (`_posteAprox`) para o mapa da aplicação (TASK-233 da aplicação).</summary>
    public class BlocoLivreTests
    {
        // altura 8: raio 32, reserva 96, bloco livre 200 (unidades do desenho)
        private static Desenho Montar()
        {
            Desenho d = new Desenho { Arquivo = @"C:\Projetos\Bloco_livre.dwg" };
            d.Postes.Add(new PosteDesenho { Handle = "B1", X = 0, Y = 0, Bloco = "POSTE", Layer = "POSTES" });
            d.Postes.Add(new PosteDesenho { Handle = "B2", X = 1000, Y = 0, Bloco = "POSTE", Layer = "POSTES" });
            d.Postes.Add(new PosteDesenho { Handle = "B3", X = 5000, Y = 0, Bloco = "POSTE", Layer = "POSTES" });   // longe de tudo: fica livre
            d.Textos.Add(new TextoDesenho { Handle = "T1", Texto = "9200D, N1", X = 120, Y = 0, Layer = "01_INSTALANDO", Cor = "#000000", Altura = 8 });
            d.Textos.Add(new TextoDesenho { Handle = "T2", Texto = "9400D, N1", X = 1110, Y = 0, Layer = "01_INSTALANDO", Cor = "#000000", Altura = 8 });
            d.Textos.Add(new TextoDesenho { Handle = "T3", Texto = "9300F, N3", X = 1150, Y = -50, Layer = "01_INSTALANDO", Cor = "#000000", Altura = 8 });
            return d;
        }

        [Fact]
        public void Anotacao_longe_vai_para_o_bloco_livre_mais_perto_um_para_um()
        {
            Obra o = LeitorProjeto.Ler(Montar(), Apoio.Pernambuco);
            Poste b1 = o.Postes.Single(p => p.Handle == "B1"), b2 = o.Postes.Single(p => p.Handle == "B2"), b3 = o.Postes.Single(p => p.Handle == "B3");
            Assert.Equal("9200D, N1", b1.AtivoLido);
            Assert.Equal("9400D, N1", b2.AtivoLido);                       // T2 (110) ganha de T3 (~158): cada bloco recebe uma anotação só
            Assert.False(b3.TemAnotacao);
            Assert.Equal(Estados.Conferir, b1.Estado);
            Assert.Contains(b1.MensagensFixas, m => m.Contains("bloco de poste livre mais perto (120 m"));
            Poste solto = o.Postes.Single(p => p.AtivoLido == "9300F, N3");
            Assert.Equal("", solto.Handle);
            Assert.Contains(solto.MensagensFixas, m => m.Contains("sem bloco de poste por perto"));
        }

        [Fact]
        public void Raio_do_bloco_livre_vem_da_configuracao()
        {
            // regras próprias (Apoio.Pernambuco é compartilhado entre as classes de teste, que rodam em paralelo)
            RegrasProjeto r = RegrasProjeto.Montar("081", "Pernambuco (081) — Leitor", "teste",
                Apoio.Dados("regras_leitor_processamento_pernambuco_seed.json"), Apoio.Dados("regras_leitor_classificacao_pernambuco_seed.json"),
                Apoio.Dados("vinculacao_pernambuco_seed.json"));
            r.Leitura = ConfigLeitura.De(r.Leitura.ParaJson().Set("raio_bloco_livre_alturas", 10.0));   // 80: nenhum bloco ao alcance
            Obra o = LeitorProjeto.Ler(Montar(), r);
            Assert.True(o.Postes.Where(p => p.Handle != "").All(p => !p.TemAnotacao));
            Assert.Equal(3, o.Postes.Count(p => p.Handle == "" && p.TemAnotacao));
            Assert.Equal(25.0, new ConfigLeitura().RaioBlocoLivreAlturas);
            Assert.Equal(25.0, ConfigLeitura.De(new ConfigLeitura().ParaJson()).RaioBlocoLivreAlturas);
        }

        [Fact]
        public void Poste_sem_bloco_leva_a_posicao_aproximada_para_a_aplicacao()
        {
            Obra o = LeitorProjeto.Ler(Montar(), Apoio.Pernambuco);
            List<JObj> outros = ObraArquivo.Dados(o).Lista("outros").Cast<JObj>().ToList();
            JObj comBloco = outros.Single(l => l.Str("ativo") == "9200D, N1"), semBloco = outros.Single(l => l.Str("ativo") == "9300F, N3");
            Assert.Equal(new object[] { 0.0, 0.0 }, comBloco.Lista("_poste").ToArray());
            Assert.False(comBloco.Has("_posteAprox"));
            Assert.False(semBloco.Has("_poste"));
            Assert.Equal(new object[] { 1150.0, -50.0 }, semBloco.Lista("_posteAprox").ToArray());
        }
    }
}
