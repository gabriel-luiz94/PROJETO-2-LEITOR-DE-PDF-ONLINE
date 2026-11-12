// TASK-018: "+ Cabo / + Outros" lendo um TEXT/MTEXT escolhido no desenho, pelas regras do leitor.
using System.Collections.Generic;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class LinhaDeTextoTests
    {
        private static Obra Ler() { return LeitorProjeto.Ler(DesenhoExemplo.Montar(), Apoio.Pernambuco); }

        private static TextoDesenho Texto(Elemento e, string handle)
        {
            return new TextoDesenho { Handle = handle, Texto = e.TextoOriginal, Layer = e.TextoLayer, Cor = e.TextoCor, X = 10, Y = 20, Altura = 1 };
        }

        [Fact]
        public void Texto_de_cabo_e_de_poste_lidos_pelas_regras_vao_para_a_tabela_certa()
        {
            Obra o = Ler();
            Vao v1 = o.VaoPorNumero("V1");
            Construtor.LeituraDeTexto c = Construtor.LerTextoParaLinha(new List<TextoDesenho> { Texto(v1, "AA1") }, Apoio.Pernambuco);
            Assert.True(c.Reconhecido);
            Assert.True(c.Cabo.Value);
            Assert.Equal(v1.AtivoLido, c.Ativo);
            Assert.Equal(v1.OperacaoLida, c.Operacao);
            Poste p = o.Postes.First(x => x.TemAnotacao);
            Construtor.LeituraDeTexto q = Construtor.LerTextoParaLinha(new List<TextoDesenho> { Texto(p, "AA2") }, Apoio.Pernambuco);
            Assert.False(q.Cabo.Value);
            Assert.Equal(p.AtivoLido, q.Ativo);
        }

        [Fact]
        public void Texto_que_as_regras_nao_entendem_entra_como_esta()
        {
            TextoDesenho t = new TextoDesenho { Handle = "AA3", Texto = "TRAFO 75 KVA NOVO", Layer = "0", Cor = "#FFFFFF", X = 1, Y = 1 };
            Construtor.LeituraDeTexto r = Construtor.LerTextoParaLinha(new List<TextoDesenho> { t }, Apoio.Pernambuco);
            Assert.False(r.Reconhecido);
            Assert.Null(r.Cabo);
            Assert.Equal("TRAFO 75 KVA NOVO", r.Ativo);
            Assert.Equal("I", r.Operacao);
        }

        [Fact]
        public void Linha_ligada_ao_texto_guarda_o_lido_e_sobrevive_ao_atualizar()
        {
            Obra o = Ler();
            Vao v1 = o.VaoPorNumero("V1");
            Construtor.LeituraDeTexto c = Construtor.LerTextoParaLinha(new List<TextoDesenho> { Texto(v1, "ZZ9") }, Apoio.Pernambuco);
            Avulso a = Construtor.NovaLinha(o, true, c.Ativo.Replace("35", "40"), "I", Apoio.Pernambuco);   // ajustou antes de adicionar
            Construtor.LigarTexto(a, c);
            Assert.Equal("ZZ9", a.TextoHandle);
            Assert.True(a.AtivoEditado);
            Assert.Equal(c.Ativo, a.AtivoDoDesenho);
            Assert.Equal(10, a.TextoX);
            Obra nova = Ler();
            Construtor.ReaplicarAjustes(nova, o, Apoio.Pernambuco);
            Avulso volta = nova.Avulsos.Single(x => x.TextoHandle == "ZZ9");
            Assert.Equal(a.AtivoLido, volta.AtivoLido);
            Assert.True(volta.Manual);
        }

        [Fact]
        public void Se_a_leitura_passa_a_pegar_o_texto_a_linha_manual_nao_fica_em_dobro()
        {
            Obra o = Ler();
            Vao v1 = o.VaoPorNumero("V1");
            Avulso a = Construtor.NovaLinha(o, true, v1.AtivoLido, "I", Apoio.Pernambuco);
            Construtor.LigarTexto(a, new Construtor.LeituraDeTexto { Ativo = v1.AtivoLido, Texto = Texto(v1, v1.TextoHandle) });
            Obra nova = Ler();
            Construtor.ReaplicarAjustes(nova, o, Apoio.Pernambuco);
            Assert.DoesNotContain(nova.Avulsos, x => x.Manual && x.TextoHandle == v1.TextoHandle);
        }

        // TASK-019: comprimento na linha de baixo vai para a linha do cabo, sem vírgula — também fora das layers de operação
        private static List<TextoDesenho> Cabo(string layer, string cor, bool mtext, params string[] linhas)
        {
            List<TextoDesenho> l = new List<TextoDesenho>();
            for (int k = 0; k < linhas.Length; k++)
                l.Add(mtext ? new TextoDesenho { Handle = "M1", Parte = k, Tipo = "MTEXT", Texto = linhas[k], Layer = layer, Cor = cor, X = 0, Y = 10 }
                            : new TextoDesenho { Handle = "T" + k, Texto = linhas[k], Layer = layer, Cor = cor, X = 0, Y = 10 - 2 * k });
            return l;
        }

        [Theory]
        [InlineData("01_INSTALANDO", "#000000", false)]
        [InlineData("0", "#FFFFFF", false)]            // layer que não é de operação: as regras não classificam, mas o processamento junta
        [InlineData("0", "#FFFFFF", true)]             // as duas linhas no mesmo MTEXT
        [InlineData("ANOTACOES", "#FF0000", true)]
        public void Comprimento_da_linha_de_baixo_entra_na_linha_do_cabo_sem_virgula(string layer, string cor, bool mtext)
        {
            Construtor.LeituraDeTexto r = Construtor.LerTextoParaLinha(Cabo(layer, cor, mtext, "3XA35P(7,94)-13,8KV", "35 M"), Apoio.Pernambuco);
            Assert.Equal("3XA35P(7,94)-13,8KV 35 m", r.Ativo);
            Assert.DoesNotContain(",  ", r.Ativo);
            Assert.True(r.Cabo.Value);
            TextoLido v = Construtor.LerAnotacao(Cabo(layer, cor, mtext, "3XA35P(7,94)-13,8KV", "35,5 m"), Apoio.Pernambuco, "CABO");    // "+ Vão"
            Assert.Equal("3XA35P(7,94)-13,8KV 35,5 m", v.R.Ativo);
        }

        [Fact]
        public void Linha_de_baixo_que_nao_e_comprimento_nao_entra_no_cabo()
        {
            Construtor.LeituraDeTexto r = Construtor.LerTextoParaLinha(Cabo("0", "#FFFFFF", false, "3XA35P(7,94)-13,8KV", "RUA DAS FLORES"), Apoio.Pernambuco);
            Assert.DoesNotContain(" m", r.Ativo.Replace("RUA", ""));
            Assert.StartsWith("3XA35P(7,94)-13,8KV", r.Ativo);
        }
    }
}
