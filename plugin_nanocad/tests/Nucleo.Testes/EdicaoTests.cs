// TASK-017: no plugin, corrigir o ativo, tirar/devolver linhas da obra e criar linhas sem apontar no desenho.
using System.Collections.Generic;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class EdicaoTests
    {
        private static Obra Ler() { return LeitorProjeto.Ler(DesenhoExemplo.Montar(), Apoio.Pernambuco); }

        private static List<string> Ativos(Obra o, string tabela)
        {
            return ((JObj)ObraArquivo.Dados(o)).Lista(tabela).OfType<JObj>().Select(l => l.Str("ativo")).ToList();
        }

        [Fact]
        public void Corrigir_o_ativo_vale_na_obra_e_voltar_ao_lido_desfaz()
        {
            Obra o = Ler();
            Vao v1 = o.VaoPorNumero("V1");
            string doDesenho = v1.AtivoLido;
            string novo = doDesenho.Replace("35", "70");
            Construtor.TrocarAtivo(o, v1, novo, Apoio.Pernambuco);
            Assert.True(v1.AtivoEditado);
            Assert.Equal(doDesenho, v1.AtivoDoDesenho);
            Assert.Contains(v1.Ativo, Ativos(o, "cabos"));
            Assert.Contains(Relatorios.Cabos(o, null), l => l.Chave == "V1" && l.Celulas.Last().Contains("corrigido"));
            Construtor.VoltarAoLido(o, v1, Apoio.Pernambuco);
            Assert.False(v1.AtivoEditado);
            Assert.Equal(doDesenho, v1.AtivoLido);
        }

        [Fact]
        public void Excluir_tira_da_obra_e_dos_vinculos_e_restaurar_devolve()
        {
            Obra o = Ler();
            Vao v1 = o.VaoPorNumero("V1");
            int cabosAntes = Ativos(o, "cabos").Count;
            Construtor.Excluir(o, v1, Apoio.Pernambuco);
            Assert.True(v1.Excluido);
            Assert.Null(v1.VinculoA);
            Assert.Null(v1.VinculoB);
            Assert.Equal(cabosAntes - 1, Ativos(o, "cabos").Count);
            Assert.DoesNotContain(Relatorios.Cabos(o, null), l => l.Chave == "V1");
            Assert.Contains(Relatorios.TextosIgnorados(o), l => l.Chave == "V1" && l.Celulas[3].Contains("restaurar"));
            Construtor.Restaurar(o, v1, Apoio.Pernambuco);
            Assert.False(v1.Excluido);
            Assert.Equal(cabosAntes, Ativos(o, "cabos").Count);
        }

        [Fact]
        public void Poste_excluido_nao_recebe_vinculo()
        {
            Obra o = Ler();
            string num = o.Vaos.SelectMany(v => v.Vinculos).Select(x => x.Poste).First();     // um poste que recebe vínculo
            Poste p = o.PostePorNumero(num);
            int outrosAntes = Ativos(o, "outros").Count;
            Construtor.Excluir(o, p, Apoio.Pernambuco);
            Assert.Equal(outrosAntes - 1, Ativos(o, "outros").Count);
            Assert.DoesNotContain(o.Vaos, v => v.Vinculos.Any(x => x.Poste == num));
        }

        [Fact]
        public void Linha_nova_sem_desenho_vai_para_cabos_ou_outros_e_excluir_apaga()
        {
            Obra o = Ler();
            Avulso c = Construtor.NovaLinha(o, true, "3xA35P(7,94)-13,8kV 40 m", "I", Apoio.Pernambuco);
            Avulso x = Construtor.NovaLinha(o, false, "2-U3", "R", Apoio.Pernambuco);
            Assert.Equal("L1", c.Numero);
            Assert.Equal("L2", x.Numero);
            Assert.Contains("3xA35P(7,94)-13,8kV 40 m", Ativos(o, "cabos"));
            Assert.Contains("2-U3", Ativos(o, "outros"));
            Assert.Contains(Relatorios.Outros(o, null), l => l.Chave == "L2");
            Construtor.Excluir(o, x, Apoio.Pernambuco);
            Assert.DoesNotContain(o.Avulsos, a => a.Numero == "L2");
        }

        [Fact]
        public void Atualizar_mantem_correcao_exclusao_e_linhas_novas()
        {
            Obra o = Ler();
            Poste p2 = o.PostePorNumero("P2");
            Vao v2 = o.VaoPorNumero("V2");
            Construtor.Excluir(o, p2, Apoio.Pernambuco);
            Construtor.TrocarAtivo(o, v2, v2.AtivoLido.Replace("35", "50"), Apoio.Pernambuco);
            Construtor.NovaLinha(o, false, "N1(9200D, P)", "I", Apoio.Pernambuco);
            Obra nova = Ler();
            Construtor.ReaplicarAjustes(nova, o, Apoio.Pernambuco);
            Assert.True(nova.PostePorNumero("P2").Excluido);
            Assert.Equal(v2.AtivoLido, nova.VaoPorNumero("V2").AtivoLido);
            Assert.True(nova.VaoPorNumero("V2").AtivoEditado);
            Assert.Contains(nova.Avulsos, a => a.Manual && a.AtivoLido == "N1(9200D, P)");
        }

        [Fact]
        public void Correcao_exclusao_e_linha_nova_vao_e_voltam_pelo_arquivo()
        {
            Obra o = Ler();
            Construtor.Excluir(o, o.PostePorNumero("P2"), Apoio.Pernambuco);
            Vao v1 = o.VaoPorNumero("V1");
            Construtor.TrocarAtivo(o, v1, v1.AtivoLido.Replace("35", "70"), Apoio.Pernambuco);
            Construtor.NovaLinha(o, true, "S04 12 m", "I", Apoio.Pernambuco);
            Obra lida = ObraArquivo.Ler(ObraArquivo.Gerar(o, System.DateTime.UtcNow), Apoio.Pernambuco);
            Assert.True(lida.PostePorNumero("P2").Excluido);
            Assert.True(lida.VaoPorNumero("V1").AtivoEditado);
            Assert.Contains(lida.Avulsos, a => a.Manual && a.AtivoLido == "S04 12 m" && a.Entidade == "CABO");
        }
    }
}
