// Painel em duas listas (Cabos · Postes e estruturas) e o filtro pelas regras do leitor (TASK-012 do plugin).
using System.Collections.Generic;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class PainelTests
    {
        private static Obra Ler() { return LeitorProjeto.Ler(DesenhoExemplo.Montar(), Apoio.Pernambuco); }

        [Fact]
        public void Cabos_e_outros_ficam_em_listas_separadas_como_no_resumo_da_aplicacao()
        {
            Obra o = Ler();
            List<LinhaTabela> cabos = Relatorios.Cabos(o, null), outros = Relatorios.Outros(o, null);
            Assert.NotEmpty(cabos);
            Assert.NotEmpty(outros);
            Assert.All(cabos, l => Assert.True(o.PorNumero(l.Chave) is Vao || o.PorNumero(l.Chave).Entidade == "CABO"));
            Assert.All(outros, l => Assert.False(o.PorNumero(l.Chave) is Vao));
            // as mesmas linhas que vão para Cabos/Outros na aplicação
            List<Elemento> c, ou;
            ObraArquivo.ElementosDasLinhas(o, out c, out ou);
            Assert.Equal(c.Select(e => e.Numero).OrderBy(x => x), cabos.Select(l => l.Chave).OrderBy(x => x));
            Assert.Equal(ou.Select(e => e.Numero).OrderBy(x => x), outros.Select(l => l.Chave).OrderBy(x => x));
            // ícones: estado, rede MT/BT, operação
            LinhaTabela v1 = cabos.First(l => l.Chave == "V1");
            Assert.Equal("estado:" + o.VaoPorNumero("V1").Estado, v1.Icones[0]);
            Assert.Equal("rede:" + o.VaoPorNumero("V1").Tipo, v1.Icones[2]);
            Assert.StartsWith("op:", v1.Icones[3]);
            Assert.Equal("poste", outros.First(l => l.Chave.StartsWith("P")).Icones[2]);
            Assert.Equal(Relatorios.Contar(o, null, null), cabos.Count + outros.Count);
        }

        [Fact]
        public void Ler_mostra_so_o_que_as_regras_classificaram_e_construir_mostra_tudo()
        {
            Obra o = Ler();
            Vao existente = new Vao { Numero = "V99", Tipo = "BT", Pontos = new List<Ponto> { new Ponto(0, 0), new Ponto(10, 0) } };
            Poste semAnot = new Poste { Numero = "P99" };
            o.Vaos.Add(existente);
            o.Postes.Add(semAnot);
            Assert.False(Relatorios.Visivel(o, existente));
            Assert.DoesNotContain(Relatorios.Cabos(o, null), l => l.Chave == "V99");
            Assert.DoesNotContain(Relatorios.Outros(o, null), l => l.Chave == "P99");
            Assert.DoesNotContain(Relatorios.Vinculos(o, Apoio.Pernambuco), l => l.Chave == "V99");
            Assert.True(Relatorios.Ocultos(o) >= 2);
            Assert.Contains("sem anotação ocultos", Relatorios.Resumo(o));

            existente.Manual = true;                                  // apontado à mão: aparece mesmo sem anotação
            Assert.Contains(Relatorios.Cabos(o, null), l => l.Chave == "V99");

            o.Modo = "construir";                                     // construir: sem filtro das regras
            Assert.Contains(Relatorios.Outros(o, null), l => l.Chave == "P99");
            Assert.Equal(0, Relatorios.Ocultos(o));
        }

        [Fact]
        public void Filtro_de_estado_vale_para_as_duas_listas()
        {
            Obra o = Ler();
            foreach (string est in new[] { Estados.Ok, Estados.Conferir, Estados.Erro })
            {
                Assert.All(Relatorios.Cabos(o, est).Concat(Relatorios.Outros(o, est)), l => Assert.Equal(est, l.Estado));
                Assert.Equal(Relatorios.Contar(o, null, est), Relatorios.Cabos(o, est).Count + Relatorios.Outros(o, est).Count);
            }
        }
    }
}
