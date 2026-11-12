// TASK-017: autocompletar do ativo no plugin (mesmo comportamento de static/ativo_dica.js da aplicação).
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class AutocompletarTests
    {
        private static CatalogoBase Cat()
        {
            return CatalogoBase.De((JObj)Json.Ler(@"{""versao"":""v1"",
              ""ativos"":[{""nome"":""N1"",""desc"":""ESTRUTURA N1"",""origem"":""OUTROS""},{""nome"":""N2"",""desc"":""ESTRUTURA N2""},
                          {""nome"":""N3"",""desc"":""ESTRUTURA N3""},{""nome"":""A35P"",""desc"":""CABO A35P"",""origem"":""CABO""},
                          {""nome"":""U3"",""desc"":""ESTRAI""},{""nome"":""DT11/300"",""desc"":""POSTE DT""}],
              ""apelidos"":{""11300D"":""DT11/300""},
              ""parametros"":{""N1"":[{""nome"":""tipocabo"",""padrao"":""P"",""valores"":[""P"",""M""],""mascaras"":[],""aberto"":false,""porNome"":false},
                                     {""nome"":""poste"",""padrao"":null,""valores"":[],""mascaras"":[],""aberto"":true,""porNome"":false}]}}"));
        }

        [Fact]
        public void Completa_os_parenteses_como_a_aplicacao()
        {
            Assert.Equal("N1(D, X)", Autocompletar.CompletarParenteses("N1(D, X"));
            Assert.Equal("IT_2(D, 3)", Autocompletar.CompletarParenteses("IT_2(D, 3, "));
            Assert.Equal("N1(A(B))", Autocompletar.CompletarParenteses("N1(A(B"));
            Assert.Equal("[CAZ7,94(", Autocompletar.CompletarParenteses("[CAZ7,94("));         // colchete aberto: não mexe
            Assert.Equal("N1(P), U3", Autocompletar.CompletarParenteses("N1(P), U3"));
        }

        [Fact]
        public void Sugere_nomes_da_palavra_atual_com_quantidade_e_apelido()
        {
            CatalogoBase c = Cat();
            PalavraAtual p = Autocompletar.Palavra("N1(P), 2-U", 10);
            Assert.Equal("2-", p.Prefixo);
            Assert.Equal("U", p.Parcial);
            Assert.Equal("U3", Autocompletar.Sugerir(c, p.Parcial, 10, false)[0].Nome);
            Assert.Equal(new[] { "N1", "N2", "N3" }, Autocompletar.Sugerir(c, "n", 10, false).Select(x => x.Nome).ToArray());
            ItemCatalogo apel = Autocompletar.Sugerir(c, "113", 10, false).First();
            Assert.Equal("DT11/300", apel.ApelidoDe);
            PalavraAtual cabo = Autocompletar.Palavra("3xA3", 4);
            Assert.Equal("3x", cabo.Prefixo);
            Assert.Equal("A35P", Autocompletar.Sugerir(c, cabo.Parcial, 10, true)[0].Nome);
            int cur;
            Assert.Equal("3xA35P", Autocompletar.Substituir("3xA3", cabo.Inicio, 4, "A35P", out cur));
            Assert.Equal(6, cur);
        }

        [Fact]
        public void Dentro_do_parentese_mostra_parametros_e_valores_do_atual()
        {
            CatalogoBase c = Cat();
            ContextoParentese ctx = Autocompletar.ContextoDoParentese("N1(", 3);
            Assert.Equal("N1", ctx.Nome);
            Assert.Equal(0, ctx.Indice);
            DicaAtivo d = Autocompletar.Dica(c, ctx.Nome);
            Assert.Equal("parametros", d.Tipo);
            Assert.Contains("»1. tipocabo [P, M] (padrão P)«", Autocompletar.TextoDaDica(d, ctx));
            Assert.Equal(new[] { "P", "M" }, Autocompletar.ValoresDoParametro(d, ctx).ToArray());
            ContextoParentese segundo = Autocompletar.ContextoDoParentese("2-N1(P, 9", 9);
            Assert.Equal("N1", segundo.Nome);
            Assert.Equal(1, segundo.Indice);
            Assert.Equal("9", segundo.Parcial);
            Assert.Equal("sem_parametros", Autocompletar.Dica(c, "U3").Tipo);
            Assert.Equal("nao_encontrado", Autocompletar.Dica(c, "XPTO").Tipo);
            Assert.Equal("sem_parametros", Autocompletar.Dica(c, "11300D").Tipo);      // apelido do DT11/300
            Assert.Null(Autocompletar.ContextoDoParentese("N1(P)", 5));
        }
    }
}
