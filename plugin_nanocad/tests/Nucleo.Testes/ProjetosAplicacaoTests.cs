// TASK-015: projetos da aplicação na tela inicial, regras do projeto escolhido e configuração de leitura guardada na aplicação.
using System.Collections.Generic;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class ProjetosAplicacaoTests
    {
        private static ClienteAplicacao Cliente(TransporteFalso t)
        {
            return new ClienteAplicacao(new ConexaoAplicacao { Url = "127.0.0.1:8000", Token = "tok" }, t);
        }

        [Fact]
        public void Lista_os_projetos_cadastrados()
        {
            TransporteFalso t = new TransporteFalso();
            t.Rotas["GET /api/plugin/projetos"] = (c, tok) => TransporteFalso.Ok("{\"projetos\":[{\"codigo\":\"081\",\"nome\":\"PERNAMBUCO\",\"tem_leitura\":true},{\"codigo\":\"071\",\"nome\":\"BAHIA\"}]}");
            List<ProjetoAplicacao> l = Cliente(t).Projetos();
            Assert.Equal(new[] { "081", "071" }, l.Select(p => p.Codigo).ToArray());
            Assert.Equal("Aplicação: BAHIA (071)", l[1].Rotulo);
            Assert.True(l[0].TemLeitura);
        }

        [Fact]
        public void Regras_trazem_a_configuracao_de_leitura_do_projeto()
        {
            JObj resp = (JObj)Json.Ler("{\"projeto\":{\"codigo\":\"081\",\"nome\":\"PERNAMBUCO\"},\"processamento\":[],\"classificacao\":[],\"vinculacao\":{},"
                + "\"leitura\":{\"layers_rede_mt\":[\"REDE_MT\",\"ANOTACOES\",\"=0\"],\"blocos_poste\":[\"pg\",\"KMLP\"]}}");
            RegrasProjeto r = ClienteAplicacao.RegrasDe(resp, "http://x");
            Assert.True(r.LeituraDaAplicacao);
            Assert.Equal(new[] { "REDE_MT", "ANOTACOES", "=0" }, r.Leitura.LayersRedeMT.ToArray());
            Assert.Equal(new[] { "pg", "KMLP" }, r.Leitura.BlocosPoste.ToArray());

            JObj sem = (JObj)Json.Ler("{\"projeto\":{\"codigo\":\"071\",\"nome\":\"BAHIA\"},\"processamento\":[],\"classificacao\":[],\"vinculacao\":{},\"leitura\":null}");
            RegrasProjeto r2 = ClienteAplicacao.RegrasDe(sem, "http://x");
            Assert.False(r2.LeituraDaAplicacao);
            Assert.Equal(new[] { "REDE_MT" }, r2.Leitura.LayersRedeMT.ToArray());     // padrão
        }

        [Fact]
        public void Salvar_leitura_manda_a_configuracao_para_o_projeto()
        {
            TransporteFalso t = new TransporteFalso();
            string corpo = null;
            t.Rotas["PUT /api/plugin/leitura"] = (c, tok) => { corpo = c; return TransporteFalso.Ok("{\"aviso\":null}"); };
            ConfigLeitura cfg = new ConfigLeitura();
            cfg.BlocosPoste = new List<string> { "pg", "KMLP" };
            Assert.Null(Cliente(t).SalvarLeitura("081", cfg));
            JObj enviado = (JObj)Json.Ler(corpo);
            Assert.Equal(new[] { "pg", "KMLP" }, Json.ListaDeTextos(enviado.Obj("leitura")["blocos_poste"]).ToArray());
            Assert.Contains(t.Chamadas, x => x.StartsWith("PUT /api/plugin/leitura"));
        }
    }
}
