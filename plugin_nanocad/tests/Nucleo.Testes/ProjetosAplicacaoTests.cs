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
        public void Codigo_027_carrega_o_pacote_da_aplicacao_independente_da_selecao_antiga()
        {
            var transporte = new TransporteFalso();
            transporte.Rotas["GET /api/plugin/regras"] = (c, tok) => TransporteFalso.Ok(
                @"{""projeto"":{""codigo"":""027"",""nome"":""PARAIBA""},""processamento"":[{""id"":""processar027""}],
                    ""classificacao"":[{""id"":""classificar027""}],""vinculacao"":{},
                    ""leitura"":{""layers_rede_mt"":[""PB_MT""],""blocos_poste"":[""PB_POSTE""]}}");
            var cliente = Cliente(transporte);

            var regras = CatalogoRegras.Resolver("027", Apoio.Pernambuco, codigo =>
            {
                Assert.Equal("027", codigo);
                return ClienteAplicacao.RegrasDe(cliente.RegrasCruas(codigo), "http://teste", codigo);
            }, () => new[] { Apoio.Pernambuco });

            Assert.Equal("027", regras.Codigo);
            Assert.Equal("processar027", Assert.Single(regras.Processamento).Str("id"));
            Assert.Equal("classificar027", Assert.Single(regras.Classificacao).Str("id"));
            Assert.True(regras.LeituraDaAplicacao);
            Assert.Equal(new[] { "PB_MT" }, regras.Leitura.LayersRedeMT);
            Assert.Equal(new[] { "PB_POSTE" }, regras.Leitura.BlocosPoste);
        }

        [Fact]
        public void Sem_download_usa_cache_de_027_ou_o_seed_de_027()
        {
            var cache = new RegrasProjeto { Codigo = "027", DaAplicacao = true, Versao = "versao-027" };
            var pendente = new RegrasProjeto { Codigo = "027", DaAplicacao = true };
            var catalogo = CatalogoRegras.Carregar(nome =>
            {
                if (nome == "regras.padrao.processamento.json") return Apoio.Dados("regras_leitor_processamento_seed.json");
                if (nome == "regras.padrao.classificacao.json") return Apoio.Dados("regras_leitor_classificacao_seed.json");
                return null;
            }, new[] { Apoio.Pernambuco });

            var regras = CatalogoRegras.Resolver("027", pendente, _ => null, () => new[] { pendente, cache }.Concat(catalogo));
            Assert.Same(cache, regras);

            var fabrica = CatalogoRegras.Resolver("027", pendente, null, () => catalogo);
            Assert.Equal("027", fabrica.Codigo);
            Assert.Equal("embutida", fabrica.Origem);
            Assert.NotEmpty(fabrica.Processamento);
            Assert.NotEmpty(fabrica.Classificacao);
            Assert.Equal(Json.Escrever(Apoio.Padrao.Processamento.Cast<object>().ToList(), false),
                         Json.Escrever(fabrica.Processamento.Cast<object>().ToList(), false));
            Assert.Equal(Json.Escrever(Apoio.Padrao.Classificacao.Cast<object>().ToList(), false),
                         Json.Escrever(fabrica.Classificacao.Cast<object>().ToList(), false));
            Assert.Null(CatalogoRegras.Resolver("999", Apoio.Pernambuco, _ => null, () => catalogo));
            Assert.Same(cache, CatalogoRegras.Resolver("027", fabrica, _ => cache, () => catalogo));
        }

        [Fact]
        public void Rejeita_resposta_da_aplicacao_com_codigo_diferente()
        {
            var resposta = new JObj().Set("projeto", new JObj().Set("codigo", "081"));
            Assert.Throws<ErroAplicacao>(() => ClienteAplicacao.RegrasDe(resposta, "http://teste", "027"));
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
