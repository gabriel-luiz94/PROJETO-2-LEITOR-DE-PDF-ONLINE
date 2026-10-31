using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    /// <summary>Transporte falso: responde por "MÉTODO caminho" e guarda as chamadas.</summary>
    public class TransporteFalso : ITransporteHttp
    {
        public readonly List<string> Chamadas = new List<string>();
        public readonly Dictionary<string, Func<string, string, RespostaHttp>> Rotas = new Dictionary<string, Func<string, string, RespostaHttp>>();

        public static RespostaHttp Ok(string json) { return new RespostaHttp { Status = 200, Corpo = json }; }

        public RespostaHttp Enviar(string metodo, string url, string corpoJson, string token, int timeoutMs)
        {
            string caminho = new Uri(url).AbsolutePath;
            Chamadas.Add(metodo + " " + caminho + " " + (token ?? "-"));
            Func<string, string, RespostaHttp> f;
            if (Rotas.TryGetValue(metodo + " " + caminho, out f)) return f(corpoJson, token);
            return new RespostaHttp { Status = 404, Corpo = "{\"detail\":\"Not Found\"}" };
        }
    }

    public class ClienteAplicacaoTests
    {
        private static ClienteAplicacao Cliente(TransporteFalso t)
        {
            return new ClienteAplicacao(new ConexaoAplicacao { Url = "127.0.0.1:8000/" }, t);
        }

        [Fact]
        public void Login_guarda_tokens_e_nao_guarda_senha()
        {
            TransporteFalso t = new TransporteFalso();
            t.Rotas["POST /api/auth/login"] = (corpo, tok) =>
            {
                JObj o = (JObj)Json.Ler(corpo);
                return o.Str("password") == "certa" ? TransporteFalso.Ok("{\"access_token\":\"A1\",\"refresh_token\":\"R1\",\"email\":\"e@x\"}")
                                                     : new RespostaHttp { Status = 401, Corpo = "{\"detail\":\"Credenciais inválidas\"}" };
            };
            ClienteAplicacao c = Cliente(t);
            int salvos = 0;
            c.TokensMudaram = x => salvos++;
            ErroAplicacao e = Assert.Throws<ErroAplicacao>(() => c.Entrar("e@x", "errada"));
            Assert.True(e.PrecisaEntrar);
            c.Entrar("e@x", "certa");
            Assert.Equal("http://127.0.0.1:8000", c.Conexao.Base);
            Assert.Equal("A1", c.Conexao.Token);
            Assert.Equal("R1", c.Conexao.TokenRenovacao);
            Assert.Equal(1, salvos);
            Assert.DoesNotContain("certa", Json.Escrever(c.Conexao.ParaJson(true), false));
        }

        [Fact]
        public void Token_vencido_renova_uma_vez_e_repete()
        {
            TransporteFalso t = new TransporteFalso();
            t.Rotas["GET /api/plugin/versao"] = (corpo, tok) => tok == "A2"
                ? TransporteFalso.Ok("{\"projeto\":{\"codigo\":\"081\",\"nome\":\"PERNAMBUCO\"},\"versao\":{\"regras\":\"r\",\"conversao\":\"c\",\"base\":\"b\",\"geral\":\"g\"}}")
                : new RespostaHttp { Status = 401, Corpo = "{}" };
            t.Rotas["POST /api/auth/refresh"] = (corpo, tok) => tok == "R1" ? TransporteFalso.Ok("{\"access_token\":\"A2\"}") : new RespostaHttp { Status = 401 };
            ClienteAplicacao c = Cliente(t);
            c.Conexao.Token = "A1"; c.Conexao.TokenRenovacao = "R1";
            VersaoAplicacao v = c.Versao("081");
            Assert.Equal("b", v.Base);
            Assert.Equal("PERNAMBUCO", v.ProjetoNome);
            Assert.Equal(new[] { "GET /api/plugin/versao A1", "POST /api/auth/refresh R1", "GET /api/plugin/versao A2" }, t.Chamadas.ToArray());
            // renovação também vencida: pede login
            c.Conexao.Token = "X"; c.Conexao.TokenRenovacao = "Y";
            Assert.True(Assert.Throws<ErroAplicacao>(() => c.Versao("081")).PrecisaEntrar);
        }

        [Fact]
        public void Sem_conexao_e_erro_da_aplicacao_viram_mensagem_clara()
        {
            TransporteFalso t = new TransporteFalso();
            t.Rotas["GET /api/plugin/regras"] = (c, tok) => new RespostaHttp { Status = 0, Erro = "Conexão recusada" };
            t.Rotas["GET /api/plugin/versao"] = (c, tok) => new RespostaHttp { Status = 404, Corpo = "{\"detail\":\"Projeto '999' não cadastrado.\"}" };
            ClienteAplicacao cl = Cliente(t);
            ErroAplicacao e = Assert.Throws<ErroAplicacao>(() => cl.RegrasCruas("081"));
            Assert.True(e.SemConexao);
            Assert.Contains("Conexão recusada", e.Message);
            Assert.Equal("Projeto '999' não cadastrado.", Assert.Throws<ErroAplicacao>(() => cl.Versao("999")).Message);
        }

        [Fact]
        public void Regras_da_aplicacao_viram_regras_do_projeto()
        {
            JObj cru = new JObj()
                .Set("projeto", new JObj().Set("codigo", "081").Set("nome", "PERNAMBUCO"))
                .Set("versao", new JObj().Set("regras", "v1"))
                .Set("processamento", Json.Ler(Apoio.Dados("regras_leitor_processamento_pernambuco_seed.json")))
                .Set("classificacao", Json.Ler(Apoio.Dados("regras_leitor_classificacao_pernambuco_seed.json")))
                .Set("vinculacao", Json.Ler(Apoio.Dados("vinculacao_pernambuco_seed.json")));
            RegrasProjeto r = ClienteAplicacao.RegrasDe(cru, "http://127.0.0.1:8000");
            Assert.True(r.DaAplicacao);
            Assert.Equal("081", r.Codigo);
            Assert.Equal("v1", r.Versao);
            Assert.Equal("Aplicação: PERNAMBUCO (081)", r.Nome);
            // lê o desenho igual às regras embutidas
            Obra a = LeitorProjeto.Ler(DesenhoExemplo.Montar(), r), b = LeitorProjeto.Ler(DesenhoExemplo.Montar(), Apoio.Pernambuco);
            Assert.Equal(ObraArquivo.Gerar(b, ObraArquivoTests.Data).Replace(Apoio.Pernambuco.Nome, ""),
                         ObraArquivo.Gerar(a, ObraArquivoTests.Data).Replace(r.Nome, ""));
        }

        [Fact]
        public void Conferencia_da_base_marca_os_elementos_certos()
        {
            Obra o = LeitorProjeto.Ler(DesenhoExemplo.Montar(), Apoio.Pernambuco);
            string antesP1 = string.Join("|", o.PostePorNumero("P1").Mensagens);
            // resposta como a de /api/plugin/conferir: P2 (outros 1) sem N4; V3 (cabos 2) sem S04
            string resp = "{\"projeto\":{\"codigo\":\"081\",\"nome\":\"PERNAMBUCO\"},\"versao\":{\"base\":\"b1\"},\"nao_encontrados\":[\"S04\"],\"linhas\":["
                + "{\"tabela\":\"outros\",\"indice\":1,\"itens\":[{\"ativo\":\"DT9/400\",\"naoEncontrado\":false},{\"ativo\":\"N1\",\"naoEncontrado\":false},"
                + "{\"ativo\":\"N4(P,3XA35P_7.94_-13.8KV)\",\"naoEncontrado\":true,\"foraDoCalculo\":false}]},"
                + "{\"tabela\":\"cabos\",\"indice\":2,\"itens\":[{\"ativo\":\"S04\",\"naoEncontrado\":true,\"foraDoCalculo\":true},{\"ativo\":\"S04\",\"naoEncontrado\":true,\"foraDoCalculo\":true}]},"
                + "{\"tabela\":\"outros\",\"indice\":0,\"itens\":[{\"ativo\":\"DT9/200\",\"naoEncontrado\":false}]}]}";
            ConferenciaBase c = ConferenciaBase.De((JObj)Json.Ler(resp));
            Assert.Equal(2, c.ComPendencia);
            c.Aplicar(o, Apoio.Pernambuco, "PERNAMBUCO");
            Poste p2 = o.PostePorNumero("P2");
            Assert.Equal(new[] { "N4" }, p2.ForaDaBase.ToArray());
            Assert.Contains("N4(P,3XA35P_7.94_-13.8KV) não está na base de orçamento (PERNAMBUCO)", p2.Mensagens);
            Vao v3 = o.VaoPorNumero("V3");
            Assert.Single(v3.MensagensBase);                           // repetido na Totalizadora, uma mensagem só
            Assert.EndsWith("fica fora do orçamento", v3.MensagensBase[0]);
            Assert.Equal(antesP1, string.Join("|", o.PostePorNumero("P1").Mensagens));   // P1 (tudo na base) não muda
            List<LinhaTabela> ativos = Relatorios.Ativos(o);
            Assert.Contains("não está na base", ativos.First(x => x.Celulas[0] == "P2 · N4").Celulas[2]);
            Assert.Contains("não está na base: S04", ativos.First(x => x.Celulas[0] == "V3").Celulas[2]);
            // nova conferência sem pendências limpa as marcas
            ConferenciaBase.De((JObj)Json.Ler("{\"linhas\":[]}")).Aplicar(o, Apoio.Pernambuco, "PERNAMBUCO");
            Assert.Empty(p2.ForaDaBase);
            Assert.DoesNotContain(p2.Mensagens, m => m.Contains("base de orçamento"));
        }

        [Fact]
        public void Linhas_da_obra_na_mesma_ordem_do_arquivo()
        {
            Obra o = LeitorProjeto.Ler(DesenhoExemplo.Montar(), Apoio.Pernambuco);
            List<Elemento> cabos, outros;
            ObraArquivo.ElementosDasLinhas(o, out cabos, out outros);
            JObj d = ObraArquivo.Dados(o);
            Assert.Equal(d.Lista("cabos").Count, cabos.Count);
            Assert.Equal(d.Lista("outros").Count, outros.Count);
            for (int i = 0; i < outros.Count; i++) Assert.Equal(((JObj)d.Lista("outros")[i]).Str("ativo"), outros[i].AtivoLido);
        }

        /// <summary>Ponta a ponta contra a aplicação rodando (uvicorn). Só roda com APP_URL, APP_EMAIL e APP_SENHA definidos.</summary>
        [Fact]
        public void Ponta_a_ponta_com_a_aplicacao_rodando()
        {
            string url = Environment.GetEnvironmentVariable("APP_URL");
            if (string.IsNullOrEmpty(url)) return;
            ClienteAplicacao c = new ClienteAplicacao(new ConexaoAplicacao { Url = url }, new TransporteHttp());
            c.Entrar(Environment.GetEnvironmentVariable("APP_EMAIL"), Environment.GetEnvironmentVariable("APP_SENHA"));
            VersaoAplicacao v = c.Versao("081");
            Assert.False(string.IsNullOrEmpty(v.Geral));
            RegrasProjeto r = ClienteAplicacao.RegrasDe(c.RegrasCruas("081"), url);
            Assert.Equal(13, r.Processamento.Count);
            Obra o = LeitorProjeto.Ler(DesenhoExemplo.Montar(), r);
            ConferenciaBase conf = c.Conferir("081", ObraArquivo.Dados(o));
            conf.Aplicar(o, r, conf.Versao.ProjetoNome);
            Assert.Equal(v.Geral, conf.Versao.Geral);
            string relatorio = string.Join("\n", o.Elementos.Where(e => e.MensagensBase.Count > 0).Select(e => e.Numero + ": " + string.Join("; ", e.MensagensBase)));
            File.WriteAllText(Path.Combine(Path.GetTempPath(), "ponta_a_ponta_plugin.txt"), relatorio);
            c.Conexao.Token = "invalido";                              // força a renovação pelo refresh token real
            Assert.Equal(v.Geral, c.Versao("081").Geral);
        }
    }
}
