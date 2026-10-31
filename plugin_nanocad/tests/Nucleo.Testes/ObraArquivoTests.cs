using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class ObraArquivoTests
    {
        public static readonly DateTime Data = new DateTime(2026, 10, 31, 12, 0, 0, DateTimeKind.Utc);

        private static Obra Ler() { return LeitorProjeto.Ler(DesenhoExemplo.Montar(), Apoio.Pernambuco); }

        [Fact]
        public void Envelope_no_formato_de_importacao_da_aplicacao()
        {
            JObj env = ObraArquivo.Envelope(Ler(), Data);
            Assert.Equal("leitor-obra", env.Str("formato"));
            Assert.Equal(1.0, env.Num("versao"));
            Assert.Equal("obra", env.Str("tipo"));
            Assert.Equal("Projeto_Rua_das_Flores", env.Str("nome"));
            Assert.Equal("081", env.Str("projeto"));
            JObj dados = env.Obj("dados");
            Assert.Equal(new[] { "cabos", "outros", "totalizadora" }, dados.Keys.ToArray());
            List<JObj> outros = dados.Lista("outros").Cast<JObj>().ToList(), cabos = dados.Lista("cabos").Cast<JObj>().ToList();
            Assert.Equal(8, outros.Count);
            Assert.Equal(7, cabos.Count);
            Assert.Equal("POSTE", outros[1].Str("entidade"));
            Assert.Equal("9400D, N1, N4", outros[1].Str("ativo"));
            Assert.Equal(2, outros[1].Lista("_poste").Count);
            JObj v1 = cabos[0];
            Assert.Equal("CABO", v1.Str("entidade"));
            Assert.Equal("I", v1.Str("operacao"));
            Assert.Equal("3xA35P(7,94)-13,8kV 45 m", v1.Str("ativo"));
            Assert.Equal("3xA35P(7,94)-13,8kV 45 m", v1.Str("textoOriginal"));
            Assert.Equal(new object[] { 0.0, 1.0 }, v1.Lista("vinculoEstruturas").ToArray());   // P1·N1 (linha 0) e P2·N1 (linha 1)
            Assert.Equal("MT", v1.Str("_geoTipo"));
            Assert.False(cabos[3].Has("vinculoEstruturas"));                                     // V4 é R: fora do vínculo
            Assert.NotNull(env.Obj("cad"));
        }

        [Fact]
        public void Segunda_estrutura_da_linha_tem_id_linha_mais_k_centesimos()
        {
            Assert.Equal(3.0, ObraArquivo.IdUnidade(3, 0));
            Assert.Equal(3.01, ObraArquivo.IdUnidade(3, 1));
            Obra o = Ler();
            Vao v = o.VaoPorNumero("V4");
            v.Operacao = "I";                                          // vira instalar: agora entra no vínculo
            Construtor.Recalcular(o, Apoio.Pernambuco);
            Assert.Equal("N4", v.VinculoA.Estrutura);
            JObj cabo = ObraArquivo.Dados(o).Lista("cabos").Cast<JObj>().ToList()[3];
            Assert.Equal(1.01, Json.ComoNumero(cabo.Lista("vinculoEstruturas")[0]));
        }

        [Fact]
        public void Ida_e_volta_pela_secao_cad_mantem_numeros_ajustes_e_vinculos()
        {
            Obra o = Ler();
            Construtor.AjustarVao(o, o.VaoPorNumero("V2"), "M", 50, Apoio.Pernambuco);
            Construtor.VincularManual(o, o.VaoPorNumero("V3"), true, 0, Apoio.Pernambuco);   // força V3 → P3·N3 (passa da capacidade)
            string json = ObraArquivo.Gerar(o, Data);
            Obra r = ObraArquivo.Ler(json, Apoio.Pernambuco);
            Assert.Equal(o.Postes.Select(p => p.Numero + p.AtivoLido + p.Handle), r.Postes.Select(p => p.Numero + p.AtivoLido + p.Handle));
            Assert.Equal(o.Vaos.Select(v => v.Numero + v.PosteA + v.PosteB + v.Ativo), r.Vaos.Select(v => v.Numero + v.PosteA + v.PosteB + v.Ativo));
            Vao v2 = r.VaoPorNumero("V2");
            Assert.Equal("M", v2.Operacao);
            Assert.Equal("I", v2.OperacaoLida);
            Assert.Equal("3xA35P(7,94)-13,8kV 50 m", v2.Ativo);
            Assert.True(r.VaoPorNumero("V3").VinculoA.Manual);
            Assert.Equal(o.Elementos.Select(e => e.Estado), r.Elementos.Select(e => e.Estado));
            Assert.Equal(ObraArquivo.Gerar(o, Data), ObraArquivo.Gerar(r, Data));
        }

        [Fact]
        public void Arquivo_ao_lado_do_desenho()
        {
            Assert.Equal(@"C:\obras\Rua.obra.json", ObraArquivo.NomeDoArquivo(@"C:\obras\Rua.dwg"));
            Assert.Equal(@"C:\o.bras\Rua", ObraArquivo.NomeDoArquivo(@"C:\o.bras\Rua").Replace(".obra.json", ""));
        }

        [Fact]
        public void Recusa_arquivo_que_nao_e_do_plugin()
        {
            Assert.Throws<ObraArquivo.ErroArquivo>(() => ObraArquivo.Ler("{\"formato\":\"outro\"}", Apoio.Pernambuco));
            Assert.Throws<ObraArquivo.ErroArquivo>(() => ObraArquivo.Ler("{\"formato\":\"leitor-obra\",\"versao\":1}", Apoio.Pernambuco));
        }

        /// <summary>Exemplo versionado em plugin_nanocad/exemplos — é o arquivo que o teste de contrato da aplicação importa
        /// (tests/test_plugin_nanocad_contrato.py). Para regerar: ATUALIZAR_EXEMPLOS=1 dotnet test.</summary>
        [Fact]
        public void Exemplo_versionado_esta_em_dia()
        {
            string caminho = Path.Combine(Apoio.Plugin, "exemplos", "Projeto_Rua_das_Flores.obra.json");
            Obra o = Ler();
            o.SalvoEm = "31/10/2026 09:00";
            string gerado = ObraArquivo.Gerar(o, Data) + "\n";
            if (Environment.GetEnvironmentVariable("ATUALIZAR_EXEMPLOS") == "1") File.WriteAllText(caminho, gerado);
            Assert.Equal(gerado, File.ReadAllText(caminho).Replace("\r\n", "\n"));
        }
    }

    public class ConstrutorTests
    {
        [Fact]
        public void Construir_do_zero_com_selecao()
        {
            RegrasProjeto rg = Apoio.Pernambuco;
            Obra o = Construtor.NovaObra(@"D:\x\Sem_Padrao.dwg", rg);
            Assert.Equal("Sem_Padrao", o.Nome);
            Poste p1 = Construtor.AdicionarPoste(o, new PosteDesenho { Handle = "B1", X = 0, Y = 0 },
                new List<TextoDesenho> { new TextoDesenho { Handle = "T1", Texto = "9200D, N1", Layer = "01_INSTALANDO", X = 3, Y = 3 } }, rg);
            // projeto sem padrão de layer/cor: as regras ignoram o texto, então vale o texto como está
            Poste p2 = Construtor.AdicionarPoste(o, new PosteDesenho { Handle = "B2", X = 100, Y = 0 },
                new List<TextoDesenho> { new TextoDesenho { Handle = "T2", Texto = "9400D, N1", Layer = "QUALQUER", Cor = "#123456", X = 103, Y = 3 } }, rg);
            Assert.Equal("P1", p1.Numero);
            Assert.Equal("P2", p2.Numero);
            Assert.Equal("POSTE", p2.Entidade);
            Assert.Equal("9400D, N1", p2.AtivoLido);
            Assert.Equal("I", p2.Operacao);
            LinhaDesenho l = new LinhaDesenho { Handle = "L1", Layer = "0", Cor = "#ff0000", Pontos = new List<Ponto> { new Ponto(1, 0), new Ponto(99, 0) } };
            Poste a, b;
            Construtor.SugerirPostes(o, l.Pontos, 10, out a, out b);
            Assert.Same(p1, a);
            Assert.Same(p2, b);
            Vao v = Construtor.AdicionarVao(o, l, new List<TextoDesenho> { new TextoDesenho { Handle = "T3", Texto = "3xA35P(7,94)-13,8kV 45 m", X = 50, Y = 5, Layer = "01_INSTALANDO" } }, a, b, rg);
            Assert.Equal("V1", v.Numero);
            Assert.Equal("N1", v.VinculoA.Estrutura);
            Assert.Equal("N1", v.VinculoB.Estrutura);
            Assert.Equal(Estados.Ok, v.Estado);
            Construtor.Remover(o, p2, rg);
            Assert.Null(v.PosteB);
            Assert.Equal(Estados.Erro, v.Estado);
        }

        [Fact]
        public void Atualizar_mantem_ajustes_e_Reler_descarta()
        {
            RegrasProjeto rg = Apoio.Pernambuco;
            Desenho d = DesenhoExemplo.Montar();
            Obra antiga = LeitorProjeto.Ler(d, rg);
            Construtor.AjustarVao(antiga, antiga.VaoPorNumero("V1"), "M", 47, rg);
            Construtor.AjustarPoste(antiga, antiga.PostePorNumero("P8"), "R", rg);
            Obra nova = LeitorProjeto.Ler(d, rg);
            Construtor.ReaplicarAjustes(nova, antiga, rg);
            Assert.Equal("M", nova.VaoPorNumero("V1").Operacao);
            Assert.Equal(47, nova.VaoPorNumero("V1").Metros);
            Assert.Equal("R", nova.PostePorNumero("P8").Operacao);
            Obra relida = LeitorProjeto.Ler(d, rg);
            Assert.Equal("I", relida.VaoPorNumero("V1").Operacao);
            Construtor.VoltarAoLido(nova, nova.VaoPorNumero("V1"), rg);
            Assert.Equal("I", nova.VaoPorNumero("V1").Operacao);
            Assert.Equal(45, nova.VaoPorNumero("V1").Metros);
        }

        [Fact]
        public void Layer_da_operacao_vem_das_regras()
        {
            Assert.Equal("01_REMOVER", Apoio.Pernambuco.LayerDaOperacao("R"));
            Assert.Equal("01_INSTALANDO", Apoio.Pernambuco.LayerDaOperacao("I"));
            Assert.Equal("01_REAPLICANDO", Apoio.Pernambuco.LayerDaOperacao("M"));
            Assert.Null(Apoio.Pernambuco.LayerDaOperacao("*I"));
        }
    }

    public class BasicosTests
    {
        [Fact]
        public void Json_ida_e_volta()
        {
            string s = "{\"a\":[1,2.5,\"x\\\"y\",true,null,{\"b\":\"ç\\u00e3\"}],\"c\":{}}";
            Assert.Equal("{\"a\":[1,2.5,\"x\\\"y\",true,null,{\"b\":\"çã\"}],\"c\":{}}", Json.Escrever(Json.Ler(s), false));
            Assert.Throws<ErroJson>(() => Json.Ler("{\"a\":}"));
            Assert.Throws<ErroJson>(() => Json.Ler("[1,2"));
        }

        [Fact]
        public void Cores_e_MText()
        {
            Assert.Equal("#ff0000", Cores.AciParaHex(1));
            Assert.Equal("#000000", Cores.AciParaHex(7));
            Assert.Equal("#00ff00", Cores.AciParaHex(3));
            Assert.Equal("#FF0000", Cores.CorDoNomeDaLayer("TXT_ff0000"));
            Assert.Null(Cores.CorDoNomeDaLayer("01_INSTALANDO"));
            List<ParteTexto> p = Cores.DividirMText("\\pxql;9400D\\P{\\C1;N1}\\PN4", "#000000");
            Assert.Equal(new[] { "9400D", "N1", "N4" }, p.Select(x => x.Texto).ToArray());
            Assert.Equal(new[] { "#000000", "#ff0000", "#000000" }, p.Select(x => x.Cor).ToArray());   // { } = cor local
            p = Cores.DividirMText("\\C1;A\\PB", "#000000");
            Assert.Equal(new[] { "#ff0000", "#ff0000" }, p.Select(x => x.Cor).ToArray());         // sem chaves: herdada
            Assert.Equal("ABC", Cores.LimparMText("{\\fArial|b0|i0;ABC}"));
            Assert.Equal("A B", Cores.LimparMText("A\\~B"));
        }

        [Fact]
        public void Poste_compacto()
        {
            Assert.Equal("DT11/300", Anotacoes.PosteOficial("11300D"));
            Assert.Equal("DT9/400", Anotacoes.PosteOficial("9400D"));
            Assert.Equal("RC10/300", Anotacoes.PosteOficial("10300R"));
            Assert.Equal("FV7/150", Anotacoes.PosteOficial("7150FV"));
            Assert.Equal("CV9/300", Anotacoes.PosteOficial("9300F"));
            PosteLido pl = Anotacoes.LerPoste("9400D, N1, 2 N4(P), 0,5 X");
            Assert.Equal("9400D", pl.CodigoPoste);
            Assert.Equal(new[] { "N1", "N4", "X" }, pl.Estruturas.Select(e => e.Nome).ToArray());
            Assert.Equal("2", pl.Estruturas[1].Qtd);
        }
    }
}
