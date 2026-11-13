// TASK-020: o Salvar do plugin grava o projeto de trabalho do OrçaCore ("<desenho>.obra"): envelope + cad + projeto_trabalho.
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class ProjetoArquivoTests
    {
        private static Obra Ler() { return LeitorProjeto.Ler(DesenhoExemplo.Montar(), Apoio.Pernambuco); }

        [Fact]
        public void Caminho_e_formato_do_projeto()
        {
            Assert.Equal(Path.ChangeExtension(@"C:\obras\Rua das Flores.dwg", ".obra"), ProjetoArquivo.Caminho(@"C:\obras\Rua das Flores.dwg"));
            Obra o = Ler();
            o.NumeroObra = "1234567";
            JObj env = ProjetoArquivo.Montar(o, null, ObraArquivoTests.Data);
            Assert.Equal("leitor-obra", env.Str("formato"));
            Assert.Equal("obra", env.Str("tipo"));
            Assert.NotNull(env.Obj("cad"));
            JObj pt = env.Obj("projeto_trabalho");
            Assert.Equal(o.Nome, pt.Str("nome"));
            Assert.Equal("1234567", pt.Str("numero_obra"));
            Assert.NotNull(pt.Obj("sap"));
            Assert.Equal(ProjetoArquivo.AssinaturaLinhas(env.Obj("dados")), env.Obj("cad").Str("assinaturaLinhas"));
        }

        [Fact]
        public void Gravar_de_novo_mantem_sap_e_resultado_do_programa()
        {
            Obra o = Ler();
            o.NumeroObra = "77";
            JObj primeiro = ProjetoArquivo.Montar(o, null, ObraArquivoTests.Data);
            // o OrçaCore preencheu o SAP e calculou o resultado (e regravou o JSON do jeito dele)
            JObj doPrograma = (JObj)Json.Ler(Json.Escrever(primeiro, false));
            doPrograma.Obj("projeto_trabalho").Set("sap", new JObj().Set("bairro", "Centro"));
            doPrograma.Obj("projeto_trabalho").Set("resultado", new JObj().Set("assinatura", "abc").Set("linhas", new List<object> { new JObj().Set("codigo", "1") }));
            Assert.Null(ProjetoArquivo.Conflito(doPrograma));                    // as linhas não mudaram: grava sem perguntar
            JObj segundo = ProjetoArquivo.Montar(o, doPrograma, ObraArquivoTests.Data);
            Assert.Equal("Centro", segundo.Obj("projeto_trabalho").Obj("sap").Str("bairro"));
            Assert.Equal("abc", segundo.Obj("projeto_trabalho").Obj("resultado").Str("assinatura"));
            Assert.Equal("77", segundo.Obj("projeto_trabalho").Str("numero_obra"));
        }

        [Fact]
        public void Linhas_alteradas_no_programa_ou_projeto_so_do_programa_pedem_confirmacao()
        {
            Obra o = Ler();
            JObj env = (JObj)Json.Ler(Json.Escrever(ProjetoArquivo.Montar(o, null, ObraArquivoTests.Data), false));
            JObj linha = env.Obj("dados").Lista("cabos").OfType<JObj>().First();
            linha.Set("ativo", linha.Str("ativo") + " X");                       // editado no Resumo do OrçaCore
            Assert.Contains("alteradas no OrçaCore", ProjetoArquivo.Conflito(env));
            JObj soPrograma = (JObj)Json.Ler(Json.Escrever(env, false));
            soPrograma.Set("cad", null);
            soPrograma = (JObj)Json.Ler(Json.Escrever(soPrograma, false).Replace(",\"cad\":null", ""));
            Assert.Contains("criado no programa", ProjetoArquivo.Conflito(soPrograma));
            Assert.Null(ProjetoArquivo.Conflito(null));
        }

        [Fact]
        public void Continuar_le_nome_e_numero_que_o_programa_guardou()
        {
            Obra o = Ler();
            o.NumeroObra = "1";
            JObj env = ProjetoArquivo.Montar(o, null, ObraArquivoTests.Data);
            env.Obj("projeto_trabalho").Set("nome", "Nome do OrçaCore").Set("numero_obra", "999");
            Obra lida = ObraArquivo.Ler(Json.Escrever(env, false), Apoio.Pernambuco);
            Assert.Equal("Nome do OrçaCore", lida.Nome);
            Assert.Equal("999", lida.NumeroObra);
        }

        [Fact]
        public void Gravacao_atomica_substitui_o_arquivo()
        {
            string pasta = Path.Combine(Path.GetTempPath(), "orc_proj_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(pasta);
            try
            {
                string arq = Path.Combine(pasta, "Rua.obra");
                ProjetoArquivo.GravarAtomico(arq, "{\"a\":1}");
                ProjetoArquivo.GravarAtomico(arq, "{\"a\":2}");
                Assert.Equal(2, (int)ProjetoArquivo.LerArquivo(arq).Num("a").Value);
                Assert.False(File.Exists(arq + ".tmp"));
            }
            finally { Directory.Delete(pasta, true); }
        }

        [Fact]
        public void Exemplo_de_projeto_versionado_esta_em_dia()
        {
            // o mesmo arquivo é aberto pelo OrçaCore nos testes da aplicação (tests/test_plugin_projeto_163.py)
            string caminho = Path.Combine(Apoio.Plugin, "exemplos", "Projeto_Rua_das_Flores.obra");
            Obra o = Ler();
            o.SalvoEm = "31/10/2026 09:00";
            o.NumeroObra = "1234567";
            string gerado = Json.Escrever(ProjetoArquivo.Montar(o, null, ObraArquivoTests.Data), true) + "\n";
            if (Environment.GetEnvironmentVariable("ATUALIZAR_EXEMPLOS") == "1") File.WriteAllText(caminho, gerado);
            Assert.Equal(gerado, File.ReadAllText(caminho).Replace("\r\n", "\n"));
        }
    }
}
