using System;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class LeituraDxfAplicacaoTests
    {
        [Fact]
        public void Preserva_os_ativos_e_a_ordem_da_aplicacao_com_ou_sem_geometria()
        {
            JObj resposta = (JObj)Json.Ler(@"{
              ""linhas"": [
                {""texto"":""CE2"",""entidade"":""AVULSO"",""operacao"":""R"",""ativo"":""CE2"",""_x"":100,""_y"":10},
                {""texto"":""A35P"",""entidade"":""CABO"",""operacao"":""I"",""ativo"":""A35P 25 m"",""_x"":100,""_y"":20},
                {""texto"":""12600D, N3"",""entidade"":""POSTE"",""operacao"":""I"",""ativo"":""12600D, N3"",""_x"":0,""_y"":0},
                {""texto"":""A35P"",""entidade"":""CABO"",""operacao"":""*R"",""ativo"":""A35P 30 m"",""_x"":10,""_y"":0},
                {""texto"":""B093116"",""entidade"":""0"",""operacao"":""I"",""ativo"":""B093116""},
                {""texto"":""N3"",""entidade"":""ESTRUTURA"",""operacao"":""R"",""ativo"":""N3"",""_x"":400,""_y"":20}
              ]}");
            var desenho = new Desenho();
            desenho.Linhas.Add(new LinhaDesenho { Layer = "REDE_MT", Cor = "#ff0000", Pontos = { new Ponto(0, 0), new Ponto(20, 0) } });
            desenho.Postes.Add(new PosteDesenho { X = 0, Y = 0, Handle = "P" });
            desenho.Textos.Add(new TextoDesenho { Texto = "CE2 27 m", Tipo = "TEXT" });  // extração nativa propositalmente divergente
            var regras = Apoio.Pernambuco;
            var obra = LeitorProjeto.Ler(desenho, regras, null, LeitorProjeto.TextosDoDxf(resposta, desenho));
            obra.FonteDxf = "obra.dxf";
            obra.HashLeitor = "regras";
            var reaberta = ObraArquivo.Ler(ObraArquivo.Gerar(obra, DateTime.UtcNow), regras);
            JObj dados = ObraArquivo.Dados(reaberta);
            Assert.Equal(new[] { "A35P 25 m", "A35P 30 m" }, dados.Lista("cabos").Cast<JObj>().Select(l => l.Str("ativo")));
            Assert.Equal(new[] { "I", "*R" }, dados.Lista("cabos").Cast<JObj>().Select(l => l.Str("operacao")));
            Assert.Equal(new[] { "CE2", "12600D, N3", "N3" }, dados.Lista("outros").Cast<JObj>().Select(l => l.Str("ativo")));
            Assert.Equal("obra.dxf", reaberta.FonteDxf);
            Assert.Equal("regras", reaberta.HashLeitor);
        }
    }
}
