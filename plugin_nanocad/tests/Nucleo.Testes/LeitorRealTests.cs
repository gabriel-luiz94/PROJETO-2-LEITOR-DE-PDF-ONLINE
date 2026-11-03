using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    /// <summary>TASK-011: desenho REAL (NOTA 9101578729, recondutagem em Recife), do ORC_DIAGNOSTICO do usuário.
    /// Rede de MT nova na layer ANOTACOES e na 0, postes existentes como pontos KML (CTW_KMLP_BASE), guias na layer TXT_POSTES em
    /// dois pedaços, várias anotações por vão (instala + retira + BT) e por poste (troca de poste).</summary>
    public class LeitorRealTests
    {
        private static DesenhoBruto Bruto()
        {
            return DesenhoBruto.De((JObj)Json.Ler(File.ReadAllText(Path.Combine(Apoio.Plugin, "referencias", "desenhos", "NOTA_9101578729.diagnostico.json"))));
        }

        private static RegrasProjeto Regras(params string[] mt)
        {
            RegrasProjeto rg = RegrasProjeto.MontarDeObjetos("081", "teste", "teste", Apoio.Pernambuco.Processamento, Apoio.Pernambuco.Classificacao, null);
            rg.Vinculo = Apoio.Pernambuco.Vinculo;
            rg.Leitura = new ConfigLeitura { BlocosPoste = new List<string> { "pg", "KMLP" }, LayersRedeMT = mt.ToList() };
            return rg;
        }

        private static Obra Ler(RegrasProjeto rg) { return LeitorProjeto.Ler(MontadorDesenho.Montar(Bruto(), rg.Leitura), rg); }

        [Fact]
        public void Diagnostico_sugere_as_layers_e_blocos_que_faltam()
        {
            RegrasProjeto rg = Regras("REDE_MT");
            rg.Leitura.BlocosPoste = new List<string> { "pg" };                      // como o usuário rodou
            DesenhoBruto b = Bruto();
            Desenho d = MontadorDesenho.Montar(b, rg.Leitura);
            List<string> l = Diagnostico.Linhas(Diagnostico.Gerar(b, d, LeitorProjeto.Ler(d, rg), rg, DateTime.Now));
            Assert.Contains(l, x => x.Contains("VERMELHAS de poste a poste na layer ANOTACOES"));
            Assert.Contains(l, x => x.Contains("na layer 0") && x.Contains("\"=0\""));
        }

        [Fact]
        public void Le_o_desenho_real_com_a_configuracao_sugerida()
        {
            Obra o = Ler(Regras("REDE_MT", "ANOTACOES", "=0"));
            Assert.Empty(o.Avulsos);                                                     // toda anotação de cabo/poste foi ligada
            Assert.True(o.Elementos.Count(e => e.GuiaUsada != null) >= 60, "anotações ligadas pela linha guia");
            Assert.True(o.Postes.Count(p => p.Placa != null) >= 10, "placas dos postes");
            Assert.Equal("B039104", o.Postes.First(p => p.AtivoLido == "11300D, B4-BT, LD").Placa);
            Assert.True(o.Vaos.Count(v => v.IrmaoDe != null) >= 15, "vãos com mais de uma anotação");
            // anotação com tensão de MT nunca fica num vão de BT (e PB/RAMA nunca num de MT)
            foreach (Vao v in o.Vaos.Where(v => v.TemAnotacao))
            {
                CaboLido c = Anotacoes.LerCabo(v.AtivoLido);
                if (!string.IsNullOrEmpty(c.Tensao) && c.Nivel == "MT") Assert.True(v.Tipo == "MT", v.Numero + " " + v.AtivoLido);
                if ((c.Codigo ?? "").StartsWith("PB")) Assert.True(v.Tipo == "BT", v.Numero + " " + v.AtivoLido);
            }
            // nenhuma linha de "rede" sem poste vira vão; troca de poste = um poste só com os dois blocos
            Assert.True(o.RedesSemPoste > 0);
            Assert.Contains(o.Postes, p => p.OutrosHandles.Count > 0 && p.Bloco.Contains("DESATIVACAO") && p.Bloco.Contains("INSTALACAO"));
            Assert.True(o.Postes.Count(p => p.Handle == "") <= 4, "postes só com anotação");
            // a obra vai para o arquivo e volta igual (irmãos inclusos)
            Obra r = ObraArquivo.Ler(ObraArquivo.Gerar(o, ObraArquivoTests.Data), Regras("REDE_MT", "ANOTACOES", "=0"));
            Assert.Equal(o.Vaos.Count(v => v.IrmaoDe != null), r.Vaos.Count(v => v.IrmaoDe != null));
            Assert.Equal(o.Elementos.Select(e => e.Estado), r.Elementos.Select(e => e.Estado));
        }

        [Fact]
        public void Nome_exato_de_layer_com_igual()
        {
            Assert.True(ConfigLeitura.Contem(new[] { "=0" }, "0"));
            Assert.False(ConfigLeitura.Contem(new[] { "=0" }, "01_INSTALANDO"));
            Assert.True(ConfigLeitura.Contem(new[] { "0" }, "01_INSTALANDO"));
        }
    }
}
