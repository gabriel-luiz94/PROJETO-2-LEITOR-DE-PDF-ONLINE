using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    /// <summary>TASK-004/005/006: o desenho BRUTO (o mesmo formato do ORC_DIAGNOSTICO) lido de um DXF de teste
    /// (tests/oraculo/gerar_dxf_teste.py → dxf_para_bruto.py) e comparado com o Leitor da APLICAÇÃO no mesmo DXF.
    /// Desenhos reais mandados pelo usuário (referencias/desenhos/*.diagnostico.json ou *.bruto.json) entram no teste de leitura.</summary>
    public class LeitorDxfTests
    {
        private static string Oraculo(string a) { return Path.Combine(Apoio.Plugin, "tests", "oraculo", a); }

        private static DesenhoBruto Bruto() { return DesenhoBruto.De((JObj)Json.Ler(File.ReadAllText(Oraculo("dxf_teste_bruto.json")))); }

        [Fact]
        public void Textos_e_regras_iguais_ao_leitor_da_aplicacao_no_mesmo_dxf()
        {
            Desenho d = MontadorDesenho.Montar(Bruto(), Apoio.Pernambuco.Leitura);
            // a aplicação não lê MLEADER do DXF; fora isso, os MESMOS textos, cores e resultado das regras, na MESMA ordem
            List<TextoLido> meus = LeitorProjeto.ClassificarTextos(d.Textos.Where(t => t.Tipo != "MLEADER"), Apoio.Pernambuco);
            List<JObj> app = ((List<object>)Json.Ler(File.ReadAllText(Oraculo("dxf_teste_esperado_app.json")))).Cast<JObj>().ToList();
            Assert.Equal(app.Count, meus.Count);
            for (int i = 0; i < app.Count; i++)
            {
                string onde = "#" + i + " '" + app[i].Str("texto") + "'";
                Assert.True(app[i].Str("texto") == meus[i].T.Texto, onde + " texto: '" + meus[i].T.Texto + "'");
                Assert.True(app[i].Str("cor") == meus[i].T.Cor, onde + " cor: " + meus[i].T.Cor + " ≠ " + app[i].Str("cor"));
                Assert.True(app[i].Str("layer") == meus[i].T.Layer, onde + " layer");
                Assert.True(app[i].Str("entidade") == meus[i].R.Entidade, onde + " entidade: " + meus[i].R.Entidade);
                Assert.True(app[i].Str("operacao") == meus[i].R.Operacao, onde + " operação: " + meus[i].R.Operacao);
                Assert.True(app[i].Str("ativo") == meus[i].R.Ativo, onde + " ativo: '" + meus[i].R.Ativo + "'");
            }
        }

        [Fact]
        public void Le_a_obra_inteira_com_LEADER_MLEADER_e_ignorados()
        {
            RegrasProjeto rg = Apoio.Pernambuco;
            Obra o = LeitorProjeto.Ler(MontadorDesenho.Montar(Bruto(), rg.Leitura), rg);
            Assert.Equal(8, o.Postes.Count);
            Assert.Equal(7, o.Vaos.Count);
            Assert.All(o.Postes, p => Assert.True(p.TemAnotacao, p.Numero + " sem anotação"));
            Assert.All(o.Vaos, v => Assert.True(v.TemAnotacao, v.Numero + " sem anotação"));
            Assert.Equal("9400D, N1, N4", o.PostePorNumero("P2").AtivoLido);                    // MTEXT de 3 parágrafos
            Poste p7 = o.PostePorNumero("P7");
            Assert.Equal("9200D, N1", p7.AtivoLido);                                              // anotação longe, pelo LEADER
            Assert.NotNull(p7.GuiaUsada);
            Vao v2 = o.VaoPorNumero("V2");
            Assert.Equal("3xA35P(7,94)-13,8kV 45 m", v2.AtivoLido);                               // MLEADER
            Assert.NotNull(v2.GuiaUsada);
            Assert.Equal("3xS04-7,97kV 38 m", o.VaoPorNumero("V3").AtivoLido);                    // comprimento na linha de baixo
            Assert.Empty(o.Avulsos);
            string[] ignorados = o.Ignorados.Select(t => t.Texto).ToArray();
            Assert.Contains("Rua das Flores", ignorados);
            Assert.Contains("12.50", ignorados);
            Assert.Contains("3xC16(16) 99 m", ignorados);                                          // azul = operação 0 (regra de motor)
            Assert.Contains("38 m", ignorados);                                                   // já usado pelo cabo
            Assert.Contains(o.Ignorados, t => t.Motivo.Contains("operação 0"));
        }

        [Fact]
        public void Configuracao_de_leitura_muda_onde_estao_rede_e_postes()
        {
            DesenhoBruto b = Bruto();
            // renomeia as layers da rede e o bloco, como num desenho de outro padrão
            foreach (EntidadeBruta e in b.Entidades)
            {
                if (e.Layer == "REDE_MT") e.Layer = "MT PROJETADA";
                if (e.Layer == "REDE_BT") e.Layer = "BT PROJETADA";
                if (e.Tipo == "INSERT") e.Bloco = "SIMB_PT";
            }
            RegrasProjeto rg = Apoio.Pernambuco;
            Obra semConfig = LeitorProjeto.Ler(MontadorDesenho.Montar(b, rg.Leitura), rg);
            Assert.Empty(semConfig.Vaos);
            Assert.All(semConfig.Postes, p => Assert.Equal("", p.Handle));                          // sem bloco: postes "só anotação"
            RegrasProjeto outra = RegrasProjeto.MontarDeObjetos("081", "teste", "teste", rg.Processamento, rg.Classificacao, null);
            outra.Vinculo = rg.Vinculo;
            outra.Leitura = ConfigLeitura.De(new JObj()
                .Set("layers_rede_mt", new List<object> { "MT PROJ" }).Set("layers_rede_bt", new List<object> { "BT PROJ" })
                .Set("blocos_poste", new List<object> { "SIMB_PT" }));
            Obra comConfig = LeitorProjeto.Ler(MontadorDesenho.Montar(b, outra.Leitura), outra);
            Assert.Equal(7, comConfig.Vaos.Count);
            Assert.Equal(8, comConfig.Postes.Count(p => p.Handle != ""));
            // e o diagnóstico avisa o que faltou na leitura sem configuração
            JObj diag = Diagnostico.Gerar(b, MontadorDesenho.Montar(b, rg.Leitura), semConfig, rg, new DateTime(2026, 11, 2));
            List<string> linhas = Diagnostico.Linhas(diag);
            Assert.Contains(linhas, l => l.Contains("nenhuma linha de rede"));
            Assert.Contains(linhas, l => l.Contains("nenhum poste achado") && l.Contains("SIMB_PT"));
            Assert.Equal("plugin-desenho-bruto", diag.Obj("bruto").Str("formato"));
            Assert.Equal(b.Entidades.Count, DesenhoBruto.De(diag).Entidades.Count);                 // o diagnóstico serve de entrada
        }

        [Fact]
        public void Contorno_do_texto_e_guias_proprias()
        {
            TextoDesenho t = new TextoDesenho { X = 0, Y = 0, Caixa = new double[] { 0, 0, 100, 10 } };
            Assert.Equal(0, t.Distancia(new Ponto(50, 5)));
            Assert.Equal(5, t.Distancia(new Ponto(105, 5)), 6);
            Assert.Equal(new Ponto(50, 5).ToString(), t.Centro.ToString());
            Assert.Equal(new Ponto(105, 0).Dist(new Ponto(0, 0)), new TextoDesenho { X = 0, Y = 0 }.Distancia(new Ponto(105, 0)));
        }

        /// <summary>Desenhos reais em referencias/desenhos (diagnóstico do plugin ou bruto do conversor): a leitura não pode quebrar.</summary>
        [Fact]
        public void Desenhos_reais_da_pasta_referencias_sao_lidos()
        {
            string pasta = Path.Combine(Apoio.Plugin, "referencias", "desenhos");
            if (!Directory.Exists(pasta)) return;
            foreach (string f in Directory.GetFiles(pasta, "*.json"))
            {
                DesenhoBruto b = DesenhoBruto.De(Json.Ler(File.ReadAllText(f)) as JObj);
                if (b.Entidades.Count == 0) continue;
                Obra o = LeitorProjeto.Ler(MontadorDesenho.Montar(b, Apoio.Pernambuco.Leitura), Apoio.Pernambuco);
                Assert.True(o.Elementos.Count() + o.Ignorados.Count > 0, Path.GetFileName(f));
            }
        }
    }
}
