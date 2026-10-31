using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class LeitorTests
    {
        private static Obra Ler() { return LeitorProjeto.Ler(DesenhoExemplo.Montar(), Apoio.Pernambuco); }

        [Fact]
        public void Numeracao_segue_a_rede_como_no_mockup()
        {
            Obra o = Ler();
            Assert.Equal(8, o.Postes.Count);
            Assert.Equal(7, o.Vaos.Count);
            string[] anot = o.Postes.Select(p => p.AtivoLido).ToArray();
            Assert.Equal(new[] { "9200D, N1", "9400D, N1, N4", "11300D, N3", "9400D, N1", "10300R, N1", "9300F, N3", "9200D, N1", "7150FV, B2" }, anot);
            string[] trechos = o.Vaos.Select(v => v.Numero + ":" + v.PosteA + "-" + v.PosteB).ToArray();
            Assert.Equal(new[] { "V1:P1-P2", "V2:P2-P3", "V3:P3-P4", "V4:P2-P5", "V5:P5-P6", "V6:P6-P7", "V7:P6-P8" }, trechos);
            Assert.Equal("3xS04-7,97kV 38 m", o.VaoPorNumero("V3").Ativo);
            Assert.Equal("R", o.VaoPorNumero("V4").Operacao);
            Assert.Equal("M", o.VaoPorNumero("V7").Operacao);
            Assert.Equal("MT", o.VaoPorNumero("V1").Tipo);
            Assert.Equal("BT", o.VaoPorNumero("V5").Tipo);
            Assert.Empty(o.Avulsos);   // "Rua das Flores" (TXT_RUAS) é ignorado pelas regras
        }

        [Fact]
        public void Vinculos_e_conferencia()
        {
            Obra o = Ler();
            Vao v1 = o.VaoPorNumero("V1"), v2 = o.VaoPorNumero("V2"), v3 = o.VaoPorNumero("V3"), v4 = o.VaoPorNumero("V4");
            Assert.Equal("P1#0", Vinculador.Chave(v1.VinculoA.Poste, v1.VinculoA.K));
            Assert.Equal("N1", v1.VinculoB.Estrutura);           // P2: N1 (1ª estrutura)
            Assert.Equal("N1", v2.VinculoA.Estrutura);           // N1 aceita 2 cabos iguais
            Assert.Equal("N3", v2.VinculoB.Estrutura);
            Assert.Null(v3.VinculoA);                             // N3 do P3 já está cheio (1 cabo)
            Assert.Equal("P4", v3.VinculoB.Poste);
            Assert.Equal(Estados.Conferir, v3.Estado);
            Assert.Null(v4.VinculoA);                             // R não vincula
            Assert.Equal(Estados.Ok, v4.Estado);
            Poste p2 = o.PostePorNumero("P2");
            Assert.Equal(new[] { 2, 0 }, p2.CabosPorEstrutura.ToArray());
            Assert.Contains("N4 0/2 cabos", p2.Mensagens);
            Poste p1 = o.PostePorNumero("P1");
            Assert.Equal("TIPOCABO", p1.ParametrosEstruturas[0][0].Key);
            Assert.Equal("P", p1.ParametrosEstruturas[0][0].Value);
            Vao v7 = o.VaoPorNumero("V7");
            Assert.Contains(v7.Mensagens, m => m.Contains("não aceita a operação M"));
            string resumo = Relatorios.Resumo(o);
            Assert.StartsWith("15 elementos", resumo);
        }

        [Fact]
        public void Ver_ativos_com_parametros()
        {
            Obra o = Ler();
            List<LinhaTabela> l = Relatorios.Ativos(o);
            Assert.Equal("3xA35P(7,94)-13,8kV", l.First(x => x.Celulas[0] == "V1").Celulas[1]);
            Assert.Equal("N1(9200D, P)", l.First(x => x.Celulas[0] == "P1 · N1").Celulas[1]);
            LinhaTabela n3 = l.First(x => x.Celulas[0] == "P3 · N3");
            Assert.Equal("N3(11300D, P, 3XA35P_7.94_-13.8KV)", n3.Celulas[1]);
            LinhaTabela n4 = l.First(x => x.Celulas[0] == "P2 · N4");
            Assert.Equal("N4(9400D, ?, ?, ?)", n4.Celulas[1]);
            Assert.Equal(Estados.Conferir, n4.Estado);
            Assert.Equal("DT9/400", l.First(x => x.Celulas[0] == "P2 · poste").Celulas[1]);
            Assert.Equal("FV7/150", l.First(x => x.Celulas[0] == "P8 · poste").Celulas[1]);
            List<LinhaTabela> vinc = Relatorios.Vinculos(o, Apoio.Pernambuco);
            Assert.Equal("N1 (slot 1/2)", vinc[0].Celulas[4]);
            Assert.Equal("N1 (slot 1/2)", vinc[0].Celulas[6]);
            Assert.Equal("N1 (slot 2/2)", vinc[1].Celulas[4]);
            Assert.Contains("não entra no vínculo", vinc[3].Celulas[3]);
        }

        [Fact]
        public void Polilinha_com_varios_postes_vira_varios_vaos_e_ponta_solta_e_erro()
        {
            Desenho d = new Desenho { Arquivo = "x.dwg" };
            d.Postes.Add(new PosteDesenho { Handle = "A", X = 0, Y = 0, Bloco = "Poste_DT" });
            d.Postes.Add(new PosteDesenho { Handle = "B", X = 100, Y = 0, Bloco = "Poste_DT" });
            d.Linhas.Add(new LinhaDesenho { Handle = "L", Layer = "REDE_BT", Cor = "#00ff00", Pontos = new List<Ponto> { new Ponto(0, 0), new Ponto(100, 0), new Ponto(150, 0) } });
            d.Textos.Add(DesenhoExemplo.Texto("3xC16(16) 30 m", 50, -10, "01_INSTALANDO"));
            Obra o = LeitorProjeto.Ler(d, Apoio.Pernambuco);
            Assert.Equal(2, o.Vaos.Count);
            Assert.Equal(new[] { 0, 1 }, o.Vaos.Select(v => v.Trecho).OrderBy(x => x).ToArray());
            Vao solto = o.Vaos.First(v => v.PosteB == null);
            Assert.Equal(Estados.Erro, solto.Estado);
            Assert.Contains("sem poste na ponta", solto.Mensagens);
            Assert.Equal("3xC16(16) 30 m", o.Vaos.First(v => v.PosteB != null).AtivoLido);
        }

        [Fact]
        public void Rede_com_cor_divergente_da_layer_e_descartada()
        {
            Assert.Equal("MT", LeitorProjeto.ClassificarRede("REDE_MT", "#ff0000"));
            Assert.Null(LeitorProjeto.ClassificarRede("REDE_BT", "#ff0000"));
            Assert.Equal("BT", LeitorProjeto.ClassificarRede("REDE_BT", "#ffffff"));
            Assert.Null(LeitorProjeto.ClassificarRede("TXT", "#ff0000"));
        }

        [Fact]
        public void Anotacao_segue_a_linha_guia()
        {
            Desenho d = new Desenho { Arquivo = "x.dwg" };
            d.Postes.Add(new PosteDesenho { Handle = "A", X = 0, Y = 0 });
            d.Postes.Add(new PosteDesenho { Handle = "B", X = 100, Y = 0 });
            d.Linhas.Add(new LinhaDesenho { Handle = "L", Layer = "REDE_MT", Cor = "#ff0000", Pontos = new List<Ponto> { new Ponto(0, 0), new Ponto(100, 0) } });
            // texto longe (100 unidades acima), ligado à rede por uma guia
            d.Textos.Add(new TextoDesenho { Handle = "T", Texto = "3xA35P(7,94)-13,8kV 45 m", X = 50, Y = 100, Altura = 8, Layer = "01_INSTALANDO" });
            d.Linhas.Add(new LinhaDesenho { Handle = "G", Layer = "GUIAS", Cor = "#ffffff", Pontos = new List<Ponto> { new Ponto(52, 98), new Ponto(50, 2) } });
            Obra o = LeitorProjeto.Ler(d, Apoio.Pernambuco);
            Assert.Single(o.Vaos);
            Assert.Equal("T", o.Vaos[0].TextoHandle);
        }

        [Fact]
        public void Poste_so_com_anotacao_e_segunda_anotacao_viram_conferir()
        {
            Desenho d = new Desenho { Arquivo = "x.dwg" };
            d.Postes.Add(new PosteDesenho { Handle = "A", X = 0, Y = 0 });
            d.Textos.Add(new TextoDesenho { Handle = "T1", Texto = "9200D, N1", X = 5, Y = 5, Altura = 8, Layer = "01_INSTALANDO" });
            d.Textos.Add(new TextoDesenho { Handle = "T2", Texto = "11300D, N3", X = -5, Y = 9, Altura = 8, Layer = "01_INSTALANDO" });
            d.Textos.Add(new TextoDesenho { Handle = "T3", Texto = "9400D", X = 500, Y = 500, Altura = 8, Layer = "01_INSTALANDO" });
            Obra o = LeitorProjeto.Ler(d, Apoio.Pernambuco);
            Assert.Equal(2, o.Postes.Count);
            Assert.Equal("T1", o.Postes.First(p => p.Handle == "A").TextoHandle);
            Poste soAnot = o.Postes.First(p => p.Handle == "");
            Assert.Equal(Estados.Conferir, soAnot.Estado);
            Assert.Single(o.Avulsos);
            Assert.Contains(o.Avulsos[0].Mensagens, m => m.Contains("já tem a sua"));
        }
    }
}
