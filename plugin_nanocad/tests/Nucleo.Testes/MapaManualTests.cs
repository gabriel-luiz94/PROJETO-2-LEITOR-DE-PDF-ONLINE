using System;
using System.Collections.Generic;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    /// <summary>TASK-038: montagem manual no formato do mapa do OrçaCore (ponto = «@nome - comentário» + campos com operação; cabo = linha/trecho + campos).</summary>
    public class MapaManualTests
    {
        private static TextoDesenho T(string texto, double x, double y, string layer = "01_INSTALANDO", int parte = 0)
        {
            return new TextoDesenho { Handle = "H" + texto.GetHashCode().ToString("X"), Texto = texto, X = x, Y = y, Layer = layer, Altura = 1.2, Parte = parte, Tipo = "MTEXT" };
        }

        [Fact]
        public void Com_regras_o_ativo_e_a_operacao_vem_da_regra_e_texto_sem_regra_vira_F()
        {
            var textos = new List<TextoDesenho> { T("N3", 10, 8, "01_REMOVER"), T("9200D, N1", 10, 10), T("TEXTO QUALQUER 123", 10, 6, "LAYER_ESTRANHA") };
            List<CampoMapa> c = MapaManual.CamposDosTextos(textos, Apoio.Pernambuco, "I", false);
            Assert.Equal(new[] { 1, 2, 3 }, c.Select(x => x.L).ToArray());                         // ordem de leitura: de cima para baixo
            Assert.Equal("9200D, N1", c[0].Texto); Assert.Equal("I", c[0].Operacao); Assert.Equal("regra", c[0].Origem);
            Assert.Equal("N3", c[1].Texto); Assert.Equal("R", c[1].Operacao);
            Assert.Equal("F", c[2].Operacao); Assert.Equal("sem regra", c[2].Origem); Assert.Equal("TEXTO QUALQUER 123", c[2].Texto);
        }

        [Fact]
        public void Sem_regras_o_texto_vai_como_esta_com_a_operacao_padrao()
        {
            List<CampoMapa> c = MapaManual.CamposDosTextos(new[] { T("cej2, flabit", 0, 5), T("it-r", 0, 3) }, null, "*m", false);
            Assert.Equal(new[] { "CEJ2, FLABIT", "IT-R" }, c.Select(x => x.Texto).ToArray());
            Assert.All(c, x => Assert.Equal("*M", x.Operacao));
            Assert.Equal("I", MapaManual.CamposDosTextos(new[] { T("N3", 0, 0) }, null, "X", false)[0].Operacao);
        }

        [Fact]
        public void Cabo_tira_a_metragem_do_campo_e_le_o_comprimento()
        {
            List<CampoMapa> c = MapaManual.CamposDosTextos(new[] { T("3xA35P(7,94)-13,8kV 45 m", 0, 0) }, null, "I", true);
            Assert.Equal("3XA35P(7,94)-13,8KV", c[0].Texto);
            Assert.Equal(45.0, MapaManual.MetrosDoTexto("3xA35P(7,94)-13,8kV 45 m"));
            Assert.Equal(38.4, MapaManual.MetrosDoTexto("PB35 38,4m"));
            Assert.Null(MapaManual.MetrosDoTexto("3xA35P(7,94)-13,8kV"));
            Assert.Equal("38,4", MapaManual.FormatarMetros(38.42)); Assert.Equal("40", MapaManual.FormatarMetros(40.01));
        }

        [Fact]
        public void Trecho_da_polilinha_entre_dois_pontos()
        {
            var linha = new List<Ponto> { new Ponto(0, 0), new Ponto(40, 0), new Ponto(40, 30) };
            Assert.Equal(70, MapaManual.Comprimento(linha), 6);
            List<Ponto> t = MapaManual.Trecho(linha, new Ponto(40.5, 20), new Ponto(10, 1));         // fora da ordem e fora da linha: projeta
            Assert.Equal(new[] { 10.0, 40.0, 40.0 }, t.Select(p => Math.Round(p.X, 6)).ToArray());
            Assert.Equal(new[] { 0.0, 0.0, 20.0 }, t.Select(p => Math.Round(p.Y, 6)).ToArray());
            Assert.Equal(50, MapaManual.Comprimento(t), 6);
            Assert.Equal("P3", MapaManual.ProximoNumero(new[] { "P1", "P2", "X9" }, "P"));
        }

        [Fact]
        public void Linhas_do_ponto_e_do_cabo_no_formato_do_mapa()
        {
            var ponto = new PontoMapa { Numero = "P12", Comentario = "Rua das  Flores", Handle = "AB", X = 287431.224, Y = 9112004.1,
                Campos = { new CampoMapa { L = 1, Texto = "12600D", Operacao = "F" }, new CampoMapa { L = 2, Texto = "N3", Operacao = "I" }, new CampoMapa { L = 2, Texto = " ", Operacao = "R" } } };
            List<JObj> l = MapaManual.LinhasDoPonto(ponto).Cast<JObj>().ToList();
            Assert.Equal(3, l.Count);
            Assert.Equal("@P12 - Rua das Flores", l[0].Str("ativo")); Assert.True(l[0].Bool("grupoDecl", false)); Assert.Equal("I", l[0].Str("operacao"));
            Assert.Equal(l[0].Str("gid"), l[1].Str("grupoDe"));
            Assert.Equal("F", l[1].Str("operacao")); Assert.True(l[1].Bool("congelada", false)); Assert.True(l[1].Bool("_congeladaPorF", false));
            Assert.Equal("POSTE", l[1].Str("entidade")); Assert.Equal("ESTRUTURA", l[2].Str("entidade"));
            Assert.Equal(new object[] { 287431.22, 9112004.1 }, l[2].Lista("_poste").ToArray());
            ponto.Handle = "";
            Assert.True(MapaManual.LinhasDoPonto(ponto).Cast<JObj>().Skip(1).All(x => x.Has("_posteAprox") && !x.Has("_poste")));   // sem símbolo: aproximada

            var cabo = new CaboMapa { Numero = "C1", Pontos = { new Ponto(0, 0), new Ponto(38.4, 0) }, Tipo = "MT", Medido = 38.4,
                Campos = { new CampoMapa { Texto = "3XA35P(7,94)-13,8KV", Operacao = "I" }, new CampoMapa { Texto = "3XC25-13,8KV 30 m", Operacao = "R" } } };
            List<JObj> c = MapaManual.LinhasDoCabo(cabo).Cast<JObj>().ToList();
            Assert.Equal(new[] { "3XA35P(7,94)-13,8KV 38,4 m", "3XC25-13,8KV 38,4 m" }, c.Select(x => x.Str("ativo")).ToArray());
            Assert.Equal("MT", c[0].Str("_geoTipo")); Assert.Equal(2, c[0].Lista("_geo").Count); Assert.Equal("R", c[1].Str("operacao"));
            cabo.Metros = 40;
            Assert.Equal("3XA35P(7,94)-13,8KV 40 m", MapaManual.LinhasDoCabo(cabo).Cast<JObj>().First().Str("ativo"));
        }

        [Fact]
        public void Pontos_e_cabos_do_mapa_vao_para_a_obra_e_voltam_da_secao_cad()
        {
            Obra o = Construtor.NovaObra(@"C:\x\Mapa.dwg", Apoio.Pernambuco);
            o.SemRegras = true;
            o.PontosMapa.Add(new PontoMapa { Numero = "P1", Nome = "P1", Handle = "AA", X = 1, Y = 2, Enviado = true, Textos = { "T1" },
                Campos = { new CampoMapa { L = 1, Texto = "N3", Operacao = "*R", Origem = "manual" } } });
            o.CabosMapa.Add(new CaboMapa { Numero = "C1", LinhaHandle = "LL", Trecho = true, Tipo = "BT", Medido = 12.5, Metros = 13,
                Pontos = { new Ponto(0, 0), new Ponto(12.5, 0) }, Campos = { new CampoMapa { Texto = "PB35", Operacao = "M" } } });
            JObj dados = ObraArquivo.Dados(o);
            Assert.Equal(2, dados.Lista("outros").Count);
            Assert.Equal("PB35 13 m", ((JObj)dados.Lista("cabos")[0]).Str("ativo"));
            string json = Json.Escrever(ObraArquivo.Envelope(o, ObraArquivoTests.Data), true);
            Obra lida = ObraArquivo.Ler(json, Apoio.Pernambuco);
            Assert.True(lida.SemRegras == true);
            PontoMapa p = Assert.Single(lida.PontosMapa);
            Assert.Equal("*R", p.Campos[0].Operacao); Assert.True(p.Enviado); Assert.Equal("AA", p.Handle); Assert.Equal(new[] { "T1" }, p.Textos.ToArray());
            CaboMapa c = Assert.Single(lida.CabosMapa);
            Assert.True(c.Trecho); Assert.Equal(13.0, c.Metros); Assert.Equal(2, c.Pontos.Count); Assert.Equal("BT", c.Tipo);
        }

        /// <summary>Montagem manual de exemplo (com as regras de Pernambuco): o contrato com a aplicação é conferido em tests/test_plugin_mapa_manual_038.py.</summary>
        public static Obra ObraDeExemplo()
        {
            RegrasProjeto rg = Apoio.Pernambuco;
            Obra o = Construtor.NovaObra(@"C:\Projetos\Montagem_Manual.dwg", rg);
            o.SemRegras = false;
            o.SalvoEm = "23/11/2026 09:00";
            o.PontosMapa.Add(new PontoMapa { Numero = "P1", Comentario = "Rua das Flores, 120", Handle = "1A2", Bloco = "POSTE", X = 1000, Y = 2000, Textos = { "T1", "T2" },
                Campos = MapaManual.CamposDosTextos(new[] { T("9200D, N1", 1003, 2004), T("TEXTO QUALQUER 123", 1003, 2001, "LAYER_ESTRANHA") }, rg, "I", false) });
            o.PontosMapa.Add(new PontoMapa { Numero = "P2", X = 1038.4, Y = 2000, Textos = { "T3" },          // sem símbolo: posição aproximada (a da anotação)
                Campos = MapaManual.CamposDosTextos(new[] { T("N3", 1038.4, 2000, "01_REMOVER") }, rg, "I", false) });
            o.PontosMapa.Add(new PontoMapa { Numero = "P3", Nome = "PX", Comentario = "poste existente", Handle = "1A9", X = 1080, Y = 2000, Textos = { "T5" },   // ajustado à mão na janela
                Campos = { new CampoMapa { L = 1, Texto = "12600D", Operacao = "F", Origem = "manual" }, new CampoMapa { L = 2, Texto = "N3", Operacao = "I", Origem = "manual" } } });
            o.CabosMapa.Add(new CaboMapa { Numero = "C1", LinhaHandle = "2B0", Tipo = "MT", Pontos = { new Ponto(1000, 2000), new Ponto(1038.4, 2000) }, Medido = 38.4, Metros = 40, Textos = { "T4" },
                Campos = MapaManual.CamposDosTextos(new[] { T("3xA35P(7,94)-13,8kV 40 m", 1019, 2002) }, rg, "I", true) });
            return o;
        }

        [Fact]
        public void Exemplo_da_montagem_manual_esta_em_dia()
        {
            Obra o = ObraDeExemplo();
            Assert.Equal(new[] { "9200D, N1:I:regra:1", "TEXTO QUALQUER 123:F:sem regra:2" }, o.PontosMapa[0].Campos.Select(c => c.Texto + ":" + c.Operacao + ":" + c.Origem + ":" + c.L).ToArray());       // texto fora das regras → F
            Assert.Equal("R", o.PontosMapa[1].Campos[0].Operacao);                                                  // N3 na layer de retirada
            string caminho = System.IO.Path.Combine(Apoio.Plugin, "exemplos", "Montagem_Manual_Mapa.obra.json");
            string gerado = ObraArquivo.Gerar(o, ObraArquivoTests.Data) + "\n";
            if (Environment.GetEnvironmentVariable("ATUALIZAR_EXEMPLOS") == "1") System.IO.File.WriteAllText(caminho, gerado);
            Assert.Equal(gerado, System.IO.File.ReadAllText(caminho).Replace("\r\n", "\n"));
        }
    
        [Fact]
        public void Pontos_e_cabos_do_mapa_aparecem_nas_abas_do_painel()
        {
            Obra o = ObraDeExemplo();
            o.PontosMapa[1].Enviado = true;
            List<LinhaTabela> pontos = MapaManual.LinhasDoPainelPontos(o);
            Assert.Equal(new[] { "M:P1", "M:P2", "M:P3" }, pontos.Select(x => x.Chave).ToArray());
            Assert.Equal(7, pontos[0].Celulas.Length);
            Assert.Equal("I/F", pontos[0].Celulas[2]);
            Assert.Equal("@P1 - Rua das Flores, 120: I 9200D, N1 · F TEXTO QUALQUER 123", pontos[0].Celulas[3]);
            Assert.Equal("sem símbolo", pontos[1].Celulas[4]); Assert.Equal("Enviada", pontos[1].Celulas[6]); Assert.Equal("op:R", pontos[1].Icones[2]);
            Assert.Equal("@PX - poste existente: F 12600D · I N3", pontos[2].Celulas[3]);
            LinhaTabela cabo = Assert.Single(MapaManual.LinhasDoPainelCabos(o));
            Assert.Equal(8, cabo.Celulas.Length);
            Assert.Equal(new[] { "", "C1", "MT", "I", "I 3xA35P(7,94)-13,8kV", "40", "no mapa", "Mapa" }, cabo.Celulas);
            Assert.Equal(new[] { new Ponto(1000, 2000).X, 1038.4 }, MapaManual.PontosParaZoom(o, "M:C1").Select(p => p.X).ToArray());
            Assert.Equal(1080, MapaManual.PontosParaZoom(o, "P3")[0].X);
        }

        private static TextoDesenho P(string handle, int parte, string texto, double y, string layer)
        {
            return new TextoDesenho { Handle = handle, Parte = parte, Texto = texto, X = 288391.2, Y = y, Layer = layer, Altura = 1.2, Tipo = "MTEXT" };
        }

        [Fact]
        public void Paragrafo_de_continuacao_do_mesmo_mtext_entra_no_campo_de_cima()
        {
            // NOTA_9101578729: «12600D, CEJ2, FLABIT, IT-R\PFLABIM, IT-R» — o 2º parágrafo virava um campo F e, trocado para I, dividia o ativo
            var textos = new[] { P("5BFC0", 0, "12600D, CEJ2, FLABIT, IT-R", 9113465.3, "POSTE_INSTALANDO_LV"), P("5BFC0", 1, "FLABIM, IT-R", 9113465.3, "POSTE_INSTALANDO_LV"),
                                 P("5BFC2", 0, "11300D, N3-2, IT-2, CFU", 9113459.1, "POSTE_REMOVENDO_LV") };
            List<CampoMapa> c = MapaManual.CamposDosTextos(textos, Apoio.Pernambuco, "I", false);
            Assert.Equal(new[] { "1 *I 12600D, CEJ2, FLABIT, IT-R, FLABIM, IT-R", "2 *R 11300D, N3-2, IT-2, CFU" }, c.Select(x => x.L + " " + x.Operacao + " " + x.Texto).ToArray());
            List<CampoMapa> sem = MapaManual.CamposDosTextos(textos, null, "I", false);
            Assert.Equal(new[] { "12600D, CEJ2, FLABIT, IT-R, FLABIM, IT-R", "11300D, N3-2, IT-2, CFU" }, sem.Select(x => x.Texto).ToArray());
        }

        [Fact]
        public void Metragem_ja_juntada_ao_cabo_nao_vira_campo_proprio()
        {
            var textos = new[] { new TextoDesenho { Handle = "5C03E", Texto = "3xA35P(7,94)", X = 288398.9, Y = 9113450.6, Layer = "POSTE_INSTALANDO_LV", Altura = 1.2, Tipo = "MTEXT" },
                                 new TextoDesenho { Handle = "5C040", Texto = "32 m", X = 288398.9, Y = 9113445.7, Layer = "0", Altura = 1.2, Tipo = "MTEXT" } };
            CampoMapa c = Assert.Single(MapaManual.CamposDosTextos(textos, Apoio.Pernambuco, "I", true));
            Assert.Equal("3xA35P(7,94)", c.Texto); Assert.Equal("*I", c.Operacao);
            Assert.Single(MapaManual.CamposDosTextos(textos, null, "I", true));                  // sem regras: a metragem também não vira campo
        }

        [Fact]
        public void Textos_de_mtexts_diferentes_continuam_em_campos_separados()
        {
            var textos = new[] { P("A1", 0, "12600D, CEJ2", 10, "POSTE_INSTALANDO_LV"), P("A2", 0, "LD", 8, "POSTE_INSTALANDO_LV") };
            Assert.Equal(new[] { "12600D, CEJ2", "LD" }, MapaManual.CamposDosTextos(textos, null, "I", false).Select(x => x.Texto).ToArray());
        }
}
}
