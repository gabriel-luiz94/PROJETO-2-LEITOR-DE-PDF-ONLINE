using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    /// <summary>O porte em C# precisa dar o MESMO resultado do código da aplicação (oráculos em tests/oraculo).</summary>
    public class ParidadeTests
    {
        private static string Oraculo(string arquivo) { return Path.Combine(Apoio.Plugin, "tests", "oraculo", arquivo); }

        private static string Rodar(string exe, string args)
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo(exe, args);
                psi.RedirectStandardOutput = true;
                psi.UseShellExecute = false;
                psi.WorkingDirectory = Apoio.RaizRepo;
                using (Process p = Process.Start(psi))
                {
                    string saida = p.StandardOutput.ReadToEnd();
                    p.WaitForExit();
                    return p.ExitCode == 0 ? saida : null;
                }
            }
            catch (System.ComponentModel.Win32Exception) { return null; }   // node/python ausente: vale o arquivo versionado
        }

        [Fact]
        public void Esperados_versionados_batem_com_o_codigo_atual_da_aplicacao()
        {
            // se o JS/Python da aplicação mudar, o esperado versionado fica velho e este teste avisa (regerar e conferir o porte)
            string js = Rodar("node", "\"" + Oraculo("oraculo_motor.js") + "\"");
            if (js != null) Assert.Equal(Json.Escrever(Json.Ler(File.ReadAllText(Oraculo("esperado_motor.json"))), false), Json.Escrever(Json.Ler(js), false));
            string py = Rodar("python3", "\"" + Oraculo("oraculo_anotacoes.py") + "\"");
            if (py != null) Assert.Equal(Json.Escrever(Json.Ler(File.ReadAllText(Oraculo("esperado_anotacoes.json"))), false), Json.Escrever(Json.Ler(py), false));
        }

        [Fact]
        public void Motor_do_leitor_igual_ao_JS()
        {
            List<object> casos = (List<object>)Json.Ler(File.ReadAllText(Oraculo("casos_motor.json")));
            List<object> esperado = (List<object>)Json.Ler(File.ReadAllText(Oraculo("esperado_motor.json")));
            int conferidos = 0;
            for (int c = 0; c < casos.Count; c++)
            {
                JObj caso = (JObj)casos[c];
                RegrasProjeto regras = caso.Str("regras") == "081" ? Apoio.Pernambuco : Apoio.Padrao;
                List<ItemLeitor> itens = caso.Lista("itens").Cast<JObj>().Select(o =>
                {
                    ItemLeitor i = new ItemLeitor();
                    i.Texto = o.Str("texto"); i.Cor = o.Str("cor"); i.Layer = o.Str("layer"); i.Pagina = (int)(o.Num("pagina") ?? 1);
                    i.X = o.Num("_x"); i.Y = o.Num("_y");
                    return i;
                }).ToList();
                List<object> res = ((JObj)esperado[c]).Lista("resultados");
                for (int k = 0; k < itens.Count; k++)
                {
                    ResultadoLeitor r = MotorLeitor.ProcessarEClassificar(itens, k, regras.Processamento, regras.Classificacao);
                    JObj e = (JObj)res[k];
                    string onde = caso.Str("nome") + " #" + k + " '" + itens[k].Texto + "'";
                    Assert.True(e.Str("operacao") == r.Operacao, onde + ": operação " + r.Operacao + " ≠ " + e.Str("operacao"));
                    Assert.True(e.Str("ativo") == r.Ativo, onde + ": ativo '" + r.Ativo + "' ≠ '" + e.Str("ativo") + "'");
                    Assert.True(e.Str("entidade") == r.Entidade, onde + ": entidade " + r.Entidade + " ≠ " + e.Str("entidade"));
                    conferidos++;
                }
            }
            Assert.True(conferidos >= 30);
        }

        [Fact]
        public void Anotacao_de_cabo_igual_ao_Python()
        {
            foreach (JObj e in ((List<object>)Json.Ler(File.ReadAllText(Oraculo("esperado_anotacoes.json")))).Cast<JObj>())
            {
                string t = e.Str("texto");
                CaboLido c = Anotacoes.LerCabo(t);
                if (e["codigo"] == null) Assert.True(c.Erro != null, t);
                else
                {
                    Assert.Equal(e.Str("codigo"), c.Codigo);
                    Assert.Equal(e.Num("comprimento"), c.Comprimento);
                }
                Assert.Equal((int)e.Num("fases").Value, c.Fases);
                Assert.Equal(e.Str("tensao"), c.Tensao);
                Assert.Equal(e.Str("nivel"), c.Nivel);
                Assert.Equal(e.Str("descricao"), c.Descricao);
            }
        }

        [Fact]
        public void Vinculo_regras_do_seed_de_Pernambuco()
        {
            RegrasVinculo rv = Apoio.Pernambuco.Vinculo;
            InfoCabo a35 = rv.InfoCabo("A35P", "I", Anotacoes.LerCabo("3xA35P(7,94)-13,8kV 45 m"));
            Assert.Equal("P", a35.ValorCabo);
            Assert.Equal(2, a35.Pontas);
            InfoEstrutura n1 = rv.InfoEstrutura("N1");
            Assert.Equal(2, n1.Qtd);
            Assert.Equal("IGUAIS", n1.Lados);
            Assert.Null(RegrasVinculo.PodeVincular(a35, n1));
            // lados iguais: N1 com A35P não aceita C16
            InfoCabo c16 = rv.InfoCabo("C16", "I", Anotacoes.LerCabo("3xC16(16) 30 m"));
            List<Ocupante> ocup = new List<Ocupante> { new Ocupante { Info = a35 } };
            Assert.Contains("exige cabos iguais", RegrasVinculo.PodeComOcupantes(c16, n1, ocup, false));
            // T3 = amarração, só cabo M; RAMA25 (multiplex) só operação I
            InfoEstrutura t3 = rv.InfoEstrutura("T3");
            Assert.Contains("só aceita cabo M", RegrasVinculo.PodeVincular(rv.InfoCabo("S04", "I", null), t3));
            Assert.Contains("não aceita a operação M", RegrasVinculo.PodeVincular(rv.InfoCabo("RAMA25", "M", null), rv.InfoEstrutura("B2")));
            // N4: tipocabo + cabo + cabo2 (descrição completa de cada cabo)
            InfoEstrutura n4 = rv.InfoEstrutura("N4");
            List<string> avisos = new List<string>();
            List<KeyValuePair<string, string>> ps = rv.Parametros(n4, new List<Ocupante> { new Ocupante { Info = a35 } }, avisos);
            Assert.Equal(new[] { "TIPOCABO", "CABO", "CABO2" }, ps.Select(p => p.Key).ToArray());
            Assert.Equal(new[] { "P", "3XA35P_7.94_-13.8KV", null }, ps.Select(p => p.Value).ToArray());
            // dois cabos diferentes: tipocabo = o maior (C > P > A)
            ps = rv.Parametros(n4, new List<Ocupante> { new Ocupante { Info = a35 }, new Ocupante { Info = c16 } }, avisos);
            Assert.Equal("C", ps[0].Value);
            Assert.Contains(avisos, a => a.Contains("usado o maior (C)"));
        }

        [Fact]
        public void Casa_lista_regex_e_vazio()
        {
            Assert.True(RegrasVinculo.Casa("", "N1"));
            Assert.True(RegrasVinculo.Casa("N1, N2", "n2"));
            Assert.False(RegrasVinculo.Casa("N1, N2", "N3"));
            Assert.True(RegrasVinculo.Casa("^(CE|[NT])[1-4]", "n4"));
            Assert.True(RegrasVinculo.Casa(new List<object> { "B2", "R2" }, "R2"));
        }
    }
}
