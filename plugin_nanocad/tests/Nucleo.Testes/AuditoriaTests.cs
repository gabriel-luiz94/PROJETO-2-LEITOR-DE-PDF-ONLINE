using System;
using System.Linq;
using System.Collections.Generic;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class AuditoriaTests
    {
        [Fact]
        public void Escolher_outro_poste_preserva_anotacao_e_escolha_ao_reabrir()
        {
            var d = new Desenho();
            d.Postes.Add(new PosteDesenho { Handle="AA", X=0 });
            d.Postes.Add(new PosteDesenho { Handle="BB", X=5 });
            d.Textos.Add(new TextoDesenho { Handle="CC", Texto="9200D, N1", X=2, Altura=1, Layer="01_INSTALANDO" });
            var r = Apoio.Pernambuco;
            var o = LeitorProjeto.Ler(d,r);
            var p = o.Postes.Single(x => x.TemAnotacao);
            Assert.Equal("AA",p.Handle);
            Assert.Contains("P:BB",p.Candidatos);
            var escolhas = new JObj().Set(Auditoria.Origem(p.TextoHandle,p.ParteTexto,p.TextoOriginal),"P:BB");
            var nova = LeitorProjeto.Ler(d,r,escolhas);
            var outro = nova.Postes.Single(x => x.TemAnotacao);
            Assert.Equal("BB",outro.Handle);
            Assert.DoesNotContain(Auditoria.Motivos(nova,outro,r,r.OperacoesPorLayer()), m => m.Contains("destino possível"));
            var chave = Auditoria.Chave(outro);
            nova.Conferidos.Set(chave,Auditoria.Assinatura(outro));
            var lida = ObraArquivo.Ler(ObraArquivo.Gerar(nova,DateTime.UtcNow),r);
            Assert.Equal(nova.Conferidos.Str(chave),lida.Conferidos.Str(chave));
            Assert.Equal("P:BB",lida.Associacoes.Str(Auditoria.Origem("CC",0,"9200D, N1")));
        }

        [Fact]
        public void Tolerancia_e_revisao_nao_modificam_o_comprimento_exportado()
        {
            var r = RegrasProjeto.MontarDeObjetos("081","Teste","",Apoio.Pernambuco.Processamento,Apoio.Pernambuco.Classificacao,null);
            var o = LeitorProjeto.Ler(DesenhoExemplo.Montar(),r);
            var v = o.Vaos.First(x => x.MetrosLidos.HasValue);
            var dados = Json.Escrever(ObraArquivo.Dados(o),false);
            Assert.DoesNotContain(Auditoria.Motivos(o,v,r,r.OperacoesPorLayer()),m => m.StartsWith("Metragem"));
            r.Leitura.ToleranciaMetros = 1;
            Assert.Contains(Auditoria.Motivos(o,v,r,r.OperacoesPorLayer()),m => m.StartsWith("Metragem"));
            Assert.Equal(dados,Json.Escrever(ObraArquivo.Dados(o),false));
            var anterior = Auditoria.Fotografar(o);
            v.Metros = v.MetrosLidos + 1;
            Assert.NotEqual(anterior.Str(Auditoria.Chave(v)),Auditoria.Assinatura(v));
        }
    }
}
