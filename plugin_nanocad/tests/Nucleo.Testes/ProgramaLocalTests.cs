// TASK-016: "Abrir na aplicação" entrega a obra ao GridReader instalado neste computador.
using System.Collections.Generic;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

namespace OrcamentoNanoCAD.Testes
{
    public class ProgramaLocalTests
    {
        [Fact]
        public void Procura_na_porta_registrada_na_padrao_e_na_do_orc_conectar_local()
        {
            ProgramaInstalado p = ProgramaInstalado.De((JObj)Json.Ler("{\"exe\":\"C:\\\\GridReader\\\\Leitor_Projetos.exe\",\"porta\":8123,\"versao\":\"2.2.0\"}"));
            Assert.Equal(8123, p.Porta);
            Assert.Equal(new[] { 8123, 8000, 8001 }, ProgramaLocal.Portas(p, "http://127.0.0.1:8001").ToArray());
            Assert.Equal(new[] { 8000 }, ProgramaLocal.Portas(null, "https://app.online.com").ToArray());
            Assert.Null(ProgramaInstalado.De((JObj)Json.Ler("{\"porta\":8000}")));
        }

        [Fact]
        public void So_conta_como_aberto_o_gridreader_instalado()
        {
            TransporteFalso t = new TransporteFalso();
            t.Rotas["GET /api/plugin/desktop"] = (c, tok) => TransporteFalso.Ok("{\"app\":\"GridReader\",\"desktop\":true}");
            Assert.True(ProgramaLocal.Aberto(t, 8000));
            TransporteFalso online = new TransporteFalso();
            online.Rotas["GET /api/plugin/desktop"] = (c, tok) => TransporteFalso.Ok("{\"app\":\"GridReader\",\"desktop\":false}");
            Assert.False(ProgramaLocal.Aberto(online, 8000));
            Assert.False(ProgramaLocal.Aberto(new TransporteFalso(), 8000));          // 404: outro programa na porta
        }

        [Fact]
        public void Entrega_chave_e_envelope_sem_token()
        {
            TransporteFalso t = new TransporteFalso();
            string corpo = null;
            t.Rotas["POST /api/plugin/receber"] = (c, tok) => { corpo = c; return TransporteFalso.Ok("{\"id\":\"x\"}"); };
            JObj env = new JObj().Set("tipo", "obra").Set("dados", new JObj());
            Assert.Null(ProgramaLocal.Entregar(t, 8000, "C:\\obras\\rua.dwg", env));
            JObj enviado = (JObj)Json.Ler(corpo);
            Assert.Equal("C:\\obras\\rua.dwg", enviado.Str("chave"));
            Assert.Equal("obra", enviado.Obj("envelope").Str("tipo"));
            Assert.Contains("POST /api/plugin/receber -", t.Chamadas);

            TransporteFalso recusa = new TransporteFalso();
            recusa.Rotas["POST /api/plugin/receber"] = (c, tok) => new RespostaHttp { Status = 403, Corpo = "{\"detail\":\"Só no programa instalado no computador.\"}" };
            Assert.Equal("HTTP 403: Só no programa instalado no computador.", ProgramaLocal.Entregar(recusa, 8000, "a", env));
            Assert.Equal("--abrir-obra \"C:\\x\\obra.json\"", ProgramaLocal.Argumentos("C:\\x\\obra.json"));
        }
    }
}
