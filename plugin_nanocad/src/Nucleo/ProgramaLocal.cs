// Nucleo/ProgramaLocal.cs — "Abrir na aplicação" entrega a obra ao GridReader instalado neste computador (TASK-016 do plugin / TASK-143).
//   1. Programa ABERTO: GET http://127.0.0.1:<porta>/api/plugin/desktop responde "GridReader" + desktop → POST /api/plugin/receber
//      {chave, envelope}: a janela do programa vem para a frente com a obra (sem login no plugin: só vale do próprio computador).
//   2. Programa FECHADO: o .exe registra onde está em %APPDATA%\OrcamentoNanoCAD\programa.json; o plugin grava {chave, envelope} num arquivo e
//      abre o .exe com --abrir-obra "<arquivo>" (camada do CAD).
//   3. Sem programa instalado: o caminho antigo (aplicação do ORC_CONECTAR, no navegador).
using System;
using System.Collections.Generic;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    public class ProgramaInstalado
    {
        public string Exe;
        public int Porta = 8000;
        public string Versao;

        public static ProgramaInstalado De(JObj o)
        {
            if (o == null || string.IsNullOrEmpty(o.Str("exe"))) return null;
            ProgramaInstalado p = new ProgramaInstalado();
            p.Exe = o.Str("exe");
            p.Porta = (int)(o.Num("porta") ?? 8000);
            p.Versao = o.Str("versao");
            return p;
        }
    }

    public static class ProgramaLocal
    {
        public const string Argumento = "--abrir-obra";

        public static string Base(int porta) { return "http://127.0.0.1:" + porta; }

        /// <summary>Portas a procurar: a registrada pelo programa, a padrão (8000) e a do ORC_CONECTAR quando é este computador.</summary>
        public static List<int> Portas(ProgramaInstalado p, string urlConexao)
        {
            List<int> l = new List<int>();
            if (p != null && p.Porta > 0) l.Add(p.Porta);
            l.Add(8000);
            Uri u;
            if (Uri.TryCreate((urlConexao ?? "").Trim(), UriKind.Absolute, out u) && (u.Host == "127.0.0.1" || u.Host == "localhost")) l.Add(u.Port);
            return l.Distinct().ToList();
        }

        /// <summary>O GridReader instalado está aberto nessa porta?</summary>
        public static bool Aberto(ITransporteHttp http, int porta)
        {
            RespostaHttp r = http.Enviar("GET", Base(porta) + "/api/plugin/desktop", null, null, 1500);
            if (r.Status != 200) return false;
            JObj o = Json.Ler(r.Corpo) as JObj;
            return o != null && (o.Str("app") ?? "") == "GridReader" && o.Bool("desktop", false);
        }

        /// <summary>Entrega a obra ao programa aberto. null = ok; senão o motivo.</summary>
        public static string Entregar(ITransporteHttp http, int porta, string chave, JObj envelope)
        {
            RespostaHttp r = http.Enviar("POST", Base(porta) + "/api/plugin/receber", Json.Escrever(Pacote(chave, envelope), false), null, 30000);
            if (r.Status == 200) return null;
            JObj o = Json.Ler(r.Corpo ?? "") as JObj;
            string det = o != null ? (o.Str("detail") ?? "") : "";
            return r.Status == 0 ? (r.Erro ?? "sem resposta") : "HTTP " + r.Status + (det.Length > 0 ? ": " + det : "");
        }

        /// <summary>O que vai no arquivo (programa fechado) e no POST (programa aberto).</summary>
        public static JObj Pacote(string chave, JObj envelope) { return new JObj().Set("chave", chave ?? "").Set("envelope", envelope); }

        /// <summary>Argumentos para abrir o .exe com a obra.</summary>
        public static string Argumentos(string arquivo) { return Argumento + " \"" + (arquivo ?? "").Replace("\"", "") + "\""; }
    }
}
