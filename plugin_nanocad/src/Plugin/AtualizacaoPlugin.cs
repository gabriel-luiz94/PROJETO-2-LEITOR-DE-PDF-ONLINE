// Plugin/AtualizacaoPlugin.cs — atualização automática do plugin (TASK-013 do plugin / TASK-130 do principal).
//   • 20 s depois de abrir o nanoCAD e a cada 4 h: pergunta à aplicação (mesmo endereço do ORC_CONECTAR) se há versão mais nova;
//     se houver, baixa em segundo plano, confere o SHA-256 e troca a DLL ao lado da atual (Nucleo.Atualizacao.Instalar).
//   • A versão nova vale na PRÓXIMA vez que o nanoCAD abrir (o nanoCAD não descarrega a DLL em uso).
//   • Arquivo travado: um processo auxiliar (PowerShell, oculto) faz a troca assim que o nanoCAD fechar.
//   • ORC_ATUALIZAR_PLUGIN confere agora · ORC_VOLTAR_VERSAO volta para a versão anterior (e essa versão não é baixada de novo).
// Estado em %APPDATA%\OrcamentoNanoCAD\atualizacao.json: versão pendente (baixada, esperando reabrir) e versão desfeita.
using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;
using OrcamentoNanoCAD.Nucleo;

namespace OrcamentoNanoCAD
{
    public static class AtualizacaoPlugin
    {
#if NANOCAD5
        public const string Alvo = "nanocad5";
#else
        public const string Alvo = "nanocad26";
#endif
        private static readonly TimeSpan Intervalo = TimeSpan.FromHours(4);
        private static string ArquivoEstado { get { return Path.Combine(Conexao.Pasta, "atualizacao.json"); } }

        private static System.Windows.Forms.Timer _timer;
        private static DateTime _proxima = DateTime.MaxValue;
        private static volatile bool _ocupado;
        private static string _mensagem;           // resultado do trabalho em segundo plano (mostrado no thread do nanoCAD)
        private static readonly object _trava = new object();
        public static string Pendente, Ignorada;    // versão baixada esperando reabrir · versão desfeita pelo usuário

        public static string DllAtual { get { return typeof(AtualizacaoPlugin).Assembly.Location; } }

        public static string Status
        {
            get
            {
                if (!Conexao.Dados.AtualizarPlugin) return "atualização automática desligada";
                if (!string.IsNullOrEmpty(Pendente)) return "versão " + Pendente + " baixada — vale ao reabrir o nanoCAD";
                return "atualização automática ligada (versão " + ObraArquivo.VersaoPlugin + ")";
            }
        }

        public static void Iniciar()
        {
            CarregarEstado();
            string dll = DllAtual;
            if (!string.IsNullOrEmpty(dll)) Atualizacao.Limpar(dll);
            if (!string.IsNullOrEmpty(Pendente) && Atualizacao.Comparar(Pendente, ObraArquivo.VersaoPlugin) <= 0)
            {
                if (Pendente == ObraArquivo.VersaoPlugin) Aplicacao.Msg("Plugin atualizado para a versão " + ObraArquivo.VersaoPlugin + ".");
                Pendente = null;
                SalvarEstado();
            }
            if (_timer != null) return;
            _proxima = DateTime.Now.AddSeconds(20);
            _timer = new System.Windows.Forms.Timer { Interval = 5000 };
            _timer.Tick += (o, e) => Tique();
            _timer.Start();
        }

        private static void Tique()
        {
            string m = null;
            lock (_trava) { if (_mensagem != null) { m = _mensagem; _mensagem = null; } }
            if (m != null) { Aplicacao.Msg(m); UI.PainelObra.AtualizarStatusConexao(); }
            if (_ocupado || DateTime.Now < _proxima) return;
            _proxima = DateTime.Now + Intervalo;
            if (Conexao.Dados.AtualizarPlugin) Conferir(false);
        }

        /// <summary>Confere em segundo plano. manual = ORC_ATUALIZAR_PLUGIN (mostra também "nada novo" e os erros).</summary>
        public static void Conferir(bool manual)
        {
            if (_ocupado) { if (manual) Aplicacao.Msg("Já estou conferindo a atualização do plugin."); return; }
            string dll = DllAtual, baseUrl = Conexao.Dados.Base;
            if (string.IsNullOrEmpty(dll) || !File.Exists(dll)) { if (manual) Aplicacao.Msg("Não achei a DLL do plugin para atualizar."); return; }
            string atual = ObraArquivo.VersaoPlugin;
            if (!string.IsNullOrEmpty(Pendente) && Atualizacao.Comparar(Pendente, atual) > 0) atual = Pendente;   // não baixa de novo
            string ignorada = Ignorada;
            _ocupado = true;
            if (manual) Aplicacao.Msg("Conferindo atualização do plugin em " + baseUrl + "…");
            ThreadPool.QueueUserWorkItem(_ =>
            {
                string msg = null;
                try
                {
                    string motivo;
                    PacoteAtualizacao p = Atualizacao.Consultar(new TransporteHttp(), baseUrl, Alvo, atual, ignorada, out motivo);
                    if (p == null) { msg = manual ? "Plugin: " + motivo + "." : null; Aplicacao.Log("atualização: " + motivo); }
                    else
                    {
                        byte[] dados = Atualizacao.Baixar(new BaixadorHttp(), baseUrl, p, out motivo);
                        if (dados == null) msg = "Atualização do plugin " + p.Versao + " não instalada: " + motivo + ".";
                        else msg = Instalar(dll, dados, p.Versao);
                    }
                }
                catch (Exception e) { msg = manual ? "Atualização do plugin: " + e.Message : null; Aplicacao.Log("atualização: " + e); }
                lock (_trava) { _mensagem = msg; }
                _ocupado = false;
            });
        }

        private static string Instalar(string dll, byte[] dados, string versao)
        {
            string motivo;
            ResultadoTroca r = Atualizacao.Instalar(dll, dados, out motivo);
            if (r == ResultadoTroca.Erro) return "Atualização do plugin " + versao + " não instalada: " + motivo + ".";
            if (r == ResultadoTroca.PrecisaAuxiliar)
            {
                Aplicacao.Log("atualização: troca agendada para quando o nanoCAD fechar (" + motivo + ")");
                if (!AgendarTroca(dll)) return "Atualização do plugin " + versao + " baixada, mas não deu para trocar o arquivo (" + motivo + ").";
            }
            Pendente = versao;
            SalvarEstado();
            Aplicacao.Log("atualização: versão " + versao + " instalada em " + dll);
            return "Versão " + versao + " do plugin baixada — passa a valer quando você reabrir o nanoCAD.";
        }

        /// <summary>PowerShell oculto: espera este nanoCAD fechar e troca dll ← dll.nova (guardando a atual em .anterior).</summary>
        private static bool AgendarTroca(string dll)
        {
            try
            {
                int pid = Process.GetCurrentProcess().Id;
                string a = dll.Replace("'", "''");
                string script = "$ErrorActionPreference='Stop'; Wait-Process -Id " + pid + " -ErrorAction SilentlyContinue; Start-Sleep -Seconds 1; "
                                + "if (Test-Path -LiteralPath '" + a + ".nova') { "
                                + "Move-Item -Force -LiteralPath '" + a + "' -Destination '" + a + ".anterior'; "
                                + "Move-Item -Force -LiteralPath '" + a + ".nova' -Destination '" + a + "' }";
                ProcessStartInfo i = new ProcessStartInfo("powershell.exe",
                    "-NoProfile -NonInteractive -WindowStyle Hidden -EncodedCommand " + Convert.ToBase64String(Encoding.Unicode.GetBytes(script)));
                i.CreateNoWindow = true;
                i.UseShellExecute = false;
                Process.Start(i);
                return true;
            }
            catch (Exception e) { Aplicacao.Log("atualização: auxiliar não iniciou: " + e.Message); return false; }
        }

        /// <summary>ORC_VOLTAR_VERSAO: volta para a DLL anterior; a versão desfeita não é baixada de novo (só uma mais nova).</summary>
        public static void Voltar()
        {
            string motivo;
            if (!Atualizacao.Reverter(DllAtual, out motivo)) { Aplicacao.Msg("Não deu para voltar a versão: " + motivo + "."); return; }
            Ignorada = ObraArquivo.VersaoPlugin;
            Pendente = null;
            SalvarEstado();
            Aplicacao.Msg("Versão anterior do plugin restaurada — vale ao reabrir o nanoCAD. A versão " + Ignorada
                          + " não será baixada de novo (só uma mais nova).");
        }

        private static void CarregarEstado()
        {
            try
            {
                if (!File.Exists(ArquivoEstado)) return;
                JObj o = Json.Ler(File.ReadAllText(ArquivoEstado, Encoding.UTF8)) as JObj;
                if (o == null) return;
                Pendente = o.Str("pendente");
                Ignorada = o.Str("ignorada");
            }
            catch (Exception) { }
        }

        private static void SalvarEstado()
        {
            try
            {
                Directory.CreateDirectory(Conexao.Pasta);
                File.WriteAllText(ArquivoEstado, Json.Escrever(new JObj().Set("pendente", Pendente).Set("ignorada", Ignorada), true), new UTF8Encoding(false));
            }
            catch (Exception) { }
        }
    }
}
