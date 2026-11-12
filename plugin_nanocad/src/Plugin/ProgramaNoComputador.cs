// Plugin/ProgramaNoComputador.cs — "Abrir na aplicação" com o GridReader instalado neste computador (TASK-016 do plugin / TASK-143).
// Aberto: entrega a obra pela porta local e a janela dele vem para a frente. Fechado: grava a obra num arquivo e abre o .exe com ela.
// Não instalado (sem %APPDATA%\OrcamentoNanoCAD\programa.json ou sem o .exe): devolve false e o comando segue o caminho antigo (navegador).
using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using OrcamentoNanoCAD.Nucleo;

namespace OrcamentoNanoCAD
{
    public static class ProgramaNoComputador
    {
        private static string ArquivoRegistro { get { return Path.Combine(Conexao.Pasta, "programa.json"); } }

        [DllImport("user32.dll")]
        private static extern bool AllowSetForegroundWindow(int processo);

        public static ProgramaInstalado Registrado()
        {
            try
            {
                if (!File.Exists(ArquivoRegistro)) return null;
                return ProgramaInstalado.De(Json.Ler(File.ReadAllText(ArquivoRegistro, Encoding.UTF8)) as JObj);
            }
            catch (Exception) { return null; }
        }

        /// <summary>true = a obra foi entregue (ou o programa foi aberto com ela).</summary>
        public static bool Abrir(string chave, JObj envelope)
        {
            ProgramaInstalado p = Registrado();
            TransporteHttp http = new TransporteHttp();
            foreach (int porta in ProgramaLocal.Portas(p, Conexao.Dados.Url))
            {
                if (!ProgramaLocal.Aberto(http, porta)) continue;
                try { AllowSetForegroundWindow(-1); } catch (Exception) { }     // deixa a janela do GridReader passar na frente do nanoCAD
                string erro = ProgramaLocal.Entregar(http, porta, chave, envelope);
                if (erro == null)
                {
                    Aplicacao.Msg("Obra enviada ao GridReader aberto neste computador (a janela dele vem para a frente; se pedir, entre com seu login).");
                    return true;
                }
                Aplicacao.Msg("O GridReader aberto recusou a obra: " + erro);
                return true;
            }
            if (p == null || string.IsNullOrEmpty(p.Exe) || !File.Exists(p.Exe)) return false;
            try
            {
                string pasta = Path.Combine(Conexao.Pasta, "obras_para_abrir");
                Directory.CreateDirectory(pasta);
                string arquivo = Path.Combine(pasta, "obra_" + DateTime.Now.ToString("yyyyMMdd_HHmmss_fff") + ".json");
                File.WriteAllText(arquivo, Json.Escrever(ProgramaLocal.Pacote(chave, envelope), false), new UTF8Encoding(false));
                ProcessStartInfo i = new ProcessStartInfo(p.Exe, ProgramaLocal.Argumentos(arquivo));
                i.UseShellExecute = false;
                i.WorkingDirectory = Path.GetDirectoryName(p.Exe);
                Process.Start(i);
                Aplicacao.Msg("Abrindo o GridReader com a obra… (entre com seu login; a obra abre em seguida)");
                return true;
            }
            catch (Exception e)
            {
                Aplicacao.Msg("Não deu para abrir o GridReader (" + p.Exe + "): " + e.Message + ". Abrindo pelo caminho antigo.");
                return false;
            }
        }
    }
}
