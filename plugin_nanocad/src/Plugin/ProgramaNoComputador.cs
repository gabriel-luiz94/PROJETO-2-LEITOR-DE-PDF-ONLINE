// Plugin/ProgramaNoComputador.cs — "Abrir na aplicação" com o OrçaCore instalado neste computador (TASK-016 / TASK-020 do plugin).
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

        /// <summary>TASK-020: abre o projeto <desenho>.obra no OrçaCore deste computador. Aberto: entrega o caminho e a janela dele vem para a
        /// frente com o projeto (Abrir/Mesclar/Sobrescrever, se já houver dados na tela). Fechado: abre o .exe já com o projeto.
        /// Sem o programa instalado: o projeto fica na pasta do desenho (abrir à mão no OrçaCore).</summary>
        public static void AbrirProjeto(string chave, string caminho)
        {
            ProgramaInstalado p = Registrado();
            TransporteHttp http = new TransporteHttp();
            foreach (int porta in ProgramaLocal.Portas(p, Conexao.Dados.Url))
            {
                if (!ProgramaLocal.Aberto(http, porta)) continue;
                try { AllowSetForegroundWindow(-1); } catch (Exception) { }     // deixa a janela do OrçaCore passar na frente do nanoCAD
                string erro = ProgramaLocal.EntregarProjeto(http, porta, chave, caminho);
                if (erro == null) Aplicacao.Msg("Projeto enviado ao OrçaCore aberto neste computador (a janela dele vem para a frente; se pedir, entre com seu login).");
                else Aplicacao.Msg("O OrçaCore aberto não abriu o projeto: " + erro + ". Abra à mão: 📁 Projeto → Abrir → " + caminho);
                return;
            }
            if (p != null && !string.IsNullOrEmpty(p.Exe) && File.Exists(p.Exe))
            {
                try
                {
                    ProcessStartInfo i = new ProcessStartInfo(p.Exe, ProgramaLocal.ArgumentosProjeto(caminho));
                    i.UseShellExecute = false;
                    i.WorkingDirectory = Path.GetDirectoryName(p.Exe);
                    Process.Start(i);
                    Aplicacao.Msg("Abrindo o OrçaCore com o projeto… (entre com seu login; o projeto abre em seguida)");
                    return;
                }
                catch (Exception e) { Aplicacao.Msg("Não deu para abrir o OrçaCore (" + p.Exe + "): " + e.Message + "."); }
            }
            Aplicacao.Msg("O OrçaCore não foi encontrado neste computador (abra-o uma vez para o plugin achá-lo). O projeto está em " + caminho
                          + ": abra no OrçaCore em 📁 Projeto → Abrir.");
        }
    }
}
