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

        public static string AcrescentarAnotacoes(string chave, JObj envelope, string lote)
        {
            Stopwatch t = Stopwatch.StartNew();
            ProgramaInstalado p = Registrado();
            if (p == null) { Aplicacao.Log("anotações: OrçaCore não registrado neste computador"); return "Abra a obra de destino no OrçaCore deste computador."; }
            var http = new TransporteHttp();
            bool aberto = ProgramaLocal.Aberto(http, p.Porta);
            Aplicacao.Log("anotações: porta " + p.Porta + (aberto ? " respondeu" : " não respondeu") + " (" + t.ElapsedMilliseconds + " ms)");
            if (!aberto) return "Abra a obra de destino no OrçaCore deste computador.";
            if (string.IsNullOrEmpty(p.ChaveAnotacoes)) return "Atualize ou reinicie o OrçaCore para receber anotações.";
            if (p.Pid > 0) { try { AllowSetForegroundWindow(p.Pid); } catch (Exception) { } }
            string erro = ProgramaLocal.EntregarAnotacoes(http, p.Porta, p.ChaveAnotacoes, chave, envelope, lote);
            Aplicacao.Log("anotações: lote " + lote + " → " + (erro ?? "acrescentadas") + " (" + t.ElapsedMilliseconds + " ms)");
            return erro;
        }

        /// <summary>TASK-020: abre o projeto <desenho>.obra no OrçaCore deste computador. Aberto: entrega o caminho e a janela dele vem para a
        /// frente com o projeto (Abrir/Mesclar/Sobrescrever, se já houver dados na tela). Fechado: abre o .exe já com o projeto.
        /// Sem o programa instalado: o projeto fica na pasta do desenho (abrir à mão no OrçaCore).
        /// 0.16.5 (TASK-039): a conversa com o OrçaCore roda em SEGUNDO PLANO — antes prendia a tela do nanoCAD enquanto esperava a resposta
        /// (até 30 s); a mensagem final volta para a linha de comando. Cada passo vai para o log com o tempo.</summary>
        public static void AbrirProjeto(string chave, string caminho)
        {
            Aplicacao.Msg("Enviando o projeto ao OrçaCore deste computador…");
            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                string msg;
                try { msg = Entregar(chave, caminho); }
                catch (Exception e) { msg = "Não deu para abrir no OrçaCore: " + e.Message + ". O projeto está em " + caminho + "."; Aplicacao.Log("abrir no OrçaCore: " + e); }
                Aplicacao.NaTela(() => Aplicacao.Msg(msg));
            });
        }

        private static string Entregar(string chave, string caminho)
        {
            Stopwatch t = Stopwatch.StartNew();
            ProgramaInstalado p = Registrado();
            TransporteHttp http = new TransporteHttp();
            foreach (int porta in ProgramaLocal.Portas(p, Conexao.Dados.Url))
            {
                bool aberto = ProgramaLocal.Aberto(http, porta);
                Aplicacao.Log("abrir no OrçaCore: porta " + porta + (aberto ? " respondeu" : " não respondeu") + " (" + t.ElapsedMilliseconds + " ms)");
                if (!aberto) continue;
                try { AllowSetForegroundWindow(-1); } catch (Exception) { }     // deixa a janela do OrçaCore passar na frente do nanoCAD
                string erro = ProgramaLocal.EntregarProjeto(http, porta, chave, caminho);
                Aplicacao.Log("abrir no OrçaCore: entrega " + (erro ?? "ok") + " (" + t.ElapsedMilliseconds + " ms)");
                return erro == null ? "Projeto enviado ao OrçaCore aberto neste computador (a janela dele vem para a frente; se pedir, entre com seu login)."
                    : "O OrçaCore aberto não abriu o projeto: " + erro + ". Abra à mão: 📁 Projeto → Abrir → " + caminho;
            }
            if (p != null && !string.IsNullOrEmpty(p.Exe) && File.Exists(p.Exe))
            {
                try
                {
                    ProcessStartInfo i = new ProcessStartInfo(p.Exe, ProgramaLocal.ArgumentosProjeto(caminho));
                    i.UseShellExecute = false;
                    i.WorkingDirectory = Path.GetDirectoryName(p.Exe);
                    Process.Start(i);
                    Aplicacao.Log("abrir no OrçaCore: programa fechado, abrindo " + p.Exe + " (" + t.ElapsedMilliseconds + " ms)");
                    return "Abrindo o OrçaCore com o projeto… (entre com seu login; o projeto abre em seguida)";
                }
                catch (Exception e) { return "Não deu para abrir o OrçaCore (" + p.Exe + "): " + e.Message + "."; }
            }
            return "O OrçaCore não foi encontrado neste computador (abra-o uma vez para o plugin achá-lo). O projeto está em " + caminho
                   + ": abra no OrçaCore em 📁 Projeto → Abrir.";
        }
    }
}
