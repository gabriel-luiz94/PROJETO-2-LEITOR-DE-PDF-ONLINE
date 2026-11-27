using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using OrcamentoNanoCAD.Nucleo;

namespace OrcamentoNanoCAD
{
    internal static class ProxyNoComputador
    {
        public static bool Disponivel(ProgramaInstalado programa)
        {
            return programa != null && programa.CadsuiteProxy && File.Exists(programa.Exe);
        }

        public static JObj Converter(string entrada, string saida)
        {
            var programa = ProgramaNoComputador.Registrado();
            if (!Disponivel(programa))
                throw new NotSupportedException("Atualize e abra o OrçaCore uma vez para registrar o conversor local de proxies.");

            string pasta = Path.Combine(Path.GetTempPath(), "orcacore-cadsuite-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(pasta);
            try
            {
                File.Copy(entrada, Path.Combine(pasta, "entrada.dxf"));
                var inicio = new ProcessStartInfo(programa.Exe, "--trabalhador-cadsuite-proxy \"" + pasta + "\"");
                inicio.UseShellExecute = false;
                inicio.CreateNoWindow = true;
                inicio.WorkingDirectory = Path.GetDirectoryName(programa.Exe);
                using (var processo = Process.Start(inicio))
                {
                    if (processo == null) throw new IOException("Não foi possível iniciar o conversor local.");
                    if (!processo.WaitForExit(90000))
                    {
                        try { processo.Kill(); processo.WaitForExit(5000); } catch (Exception) { }
                        throw new TimeoutException("O conversor local excedeu 90 segundos.");
                    }
                    string json = Path.Combine(pasta, "resultado.json");
                    var r = File.Exists(json) ? Json.Ler(File.ReadAllText(json, Encoding.UTF8)) as JObj : null;
                    if (processo.ExitCode != 0 || r == null)
                        throw new IOException("Falha no conversor local: " + (r == null ? "sem relatório." : r.Str("erro")));
                    if ((r.Num("convertidos") ?? 0) > 0) File.Copy(Path.Combine(pasta, "saida.dxf"), saida, true);
                    return r;
                }
            }
            finally
            {
                try { Directory.Delete(pasta, true); } catch (IOException) { } catch (UnauthorizedAccessException) { }
            }
        }
    }
}
