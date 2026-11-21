using System;
using System.IO;

namespace OrcamentoNanoCAD.Nucleo
{
    internal static class FonteDxfAtiva
    {
        public static T Usar<T>(string arquivo, Func<bool> arquivoAtual, Action<string> exportar,
            Func<string, bool, T> ler, Action<Exception> registrarFalhaLimpeza)
        {
            if (string.Equals(Path.GetExtension(arquivo), ".dxf", StringComparison.OrdinalIgnoreCase)
                && File.Exists(arquivo) && arquivoAtual())
                return ler(arquivo, false);

            string pasta = Path.Combine(Path.GetTempPath(), "OrcaCore-DXF-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(pasta);
            try
            {
                string temporario = Path.Combine(pasta, "desenho.dxf");
                exportar(temporario);
                if (!File.Exists(temporario) || new FileInfo(temporario).Length == 0)
                    throw new IOException("O nanoCAD não gerou um DXF válido do desenho ativo.");
                return ler(temporario, true);
            }
            finally
            {
                try { Directory.Delete(pasta, true); }
                catch (IOException erro) { registrarFalhaLimpeza(erro); }
                catch (UnauthorizedAccessException erro) { registrarFalhaLimpeza(erro); }
            }
        }
    }
}
