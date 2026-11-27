using System;
using System.IO;
using System.Text;
using System.Windows.Forms;
using OrcamentoNanoCAD.Nucleo;
using Teigha.Runtime;

namespace OrcamentoNanoCAD
{
    public partial class Comandos
    {
        [CommandMethod("CSEXPORTAR")]
        public void ExportarDesenhoJson()
        {
            Rodar("CSEXPORTAR", s =>
            {
                using (var janela = new SaveFileDialog { Filter = "Desenho JSON|*.json", OverwritePrompt = true })
                {
                    if (janela.ShowDialog() != DialogResult.OK) return;
                    string temp = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".dxf");
                    try
                    {
                        DxfDoDesenho.Exportar(s.Doc.Database, temp);
                        JObj desenho;
                        using (var leitor = new StreamReader(temp, new UTF8Encoding(false, true), true))
                            desenho = DesenhoJson.LerDxf(leitor);
                        File.WriteAllText(janela.FileName, Json.Escrever(desenho, false), new UTF8Encoding(false));
                        Aplicacao.Msg("JSON exportado. Esta etapa inclui todas as seções do DXF; use uma cópia sem layouts para o editor web.");
                    }
                    finally { if (File.Exists(temp)) File.Delete(temp); }
                }
            });
        }

        [CommandMethod("CSIMPORTAR")]
        public void ImportarDesenhoJson()
        {
            Rodar("CSIMPORTAR", s =>
            {
                using (var origem = new OpenFileDialog { Filter = "Desenho JSON|*.json" })
                using (var destino = new SaveFileDialog { Filter = "Desenho DXF|*.dxf", OverwritePrompt = true })
                {
                    if (origem.ShowDialog() != DialogResult.OK || destino.ShowDialog() != DialogResult.OK) return;
                    if (File.Exists(destino.FileName)) throw new IOException("Escolha um novo arquivo DXF; não sobrescrevemos desenhos.");
                    JObj desenho = Json.Ler(File.ReadAllText(origem.FileName, new UTF8Encoding(false, true))) as JObj;
                    using (var buffer = new StringWriter(System.Globalization.CultureInfo.InvariantCulture))
                    {
                        DesenhoJson.EscreverDxf(desenho, buffer);
                        using (var arquivo = new FileStream(destino.FileName, FileMode.CreateNew))
                        using (var escritor = new StreamWriter(arquivo, new UTF8Encoding(false)))
                            escritor.Write(buffer.ToString());
                    }
                    Aplicacao.Msg("DXF gerado. Abra o novo arquivo para conferir; o desenho ativo não foi alterado.");
                }
            });
        }
    }
}
