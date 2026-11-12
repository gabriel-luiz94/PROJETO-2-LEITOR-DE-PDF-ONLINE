// Plugin/Catalogos.cs — a Base de Orçamento de cada projeto para o autocompletar (TASK-017 do plugin).
// Na memória por 10 min; cópia em %APPDATA%\OrcamentoNanoCAD\cache\catalogo_<projeto>.json para usar sem a aplicação.
// Obter devolve na hora o que houver (memória ou cópia) e, se estiver velho, busca na aplicação em segundo plano e chama `pronto`
// no thread da tela quando chegar.
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using OrcamentoNanoCAD.Nucleo;

namespace OrcamentoNanoCAD
{
    public static class Catalogos
    {
        private class Entrada { public CatalogoBase Cat; public DateTime Quando; }
        private static readonly Dictionary<string, Entrada> _mem = new Dictionary<string, Entrada>();
        private static readonly object _trava = new object();

        private static string Arquivo(string projeto)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in projeto ?? "") sb.Append(char.IsLetterOrDigit(c) ? c : '_');
            return Path.Combine(Conexao.Pasta, "cache", "catalogo_" + sb + ".json");
        }

        public static CatalogoBase Obter(string projeto, Control tela, Action<CatalogoBase, string> pronto)
        {
            if (string.IsNullOrEmpty(projeto)) return null;
            Entrada e;
            lock (_trava) _mem.TryGetValue(projeto, out e);
            if (e == null)
            {
                try
                {
                    string f = Arquivo(projeto);
                    if (File.Exists(f)) e = new Entrada { Cat = CatalogoBase.De(Json.Ler(File.ReadAllText(f, Encoding.UTF8)) as JObj), Quando = DateTime.MinValue };
                }
                catch (Exception) { }
            }
            if ((e == null || DateTime.Now - e.Quando > TimeSpan.FromMinutes(10)) && Conexao.Logado)
            {
                ClienteAplicacao cli = Conexao.Cliente;
                ThreadPool.QueueUserWorkItem(_ =>
                {
                    CatalogoBase novo = null;
                    string erro = null;
                    try
                    {
                        JObj cru = cli.CatalogoCru(projeto);
                        novo = CatalogoBase.De(cru);
                        lock (_trava) _mem[projeto] = new Entrada { Cat = novo, Quando = DateTime.Now };
                        try
                        {
                            Directory.CreateDirectory(Path.GetDirectoryName(Arquivo(projeto)));
                            File.WriteAllText(Arquivo(projeto), Json.Escrever(cru, false), new UTF8Encoding(false));
                        }
                        catch (Exception) { }
                    }
                    catch (Exception ex) { erro = ex.Message; }
                    try
                    {
                        if (pronto != null && tela != null && !tela.IsDisposed && tela.IsHandleCreated)
                            tela.BeginInvoke((Action)(() => { if (!tela.IsDisposed) pronto(novo, erro); }));
                    }
                    catch (Exception) { }
                });
            }
            return e != null ? e.Cat : null;
        }
    }
}
