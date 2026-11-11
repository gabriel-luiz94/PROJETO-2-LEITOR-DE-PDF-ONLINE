// Plugin/Aplicacao.cs — entrada do plugin no nanoCAD: carga (APPLOAD/NETLOAD), regras disponíveis e a sessão de cada desenho.
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using HostMgd.ApplicationServices;
using HostMgd.EditorInput;
using OrcamentoNanoCAD.Nucleo;
using Teigha.Runtime;

[assembly: ExtensionApplication(typeof(OrcamentoNanoCAD.Aplicacao))]

namespace OrcamentoNanoCAD
{
    public class Aplicacao : IExtensionApplication
    {
        public void Initialize()
        {
            try
            {
                Editor ed = Ed();
                if (ed != null)
                    ed.WriteMessage("\nOrçamento (Leitor de Projetos) " + ObraArquivo.VersaoPlugin + " carregado. Comando: ORCAMENTO (ajuda: ORC_AJUDA).\n");
            }
            catch (System.Exception) { /* sem documento aberto: tudo bem */ }
            try { AtualizacaoPlugin.Iniciar(); }
            catch (System.Exception e) { Log("atualização: " + e); }
            try { Sincronizador.Iniciar(); }
            catch (System.Exception e) { Msg("Ligação com a aplicação indisponível: " + e.Message); }
        }

        public void Terminate() { }

        public static Document Doc() { return Application.DocumentManager.MdiActiveDocument; }

        public static Editor Ed()
        {
            Document d = Doc();
            return d == null ? null : d.Editor;
        }

        public static void Msg(string texto)
        {
            Log(texto);
            Editor ed = Ed();
            if (ed != null) ed.WriteMessage("\n" + texto);
        }

        // ── log (TASK-004): cada comando, mensagem e erro, para o usuário mandar quando algo não funciona ──
        public static string ArquivoLog { get { return Path.Combine(Conexao.Pasta, "log.txt"); } }
        private static readonly object _travaLog = new object();

        public static void Log(string texto)
        {
            try
            {
                lock (_travaLog)
                {
                    Directory.CreateDirectory(Conexao.Pasta);
                    FileInfo f = new FileInfo(ArquivoLog);
                    if (f.Exists && f.Length > 1024 * 1024) File.Copy(ArquivoLog, ArquivoLog + ".1", true);
                    if (f.Exists && f.Length > 1024 * 1024) File.Delete(ArquivoLog);
                    File.AppendAllText(ArquivoLog, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + texto + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch (System.Exception) { }
        }

        // ── regras ──
        private static List<RegrasProjeto> _regras;

        public static string PastaDaDll
        {
            get { return Path.GetDirectoryName(typeof(Aplicacao).Assembly.Location) ?? ""; }
        }

        private static string LerRecurso(string nome)
        {
            using (Stream s = typeof(Aplicacao).Assembly.GetManifestResourceStream(nome))
            {
                if (s == null) return null;
                using (StreamReader r = new StreamReader(s, Encoding.UTF8)) return r.ReadToEnd();
            }
        }

        private static string LerSeExiste(string pasta, string arquivo)
        {
            string c = Path.Combine(pasta, arquivo);
            return File.Exists(c) ? File.ReadAllText(c, Encoding.UTF8) : null;
        }

        /// <summary>Regras de uma pasta "regras\&lt;codigo&gt;\" ao lado da DLL (processamento.json, classificacao.json, vinculacao.json, nome.txt).</summary>
        private static List<RegrasProjeto> RegrasExternas()
        {
            List<RegrasProjeto> l = new List<RegrasProjeto>();
            string raiz = Path.Combine(PastaDaDll, "regras");
            if (!Directory.Exists(raiz)) return l;
            foreach (string pasta in Directory.GetDirectories(raiz))
            {
                try
                {
                    string proc = LerSeExiste(pasta, "processamento.json");
                    if (proc == null) continue;
                    string codigo = Path.GetFileName(pasta);
                    string nome = LerSeExiste(pasta, "nome.txt");
                    l.Add(RegrasProjeto.Montar(codigo, string.IsNullOrEmpty(nome) ? codigo + " (pasta regras)" : nome.Trim(), pasta, proc,
                        LerSeExiste(pasta, "classificacao.json"), LerSeExiste(pasta, "vinculacao.json")));
                }
                catch (System.Exception e) { Msg("Regras ignoradas em " + pasta + ": " + e.Message); }
            }
            return l;
        }

        public static List<RegrasProjeto> Regras
        {
            get
            {
                if (_regras == null)
                {
                    _regras = CatalogoRegras.Carregar(LerRecurso, RegrasExternas());
                    foreach (RegrasProjeto r in _regras) ConfigsLocais.Aplicar(r);
                }
                return _regras;
            }
        }

        public static void RecarregarRegras() { _regras = null; }

        /// <summary>Regras para a tela inicial: as baixadas da aplicação (cache) primeiro, depois as embutidas / da pasta.</summary>
        public static List<RegrasProjeto> RegrasParaEscolher()
        {
            List<RegrasProjeto> l = new List<RegrasProjeto>();
            if (Conexao.Logado) l.AddRange(Conexao.RegrasEmCache());
            l.AddRange(Regras);
            return l;
        }

        // ── sessão por desenho ──
        private const string ChaveSessao = "OrcamentoNanoCAD.Sessao";
        private static readonly Dictionary<Document, Sessao> _sessoes = new Dictionary<Document, Sessao>();

        public static Sessao Sessao(Document doc, bool criar)
        {
            if (doc == null) return null;
            Sessao s = Compat.Tentar("dados do desenho (UserData)", () => doc.UserData.ContainsKey(ChaveSessao) ? doc.UserData[ChaveSessao] as Sessao : null, null);
            if (s == null) lock (_sessoes) _sessoes.TryGetValue(doc, out s);      // nanoCAD sem UserData: guarda aqui
            if (s == null && criar)
            {
                s = new Sessao(doc);
                Sessao nova = s;
                if (!Compat.Tentar("dados do desenho (UserData)", () => { doc.UserData[ChaveSessao] = nova; }))
                    lock (_sessoes) _sessoes[doc] = s;
            }
            return s;
        }
    }

    /// <summary>Estado do plugin num desenho: a obra montada, as regras escolhidas e as marcas na tela.</summary>
    public class Sessao
    {
        public readonly Document Doc;
        public Obra Obra;
        public RegrasProjeto Regras;
        public bool Modificada;
        public bool MarcasVisiveis = true;
        public readonly Marcas Marcas = new Marcas();
        public DesenhoBruto UltimoBruto;       // o último DWG lido, cru (refaz a obra quando regras/configuração mudam, sem reler o DWG)
        public Desenho UltimoDesenho;
        // ligação com a aplicação (Sincronizador): cada mudança na obra sobe a revisão; a conferência com a base vale para uma revisão
        public int Revisao;
        public int RevisaoConferida = -1;
        public VersaoAplicacao VersaoAplicacao;
        public VersaoAplicacao VersaoConferida;
        public DateTime? UltimaConferencia;
        public int ConferenciaPendencias;

        public Sessao(Document doc)
        {
            Doc = doc;
            Regras = Aplicacao.Regras.FirstOrDefault();
        }

        public string Arquivo { get { return Doc.Name ?? ""; } }

        public bool DesenhoSalvo
        {
            get
            {
                string n = Doc.Name ?? "";
                return n.Length > 0 && Path.IsPathRooted(n) && File.Exists(n);
            }
        }

        public string ArquivoObra { get { return DesenhoSalvo ? ObraArquivo.NomeDoArquivo(Arquivo) : null; } }

        /// <summary>Depois de qualquer mudança: marcas e painel em dia.</summary>
        public void Atualizar()
        {
            if (Obra == null) return;
            try { if (MarcasVisiveis) Marcas.Mostrar(Obra, Regras); else Marcas.Limpar(); }
            catch (System.Exception e) { Aplicacao.Msg("Marcas na tela indisponíveis: " + e.Message); }
            UI.PainelObra.AtualizarSeAberto(this);
        }
    }
}
