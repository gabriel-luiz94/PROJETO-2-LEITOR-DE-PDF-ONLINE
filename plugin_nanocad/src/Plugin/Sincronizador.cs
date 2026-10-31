// Plugin/Sincronizador.cs — ligação com a aplicação aberta (localhost ou online), TASK-003 do plugin.
//   Conexao       endereço, usuário e tokens (DPAPI do Windows, só este usuário), cache das regras baixadas (%APPDATA%\OrcamentoNanoCAD)
//   Sincronizador a cada segundo (timer da tela, no thread do nanoCAD): consulta a versão da aplicação no intervalo escolhido;
//                 regras mudaram → baixa e refaz a obra (mesmo desenho, ajustes mantidos); base/conversão mudou ou a obra mudou →
//                 confere de novo na aplicação quais ativos não estão na Base de Orçamento. A rede roda em segundo plano; a obra
//                 só é mexida no thread do nanoCAD (no "tique" do timer).
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using HostMgd.ApplicationServices;
using OrcamentoNanoCAD.Nucleo;

namespace OrcamentoNanoCAD
{
    public static class Conexao
    {
        public static readonly string Pasta = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OrcamentoNanoCAD");
        private static string Arquivo { get { return Path.Combine(Pasta, "conexao.json"); } }
        private static readonly byte[] Entropia = Encoding.UTF8.GetBytes("OrcamentoNanoCAD.tokens");

        public static ConexaoAplicacao Dados = new ConexaoAplicacao();
        public static ClienteAplicacao Cliente;

        static Conexao()
        {
            try { Carregar(); } catch (Exception) { Dados = new ConexaoAplicacao(); }
            NovoCliente();
        }

        private static void NovoCliente()
        {
            Cliente = new ClienteAplicacao(Dados, new TransporteHttp());
            Cliente.TokensMudaram = c => Salvar();
        }

        public static bool Logado { get { return Dados.Logado; } }

        private static void Carregar()
        {
            if (!File.Exists(Arquivo)) return;
            JObj o = Json.Ler(File.ReadAllText(Arquivo, Encoding.UTF8)) as JObj;
            Dados = ConexaoAplicacao.De(o);
            string prot = o != null ? o.Str("tokens") : null;
            if (string.IsNullOrEmpty(prot)) return;
            try
            {
                byte[] b = ProtectedData.Unprotect(Convert.FromBase64String(prot), Entropia, DataProtectionScope.CurrentUser);
                JObj t = Json.Ler(Encoding.UTF8.GetString(b)) as JObj;
                if (t != null) { Dados.Token = t.Str("token"); Dados.TokenRenovacao = t.Str("renovacao"); }
            }
            catch (Exception) { /* outro usuário/máquina: pede login de novo */ }
        }

        public static void Salvar()
        {
            Directory.CreateDirectory(Pasta);
            JObj o = Dados.ParaJson(false);
            if (Dados.Logado)
            {
                JObj t = new JObj().Set("token", Dados.Token).Set("renovacao", Dados.TokenRenovacao);
                byte[] b = ProtectedData.Protect(Encoding.UTF8.GetBytes(Json.Escrever(t, false)), Entropia, DataProtectionScope.CurrentUser);
                o.Set("tokens", Convert.ToBase64String(b));
            }
            File.WriteAllText(Arquivo, Json.Escrever(o, true), new UTF8Encoding(false));
        }

        /// <summary>Troca endereço/usuário: entra com a senha (que não é guardada).</summary>
        public static void Entrar(string url, string email, string senha, int intervalo, bool automatico)
        {
            ConexaoAplicacao nova = new ConexaoAplicacao();
            nova.Url = url; nova.IntervaloSegundos = Math.Max(2, intervalo); nova.Automatico = automatico;
            ClienteAplicacao c = new ClienteAplicacao(nova, new TransporteHttp());
            c.Entrar(email, senha);
            Dados = nova;
            NovoCliente();
            Salvar();
        }

        public static void Ajustar(int intervalo, bool automatico)
        {
            Dados.IntervaloSegundos = Math.Max(2, intervalo);
            Dados.Automatico = automatico;
            Salvar();
        }

        public static void Sair()
        {
            Cliente.Sair();
            Salvar();
        }

        // ── cache das regras baixadas (para abrir sem a aplicação) ──
        private static string PastaCache
        {
            get
            {
                string h;
                using (SHA1 sha = SHA1.Create())
                    h = BitConverter.ToString(sha.ComputeHash(Encoding.UTF8.GetBytes(Dados.Base.ToLowerInvariant()))).Replace("-", "").Substring(0, 12);
                return Path.Combine(Pasta, "cache", h);
            }
        }

        public static void SalvarCache(string projeto, JObj cru)
        {
            Directory.CreateDirectory(PastaCache);
            File.WriteAllText(Path.Combine(PastaCache, "regras_" + Seguro(projeto) + ".json"), Json.Escrever(cru, false), new UTF8Encoding(false));
        }

        private static string Seguro(string s)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in s ?? "") sb.Append(char.IsLetterOrDigit(c) ? c : '_');
            return sb.ToString();
        }

        /// <summary>Regras de projetos já baixadas desta aplicação (mostradas na tela inicial).</summary>
        public static List<RegrasProjeto> RegrasEmCache()
        {
            List<RegrasProjeto> l = new List<RegrasProjeto>();
            try
            {
                if (!Directory.Exists(PastaCache)) return l;
                foreach (string f in Directory.GetFiles(PastaCache, "regras_*.json"))
                {
                    try { l.Add(ClienteAplicacao.RegrasDe((JObj)Json.Ler(File.ReadAllText(f, Encoding.UTF8)), Dados.Base)); }
                    catch (Exception) { }
                }
            }
            catch (Exception) { }
            return l;
        }

        /// <summary>Baixa agora (bloqueia até o timeout) as regras do projeto; guarda no cache.</summary>
        public static RegrasProjeto BaixarRegras(string projeto)
        {
            JObj cru = Cliente.RegrasCruas(projeto);
            SalvarCache(projeto, cru);
            return ClienteAplicacao.RegrasDe(cru, Dados.Base);
        }
    }

    public static class Sincronizador
    {
        private enum Tipo { Versao, Regras, Conferir }

        private class Trabalho
        {
            public Tipo Tipo;
            public Sessao Sessao;
            public string Projeto;
            public int Revisao;
            public JObj Dados;
            public VersaoAplicacao Versao;
            public JObj RegrasCruas;
            public ConferenciaBase Conferencia;
            public ErroAplicacao Erro;
        }

        private static System.Windows.Forms.Timer _timer;
        private static readonly object _trava = new object();
        private static Trabalho _pronto;
        private static bool _ocupado;
        private static DateTime _proximaVersao = DateTime.MinValue;
        private static DateTime _esperarAte = DateTime.MinValue;
        private static int _falhas;
        private static bool _pararAteEntrar;
        private static string _status = "";

        public static string Status { get { return _status; } }

        public static void Iniciar()
        {
            if (_timer != null) return;
            _timer = new System.Windows.Forms.Timer();
            _timer.Interval = 1000;
            _timer.Tick += (o, e) => Tique();
            _timer.Start();
            AtualizarStatus();
        }

        /// <summary>Depois de conectar/sincronizar à mão: consulta já e volta a tentar mesmo depois de erro.</summary>
        public static void Agora()
        {
            _proximaVersao = DateTime.MinValue;
            _esperarAte = DateTime.MinValue;
            _falhas = 0;
            _pararAteEntrar = false;
            AtualizarStatus();
        }

        private static void Definir(string s)
        {
            if (s == _status) return;
            _status = s;
            UI.PainelObra.AtualizarStatusConexao();
        }

        private static void AtualizarStatus()
        {
            if (!Conexao.Logado) { Definir("⚪ Sem ligação com a aplicação — regras " + RegrasTexto() + ". Conectar: ORC_CONECTAR"); return; }
            Sessao s = Aplicacao.Sessao(Aplicacao.Doc(), false);
            string conf = s != null && s.UltimaConferencia.HasValue
                ? " · base conferida " + s.UltimaConferencia.Value.ToString("HH:mm:ss") + (s.ConferenciaPendencias > 0 ? " (" + s.ConferenciaPendencias + " linha(s) com ativo fora da base)" : " (tudo na base)")
                : "";
            Definir("🔗 " + Conexao.Dados.Base + (Conexao.Dados.Automatico ? "" : " (manual)") + " · regras " + RegrasTexto() + conf);
        }

        private static string RegrasTexto()
        {
            Sessao s = Aplicacao.Sessao(Aplicacao.Doc(), false);
            if (s == null || s.Regras == null) return "—";
            return s.Regras.DaAplicacao ? "da aplicação" : "embutidas (" + s.Regras.Nome + ")";
        }

        private static void Tique()
        {
            try
            {
                Trabalho t = null;
                lock (_trava) { if (_pronto != null) { t = _pronto; _pronto = null; } }
                if (t != null) Aplicar(t);
                if (_ocupado || !Conexao.Logado || _pararAteEntrar || DateTime.Now < _esperarAte) return;
                Document d = Aplicacao.Doc();
                Sessao s = Aplicacao.Sessao(d, false);
                if (s == null || s.Obra == null || string.IsNullOrEmpty(s.Obra.Projeto)) return;
                if (s.Revisao != s.RevisaoConferida && Conexao.Dados.Automatico)
                {
                    Iniciar(new Trabalho { Tipo = Tipo.Conferir, Sessao = s, Projeto = s.Obra.Projeto, Revisao = s.Revisao, Dados = ObraArquivo.Dados(s.Obra) });
                    return;
                }
                if (Conexao.Dados.Automatico && DateTime.Now >= _proximaVersao)
                {
                    _proximaVersao = DateTime.Now.AddSeconds(Conexao.Dados.IntervaloSegundos);
                    Iniciar(new Trabalho { Tipo = Tipo.Versao, Sessao = s, Projeto = s.Obra.Projeto });
                }
            }
            catch (Exception e) { Definir("⚠ Ligação com a aplicação: " + e.Message); }
        }

        private static void Iniciar(Trabalho t)
        {
            _ocupado = true;
            ClienteAplicacao cliente = Conexao.Cliente;
            ThreadPool.QueueUserWorkItem(_ =>
            {
                try
                {
                    if (t.Tipo == Tipo.Versao) t.Versao = cliente.Versao(t.Projeto);
                    else if (t.Tipo == Tipo.Regras) t.RegrasCruas = cliente.RegrasCruas(t.Projeto);
                    else t.Conferencia = cliente.Conferir(t.Projeto, t.Dados);
                }
                catch (ErroAplicacao e) { t.Erro = e; }
                catch (Exception e) { t.Erro = new ErroAplicacao(0, e.Message); }
                lock (_trava) { _pronto = t; }
            });
        }

        private static void Aplicar(Trabalho t)
        {
            _ocupado = false;
            Sessao s = t.Sessao;
            if (t.Erro != null)
            {
                _falhas++;
                if (t.Erro.PrecisaEntrar) { _pararAteEntrar = true; Definir("⚠ " + t.Erro.Message); return; }
                _esperarAte = DateTime.Now.AddSeconds(Math.Min(60, Conexao.Dados.IntervaloSegundos * Math.Pow(2, Math.Min(_falhas, 4))));
                Definir("⚠ " + t.Erro.Message + " — tentando de novo às " + _esperarAte.ToString("HH:mm:ss") + " (o plugin segue com o que já tem)");
                return;
            }
            _falhas = 0;
            if (s.Obra == null || s.Obra.Projeto != t.Projeto) { AtualizarStatus(); return; }
            if (t.Tipo == Tipo.Versao)
            {
                s.VersaoAplicacao = t.Versao;
                if (s.Regras != null && s.Regras.DaAplicacao && s.Regras.Codigo == t.Projeto && t.Versao.Regras != s.Regras.Versao)
                    Iniciar(new Trabalho { Tipo = Tipo.Regras, Sessao = s, Projeto = t.Projeto });
                else if (s.VersaoConferida == null || s.VersaoConferida.Base != t.Versao.Base || s.VersaoConferida.Conversao != t.Versao.Conversao
                         || s.VersaoConferida.Regras != t.Versao.Regras)
                    s.RevisaoConferida = -1;                       // base/conversão/vínculo mudou: conferir de novo
            }
            else if (t.Tipo == Tipo.Regras)
            {
                Conexao.SalvarCache(t.Projeto, t.RegrasCruas);
                TrocarRegras(s, ClienteAplicacao.RegrasDe(t.RegrasCruas, Conexao.Dados.Base));
                Aplicacao.Msg("Regras do projeto " + t.Projeto + " atualizadas pela aplicação; obra refeita (ajustes mantidos). " + Relatorios.Resumo(s.Obra));
            }
            else if (t.Revisao == s.Revisao)
            {
                t.Conferencia.Aplicar(s.Obra, s.Regras, t.Conferencia.Versao.ProjetoNome);
                s.RevisaoConferida = s.Revisao;
                s.VersaoConferida = t.Conferencia.Versao;
                s.UltimaConferencia = DateTime.Now;
                s.ConferenciaPendencias = t.Conferencia.ComPendencia;
                s.Atualizar();
            }
            AtualizarStatus();
        }

        /// <summary>Passa a sessão para outras regras e refaz a obra: relê o último desenho lido (sem tocar no DWG) mantendo os ajustes.</summary>
        public static void TrocarRegras(Sessao s, RegrasProjeto novas)
        {
            s.Regras = novas;
            Obra antiga = s.Obra;
            if (antiga != null)
            {
                if (antiga.Modo == "ler" && s.UltimoDesenho != null)
                {
                    Obra nova = LeitorProjeto.Ler(s.UltimoDesenho, novas);
                    nova.Projeto = antiga.Projeto; nova.Nome = antiga.Nome; nova.SalvoEm = antiga.SalvoEm;
                    Construtor.ReaplicarAjustes(nova, antiga, novas);
                    s.Obra = nova;
                }
                else
                {
                    antiga.NomeRegras = novas.Nome;
                    Construtor.Recalcular(antiga, novas);
                }
            }
            s.Revisao++;
            s.Atualizar();
        }
    }
}
