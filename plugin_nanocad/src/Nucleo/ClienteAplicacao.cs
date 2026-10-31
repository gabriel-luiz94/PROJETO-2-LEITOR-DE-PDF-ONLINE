// Nucleo/ClienteAplicacao.cs — conversa com a aplicação (localhost ou online) pela API (TASK-003 do plugin / TASK-104 do principal):
//   POST /api/auth/login, /api/auth/refresh      login e renovação do token (o token renova sozinho por até 30 dias)
//   GET  /api/plugin/versao?projeto=             impressões digitais das regras / conversão / base (consultado a cada poucos segundos)
//   GET  /api/plugin/regras?projeto=             regras do leitor + regras de vínculo gravadas na aplicação
//   POST /api/plugin/conferir                    a obra (linhas do .obra.json) conferida contra a Base de Orçamento
// O transporte HTTP fica atrás de ITransporteHttp para os testes rodarem sem rede. Sem dependências: HttpWebRequest existe no .NET 4.
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace OrcamentoNanoCAD.Nucleo
{
    public class RespostaHttp
    {
        public int Status;      // 0 = não conectou
        public string Corpo = "";
        public string Erro;
    }

    public interface ITransporteHttp
    {
        RespostaHttp Enviar(string metodo, string url, string corpoJson, string token, int timeoutMs);
    }

    public class TransporteHttp : ITransporteHttp
    {
        static TransporteHttp()
        {
            // .NET 4.0 (nanoCAD 5) não liga o TLS 1.2 sozinho; 3072 = Tls12 (precisa do .NET 4.5+ instalado no Windows)
            try { ServicePointManager.SecurityProtocol |= (SecurityProtocolType)3072; }
            catch (NotSupportedException) { }
        }

        public RespostaHttp Enviar(string metodo, string url, string corpoJson, string token, int timeoutMs)
        {
            RespostaHttp r = new RespostaHttp();
            try
            {
#pragma warning disable SYSLIB0014
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
#pragma warning restore SYSLIB0014
                req.Method = metodo;
                req.Timeout = timeoutMs;
                req.ReadWriteTimeout = timeoutMs;
                req.Accept = "application/json";
                if (!string.IsNullOrEmpty(token)) req.Headers["Authorization"] = "Bearer " + token;
                if (corpoJson != null)
                {
                    byte[] b = Encoding.UTF8.GetBytes(corpoJson);
                    req.ContentType = "application/json; charset=utf-8";
                    req.ContentLength = b.Length;
                    using (Stream s = req.GetRequestStream()) s.Write(b, 0, b.Length);
                }
                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse()) Ler(resp, r);
            }
            catch (WebException e)
            {
                HttpWebResponse resp = e.Response as HttpWebResponse;
                if (resp != null) using (resp) Ler(resp, r);
                else { r.Status = 0; r.Erro = e.Message; }
            }
            catch (Exception e) { r.Status = 0; r.Erro = e.Message; }
            return r;
        }

        private static void Ler(HttpWebResponse resp, RespostaHttp r)
        {
            r.Status = (int)resp.StatusCode;
            using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8)) r.Corpo = sr.ReadToEnd();
        }
    }

    /// <summary>Endereço, usuário e tokens. A senha nunca é guardada.</summary>
    public class ConexaoAplicacao
    {
        public string Url = "http://127.0.0.1:8000";
        public string Email = "";
        public string Token;
        public string TokenRenovacao;
        public int IntervaloSegundos = 5;
        public bool Automatico = true;

        public bool Logado { get { return !string.IsNullOrEmpty(Token) || !string.IsNullOrEmpty(TokenRenovacao); } }

        public string Base
        {
            get
            {
                string u = (Url ?? "").Trim().TrimEnd('/');
                if (u.Length > 0 && !u.StartsWith("http://", StringComparison.OrdinalIgnoreCase) && !u.StartsWith("https://", StringComparison.OrdinalIgnoreCase))
                    u = "http://" + u;
                return u;
            }
        }

        public JObj ParaJson(bool comTokens)
        {
            JObj o = new JObj();
            o.Set("url", Url); o.Set("email", Email); o.Set("intervalo", IntervaloSegundos); o.Set("automatico", Automatico);
            if (comTokens) { o.Set("token", Token); o.Set("renovacao", TokenRenovacao); }
            return o;
        }

        public static ConexaoAplicacao De(JObj o)
        {
            ConexaoAplicacao c = new ConexaoAplicacao();
            if (o == null) return c;
            c.Url = o.Str("url") ?? c.Url;
            c.Email = o.Str("email") ?? "";
            c.IntervaloSegundos = Math.Max(2, (int)(o.Num("intervalo") ?? 5));
            c.Automatico = o.Bool("automatico", true);
            c.Token = o.Str("token");
            c.TokenRenovacao = o.Str("renovacao");
            return c;
        }
    }

    public class ErroAplicacao : Exception
    {
        public readonly int Status;
        public ErroAplicacao(int status, string msg) : base(msg) { Status = status; }
        public bool SemConexao { get { return Status == 0; } }
        public bool PrecisaEntrar { get { return Status == 401; } }
    }

    /// <summary>Impressões digitais devolvidas por /api/plugin/versao (mudou = baixar/conferir de novo).</summary>
    public class VersaoAplicacao
    {
        public string Regras, Conversao, Base, Geral;
        public string ProjetoNome;

        public static VersaoAplicacao De(JObj resposta)
        {
            VersaoAplicacao v = new VersaoAplicacao();
            JObj o = resposta != null ? resposta.Obj("versao") : null;
            if (o != null) { v.Regras = o.Str("regras"); v.Conversao = o.Str("conversao"); v.Base = o.Str("base"); v.Geral = o.Str("geral"); }
            JObj p = resposta != null ? resposta.Obj("projeto") : null;
            if (p != null) v.ProjetoNome = p.Str("nome");
            return v;
        }
    }

    public class ClienteAplicacao
    {
        public readonly ConexaoAplicacao Conexao;
        private readonly ITransporteHttp _http;
        public int TimeoutMs = 15000;
        /// <summary>Chamado quando os tokens mudam (login/renovação), para a camada do CAD guardar.</summary>
        public Action<ConexaoAplicacao> TokensMudaram;

        public ClienteAplicacao(ConexaoAplicacao conexao, ITransporteHttp http)
        {
            Conexao = conexao;
            _http = http;
        }

        private static string Detalhe(RespostaHttp r)
        {
            try
            {
                JObj o = Json.Ler(r.Corpo) as JObj;
                if (o != null && o["detail"] is string) return (string)o["detail"];
                if (o != null && o.Obj("detail") != null) return string.Join("; ", Json.ListaDeTextos(o.Obj("detail")["erros"]).ToArray());
            }
            catch (ErroJson) { }
            return r.Erro ?? ("HTTP " + r.Status);
        }

        private static ErroAplicacao Erro(RespostaHttp r)
        {
            if (r.Status == 0) return new ErroAplicacao(0, "sem conexão com a aplicação (" + (r.Erro ?? "?") + ")");
            if (r.Status == 401) return new ErroAplicacao(401, "sessão expirada ou não autorizada: entre de novo (ORC_CONECTAR)");
            return new ErroAplicacao(r.Status, Detalhe(r));
        }

        public void Entrar(string email, string senha)
        {
            JObj corpo = new JObj().Set("email", email).Set("password", senha).Set("lembrar", true);
            RespostaHttp r = _http.Enviar("POST", Conexao.Base + "/api/auth/login", Json.Escrever(corpo, false), null, TimeoutMs);
            if (r.Status != 200) throw Erro(r);
            JObj o = (JObj)Json.Ler(r.Corpo);
            Conexao.Email = email;
            Conexao.Token = o.Str("access_token");
            Conexao.TokenRenovacao = o.Str("refresh_token");
            if (TokensMudaram != null) TokensMudaram(Conexao);
        }

        public void Sair()
        {
            Conexao.Token = null;
            Conexao.TokenRenovacao = null;
            if (TokensMudaram != null) TokensMudaram(Conexao);
        }

        private bool Renovar()
        {
            if (string.IsNullOrEmpty(Conexao.TokenRenovacao)) return false;
            RespostaHttp r = _http.Enviar("POST", Conexao.Base + "/api/auth/refresh", null, Conexao.TokenRenovacao, TimeoutMs);
            if (r.Status != 200) return false;
            Conexao.Token = ((JObj)Json.Ler(r.Corpo)).Str("access_token");
            if (TokensMudaram != null) TokensMudaram(Conexao);
            return true;
        }

        /// <summary>Chamada autenticada; um 401 tenta renovar o token uma vez.</summary>
        private JObj Chamar(string metodo, string caminho, JObj corpo)
        {
            string json = corpo == null ? null : Json.Escrever(corpo, false);
            RespostaHttp r = _http.Enviar(metodo, Conexao.Base + caminho, json, Conexao.Token, TimeoutMs);
            if (r.Status == 401 && Renovar()) r = _http.Enviar(metodo, Conexao.Base + caminho, json, Conexao.Token, TimeoutMs);
            if (r.Status != 200) throw Erro(r);
            JObj o = Json.Ler(r.Corpo) as JObj;
            if (o == null) throw new ErroAplicacao(r.Status, "resposta inesperada da aplicação");
            return o;
        }

        private static string Q(string projeto) { return Uri.EscapeDataString((projeto ?? "").Trim()); }

        public VersaoAplicacao Versao(string projeto)
        {
            return VersaoAplicacao.De(Chamar("GET", "/api/plugin/versao?projeto=" + Q(projeto), null));
        }

        /// <summary>Resposta crua de /api/plugin/regras (guardada em cache pela camada do CAD).</summary>
        public JObj RegrasCruas(string projeto)
        {
            return Chamar("GET", "/api/plugin/regras?projeto=" + Q(projeto), null);
        }

        public ConferenciaBase Conferir(string projeto, JObj dados)
        {
            JObj corpo = new JObj().Set("projeto", projeto).Set("dados", dados);
            return ConferenciaBase.De(Chamar("POST", "/api/plugin/conferir", corpo));
        }

        /// <summary>Resposta de /api/plugin/regras → regras do projeto (origem = endereço da aplicação).</summary>
        public static RegrasProjeto RegrasDe(JObj resposta, string origem)
        {
            JObj p = resposta.Obj("projeto") ?? new JObj();
            string codigo = p.Str("codigo") ?? "";
            RegrasProjeto r = RegrasProjeto.MontarDeObjetos(codigo, "Aplicação: " + (p.Str("nome") ?? codigo) + " (" + codigo + ")", origem,
                (resposta.Lista("processamento") ?? new List<object>()).OfType<JObj>().ToList(),
                (resposta.Lista("classificacao") ?? new List<object>()).OfType<JObj>().ToList(),
                resposta.Obj("vinculacao"));
            r.Versao = VersaoAplicacao.De(resposta).Regras;
            r.DaAplicacao = true;
            return r;
        }
    }

    public class ItemConferido
    {
        public string Ativo;
        public bool NaoEncontrado;
        public bool ForaDoCalculo;
    }

    /// <summary>Resultado de /api/plugin/conferir: os ativos que cada linha gera e quais não estão na Base de Orçamento.</summary>
    public class ConferenciaBase
    {
        public VersaoAplicacao Versao;
        public Dictionary<string, List<ItemConferido>> PorLinha = new Dictionary<string, List<ItemConferido>>();
        public List<string> NaoEncontrados = new List<string>();
        public int ComPendencia;

        public static string Chave(string tabela, int indice) { return tabela + "#" + indice; }

        public static ConferenciaBase De(JObj o)
        {
            ConferenciaBase c = new ConferenciaBase();
            c.Versao = VersaoAplicacao.De(o);
            c.NaoEncontrados = Json.ListaDeTextos(o["nao_encontrados"]);
            foreach (JObj l in (o.Lista("linhas") ?? new List<object>()).OfType<JObj>())
            {
                List<ItemConferido> itens = new List<ItemConferido>();
                foreach (JObj x in (l.Lista("itens") ?? new List<object>()).OfType<JObj>())
                {
                    ItemConferido it = new ItemConferido();
                    it.Ativo = x.Str("ativo") ?? "";
                    it.NaoEncontrado = x.Bool("naoEncontrado", false);
                    it.ForaDoCalculo = x.Bool("foraDoCalculo", false);
                    itens.Add(it);
                }
                c.PorLinha[Chave(l.Str("tabela"), (int)(l.Num("indice") ?? -1))] = itens;
                if (itens.Any(i => i.NaoEncontrado)) c.ComPendencia++;
            }
            return c;
        }

        /// <summary>Leva o resultado para os elementos da obra (mesma ordem de linhas de ObraArquivo.Dados) e refaz a conferência.</summary>
        public void Aplicar(Obra obra, RegrasProjeto regras, string projetoNome)
        {
            List<Elemento> cabos, outros;
            ObraArquivo.ElementosDasLinhas(obra, out cabos, out outros);
            foreach (Elemento e in obra.Elementos) { e.ForaDaBase.Clear(); e.MensagensBase.Clear(); }
            Marcar(cabos, "cabos", projetoNome);
            Marcar(outros, "outros", projetoNome);
            Conferencia.Avaliar(obra, regras);
        }

        private void Marcar(List<Elemento> elementos, string tabela, string projetoNome)
        {
            for (int i = 0; i < elementos.Count; i++)
            {
                List<ItemConferido> itens;
                if (!PorLinha.TryGetValue(Chave(tabela, i), out itens)) continue;
                Elemento e = elementos[i];
                foreach (ItemConferido it in itens.Where(x => x.NaoEncontrado))
                {
                    string nome = Anotacoes.NomeBase(it.Ativo);
                    if (e.ForaDaBase.Contains(nome)) continue;
                    e.ForaDaBase.Add(nome);
                    e.MensagensBase.Add(it.Ativo + " não está na base de orçamento" + (string.IsNullOrEmpty(projetoNome) ? "" : " (" + projetoNome + ")")
                                        + (it.ForaDoCalculo ? " — fica fora do orçamento" : ""));
                }
            }
        }
    }
}
