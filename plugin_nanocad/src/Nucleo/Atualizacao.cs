// Nucleo/Atualizacao.cs — atualização automática do plugin (TASK-013 do plugin / TASK-130 do principal).
// A aplicação publica a DLL do repositório em GET /api/plugin/atualizacao (versão + SHA-256) e .../arquivo (a DLL).
// O nanoCAD trava a DLL carregada e o .NET não a descarrega; mas o Windows deixa RENOMEAR um arquivo em uso. Então:
//   OrcamentoNanoCAD.dll (em uso) → OrcamentoNanoCAD.dll.anterior;   nova → OrcamentoNanoCAD.dll
// e a versão nova vale na próxima vez que o nanoCAD abrir (o APPLOAD carrega o mesmo caminho de sempre).
// Se o rename não for possível, a troca é feita por um processo auxiliar quando o nanoCAD fechar (camada do CAD).
// Sem dependência do CAD: os testes rodam no Linux.
using System;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;

namespace OrcamentoNanoCAD.Nucleo
{
    public class PacoteAtualizacao
    {
        public string Alvo, Versao, Sha256, Url;
        public long Tamanho;

        public static PacoteAtualizacao De(JObj o)
        {
            if (o == null) return null;
            PacoteAtualizacao p = new PacoteAtualizacao();
            p.Alvo = o.Str("alvo");
            p.Versao = o.Str("versao");
            p.Sha256 = (o.Str("sha256") ?? "").ToLowerInvariant();
            p.Url = o.Str("url");
            p.Tamanho = (long)(o.Num("tamanho") ?? 0);
            if (string.IsNullOrEmpty(p.Versao) || p.Sha256.Length != 64 || string.IsNullOrEmpty(p.Url)) return null;
            return p;
        }
    }

    /// <summary>Download de arquivo binário (separado de ITransporteHttp, que é texto/JSON).</summary>
    public interface IBaixador
    {
        byte[] Baixar(string url, int timeoutMs, out int status, out string erro);
    }

    public class BaixadorHttp : IBaixador
    {
        public byte[] Baixar(string url, int timeoutMs, out int status, out string erro)
        {
            status = 0; erro = null;
            try
            {
                new TransporteHttp();      // garante o TLS 1.2 ligado (construtor estático)
#pragma warning disable SYSLIB0014
                HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
#pragma warning restore SYSLIB0014
                req.Method = "GET";
                req.Timeout = timeoutMs;
                req.ReadWriteTimeout = timeoutMs;
                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                using (Stream s = resp.GetResponseStream())
                using (MemoryStream m = new MemoryStream())
                {
                    status = (int)resp.StatusCode;
                    byte[] buf = new byte[81920];
                    int n;
                    while ((n = s.Read(buf, 0, buf.Length)) > 0)
                    {
                        m.Write(buf, 0, n);
                        if (m.Length > Atualizacao.TamanhoMaximo) { erro = "arquivo grande demais"; return null; }
                    }
                    return m.ToArray();
                }
            }
            catch (WebException e)
            {
                HttpWebResponse r = e.Response as HttpWebResponse;
                if (r != null) { status = (int)r.StatusCode; r.Close(); }
                erro = e.Message;
            }
            catch (Exception e) { erro = e.Message; }
            return null;
        }
    }

    public enum ResultadoTroca { Trocada, PrecisaAuxiliar, Erro }

    public static class Atualizacao
    {
        public const string ArquivoDll = "OrcamentoNanoCAD.dll";
        public const long TamanhoMaximo = 50L * 1024 * 1024;

        /// <summary>Compara "0.5.0" com "0.10.1" número a número (sufixos depois de "-" são ignorados). &lt;0, 0, &gt;0.</summary>
        public static int Comparar(string a, string b)
        {
            string[] x = Partes(a), y = Partes(b);
            for (int i = 0; i < Math.Max(x.Length, y.Length); i++)
            {
                int va = i < x.Length ? Numero(x[i]) : 0, vb = i < y.Length ? Numero(y[i]) : 0;
                if (va != vb) return va < vb ? -1 : 1;
            }
            return 0;
        }

        private static string[] Partes(string v)
        {
            string s = (v ?? "").Trim().TrimStart('v', 'V');
            int h = s.IndexOf('-');
            if (h >= 0) s = s.Substring(0, h);
            return s.Split('.');
        }

        private static int Numero(string s)
        {
            int n;
            return int.TryParse(s, out n) ? n : 0;
        }

        public static string Sha256(byte[] dados)
        {
            using (SHA256 h = SHA256.Create())
            {
                byte[] r = h.ComputeHash(dados);
                StringBuilder sb = new StringBuilder();
                foreach (byte b in r) sb.Append(b.ToString("x2"));
                return sb.ToString();
            }
        }

        /// <summary>Só baixa da aplicação configurada por HTTPS, ou por HTTP no próprio computador (localhost).</summary>
        public static bool OrigemPermitida(string baseUrl)
        {
            Uri u;
            if (!Uri.TryCreate((baseUrl ?? "").Trim(), UriKind.Absolute, out u)) return false;
            if (u.Scheme == "https") return true;
            if (u.Scheme != "http") return false;
            string h = u.Host.ToLowerInvariant();
            return h == "localhost" || h == "127.0.0.1" || h == "[::1]" || h == "::1";
        }

        /// <summary>Endereço absoluto do arquivo: o "url" da resposta é relativo à aplicação e tem de ficar NELA.</summary>
        public static string UrlDoArquivo(string baseUrl, string url)
        {
            Uri b = new Uri(baseUrl.TrimEnd('/') + "/");
            Uri r = new Uri(b, url ?? "");
            if (!string.Equals(r.Scheme, b.Scheme, StringComparison.OrdinalIgnoreCase) || !string.Equals(r.Authority, b.Authority, StringComparison.OrdinalIgnoreCase))
                return null;
            return r.ToString();
        }

        /// <summary>Consulta a aplicação. Devolve o pacote se houver versão MAIS NOVA que a instalada (e não ignorada); senão null + motivo.</summary>
        public static PacoteAtualizacao Consultar(ITransporteHttp http, string baseUrl, string alvo, string versaoAtual, string versaoIgnorada, out string motivo)
        {
            motivo = null;
            if (!OrigemPermitida(baseUrl)) { motivo = "endereço da aplicação não é HTTPS nem localhost: atualização desligada por segurança"; return null; }
            RespostaHttp r = http.Enviar("GET", baseUrl.TrimEnd('/') + "/api/plugin/atualizacao?alvo=" + Uri.EscapeDataString(alvo), null, null, 15000);
            if (r.Status == 0) { motivo = "aplicação fora do ar (" + r.Erro + ")"; return null; }
            if (r.Status == 404) { motivo = "a aplicação não publica o plugin para " + alvo; return null; }
            if (r.Status != 200) { motivo = "a aplicação respondeu " + r.Status; return null; }
            PacoteAtualizacao p = PacoteAtualizacao.De(Json.Ler(r.Corpo) as JObj);
            if (p == null) { motivo = "resposta de atualização incompleta"; return null; }
            if (Comparar(p.Versao, versaoAtual) <= 0) { motivo = "já está na versão mais nova (" + versaoAtual + ")"; return null; }
            if (!string.IsNullOrEmpty(versaoIgnorada) && Comparar(p.Versao, versaoIgnorada) <= 0)
            {
                motivo = "versão " + p.Versao + " foi desfeita neste computador (ORC_VOLTAR_VERSAO); espera a próxima";
                return null;
            }
            return p;
        }

        /// <summary>Baixa e confere (tamanho e SHA-256). null + motivo se algo não bate.</summary>
        public static byte[] Baixar(IBaixador baixador, string baseUrl, PacoteAtualizacao p, out string motivo)
        {
            motivo = null;
            string url = UrlDoArquivo(baseUrl, p.Url);
            if (url == null) { motivo = "o arquivo apontado está fora da aplicação configurada"; return null; }
            int status;
            string erro;
            byte[] dados = baixador.Baixar(url, 120000, out status, out erro);
            if (dados == null || status != 200) { motivo = "download falhou (" + (erro ?? ("status " + status)) + ")"; return null; }
            if (p.Tamanho > 0 && dados.Length != p.Tamanho) { motivo = "tamanho do arquivo não confere"; return null; }
            if (Sha256(dados) != p.Sha256) { motivo = "o arquivo baixado não confere com a impressão digital (SHA-256)"; return null; }
            if (dados.Length < 2 || dados[0] != (byte)'M' || dados[1] != (byte)'Z') { motivo = "o arquivo baixado não é uma DLL"; return null; }
            return dados;
        }

        /// <summary>Grava a versão nova ao lado da DLL em uso e troca: atual → .anterior, nova → atual. Se a troca não puder ser feita
        /// agora (arquivo travado), deixa "&lt;dll&gt;.nova" pronta e devolve PrecisaAuxiliar.</summary>
        public static ResultadoTroca Instalar(string dllAtual, byte[] dados, out string motivo)
        {
            motivo = null;
            string nova = dllAtual + ".nova", anterior = dllAtual + ".anterior";
            try
            {
                File.WriteAllBytes(nova, dados);
            }
            catch (Exception e) { motivo = "não deu para gravar na pasta do plugin (" + e.Message + ")"; return ResultadoTroca.Erro; }
            try
            {
                if (File.Exists(anterior)) File.Delete(anterior);
                File.Move(dllAtual, anterior);
            }
            catch (Exception e) { motivo = e.Message; return ResultadoTroca.PrecisaAuxiliar; }
            try
            {
                File.Move(nova, dllAtual);
                return ResultadoTroca.Trocada;
            }
            catch (Exception e)
            {
                // volta como estava para nunca deixar o APPLOAD sem DLL
                try { if (!File.Exists(dllAtual) && File.Exists(anterior)) File.Move(anterior, dllAtual); } catch (Exception) { }
                motivo = e.Message;
                return ResultadoTroca.PrecisaAuxiliar;
            }
        }

        /// <summary>Volta para a versão anterior (atual → .desfeita, .anterior → atual). false + motivo se não houver anterior.</summary>
        public static bool Reverter(string dllAtual, out string motivo)
        {
            motivo = null;
            string anterior = dllAtual + ".anterior", desfeita = dllAtual + ".desfeita";
            if (!File.Exists(anterior)) { motivo = "não há versão anterior guardada"; return false; }
            try
            {
                if (File.Exists(desfeita)) File.Delete(desfeita);
                File.Move(dllAtual, desfeita);
                File.Move(anterior, dllAtual);
                return true;
            }
            catch (Exception e)
            {
                try { if (!File.Exists(dllAtual) && File.Exists(desfeita)) File.Move(desfeita, dllAtual); } catch (Exception) { }
                motivo = e.Message;
                return false;
            }
        }

        /// <summary>Restos de trocas anteriores que já podem ir embora (chamado ao abrir: a DLL em uso é a atual).</summary>
        public static void Limpar(string dllAtual)
        {
            foreach (string s in new[] { ".desfeita", ".nova" })
                try { if (File.Exists(dllAtual + s)) File.Delete(dllAtual + s); } catch (Exception) { }
        }
    }
}
