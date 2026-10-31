// Nucleo/Modelo.cs — o que o plugin lê do desenho (Desenho*) e a obra que ele monta (Obra, Poste, Vao).
// Nada aqui depende do nanoCAD: a camada do CAD só preenche `Desenho` e desenha/mostra a `Obra`.
using System;
using System.Collections.Generic;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    // ── entrada: o desenho como o CAD entrega ──
    public class TextoDesenho
    {
        public string Handle;         // handle da entidade (TEXT, MTEXT, ATTRIB); parágrafos do mesmo MTEXT repetem o handle
        public int Parte;             // nº do parágrafo dentro do MTEXT (0 para TEXT)
        public string Texto;
        public string Cor = "#000000";
        public string Layer = "";
        public double X, Y;
        public double Altura;
        public string Tipo = "TEXT";  // TEXT | MTEXT | ATTRIB | BLOCO ("[BLOCO: nome]", como o leitor de DXF)
    }

    public class PosteDesenho
    {
        public string Handle;
        public double X, Y;
        public string Bloco;
        public string Layer;
    }

    public class LinhaDesenho
    {
        public string Handle;
        public string Layer = "";
        public string Cor = "#000000";
        public List<Ponto> Pontos = new List<Ponto>();
        public bool Leader;
    }

    public class Desenho
    {
        public string Arquivo;        // caminho do DWG (pode ser vazio: desenho novo)
        public List<TextoDesenho> Textos = new List<TextoDesenho>();
        public List<PosteDesenho> Postes = new List<PosteDesenho>();
        public List<LinhaDesenho> Linhas = new List<LinhaDesenho>();
    }

    // ── saída: a obra ──
    public static class Estados
    {
        public const string Ok = "OK";
        public const string Conferir = "CONFERIR";
        public const string Erro = "ERRO";

        public static string Pior(string a, string b)
        {
            if (a == Erro || b == Erro) return Erro;
            if (a == Conferir || b == Conferir) return Conferir;
            return Ok;
        }
    }

    /// <summary>Ponta de vão ligada a uma estrutura de um poste (K = posição da estrutura na anotação do poste, 0 = a 1ª).</summary>
    public class Vinculo
    {
        public string Poste;
        public string Estrutura;
        public int K;
        public bool Manual;
    }

    public abstract class Elemento
    {
        public string Numero;
        public string TextoHandle;    // anotação associada
        public string TextoOriginal = "";
        public string TextoLayer = "";
        public string TextoCor = "";
        public double? TextoX, TextoY;
        public string Entidade = "0";
        public string AtivoLido = ""; // ativo como o motor do leitor montou
        public string OperacaoLida = "";
        public string Operacao = "";  // pode ser ajustada no pop-up
        public bool Manual;           // montado em "Construir do zero" ou com anotação trocada à mão
        public string Estado = Estados.Ok;
        public List<string> Mensagens = new List<string>();
        public List<string> MensagensFixas = new List<string>();   // da leitura (não mudam na conferência)
        public List<string> ForaDaBase = new List<string>();       // ativos desta linha que a aplicação não achou na Base de Orçamento
        public List<string> MensagensBase = new List<string>();    // ... e as mensagens (vêm de /api/plugin/conferir)

        public bool TemAnotacao { get { return !string.IsNullOrEmpty(AtivoLido); } }
        public bool OperacaoAjustada { get { return Operacao != OperacaoLida; } }

        public void Marcar(string estado, string msg)
        {
            Estado = Estados.Pior(Estado, estado);
            if (!string.IsNullOrEmpty(msg) && !Mensagens.Contains(msg)) Mensagens.Add(msg);
        }
    }

    public class Poste : Elemento
    {
        public string Handle;         // bloco do poste (vazio se o poste foi criado só pela anotação)
        public double X, Y;
        public string Bloco = "";
        public PosteLido Lido = new PosteLido();
        // parâmetros calculados de cada estrutura (mesma ordem de Lido.Estruturas)
        public List<List<KeyValuePair<string, string>>> ParametrosEstruturas = new List<List<KeyValuePair<string, string>>>();
        public List<int> CabosPorEstrutura = new List<int>();
        public List<string> AvisosParametros = new List<string>();
        public Ponto Ponto { get { return new Ponto(X, Y); } }
    }

    public class Vao : Elemento
    {
        public string LinhaHandle;
        public int Trecho;            // a mesma linha da rede pode virar vários vãos (um por par de postes)
        public List<Ponto> Pontos = new List<Ponto>();
        public string Tipo = "";      // MT | BT (pela cor + layer da rede)
        public string LinhaLayer = "";
        public string LinhaCor = "";
        public string PosteA, PosteB; // números (P1…) ou null = ponta sem poste
        public Poste PosteObjA, PosteObjB;
        public double? MetrosLidos;
        public double? Metros;        // pode ser ajustado no pop-up
        public CaboLido Cabo;
        public InfoCabo InfoCabo;
        public Vinculo VinculoA, VinculoB;
        public string MotivoVinculo;  // por que a última tentativa de vínculo automático falhou

        /// <summary>Ativo que vai para a obra (com o comprimento ajustado, se houver).</summary>
        public string Ativo
        {
            get
            {
                if (!TemAnotacao) return "";
                if (Metros.HasValue && (!MetrosLidos.HasValue || Math.Abs(Metros.Value - MetrosLidos.Value) > 1e-9))
                    return Anotacoes.ComMetros(AtivoLido, Metros.Value);
                return AtivoLido;
            }
        }

        public bool ParticipaVinculo { get { return TemAnotacao && RegrasVinculo.ParticipaVinculo(Operacao); } }

        public IEnumerable<Vinculo> Vinculos
        {
            get
            {
                if (VinculoA != null) yield return VinculoA;
                if (VinculoB != null) yield return VinculoB;
            }
        }
    }

    /// <summary>Texto lido que não é cabo nem poste (vai para Outros como está, sem vínculo).</summary>
    public class Avulso : Elemento
    {
        public Elemento PertoDe;      // segunda anotação perto de um poste/vão que já tem a sua
    }

    public class Obra
    {
        public string Nome = "obra";
        public string Projeto = "081";
        public string NomeRegras = "";
        public string Desenho = "";
        public string Modo = "ler";   // ler | construir
        public string SalvoEm;        // data/hora (local) do último salvar
        public double AlturaTexto = 1; // altura típica dos textos do desenho (tamanho das marcas P1/V1 na tela)
        public List<Poste> Postes = new List<Poste>();
        public List<Vao> Vaos = new List<Vao>();
        public List<Avulso> Avulsos = new List<Avulso>();

        public Poste PostePorNumero(string n) { return n == null ? null : Postes.FirstOrDefault(p => p.Numero == n); }
        public Vao VaoPorNumero(string n) { return Vaos.FirstOrDefault(v => v.Numero == n); }

        public Elemento PorNumero(string n)
        {
            if (string.IsNullOrEmpty(n)) return null;
            Elemento e = PostePorNumero(n);
            if (e != null) return e;
            e = VaoPorNumero(n);
            if (e != null) return e;
            return Avulsos.FirstOrDefault(a => a.Numero == n);
        }

        public IEnumerable<Elemento> Elementos
        {
            get
            {
                foreach (Poste p in Postes) yield return p;
                foreach (Vao v in Vaos) yield return v;
                foreach (Avulso a in Avulsos) yield return a;
            }
        }

        public Elemento PorHandle(string handle)
        {
            if (string.IsNullOrEmpty(handle)) return null;
            foreach (Poste p in Postes) if (p.Handle == handle || p.TextoHandle == handle) return p;
            foreach (Vao v in Vaos) if (v.LinhaHandle == handle || v.TextoHandle == handle) return v;
            foreach (Avulso a in Avulsos) if (a.TextoHandle == handle) return a;
            return null;
        }
    }
}
