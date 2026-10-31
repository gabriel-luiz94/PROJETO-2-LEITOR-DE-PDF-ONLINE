using System.Collections.Generic;
using OrcamentoNanoCAD.Nucleo;

namespace OrcamentoNanoCAD.Testes
{
    /// <summary>O desenho dos mockups (referencias/mockups/C_ler.png): 8 postes, 7 vãos, regras de Pernambuco.
    /// Coordenadas da tela do mockup com o Y invertido (no CAD o Y cresce para cima).</summary>
    public static class DesenhoExemplo
    {
        public const double H = 8;   // altura dos textos

        [System.ThreadStatic] private static int _h;   // por thread: o xUnit roda as classes de teste em paralelo
        private static string Handle() { return (_h++).ToString("X"); }

        public static TextoDesenho Texto(string t, double x, double y, string layer, string cor = "#000000")
        {
            return new TextoDesenho { Handle = Handle(), Texto = t, X = x, Y = -y, Layer = layer, Cor = cor, Altura = H };
        }

        public static Desenho Montar()
        {
            _h = 0x100;   // handles iguais a cada montagem (o exemplo versionado depende disso)
            Desenho d = new Desenho { Arquivo = @"C:\Projetos\Projeto_Rua_das_Flores.dwg" };
            var postes = new[] {
                new { x = 105.0, y = 305.0, t = "9200D, N1" },
                new { x = 325.0, y = 275.0, t = "9400D, N1, N4" },
                new { x = 555.0, y = 335.0, t = "11300D, N3" },
                new { x = 785.0, y = 285.0, t = "9400D, N1" },
                new { x = 325.0, y = 525.0, t = "10300R, N1" },
                new { x = 565.0, y = 575.0, t = "9300F, N3" },
                new { x = 805.0, y = 535.0, t = "9200D, N1" },
                new { x = 565.0, y = 745.0, t = "7150FV, B2" },
            };
            // a ordem de inserção é bagunçada de propósito: a numeração tem de seguir a rede, não a ordem do DWG
            int[] ordem = { 6, 2, 0, 7, 4, 1, 5, 3 };
            foreach (int i in ordem)
            {
                d.Postes.Add(new PosteDesenho { Handle = Handle(), X = postes[i].x, Y = -postes[i].y, Bloco = "POSTE", Layer = "POSTES" });
                d.Textos.Add(Texto(postes[i].t, postes[i].x - 22, postes[i].y - 12, "01_INSTALANDO"));
            }
            Rede(d, postes[0].x, postes[0].y, postes[1].x, postes[1].y, "MT", "3xA35P(7,94)-13,8kV 45 m", "01_INSTALANDO");
            Rede(d, postes[1].x, postes[1].y, postes[2].x, postes[2].y, "MT", "3xA35P(7,94)-13,8kV 45 m", "01_INSTALANDO");
            Rede(d, postes[2].x, postes[2].y, postes[3].x, postes[3].y, "MT", "3xS04-7,97kV 38 m", "01_INSTALANDO");
            Rede(d, postes[1].x, postes[1].y, postes[4].x, postes[4].y, "MT", "3xA35P(7,94)-13,8kV 52 m", "01_REMOVER");
            Rede(d, postes[4].x, postes[4].y, postes[5].x, postes[5].y, "BT", "3xC16(16) 30 m", "01_INSTALANDO");
            Rede(d, postes[5].x, postes[5].y, postes[6].x, postes[6].y, "BT", "3xC16(16) 28 m", "01_INSTALANDO");
            Rede(d, postes[5].x, postes[5].y, postes[7].x, postes[7].y, "BT", "RAMA25 20 m", "01_REAPLICANDO");
            d.Textos.Add(Texto("Rua das Flores", 450, 420, "TXT_RUAS"));
            return d;
        }

        private static void Rede(Desenho d, double x1, double y1, double x2, double y2, string tipo, string anotacao, string layerTexto)
        {
            LinhaDesenho l = new LinhaDesenho
            {
                Handle = Handle(),
                Layer = tipo == "MT" ? "REDE_MT" : "REDE_BT",
                Cor = tipo == "MT" ? "#ff0000" : "#00ff00",
                Pontos = new List<Ponto> { new Ponto(x1, -y1), new Ponto(x2, -y2) },
            };
            d.Linhas.Add(l);
            double mx = (x1 + x2) / 2, my = (y1 + y2) / 2;
            bool vertical = System.Math.Abs(x2 - x1) < System.Math.Abs(y2 - y1);
            d.Textos.Add(vertical ? Texto(anotacao, mx - 15, my, layerTexto) : Texto(anotacao, mx, my - 12, layerTexto));
        }
    }
}
