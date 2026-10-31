// Nucleo/Geometria.cs — pontos, distâncias e trechos de polilinha (sem CAD).
using System;
using System.Collections.Generic;

namespace OrcamentoNanoCAD.Nucleo
{
    public struct Ponto
    {
        public double X;
        public double Y;
        public Ponto(double x, double y) { X = x; Y = y; }
        public double Dist(Ponto o) { return Math.Sqrt((X - o.X) * (X - o.X) + (Y - o.Y) * (Y - o.Y)); }
        public override string ToString() { return Json.Numero(Math.Round(X, 2)) + "," + Json.Numero(Math.Round(Y, 2)); }
    }

    /// <summary>Projeção de um ponto numa polilinha: distância até ela e a "estação" (comprimento percorrido desde o início).</summary>
    public struct Projecao
    {
        public double Distancia;
        public double Estacao;
        public Ponto Pe;
    }

    public static class Geometria
    {
        public static Projecao ProjetarSegmento(Ponto p, Ponto a, Ponto b)
        {
            double dx = b.X - a.X, dy = b.Y - a.Y;
            double t = (dx == 0 && dy == 0) ? 0 : Math.Max(0, Math.Min(1, ((p.X - a.X) * dx + (p.Y - a.Y) * dy) / (dx * dx + dy * dy)));
            Projecao r = new Projecao();
            r.Pe = new Ponto(a.X + t * dx, a.Y + t * dy);
            r.Distancia = p.Dist(r.Pe);
            r.Estacao = t * Math.Sqrt(dx * dx + dy * dy);
            return r;
        }

        public static Projecao Projetar(Ponto p, IList<Ponto> pts)
        {
            Projecao melhor = new Projecao();
            melhor.Distancia = double.MaxValue;
            double acumulado = 0;
            for (int i = 0; i + 1 < pts.Count; i++)
            {
                Projecao pr = ProjetarSegmento(p, pts[i], pts[i + 1]);
                if (pr.Distancia < melhor.Distancia)
                {
                    melhor = pr;
                    melhor.Estacao = acumulado + pr.Estacao;
                }
                acumulado += pts[i].Dist(pts[i + 1]);
            }
            if (pts.Count == 1) { melhor.Distancia = p.Dist(pts[0]); melhor.Estacao = 0; melhor.Pe = pts[0]; }
            return melhor;
        }

        public static double Comprimento(IList<Ponto> pts)
        {
            double s = 0;
            for (int i = 0; i + 1 < pts.Count; i++) s += pts[i].Dist(pts[i + 1]);
            return s;
        }

        /// <summary>Ponto na estação `s` da polilinha.</summary>
        public static Ponto PontoNaEstacao(IList<Ponto> pts, double s)
        {
            if (pts.Count == 0) return new Ponto();
            double acumulado = 0;
            for (int i = 0; i + 1 < pts.Count; i++)
            {
                double l = pts[i].Dist(pts[i + 1]);
                if (acumulado + l >= s && l > 0)
                {
                    double t = (s - acumulado) / l;
                    return new Ponto(pts[i].X + t * (pts[i + 1].X - pts[i].X), pts[i].Y + t * (pts[i + 1].Y - pts[i].Y));
                }
                acumulado += l;
            }
            return pts[pts.Count - 1];
        }

        /// <summary>Pedaço da polilinha entre as estações s1 e s2 (s1 &lt; s2), com os vértices do meio.</summary>
        public static List<Ponto> Trecho(IList<Ponto> pts, double s1, double s2)
        {
            List<Ponto> r = new List<Ponto>();
            r.Add(PontoNaEstacao(pts, s1));
            double acumulado = 0;
            for (int i = 0; i + 1 < pts.Count; i++)
            {
                acumulado += pts[i].Dist(pts[i + 1]);
                if (acumulado > s1 && acumulado < s2 && i + 1 < pts.Count - 1) r.Add(pts[i + 1]);
            }
            r.Add(PontoNaEstacao(pts, s2));
            return r;
        }

        /// <summary>Ponto no meio do comprimento da polilinha.</summary>
        public static Ponto Meio(IList<Ponto> pts)
        {
            return PontoNaEstacao(pts, Comprimento(pts) / 2);
        }

        public static double Mediana(List<double> v)
        {
            if (v.Count == 0) return 0;
            List<double> c = new List<double>(v);
            c.Sort();
            return c[c.Count / 2];
        }

        /// <summary>Ângulo (rad, 0..π) entre duas direções.</summary>
        public static double Desvio(double dx1, double dy1, double dx2, double dy2)
        {
            double n1 = Math.Sqrt(dx1 * dx1 + dy1 * dy1), n2 = Math.Sqrt(dx2 * dx2 + dy2 * dy2);
            if (n1 == 0 || n2 == 0) return 0;
            double c = (dx1 * dx2 + dy1 * dy2) / (n1 * n2);
            return Math.Acos(Math.Max(-1, Math.Min(1, c)));
        }
    }
}
