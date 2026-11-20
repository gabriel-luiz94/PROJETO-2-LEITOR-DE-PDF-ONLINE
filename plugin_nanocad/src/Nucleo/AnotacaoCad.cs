using System;

namespace OrcamentoNanoCAD.Nucleo
{
    public static class AnotacaoCad
    {
        public static string AlternarLinhaViva(string operacao)
        {
            switch (operacao)
            {
                case "I": case "R": case "M": return "*" + operacao;
                case "*I": case "*R": case "*M": return operacao.Substring(1);
                default: return null;
            }
        }

        public static Ponto CentroDaBorda(Ponto minimo, Ponto maximo, Ponto alvo)
        {
            Ponto centro = new Ponto((minimo.X + maximo.X) / 2, (minimo.Y + maximo.Y) / 2);
            double dx = alvo.X - centro.X, dy = alvo.Y - centro.Y;
            double taxaX = dx == 0 ? double.PositiveInfinity : (maximo.X - centro.X) / Math.Abs(dx);
            double taxaY = dy == 0 ? double.PositiveInfinity : (maximo.Y - centro.Y) / Math.Abs(dy);
            double taxa = Math.Min(1, Math.Min(taxaX, taxaY));
            return new Ponto(centro.X + taxa * dx, centro.Y + taxa * dy);
        }
    }
}
