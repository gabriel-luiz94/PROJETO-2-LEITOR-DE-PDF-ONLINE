// Plugin/Icones.cs — ícones DESENHADOS (GDI+), para não depender da fonte de emoji do Windows/nanoCAD (que some ou sai em
// preto e branco em listas e botões). Dois usos:
//   • nas listas: Desenhar(g, área, "estado:OK" | "rede:MT" | "op:I" | "poste" | "avulso" | "cabos" | "outros")
//   • nos botões: Imagem("inicio" | "fechar" | "leitura" | ...) → bitmap 18×18 na cor do texto
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using OrcamentoNanoCAD.Nucleo;

namespace OrcamentoNanoCAD.UI
{
    public static class Icones
    {
        public static readonly Color Roxo = Color.FromArgb(0xa3, 0x71, 0xf7);
        public static readonly Color AzulPoste = Color.FromArgb(0x58, 0xa6, 0xff);
        private static readonly Font Letra = new Font("Segoe UI", 7.5f, FontStyle.Bold);
        private static readonly Dictionary<string, Bitmap> _cache = new Dictionary<string, Bitmap>();

        public static Color CorDaOperacao(string op)
        {
            string o = (op ?? "").Replace("*", "").Trim().ToUpperInvariant();
            if (o == "I") return Visual.Ok;
            if (o == "R") return Visual.Erro;
            if (o == "M") return Roxo;
            if (o.Length == 0) return Color.Empty;
            return Visual.Azul;
        }

        private static string NomeOperacao(string op)
        {
            string o = (op ?? "").Replace("*", "").Trim().ToUpperInvariant();
            if (o == "I") return "+ " + op;
            if (o == "R") return "− " + op;
            if (o == "M") return "↻ " + op;
            return op;
        }

        private static GraphicsPath Arredondado(RectangleF r, float raio)
        {
            GraphicsPath p = new GraphicsPath();
            float d = raio * 2;
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        /// <summary>Etiqueta colorida com texto (MT, BT, + I, − R).</summary>
        private static void Etiqueta(Graphics g, Rectangle area, string texto, Color cor)
        {
            SizeF tam = g.MeasureString(texto, Letra);
            float w = Math.Max(tam.Width + 8, 26), h = 16;
            RectangleF r = new RectangleF(area.X + 3, area.Y + (area.Height - h) / 2f, w, h);
            using (GraphicsPath p = Arredondado(r, 4))
            using (SolidBrush f = new SolidBrush(Color.FromArgb(60, cor)))
            using (Pen borda = new Pen(cor))
            {
                g.FillPath(f, p);
                g.DrawPath(borda, p);
            }
            using (SolidBrush t = new SolidBrush(cor))
            using (StringFormat sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                g.DrawString(texto, Letra, t, r, sf);
        }

        /// <summary>Desenha o ícone de uma célula. Devolve false se o código não é conhecido (aí a célula mostra só o texto).</summary>
        public static bool Desenhar(Graphics g, Rectangle area, string codigo)
        {
            if (string.IsNullOrEmpty(codigo)) return false;
            SmoothingMode antes = g.SmoothingMode;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            try
            {
                int lado = Math.Min(16, area.Height - 4);
                Rectangle q = new Rectangle(area.X + (Math.Min(area.Width, 28) - lado) / 2, area.Y + (area.Height - lado) / 2, lado, lado);
                if (codigo.StartsWith("estado:")) { Estado(g, q, codigo.Substring(7)); return true; }
                if (codigo.StartsWith("rede:"))
                {
                    string t = codigo.Substring(5);
                    Etiqueta(g, area, t.Length == 0 ? "?" : t, t == "MT" ? Visual.Erro : t == "BT" ? Visual.Ok : Visual.Suave);
                    return true;
                }
                if (codigo.StartsWith("op:"))
                {
                    string op = codigo.Substring(3);
                    if (op.Length > 0) Etiqueta(g, area, NomeOperacao(op), CorDaOperacao(op));
                    return true;
                }
                if (codigo == "poste" || codigo == "outros") { Poste(g, q, AzulPoste); return true; }
                if (codigo == "cabos") { Cabo(g, q); return true; }
                if (codigo == "avulso") { Avulso(g, q); return true; }
                return false;
            }
            finally { g.SmoothingMode = antes; }
        }

        private static void Estado(Graphics g, Rectangle q, string estado)
        {
            if (estado == Estados.Conferir)
            {
                PointF[] tri = { new PointF(q.X + q.Width / 2f, q.Y), new PointF(q.Right, q.Bottom), new PointF(q.X, q.Bottom) };
                using (SolidBrush f = new SolidBrush(Visual.Conferir)) g.FillPolygon(f, tri);
                using (Pen p = new Pen(Visual.Fundo, 2f))
                {
                    float x = q.X + q.Width / 2f;
                    g.DrawLine(p, x, q.Y + q.Height * 0.38f, x, q.Y + q.Height * 0.68f);
                    g.DrawLine(p, x, q.Y + q.Height * 0.8f, x, q.Y + q.Height * 0.88f);
                }
                return;
            }
            Color cor = estado == Estados.Erro ? Visual.Erro : Visual.Ok;
            using (SolidBrush f = new SolidBrush(cor)) g.FillEllipse(f, q);
            using (Pen p = new Pen(Color.White, 2f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            {
                float a = q.Width * 0.3f;
                if (estado == Estados.Erro)
                {
                    g.DrawLine(p, q.X + a, q.Y + a, q.Right - a, q.Bottom - a);
                    g.DrawLine(p, q.Right - a, q.Y + a, q.X + a, q.Bottom - a);
                }
                else
                    g.DrawLines(p, new[] { new PointF(q.X + q.Width * 0.27f, q.Y + q.Height * 0.52f), new PointF(q.X + q.Width * 0.44f, q.Y + q.Height * 0.7f),
                                           new PointF(q.X + q.Width * 0.75f, q.Y + q.Height * 0.32f) });
            }
        }

        /// <summary>Poste: haste vertical com cruzeta e base.</summary>
        private static void Poste(Graphics g, Rectangle q, Color cor)
        {
            using (Pen p = new Pen(cor, 2f) { StartCap = LineCap.Round, EndCap = LineCap.Round })
            {
                float x = q.X + q.Width / 2f;
                g.DrawLine(p, x, q.Y + 1, x, q.Bottom - 1);
                g.DrawLine(p, q.X + 2, q.Y + q.Height * 0.25f, q.Right - 2, q.Y + q.Height * 0.25f);
                g.DrawLine(p, x - q.Width * 0.3f, q.Bottom - 1, x + q.Width * 0.3f, q.Bottom - 1);
            }
            using (SolidBrush f = new SolidBrush(cor))
            {
                g.FillEllipse(f, q.X + 1, q.Y + q.Height * 0.25f - 4, 3, 3);
                g.FillEllipse(f, q.Right - 4, q.Y + q.Height * 0.25f - 4, 3, 3);
            }
        }

        /// <summary>Cabos: dois fios (MT vermelho, BT verde) entre dois pontos.</summary>
        private static void Cabo(Graphics g, Rectangle q)
        {
            using (Pen mt = new Pen(Visual.Erro, 2f))
            using (Pen bt = new Pen(Visual.Ok, 2f))
            {
                g.DrawBezier(mt, q.X, q.Y + 4, q.X + q.Width / 3f, q.Y + 8, q.Right - q.Width / 3f, q.Y + 8, q.Right, q.Y + 4);
                g.DrawBezier(bt, q.X, q.Y + 10, q.X + q.Width / 3f, q.Y + 14, q.Right - q.Width / 3f, q.Y + 14, q.Right, q.Y + 10);
            }
            using (SolidBrush f = new SolidBrush(Visual.Texto))
            {
                g.FillRectangle(f, q.X, q.Y + 1, 2, q.Height - 1);
                g.FillRectangle(f, q.Right - 2, q.Y + 1, 2, q.Height - 1);
            }
        }

        /// <summary>Anotação solta (sem poste/vão): quadrado tracejado com "?".</summary>
        private static void Avulso(Graphics g, Rectangle q)
        {
            using (Pen p = new Pen(Visual.Suave, 1.5f) { DashStyle = DashStyle.Dash })
                g.DrawRectangle(p, q.X + 1, q.Y + 1, q.Width - 2, q.Height - 2);
            using (SolidBrush t = new SolidBrush(Visual.Suave))
            using (StringFormat sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                g.DrawString("?", Letra, t, new RectangleF(q.X, q.Y, q.Width, q.Height), sf);
        }

        // ── ícones dos botões ──

        /// <summary>Bitmap 18×18 do ícone (guardado em cache).</summary>
        public static Bitmap Imagem(string nome, Color cor)
        {
            string chave = nome + "|" + cor.ToArgb();
            Bitmap b;
            if (_cache.TryGetValue(chave, out b)) return b;
            b = new Bitmap(18, 18);
            using (Graphics g = Graphics.FromImage(b))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                Botao(g, nome, cor, new Rectangle(1, 1, 16, 16));
            }
            _cache[chave] = b;
            return b;
        }

        private static void Botao(Graphics g, string nome, Color cor, Rectangle q)
        {
            using (Pen p = new Pen(cor, 1.7f) { StartCap = LineCap.Round, EndCap = LineCap.Round, LineJoin = LineJoin.Round })
            using (SolidBrush f = new SolidBrush(cor))
            {
                float cx = q.X + q.Width / 2f, cy = q.Y + q.Height / 2f;
                switch (nome)
                {
                    case "inicio":      // casa
                        g.DrawLines(p, new[] { new PointF(q.X + 1, cy), new PointF(cx, q.Y + 1), new PointF(q.Right - 1, cy) });
                        g.DrawLines(p, new[] { new PointF(q.X + 3, cy - 1), new PointF(q.X + 3, q.Bottom - 1), new PointF(q.Right - 3, q.Bottom - 1), new PointF(q.Right - 3, cy - 1) });
                        g.DrawRectangle(p, cx - 2, q.Bottom - 6, 4, 5);
                        break;
                    case "fechar":      // X
                        g.DrawLine(p, q.X + 3, q.Y + 3, q.Right - 3, q.Bottom - 3);
                        g.DrawLine(p, q.Right - 3, q.Y + 3, q.X + 3, q.Bottom - 3);
                        break;
                    case "leitura":     // controles deslizantes (configurar)
                        for (int i = 0; i < 3; i++)
                        {
                            float y = q.Y + 3 + i * 5;
                            g.DrawLine(p, q.X + 1, y, q.Right - 1, y);
                            float x = q.X + 4 + ((i * 5) % 9);
                            g.FillEllipse(f, x - 2.5f, y - 2.5f, 5, 5);
                        }
                        break;
                    case "diagnostico": // lupa
                        g.DrawEllipse(p, q.X + 1, q.Y + 1, 10, 10);
                        g.DrawLine(p, q.X + 10, q.Y + 10, q.Right - 1, q.Bottom - 1);
                        break;
                    case "reler":       // seta circular
                    case "sincronizar":
                        g.DrawArc(p, q.X + 2, q.Y + 2, 12, 12, 40, 280);
                        g.FillPolygon(f, new[] { new PointF(q.Right - 1, q.Y + 3), new PointF(q.Right - 1, q.Y + 9), new PointF(q.Right - 7, q.Y + 6) });
                        if (nome == "sincronizar") g.FillEllipse(f, cx - 1.5f, cy - 1.5f, 3, 3);
                        break;
                    case "atualizar":   // seta circular + ponto (mantém ajustes)
                        g.DrawArc(p, q.X + 2, q.Y + 2, 12, 12, 220, 280);
                        g.FillPolygon(f, new[] { new PointF(q.X + 1, q.Bottom - 3), new PointF(q.X + 1, q.Bottom - 9), new PointF(q.X + 7, q.Bottom - 6) });
                        g.FillEllipse(f, cx - 2, cy - 2, 4, 4);
                        break;
                    case "poste":
                        Poste(g, q, cor);
                        break;
                    case "vao":         // linha entre dois postes
                        g.DrawLine(p, q.X + 2, q.Y + 3, q.X + 2, q.Bottom - 1);
                        g.DrawLine(p, q.Right - 2, q.Y + 3, q.Right - 2, q.Bottom - 1);
                        g.DrawBezier(p, q.X + 2, q.Y + 5, cx - 2, q.Y + 10, cx + 2, q.Y + 10, q.Right - 2, q.Y + 5);
                        break;
                    case "vincular":    // elos de corrente
                    case "vinculos":
                    case "conectar":
                        g.DrawArc(p, q.X + 1, q.Y + 5, 9, 6, 90, 180);
                        g.DrawLine(p, q.X + 5.5f, q.Y + 5, q.X + 9, q.Y + 5);
                        g.DrawLine(p, q.X + 5.5f, q.Y + 11, q.X + 9, q.Y + 11);
                        g.DrawArc(p, q.Right - 10, q.Y + 5, 9, 6, 270, 180);
                        g.DrawLine(p, q.Right - 9, q.Y + 5, q.Right - 5.5f, q.Y + 5);
                        g.DrawLine(p, q.Right - 9, q.Y + 11, q.Right - 5.5f, q.Y + 11);
                        g.DrawLine(p, q.X + 6, cy, q.Right - 6, cy);
                        break;
                    case "salvar":      // disquete
                        g.DrawLines(p, new[] { new PointF(q.X + 1, q.Y + 1), new PointF(q.Right - 4, q.Y + 1), new PointF(q.Right - 1, q.Y + 4),
                                               new PointF(q.Right - 1, q.Bottom - 1), new PointF(q.X + 1, q.Bottom - 1), new PointF(q.X + 1, q.Y + 1) });
                        g.DrawRectangle(p, q.X + 4, q.Y + 1, 7, 4);
                        g.DrawRectangle(p, q.X + 4, q.Bottom - 7, 8, 6);
                        break;
                    case "app":         // globo
                        g.DrawEllipse(p, q.X + 1, q.Y + 1, 14, 14);
                        g.DrawEllipse(p, cx - 3, q.Y + 1, 6, 14);
                        g.DrawLine(p, q.X + 1, cy, q.Right - 1, cy);
                        break;
                    case "ativos":      // caixa
                        g.DrawPolygon(p, new[] { new PointF(cx, q.Y + 1), new PointF(q.Right - 1, q.Y + 4.5f), new PointF(q.Right - 1, q.Bottom - 4.5f),
                                                 new PointF(cx, q.Bottom - 1), new PointF(q.X + 1, q.Bottom - 4.5f), new PointF(q.X + 1, q.Y + 4.5f) });
                        g.DrawLines(p, new[] { new PointF(q.X + 1, q.Y + 4.5f), new PointF(cx, q.Y + 8), new PointF(q.Right - 1, q.Y + 4.5f) });
                        g.DrawLine(p, cx, q.Y + 8, cx, q.Bottom - 1);
                        break;
                    case "marcas":      // olho
                        g.DrawBezier(p, q.X + 1, cy, q.X + 5, q.Y + 3, q.Right - 5, q.Y + 3, q.Right - 1, cy);
                        g.DrawBezier(p, q.X + 1, cy, q.X + 5, q.Bottom - 3, q.Right - 5, q.Bottom - 3, q.Right - 1, cy);
                        g.FillEllipse(f, cx - 2.5f, cy - 2.5f, 5, 5);
                        break;
                    case "ok":
                    case "conferir":
                    case "erro":
                        Estado(g, new Rectangle(q.X + 1, q.Y + 1, 14, 14), nome == "ok" ? Estados.Ok : nome == "conferir" ? Estados.Conferir : Estados.Erro);
                        break;
                    case "ignorados":   // texto riscado
                        g.DrawLine(p, q.X + 2, q.Y + 4, q.Right - 2, q.Y + 4);
                        g.DrawLine(p, cx, q.Y + 4, cx, q.Bottom - 2);
                        g.DrawLine(p, q.X + 1, q.Bottom - 1, q.Right - 1, q.Y + 1);
                        break;
                    case "todos":       // lista
                        for (int i = 0; i < 3; i++)
                        {
                            float y = q.Y + 3 + i * 5;
                            g.FillEllipse(f, q.X + 1, y - 1.5f, 3, 3);
                            g.DrawLine(p, q.X + 6, y, q.Right - 1, y);
                        }
                        break;
                    case "cabos":
                        Cabo(g, q);
                        break;
                }
            }
        }
    }
}
