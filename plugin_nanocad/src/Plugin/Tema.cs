// Plugin/Tema.cs — visual do plugin com a PALETA DA APLICAÇÃO (static/styles.css: fundo #0d1117, cartões #161b22, bordas #30363d,
// destaque laranja #d97706, texto #e6edf3 / #8b949e; estados: verde #3fb950, âmbar #d29922, vermelho #f85149) — TASK-009 do plugin.
// Também: disparo confiável de comandos a partir das janelas (TASK-007) e o log do plugin.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OrcamentoNanoCAD.Nucleo;

namespace OrcamentoNanoCAD.UI
{
    public static class Visual
    {
        public static readonly Color Fundo = Color.FromArgb(0x0d, 0x11, 0x17);
        public static readonly Color Cartao = Color.FromArgb(0x16, 0x1b, 0x22);
        public static readonly Color Realce = Color.FromArgb(0x21, 0x26, 0x2d);
        public static readonly Color Borda = Color.FromArgb(0x30, 0x36, 0x3d);
        public static readonly Color Destaque = Color.FromArgb(0xd9, 0x77, 0x06);
        public static readonly Color DestaqueEscuro = Color.FromArgb(0xb4, 0x53, 0x09);
        public static readonly Color Texto = Color.FromArgb(0xe6, 0xed, 0xf3);
        public static readonly Color Suave = Color.FromArgb(0x8b, 0x94, 0x9e);
        public static readonly Color Ok = Color.FromArgb(0x3f, 0xb9, 0x50);
        public static readonly Color Conferir = Color.FromArgb(0xd2, 0x99, 0x22);
        public static readonly Color Erro = Color.FromArgb(0xf8, 0x51, 0x49);
        public static readonly Color Azul = Color.FromArgb(0x1f, 0x6f, 0xeb);

        private static readonly string Familia = EscolherFamilia();
        public static readonly Font Titulo = new Font(Familia, 12f, FontStyle.Bold);
        public static readonly Font Normal = new Font(Familia, 9f);
        public static readonly Font Mono = new Font("Consolas", 9.5f);
        public static readonly Font Secao = new Font(Familia, 8f, FontStyle.Bold);
        public static readonly Font Grande = new Font(Familia, 11f, FontStyle.Bold);

        private static string EscolherFamilia()
        {
            try
            {
                foreach (FontFamily f in FontFamily.Families) if (f.Name == "Inter") return "Inter";       // a fonte da aplicação, se instalada
            }
            catch (Exception) { }
            return "Segoe UI";
        }

        public const string Primario = "primario";

        public static Color DoEstado(string estado)
        {
            if (estado == Estados.Erro) return Erro;
            if (estado == Estados.Conferir) return Conferir;
            return Ok;
        }

        public static Label Rotulo(string texto, Font fonte, Color? cor = null)
        {
            Label l = new Label { Text = texto, Font = fonte ?? Normal, AutoSize = true, Margin = new Padding(0, 2, 0, 2), ForeColor = cor ?? Texto, BackColor = Color.Transparent };
            l.Tag = cor.HasValue ? "cor" : null;
            return l;
        }

        public static Button Botao(string texto, EventHandler acao, bool primario = false, string icone = null)
        {
            Button b = new Button { Text = texto, AutoSize = true, Font = Normal, Margin = new Padding(3), Padding = new Padding(6, 3, 6, 3), Cursor = Cursors.Hand };
            if (primario) b.Tag = Primario;
            if (icone != null)
            {
                b.Image = Icones.Imagem(icone, primario ? Color.White : Texto);
                b.ImageAlign = ContentAlignment.MiddleLeft;
                b.TextImageRelation = TextImageRelation.ImageBeforeText;
                if (string.IsNullOrEmpty(texto)) b.Padding = new Padding(3);
            }
            b.Click += acao;
            Estilo(b);
            return b;
        }

        /// <summary>Botão só com ícone (o nome aparece na dica ao passar o mouse).</summary>
        public static Button BotaoIcone(string icone, string dica, EventHandler acao, ToolTip dicas)
        {
            Button b = Botao("", acao, false, icone);
            b.AutoSize = false;
            b.Size = new Size(32, 30);
            b.ImageAlign = ContentAlignment.MiddleCenter;
            if (dicas != null) dicas.SetToolTip(b, dica);
            return b;
        }

        /// <summary>Faixa de título de uma lista do painel: ícone desenhado + TÍTULO + contagem (como os cabeçalhos do Resumo da aplicação).</summary>
        public static Panel Cabecalho(string icone, Func<string> texto)
        {
            Panel p = new Panel { Dock = DockStyle.Top, Height = 30, BackColor = Cartao, Tag = "cartao" };
            p.Paint += (o, e) =>
            {
                e.Graphics.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                Icones.Desenhar(e.Graphics, new Rectangle(6, 0, 28, p.Height), icone);
                TextRenderer.DrawText(e.Graphics, texto(), Grande, new Rectangle(36, 0, p.Width - 40, p.Height), Texto,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
                using (Pen b = new Pen(Destaque, 2)) e.Graphics.DrawLine(b, 0, p.Height - 1, p.Width, p.Height - 1);
            };
            return p;
        }

        private static void Estilo(Button b)
        {
            bool prim = (b.Tag as string) == Primario;
            b.FlatStyle = FlatStyle.Flat;
            b.UseVisualStyleBackColor = false;
            b.BackColor = prim ? Destaque : Realce;
            b.ForeColor = prim ? Color.White : Texto;
            b.FlatAppearance.BorderColor = prim ? Destaque : Borda;
            b.FlatAppearance.BorderSize = 1;
            b.FlatAppearance.MouseOverBackColor = prim ? DestaqueEscuro : Borda;
            b.FlatAppearance.MouseDownBackColor = prim ? DestaqueEscuro : Fundo;
        }

        /// <summary>Aplica o tema na janela inteira (chamar no fim do construtor).</summary>
        public static void Aplicar(Control raiz)
        {
            if (raiz is Form)
            {
                raiz.BackColor = Fundo;
                raiz.ForeColor = Texto;
                ((Form)raiz).HandleCreated += (o, e) => BarraEscura((Form)o);
            }
            foreach (Control c in raiz.Controls) Controle(c);
        }

        private static void Controle(Control c)
        {
            if (c is Button) Estilo((Button)c);
            else if (c is Label) { if ((c.Tag as string) != "cor") c.ForeColor = Texto; c.BackColor = Color.Transparent; }
            else if (c is TextBox) { TextBox t = (TextBox)c; t.BackColor = Fundo; t.ForeColor = Texto; t.BorderStyle = BorderStyle.FixedSingle; }
            else if (c is NumericUpDown) { c.BackColor = Fundo; c.ForeColor = Texto; ((NumericUpDown)c).BorderStyle = BorderStyle.FixedSingle; }
            else if (c is ComboBox) { ComboBox cb = (ComboBox)c; cb.FlatStyle = FlatStyle.Flat; cb.BackColor = Fundo; cb.ForeColor = Texto; }
            else if (c is RadioButton && ((RadioButton)c).Appearance == Appearance.Button)
            {
                RadioButton r = (RadioButton)c;
                r.FlatStyle = FlatStyle.Flat;
                r.BackColor = Realce;
                if ((r.Tag as string) == null || r.ForeColor == SystemColors.ControlText) r.ForeColor = Texto;
                r.FlatAppearance.BorderColor = Borda;
                r.FlatAppearance.CheckedBackColor = Color.FromArgb(0x3a, 0x2a, 0x10);
                r.FlatAppearance.MouseOverBackColor = Borda;
            }
            else if (c is CheckBox || c is RadioButton) { c.ForeColor = Texto; c.BackColor = Color.Transparent; }
            else if (c is ListView) Lista((ListView)c);
            else if (c is Panel || c is TableLayoutPanel || c is FlowLayoutPanel) c.BackColor = c.Tag as string == "cartao" ? Cartao : Color.Transparent;
            foreach (Control f in c.Controls) Controle(f);
        }

        public static Panel Cartao_(Control conteudo)
        {
            Panel p = new Panel { AutoSize = true, Padding = new Padding(10), Margin = new Padding(6), BackColor = Cartao, Tag = "cartao" };
            p.Paint += (o, e) => { using (Pen pen = new Pen(Borda)) e.Graphics.DrawRectangle(pen, 0, 0, p.Width - 1, p.Height - 1); };
            p.Controls.Add(conteudo);
            return p;
        }

        // ── lista com cabeçalho escuro (o padrão do Windows ignora a cor do cabeçalho) ──
        public static ListView Lista(string[] colunas)
        {
            ListView lv = new ListView { View = View.Details, FullRowSelect = true, GridLines = false, MultiSelect = false, HideSelection = false, Font = Normal, Dock = DockStyle.Fill };
            foreach (string c in colunas) lv.Columns.Add(c);
            lv.SmallImageList = new ImageList { ImageSize = new Size(1, 24) };      // linhas mais altas (cabem os ícones)
            Lista(lv);
            return lv;
        }

        private static void Lista(ListView lv)
        {
            lv.BackColor = Cartao;
            lv.ForeColor = Texto;
            lv.BorderStyle = BorderStyle.None;
            if (lv.OwnerDraw) return;
            lv.OwnerDraw = true;
            lv.DrawColumnHeader += (o, e) =>
            {
                using (SolidBrush f = new SolidBrush(Fundo)) e.Graphics.FillRectangle(f, e.Bounds);
                using (Pen p = new Pen(Borda)) e.Graphics.DrawLine(p, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
                TextRenderer.DrawText(e.Graphics, e.Header.Text, Secao, new Rectangle(e.Bounds.X + 4, e.Bounds.Y, e.Bounds.Width - 4, e.Bounds.Height), Suave,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
            };
            lv.DrawItem += (o, e) => { };
            lv.DrawSubItem += (o, e) =>
            {
                bool sel = e.Item.Selected;
                using (SolidBrush f = new SolidBrush(sel ? Color.FromArgb(0x3a, 0x2a, 0x10) : (e.ItemIndex % 2 == 0 ? Cartao : Realce))) e.Graphics.FillRectangle(f, e.Bounds);
                Color cor = e.Item.ForeColor == lv.ForeColor || e.Item.ForeColor.IsEmpty ? Texto : e.Item.ForeColor;
                if (e.ColumnIndex == 0 && !sel) cor = e.Item.ForeColor == lv.ForeColor ? Texto : cor;
                string icone = e.SubItem.Tag as string;
                if (icone != null && Icones.Desenhar(e.Graphics, e.Bounds, icone)) return;
                TextRenderer.DrawText(e.Graphics, e.SubItem.Text, lv.Font, new Rectangle(e.Bounds.X + 4, e.Bounds.Y, e.Bounds.Width - 4, e.Bounds.Height), cor,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
            };
        }

        public static void Preencher(ListView lv, List<LinhaTabela> linhas)
        {
            lv.BeginUpdate();
            lv.Items.Clear();
            foreach (LinhaTabela l in linhas)
            {
                ListViewItem it = new ListViewItem(l.Celulas);
                it.Tag = l.Chave;
                it.ForeColor = l.Estado == Estados.Ok ? Texto : DoEstado(l.Estado);
                if (l.Icones != null)
                    for (int i = 0; i < l.Icones.Length && i < it.SubItems.Count; i++) it.SubItems[i].Tag = l.Icones[i];
                lv.Items.Add(it);
            }
            foreach (ColumnHeader c in lv.Columns) c.Width = -2;
            for (int i = 0; i < lv.Columns.Count; i++)
            {
                ColumnHeader c = lv.Columns[i];
                bool soIcone = c.Text.Length == 0;
                if (soIcone) c.Width = 30;                                  // coluna do estado (só o ícone)
                else if (c.Width < 60) c.Width = 60;
                if (c.Width > 320) c.Width = 320;
            }
            lv.EndUpdate();
        }

        // barra de título escura no Windows 10/11 (sem efeito em versões antigas)
        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int valor, int tamanho);

        private static void BarraEscura(Form f)
        {
            try { int um = 1; if (DwmSetWindowAttribute(f.Handle, 20, ref um, 4) != 0) DwmSetWindowAttribute(f.Handle, 19, ref um, 4); }
            catch (Exception) { }
        }

        // ── comandos a partir das janelas ──
        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hwnd);

        /// <summary>Manda um comando para o desenho ativo. Termina com Enter (o nanoCAD não aceita só o espaço vindo de janela) e devolve
        /// o foco ao nanoCAD, para os pedidos de seleção aparecerem no desenho. Com um comando em andamento, avisa em vez de atropelar.</summary>
        public static void Executar(string comando)
        {
            HostMgd.ApplicationServices.Document d = Aplicacao.Doc();
            if (d == null) return;
            string atual = "";
            try { atual = d.CommandInProgress ?? ""; } catch (Exception) { }
            if (atual.Length > 0 && !atual.StartsWith("ORC_", StringComparison.OrdinalIgnoreCase) && atual != "ORCAMENTO")
            {
                Aplicacao.Msg("Termine (ou cancele com Esc) o comando " + atual + " antes de usar o painel.");
                return;
            }
            Aplicacao.Log("painel → " + comando);
            try { SetForegroundWindow(Process.GetCurrentProcess().MainWindowHandle); } catch (Exception) { }
            d.SendStringToExecute(comando + "\n", true, false, true);
        }
    }
}
