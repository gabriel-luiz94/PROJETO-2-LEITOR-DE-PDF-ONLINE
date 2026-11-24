// Plugin/TelaCamposMapa.cs — TASK-038: janela do ponto/cabo no formato do mapa, no molde do «anota poste»:
// CAMPOS (L = linha do texto), cada campo com o ativo e a operação (F, I, *I, M, *M, R, *R); + Novo campo / + Nova linha; «Todos» por operação.
// Ponto: nome - comentário. Cabo (0.16.4): «Medido na linha» (só informa), «Comprimento usado» (começa no medido sem casas decimais) e
// um comprimento POR LINHA: a linha cuja anotação traz metragem («PB70 32 m», «PB70/32 m») fica com ela; as outras acompanham o usado.
// Ctrl+Enter confirma, Esc cancela. C# 5 (nanoCAD 5).
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using OrcamentoNanoCAD.Nucleo;

namespace OrcamentoNanoCAD.UI
{
    internal sealed class TelaCamposMapa : Form
    {
        private sealed class LinhaCampo
        {
            public Panel Painel;
            public NumericUpDown L;
            public TextBox Texto;
            public List<RadioButton> Ops = new List<RadioButton>();
            public Label Origem;
            public string Entidade;
            public TextBox M;                 // cabo: comprimento desta linha
            public bool Proprio;              // true = a linha tem o seu comprimento; false = acompanha o «Comprimento usado»
            public string Operacao { get { RadioButton r = Ops.FirstOrDefault(x => x.Checked); return r == null ? "I" : r.Text; } }
        }

        private readonly bool _cabo;
        private readonly FlowLayoutPanel _lista = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoScroll = true, FlowDirection = FlowDirection.TopDown, WrapContents = false };
        private readonly List<LinhaCampo> _linhas = new List<LinhaCampo>();
        private readonly TextBox _nome = new TextBox { Width = 150 };
        private readonly TextBox _comentario = new TextBox { Width = 520 };
        private readonly TextBox _metros = new TextBox { Width = 90 };
        private readonly Label _resumo;
        private readonly double _medido;
        private bool _sincronizando;

        public string Nome { get { return (_nome.Text ?? "").Trim(); } }
        public string Comentario { get { return (_comentario.Text ?? "").Trim(); } }
        public double? Metros
        {
            get
            {
                double v;
                string t = (_metros.Text ?? "").Trim().Replace(',', '.');
                return double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out v) && v > 0 ? (double?)v : null;
            }
        }
        public List<CampoMapa> Campos
        {
            get
            {
                return _linhas.Where(l => (l.Texto.Text ?? "").Trim().Length > 0).Select(l => new CampoMapa {
                    L = (int)l.L.Value, Texto = l.Texto.Text.Trim().ToUpperInvariant(), Operacao = l.Operacao,
                    Origem = (string)l.Origem.Tag ?? "manual", Entidade = l.Entidade, Metros = _cabo && l.Proprio ? Numero(l.M.Text) : null }).ToList();
            }
        }

        private static double? Numero(string texto)
        {
            double v;
            string t = (texto ?? "").Trim().Replace(',', '.');
            return double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out v) && v > 0 ? (double?)v : null;
        }

        /// <summary>`titulo` = «Ponto P12 · poste (bloco …)» ou «Cabo C3 · trecho 2 da polilinha». Para cabo: `medido` (m, na linha) e `usado`
        /// (o comprimento usado já escolhido; null = o medido sem casas decimais).</summary>
        public TelaCamposMapa(string titulo, bool cabo, List<CampoMapa> campos, string nome, string comentario, double medido, double? usado, string infoRegras, bool podeExcluir = false)
        {
            _cabo = cabo; _medido = medido;
            Text = titulo; Width = 960; Height = 600; StartPosition = FormStartPosition.CenterParent; KeyPreview = true; MinimizeBox = false;
            var topo = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(8), WrapContents = true };
            if (!cabo)
            {
                _nome.Text = nome ?? ""; _comentario.Text = comentario ?? "";
                topo.Controls.Add(Visual.Rotulo("Nome", Visual.Normal)); topo.Controls.Add(_nome);
                topo.Controls.Add(Visual.Rotulo("-", Visual.Grande)); topo.Controls.Add(Visual.Rotulo("Comentário", Visual.Normal)); topo.Controls.Add(_comentario);
            }
            else
            {
                _metros.Text = MapaManual.FormatarMetros(usado ?? MapaManual.UsadoPadrao(medido));
                topo.Controls.Add(Visual.Rotulo("Medido na linha: " + MapaManual.FormatarMetros(medido) + " m", Visual.Secao));
                topo.Controls.Add(Visual.Rotulo("     Comprimento usado (m)", Visual.Normal)); topo.Controls.Add(_metros);
                topo.Controls.Add(Visual.Rotulo("vai para as linhas sem metragem própria", Visual.Normal, Visual.Suave));
                _metros.TextChanged += (s, e) => SincronizarUsado();
            }
            if (!string.IsNullOrEmpty(infoRegras)) { Label r = Visual.Rotulo(infoRegras, Visual.Normal, Visual.Suave); topo.Controls.Add(r); topo.SetFlowBreak(r, true); }

            var barraCampos = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(8, 0, 8, 4) };
            barraCampos.Controls.Add(Visual.Rotulo("CAMPOS  (L = linha do texto)", Visual.Secao));
            barraCampos.Controls.Add(Visual.Rotulo("   Todos:", Visual.Normal, Visual.Suave));
            foreach (string op in MapaManual.Operacoes)
            {
                string o = op;
                Button b = Visual.Botao(o, (s, e) => { foreach (LinhaCampo l in _linhas) Marcar(l, o); AtualizarResumo(); });
                b.Width = 40; b.Padding = new Padding(0);
                barraCampos.Controls.Add(b);
            }

            var rodapeCampos = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(8, 4, 8, 4) };
            rodapeCampos.Controls.Add(Visual.Botao("+ Novo campo", (s, e) => { LinhaCampo u = _linhas.LastOrDefault(); Adicionar(new CampoMapa { L = u == null ? 1 : (int)u.L.Value, Operacao = u == null ? "I" : u.Operacao, Origem = "manual" }, true); }));
            rodapeCampos.Controls.Add(Visual.Botao("+ Nova linha", (s, e) => { LinhaCampo u = _linhas.LastOrDefault(); Adicionar(new CampoMapa { L = _linhas.Count == 0 ? 1 : (int)_linhas.Max(x => x.L.Value) + 1, Operacao = u == null ? "I" : u.Operacao, Origem = "manual" }, true); }));
            _resumo = Visual.Rotulo("", Visual.Normal, Visual.Suave);
            _resumo.MaximumSize = new Size(880, 0);
            rodapeCampos.Controls.Add(_resumo);
            rodapeCampos.SetFlowBreak(_resumo, true);

            var rodape = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = FlowDirection.RightToLeft, Padding = new Padding(8) };
            Button ok = Visual.Botao(cabo ? "Confirmar cabo (Ctrl+Enter)" : "Confirmar ponto (Ctrl+Enter)", (s, e) => Confirmar(), true);
            Button cancelar = Visual.Botao("Cancelar", (s, e) => { DialogResult = DialogResult.Cancel; });
            CancelButton = cancelar;
            rodape.Controls.Add(ok); rodape.Controls.Add(cancelar);
            if (podeExcluir)
                rodape.Controls.Add(Visual.Botao(cabo ? "Excluir cabo" : "Excluir ponto", (s, e) =>
                {
                    if (MessageBox.Show(this, cabo ? "Excluir este cabo da montagem?" : "Excluir este ponto da montagem?", "OrçaCore", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        DialogResult = DialogResult.Abort;
                }));
            rodape.Controls.Add(Visual.Rotulo("Ctrl+Enter confirma · Esc cancela", Visual.Normal, Visual.Suave));

            Controls.Add(_lista); Controls.Add(rodapeCampos); Controls.Add(rodape); Controls.Add(barraCampos); Controls.Add(topo);
            foreach (CampoMapa c in campos ?? new List<CampoMapa>()) Adicionar(c, false);
            if (_linhas.Count == 0) Adicionar(new CampoMapa { L = 1, Operacao = "I", Origem = "manual" }, false);
            Visual.Aplicar(this);
            CoresDosComprimentos();
            AtualizarResumo();
            KeyDown += (s, e) => { if (e.Control && e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; Confirmar(); } };
        }

        private void Marcar(LinhaCampo l, string op)
        {
            foreach (RadioButton r in l.Ops) r.Checked = r.Text == op;
        }

        private void Adicionar(CampoMapa c, bool focar)
        {
            var l = new LinhaCampo { Entidade = c.Entidade };
            l.Painel = new Panel { Width = 900, Height = 32, Margin = new Padding(0, 1, 0, 1) };
            l.L = new NumericUpDown { Minimum = 1, Maximum = 99, Value = Math.Max(1, Math.Min(99, c.L)), Width = 44, Left = 4, Top = 5 };
            l.Texto = new TextBox { Text = c.Texto ?? "", Width = _cabo ? 260 : 330, Left = 54, Top = 5 };
            l.Texto.TextChanged += (s, e) => AtualizarResumo();
            int x = _cabo ? 322 : 392;
            foreach (string op in MapaManual.Operacoes)
            {
                var r = new RadioButton { Text = op, Appearance = Appearance.Button, AutoSize = false, Width = 38, Height = 24, Left = x, Top = 4,
                    TextAlign = ContentAlignment.MiddleCenter, Checked = op == (MapaManual.NormalizarOperacao(c.Operacao) ?? "I") };
                r.CheckedChanged += (s, e) => { if (((RadioButton)s).Checked) { l.Origem.Text = ""; l.Origem.Tag = "manual"; AtualizarResumo(); } };
                l.Ops.Add(r); l.Painel.Controls.Add(r);
                x += op == "F" || op == "*I" || op == "*M" ? 44 : 39;
            }
            if (_cabo)
            {
                // 0.16.4: comprimento da linha — o do texto (próprio) ou o «Comprimento usado» (acompanha; em cinza)
                l.Proprio = c.Metros.HasValue;
                l.M = new TextBox { Width = 52, Left = x + 4, Top = 5, TextAlign = HorizontalAlignment.Right,
                    Text = MapaManual.FormatarMetros(c.Metros ?? (Numero(_metros.Text) ?? MapaManual.UsadoPadrao(_medido))) };
                l.M.TextChanged += (s, e) =>
                {
                    if (_sincronizando) return;
                    l.Proprio = (l.M.Text ?? "").Trim().Length > 0;
                    if (!l.Proprio) { _sincronizando = true; l.M.Text = _metros.Text; _sincronizando = false; }
                    CoresDosComprimentos();
                };
                Label m = Visual.Rotulo("m", Visual.Normal, Visual.Suave);
                m.Left = x + 58; m.Top = 8; m.AutoSize = true;
                l.Painel.Controls.Add(l.M); l.Painel.Controls.Add(m);
                x += 74;
            }
            l.Origem = Visual.Rotulo(c.Origem == "regra" ? "regra" : c.Origem == "sem regra" ? "sem regra → F" : "", Visual.Normal, c.Origem == "sem regra" ? Visual.Conferir : Visual.Ok);
            l.Origem.Tag = c.Origem ?? "manual"; l.Origem.Left = x + 4; l.Origem.Top = 8; l.Origem.AutoSize = true;
            Button tirar = Visual.Botao("×", (s, e) => { _lista.Controls.Remove(l.Painel); _linhas.Remove(l); AtualizarResumo(); });
            tirar.Width = 30; tirar.Left = 860; tirar.Top = 3; tirar.Padding = new Padding(0);
            l.Painel.Controls.AddRange(new Control[] { l.L, l.Texto, l.Origem, tirar });
            _linhas.Add(l); _lista.Controls.Add(l.Painel);
            if (focar) { Visual.Aplicar(l.Painel); CoresDosComprimentos(); l.Texto.Focus(); }
        }

        /// <summary>O «Comprimento usado» mudou: as linhas sem comprimento próprio acompanham.</summary>
        private void SincronizarUsado()
        {
            _sincronizando = true;
            foreach (LinhaCampo l in _linhas) if (_cabo && !l.Proprio && l.M != null) l.M.Text = _metros.Text;
            _sincronizando = false;
            CoresDosComprimentos();
        }

        private void CoresDosComprimentos()
        {
            foreach (LinhaCampo l in _linhas) if (l.M != null) l.M.ForeColor = l.Proprio ? Visual.Texto : Visual.Suave;
        }

        private void AtualizarResumo()
        {
            if (_resumo == null) return;
            var partes = MapaManual.Operacoes.Select(op => new { op, itens = _linhas.Where(l => l.Operacao == op && (l.Texto.Text ?? "").Trim().Length > 0).Select(l => l.Texto.Text.Trim().ToUpperInvariant()).ToList() })
                .Where(p => p.itens.Count > 0).Select(p => p.op + ": " + string.Join(", ", p.itens.ToArray()));
            _resumo.Text = string.Join("   ·   ", partes.ToArray());
        }

        private void Confirmar()
        {
            if (Campos.Count == 0) { MessageBox.Show(this, "Preencha ao menos um campo.", "OrçaCore", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (_cabo && !Metros.HasValue) { MessageBox.Show(this, "Informe o comprimento usado do cabo, em metros.", "OrçaCore", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (_cabo && _linhas.Any(l => l.Proprio && (l.Texto.Text ?? "").Trim().Length > 0 && !Numero(l.M.Text).HasValue))
            { MessageBox.Show(this, "Há linha com comprimento inválido. Use um número maior que zero (ou apague para usar o comprimento usado).", "OrçaCore", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            DialogResult = DialogResult.OK;
        }
    }
}
