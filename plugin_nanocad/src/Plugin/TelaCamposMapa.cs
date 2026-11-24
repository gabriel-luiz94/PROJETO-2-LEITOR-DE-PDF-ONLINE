// Plugin/TelaCamposMapa.cs — TASK-038: janela do ponto/cabo no formato do mapa, no molde do «anota poste»:
// CAMPOS (L = linha do texto), cada campo com o ativo e a operação (F, I, *I, M, *M, R, *R); + Novo campo / + Nova linha; «Todos» por operação.
// Ponto: nome - comentário. Cabo: comprimento (anotado ou medido na linha). Ctrl+Enter confirma, Esc cancela. C# 5 (nanoCAD 5).
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
        private readonly double? _anotado;

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
                    Origem = (string)l.Origem.Tag ?? "manual", Entidade = l.Entidade }).ToList();
            }
        }

        /// <summary>`titulo` = «Ponto P12 · poste (bloco …)» ou «Cabo C3 · trecho de 38,4 m». Para cabo: `medido` (m) e `anotado` (metragem do texto).</summary>
        public TelaCamposMapa(string titulo, bool cabo, List<CampoMapa> campos, string nome, string comentario, double medido, double? anotado, string infoRegras, bool podeExcluir = false)
        {
            _cabo = cabo; _medido = medido; _anotado = anotado;
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
                _metros.Text = MapaManual.FormatarMetros(anotado ?? medido);
                topo.Controls.Add(Visual.Rotulo("Comprimento (m)", Visual.Normal)); topo.Controls.Add(_metros);
                topo.Controls.Add(Visual.Rotulo("medido na linha: " + MapaManual.FormatarMetros(medido) + " m" + (anotado.HasValue ? " · anotado no texto: " + MapaManual.FormatarMetros(anotado.Value) + " m" : " · sem metragem no texto"), Visual.Normal, Visual.Suave));
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
            l.Texto = new TextBox { Text = c.Texto ?? "", Width = 330, Left = 54, Top = 5 };
            l.Texto.TextChanged += (s, e) => AtualizarResumo();
            int x = 392;
            foreach (string op in MapaManual.Operacoes)
            {
                var r = new RadioButton { Text = op, Appearance = Appearance.Button, AutoSize = false, Width = 38, Height = 24, Left = x, Top = 4,
                    TextAlign = ContentAlignment.MiddleCenter, Checked = op == (MapaManual.NormalizarOperacao(c.Operacao) ?? "I") };
                r.CheckedChanged += (s, e) => { if (((RadioButton)s).Checked) { l.Origem.Text = ""; l.Origem.Tag = "manual"; AtualizarResumo(); } };
                l.Ops.Add(r); l.Painel.Controls.Add(r);
                x += op == "F" || op == "*I" || op == "*M" ? 44 : 39;
            }
            l.Origem = Visual.Rotulo(c.Origem == "regra" ? "regra" : c.Origem == "sem regra" ? "sem regra → F" : "", Visual.Normal, c.Origem == "sem regra" ? Visual.Conferir : Visual.Ok);
            l.Origem.Tag = c.Origem ?? "manual"; l.Origem.Left = x + 4; l.Origem.Top = 8; l.Origem.AutoSize = true;
            Button tirar = Visual.Botao("×", (s, e) => { _lista.Controls.Remove(l.Painel); _linhas.Remove(l); AtualizarResumo(); });
            tirar.Width = 30; tirar.Left = 860; tirar.Top = 3; tirar.Padding = new Padding(0);
            l.Painel.Controls.AddRange(new Control[] { l.L, l.Texto, l.Origem, tirar });
            _linhas.Add(l); _lista.Controls.Add(l.Painel);
            if (focar) { Visual.Aplicar(l.Painel); l.Texto.Focus(); }
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
            if (_cabo && !Metros.HasValue) { MessageBox.Show(this, "Informe o comprimento do cabo, em metros.", "OrçaCore", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            DialogResult = DialogResult.OK;
        }
    }
}
