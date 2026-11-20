using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OrcamentoNanoCAD.Nucleo;

namespace OrcamentoNanoCAD.UI
{
    internal sealed class TelaNota : Form
    {
        private readonly TextBox _texto = new TextBox { Dock = DockStyle.Fill, Multiline = true, ScrollBars = ScrollBars.Vertical };
        private readonly ComboBox _operacao = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 68 };
        private readonly ComboBox _alinhamento = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 115 };
        private readonly NumericUpDown _altura = new NumericUpDown { DecimalPlaces = 3, Minimum = 0.001m, Maximum = 10000, Width = 85 };
        private readonly CheckBox _guia = new CheckBox { Text = "Criar linha guia", Checked = true, AutoSize = true };
        private readonly int _alinhamentoInicial;

        public string Conteudo { get { return _texto.Text.Replace("\r\n", "\\P").Replace("\n", "\\P"); } }
        public string Operacao { get { return _operacao.SelectedIndex == 0 ? null : _operacao.Text; } }
        public double Altura { get { return (double)_altura.Value; } }
        public bool Guia { get { return _guia.Checked; } }
        public int Alinhamento { get { return _alinhamento.SelectedIndex == 0 ? _alinhamentoInicial : _alinhamento.SelectedIndex; } }

        public TelaNota(RegrasProjeto regras, string texto, string operacao, double altura, int alinhamento)
        {
            bool criar = alinhamento == 0;
            _alinhamentoInicial = alinhamento;
            Text = criar ? "Criar anotação" : "Editar anotação";
            Width = 520; Height = 365; MinimumSize = new Size(460, 280);
            _texto.Text = texto.Replace("\\P", Environment.NewLine);
            _altura.Value = (decimal)Math.Min(10000, Math.Max(0.001, altura));
            _operacao.Items.Add("Manter");
            _operacao.Items.AddRange(regras.Operacoes().Where(op => regras.LayerDaOperacao(op) != null).ToArray());
            _operacao.SelectedIndex = Math.Max(0, _operacao.Items.IndexOf(operacao ?? "Manter"));
            _alinhamento.Items.AddRange(new object[] { "Manter", "Esquerda", "Centro", "Direita" });
            _alinhamento.SelectedIndex = criar ? 1 : 0;
            var opcoes = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(5) };
            opcoes.Controls.AddRange(new Control[] { new Label { Text = "Operação", AutoSize = true }, _operacao,
                new Label { Text = "Alinhamento", AutoSize = true }, _alinhamento, new Label { Text = "Altura", AutoSize = true }, _altura });
            if (criar) opcoes.Controls.Add(_guia);
            var botoes = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(5) };
            var salvar = Visual.Botao("Confirmar", (sender, args) => DialogResult = DialogResult.OK, true);
            salvar.Enabled = !string.IsNullOrEmpty(_texto.Text.Trim());
            _texto.TextChanged += (sender, args) => salvar.Enabled = !string.IsNullOrEmpty(_texto.Text.Trim());
            var cancelar = Visual.Botao("Cancelar", (sender, args) => DialogResult = DialogResult.Cancel);
            botoes.Controls.Add(salvar); botoes.Controls.Add(cancelar);
            CancelButton = cancelar;
            Controls.Add(_texto); Controls.Add(botoes); Controls.Add(opcoes);
            Visual.Aplicar(this);
        }
    }
}
