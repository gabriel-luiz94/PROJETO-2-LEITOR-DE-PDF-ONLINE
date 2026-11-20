using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using OrcamentoNanoCAD.Nucleo;

namespace OrcamentoNanoCAD.UI
{
    internal sealed class LoteOperacao : Form
    {
        private readonly ComboBox _op = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly CheckBox _dwg = new CheckBox { AutoSize = true, Text = "Alterar também a layer dos textos no DWG" };
        public string Operacao { get { return _op.Text; } }
        public bool AlterarLayer { get { return _dwg.Checked; } }

        public LoteOperacao(Sessao s, List<Elemento> itens)
        {
            Text = "Prévia da alteração"; Width = 610; Height = 360; StartPosition = FormStartPosition.CenterParent;
            var topo = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true };
            var rodape = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true };
            var lista = Visual.Lista(new[] { "Item", "Ativo", "Atual", "Proposta", "Layer no DWG" });
            var aplicar = Visual.Botao("Aplicar", (o,e) => DialogResult = DialogResult.OK, true);
            var cancelar = Visual.Botao("Cancelar", (o,e) => DialogResult = DialogResult.Cancel);
            CancelButton = cancelar;
            var handles = new HashSet<string>(itens.Select(x => x.TextoHandle));
            bool compartilhado = s.Obra.Elementos.Except(itens).Any(e => handles.Contains(e.TextoHandle))
                || s.Obra.Ignorados.Any(t => handles.Contains(t.Handle));
            Action previa = () => {
                string layer = s.Regras.LayerDaOperacao(Operacao);
                _dwg.Enabled = !string.IsNullOrEmpty(layer) && !compartilhado && itens.All(e => !string.IsNullOrEmpty(e.TextoHandle));
                if (!_dwg.Enabled) _dwg.Checked = false;
                Visual.Preencher(lista,itens.Select(e => new LinhaTabela { Chave = e.Numero,
                    Celulas = new[] { e.Numero, e.AtivoLido, e.Operacao, Operacao, AlterarLayer ? e.LayerAtual + " → " + layer : "Somente obra" } }).ToList());
            };
            _op.Items.AddRange(s.Regras.Operacoes().ToArray()); _op.SelectedIndex = Math.Max(0,_op.Items.IndexOf(itens[0].Operacao));
            _op.SelectedIndexChanged += (o,e) => previa(); _dwg.CheckedChanged += (o,e) => previa();
            topo.Controls.Add(_op); topo.Controls.Add(_dwg); rodape.Controls.Add(aplicar); rodape.Controls.Add(cancelar);
            Controls.Add(lista); Controls.Add(rodape); Controls.Add(topo); Visual.Aplicar(this); previa();
        }
    }
}
