// Plugin/CampoAtivo.cs — campo de ativo com o autocompletar da aplicação (TASK-017 do plugin / TASK-144):
//   • digitando um nome: lista de sugestões da Base de Orçamento do projeto (com descrição; apelido → oficial);
//   • dentro de "(": a dica dos parâmetros (o atual em destaque) e os valores dele na lista;
//   • ↑/↓ escolhem, Enter/Tab aceitam, Esc fecha; ao confirmar ou sair do campo, o ")" que falta é completado.
// O catálogo vem de Catalogos (aplicação conectada, com cópia local); sem catálogo, o campo funciona como texto comum.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OrcamentoNanoCAD.Nucleo;

namespace OrcamentoNanoCAD.UI
{
    public class CampoAtivo : FlowLayoutPanel
    {
        private readonly TextBox _txt;
        private readonly ListBox _lista;
        private readonly Label _dica;
        private CatalogoBase _cat;
        private readonly bool _cabo;
        private int _inicioTroca = -1;        // trecho que a escolha da lista substitui
        private bool _ehValor;

        public CampoAtivo(string texto, bool cabo, int largura)
        {
            _cabo = cabo;
            FlowDirection = FlowDirection.TopDown;
            WrapContents = false;
            AutoSize = true;
            Margin = new Padding(0, 2, 0, 4);
            _txt = new TextBox { Width = largura, Text = texto ?? "", Font = Visual.Mono };
            _lista = new ListBox { Width = largura, Height = 132, Visible = false, Font = Visual.Normal, IntegralHeight = false };
            _dica = Visual.Rotulo("", Visual.Normal, Visual.Suave);
            _dica.MaximumSize = new Size(largura, 0);
            Controls.Add(_txt);
            Controls.Add(_lista);
            Controls.Add(_dica);
            _txt.TextChanged += (o, e) => Atualizar();
            _txt.KeyUp += (o, e) => { if (e.KeyCode == Keys.Left || e.KeyCode == Keys.Right || e.KeyCode == Keys.Home || e.KeyCode == Keys.End) Atualizar(); };
            _txt.MouseUp += (o, e) => Atualizar();
            _txt.KeyDown += Tecla;
            _txt.PreviewKeyDown += (o, e) => { if (e.KeyCode == Keys.Tab && _lista.Visible) e.IsInputKey = true; };
            _txt.Leave += (o, e) => { if (!_lista.Focused) { Completar(); Esconder(); } };
            _lista.MouseClick += (o, e) => Aceitar();
            _lista.KeyDown += (o, e) => { if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab) { Aceitar(); e.Handled = true; e.SuppressKeyPress = true; } };
        }

        public TextBox Caixa { get { return _txt; } }

        /// <summary>O texto, com os ")" que faltam completados.</summary>
        public string Texto { get { return Autocompletar.CompletarParenteses((_txt.Text ?? "").Trim()); } }

        public void DefinirCatalogo(CatalogoBase c)
        {
            _cat = c;
            if (_txt.Focused) Atualizar();
            else if (c == null || c.Vazio) _dica.Text = "";
        }

        public void Situacao(string texto) { if (!_txt.Focused) _dica.Text = texto ?? ""; }

        private void Tecla(object o, KeyEventArgs e)
        {
            if (_lista.Visible && _lista.Items.Count > 0)
            {
                if (e.KeyCode == Keys.Down) { _lista.SelectedIndex = Math.Min(_lista.Items.Count - 1, _lista.SelectedIndex + 1); e.Handled = true; return; }
                if (e.KeyCode == Keys.Up) { _lista.SelectedIndex = Math.Max(0, _lista.SelectedIndex - 1); e.Handled = true; return; }
                if ((e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab) && _lista.SelectedIndex >= 0) { Aceitar(); e.Handled = true; e.SuppressKeyPress = true; return; }
            }
            if (e.KeyCode == Keys.Escape && _lista.Visible) { Esconder(); e.Handled = true; e.SuppressKeyPress = true; return; }
            if (e.KeyCode == Keys.Enter) Completar();          // Enter sem lista: completa ")" e deixa o diálogo seguir (botão padrão)
        }

        private void Completar()
        {
            string novo = Autocompletar.CompletarParenteses(_txt.Text);
            if (novo != _txt.Text) { _txt.Text = novo; _txt.SelectionStart = novo.Length; }
        }

        private void Esconder() { _lista.Visible = false; _lista.Items.Clear(); _inicioTroca = -1; }

        private void Atualizar()
        {
            if (_cat == null || _cat.Vazio) { Esconder(); return; }
            string t = _txt.Text ?? "";
            int cur = _txt.SelectionStart;
            ContextoParentese ctx = Autocompletar.ContextoDoParentese(t, cur);
            if (ctx != null && ctx.Nome.Length > 0)
            {
                DicaAtivo d = Autocompletar.Dica(_cat, ctx.Nome);
                _dica.Text = Autocompletar.TextoDaDica(d, ctx);
                _dica.ForeColor = d.Tipo == "parametros" ? Visual.Texto : d.Tipo == "nao_encontrado" ? Visual.Erro : Visual.Conferir;
                Mostrar(Autocompletar.ValoresDoParametro(d, ctx).Cast<object>().ToList(), ctx.InicioParcial, true);
                return;
            }
            PalavraAtual p = Autocompletar.Palavra(t, cur);
            List<ItemCatalogo> sug = Autocompletar.Sugerir(_cat, p.Parcial, 12, _cabo);
            _dica.ForeColor = Visual.Suave;
            _dica.Text = p.Parcial.Length == 0 ? "Digite para ver os ativos da Base de Orçamento · ( mostra os parâmetros"
                : sug.Count == 0 ? (_cat.Existe(p.Parcial) ? "" : p.Parcial.ToUpperInvariant() + ": não está na base") : "↑/↓ escolhe · Enter aceita · Esc fecha";
            Mostrar(sug.Cast<object>().ToList(), p.Inicio, false);
        }

        private void Mostrar(List<object> itens, int inicio, bool valores)
        {
            if (itens.Count == 0) { Esconder(); return; }
            _lista.BeginUpdate();
            _lista.Items.Clear();
            foreach (object i in itens) _lista.Items.Add(i);
            _lista.SelectedIndex = 0;
            _lista.EndUpdate();
            _lista.Visible = true;
            _inicioTroca = inicio;
            _ehValor = valores;
        }

        private void Aceitar()
        {
            if (_lista.SelectedItem == null || _inicioTroca < 0) return;
            ItemCatalogo it = _lista.SelectedItem as ItemCatalogo;
            string valor = it != null ? (it.ApelidoDe ?? it.Nome) : _lista.SelectedItem.ToString();
            int cur;
            string novo = Autocompletar.Substituir(_txt.Text, _inicioTroca, _txt.SelectionStart, valor, out cur);
            Esconder();
            _txt.Text = novo;
            _txt.Focus();
            _txt.SelectionStart = cur;
            _txt.SelectionLength = 0;
            Atualizar();
        }
    }
}
