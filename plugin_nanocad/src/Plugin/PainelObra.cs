// Plugin/PainelObra.cs — painel (janela não modal) da obra: lista de postes/vãos com filtros Ok/Conferir/Erro, e os botões
// Reler (do zero) / Atualizar / Salvar / Vínculos / Ativos; em "Construir do zero", + Poste / + Vão / Vincular automático.
// Os botões disparam comandos (ORC_*), que rodam no contexto de comando do nanoCAD.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using OrcamentoNanoCAD.Nucleo;

namespace OrcamentoNanoCAD.UI
{
    public class PainelObra : Form
    {
        private static PainelObra _aberto;

        private Sessao _s;
        private string _filtro;                    // null = todos
        private readonly Label _titulo, _rodape;
        private readonly RadioButton _todos, _ok, _conf, _erro;
        private readonly FlowLayoutPanel _botoesLer, _botoesConstruir;
        private readonly ListView _lista;

        public static void Abrir(Sessao s)
        {
            if (_aberto == null || _aberto.IsDisposed)
            {
                _aberto = new PainelObra();
                _aberto.Mostrar(s);
                HostMgd.ApplicationServices.Application.ShowModelessDialog(_aberto);
            }
            else
            {
                _aberto.Mostrar(s);
                _aberto.Visible = true;
                _aberto.Activate();
            }
        }

        public static void AtualizarSeAberto(Sessao s)
        {
            if (_aberto != null && !_aberto.IsDisposed && _aberto._s == s) _aberto.Mostrar(s);
        }

        private PainelObra()
        {
            Text = "Orçamento — obra";
            Width = 470; Height = 640;
            StartPosition = FormStartPosition.Manual;
            Location = new Point(Screen.PrimaryScreen.WorkingArea.Right - Width - 20, Screen.PrimaryScreen.WorkingArea.Top + 80);
            Font = Visual.Normal;
            ShowInTaskbar = false;

            FlowLayoutPanel topo = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(6) };
            _titulo = Visual.Rotulo("", Visual.Titulo);
            topo.Controls.Add(_titulo);
            FlowLayoutPanel filtros = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            _todos = Filtro(filtros, null);
            _ok = Filtro(filtros, Estados.Ok);
            _conf = Filtro(filtros, Estados.Conferir);
            _erro = Filtro(filtros, Estados.Erro);
            topo.Controls.Add(filtros);

            _botoesLer = new FlowLayoutPanel { AutoSize = true, WrapContents = true, MaximumSize = new Size(450, 0) };
            _botoesLer.Controls.Add(Visual.Botao("⟳ Reler (do zero)", (o, e) => Visual.Executar("ORC_RELER")));
            _botoesLer.Controls.Add(Visual.Botao("↻ Atualizar", (o, e) => Visual.Executar("ORC_ATUALIZAR")));
            _botoesConstruir = new FlowLayoutPanel { AutoSize = true, WrapContents = true, MaximumSize = new Size(450, 0) };
            _botoesConstruir.Controls.Add(Visual.Botao("+ Poste (bloco + anotação)", (o, e) => Visual.Executar("ORC_POSTE")));
            _botoesConstruir.Controls.Add(Visual.Botao("+ Vão (linha + anotação)", (o, e) => Visual.Executar("ORC_VAO")));
            _botoesConstruir.Controls.Add(Visual.Botao("Vincular automático", (o, e) => Visual.Executar("ORC_VINCULAR")));
            FlowLayoutPanel comuns = new FlowLayoutPanel { AutoSize = true, WrapContents = true, MaximumSize = new Size(450, 0) };
            Button salvar = Visual.Botao("💾 Salvar", (o, e) => Visual.Executar("ORC_SALVAR"));
            salvar.BackColor = Visual.Ok; salvar.ForeColor = Color.White;
            comuns.Controls.Add(salvar);
            comuns.Controls.Add(Visual.Botao("🔗 Vínculos", (o, e) => Visual.Executar("ORC_VINCULOS")));
            comuns.Controls.Add(Visual.Botao("📦 Ativos", (o, e) => Visual.Executar("ORC_ATIVOS")));
            comuns.Controls.Add(Visual.Botao("Marcas", (o, e) => Visual.Executar("ORC_MARCAS")));
            topo.Controls.Add(_botoesLer);
            topo.Controls.Add(_botoesConstruir);
            topo.Controls.Add(comuns);

            _lista = Visual.Lista(Relatorios.ColunasConferencia);
            _lista.Click += (o, e) => { string n = Selecionado(); if (n != null) Visual.Executar("ORC_ZOOM " + n); };
            _lista.ItemActivate += (o, e) => { string n = Selecionado(); if (n != null) Visual.Executar("ORC_EDITAR " + n); };

            _rodape = Visual.Rotulo("", Visual.Normal, Color.DimGray);
            _rodape.Dock = DockStyle.Bottom; _rodape.AutoSize = false; _rodape.Height = 54; _rodape.Padding = new Padding(6);

            Controls.Add(_lista);
            Controls.Add(_rodape);
            Controls.Add(topo);
        }

        private RadioButton Filtro(FlowLayoutPanel onde, string estado)
        {
            RadioButton r = new RadioButton { Appearance = Appearance.Button, AutoSize = true, Tag = estado, Checked = estado == null, Margin = new Padding(2) };
            if (estado != null) r.ForeColor = Visual.DoEstado(estado);
            r.CheckedChanged += (o, e) => { if (r.Checked) { _filtro = (string)r.Tag; Preencher(); } };
            onde.Controls.Add(r);
            return r;
        }

        private string Selecionado()
        {
            return _lista.SelectedItems.Count > 0 ? _lista.SelectedItems[0].Tag as string : null;
        }

        private void Mostrar(Sessao s)
        {
            _s = s;
            Obra o = s.Obra;
            if (o == null) return;
            bool construir = o.Modo == "construir";
            _titulo.Text = (construir ? "🧱 Construir do zero" : "📖 Projeto lido — conferência") + (s.Modificada ? "  •" : "");
            _botoesLer.Visible = !construir;
            _botoesConstruir.Visible = construir;
            _todos.Text = "Todos (" + Conferencia.Contar(o, null) + ")";
            _ok.Text = "✔ Ok (" + Conferencia.Contar(o, Estados.Ok) + ")";
            _conf.Text = "⚠ Conferir (" + Conferencia.Contar(o, Estados.Conferir) + ")";
            _erro.Text = "✖ Erro (" + Conferencia.Contar(o, Estados.Erro) + ")";
            _rodape.Text = Relatorios.Resumo(o) + "\nClique numa linha para localizar no desenho; duplo clique (ou ORC_EDITAR) abre o pop-up de ajuste.";
            Preencher();
        }

        private void Preencher()
        {
            if (_s == null || _s.Obra == null) return;
            Visual.Preencher(_lista, Relatorios.Conferencia(_s.Obra, _filtro));
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Hide(); }   // reabre com ORC_PAINEL
            base.OnFormClosing(e);
        }
    }
}
