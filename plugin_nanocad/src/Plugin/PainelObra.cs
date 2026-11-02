// Plugin/PainelObra.cs — painel (janela não modal) da obra, no visual da aplicação:
//   navegação  ⌂ Início · ✕ Fechar obra · ⚙ Leitura · 🩺 Diagnóstico
//   obra       Ler: ⟳ Reler (do zero) · ↻ Atualizar   |   Construir: + Poste · + Vão · Vincular automático
//   saída      💾 Salvar · 🌐 Abrir na aplicação · 🔗 Vínculos · 📦 Ativos · Marcas
//   aplicação  🔗 Conectar · ⇅ Sincronizar + situação da ligação
//   lista      Todos / Ok / Conferir / Erro / Ignorados (textos que as regras descartaram, com o motivo)
// Os botões disparam comandos (ORC_*), que rodam no contexto de comando do nanoCAD.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OrcamentoNanoCAD.Nucleo;

namespace OrcamentoNanoCAD.UI
{
    public class PainelObra : Form
    {
        private static PainelObra _aberto;
        private const string Ignorados = "IGNORADOS";

        private Sessao _s;
        private string _filtro;                    // null = todos; Estados.*; ou Ignorados
        private readonly Label _titulo, _subtitulo, _rodape, _conexao;
        private readonly RadioButton _todos, _ok, _conf, _erro, _ign;
        private readonly FlowLayoutPanel _botoesLer, _botoesConstruir;
        private readonly ListView _lista, _listaIgnorados;

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

        public static void Esconder()
        {
            if (_aberto != null && !_aberto.IsDisposed) _aberto.Hide();
        }

        public static void AtualizarStatusConexao()
        {
            if (_aberto != null && !_aberto.IsDisposed) _aberto._conexao.Text = Sincronizador.Status;
        }

        public static void AtualizarSeAberto(Sessao s)
        {
            if (_aberto != null && !_aberto.IsDisposed && _aberto._s == s) _aberto.Mostrar(s);
        }

        private static FlowLayoutPanel Linha()
        {
            return new FlowLayoutPanel { AutoSize = true, WrapContents = true, MaximumSize = new Size(520, 0), Margin = new Padding(0, 2, 0, 2) };
        }

        private PainelObra()
        {
            Text = "Orçamento — obra";
            Width = 540; Height = 720;
            StartPosition = FormStartPosition.Manual;
            Location = new Point(Screen.PrimaryScreen.WorkingArea.Right - Width - 20, Screen.PrimaryScreen.WorkingArea.Top + 60);
            Font = Visual.Normal;
            ShowInTaskbar = false;

            FlowLayoutPanel topo = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(10, 8, 10, 4) };
            FlowLayoutPanel nav = Linha();
            nav.Controls.Add(Visual.Botao("⌂ Início", (o, e) => Visual.Executar("ORCAMENTO")));
            nav.Controls.Add(Visual.Botao("✕ Fechar obra", (o, e) => Visual.Executar("ORC_FECHAR")));
            nav.Controls.Add(Visual.Botao("⚙ Leitura", (o, e) => Visual.Executar("ORC_LEITURA")));
            nav.Controls.Add(Visual.Botao("🩺 Diagnóstico", (o, e) => Visual.Executar("ORC_DIAGNOSTICO")));
            topo.Controls.Add(nav);
            _titulo = Visual.Rotulo("", Visual.Titulo);
            _titulo.Margin = new Padding(0, 8, 0, 0);
            topo.Controls.Add(_titulo);
            _subtitulo = Visual.Rotulo("", Visual.Normal, Visual.Suave);
            _subtitulo.MaximumSize = new Size(510, 0);
            topo.Controls.Add(_subtitulo);

            FlowLayoutPanel filtros = Linha();
            _todos = Filtro(filtros, null);
            _ok = Filtro(filtros, Estados.Ok);
            _conf = Filtro(filtros, Estados.Conferir);
            _erro = Filtro(filtros, Estados.Erro);
            _ign = Filtro(filtros, Ignorados);
            topo.Controls.Add(filtros);

            _botoesLer = Linha();
            _botoesLer.Controls.Add(Visual.Botao("⟳ Reler (do zero)", (o, e) => Visual.Executar("ORC_RELER"), true));
            _botoesLer.Controls.Add(Visual.Botao("↻ Atualizar", (o, e) => Visual.Executar("ORC_ATUALIZAR")));
            _botoesConstruir = Linha();
            _botoesConstruir.Controls.Add(Visual.Botao("+ Poste (bloco + anotação)", (o, e) => Visual.Executar("ORC_POSTE"), true));
            _botoesConstruir.Controls.Add(Visual.Botao("+ Vão (linha + anotação)", (o, e) => Visual.Executar("ORC_VAO"), true));
            _botoesConstruir.Controls.Add(Visual.Botao("Vincular automático", (o, e) => Visual.Executar("ORC_VINCULAR")));
            topo.Controls.Add(_botoesLer);
            topo.Controls.Add(_botoesConstruir);

            FlowLayoutPanel saida = Linha();
            saida.Controls.Add(Visual.Botao("💾 Salvar", (o, e) => Visual.Executar("ORC_SALVAR")));
            saida.Controls.Add(Visual.Botao("🌐 Abrir na aplicação", (o, e) => Visual.Executar("ORC_ABRIR_APP")));
            saida.Controls.Add(Visual.Botao("🔗 Vínculos", (o, e) => Visual.Executar("ORC_VINCULOS")));
            saida.Controls.Add(Visual.Botao("📦 Ativos", (o, e) => Visual.Executar("ORC_ATIVOS")));
            saida.Controls.Add(Visual.Botao("Marcas", (o, e) => Visual.Executar("ORC_MARCAS")));
            topo.Controls.Add(saida);

            FlowLayoutPanel app = Linha();
            app.Controls.Add(Visual.Botao("🔗 Conectar à aplicação", (o, e) => Visual.Executar("ORC_CONECTAR")));
            app.Controls.Add(Visual.Botao("⇅ Sincronizar", (o, e) => Visual.Executar("ORC_SINCRONIZAR")));
            topo.Controls.Add(app);
            _conexao = Visual.Rotulo(Sincronizador.Status, Visual.Normal, Visual.Suave);
            _conexao.MaximumSize = new Size(510, 0);
            topo.Controls.Add(_conexao);

            _lista = Visual.Lista(Relatorios.ColunasConferencia);
            _lista.Click += (o, e) => { string n = Selecionado(_lista); if (n != null) Visual.Executar("ORC_ZOOM " + n); };
            _lista.ItemActivate += (o, e) => { string n = Selecionado(_lista); if (n != null) Visual.Executar("ORC_EDITAR " + n); };
            _listaIgnorados = Visual.Lista(Relatorios.ColunasIgnorados);
            _listaIgnorados.Visible = false;
            _listaIgnorados.Click += (o, e) => { string n = Selecionado(_listaIgnorados); if (n != null) Visual.Executar("ORC_ZOOM " + n); };

            _rodape = Visual.Rotulo("", Visual.Normal, Visual.Suave);
            _rodape.Dock = DockStyle.Bottom; _rodape.AutoSize = false; _rodape.Height = 56; _rodape.Padding = new Padding(10, 6, 10, 6);

            Controls.Add(_lista);
            Controls.Add(_listaIgnorados);
            Controls.Add(_rodape);
            Controls.Add(topo);
            Visual.Aplicar(this);
        }

        private RadioButton Filtro(FlowLayoutPanel onde, string estado)
        {
            RadioButton r = new RadioButton { Appearance = Appearance.Button, AutoSize = true, Tag = estado, Checked = estado == null, Margin = new Padding(2), Padding = new Padding(4, 1, 4, 1) };
            r.ForeColor = estado == null || estado == Ignorados ? Visual.Texto : Visual.DoEstado(estado);
            r.CheckedChanged += (o, e) => { if (r.Checked) { _filtro = (string)r.Tag; Preencher(); } };
            onde.Controls.Add(r);
            return r;
        }

        private static string Selecionado(ListView l)
        {
            return l.SelectedItems.Count > 0 ? l.SelectedItems[0].Tag as string : null;
        }

        private void Mostrar(Sessao s)
        {
            _s = s;
            Obra o = s.Obra;
            if (o == null) return;
            bool construir = o.Modo == "construir";
            _titulo.Text = (construir ? "🧱 Construir do zero" : "📖 Projeto lido — conferência") + (s.Modificada ? "  •" : "");
            _subtitulo.Text = o.Nome + " · projeto " + (string.IsNullOrEmpty(o.Projeto) ? "(sem código)" : o.Projeto) + " · " + (s.Regras != null ? s.Regras.Nome : "")
                              + (s.Modificada ? " · mudanças não salvas" : (o.SalvoEm != null ? " · salvo " + o.SalvoEm : ""));
            _botoesLer.Visible = !construir;
            _botoesConstruir.Visible = construir;
            _todos.Text = "Todos (" + Conferencia.Contar(o, null) + ")";
            _ok.Text = "✔ Ok (" + Conferencia.Contar(o, Estados.Ok) + ")";
            _conf.Text = "⚠ Conferir (" + Conferencia.Contar(o, Estados.Conferir) + ")";
            _erro.Text = "✖ Erro (" + Conferencia.Contar(o, Estados.Erro) + ")";
            _ign.Text = "Ignorados (" + o.Ignorados.Count + ")";
            _rodape.Text = Relatorios.Resumo(o) + "\nClique numa linha para localizar no desenho; duplo clique abre o pop-up de ajuste. "
                           + (construir ? "Use + Poste e + Vão para montar a obra." : "Ignorados = textos que as regras do leitor descartaram.");
            _conexao.Text = Sincronizador.Status;
            Preencher();
        }

        private void Preencher()
        {
            if (_s == null || _s.Obra == null) return;
            bool ign = _filtro == Ignorados;
            _lista.Visible = !ign;
            _listaIgnorados.Visible = ign;
            if (ign) Visual.Preencher(_listaIgnorados, Relatorios.TextosIgnorados(_s.Obra));
            else Visual.Preencher(_lista, Relatorios.Conferencia(_s.Obra, _filtro));
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Hide(); }   // reabre com ORC_PAINEL
            base.OnFormClosing(e);
        }
    }
}
