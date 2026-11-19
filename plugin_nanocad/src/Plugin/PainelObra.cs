// Plugin/PainelObra.cs — painel (janela não modal) da obra, no visual da aplicação (TASK-012 do plugin):
//   barra       Projeto, Conferir e Configurações; Salvar e Abrir no OrçaCore ficam visíveis.
//   obra        Ler: Reler (do zero) · Atualizar   |   Construir: + Poste · + Vão · Vincular automático
//   filtros     Todos / Ok / Conferir / Erro / Ignorados (textos que as regras descartaram, com o motivo)
//   listas      CABOS (vãos) e POSTES E ESTRUTURAS (outros) separadas, como a Tabela de Cabos e a Tabela de Outros do Resumo.
//               Ler: só o que as regras do leitor classificaram; Construir: tudo o que o usuário montou (sem filtro das regras).
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
        private readonly ListView _cabos, _outros, _listaIgnorados;
        private readonly SplitContainer _split;
        private int _nCabos, _nOutros;

        public static void Abrir(Sessao s)
        {
            if (_aberto == null || _aberto.IsDisposed)
            {
                _aberto = new PainelObra();
                _aberto.Mostrar(s);
                Visual.MostrarNaoModal(_aberto);
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
            return new FlowLayoutPanel { AutoSize = true, WrapContents = true, MaximumSize = new Size(640, 0), Margin = new Padding(0, 2, 0, 2) };
        }

        private static ToolStripMenuItem Grupo(MenuStrip menu, string nome)
        {
            ToolStripMenuItem grupo = new ToolStripMenuItem(nome);
            menu.Items.Add(grupo);
            return grupo;
        }

        private static void Acao(ToolStripMenuItem grupo, string nome, string comando)
        {
            grupo.DropDownItems.Add(nome, null, (o, e) => Visual.Executar(comando));
        }

        private PainelObra()
        {
            Text = "Orçamento — obra";
            Width = 660; Height = 820;
            StartPosition = FormStartPosition.Manual;
            Location = new Point(Screen.PrimaryScreen.WorkingArea.Right - Width - 20, Screen.PrimaryScreen.WorkingArea.Top + 40);
            Font = Visual.Normal;
            ShowInTaskbar = false;

            FlowLayoutPanel topo = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, FlowDirection = FlowDirection.TopDown, WrapContents = false, Padding = new Padding(10, 8, 10, 4) };

            _titulo = Visual.Rotulo("", Visual.Titulo);
            _titulo.Margin = new Padding(0, 6, 0, 0);
            topo.Controls.Add(_titulo);
            _subtitulo = Visual.Rotulo("", Visual.Normal, Visual.Suave);
            _subtitulo.MaximumSize = new Size(620, 0);
            topo.Controls.Add(_subtitulo);
            _conexao = Visual.Rotulo(Sincronizador.Status, Visual.Normal, Visual.Suave);
            _conexao.MaximumSize = new Size(620, 0);
            topo.Controls.Add(_conexao);

            FlowLayoutPanel barra = Linha();
            MenuStrip menu = new MenuStrip { AutoSize = true, Dock = DockStyle.None, GripStyle = ToolStripGripStyle.Hidden };
            ToolStripMenuItem projeto = Grupo(menu, "Projeto");
            Acao(projeto, "Início", "ORCAMENTO");
            Acao(projeto, "Nome e número da obra", "ORC_PROJETO");
            Acao(projeto, "Fechar obra", "ORC_FECHAR");
            ToolStripMenuItem conferir = Grupo(menu, "Conferir");
            Acao(conferir, "Vínculos", "ORC_VINCULOS");
            Acao(conferir, "Ativos com parâmetros", "ORC_ATIVOS");
            Acao(conferir, "Mostrar ou ocultar marcas", "ORC_MARCAS");
            ToolStripMenuItem config = Grupo(menu, "Configurações");
            Acao(config, "Leitura do desenho", "ORC_LEITURA");
            Acao(config, "Conectar à aplicação", "ORC_CONECTAR");
            Acao(config, "Sincronizar regras", "ORC_SINCRONIZAR");
            Acao(config, "Diagnóstico", "ORC_DIAGNOSTICO");
            barra.Controls.Add(menu);
            barra.Controls.Add(Visual.Botao("Salvar", (o, e) => Visual.Executar("ORC_SALVAR"), false, "salvar"));
            barra.Controls.Add(Visual.Botao("Abrir no OrçaCore", (o, e) => Visual.Executar("ORC_ABRIR_APP"), false, "app"));
            topo.Controls.Add(barra);

            _botoesLer = Linha();
            _botoesLer.Controls.Add(Visual.Botao("Atualizar (mantém ajustes)", (o, e) => Visual.Executar("ORC_ATUALIZAR"), true, "atualizar"));
            _botoesLer.Controls.Add(Visual.Botao("Reler (do zero)", (o, e) => Visual.Executar("ORC_RELER"), false, "reler"));
            _botoesLer.Controls.Add(Visual.Botao("+ Cabo", (o, e) => Visual.Executar("ORC_NOVO_CABO"), false, "cabos"));
            _botoesLer.Controls.Add(Visual.Botao("+ Outros", (o, e) => Visual.Executar("ORC_NOVO_OUTROS"), false, "ativos"));
            _botoesConstruir = Linha();
            _botoesConstruir.Controls.Add(Visual.Botao("Poste (bloco + anotação)", (o, e) => Visual.Executar("ORC_POSTE"), true, "poste"));
            _botoesConstruir.Controls.Add(Visual.Botao("Vão (linha + anotação)", (o, e) => Visual.Executar("ORC_VAO"), true, "vao"));
            _botoesConstruir.Controls.Add(Visual.Botao("Vincular automático", (o, e) => Visual.Executar("ORC_VINCULAR"), false, "vincular"));
            _botoesConstruir.Controls.Add(Visual.Botao("+ Cabo", (o, e) => Visual.Executar("ORC_NOVO_CABO"), false, "cabos"));
            _botoesConstruir.Controls.Add(Visual.Botao("+ Outros", (o, e) => Visual.Executar("ORC_NOVO_OUTROS"), false, "ativos"));
            topo.Controls.Add(_botoesLer);
            topo.Controls.Add(_botoesConstruir);

            FlowLayoutPanel filtros = Linha();
            _todos = Filtro(filtros, null, "todos");
            _ok = Filtro(filtros, Estados.Ok, "ok");
            _conf = Filtro(filtros, Estados.Conferir, "conferir");
            _erro = Filtro(filtros, Estados.Erro, "erro");
            _ign = Filtro(filtros, Ignorados, "ignorados");
            topo.Controls.Add(filtros);

            // as duas listas, separadas como no Resumo da aplicação
            _cabos = Visual.Lista(Relatorios.ColunasCabos);
            _outros = Visual.Lista(Relatorios.ColunasOutros);
            foreach (ListView l in new[] { _cabos, _outros }) Ligar(l, true);
            _split = new SplitContainer { Dock = DockStyle.Fill, Orientation = Orientation.Horizontal, SplitterWidth = 6, BackColor = Visual.Fundo };
            _split.Panel1.Controls.Add(_cabos);
            _split.Panel1.Controls.Add(Visual.Cabecalho("cabos", () => "CABOS (vãos) · " + _nCabos));
            _split.Panel2.Controls.Add(_outros);
            _split.Panel2.Controls.Add(Visual.Cabecalho("outros", () => "POSTES E ESTRUTURAS (outros) · " + _nOutros));

            _listaIgnorados = Visual.Lista(Relatorios.ColunasIgnorados);
            _listaIgnorados.Visible = false;
            Ligar(_listaIgnorados, false);
            // TASK-017: as linhas excluídas aparecem aqui; duplo clique abre o pop-up com "Restaurar à obra"
            _listaIgnorados.ItemActivate += (o, e) => { string n = Selecionado(_listaIgnorados); if (n != null && !n.StartsWith("H:")) Visual.Executar("ORC_EDITAR " + n); };

            _rodape = Visual.Rotulo("", Visual.Normal, Visual.Suave);
            _rodape.Dock = DockStyle.Bottom; _rodape.AutoSize = false; _rodape.Height = 58; _rodape.Padding = new Padding(10, 6, 10, 6);

            Controls.Add(_split);
            Controls.Add(_listaIgnorados);
            Controls.Add(_rodape);
            Controls.Add(topo);
            Visual.Aplicar(this);
            _split.BackColor = Visual.Borda;                        // a faixa entre as listas
            _split.Panel1.BackColor = Visual.Fundo;
            _split.Panel2.BackColor = Visual.Fundo;
            Shown += (o, e) => { try { _split.SplitterDistance = Math.Max(120, (_split.Height - _split.SplitterWidth) * 55 / 100); } catch (Exception) { } };
        }

        private void Ligar(ListView l, bool editar)
        {
            l.Click += (o, e) => { string n = Selecionado(l); if (n != null) Visual.Executar("ORC_ZOOM " + n); };
            if (editar) l.ItemActivate += (o, e) => { string n = Selecionado(l); if (n != null) Visual.Executar("ORC_EDITAR " + n); };
        }

        private RadioButton Filtro(FlowLayoutPanel onde, string estado, string icone)
        {
            RadioButton r = new RadioButton { Appearance = Appearance.Button, AutoSize = true, Tag = estado, Checked = estado == null, Margin = new Padding(2), Padding = new Padding(4, 2, 6, 2) };
            r.ForeColor = estado == null || estado == Ignorados ? Visual.Texto : Visual.DoEstado(estado);
            r.Image = Icones.Imagem(icone, estado == null || estado == Ignorados ? Visual.Texto : Visual.DoEstado(estado));
            r.ImageAlign = ContentAlignment.MiddleLeft;
            r.TextImageRelation = TextImageRelation.ImageBeforeText;
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
            _titulo.Text = (construir ? "Construir do zero" : "Projeto lido — conferência") + (s.Modificada ? "  •" : "");
            _subtitulo.Text = o.Nome + (string.IsNullOrEmpty(o.NumeroObra) ? "" : " · obra nº " + o.NumeroObra) + " · projeto " + (string.IsNullOrEmpty(o.Projeto) ? "(sem código)" : o.Projeto) + " · " + (s.Regras != null ? s.Regras.Nome + " · regras: " + s.Regras.Origem : "")
                              + (s.Modificada ? " · mudanças não salvas" : (o.SalvoEm != null ? " · salvo " + o.SalvoEm : ""));
            _botoesLer.Visible = !construir;
            _botoesConstruir.Visible = construir;
            _todos.Text = "Todos " + Relatorios.Contar(o, null, null);
            _ok.Text = "Ok " + Relatorios.Contar(o, null, Estados.Ok);
            _conf.Text = "Conferir " + Relatorios.Contar(o, null, Estados.Conferir);
            _erro.Text = "Erro " + Relatorios.Contar(o, null, Estados.Erro);
            _ign.Text = "Ignorados " + (o.Ignorados.Count + Relatorios.Excluidos(o));
            int ocultos = Relatorios.Ocultos(o);
            _rodape.Text = Relatorios.Resumo(o) + "\nClique numa linha para localizar no desenho; duplo clique abre o pop-up de ajuste. "
                           + (construir ? "Construir: você escolhe poste/vão e anotação — as regras do leitor não filtram nada aqui."
                                        : "Só aparece o que as regras do leitor classificaram" + (ocultos > 0 ? " (" + ocultos + " sem anotação: rede existente/postes sem anotação)" : "") + ".");
            _conexao.Text = Sincronizador.Status;
            Preencher();
        }

        private void Preencher()
        {
            if (_s == null || _s.Obra == null) return;
            bool ign = _filtro == Ignorados;
            _split.Visible = !ign;
            _listaIgnorados.Visible = ign;
            if (ign) { Visual.Preencher(_listaIgnorados, Relatorios.TextosIgnorados(_s.Obra)); return; }
            List<LinhaTabela> cabos = Relatorios.Cabos(_s.Obra, _filtro), outros = Relatorios.Outros(_s.Obra, _filtro);
            _nCabos = cabos.Count;
            _nOutros = outros.Count;
            Visual.Preencher(_cabos, cabos);
            Visual.Preencher(_outros, outros);
            _split.Panel1.Invalidate(true);
            _split.Panel2.Invalidate(true);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Hide(); }   // reabre com ORC_PAINEL
            base.OnFormClosing(e);
        }
    }
}
