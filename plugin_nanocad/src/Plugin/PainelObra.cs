// Painel de auditoria e revisão da leitura.
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
        private Sessao _s;
        private Obra _ultima;
        private readonly TabControl _abas = new TabControl { Dock = DockStyle.Fill };
        private readonly ListView _pendencias, _cabos, _outros, _ignorados;
        private readonly TextBox _detalhe = new TextBox { Multiline = true, ReadOnly = true, Dock = DockStyle.Bottom, Height = 114, ScrollBars = ScrollBars.Vertical };
        private readonly ComboBox _poste = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 155 };
        private readonly ComboBox _estado = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 85 };
        private readonly CheckBox _mudancas = new CheckBox { Text = "Só mudanças", AutoSize = true };
        private readonly Label _titulo, _subtitulo;
        private readonly ToolTip _dicas = new ToolTip();
        private readonly TextBox _busca = new TextBox { Width = 145, AccessibleName = "Buscar texto, placa ou layer" };
        private readonly FlowLayoutPanel _botoesLer, _botoesConstruir;
        private readonly FlowLayoutPanel _editorManual, _rodapeManual;
        private readonly TextBox _textoManual = new TextBox { Width = 390, AccessibleName = "Anotação" };
        private readonly CheckBox _comentario = new CheckBox { Text = "Comentário", AutoSize = true };
        private readonly CheckBox _congelada = new CheckBox { Text = "Congelada", AutoSize = true };
        private readonly CheckBox _bloqueada = new CheckBox { Text = "Bloqueada", AutoSize = true };
        private readonly Label _contextoManual, _entregaManual;
        private readonly Button _aplicarManual, _enviarManual;
        private readonly List<RadioButton> _operacoesManuais = new List<RadioButton>();
        private Elemento _linhaManual;
        private readonly ComboBox _regrasManuais = new ComboBox { Width = 330, DropDownStyle = ComboBoxStyle.DropDownList };
        private readonly Button _editar, _associar, _conferir, _lote;
        private bool _preenchendo;
        private Dictionary<string, TextoIgnorado> _fora = new Dictionary<string, TextoIgnorado>();
        private Dictionary<string, string> _avisos = new Dictionary<string, string>();

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
            if (_aberto != null && !_aberto.IsDisposed) _aberto._dicas.SetToolTip(_aberto._subtitulo, Sincronizador.Status);
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
            Width = 670; Height = Math.Min(820, Screen.PrimaryScreen.WorkingArea.Height);
            MinimumSize = new Size(500, 470);
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
            _dicas.SetToolTip(_subtitulo, Sincronizador.Status);

            FlowLayoutPanel barra = Linha();
            MenuStrip menu = new MenuStrip { AutoSize = true, Dock = DockStyle.None, GripStyle = ToolStripGripStyle.Hidden };
            ToolStripMenuItem projeto = Grupo(menu, "Projeto");
            Acao(projeto, "Início", "ORCACORE");
            Acao(projeto, "Nome e número da obra", "ORC_PROJETO");
            Acao(projeto, "Fechar obra", "ORC_FECHAR");
            ToolStripMenuItem notas = Grupo(menu, "Anotações");
            Acao(notas, "Criar nota…", "ORC_NOTA");
            Acao(notas, "Editar nota…", "ORC_EDITAR_NOTA");
            Acao(notas, "Mover nota…", "ORC_MOVER_NOTA");
            Acao(notas, "Reapontar guia…", "ORC_REAPONTAR_NOTA");
            Acao(notas, "Atualizar guias", "ORC_ATUALIZAR_GUIAS");
            Acao(notas, "Auditoria de layers…", "ORC_AUDITAR_LAYERS");
            ToolStripMenuItem conferir = Grupo(menu, "Conferir");
            Acao(conferir, "Vínculos", "ORC_VINCULOS");
            Acao(conferir, "Ativos com parâmetros", "ORC_ATIVOS");
            Acao(conferir, "Mostrar ou ocultar marcas", "ORC_MARCAS");
            ToolStripMenuItem config = Grupo(menu, "Configurações");
            Acao(config, "Leitura do desenho", "ORC_LEITURA");
            Acao(config, "Conectar à aplicação", "ORC_CONECTAR");
            Acao(config, "Sincronizar regras", "ORC_SINCRONIZAR");
            Acao(config, "Diagnóstico", "ORC_DIAGNOSTICO");
            Acao(projeto, "Reler do zero…", "ORC_RELER");
            conferir.DropDownItems.Add("Cores por operação · I vermelho / R magenta / M azul / * LV",null,(o,e) => {
                if (!Ativa) return;
                _s.AnaliseCores = !_s.AnaliseCores;
                if (_s.AnaliseCores) _s.MarcasVisiveis = true;
                _s.Atualizar();
            });
            barra.Controls.Add(menu);
            barra.Controls.Add(Visual.Botao("Salvar", (o, e) => Visual.Executar("ORC_SALVAR"), false, "salvar"));
            barra.Controls.Add(Visual.Botao("Abrir no OrçaCore", (o, e) => Visual.Executar("ORC_ABRIR_APP"), false, "app"));
            topo.Controls.Add(barra);

            _botoesLer = Linha();
            _botoesLer.Controls.Add(Visual.Botao("Atualizar leitura", (o, e) => Visual.Executar("ORC_ATUALIZAR"), true, "atualizar"));

            _botoesLer.Controls.Add(Visual.Botao("+ Cabo", (o, e) => Visual.Executar("ORC_NOVO_CABO"), false, "cabos"));
            _botoesLer.Controls.Add(Visual.Botao("+ Outros", (o, e) => Visual.Executar("ORC_NOVO_OUTROS"), false, "ativos"));
            _botoesConstruir = Linha();
            _botoesConstruir.Controls.Add(Visual.Rotulo("Projeto e regras", Visual.Normal));
            _botoesConstruir.Controls.Add(_regrasManuais);
            _botoesConstruir.SetFlowBreak(_regrasManuais, true);
            _regrasManuais.SelectionChangeCommitted += (o,e) =>
            {
                var regras = _regrasManuais.SelectedItem as RegrasProjeto;
                if (Ativa && regras != null && regras.Codigo != _s.Obra.Projeto)
                    Visual.Executar("ORC_REGRAS_MONTAGEM " + regras.Codigo);
            };
            _botoesConstruir.Controls.Add(Visual.Botao("Ler por região", (o, e) => Visual.Executar("ORC_REGIAO"), true));
            // TASK-038: no formato do mapa do OrçaCore (ponto = «@nome - comentário» + campos com operação; cabo = linha/trecho + campos)
            _botoesConstruir.Controls.Add(Visual.Botao("Selecionar poste", (o, e) => Visual.Executar("ORC_PONTO_MAPA"), false, "poste"));
            _botoesConstruir.Controls.Add(Visual.Botao("Selecionar cabo", (o, e) => Visual.Executar("ORC_CABO_MAPA"), false, "vao"));
            _botoesConstruir.Controls.Add(Visual.Botao("Regras / sem regras", (o, e) => Visual.Executar("ORC_MODO_MANUAL")));
            _botoesConstruir.Controls.Add(Visual.Botao("Limpar", (o, e) => Visual.Executar("ORC_LIMPAR_MONTAGEM")));
            _botoesConstruir.Controls.Add(Visual.Botao("Vincular automático", (o, e) => Visual.Executar("ORC_VINCULAR"), false, "vincular"));
            _botoesConstruir.Controls.Add(Visual.Botao("Selecionar cabos", (o,e) => Visual.Executar("ORC_SELECIONAR_CABOS")));
            _botoesConstruir.Controls.Add(Visual.Botao("+ Cabo", (o, e) => Visual.Executar("ORC_NOVO_CABO"), false, "cabos"));
            _botoesConstruir.Controls.Add(Visual.Botao("+ Outros", (o, e) => Visual.Executar("ORC_NOVO_OUTROS"), false, "ativos"));
            topo.Controls.Add(_botoesLer);
            topo.Controls.Add(_botoesConstruir);

            FlowLayoutPanel filtros = Linha();
            _estado.Items.AddRange(new object[] { "Todos", Estados.Ok, Estados.Conferir, Estados.Erro });
            _estado.SelectedIndex = 0;
            filtros.Controls.Add(Visual.Rotulo("Buscar", Visual.Normal));
            filtros.Controls.Add(_busca);
            filtros.Controls.Add(_poste); filtros.Controls.Add(_estado); filtros.Controls.Add(_mudancas);
            topo.Controls.Add(filtros);
            _pendencias = Aba("Pendências", new[] { "Item", "Operação", "Conferir" });
            _outros = Aba("Postes e estruturas", Relatorios.ColunasOutros);
            _cabos = Aba("Vãos e cabos", Relatorios.ColunasCabos);
            _ignorados = Aba("Fora da obra", Relatorios.ColunasIgnorados);
            FlowLayoutPanel acoes = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, Padding = new Padding(5) };
            acoes.Controls.Add(Visual.Botao("Anterior", (o,e) => Avancar(-1)));
            acoes.Controls.Add(Visual.Botao("Próxima", (o,e) => Avancar(1)));
            _editar = Visual.Botao("Ajustar", (o,e) => Comando("ORC_EDITAR"));
            _associar = Visual.Botao("Escolher destino", (o,e) => Comando("ORC_REASSOCIAR"));
            _conferir = Visual.Botao("Conferido", (o,e) => Conferir());
            _lote = Visual.Botao("Operação em lote…", (o,e) => Lote());
            acoes.Controls.AddRange(new Control[] { _editar, _associar, _conferir, _lote });
            _editorManual = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, WrapContents = false,
                FlowDirection = FlowDirection.TopDown, Padding = new Padding(10, 5, 10, 5), Tag = "cartao" };
            _contextoManual = Visual.Rotulo("Selecione uma anotação", Visual.Secao);
            _editorManual.Controls.Add(_contextoManual);
            _editorManual.Controls.Add(_textoManual);
            FlowLayoutPanel operacoes = Linha();
            operacoes.Controls.Add(Visual.Rotulo("Operação", Visual.Normal));
            foreach (string op in MontagemManual.Operacoes)
            {
                var b = new RadioButton { Text = op, Appearance = Appearance.Button, AutoSize = false,
                    Width = 43, Height = 27, TextAlign = ContentAlignment.MiddleCenter, Margin = new Padding(1) };
                _operacoesManuais.Add(b);
                operacoes.Controls.Add(b);
            }
            _editorManual.Controls.Add(operacoes);
            FlowLayoutPanel marcas = Linha();
            _aplicarManual = Visual.Botao("Aplicar edição", (o, e) => AplicarEdicaoManual());
            marcas.Controls.AddRange(new Control[] { _comentario, _congelada, _bloqueada, _aplicarManual });
            _editorManual.Controls.Add(marcas);
            _comentario.CheckedChanged += (o, e) => { if (_comentario.Checked) _congelada.Checked = true; EstadoEditorManual(); };
            _bloqueada.CheckedChanged += (o, e) => EstadoEditorManual();

            _rodapeManual = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, WrapContents = false,
                FlowDirection = FlowDirection.TopDown, Padding = new Padding(10, 5, 10, 5) };
            _entregaManual = Visual.Rotulo("Destino · obra aberta no OrçaCore", Visual.Normal, Visual.Suave);
            _entregaManual.MaximumSize = new Size(550, 0);
            _enviarManual = Visual.Botao("Adicionar anotações ao OrçaCore", (o, e) => EnviarAnotacoes(), true, "app");
            _rodapeManual.Controls.Add(_entregaManual);
            _rodapeManual.Controls.Add(_enviarManual);
            Controls.Add(_abas); Controls.Add(_detalhe); Controls.Add(acoes);
            Controls.Add(_editorManual); Controls.Add(_rodapeManual); Controls.Add(topo);
            _abas.SelectedIndexChanged += (o,e) => Selecao();
            _estado.SelectedIndexChanged += (o,e) => Preencher();
            _poste.SelectedIndexChanged += (o,e) => Preencher();
            _mudancas.CheckedChanged += (o,e) => Preencher();
            _busca.TextChanged += (o,e) => Preencher();
            Visual.Aplicar(this);
        }

        private ListView Aba(string titulo, string[] colunas)
        {
            TabPage tab = new TabPage(titulo);
            ListView lista = Visual.Lista(colunas);
            lista.MultiSelect = true;
            tab.Controls.Add(lista);
            _abas.TabPages.Add(tab);
            lista.SelectedIndexChanged += (o,e) => Selecao();
            lista.ItemActivate += (o,e) => Comando("ORC_ZOOM");
            return lista;
        }

        private ListView Lista
        {
            get
            {
                // Mostrar preenche o painel antes de o nanoCAD exibir o formulário.
                TabPage aba = _abas.SelectedTab;
                return aba == null ? _pendencias : (ListView)aba.Controls[0];
            }
        }
        private string Id { get { return Lista.SelectedItems.Count == 0 ? null : Lista.SelectedItems[0].Tag as string; } }
        private bool Ativa { get { return _s != null && _s.Obra != null && Aplicacao.Doc() == _s.Doc; } }
        private void Comando(string cmd)
        {
            if (!Ativa || Id == null) return;
            TextoIgnorado t;
            Visual.Executar(cmd + " " + (_fora.TryGetValue(Id, out t) ? "H:" + t.Handle : Id));
        }

        private void Mostrar(Sessao s)
        {
            if (s.Obra == null) return;
            bool mudou = _s != s || _ultima != s.Obra;
            string poste = _s == s ? _poste.Text : "";
            _s = s; _ultima = s.Obra; _preenchendo = true;
            bool manual = s.Obra.Modo == "construir";
            _titulo.Text = (manual ? "Montar manualmente · " : "") + s.Obra.Nome + (s.Modificada ? "   • Alterações por salvar" : "");
            _subtitulo.Text = s.Regras.Nome + " · " + s.Obra.Projeto + "   |   " + s.Regras.Origem;
            _botoesLer.Visible = !manual; _botoesConstruir.Visible = manual;
            _editorManual.Visible = _rodapeManual.Visible = manual;
            _detalhe.Visible = !manual;
            if (manual)
            {
                _regrasManuais.Items.Clear();
                _regrasManuais.Items.Add(s.Regras);
                foreach (RegrasProjeto r in Aplicacao.RegrasParaEscolher().Where(r => r.Codigo != s.Regras.Codigo && !string.IsNullOrEmpty(r.Codigo)))
                    _regrasManuais.Items.Add(r);
                _regrasManuais.SelectedIndex = 0;
                _regrasManuais.Enabled = !s.Obra.Elementos.Any() && !s.EnviandoAnotacoes;
            }
            if (mudou && !s.EnviandoAnotacoes) _entregaManual.Text = "Destino · obra aberta no OrçaCore";
            _enviarManual.Enabled = _botoesConstruir.Enabled = !s.EnviandoAnotacoes;
            _poste.Items.Clear(); _poste.Items.Add("Todos os postes");
            foreach (Poste p in s.Obra.Postes.Where(x => x.IrmaoDe == null))
                _poste.Items.Add((p.Numero + " " + p.Placa).Trim());
            _poste.SelectedIndex = Math.Max(0, _poste.Items.IndexOf(poste));
            _mudancas.Enabled = s.Obra.LeituraAnterior != null;
            if (!_mudancas.Enabled) _mudancas.Checked = false;
            _preenchendo = false;
            Preencher();
            if (mudou)
            {
                _abas.SelectedIndex = manual ? 1 : _pendencias.Items.Count > 0 ? 0 : (_outros.Items.Count > 0 ? 1 : 2);
                if (Lista.SelectedItems.Count == 0 && Lista.Items.Count > 0) Lista.Items[0].Selected = true;
            }
        }

        private bool NoPoste(Elemento e)
        {
            if (_poste.SelectedIndex <= 0) return true;
            string n = _poste.Text.Split(' ')[0];
            Vao v = e as Vao;
            return e.Numero == n || e.IrmaoDe == n || (v != null && (v.PosteA == n || v.PosteB == n));
        }

        private bool Corresponde(params string[] campos)
        {
            string termo = _busca.Text.Trim();
            return termo.Length == 0 || campos.Any(c => (c ?? "").IndexOf(termo, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private void Preencher()
        {
            if (_preenchendo || _s == null || _s.Obra == null) return;
            _preenchendo = true;
            try
            {
                Obra o = _s.Obra;
                var layers = _s.Regras.OperacoesPorLayer();
                var elementos = Relatorios.DoPainel(o, null).ToList();
                var textos = new HashSet<string>(elementos.Select(e => (e.TextoOriginal ?? "").Trim()), StringComparer.OrdinalIgnoreCase);
                var mostrar = new HashSet<string>();
                var pend = new List<LinhaTabela>();
                _avisos.Clear();
                foreach (Elemento e in elementos)
                {
                    Poste p = e as Poste;
                    if (!Corresponde(e.Numero, e.TextoOriginal, e.LayerAtual, e.AtivoLido, p == null ? null : p.Placa) || !NoPoste(e) || (_estado.SelectedIndex > 0 && _estado.Text != e.Estado)) continue;
                    string chave = Auditoria.Chave(e);
                    string conferido = o.Conferidos.Str(chave);
                    string assinatura = _mudancas.Checked || conferido != null ? Auditoria.Assinatura(e) : null;
                    string anterior = o.LeituraAnterior == null ? null : o.LeituraAnterior.Str(chave);
                    if (_mudancas.Checked && anterior == assinatura) continue;
                    mostrar.Add(e.Numero);
                    string aviso = string.Join("; ", Auditoria.Motivos(o, e, _s.Regras,layers).Distinct().ToArray());
                    if (aviso.Length == 0 && _mudancas.Checked) aviso = anterior == null ? "Novo na leitura" : "Alterado desde a leitura anterior";
                    _avisos[e.Numero] = aviso;
                    if (aviso.Length == 0 || conferido == assinatura + aviso) continue;
                    pend.Add(new LinhaTabela { Chave = e.Numero, Estado = e.Estado == Estados.Erro ? Estados.Erro : Estados.Conferir,
                        Celulas = new[] { e.Numero, e.Operacao, aviso } });
                }
                var fora = Relatorios.TextosIgnorados(o).Where(x => !x.Chave.StartsWith("H:") && NoPoste(o.PorNumero(x.Chave))).ToList();
                _fora.Clear();
                foreach (TextoIgnorado t in o.Ignorados)
                {
                    string id = "H:" + t.Handle + "/" + t.Parte;
                    _fora[id] = t;
                    if (_poste.SelectedIndex > 0 || _mudancas.Checked || _estado.SelectedIndex > 0 || !Corresponde(t.Texto,t.Layer)) continue;
                    fora.Add(new LinhaTabela { Chave = id, Celulas = new[] { t.Texto, t.Layer, t.Cor, t.Motivo } });
                    if (Auditoria.Suspeito(o,t,_s.Regras,textos,layers) && o.Conferidos.Str(id) != AssinaturaIgnorado(t))
                        pend.Add(new LinhaTabela { Chave = id, Estado = Estados.Conferir, Celulas = new[] { t.Texto, "—", t.Motivo } });
                }
                Popular(_pendencias, pend.OrderByDescending(x => x.Estado == Estados.Erro).ToList());
                Popular(_outros, LinhasManuais(Relatorios.Outros(o,null), o, _outros, false).Where(x => mostrar.Contains(x.Chave)).ToList());
                Popular(_cabos, LinhasManuais(Relatorios.Cabos(o,null), o, _cabos, true).Where(x => mostrar.Contains(x.Chave)).ToList());
                Popular(_ignorados, fora);
                List<Elemento> cabos, outros;
                ObraArquivo.ElementosDasLinhas(o, out cabos, out outros);
                _abas.TabPages[0].Text = "Pendências (" + pend.Count + ")";
                _abas.TabPages[1].Text = (o.Modo == "construir" ? "Postes e outros" : "Postes e estruturas") + " (" + outros.Count + ")";
                _abas.TabPages[2].Text = "Vãos e cabos (" + cabos.Count + ")";
                _abas.TabPages[3].Text = "Fora (" + (o.Ignorados.Count + Relatorios.Excluidos(o)) + ")";
            }
            finally { _preenchendo = false; }
            Selecao();
        }

        private static string AssinaturaIgnorado(TextoIgnorado t)
        {
            return OrigemIgnorado(t) + "|" + t.Motivo;
        }

        private static string OrigemIgnorado(TextoIgnorado t)
        {
            return Auditoria.Origem(t.Layer, t.Parte, t.Texto) + "|" + t.X + "|" + t.Y;
        }

        private static void Popular(ListView lista, List<LinhaTabela> linhas)
        {
            var ids = new HashSet<string>(lista.SelectedItems.Cast<ListViewItem>().Select(i => (string)i.Tag));
            int topo = lista.Items.Count > 0 && lista.TopItem != null ? lista.TopItem.Index : 0;
            Visual.Preencher(lista, linhas);
            foreach (ListViewItem i in lista.Items) if (ids.Contains((string)i.Tag)) i.Selected = true;
            if (lista.Visible && lista.Items.Count > 0) lista.TopItem = lista.Items[Math.Min(topo, lista.Items.Count-1)];
        }

        private void Selecao()
        {
            if (_preenchendo || _s == null || _s.Obra == null) return;
            string id = Id;
            Elemento e = _s.Obra.PorNumero(id);
            TextoIgnorado t = null;
            if (id != null) _fora.TryGetValue(id, out t);
            MostrarEdicaoManual(e);
            _editar.Enabled = e != null && (!e.Bloqueada && !e.EnviadaNaMontagem && !_s.EnviandoAnotacoes);
            _editar.Text = e != null && e.Excluido ? "Restaurar…" : e is Vao ? "Ajustar vão…" : "Ajustar…";
            _associar.Enabled = _s.Obra.Modo != "construir" && e != null && e.Candidatos.Count > 1;
            _conferir.Enabled = id != null && Lista == _pendencias;
            _lote.Enabled = e != null && !e.Excluido && (!e.Bloqueada && !e.EnviadaNaMontagem && !_s.EnviandoAnotacoes);
            _detalhe.Text = e != null ? Auditoria.Detalhe(_s.Obra,e) + Environment.NewLine + (_avisos.ContainsKey(id) ? _avisos[id] : "") :
                t != null ? t.Texto + Environment.NewLine + "Layer " + t.Layer + " · " + t.Cor + Environment.NewLine + t.Motivo + " · Fora da exportação" : "";
            if (Ativa && Visible) _s.Destacar(e,t);
        }

        private static List<LinhaTabela> LinhasManuais(List<LinhaTabela> linhas, Obra obra, ListView lista, bool cabos)
        {
            bool manual = obra.Modo == "construir";
            string[] colunas = manual
                ? (cabos ? new[] { "", "Vão", "Rede", "Op.", "Anotação", "m", "Postes", "Marcas" }
                         : new[] { "", "Nº", "Op.", "Anotação", "Layer CAD", "Cor", "Marcas" })
                : (cabos ? Relatorios.ColunasCabos : Relatorios.ColunasOutros);
            if (!lista.Columns.Cast<ColumnHeader>().Select(c => c.Text).SequenceEqual(colunas))
            {
                lista.Columns.Clear();
                foreach (string coluna in colunas) lista.Columns.Add(coluna);
            }
            if (!manual) return linhas;
            foreach (LinhaTabela l in linhas)
            {
                Elemento e = obra.PorNumero(l.Chave);
                if (e == null) continue;
                var marcas = new List<string>();
                if (e.Comentario) marcas.Add("Comentário");
                else if (e.Congelada) marcas.Add("Congelada");
                if (e.Bloqueada) marcas.Add("Bloqueada");
                if (e.EnviadaNaMontagem) marcas.Add("Enviada");
                string estado = string.Join(", ", marcas.ToArray());
                if (cabos) l.Celulas[7] = estado;
                else
                {
                    l.Celulas = new[] { "", e.Numero, e.Operacao, l.Celulas[4], e.LayerAtual, e.TextoCor, estado };
                    l.Icones = new[] { "estado:" + e.Estado, null, "op:" + e.Operacao, null, null, "cor:" + e.TextoCor, null };
                }
            }
            return linhas;
        }

        private void MostrarEdicaoManual(Elemento e)
        {
            if (_s.Obra.Modo != "construir") return;
            if (e != null && e.Excluido) e = null;
            _linhaManual = e;
            Poste p = e as Poste;
            _contextoManual.Text = e == null ? "Selecione uma anotação para editar" :
                (p == null ? e.Numero : "Poste " + (p.IrmaoDe ?? p.Numero)) +
                (string.IsNullOrEmpty(e.LayerAtual) ? "" : " · " + e.LayerAtual + " · " + e.TextoCor);
            _textoManual.Text = e == null ? "" : e.AtivoLido;
            _comentario.Checked = e != null && e.Comentario;
            _congelada.Checked = e != null && (e.Congelada || e.Comentario);
            _bloqueada.Checked = e != null && e.Bloqueada;
            foreach (RadioButton b in _operacoesManuais) b.Checked = e != null && b.Text == e.Operacao;
            if (e != null && !_operacoesManuais.Any(b => b.Checked)) _operacoesManuais[0].Checked = true;
            EstadoEditorManual();
        }

        private void EstadoEditorManual()
        {
            bool ativa = _linhaManual != null && !_linhaManual.EnviadaNaMontagem && _s != null && !_s.EnviandoAnotacoes;
            bool editavel = ativa && !_linhaManual.Bloqueada;
            _textoManual.ReadOnly = !editavel;
            foreach (RadioButton b in _operacoesManuais) b.Enabled = editavel;
            _comentario.Enabled = editavel;
            _congelada.Enabled = editavel && !_comentario.Checked;
            _bloqueada.Enabled = ativa;
            _aplicarManual.Enabled = ativa && (editavel || !_bloqueada.Checked);
        }

        private void AplicarEdicaoManual()
        {
            if (!Ativa || !_aplicarManual.Enabled) return;
            var op = _operacoesManuais.FirstOrDefault(b => b.Checked);
            if (op == null || string.IsNullOrEmpty((_textoManual.Text ?? "").Trim())) return;
            MontagemManual.Editar(_s.Obra, _linhaManual, _textoManual.Text.Trim(), op.Text,
                _comentario.Checked, _congelada.Checked, _bloqueada.Checked, _s.Regras);
            _s.Modificada = true;
            _s.Revisao++;
            _s.Atualizar();
        }

        private void EnviarAnotacoes()
        {
            if (!Ativa || _s.EnviandoAnotacoes) return;
            Sessao s = _s;
            Obra atual = s.Obra;
            List<Elemento> cabos, outros;
            ObraArquivo.ElementosDasLinhas(atual, out cabos, out outros);
            var novas = new HashSet<Elemento>(cabos.Concat(outros).Where(e => !e.EnviadaNaMontagem));
            var pontosNovos = atual.PontosMapa.Where(p => !p.Enviado).ToList();                         // TASK-038
            var cabosNovos = atual.CabosMapa.Where(c => !c.Enviado).ToList();
            int total = novas.Count + pontosNovos.Count + cabosNovos.Count;
            if (total == 0) { _entregaManual.Text = "Nenhuma anotação nova para enviar."; return; }
            var obra = new Obra { Projeto = atual.Projeto, Nome = atual.Nome, Modo = atual.Modo,
                Postes = atual.Postes.Where(novas.Contains).ToList(), Vaos = atual.Vaos.Where(novas.Contains).ToList(),
                Avulsos = atual.Avulsos.Where(novas.Contains).ToList(), PontosMapa = pontosNovos, CabosMapa = cabosNovos };
            var env = new JObj().Set("formato", ObraArquivo.Formato).Set("versao", ObraArquivo.Versao)
                .Set("tipo", "obra").Set("projeto", obra.Projeto).Set("dados", ObraArquivo.Dados(obra));
            string conteudo = Json.Escrever(env, false);
            if (s.LoteAnotacoes == null || s.ConteudoAnotacoes != conteudo)
            {
                s.LoteAnotacoes = Guid.NewGuid().ToString("N");
                s.ConteudoAnotacoes = conteudo;
            }
            string lote = s.LoteAnotacoes, chave = s.Arquivo;
            s.EnviandoAnotacoes = true;
            _enviarManual.Enabled = _botoesConstruir.Enabled = false;
            EstadoEditorManual();
            _entregaManual.Text = "Enviando " + total + " anotações…";
            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                string erro;
                try { erro = ProgramaNoComputador.AcrescentarAnotacoes(chave, env, lote); }
                catch (Exception ex) { erro = ex.Message; }
                try
                {
                    BeginInvoke((Action)(() =>
                    {
                        s.EnviandoAnotacoes = false;
                        if (erro == null && s.Obra == atual)
                        {
                            foreach (Elemento linha in novas) linha.EnviadaNaMontagem = true;
                            foreach (PontoMapa p in pontosNovos) p.Enviado = true;
                            foreach (CaboMapa c in cabosNovos) c.Enviado = true;
                        }
                        if (_s != s || s.Obra != atual) return;
                        Mostrar(s);
                        _entregaManual.Text = erro ?? (total + " anotações acrescentadas à obra aberta"
                            + (pontosNovos.Count + cabosNovos.Count > 0 ? " (" + pontosNovos.Count + " ponto(s) e " + cabosNovos.Count + " cabo(s) no formato do mapa)." : "."));
                    }));
                }
                catch (InvalidOperationException) { s.EnviandoAnotacoes = false; }
            });
        }

        private void Avancar(int passo)
        {
            if (Lista.Items.Count == 0) return;
            int i = Lista.SelectedIndices.Count == 0 ? 0 : Math.Max(0,Math.Min(Lista.Items.Count-1, Lista.SelectedIndices[0]+passo));
            Lista.SelectedIndices.Clear(); Lista.Items[i].Selected = true; Lista.Items[i].EnsureVisible(); Lista.Focus();
        }

        private void Conferir()
        {
            if (!Ativa) return;
            int indice = Lista.SelectedIndices.Count > 0 ? Lista.SelectedIndices[0] : 0;
            foreach (ListViewItem item in Lista.SelectedItems)
            {
                string id = (string)item.Tag;
                TextoIgnorado t;
                if (_fora.TryGetValue(id,out t)) _s.Obra.Conferidos.Set(id,AssinaturaIgnorado(t));
                else
                {
                    Elemento e = _s.Obra.PorNumero(id);
                    if (e != null) _s.Obra.Conferidos.Set(Auditoria.Chave(e), Auditoria.Assinatura(e) + _avisos[id]);
                }
            }
            _s.Modificada = true; Mostrar(_s);
            if (Lista.Items.Count > 0) { var item = Lista.Items[Math.Min(indice,Lista.Items.Count-1)]; item.Selected = true; item.EnsureVisible(); }
        }

        private void Lote()
        {
            if (!Ativa) return;
            var itens = Lista.SelectedItems.Cast<ListViewItem>().Select(i => _s.Obra.PorNumero((string)i.Tag)).ToList();
            if (itens.Count == 0 || itens.Any(e => e == null || e.Excluido || e.Bloqueada || e.EnviadaNaMontagem)) return;
            using (var janela = new LoteOperacao(_s, itens))
                if (janela.ShowDialog(this) == DialogResult.OK && Ativa) Comandos.AplicarLote(_s,itens,janela.Operacao,janela.AlterarLayer);
        }

        protected override void OnVisibleChanged(EventArgs e)
        {
            base.OnVisibleChanged(e);
            if (!Visible && Ativa)
            {
                _s.Destacar(null,null);
                if (_s.AnaliseCores) { _s.AnaliseCores = false; _s.Atualizar(); }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing) { e.Cancel = true; Hide(); }
            base.OnFormClosing(e);
        }
    }
}
