// Plugin/Telas.cs — telas (WinForms, sem designer): início, tabelas (vínculos / ativos) e o pop-up do vão/poste.
// O painel de conferência fica em PainelObra.cs. Ações que mexem no desenho viram comandos (SendStringToExecute), para rodar
// no contexto de comando do nanoCAD.
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OrcamentoNanoCAD.Nucleo;

namespace OrcamentoNanoCAD.UI
{
    // ── tela inicial (mockup A) ──
    public enum Escolha { Nada, Construir, Ler, Continuar, Painel, Leitura }

    public class TelaInicio : Form
    {
        public Escolha Escolha = Escolha.Nada;
        public RegrasProjeto Regras;
        public string Projeto;
        private readonly ComboBox _regras;
        private readonly TextBox _projeto;
        private readonly Label _status;
        private readonly Sessao _s;

        public TelaInicio(Sessao s, PersistenciaDwg.ObraSalva salva)
        {
            _s = s;
            Text = "ReadGrid — orçamento no nanoCAD";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            AutoSize = true; AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Font = Visual.Normal;
            Padding = new Padding(16);

            FlowLayoutPanel raiz = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            raiz.Controls.Add(Visual.Rotulo("⚡ O que deseja fazer neste desenho?", Visual.Titulo));
            if (s.Obra != null)
            {
                FlowLayoutPanel aberta = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 6, 0, 0) };
                aberta.Controls.Add(Visual.Rotulo("Obra aberta: " + s.Obra.Nome + " (" + (s.Obra.Modo == "construir" ? "construir do zero" : "projeto lido") + ")"
                                                  + (s.Modificada ? " — com mudanças não salvas" : ""), Visual.Normal, s.Modificada ? Visual.Conferir : Visual.Suave));
                aberta.Controls.Add(Visual.Botao("↩ Voltar ao painel", (o, e) => Escolher(Escolha.Painel), true));
                raiz.Controls.Add(aberta);
            }
            FlowLayoutPanel cartoes = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 8, 0, 0) };
            cartoes.Controls.Add(Opcao("🧱  Construir do zero",
                "Você seleciona no desenho:\n•  os postes e suas anotações\n•  os vãos e suas anotações", "Para projetos sem padrão de layers/cores.", Escolha.Construir, false));
            cartoes.Controls.Add(Opcao("📖  Ler projeto",
                "Lê o desenho inteiro com as regras do projeto (layer + cor).\nMonta postes, vãos e vínculos sozinho.", "Sobra só a conferência e os ajustes.", Escolha.Ler, true));
            raiz.Controls.Add(cartoes);

            FlowLayoutPanel baixo = new FlowLayoutPanel { AutoSize = true, WrapContents = false, Margin = new Padding(0, 10, 0, 0) };
            FlowLayoutPanel colRegras = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown };
            colRegras.Controls.Add(Visual.Rotulo("REGRAS DO PROJETO", Visual.Secao, Visual.Suave));
            _regras = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 290 };
            PreencherRegras(s.Regras);
            colRegras.Controls.Add(_regras);
            FlowLayoutPanel botoesRegras = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            botoesRegras.Controls.Add(Visual.Botao("⚙ Configurar leitura…", (o, e) => Escolher(Escolha.Leitura)));
            botoesRegras.Controls.Add(Visual.Botao("🔗 Aplicação…", (o, e) =>
            {
                using (TelaConexao tc = new TelaConexao()) tc.ShowDialog(this);
                PreencherRegras(_regras.SelectedItem as RegrasProjeto);
                _status.Text = Sincronizador.Status;
            }));
            colRegras.Controls.Add(botoesRegras);
            FlowLayoutPanel colProj = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, Margin = new Padding(12, 0, 0, 0) };
            colProj.Controls.Add(Visual.Rotulo("PROJETO NA APLICAÇÃO", Visual.Secao, Visual.Suave));
            _projeto = new TextBox { Width = 90, Text = s.Obra != null ? s.Obra.Projeto : (s.Regras != null ? s.Regras.Codigo : "") };
            colProj.Controls.Add(_projeto);
            _regras.SelectedIndexChanged += (o, e) => { RegrasProjeto r = _regras.SelectedItem as RegrasProjeto; if (r != null && r.Codigo.Length > 0 && _projeto.Text != r.Codigo) _projeto.Text = r.Codigo; };
            // digitou o código de um projeto da aplicação: escolhe as regras dele
            _projeto.TextChanged += (o, e) =>
            {
                string cod = (_projeto.Text ?? "").Trim();
                foreach (RegrasProjeto r in _regras.Items)
                    if (r.DaAplicacao && r.Codigo == cod) { if (_regras.SelectedItem != r) _regras.SelectedItem = r; break; }
            };
            if (Conexao.Logado) AtualizarProjetos();
            FlowLayoutPanel colSalva = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, Margin = new Padding(12, 0, 0, 0) };
            colSalva.Controls.Add(Visual.Rotulo("OBRA SALVA NESTE DESENHO", Visual.Secao, Visual.Suave));
            colSalva.Controls.Add(Visual.Rotulo(salva == null ? "(nenhuma)" : ((salva.Quando ?? "") + " · " + salva.Onde), Visual.Mono));
            Button continuar = Visual.Botao("Continuar obra salva", (o, e) => Escolher(Escolha.Continuar), salva != null);
            continuar.Enabled = salva != null;
            colSalva.Controls.Add(continuar);
            baixo.Controls.Add(colRegras);
            baixo.Controls.Add(colProj);
            baixo.Controls.Add(colSalva);
            raiz.Controls.Add(baixo);
            _status = Visual.Rotulo(Sincronizador.Status, Visual.Normal, Visual.Suave);
            _status.MaximumSize = new Size(700, 0);
            _status.Margin = new Padding(0, 10, 0, 0);
            raiz.Controls.Add(_status);
            Label credito = Visual.Rotulo("ReadGrid " + ObraArquivo.VersaoPlugin + " · Desenvolvido por Gabriel Sales", Visual.Normal, Visual.Suave);
            credito.Margin = new Padding(0, 8, 0, 0);
            raiz.Controls.Add(credito);
            Controls.Add(raiz);
            Visual.Aplicar(this);
        }

        /// <summary>Busca a lista de projetos da aplicação em segundo plano (a tela abre na hora com a lista guardada).</summary>
        private void AtualizarProjetos()
        {
            System.Threading.ThreadPool.QueueUserWorkItem(_ =>
            {
                string erro = null;
                try { Conexao.BaixarProjetos(); }
                catch (Exception e) { erro = e.Message; }
                try
                {
                    if (IsDisposed || !IsHandleCreated) return;
                    BeginInvoke((Action)(() =>
                    {
                        if (IsDisposed) return;
                        string cod = (_projeto.Text ?? "").Trim();
                        PreencherRegras(_regras.SelectedItem as RegrasProjeto);
                        foreach (RegrasProjeto r in _regras.Items) if (r.DaAplicacao && r.Codigo == cod) { _regras.SelectedItem = r; break; }
                        if (erro != null) _status.Text = "Lista de projetos da aplicação indisponível (" + erro + "): mostrando a última guardada.";
                    }));
                }
                catch (Exception) { }
            });
        }

        private void PreencherRegras(RegrasProjeto atual)
        {
            _regras.Items.Clear();
            foreach (RegrasProjeto r in Aplicacao.RegrasParaEscolher()) _regras.Items.Add(r);
            RegrasProjeto sel = null;
            foreach (RegrasProjeto r in _regras.Items) if (atual != null && r.Nome == atual.Nome) sel = r;
            if (sel == null && _regras.Items.Count > 0) sel = (RegrasProjeto)_regras.Items[0];
            _regras.SelectedItem = sel;
        }

        /// <summary>Cartão clicável (o destacado = a opção recomendada, como no mockup).</summary>
        private Control Opcao(string titulo, string descricao, string rodape, Escolha escolha, bool destacado)
        {
            FlowLayoutPanel f = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = false, Width = 320, Height = 170, WrapContents = false, Padding = new Padding(14), Margin = new Padding(0, 0, 10, 0), Cursor = Cursors.Hand };
            f.Controls.Add(Visual.Rotulo(titulo, Visual.Grande));
            Label d = Visual.Rotulo(descricao, Visual.Normal, Visual.Suave);
            d.MaximumSize = new Size(290, 0); d.Margin = new Padding(0, 8, 0, 8);
            f.Controls.Add(d);
            f.Controls.Add(Visual.Rotulo(rodape, Visual.Normal, destacado ? Visual.Ok : Color.FromArgb(0x58, 0xa6, 0xff)));
            f.Tag = "cartao";
            Color borda = destacado ? Visual.Azul : Visual.Borda;
            f.Paint += (o, e) => { using (Pen p = new Pen(borda, destacado ? 2 : 1)) e.Graphics.DrawRectangle(p, 1, 1, f.Width - 2, f.Height - 2); };
            EventHandler clique = (o, e) => Escolher(escolha);
            f.Click += clique;
            foreach (Control c in f.Controls) { c.Click += clique; c.Cursor = Cursors.Hand; }
            return f;
        }

        private void Escolher(Escolha e)
        {
            if ((e == Escolha.Ler || e == Escolha.Construir || e == Escolha.Continuar) && _s.Obra != null && _s.Modificada
                && MessageBox.Show(this, "A obra aberta tem mudanças não salvas. Começar outra e descartar essas mudanças?", "Orçamento",
                                   MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            Escolha = e;
            Regras = _regras.SelectedItem as RegrasProjeto;
            Projeto = (_projeto.Text ?? "").Trim();
            DialogResult = DialogResult.OK;
            Close();
        }
    }

    // ── configuração da leitura (TASK-006): onde estão rede, postes e guias neste padrão de desenho ──
    public class TelaLeitura : Form
    {
        public ConfigLeitura Resultado;
        private readonly TextBox _mt, _bt, _blocos, _layersPoste, _layersGuia;
        private readonly CheckBox _cor;
        private readonly NumericUpDown _raio, _guia, _tol, _toque, _reserva;
        private readonly TextBox _placa;

        public TelaLeitura(RegrasProjeto regras, IEnumerable<string> layersDoDesenho, IEnumerable<string> blocosDoDesenho)
        {
            ConfigLeitura c = regras.Leitura ?? new ConfigLeitura();
            Text = "⚙ Configurar leitura — " + regras.Nome;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            AutoSize = true; AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Font = Visual.Normal;
            Padding = new Padding(16);
            FlowLayoutPanel f = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            f.Controls.Add(Visual.Rotulo("Onde o plugin procura a rede, os postes e as linhas guias neste padrão de desenho.\n"
                + "Listas separadas por vírgula; vale quando o nome CONTÉM o trecho (sem diferença de maiúsculas); \"=NOME\" = nome exato (ex.: =0).", Visual.Normal, Visual.Suave));
            f.Controls.Add(Visual.Rotulo(regras.DaAplicacao && Conexao.Logado
                ? (regras.LeituraDaAplicacao ? "✔ Guardada na aplicação para este projeto: vale em todos os computadores." : "Ao salvar, fica guardada na aplicação para este projeto (vale em todos os computadores).")
                : "Guardada só neste computador (para guardar na aplicação: conecte-se e escolha as regras \"Aplicação: …\").",
                Visual.Normal, regras.LeituraDaAplicacao ? Visual.Ok : Visual.Suave));
            _mt = Campo(f, "Layers da rede de MÉDIA tensão", string.Join(", ", c.LayersRedeMT.ToArray()));
            _bt = Campo(f, "Layers da rede de BAIXA tensão", string.Join(", ", c.LayersRedeBT.ToArray()));
            _cor = new CheckBox { Text = "A cor confirma a layer (vermelho = MT, verde = BT; cor contrária descarta a linha)", Checked = c.CorConfirma, AutoSize = true };
            f.Controls.Add(_cor);
            _blocos = Campo(f, "Blocos de POSTE (trecho do nome do bloco)", string.Join(", ", c.BlocosPoste.ToArray()));
            _layersPoste = Campo(f, "…ou layers de poste (qualquer bloco ou círculo nelas é poste)", string.Join(", ", c.LayersPoste.ToArray()));
            _layersGuia = Campo(f, "Layers das linhas guias (vazio = qualquer linha curta fora da rede; LEADER e MLEADER sempre valem)", string.Join(", ", c.LayersGuia.ToArray()));
            FlowLayoutPanel nums = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            _raio = Numero(nums, "Busca a partir do texto (× altura)", c.RaioAlturas);
            _guia = Numero(nums, "Guia de até (× altura)", c.GuiaMaxAlturas);
            _tol = Numero(nums, "Poste sobre a rede até (× altura)", c.TolPosteAlturas);
            f.Controls.Add(nums);
            FlowLayoutPanel nums2 = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            _toque = Numero(nums2, "Guia começa a até (× altura do texto)", c.GuiaToqueAlturas);
            _reserva = Numero(nums2, "Sem guia: elemento livre até (× altura)", c.RaioReservaAlturas);
            f.Controls.Add(nums2);
            _placa = Campo(f, "Placa do poste (expressão regular; vazio = não procurar)", c.PlacaRegex);
            Label ajuda = Visual.Rotulo("Layers deste desenho: " + Resumir(layersDoDesenho) + "\nBlocos deste desenho: " + Resumir(blocosDoDesenho), Visual.Normal, Visual.Suave);
            ajuda.MaximumSize = new Size(640, 0);
            ajuda.Margin = new Padding(0, 8, 0, 8);
            f.Controls.Add(ajuda);
            FlowLayoutPanel b = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            Button salvar = Visual.Botao("Salvar e reler", (o, e) => Salvar(), true);
            b.Controls.Add(salvar);
            b.Controls.Add(Visual.Botao("Padrão (Pernambuco)", (o, e) => { Resultado = new ConfigLeitura(); DialogResult = DialogResult.OK; Close(); }));
            b.Controls.Add(Visual.Botao("Cancelar", (o, e) => Close()));
            f.Controls.Add(b);
            Controls.Add(f);
            AcceptButton = salvar;
            Visual.Aplicar(this);
        }

        private static string Resumir(IEnumerable<string> l)
        {
            List<string> x = (l ?? new string[0]).Where(s => !string.IsNullOrEmpty(s)).Distinct().OrderBy(s => s).ToList();
            return x.Count == 0 ? "(leia o desenho primeiro)" : string.Join(", ", x.Take(40).ToArray()) + (x.Count > 40 ? "…" : "");
        }

        private static TextBox Campo(FlowLayoutPanel f, string rotulo, string valor)
        {
            f.Controls.Add(Visual.Rotulo(rotulo, Visual.Secao, Visual.Suave));
            TextBox t = new TextBox { Width = 640, Text = valor };
            f.Controls.Add(t);
            return t;
        }

        private static NumericUpDown Numero(FlowLayoutPanel f, string rotulo, double valor)
        {
            f.Controls.Add(Visual.Rotulo(rotulo, Visual.Normal, Visual.Suave));
            NumericUpDown n = new NumericUpDown { Minimum = 0.1m, Maximum = 1000, DecimalPlaces = 1, Increment = 0.5m, Width = 70, Value = (decimal)Math.Max(0.1, Math.Min(1000, valor)), Margin = new Padding(4, 0, 16, 0) };
            f.Controls.Add(n);
            return n;
        }

        private void Salvar()
        {
            ConfigLeitura c = new ConfigLeitura();
            c.LayersRedeMT = ConfigLeitura.Lista(_mt.Text);
            c.LayersRedeBT = ConfigLeitura.Lista(_bt.Text);
            c.CorConfirma = _cor.Checked;
            c.BlocosPoste = ConfigLeitura.Lista(_blocos.Text);
            c.LayersPoste = ConfigLeitura.Lista(_layersPoste.Text);
            c.LayersGuia = ConfigLeitura.Lista(_layersGuia.Text);
            c.RaioAlturas = (double)_raio.Value;
            c.GuiaMaxAlturas = (double)_guia.Value;
            c.TolPosteAlturas = (double)_tol.Value;
            c.GuiaToqueAlturas = (double)_toque.Value;
            c.RaioReservaAlturas = (double)_reserva.Value;
            c.PlacaRegex = (_placa.Text ?? "").Trim();
            if (c.LayersRedeMT.Count + c.LayersRedeBT.Count == 0 || c.BlocosPoste.Count + c.LayersPoste.Count == 0)
            {
                MessageBox.Show(this, "Informe ao menos uma layer de rede e um bloco (ou layer) de poste.", "Orçamento", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            Resultado = c;
            DialogResult = DialogResult.OK;
            Close();
        }
    }

    // ── tabelas: "Ver vínculos" e "Ver ativos com parâmetros" ──
    public class TelaTabela : Form
    {
        private readonly ListView _lista;
        private readonly Func<List<LinhaTabela>> _linhas;

        public TelaTabela(string titulo, string[] colunas, Func<List<LinhaTabela>> linhas, string rodape)
        {
            Text = titulo;
            Width = 1000; Height = 520;
            StartPosition = FormStartPosition.CenterScreen;
            Font = Visual.Normal;
            _linhas = linhas;
            _lista = Visual.Lista(colunas);
            _lista.ItemActivate += (o, e) => { if (_lista.SelectedItems.Count > 0) Visual.Executar("ORC_EDITAR " + _lista.SelectedItems[0].Tag); };
            _lista.Click += (o, e) => { if (_lista.SelectedItems.Count > 0) Visual.Executar("ORC_ZOOM " + _lista.SelectedItems[0].Tag); };
            Controls.Add(_lista);
            Label r = Visual.Rotulo(rodape, Visual.Normal, Visual.Suave);
            r.Dock = DockStyle.Bottom; r.AutoSize = false; r.Height = 36; r.Padding = new Padding(6);
            Controls.Add(r);
            Atualizar();
            Visual.Aplicar(this);
        }

        public void Atualizar() { Visual.Preencher(_lista, _linhas()); }
    }

    // ── pop-up do vão / poste ──
    public enum AcaoPopup { Nada, Aplicar, VoltarAoLido, Excluir }

    public class TelaElemento : Form
    {
        public AcaoPopup Acao = AcaoPopup.Nada;
        public string Operacao;
        public double? Metros;
        public int VinculoA = -2, VinculoB = -2;      // -2 = não mexer, -1 = automático, k = estrutura k do poste
        public bool AlterarLayer = true;

        private readonly Elemento _e;
        private readonly RegrasProjeto _regras;
        private readonly ComboBox _op;
        private readonly TextBox _metros;
        private readonly ComboBox _vincA, _vincB;
        private readonly CheckBox _layer;
        private readonly Label _aviso;
        private bool _mudaLayer;

        public TelaElemento(Obra obra, Elemento e, RegrasProjeto regras, bool podeExcluir)
        {
            _e = e;
            _regras = regras;
            Vao v = e as Vao;
            Poste p = e as Poste;
            Text = v != null ? "Vão " + v.Numero + " · " + (v.PosteA ?? "—") + " → " + (v.PosteB ?? "—") + "   [" + v.Operacao + "]"
                 : p != null ? "Poste " + p.Numero + (p.Lido.CodigoPoste != null ? " · " + p.Lido.CodigoPoste : "") + "   [" + p.Operacao + "]"
                 : "Texto " + e.Numero;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            AutoSize = true; AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Font = Visual.Normal;
            Padding = new Padding(10);

            TableLayoutPanel t = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Fill };
            FlowLayoutPanel entrada = Coluna("ENTRADA (DO DESENHO)"), saida = Coluna("SAÍDA (PARA A OBRA)");
            Campo(entrada, "Anotação (texto)", string.IsNullOrEmpty(e.TextoOriginal) ? "(sem anotação)" : e.TextoOriginal);
            Campo(entrada, "Layer · Cor do texto", (e.TextoLayer ?? "") + " · " + (e.TextoCor ?? ""));
            if (v != null)
            {
                Campo(entrada, "Linha", (v.LinhaLayer ?? "") + " · " + (v.Tipo == "MT" ? "média tensão" : v.Tipo == "BT" ? "baixa tensão" : "?"));
                Campo(entrada, "Postes", Descr(obra, v.PosteA) + " → " + Descr(obra, v.PosteB));
                string tipo = v.InfoCabo != null
                    ? v.InfoCabo.Codigo + " · tipocabo " + (v.InfoCabo.ValorCabo ?? "—") + " · " + (v.Cabo != null ? v.Cabo.Fases + " fase(s) · " + v.Cabo.Nivel : "")
                    : (v.Cabo != null && v.Cabo.Erro != null ? v.Cabo.Erro : "—");
                Campo(saida, "Ativo · tipo", tipo);
                Campo(saida, "Estruturas", Ponta(v.VinculoA, v.PosteA) + " ⇄ " + Ponta(v.VinculoB, v.PosteB));
                Campo(saida, "Ativo na obra", v.Ativo);
            }
            if (p != null)
            {
                Campo(entrada, "Bloco", string.IsNullOrEmpty(p.Handle) ? "(sem bloco)" : p.Bloco);
                Campo(saida, "Poste", Anotacoes.PosteOficial(p.Lido.CodigoPoste ?? "") ?? (p.Lido.CodigoPoste ?? "—"));
                for (int k = 0; k < p.Lido.Estruturas.Count; k++)
                {
                    string pars;
                    bool pend;
                    string ativo = Relatorios.AtivoDaEstrutura(p, k, out pars, out pend);
                    int cabos = k < p.CabosPorEstrutura.Count ? p.CabosPorEstrutura[k] : 0;
                    Campo(saida, p.Lido.Estruturas[k].Nome + " · " + cabos + " cabo(s)", ativo + "\n" + pars, pend ? Visual.Conferir : (Color?)null);
                }
            }
            if (e.Mensagens.Count > 0) Campo(saida, "Conferência", string.Join("\n", e.Mensagens.ToArray()), Visual.DoEstado(e.Estado));
            t.Controls.Add(entrada, 0, 0);
            t.Controls.Add(saida, 1, 0);

            // ajustes
            FlowLayoutPanel aj = Coluna("AJUSTES");
            FlowLayoutPanel linha = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            linha.Controls.Add(Visual.Rotulo("Operação", Visual.Normal));
            _op = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 150 };
            foreach (string op in regras.Operacoes()) _op.Items.Add(op);
            if (!_op.Items.Contains(e.Operacao ?? "")) _op.Items.Add(e.Operacao ?? "");
            _op.SelectedItem = e.Operacao ?? "";
            linha.Controls.Add(_op);
            if (v != null)
            {
                linha.Controls.Add(Visual.Rotulo("   Metros", Visual.Normal));
                _metros = new TextBox { Width = 70, Text = v.Metros.HasValue ? Anotacoes.FormatarMetros(v.Metros.Value) : "" };
                linha.Controls.Add(_metros);
            }
            aj.Controls.Add(linha);
            if (v != null && v.ParticipaVinculo)
            {
                FlowLayoutPanel lv = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
                _vincA = ComboVinculo(obra, v.PosteA, v.VinculoA);
                _vincB = ComboVinculo(obra, v.PosteB, v.VinculoB);
                lv.Controls.Add(Visual.Rotulo("Estrutura em " + (v.PosteA ?? "—"), Visual.Normal));
                lv.Controls.Add(_vincA);
                lv.Controls.Add(Visual.Rotulo("   em " + (v.PosteB ?? "—"), Visual.Normal));
                lv.Controls.Add(_vincB);
                aj.Controls.Add(lv);
            }
            _aviso = Visual.Rotulo("", Visual.Normal, Visual.Conferir);
            _layer = new CheckBox { Text = "Alterar a layer do texto no desenho", Checked = true, AutoSize = true };
            aj.Controls.Add(_aviso);
            aj.Controls.Add(_layer);
            _op.SelectedIndexChanged += (o, ev) => MostrarAviso();
            MostrarAviso();

            FlowLayoutPanel botoes = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            Button aplicar = Visual.Botao("Aplicar", (o, ev) => Fechar(AcaoPopup.Aplicar), true);
            botoes.Controls.Add(aplicar);
            botoes.Controls.Add(Visual.Botao("Voltar ao lido", (o, ev) => Fechar(AcaoPopup.VoltarAoLido)));
            if (podeExcluir) botoes.Controls.Add(Visual.Botao("Excluir da obra", (o, ev) => Fechar(AcaoPopup.Excluir)));
            botoes.Controls.Add(Visual.Botao("Fechar", (o, ev) => Fechar(AcaoPopup.Nada)));
            aj.Controls.Add(botoes);
            t.Controls.Add(aj, 0, 1);
            t.SetColumnSpan(aj, 2);
            Controls.Add(t);
            AcceptButton = aplicar;
            Visual.Aplicar(this);
        }

        private static string Descr(Obra obra, string numero)
        {
            Poste p = obra.PostePorNumero(numero);
            return p == null ? "—" : p.Numero + " (" + (p.Lido.CodigoPoste ?? "?") + ")";
        }

        private static string Ponta(Vinculo vi, string poste)
        {
            if (poste == null) return "sem poste";
            return vi == null ? "? (" + poste + ")" : vi.Estrutura + " (" + vi.Poste + ")" + (vi.Manual ? " ✋" : "");
        }

        private ComboBox ComboVinculo(Obra obra, string numPoste, Vinculo atual)
        {
            ComboBox c = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 130 };
            c.Items.Add("(automático)");
            Poste p = obra.PostePorNumero(numPoste);
            if (p != null) for (int k = 0; k < p.Lido.Estruturas.Count; k++) c.Items.Add(p.Lido.Estruturas[k].Nome + (k > 0 ? " #" + (k + 1) : ""));
            c.SelectedIndex = atual != null && atual.Manual && atual.K + 1 < c.Items.Count ? atual.K + 1 : 0;
            c.Enabled = p != null;
            c.Tag = c.SelectedIndex;
            return c;
        }

        private void MostrarAviso()
        {
            string nova = _op.SelectedItem as string ?? "";
            string camadaNova = _regras.LayerDaOperacao(nova);
            bool muda = nova != (_e.Operacao ?? "") && camadaNova != null && !string.IsNullOrEmpty(_e.TextoHandle)
                        && !string.Equals(camadaNova, _e.TextoLayer, StringComparison.OrdinalIgnoreCase);
            _mudaLayer = muda;
            _aviso.Text = muda ? "Ao aplicar, a anotação muda de layer: " + _e.TextoLayer + " → " + camadaNova + " (o desenho é alterado)." : "";
            _layer.Visible = muda;
        }

        private static FlowLayoutPanel Coluna(string titulo)
        {
            FlowLayoutPanel f = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Margin = new Padding(4, 4, 18, 4) };
            f.Controls.Add(Visual.Rotulo(titulo, Visual.Secao, Visual.Suave));
            return f;
        }

        private static void Campo(FlowLayoutPanel f, string rotulo, string valor, Color? cor = null)
        {
            f.Controls.Add(Visual.Rotulo(rotulo, Visual.Normal, Visual.Suave));
            Label v = Visual.Rotulo(valor ?? "", Visual.Mono, cor);
            v.MaximumSize = new Size(420, 0);
            v.Margin = new Padding(0, 0, 0, 6);
            f.Controls.Add(v);
        }

        private void Fechar(AcaoPopup a)
        {
            Acao = a;
            Operacao = _op.SelectedItem as string;
            if (_metros != null)
            {
                double? m = Anotacoes.Numero((_metros.Text ?? "").Trim());
                if (_metros.Text.Trim().Length > 0 && !m.HasValue)
                {
                    MessageBox.Show(this, "Metros inválido: " + _metros.Text, "Orçamento", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                Metros = m;
            }
            if (_vincA != null && _vincA.SelectedIndex != (int)_vincA.Tag) VinculoA = _vincA.SelectedIndex - 1;
            if (_vincB != null && _vincB.SelectedIndex != (int)_vincB.Tag) VinculoB = _vincB.SelectedIndex - 1;
            AlterarLayer = _mudaLayer && _layer.Checked;
            DialogResult = DialogResult.OK;
            Close();
        }
    }

    // ── ligação com a aplicação ──
    public class TelaConexao : Form
    {
        private readonly TextBox _url, _email, _senha;
        private readonly NumericUpDown _intervalo;
        private readonly CheckBox _auto, _atualizarPlugin;
        private readonly Label _situacao;

        public TelaConexao()
        {
            Text = "🔗 Ligação com o ReadGrid";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            AutoSize = true; AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Font = Visual.Normal;
            Padding = new Padding(12);
            ConexaoAplicacao c = Conexao.Dados;
            FlowLayoutPanel f = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            f.Controls.Add(Visual.Rotulo("Com a aplicação aberta, o plugin usa as regras do leitor e de vínculo GRAVADAS nela e confere a obra\n"
                + "com a Base de Orçamento (o que não está cadastrado vira Conferir). Mudou lá, aparece aqui em poucos segundos.", Visual.Normal, Visual.Suave));
            f.Controls.Add(Visual.Rotulo("Endereço (local: http://127.0.0.1:8000 · ou o endereço online)", Visual.Secao));
            _url = new TextBox { Width = 420, Text = c.Url };
            f.Controls.Add(_url);
            f.Controls.Add(Visual.Rotulo("E-mail", Visual.Secao));
            _email = new TextBox { Width = 420, Text = c.Email };
            f.Controls.Add(_email);
            f.Controls.Add(Visual.Rotulo("Senha (não é guardada; o acesso fica válido por até 30 dias)", Visual.Secao));
            _senha = new TextBox { Width = 420, UseSystemPasswordChar = true };
            f.Controls.Add(_senha);
            FlowLayoutPanel l = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            l.Controls.Add(Visual.Rotulo("Consultar a aplicação a cada", Visual.Normal));
            _intervalo = new NumericUpDown { Minimum = 2, Maximum = 300, Value = Math.Max(2, Math.Min(300, c.IntervaloSegundos)), Width = 60 };
            l.Controls.Add(_intervalo);
            l.Controls.Add(Visual.Rotulo("s", Visual.Normal));
            _auto = new CheckBox { Text = "atualizar sozinho", Checked = c.Automatico, AutoSize = true, Margin = new Padding(16, 4, 0, 0) };
            l.Controls.Add(_auto);
            f.Controls.Add(l);
            _atualizarPlugin = new CheckBox { Text = "Atualizar o plugin sozinho (baixa as versões novas publicadas na aplicação; valem ao reabrir o nanoCAD)",
                                              Checked = c.AtualizarPlugin, AutoSize = true, Margin = new Padding(0, 4, 0, 0) };
            f.Controls.Add(_atualizarPlugin);
            f.Controls.Add(Visual.Rotulo("Plugin " + ObraArquivo.VersaoPlugin + " · " + AtualizacaoPlugin.Status, Visual.Normal, Visual.Suave));
            _situacao = Visual.Rotulo(Conexao.Logado ? "Conectado como " + c.Email + " em " + c.Base : "Não conectado.", Visual.Normal, Conexao.Logado ? Visual.Ok : Visual.Suave);
            _situacao.MaximumSize = new Size(440, 0);
            f.Controls.Add(_situacao);
            FlowLayoutPanel b = new FlowLayoutPanel { AutoSize = true, WrapContents = false };
            Button entrar = Visual.Botao("Entrar", (o, e) => Entrar(), true);
            b.Controls.Add(entrar);
            b.Controls.Add(Visual.Botao("Salvar opções", (o, e) => { Conexao.Ajustar((int)_intervalo.Value, _auto.Checked, _atualizarPlugin.Checked); _situacao.Text = "Opções salvas."; }));
            b.Controls.Add(Visual.Botao("Desconectar", (o, e) => { Conexao.Sair(); _situacao.Text = "Desconectado: o plugin volta às regras embutidas."; _situacao.ForeColor = Visual.Suave; }));
            b.Controls.Add(Visual.Botao("Fechar", (o, e) => Close()));
            f.Controls.Add(b);
            Controls.Add(f);
            AcceptButton = entrar;
            Visual.Aplicar(this);
        }

        private void Entrar()
        {
            Cursor = Cursors.WaitCursor;
            try
            {
                Conexao.Entrar(_url.Text.Trim(), _email.Text.Trim(), _senha.Text, (int)_intervalo.Value, _auto.Checked);
                Conexao.Ajustar((int)_intervalo.Value, _auto.Checked, _atualizarPlugin.Checked);
                _senha.Text = "";
                _situacao.Text = "Conectado como " + Conexao.Dados.Email + " em " + Conexao.Dados.Base + ".";
                _situacao.ForeColor = Visual.Ok;
            }
            catch (ErroAplicacao e)
            {
                _situacao.Text = e.Status == 401 ? "E-mail ou senha não conferem." : "Não conectou: " + e.Message;
                _situacao.ForeColor = Visual.Erro;
            }
            finally { Cursor = Cursors.Default; }
        }
    }
}
