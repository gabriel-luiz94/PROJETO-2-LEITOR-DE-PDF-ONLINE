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
    public static class Visual
    {
        public static readonly Color Ok = Color.FromArgb(30, 130, 60);
        public static readonly Color Conferir = Color.FromArgb(190, 120, 0);
        public static readonly Color Erro = Color.FromArgb(200, 40, 40);
        public static readonly Font Titulo = new Font("Segoe UI", 11f, FontStyle.Bold);
        public static readonly Font Normal = new Font("Segoe UI", 9f);
        public static readonly Font Mono = new Font("Consolas", 9.5f);
        public static readonly Font Secao = new Font("Segoe UI", 8f, FontStyle.Bold);

        public static Color DoEstado(string estado)
        {
            if (estado == Estados.Erro) return Erro;
            if (estado == Estados.Conferir) return Conferir;
            return Ok;
        }

        public static Label Rotulo(string texto, Font fonte, Color? cor = null)
        {
            Label l = new Label { Text = texto, Font = fonte ?? Normal, AutoSize = true, Margin = new Padding(0, 2, 0, 2) };
            if (cor.HasValue) l.ForeColor = cor.Value;
            return l;
        }

        public static Button Botao(string texto, EventHandler acao)
        {
            Button b = new Button { Text = texto, AutoSize = true, Font = Normal, Margin = new Padding(2), Padding = new Padding(4, 2, 4, 2) };
            b.Click += acao;
            return b;
        }

        public static ListView Lista(string[] colunas)
        {
            ListView lv = new ListView { View = View.Details, FullRowSelect = true, GridLines = true, MultiSelect = false, HideSelection = false, Font = Normal, Dock = DockStyle.Fill };
            foreach (string c in colunas) lv.Columns.Add(c);
            return lv;
        }

        public static void Preencher(ListView lv, List<LinhaTabela> linhas)
        {
            lv.BeginUpdate();
            lv.Items.Clear();
            foreach (LinhaTabela l in linhas)
            {
                ListViewItem it = new ListViewItem(l.Celulas);
                it.Tag = l.Chave;
                if (l.Estado != Estados.Ok) it.ForeColor = DoEstado(l.Estado);
                lv.Items.Add(it);
            }
            foreach (ColumnHeader c in lv.Columns) c.Width = -2;
            lv.EndUpdate();
        }

        public static void Executar(string comando)
        {
            HostMgd.ApplicationServices.Document d = Aplicacao.Doc();
            if (d != null) d.SendStringToExecute(comando + " ", true, false, false);
        }
    }

    // ── tela inicial ──
    public enum Escolha { Nada, Construir, Ler, Continuar }

    public class TelaInicio : Form
    {
        public Escolha Escolha = Escolha.Nada;
        public RegrasProjeto Regras;
        public string Projeto;
        private readonly ComboBox _regras;
        private readonly TextBox _projeto;

        public TelaInicio(Sessao s, PersistenciaDwg.ObraSalva salva)
        {
            Text = "⚡ Orçamento — o que deseja fazer neste desenho?";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false; MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            AutoSize = true; AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Font = Visual.Normal;
            Padding = new Padding(12);

            TableLayoutPanel raiz = new TableLayoutPanel { ColumnCount = 2, AutoSize = true, Dock = DockStyle.Fill };
            raiz.Controls.Add(Opcao("🧱 Construir do zero",
                "Você seleciona no desenho:\n• os postes e suas anotações\n• os vãos e suas anotações\n\nPara projetos sem padrão de layers/cores.",
                Escolha.Construir), 0, 0);
            raiz.Controls.Add(Opcao("📖 Ler projeto",
                "Lê o desenho inteiro com as regras do projeto (layer + cor).\nMonta postes, vãos e vínculos sozinho.\n\nSobra só a conferência e os ajustes.",
                Escolha.Ler), 1, 0);

            FlowLayoutPanel baixo = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.LeftToRight, WrapContents = false, Margin = new Padding(0, 10, 0, 0) };
            FlowLayoutPanel colRegras = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown };
            colRegras.Controls.Add(Visual.Rotulo("Regras do projeto", Visual.Secao));
            _regras = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 240 };
            foreach (RegrasProjeto r in Aplicacao.Regras) _regras.Items.Add(r);
            _regras.SelectedItem = s.Regras ?? Aplicacao.Regras.FirstOrDefault();
            colRegras.Controls.Add(_regras);
            FlowLayoutPanel colProj = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown };
            colProj.Controls.Add(Visual.Rotulo("Projeto (código na aplicação)", Visual.Secao));
            _projeto = new TextBox { Width = 90, Text = s.Obra != null ? s.Obra.Projeto : (s.Regras != null ? s.Regras.Codigo : "") };
            colProj.Controls.Add(_projeto);
            _regras.SelectedIndexChanged += (o, e) => { RegrasProjeto r = _regras.SelectedItem as RegrasProjeto; if (r != null && r.Codigo.Length > 0) _projeto.Text = r.Codigo; };
            FlowLayoutPanel colSalva = new FlowLayoutPanel { AutoSize = true, FlowDirection = FlowDirection.TopDown, Margin = new Padding(16, 0, 0, 0) };
            colSalva.Controls.Add(Visual.Rotulo("Obra salva neste desenho", Visual.Secao));
            colSalva.Controls.Add(Visual.Rotulo(salva == null ? "(nenhuma)" : ((salva.Quando ?? "") + " · " + salva.Onde), Visual.Mono));
            Button continuar = Visual.Botao("Continuar obra salva", (o, e) => Escolher(Escolha.Continuar));
            continuar.Enabled = salva != null;
            colSalva.Controls.Add(continuar);
            baixo.Controls.Add(colRegras);
            baixo.Controls.Add(colProj);
            baixo.Controls.Add(colSalva);
            raiz.Controls.Add(baixo, 0, 1);
            raiz.SetColumnSpan(baixo, 2);
            Controls.Add(raiz);
        }

        private Control Opcao(string titulo, string descricao, Escolha escolha)
        {
            Button b = new Button
            {
                Text = titulo + "\n\n" + descricao,
                Width = 300, Height = 170, Font = Visual.Normal, TextAlign = ContentAlignment.MiddleCenter, Margin = new Padding(6),
            };
            b.Click += (o, e) => Escolher(escolha);
            return b;
        }

        private void Escolher(Escolha e)
        {
            Escolha = e;
            Regras = _regras.SelectedItem as RegrasProjeto;
            Projeto = (_projeto.Text ?? "").Trim();
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
            Label r = Visual.Rotulo(rodape, Visual.Normal, Color.DimGray);
            r.Dock = DockStyle.Bottom; r.AutoSize = false; r.Height = 36; r.Padding = new Padding(6);
            Controls.Add(r);
            Atualizar();
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
            Button aplicar = Visual.Botao("Aplicar", (o, ev) => Fechar(AcaoPopup.Aplicar));
            aplicar.BackColor = Visual.Ok; aplicar.ForeColor = Color.White;
            botoes.Controls.Add(aplicar);
            botoes.Controls.Add(Visual.Botao("Voltar ao lido", (o, ev) => Fechar(AcaoPopup.VoltarAoLido)));
            if (podeExcluir) botoes.Controls.Add(Visual.Botao("Excluir da obra", (o, ev) => Fechar(AcaoPopup.Excluir)));
            botoes.Controls.Add(Visual.Botao("Fechar", (o, ev) => Fechar(AcaoPopup.Nada)));
            aj.Controls.Add(botoes);
            t.Controls.Add(aj, 0, 1);
            t.SetColumnSpan(aj, 2);
            Controls.Add(t);
            AcceptButton = aplicar;
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
            f.Controls.Add(Visual.Rotulo(titulo, Visual.Secao, Color.DimGray));
            return f;
        }

        private static void Campo(FlowLayoutPanel f, string rotulo, string valor, Color? cor = null)
        {
            f.Controls.Add(Visual.Rotulo(rotulo, Visual.Normal, Color.DimGray));
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
}
