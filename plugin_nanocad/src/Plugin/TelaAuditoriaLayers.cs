using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using OrcamentoNanoCAD.Nucleo;
using Teigha.DatabaseServices;

namespace OrcamentoNanoCAD.UI
{
    internal sealed class TelaAuditoriaLayers : Form
    {
        private static TelaAuditoriaLayers _aberta;
        private readonly Sessao _sessao;
        private readonly TabControl _abas = new TabControl { Dock = DockStyle.Fill };
        private readonly ListView _lista = Visual.Lista(new[] { "Texto", "Layer", "Operação", "Handle" });
        private readonly ComboBox _destino = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 100 };
        private readonly CheckedListBox _layers = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true };
        private readonly Marcas _marcas = new Marcas();
        private readonly VisibilidadeDeNotas _visibilidade = new VisibilidadeDeNotas();
        private HashSet<string> _ocultas = new HashSet<string>();
        private List<EntidadeBruta> _textos = new List<EntidadeBruta>();
        private Dictionary<string, string> _operacoes;
        private bool _carregando;
        private ObjectId[] _enviadas;
        private string _acaoEnviada;

        internal static void AplicarSelecao(Sessao s)
        {
            if (_aberta == null || _aberta._sessao != s || _aberta._enviadas == null) return;
            ObjectId[] ids = _aberta._enviadas;
            string acao = _aberta._acaoEnviada;
            _aberta._enviadas = null;
            AuditoriaLayersCad.Aplicar(s, ids, acao, false);
        }

        private bool Ativa { get { return Aplicacao.Doc() == _sessao.Doc && Visible; } }
        private bool Analise { get { return _abas.SelectedIndex == 4; } }

        public static void Abrir(Sessao s)
        {
            if (_aberta != null) _aberta.Close();
            var nova = new TelaAuditoriaLayers(s);
            _aberta = nova;
            Visual.MostrarNaoModal(nova);
        }

        public static void Atualizar(Sessao s)
        {
            if (_aberta != null && !_aberta.IsDisposed && _aberta._sessao == s) _aberta.Carregar();
        }

        private TelaAuditoriaLayers(Sessao s)
        {
            _sessao = s;
            Text = "OrçaCore · Auditoria de layers · " + s.Regras.Nome;
            Width = 700; Height = 520; MinimumSize = new Size(490, 360);
            _lista.MultiSelect = true;
            _abas.TabPages.Add("Instalando"); _abas.TabPages.Add("Removendo"); _abas.TabPages.Add("Reaplicando");
            _abas.TabPages.Add("Fora"); _abas.TabPages.Add("Análise");
            _abas.TabPages[0].Controls.Add(_lista);
            _abas.TabPages[4].Controls.Add(_layers);
            var legenda = new Label { Text = "Marcas visíveis por layer · vermelho I · magenta R · azul M · LV linha viva", Dock = DockStyle.Top, Height = 28 };
            _abas.TabPages[4].Controls.Add(legenda);
            var barra = new FlowLayoutPanel { Dock = DockStyle.Top, AutoSize = true, Padding = new Padding(5) };
            _destino.Items.AddRange(s.Regras.Operacoes().Where(op => s.Regras.LayerDaOperacao(op) != null).Concat(new[] { "0" }).ToArray());
            _destino.SelectedIndex = 0;
            barra.Controls.Add(_destino);
            barra.Controls.Add(Visual.Botao("Mover para layer", (o,e) => Aplicar(_destino.Text)));
            barra.Controls.Add(Visual.Botao("Alternar LV", (o,e) => Aplicar("LV")));
            barra.Controls.Add(Visual.Botao("Excluir…", (o,e) => Aplicar("EXCLUIR")));
            barra.Controls.Add(Visual.Botao("Apontar no desenho", (o,e) => NaTela(_destino.Text)));
            barra.Controls.Add(Visual.Botao("LV no desenho", (o,e) => NaTela("LV")));
            barra.Controls.Add(Visual.Botao("Excluir no desenho…", (o,e) => NaTela("EXCLUIR")));
            barra.Controls.Add(Visual.Botao("Atualizar", (o,e) => { if (Ativa) Carregar(); }));
            var filtros = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true };
            filtros.Controls.Add(Visual.Botao("Ocultar todas", (o,e) => MarcarLayers(false)));
            filtros.Controls.Add(Visual.Botao("Mostrar todas", (o,e) => MarcarLayers(true)));
            _abas.TabPages[4].Controls.Add(filtros);
            Controls.Add(_abas); Controls.Add(barra);
            _abas.SelectedIndexChanged += (o,e) => MostrarAba();
            _lista.SelectedIndexChanged += (o,e) => Destacar();
            _lista.Click += (o,e) => { if (_abas.SelectedIndex == 3 && _lista.SelectedItems.Count == 1) Localizar(); };
            _lista.ItemActivate += (o,e) => { if (_abas.SelectedIndex != 3) Localizar(); };
            _lista.KeyDown += (o,e) =>
            {
                if (e.Control && e.KeyCode == Keys.A) { foreach (ListViewItem item in _lista.Items) item.Selected = true; e.SuppressKeyPress = true; }
                if (e.KeyCode == Keys.Delete) { Aplicar("EXCLUIR"); e.SuppressKeyPress = true; }
            };
            _layers.ItemCheck += (o,e) => { if (!_carregando && IsHandleCreated) BeginInvoke((Action)Destacar); };
            Shown += (o,e) => Destacar();
            FormClosed += (o,e) => { RemoverVisibilidade(); _marcas.Limpar(); if (_aberta == this) _aberta = null; };
            Visual.Aplicar(this);
            Carregar();
        }

        private string Operacao(EntidadeBruta t)
        {
            string op;
            return _operacoes.TryGetValue(t.Layer, out op) ? op : "";
        }

        private void Carregar()
        {
            _operacoes = _sessao.Regras.OperacoesPorLayer();
            _textos = AuditoriaLayersCad.Ler(_sessao);
            var ocultas = new HashSet<string>(_layers.Items.Cast<string>().Except(_layers.CheckedItems.Cast<string>()));
            _carregando = true;
            _layers.Items.Clear();
            foreach (string layer in _textos.Select(t => t.Layer).Distinct().OrderBy(x => x))
                _layers.Items.Add(layer, !ocultas.Contains(layer));
            _carregando = false;
            MostrarAba();
        }

        private void MostrarAba()
        {
            if (!Analise)
            {
                _abas.SelectedTab.Controls.Add(_lista);
                string operacao = new[] { "I", "R", "M", "" }[_abas.SelectedIndex];
                _carregando = true;
                Visual.Preencher(_lista, _textos.Where(t => Operacao(t).TrimStart('*') == operacao).Select(t =>
                    new LinhaTabela { Chave = t.Handle, Celulas = new[] { MotorLeitor.LimparTexto(t.Texto), t.Layer,
                        Operacao(t), t.Handle } }).ToList());
                _carregando = false;
            }
            Destacar();
        }

        private void MarcarLayers(bool visivel)
        {
            _carregando = true;
            for (int i = 0; i < _layers.Items.Count; i++) _layers.SetItemChecked(i, visivel);
            _carregando = false;
            Destacar();
        }

        private void Destacar()
        {
            if (_carregando || !Ativa) return;
            var ocultar = new HashSet<string>(Analise ? _layers.Items.Cast<string>().Except(_layers.CheckedItems.Cast<string>()) : new string[0]);
            if (!_ocultas.SetEquals(ocultar))
            {
                var ids = _textos.Where(t => ocultar.Contains(t.Layer)).Select(t => GuiasDeNota.Id(_sessao.Doc.Database, t.Handle)).Where(id => !id.IsNull).ToArray();
                if (Compat.Tentar("visibilidade de notas na viewport", () => _visibilidade.Aplicar(ids), false)) _sessao.Doc.Editor.Regen();
                _ocultas = ocultar;
            }
            _marcas.Limpar();
            var escolhidas = new HashSet<string>(_lista.SelectedItems.Cast<ListViewItem>().Select(i => (string)i.Tag));
            var visiveis = new HashSet<string>(_layers.CheckedItems.Cast<string>());
            foreach (EntidadeBruta texto in _textos)
            {
                if (Analise ? !visiveis.Contains(texto.Layer) : !escolhidas.Contains(texto.Handle)) continue;
                string op = Operacao(texto);
                _marcas.MarcarTexto(new Ponto(texto.X,texto.Y), texto.Altura, op.StartsWith("*") ? op + " LV" : (op.Length == 0 ? "Fora" : op),
                    Analise ? Cor(op) : (short)2);
            }
        }

        private void RemoverVisibilidade()
        {
            _visibilidade.Dispose();
            if (_ocultas.Count > 0 && Aplicacao.Doc() == _sessao.Doc) _sessao.Doc.Editor.Regen();
            _ocultas.Clear();
        }

        private static short Cor(string op)
        {
            switch (op)
            {
                case "I": return 1; case "*I": return 12;
                case "R": return 6; case "*R": return 212;
                case "M": return 5; case "*M": return 172;
                default: return 7;
            }
        }

        private void Localizar()
        {
            if (Ativa && _lista.SelectedItems.Count > 0)
                Visual.Executar("ORC_ZOOM_NOTA " + (string)_lista.SelectedItems[0].Tag);
        }

        private void NaTela(string operacao)
        {
            if (Ativa) Visual.Executar("ORC_LAYER_NOTA " + operacao + (Analise ? " SIM" : " NAO"));
        }

        private void Aplicar(string acao)
        {
            if (!Ativa || Analise || _lista.SelectedItems.Count == 0) return;
            var ids = _lista.SelectedItems.Cast<ListViewItem>().Select(item => GuiasDeNota.Id(_sessao.Doc.Database, (string)item.Tag)).ToArray();
            string pergunta = acao == "EXCLUIR" ? "Excluir do DWG " + ids.Length + " anotação(ões) e suas guias do OrçaCore?" :
                "Alterar a layer no DWG de " + ids.Length + " anotação(ões)?";
            if (MessageBox.Show(this, pergunta, Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            _enviadas = ids; _acaoEnviada = acao;
            Visual.Executar("ORC_LAYERS_LISTA");
        }
    }
}
