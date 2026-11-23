// Plugin/ComandosMapa.cs — TASK-038: montagem manual no formato do mapa do OrçaCore.
//   ORC_PONTO_MAPA: poste (bloco/círculo ou Enter = sem símbolo) → anotações (vários TEXT/MTEXT) → Enter → janela dos campos → ponto guardado.
//   ORC_CABO_MAPA: linha/polilinha (inteira ou [T]recho entre dois pontos) → anotações → janela (campos + comprimento) → cabo guardado.
//   ORC_MODO_MANUAL: escolhe entre as regras de um projeto e «sem regras».
// O envio à obra aberta no OrçaCore é o mesmo do painel («Adicionar anotações ao OrçaCore»). O vínculo é feito no OrçaCore. C# 5 (nanoCAD 5).
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using HostMgd.EditorInput;
using OrcamentoNanoCAD.Nucleo;
using OrcamentoNanoCAD.UI;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;
using Application = HostMgd.ApplicationServices.Application;

namespace OrcamentoNanoCAD
{
    internal sealed class TelaModoManual : Form
    {
        private readonly RadioButton _com = new RadioButton { Text = "Usar as regras de um projeto (ativo e operação pelas regras do Leitor; texto sem regra entra como F)", AutoSize = true, Checked = true };
        private readonly RadioButton _sem = new RadioButton { Text = "Sem regras (o texto vai como está; você marca a operação)", AutoSize = true };
        private readonly ComboBox _regras = new ComboBox { Width = 420, DropDownStyle = ComboBoxStyle.DropDownList };
        public bool SemRegras { get { return _sem.Checked; } }
        public RegrasProjeto Regras { get { return _regras.SelectedItem as RegrasProjeto; } }

        public TelaModoManual(RegrasProjeto atuais, bool? semRegras)
        {
            Text = "Montagem manual"; Width = 640; Height = 260; StartPosition = FormStartPosition.CenterParent; MinimizeBox = false; MaximizeBox = false;
            var corpo = new FlowLayoutPanel { Dock = DockStyle.Fill, FlowDirection = System.Windows.Forms.FlowDirection.TopDown, WrapContents = false, Padding = new Padding(12) };
            foreach (RegrasProjeto r in Aplicacao.RegrasParaEscolher().Where(r => !string.IsNullOrEmpty(r.Codigo))) _regras.Items.Add(r);
            if (atuais != null && !_regras.Items.Cast<RegrasProjeto>().Any(r => r.Codigo == atuais.Codigo)) _regras.Items.Insert(0, atuais);
            _regras.SelectedItem = _regras.Items.Cast<RegrasProjeto>().FirstOrDefault(r => atuais != null && r.Codigo == atuais.Codigo) ?? (_regras.Items.Count > 0 ? _regras.Items[0] : null);
            if (semRegras == true) _sem.Checked = true;
            _com.CheckedChanged += (o, e) => _regras.Enabled = _com.Checked;
            _regras.Enabled = _com.Checked;
            corpo.Controls.Add(_com); corpo.Controls.Add(_regras); corpo.Controls.Add(_sem);
            var rodape = new FlowLayoutPanel { Dock = DockStyle.Bottom, AutoSize = true, FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft, Padding = new Padding(8) };
            Button ok = Visual.Botao("Começar", (o, e) => { if (_com.Checked && Regras == null) return; DialogResult = DialogResult.OK; }, true);
            Button cancelar = Visual.Botao("Cancelar", (o, e) => DialogResult = DialogResult.Cancel);
            AcceptButton = ok; CancelButton = cancelar;
            rodape.Controls.Add(ok); rodape.Controls.Add(cancelar);
            Controls.Add(corpo); Controls.Add(rodape);
            Visual.Aplicar(this);
        }
    }

    public partial class Comandos
    {
        /// <summary>Pergunta (uma vez por obra, ou quando pedido) se a montagem usa as regras de um projeto ou nenhuma. false = cancelou.</summary>
        private static bool EscolherModoManual(Sessao s, bool forcar)
        {
            if (s.Obra.SemRegras.HasValue && !forcar) return true;
            using (var tela = new TelaModoManual(s.Regras, s.Obra.SemRegras))
            {
                if (Application.ShowModalDialog(tela) != DialogResult.OK) return false;
                if (!tela.SemRegras && tela.Regras != null && tela.Regras.Codigo != s.Regras.Codigo)
                {
                    RegrasProjeto r = RegrasFrescas(tela.Regras, tela.Regras.Codigo);
                    if (r == null) return false;
                    s.Regras = r; s.Obra.Projeto = r.Codigo; s.Obra.NomeRegras = r.Nome;
                }
                s.Obra.SemRegras = tela.SemRegras;
            }
            Aplicacao.Msg(s.Obra.SemRegras == true ? "Montagem manual sem regras: o texto vai como está." : "Montagem manual com as regras de " + s.Regras.Nome + ".");
            Mudou(s, true);
            return true;
        }

        private static string InfoRegras(Sessao s)
        {
            return s.Obra.SemRegras == true ? "Sem regras: o texto vai como está." : "Regras: " + s.Regras.Nome + " · texto sem regra entra como F (fictício).";
        }

        [CommandMethod("ORC_MODO_MANUAL")]
        public void CmdModoManual() { Rodar("ORC_MODO_MANUAL", s => { PrepararConstrucao(s); EscolherModoManual(s, true); }); }

        [CommandMethod("ORC_PONTO_MAPA")]
        public void CmdPontoMapa()
        {
            Rodar("ORC_PONTO_MAPA", s =>
            {
                if (s.EnviandoAnotacoes) return;
                PrepararConstrucao(s);
                if (!EscolherModoManual(s, false)) return;
                Editor ed = s.Doc.Editor;
                string ultimaOp = "I";
                while (true)
                {
                    PromptEntityOptions o = new PromptEntityOptions("\nSelecione o POSTE — bloco ou círculo (Enter = sem símbolo, Esc = terminar): ");
                    o.AllowNone = true;
                    PromptEntityResult r = ed.GetEntity(o);
                    if (r.Status == PromptStatus.Cancel || r.Status == PromptStatus.Error) break;
                    PosteDesenho bloco = null;
                    if (r.Status == PromptStatus.OK)
                    {
                        DesenhoBruto b = ExtratorDesenho.LerSelecao(s.Doc, new[] { r.ObjectId });
                        EntidadeBruta e = b.Entidades.FirstOrDefault();
                        if (e == null || (e.Tipo != "INSERT" && e.Tipo != "CIRCLE")) { Aplicacao.Msg("Isso não é bloco nem círculo. Tente de novo (Enter = ponto sem símbolo)."); continue; }
                        bloco = new PosteDesenho { Handle = e.Handle, X = e.X, Y = e.Y, Bloco = e.Tipo == "CIRCLE" ? "(círculo)" : (e.Bloco ?? ""), Layer = e.Layer };
                    }
                    List<TextoDesenho> textos = SelecionarTextos(s, "\nSelecione as anotações do ponto (TEXT/MTEXT, várias) e Enter: ");
                    if (bloco == null && textos.Count == 0) break;
                    List<CampoMapa> campos = MapaManual.CamposDosTextos(textos, s.Obra.SemRegras == true ? null : s.Regras, ultimaOp, false);
                    string numero = MapaManual.ProximoNumero(s.Obra.PontosMapa.Select(p => p.Numero), "P");
                    string titulo = "Ponto " + numero + (bloco != null ? " · poste (" + (string.IsNullOrEmpty(bloco.Bloco) ? "bloco" : bloco.Bloco) + ")" : " · sem símbolo (posição da anotação)");
                    using (var tela = new TelaCamposMapa(titulo, false, campos, numero, "", 0, null, InfoRegras(s)))
                    {
                        if (Application.ShowModalDialog(tela) != DialogResult.OK) continue;
                        var ponto = new PontoMapa { Numero = numero, Nome = tela.Nome, Comentario = tela.Comentario, Campos = tela.Campos,
                            Textos = textos.Select(t => t.Handle).Distinct().ToList() };
                        if (bloco != null) { ponto.Handle = bloco.Handle; ponto.Bloco = bloco.Bloco; ponto.X = bloco.X; ponto.Y = bloco.Y; }
                        else { TextoDesenho t0 = MapaManual.OrdemDeLeitura(textos).First(); ponto.X = t0.X; ponto.Y = t0.Y; }
                        s.Obra.PontosMapa.Add(ponto);
                        if (ponto.Campos.Count > 0) ultimaOp = ponto.Campos[ponto.Campos.Count - 1].Operacao;
                        Aplicacao.Msg("@" + ponto.NomeFinal + (ponto.Comentario.Length > 0 ? " - " + ponto.Comentario : "") + ": "
                                      + string.Join(" · ", ponto.Campos.Select(c => c.Operacao + " " + c.Texto).ToArray()));
                        Mudou(s, true);
                    }
                }
            });
        }

        [CommandMethod("ORC_CABO_MAPA")]
        public void CmdCaboMapa()
        {
            Rodar("ORC_CABO_MAPA", s =>
            {
                if (s.EnviandoAnotacoes) return;
                PrepararConstrucao(s);
                if (!EscolherModoManual(s, false)) return;
                Editor ed = s.Doc.Editor;
                string ultimaOp = "I";
                while (true)
                {
                    PromptEntityOptions o = new PromptEntityOptions("\nSelecione a LINHA do cabo — linha ou polilinha (Esc = terminar): ");
                    o.AllowNone = true;
                    PromptEntityResult r = ed.GetEntity(o);
                    if (r.Status != PromptStatus.OK) break;
                    LinhaDesenho linha = Montar(s, new[] { r.ObjectId }).Linhas.FirstOrDefault(l => !l.Leader);
                    if (linha == null || linha.Pontos.Count < 2) { Aplicacao.Msg("Isso não é linha nem polilinha. Tente de novo."); continue; }
                    List<Ponto> pontos = linha.Pontos;
                    bool trecho = false;
                    PromptKeywordOptions k = new PromptKeywordOptions("\nUsar [Inteira/Trecho] <Inteira>: ", "Inteira Trecho") { AllowNone = true };
                    PromptResult kr = ed.GetKeywords(k);
                    if (kr.Status == PromptStatus.Cancel) break;
                    if (kr.Status == PromptStatus.OK && kr.StringResult == "Trecho")
                    {
                        PromptPointResult a = ed.GetPoint(new PromptPointOptions("\nPrimeiro ponto do trecho (sobre a linha): "));
                        if (a.Status != PromptStatus.OK) continue;
                        PromptPointResult b = ed.GetPoint(new PromptPointOptions("\nSegundo ponto do trecho: "));
                        if (b.Status != PromptStatus.OK) continue;
                        pontos = MapaManual.Trecho(linha.Pontos, new Ponto(a.Value.X, a.Value.Y), new Ponto(b.Value.X, b.Value.Y));
                        trecho = true;
                    }
                    double medido = MapaManual.Comprimento(pontos);
                    List<TextoDesenho> textos = SelecionarTextos(s, "\nSelecione as anotações do cabo (TEXT/MTEXT, várias) e Enter: ");
                    List<CampoMapa> campos = MapaManual.CamposDosTextos(textos, s.Obra.SemRegras == true ? null : s.Regras, ultimaOp, true);
                    double? anotado = MapaManual.OrdemDeLeitura(textos).Select(t => MapaManual.MetrosDoTexto(t.Texto)).FirstOrDefault(m => m.HasValue);
                    string numero = MapaManual.ProximoNumero(s.Obra.CabosMapa.Select(c => c.Numero), "C");
                    string tipo = s.Regras.Leitura.ClassificarRede(linha.Layer, linha.Cor) ?? "";
                    string titulo = "Cabo " + numero + " · " + (trecho ? "trecho de " : "linha de ") + MapaManual.FormatarMetros(medido) + " m" + (tipo.Length > 0 ? " · " + tipo : "");
                    using (var tela = new TelaCamposMapa(titulo, true, campos, "", "", medido, anotado, InfoRegras(s)))
                    {
                        if (Application.ShowModalDialog(tela) != DialogResult.OK) continue;
                        var cabo = new CaboMapa { Numero = numero, LinhaHandle = linha.Handle ?? "", Trecho = trecho, Pontos = pontos, Tipo = tipo, Medido = medido,
                            Metros = tela.Metros, Campos = tela.Campos, Textos = textos.Select(t => t.Handle).Distinct().ToList() };
                        s.Obra.CabosMapa.Add(cabo);
                        if (cabo.Campos.Count > 0) ultimaOp = cabo.Campos[cabo.Campos.Count - 1].Operacao;
                        Aplicacao.Msg(numero + " (" + MapaManual.FormatarMetros(cabo.Metros ?? medido) + " m): "
                                      + string.Join(" · ", cabo.Campos.Select(c => c.Operacao + " " + c.Texto).ToArray()));
                        Mudou(s, true);
                    }
                }
            });
        }
    }
}
