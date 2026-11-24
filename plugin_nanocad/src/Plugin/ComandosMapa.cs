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

        // ── escolha do objeto base (usada ao coletar e em «Definir poste/cabo») ──

        /// <summary>Bloco ou círculo do poste. 1 = escolheu (bloco), 0 = Enter (sem símbolo), -1 = cancelou.</summary>
        private static int EscolherPoste(Sessao s, string mensagem, out PosteDesenho bloco)
        {
            bloco = null;
            Editor ed = s.Doc.Editor;
            while (true)
            {
                PromptEntityOptions o = new PromptEntityOptions(mensagem);
                o.AllowNone = true;
                PromptEntityResult r = ed.GetEntity(o);
                if (r.Status == PromptStatus.None) return 0;
                if (r.Status != PromptStatus.OK) return -1;
                DesenhoBruto b = ExtratorDesenho.LerSelecao(s.Doc, new[] { r.ObjectId });
                EntidadeBruta e = b.Entidades.FirstOrDefault();
                if (e == null || (e.Tipo != "INSERT" && e.Tipo != "CIRCLE")) { Aplicacao.Msg("Isso não é bloco nem círculo. Tente de novo (Enter = sem símbolo)."); continue; }
                bloco = new PosteDesenho { Handle = e.Handle, X = e.X, Y = e.Y, Bloco = e.Tipo == "CIRCLE" ? "(círculo)" : (e.Bloco ?? ""), Layer = e.Layer };
                return 1;
            }
        }

        /// <summary>Linha/polilinha do cabo, inteira ou um trecho entre dois pontos. false = cancelou.</summary>
        private static bool EscolherLinhaDoCabo(Sessao s, string mensagem, out LinhaDesenho linha, out List<Ponto> pontos, out bool trecho)
        {
            linha = null; pontos = null; trecho = false;
            Editor ed = s.Doc.Editor;
            while (true)
            {
                PromptEntityOptions o = new PromptEntityOptions(mensagem);
                o.AllowNone = true;
                PromptEntityResult r = ed.GetEntity(o);
                if (r.Status != PromptStatus.OK) return false;
                linha = Montar(s, new[] { r.ObjectId }).Linhas.FirstOrDefault(l => !l.Leader);
                if (linha == null || linha.Pontos.Count < 2) { Aplicacao.Msg("Isso não é linha nem polilinha. Tente de novo."); continue; }
                break;
            }
            pontos = linha.Pontos;
            PromptKeywordOptions k = new PromptKeywordOptions("\nUsar [Inteira/Trecho] <Inteira>: ", "Inteira Trecho") { AllowNone = true };
            PromptResult kr = ed.GetKeywords(k);
            if (kr.Status == PromptStatus.Cancel) return false;
            if (kr.Status == PromptStatus.OK && kr.StringResult == "Trecho")
            {
                PromptPointResult a = ed.GetPoint(new PromptPointOptions("\nPrimeiro ponto do trecho (sobre a linha): "));
                if (a.Status != PromptStatus.OK) return false;
                PromptPointResult b = ed.GetPoint(new PromptPointOptions("\nSegundo ponto do trecho: "));
                if (b.Status != PromptStatus.OK) return false;
                pontos = MapaManual.Trecho(linha.Pontos, new Ponto(a.Value.X, a.Value.Y), new Ponto(b.Value.X, b.Value.Y));
                trecho = true;
            }
            return true;
        }

        [CommandMethod("ORC_PONTO_MAPA")]
        public void CmdPontoMapa()
        {
            Rodar("ORC_PONTO_MAPA", s =>
            {
                if (s.EnviandoAnotacoes) return;
                PrepararConstrucao(s);
                if (!EscolherModoManual(s, false)) return;
                string ultimaOp = "I";
                while (true)
                {
                    PosteDesenho bloco;
                    if (EscolherPoste(s, "\nSelecione o POSTE — bloco ou círculo (Enter = sem símbolo, Esc = terminar): ", out bloco) < 0) break;
                    List<TextoDesenho> textos = SelecionarTextos(s, "\nSelecione as anotações do ponto (TEXT/MTEXT, várias) e Enter: ");
                    if (bloco == null && textos.Count == 0) break;
                    List<CampoMapa> campos = MapaManual.CamposDosTextos(textos, s.Obra.SemRegras == true ? null : s.Regras, ultimaOp, false);
                    string numero = MapaManual.ProximoNumero(s.Obra.PontosMapa.Select(p => p.Numero), "P");
                    string titulo = "Ponto " + numero + (bloco != null ? " · poste (" + (string.IsNullOrEmpty(bloco.Bloco) ? "bloco" : bloco.Bloco) + ")" : " · sem símbolo (posição da anotação)");
                    using (var tela = new TelaCamposMapa(titulo, false, campos, numero, "", 0, null, InfoRegras(s)))
                    {
                        if (Application.ShowModalDialog(tela) != DialogResult.OK) continue;
                        var ponto = new PontoMapa { Numero = numero, Nome = tela.Nome, Comentario = tela.Comentario, Campos = tela.Campos,
                            Textos = textos.Select(t => t.Handle).Distinct().ToList(), Caixas = MapaManual.CaixasDosTextos(textos) };
                        if (bloco != null) { ponto.Handle = bloco.Handle; ponto.Bloco = bloco.Bloco; ponto.X = bloco.X; ponto.Y = bloco.Y; }
                        else { TextoDesenho t0 = MapaManual.OrdemDeLeitura(textos).First(); ponto.X = t0.X; ponto.Y = t0.Y; }
                        s.Obra.PontosMapa.Add(ponto);
                        if (ponto.Campos.Count > 0) ultimaOp = ponto.Campos[ponto.Campos.Count - 1].Operacao;
                        Aplicacao.Msg("@" + ponto.NomeFinal + (ponto.Comentario.Length > 0 ? " - " + ponto.Comentario : "") + ": " + MapaManual.ResumoCampos(ponto.Campos));
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
                string ultimaOp = "I";
                while (true)
                {
                    LinhaDesenho linha; List<Ponto> pontos; bool trecho;
                    if (!EscolherLinhaDoCabo(s, "\nSelecione a LINHA do cabo — linha ou polilinha (Esc = terminar): ", out linha, out pontos, out trecho)) break;
                    double medido = MapaManual.Comprimento(pontos);
                    List<TextoDesenho> textos = SelecionarTextos(s, "\nSelecione as anotações do cabo (TEXT/MTEXT, várias) e Enter: ");
                    List<CampoMapa> campos = MapaManual.CamposDosTextos(textos, s.Obra.SemRegras == true ? null : s.Regras, ultimaOp, true);
                    double? anotado = MapaManual.OrdemDeLeitura(textos).Select(t => MapaManual.MetrosDoTexto(t.Texto)).FirstOrDefault(m => m.HasValue);
                    string numero = MapaManual.ProximoNumero(s.Obra.CabosMapa.Select(c => c.Numero), "C");
                    string tipoPrevio = MapaManual.TipoDoCabo(s.Regras.Leitura, linha.Layer, linha.Cor, campos);
                    string titulo = "Cabo " + numero + " · " + (trecho ? "trecho de " : "linha de ") + MapaManual.FormatarMetros(medido) + " m" + (tipoPrevio.Length > 0 ? " · " + tipoPrevio : "");
                    using (var tela = new TelaCamposMapa(titulo, true, campos, "", "", medido, anotado, InfoRegras(s)))
                    {
                        if (Application.ShowModalDialog(tela) != DialogResult.OK) continue;
                        var cabo = new CaboMapa { Numero = numero, LinhaHandle = linha.Handle ?? "", Trecho = trecho, Pontos = pontos, Medido = medido,
                            Metros = tela.Metros, Campos = tela.Campos, Textos = textos.Select(t => t.Handle).Distinct().ToList(), Caixas = MapaManual.CaixasDosTextos(textos) };
                        cabo.Tipo = MapaManual.TipoDoCabo(s.Regras.Leitura, linha.Layer, linha.Cor, cabo.Campos);       // 0.16.3: a cor da linha manda
                        s.Obra.CabosMapa.Add(cabo);
                        if (cabo.Campos.Count > 0) ultimaOp = cabo.Campos[cabo.Campos.Count - 1].Operacao;
                        Aplicacao.Msg(numero + " (" + MapaManual.FormatarMetros(cabo.Metros ?? medido) + " m" + (cabo.Tipo.Length > 0 ? ", " + cabo.Tipo : "") + "): " + MapaManual.ResumoCampos(cabo.Campos));
                        Mudou(s, true);
                    }
                }
            });
        }

        // ── item do mapa escolhido no painel («M:P1» / «M:C1») ──

        private static string LerItemDoMapa(Sessao s, string prompt)
        {
            PromptResult r = s.Doc.Editor.GetString(new PromptStringOptions("\n" + prompt + " (P1, C2…): ") { AllowSpaces = false });
            if (r.Status != PromptStatus.OK) return null;
            string n = (r.StringResult ?? "").Trim().ToUpperInvariant();
            return n.StartsWith(MapaManual.PrefixoPainel) ? n.Substring(MapaManual.PrefixoPainel.Length) : n;
        }

        private static string AvisoEnviado(bool enviado)
        {
            return enviado ? "Já enviado ao OrçaCore: o ajuste fica aqui e não é reenviado (altere também no OrçaCore)." : null;
        }

        /// <summary>Clique no painel: enquadra o item e o destaca junto com o objeto base (poste ou linha do cabo).</summary>
        [CommandMethod("ORC_MAPA_ZOOM")]
        public void CmdMapaZoom()
        {
            Rodar("ORC_MAPA_ZOOM", s =>
            {
                if (!PrecisaObra(s)) return;
                string n = LerItemDoMapa(s, "Item do mapa");
                if (n == null) return;
                List<Ponto> pts = MapaManual.PontosParaZoom(s.Obra, n);
                PontoMapa p = s.Obra.PontosMapa.FirstOrDefault(x => x.Numero == n);
                CaboMapa c = p == null ? s.Obra.CabosMapa.FirstOrDefault(x => x.Numero == n) : null;
                foreach (double[] cx in p != null ? p.Caixas : c != null ? c.Caixas : new List<double[]>())
                { pts.Add(new Ponto(cx[0], cx[1])); pts.Add(new Ponto(cx[2], cx[3])); }
                if (pts.Count > 0) ZoomPontos(s, pts);
                s.DestacarMapa(n);
            });
        }

        [CommandMethod("ORC_MAPA_ITEM")]
        public void CmdMapaItem()
        {
            Rodar("ORC_MAPA_ITEM", s =>
            {
                if (!PrecisaObra(s) || s.EnviandoAnotacoes) return;
                string n = LerItemDoMapa(s, "Item do mapa");
                if (n == null) return;
                List<Ponto> pts = MapaManual.PontosParaZoom(s.Obra, n);
                if (pts.Count > 0) ZoomPontos(s, pts);
                s.DestacarMapa(n);
                PontoMapa ponto = s.Obra.PontosMapa.FirstOrDefault(x => x.Numero == n);
                CaboMapa cabo = ponto == null ? s.Obra.CabosMapa.FirstOrDefault(x => x.Numero == n) : null;
                if (ponto == null && cabo == null) { Aplicacao.Msg("Item " + n + " não existe na montagem."); return; }
                string aviso = AvisoEnviado(ponto != null ? ponto.Enviado : cabo.Enviado) ?? InfoRegras(s);
                if (ponto != null)
                {
                    using (var tela = new TelaCamposMapa("Ponto " + ponto.Numero + (string.IsNullOrEmpty(ponto.Handle) ? " · sem símbolo" : " · poste"), false,
                        ponto.Campos, ponto.NomeFinal, ponto.Comentario, 0, null, aviso, true))
                    {
                        DialogResult d = Application.ShowModalDialog(tela);
                        if (d == DialogResult.Abort) { s.Obra.PontosMapa.Remove(ponto); Aplicacao.Msg("Ponto " + ponto.Numero + " excluído da montagem."); s.DestacarMapa(null); Mudou(s, true); }
                        else if (d == DialogResult.OK) { ponto.Nome = tela.Nome; ponto.Comentario = tela.Comentario; ponto.Campos = tela.Campos; Mudou(s, true); }
                    }
                    return;
                }
                double? anotado = cabo.Metros.HasValue && System.Math.Abs(cabo.Metros.Value - cabo.Medido) > 0.05 ? cabo.Metros : null;
                using (var tela = new TelaCamposMapa("Cabo " + cabo.Numero + " · " + (cabo.Trecho ? "trecho de " : "linha de ") + MapaManual.FormatarMetros(cabo.Medido) + " m"
                    + (string.IsNullOrEmpty(cabo.Tipo) ? "" : " · " + cabo.Tipo), true, cabo.Campos, "", "", cabo.Medido, anotado, aviso, true))
                {
                    DialogResult d = Application.ShowModalDialog(tela);
                    if (d == DialogResult.Abort) { s.Obra.CabosMapa.Remove(cabo); Aplicacao.Msg("Cabo " + cabo.Numero + " excluído da montagem."); s.DestacarMapa(null); Mudou(s, true); }
                    else if (d == DialogResult.OK) { cabo.Campos = tela.Campos; cabo.Metros = tela.Metros; Mudou(s, true); }
                }
            });
        }

        /// <summary>Excluir o item (o painel já confirmou).</summary>
        [CommandMethod("ORC_MAPA_EXCLUIR")]
        public void CmdMapaExcluir()
        {
            Rodar("ORC_MAPA_EXCLUIR", s =>
            {
                if (!PrecisaObra(s) || s.EnviandoAnotacoes) return;
                string n = LerItemDoMapa(s, "Item do mapa a excluir");
                if (n == null) return;
                int tirados = s.Obra.PontosMapa.RemoveAll(x => x.Numero == n) + s.Obra.CabosMapa.RemoveAll(x => x.Numero == n);
                if (tirados == 0) { Aplicacao.Msg("Item " + n + " não existe na montagem."); return; }
                s.DestacarMapa(null);
                Aplicacao.Msg((n.StartsWith("C") ? "Cabo " : "Ponto ") + n + " excluído da montagem.");
                Mudou(s, true);
            });
        }

        /// <summary>Definir o objeto base: outro bloco/círculo para o ponto (Enter = sem símbolo) ou outra linha/trecho para o cabo. Os campos ficam.</summary>
        [CommandMethod("ORC_MAPA_OBJETO")]
        public void CmdMapaObjeto()
        {
            Rodar("ORC_MAPA_OBJETO", s =>
            {
                if (!PrecisaObra(s) || s.EnviandoAnotacoes) return;
                string n = LerItemDoMapa(s, "Item do mapa");
                if (n == null) return;
                PontoMapa ponto = s.Obra.PontosMapa.FirstOrDefault(x => x.Numero == n);
                CaboMapa cabo = ponto == null ? s.Obra.CabosMapa.FirstOrDefault(x => x.Numero == n) : null;
                if (ponto == null && cabo == null) { Aplicacao.Msg("Item " + n + " não existe na montagem."); return; }
                s.DestacarMapa(n);
                string aviso = AvisoEnviado(ponto != null ? ponto.Enviado : cabo.Enviado);
                if (aviso != null) Aplicacao.Msg(aviso);
                if (ponto != null)
                {
                    PosteDesenho bloco;
                    int r = EscolherPoste(s, "\nNovo POSTE de " + ponto.NomeFinal + " — bloco ou círculo (Enter = sem símbolo, Esc = manter): ", out bloco);
                    if (r < 0) return;
                    if (bloco != null) { ponto.Handle = bloco.Handle; ponto.Bloco = bloco.Bloco; ponto.X = bloco.X; ponto.Y = bloco.Y; }
                    else
                    {
                        ponto.Handle = ""; ponto.Bloco = "";
                        if (ponto.Caixas.Count > 0) { ponto.X = ponto.Caixas[0][0]; ponto.Y = ponto.Caixas[0][3]; }
                    }
                    Aplicacao.Msg("Ponto " + ponto.NomeFinal + ": " + (bloco != null ? "poste (" + (string.IsNullOrEmpty(bloco.Bloco) ? "bloco" : bloco.Bloco) + ")" : "sem símbolo (posição da anotação)") + ".");
                }
                else
                {
                    LinhaDesenho linha; List<Ponto> pontos; bool trecho;
                    if (!EscolherLinhaDoCabo(s, "\nNova LINHA do cabo " + cabo.Numero + " — linha ou polilinha (Esc = manter): ", out linha, out pontos, out trecho)) return;
                    bool eraMedido = !cabo.Metros.HasValue || System.Math.Abs(cabo.Metros.Value - cabo.Medido) <= 0.05;
                    cabo.LinhaHandle = linha.Handle ?? ""; cabo.Trecho = trecho; cabo.Pontos = pontos; cabo.Medido = MapaManual.Comprimento(pontos);
                    if (eraMedido) cabo.Metros = cabo.Medido;                                   // o comprimento anotado no texto fica; o medido acompanha a linha nova
                    cabo.Tipo = MapaManual.TipoDoCabo(s.Regras.Leitura, linha.Layer, linha.Cor, cabo.Campos);
                    Aplicacao.Msg("Cabo " + cabo.Numero + ": " + (trecho ? "trecho de " : "linha de ") + MapaManual.FormatarMetros(cabo.Medido) + " m"
                        + (cabo.Tipo.Length > 0 ? " · " + cabo.Tipo : "") + " · vale " + MapaManual.FormatarMetros(cabo.Metros ?? cabo.Medido) + " m.");
                }
                Mudou(s, true);
                s.DestacarMapa(n);
            });
        }
    }
}
