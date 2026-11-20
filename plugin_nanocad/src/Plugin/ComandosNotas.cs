using System;
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
    public partial class Comandos
    {
        [CommandMethod("ORC_NOTA")]
        public void CmdNota()
        {
            Rodar("ORC_NOTA", s =>
            {
                var alvo = s.Doc.Editor.GetEntity("\nSelecione o poste, cabo ou objeto da nota: ");
                if (alvo.Status != PromptStatus.OK) return;
                using (var tela = new TelaNota(s.Regras, "", "I", s.Obra == null ? 1 : s.Obra.AlturaTexto, 0))
                {
                    if (Application.ShowModalDialog(tela) != DialogResult.OK) return;
                    Point3d? ponto = PontoDaNota(s.Doc.Editor, "\nPosição da anotação: ");
                    if (!ponto.HasValue) return;
                    Database db = s.Doc.Database;
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        var entidade = (Entity)tr.GetObject(alvo.ObjectId, OpenMode.ForRead);
                        var nota = new MText { Location = ponto.Value };
                        nota.SetDatabaseDefaults(db);
                        GuiasDeNota.Modelo(db, tr, OpenMode.ForWrite).AppendEntity(nota);
                        tr.AddNewlyCreatedDBObject(nota, true);
                        AplicarNota(db, tr, nota, tela, s.Regras);
                        if (tela.Guia) GuiasDeNota.Criar(db, tr, nota, entidade);
                        tr.Commit();
                    }
                    NotaAlterada(s);
                }
            });
        }

        [CommandMethod("ORC_EDITAR_NOTA")]
        public void CmdEditarNota()
        {
            Rodar("ORC_EDITAR_NOTA", s =>
            {
                ObjectId id = PedirNota(s.Doc.Editor);
                if (id.IsNull) return;
                Database db = s.Doc.Database;
                string texto;
                double altura;
                int alinhamento;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var nota = (MText)tr.GetObject(id, OpenMode.ForRead);
                    texto = nota.Contents; altura = nota.TextHeight;
                    alinhamento = (int)nota.Attachment;
                }
                using (var tela = new TelaNota(s.Regras, texto, null, altura, alinhamento))
                {
                    if (Application.ShowModalDialog(tela) != DialogResult.OK) return;
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        var nota = (MText)tr.GetObject(id, OpenMode.ForWrite);
                        AplicarNota(db, tr, nota, tela, s.Regras);
                        GuiasDeNota.AtualizarNoModelo(db, tr);
                        tr.Commit();
                    }
                    NotaAlterada(s);
                }
            });
        }

        [CommandMethod("ORC_MOVER_NOTA")]
        public void CmdMoverNota()
        {
            Rodar("ORC_MOVER_NOTA", s =>
            {
                Editor ed = s.Doc.Editor;
                ObjectId id = PedirNota(ed);
                if (id.IsNull) return;
                var opcoes = new PromptKeywordOptions("\nMovimento [Nota/Grupo/Alinhar/Vertical] <Nota>: ", "Nota Grupo Alinhar Vertical") { AllowNone = true };
                opcoes.Keywords.Default = "Nota";
                PromptResult modo = ed.GetKeywords(opcoes);
                if (modo.Status != PromptStatus.OK && modo.Status != PromptStatus.None) return;
                bool vertical = modo.StringResult == "Vertical", alinhar = modo.StringResult == "Alinhar";
                var ids = new HashSet<ObjectId> { id };
                if (modo.StringResult == "Grupo" || alinhar || vertical)
                    ids.UnionWith(SelecionarNotas(ed, "\nSelecione os MTEXT que acompanham a nota: "));
                Point3d? destino = PontoDaNota(ed, "\nNova posição da nota (Enter mantém): ");
                if (!destino.HasValue) return;
                Database db = s.Doc.Database;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    var principal = (MText)tr.GetObject(id, OpenMode.ForRead);
                    Vector3d deslocamento = destino.Value - principal.Location;
                    if (vertical) deslocamento = new Vector3d(0, deslocamento.Y, 0);
                    foreach (ObjectId item in ids)
                    {
                        var nota = (MText)tr.GetObject(item, OpenMode.ForWrite);
                        Point3d posicao = nota.Location + deslocamento;
                        nota.Location = alinhar ? new Point3d(destino.Value.X, posicao.Y, posicao.Z) : posicao;
                    }
                    GuiasDeNota.AtualizarNoModelo(db, tr);
                    tr.Commit();
                }
                NotaAlterada(s);
            });
        }

        [CommandMethod("ORC_REAPONTAR_NOTA")]
        public void CmdReapontarNota()
        {
            Rodar("ORC_REAPONTAR_NOTA", s =>
            {
                ObjectId id = PedirNota(s.Doc.Editor);
                if (id.IsNull) return;
                PromptEntityResult escolha = s.Doc.Editor.GetEntity("\nSelecione o novo objeto de destino da guia: ");
                if (escolha.Status != PromptStatus.OK || escolha.ObjectId == id) return;
                var db = s.Doc.Database;
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var nota = (MText)tr.GetObject(id, OpenMode.ForRead);
                    var alvo = (Entity)tr.GetObject(escolha.ObjectId, OpenMode.ForRead);
                    GuiasDeNota.AtualizarNoModelo(db, tr, new HashSet<string> { nota.Handle.ToString() });
                    GuiasDeNota.Criar(db, tr, nota, alvo);
                    tr.Commit();
                }
                NotaAlterada(s);
            });
        }

        [CommandMethod("ORC_ATUALIZAR_GUIAS")]
        public void CmdAtualizarGuias() { Rodar("ORC_ATUALIZAR_GUIAS", s => GuiasDeNota.Atualizar(s.Doc)); }

        [CommandMethod("ORC_AUDITAR_LAYERS")]
        public void CmdAuditarLayers() { Rodar("ORC_AUDITAR_LAYERS", s => TelaAuditoriaLayers.Abrir(s)); }

        [CommandMethod("ORC_ZOOM_NOTA")]
        public void CmdZoomNota()
        {
            Rodar("ORC_ZOOM_NOTA", s =>
            {
                string handle = PedirPalavra(s.Doc.Editor, "\nHandle do MTEXT: ");
                ObjectId id = GuiasDeNota.Id(s.Doc.Database, handle);
                if (id.IsNull || id.IsErased) return;
                using (var tr = s.Doc.Database.TransactionManager.StartTransaction())
                {
                    var nota = tr.GetObject(id, OpenMode.ForRead) as MText;
                    if (nota == null) return;
                    var caixa = nota.GeometricExtents;
                    ZoomPontos(s, new List<Ponto> { new Ponto(caixa.MinPoint.X, caixa.MinPoint.Y), new Ponto(caixa.MaxPoint.X, caixa.MaxPoint.Y) });
                }
            });
        }

        [CommandMethod("ORC_LAYERS_LISTA")]
        public void CmdLayersLista()
        {
            Rodar("ORC_LAYERS_LISTA", s => TelaAuditoriaLayers.AplicarSelecao(s));
        }

        [CommandMethod("ORC_LAYER_NOTA")]
        public void CmdLayerNota()
        {
            Rodar("ORC_LAYER_NOTA", s =>
            {
                Editor ed = s.Doc.Editor;
                string acao = PedirPalavra(ed, "\nOperação I/R/M, *I/*R/*M, LV, 0 ou EXCLUIR: ");
                if (acao == null) return;
                string modo = PedirPalavra(ed, "\nModo análise (SIM/NAO): ");
                if (modo == null) return;
                if (modo == "SIM")
                {
                    // Em análise, cada clique aplica a ação. Enter ou Esc fecha a sequência.
                    for (;;)
                    {
                        ObjectId nota = PedirNota(ed);
                        if (nota.IsNull) break;
                        if (ConfirmarLayerNota(s, new[] { nota }, acao)) AuditoriaLayersCad.Aplicar(s, new[] { nota }, acao, true);
                    }
                    return;
                }
                var selecionadas = SelecionarNotas(ed, "\nSelecione os MTEXT: ");
                if (selecionadas.Length > 0 && ConfirmarLayerNota(s, selecionadas, acao)) AuditoriaLayersCad.Aplicar(s, selecionadas, acao, false);
            });
        }

        private static bool ConfirmarLayerNota(Sessao s, ObjectId[] notas, string acao)
        {
            if (acao == "EXCLUIR") return Confirmar(s.Doc.Editor, "Excluir do DWG os MTEXT selecionados e suas guias do OrçaCore?");
            return !AuditoriaLayersCad.MudaOperacao(s, notas, acao) || Confirmar(s.Doc.Editor, "Mover o texto para outra layer de operação no DWG?");
        }

        private static string PedirPalavra(Editor ed, string mensagem)
        {
            var resposta = ed.GetString(new PromptStringOptions(mensagem) { AllowSpaces = false });
            return resposta.Status == PromptStatus.OK ? (resposta.StringResult ?? "").Trim().ToUpperInvariant() : null;
        }

        private static void AplicarNota(Database db, Transaction tr, MText nota, TelaNota tela, RegrasProjeto regras)
        {
            nota.Contents = tela.Conteudo;
            nota.TextHeight = tela.Altura;
            if (tela.Alinhamento > 0) nota.Attachment = (AttachmentPoint)tela.Alinhamento;
            string layer = regras.LayerDaOperacao(tela.Operacao);
            if (layer != null) DefinirLayer(db, tr, nota.ObjectId, layer);
        }

        internal static void NotaAlterada(Sessao s)
        {
            s.UltimoBruto = null;
            s.UltimoDesenho = null;
            if (s.Obra != null) Aplicacao.Msg("DWG alterado. Use Atualizar leitura no painel para trazer as mudanças para a obra.");
            TelaAuditoriaLayers.Atualizar(s);
        }

        internal static ObjectId[] SelecionarNotas(Editor editor, string mensagem)
        {
            var filtro = new SelectionFilter(new[] { new TypedValue((int)DxfCode.Start, "MTEXT") });
            PromptSelectionResult selecao = editor.GetSelection(new PromptSelectionOptions { MessageForAdding = mensagem }, filtro);
            return selecao.Status == PromptStatus.OK ? selecao.Value.GetObjectIds() : new ObjectId[0];
        }

        private static ObjectId PedirNota(Editor ed)
        {
            var opcoes = new PromptEntityOptions("\nSelecione a anotação MTEXT: ");
            opcoes.AllowNone = true;
            opcoes.SetRejectMessage("\nSelecione um MTEXT.");
            opcoes.AddAllowedClass(typeof(MText), false);
            PromptEntityResult r = ed.GetEntity(opcoes);
            return r.Status == PromptStatus.OK ? r.ObjectId : ObjectId.Null;
        }

        private static Point3d? PontoDaNota(Editor ed, string mensagem)
        {
            PromptPointResult r = ed.GetPoint(new PromptPointOptions(mensagem) { AllowNone = true });
            return r.Status == PromptStatus.OK ? (Point3d?)r.Value.TransformBy(ed.CurrentUserCoordinateSystem) : null;
        }
    }
}
