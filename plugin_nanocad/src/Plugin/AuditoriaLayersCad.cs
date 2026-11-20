using System;
using System.Collections.Generic;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Teigha.DatabaseServices;

namespace OrcamentoNanoCAD
{
    internal static class AuditoriaLayersCad
    {
        public static List<EntidadeBruta> Ler(Sessao s)
        {
            var db = s.Doc.Database;
            var textos = new List<EntidadeBruta>();
            using (var tr = db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId id in GuiasDeNota.Modelo(db, tr, OpenMode.ForRead))
                {
                    if (id.ObjectClass.DxfName != "MTEXT") continue;
                    var nota = tr.GetObject(id, OpenMode.ForRead) as MText;
                    if (nota != null) textos.Add(ExtratorDesenho.Bruta(nota, tr, false));
                }
            }
            return textos;
        }

        public static bool MudaOperacao(Sessao s, ObjectId[] ids, string acao)
        {
            if (acao == "LV") return false;
            var mapa = s.Regras.OperacoesPorLayer();
            using (var tr = s.Doc.Database.TransactionManager.StartTransaction())
                foreach (ObjectId id in ids)
                {
                    var nota = tr.GetObject(id, OpenMode.ForRead) as MText;
                    string atual;
                    if (nota != null && (!mapa.TryGetValue(nota.Layer, out atual) || atual != acao)) return true;
                }
            return false;
        }

        public static void Aplicar(Sessao s, ObjectId[] ids, string acao, bool alternarComZero)
        {
            if (ids.Length == 0) return;
            var operacoesPorLayer = s.Regras.OperacoesPorLayer();
            string layerAlvo = s.Regras.LayerDaOperacao(acao);
            if (acao != "LV" && acao != "EXCLUIR" && acao != "0" && layerAlvo == null)
                throw new InvalidOperationException("O projeto não tem uma layer para a operação " + acao + ".");
            Database db = s.Doc.Database;
            using (s.Doc.LockDocument())
            using (var tr = db.TransactionManager.StartTransaction())
            {
                var excluidas = new HashSet<string>();
                foreach (ObjectId id in ids.Distinct())
                {
                    if (id.IsNull || id.IsErased) continue;
                    var nota = tr.GetObject(id, OpenMode.ForRead) as MText;
                    if (nota == null) continue;
                    if (acao == "EXCLUIR")
                    {
                        excluidas.Add(nota.Handle.ToString());
                        nota.UpgradeOpen(); nota.Erase();
                        continue;
                    }
                    string destino = acao == "0" ? "0" : layerAlvo;
                    if (acao == "LV")
                    {
                        string atual;
                        if (!operacoesPorLayer.TryGetValue(nota.Layer, out atual))
                            throw new InvalidOperationException("Defina primeiro a operação das notas fora das layers do projeto.");
                        destino = s.Regras.LayerDaOperacao(AnotacaoCad.AlternarLinhaViva(atual));
                        if (destino == null) throw new InvalidOperationException("O projeto não tem a layer correspondente para alternar linha viva.");
                    }
                    if (alternarComZero && acao != "LV" && destino == nota.Layer) destino = "0";
                    if (!string.Equals(destino, nota.Layer, StringComparison.OrdinalIgnoreCase)) Comandos.DefinirLayer(db, tr, id, destino);
                }
                if (excluidas.Count > 0) GuiasDeNota.AtualizarNoModelo(db, tr, excluidas);
                tr.Commit();
            }
            Comandos.NotaAlterada(s);
        }
    }
}
