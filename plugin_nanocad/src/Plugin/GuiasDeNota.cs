using System;
using System.Collections.Generic;
using HostMgd.ApplicationServices;
using OrcamentoNanoCAD.Nucleo;
using Teigha.DatabaseServices;
using Teigha.Geometry;

namespace OrcamentoNanoCAD
{
    internal static class GuiasDeNota
    {
        private const string Aplicativo = "ORCACORE_NOTAS";
        private const string Layer = "ORC_GUIAS";
        private const int NomeApp = 1001, Referencia = 1005;
        private static readonly HashSet<string> ComandosDeMovimento = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        { "MOVE", "STRETCH", "GRIP_STRETCH", "GRIP_MOVE", "GRIP_ROTATE", "GRIP_SCALE", "GRIP_MIRROR", "ROTATE", "SCALE", "ALIGN", "PROPERTIES" };

        public static void Iniciar()
        {
            Application.DocumentManager.DocumentActivated += (sender, args) => Ligar(args.Document);
            Ligar(Aplicacao.Doc());
        }

        private static void Ligar(Document doc)
        {
            if (doc == null) return;
            doc.CommandEnded -= AoTerminar;
            doc.CommandEnded += AoTerminar;
        }

        private static void AoTerminar(object sender, CommandEventArgs args)
        {
            Document doc = sender as Document;
            if (doc == null || !ComandosDeMovimento.Contains((args.GlobalCommandName ?? "").TrimStart('_', '.'))) return;
            try
            {
                Atualizar(doc);
                Sessao s = Aplicacao.Sessao(doc, false);
                if (s != null)
                {
                    s.UltimoBruto = null;
                    s.UltimoDesenho = null;
                    UI.TelaAuditoriaLayers.Atualizar(s);
                }
            }
            catch (System.Exception erro) { Aplicacao.Log("Atualizar guias de nota: " + erro.Message); }
        }

        public static void Atualizar(Document doc)
        {
            Database db = doc.Database;
            using (doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                var aplicativos = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
                if (!aplicativos.Has(Aplicativo)) return;
                AtualizarNoModelo(db, tr);
                tr.Commit();
            }
        }

        public static BlockTableRecord Modelo(Database db, Transaction tr, OpenMode modo)
        {
            var blocos = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            return (BlockTableRecord)tr.GetObject(blocos[BlockTableRecord.ModelSpace], modo);
        }

        private static string[] Referencias(Line linha)
        {
            using (ResultBuffer dados = linha.GetXDataForApplication(Aplicativo))
            {
                if (dados == null) return null;
                var handles = new List<string>();
                foreach (TypedValue valor in dados)
                    if (valor.TypeCode == Referencia) handles.Add(valor.Value as string);
                return handles.Count == 2 ? handles.ToArray() : null;
            }
        }

        public static string NotaDaGuia(Line linha)
        {
            return Compat.Tentar("referências das guias de nota", () => { string[] refs = Referencias(linha); return refs == null ? null : refs[0]; }, (string)null);
        }

        public static ObjectId Id(Database db, string handle)
        {
            long numero;
            if (!long.TryParse(handle, System.Globalization.NumberStyles.HexNumber, System.Globalization.CultureInfo.InvariantCulture, out numero))
                return ObjectId.Null;
            try { return db.GetObjectId(false, new Handle(numero), 0); }
            catch (Teigha.Runtime.Exception) { return ObjectId.Null; }
        }

        private static Entity Entidade(Database db, Transaction tr, string handle)
        {
            ObjectId id = Id(db, handle);
            return id.IsNull || id.IsErased ? null : tr.GetObject(id, OpenMode.ForRead) as Entity;
        }

        public static void AtualizarNoModelo(Database db, Transaction tr, HashSet<string> notasExcluidas = null)
        {
            foreach (ObjectId id in Modelo(db, tr, OpenMode.ForRead))
            {
                if (id.IsErased || id.ObjectClass.DxfName != "LINE") continue;
                var linha = tr.GetObject(id, OpenMode.ForRead) as Line;
                if (linha == null) continue;
                string[] refs = Referencias(linha);
                if (refs == null) continue;
                if (notasExcluidas != null)
                {
                    if (notasExcluidas.Contains(refs[0])) { linha.UpgradeOpen(); linha.Erase(); }
                    continue;
                }
                var nota = Entidade(db, tr, refs[0]) as MText;
                var alvo = Entidade(db, tr, refs[1]);
                if (nota != null && alvo != null) AtualizarExtremos(linha, nota, alvo, true);
            }
        }

        public static void Criar(Database db, Transaction tr, MText nota, Entity alvo)
        {
            var aplicativos = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
            if (!aplicativos.Has(Aplicativo))
            {
                aplicativos.UpgradeOpen();
                var registro = new RegAppTableRecord { Name = Aplicativo };
                aplicativos.Add(registro);
                tr.AddNewlyCreatedDBObject(registro, true);
            }
            var linha = new Line(nota.Location, Destino(alvo, nota.Location));
            AtualizarExtremos(linha, nota, alvo, false);
            Modelo(db, tr, OpenMode.ForWrite).AppendEntity(linha);
            tr.AddNewlyCreatedDBObject(linha, true);
            Comandos.DefinirLayer(db, tr, linha.ObjectId, Layer);
            using (var dados = new ResultBuffer(new TypedValue(NomeApp, Aplicativo),
                new TypedValue(Referencia, nota.Handle.ToString()), new TypedValue(Referencia, alvo.Handle.ToString())))
                linha.XData = dados;
        }

        private static Point3d Destino(Entity alvo, Point3d de)
        {
            var bloco = alvo as BlockReference;
            if (bloco != null) return bloco.Position;
            var circulo = alvo as Circle;
            if (circulo != null) return circulo.Center;
            var curva = alvo as Curve;
            if (curva != null) return curva.GetClosestPointTo(de, false);
            var caixa = alvo.GeometricExtents;
            return caixa.MinPoint + (caixa.MaxPoint - caixa.MinPoint) / 2;
        }

        private static void AtualizarExtremos(Line linha, MText nota, Entity alvo, bool abrir)
        {
            Point3d destino = Destino(alvo, nota.Location);
            Extents3d caixa = nota.GeometricExtents;
            Ponto borda = AnotacaoCad.CentroDaBorda(new Ponto(caixa.MinPoint.X, caixa.MinPoint.Y),
                new Ponto(caixa.MaxPoint.X, caixa.MaxPoint.Y), new Ponto(destino.X, destino.Y));
            var inicio = new Point3d(borda.X, borda.Y, nota.Location.Z);
            if (linha.StartPoint.DistanceTo(inicio) < 1e-6 && linha.EndPoint.DistanceTo(destino) < 1e-6) return;
            if (abrir) linha.UpgradeOpen();
            linha.StartPoint = inicio;
            linha.EndPoint = destino;
        }
    }
}
