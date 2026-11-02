// Plugin/ExtratorDesenho.cs — lê o DWG aberto como DesenhoBruto (o que o CAD entrega, sem interpretar): textos com o conteúdo
// original, cores sem resolver, blocos (atributos + textos da definição), linhas, LEADER, MLEADER (texto + linhas de chamada),
// círculos e o contorno de cada texto. A interpretação (cor final, parágrafos, postes, redes, guias) é do núcleo (MontadorDesenho).
using System;
using System.Collections;
using System.Collections.Generic;
using HostMgd.ApplicationServices;
using OrcamentoNanoCAD.Nucleo;
using Teigha.Colors;
using Teigha.DatabaseServices;
using Teigha.Geometry;

namespace OrcamentoNanoCAD
{
    public static class ExtratorDesenho
    {
        public static CorBruta CorBrutaDe(Color c)
        {
            CorBruta r = new CorBruta();
            if (c == null) return r;
            if (c.ColorMethod == ColorMethod.ByColor)
            {
                System.Drawing.Color v = c.ColorValue;
                r.Metodo = "rgb"; r.Rgb = (v.R << 16) | (v.G << 8) | v.B;
            }
            else if (c.ColorMethod == ColorMethod.ByAci) { r.Metodo = "aci"; r.Aci = c.ColorIndex; }
            else if (c.ColorMethod == ColorMethod.ByBlock) r.Metodo = "block";
            else r.Metodo = "layer";
            return r;
        }

        public static string NomeDoBloco(BlockReference br, Transaction tr)
        {
            try
            {
                if (br.IsDynamicBlock)
                {
                    BlockTableRecord btr = tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    if (btr != null) return btr.Name;
                }
            }
            catch (System.Exception) { }
            return br.Name;
        }

        private static double[] Caixa(Entity e)
        {
            try
            {
                Extents3d x = e.GeometricExtents;
                return new[] { x.MinPoint.X, x.MinPoint.Y, x.MaxPoint.X, x.MaxPoint.Y };
            }
            catch (System.Exception) { return null; }
        }

        private static Ponto P(Point3d p) { return new Ponto(p.X, p.Y); }

        private static EntidadeBruta Base(Entity ent, string tipo)
        {
            EntidadeBruta b = new EntidadeBruta();
            b.Tipo = tipo; b.Handle = ent.Handle.ToString(); b.Layer = ent.Layer; b.Cor = CorBrutaDe(ent.Color);
            return b;
        }

        /// <summary>Uma entidade → bruta (null = tipo que o leitor não usa).</summary>
        public static EntidadeBruta Bruta(Entity ent, Transaction tr, bool comDefinicao)
        {
            MText mt = ent as MText;
            if (mt != null)
            {
                EntidadeBruta b = Base(ent, "MTEXT");
                b.Texto = mt.Contents; b.X = mt.Location.X; b.Y = mt.Location.Y; b.Altura = mt.TextHeight; b.Caixa = Caixa(mt);
                return b;
            }
            AttributeReference ar = ent as AttributeReference;
            if (ar != null)
            {
                EntidadeBruta b = Base(ent, "ATTRIB");
                b.Texto = ar.TextString; b.X = ar.Position.X; b.Y = ar.Position.Y; b.Altura = ar.Height; b.Invisivel = ar.Invisible; b.Caixa = Caixa(ar);
                return b;
            }
            if (ent is AttributeDefinition) return null;
            DBText dt = ent as DBText;
            if (dt != null)
            {
                EntidadeBruta b = Base(ent, "TEXT");
                b.Texto = dt.TextString; b.X = dt.Position.X; b.Y = dt.Position.Y; b.Altura = dt.Height; b.Caixa = Caixa(dt);
                return b;
            }
            MLeader ml = ent as MLeader;
            if (ml != null) return Multileader(ml);
            BlockReference br = ent as BlockReference;
            if (br != null)
            {
                EntidadeBruta b = Base(ent, "INSERT");
                b.Bloco = NomeDoBloco(br, tr); b.X = br.Position.X; b.Y = br.Position.Y;
                foreach (ObjectId aid in br.AttributeCollection)
                {
                    AttributeReference a = tr.GetObject(aid, OpenMode.ForRead) as AttributeReference;
                    if (a != null) b.Filhos.Add(Bruta(a, tr, false));
                }
                if (comDefinicao)
                {
                    // como a aplicação: TEXT/MTEXT da DEFINIÇÃO do bloco, com as coordenadas da definição
                    BlockTableRecord def = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                    if (def != null)
                        foreach (ObjectId id in def)
                        {
                            Entity e = tr.GetObject(id, OpenMode.ForRead) as Entity;
                            if ((e is DBText && !(e is AttributeDefinition)) || e is MText) b.Definicao.Add(Bruta(e, tr, false));
                        }
                }
                return b;
            }
            Circle c = ent as Circle;
            if (c != null)
            {
                EntidadeBruta b = Base(ent, "CIRCLE");
                b.X = c.Center.X; b.Y = c.Center.Y; b.Raio = c.Radius;
                return b;
            }
            List<Ponto> pts = new List<Ponto>();
            string tipo;
            Line ln = ent as Line;
            Polyline pl = ent as Polyline;
            Polyline2d p2 = ent as Polyline2d;
            Polyline3d p3 = ent as Polyline3d;
            Leader ld = ent as Leader;
            if (ln != null) { tipo = "LINE"; pts.Add(P(ln.StartPoint)); pts.Add(P(ln.EndPoint)); }
            else if (pl != null)
            {
                tipo = "LWPOLYLINE";
                for (int i = 0; i < pl.NumberOfVertices; i++) { Point2d q = pl.GetPoint2dAt(i); pts.Add(new Ponto(q.X, q.Y)); }
                if (pl.Closed && pts.Count > 0) pts.Add(pts[0]);
            }
            else if (p2 != null)
            {
                tipo = "POLYLINE";
                foreach (ObjectId vid in p2) { Vertex2d v = tr.GetObject(vid, OpenMode.ForRead) as Vertex2d; if (v != null) pts.Add(P(v.Position)); }
                if (p2.Closed && pts.Count > 0) pts.Add(pts[0]);
            }
            else if (p3 != null)
            {
                tipo = "POLYLINE";
                foreach (ObjectId vid in p3) { PolylineVertex3d v = tr.GetObject(vid, OpenMode.ForRead) as PolylineVertex3d; if (v != null) pts.Add(P(v.Position)); }
                if (p3.Closed && pts.Count > 0) pts.Add(pts[0]);
            }
            else if (ld != null) { tipo = "LEADER"; for (int i = 0; i < ld.NumVertices; i++) pts.Add(P(ld.VertexAt(i))); }
            else return null;
            EntidadeBruta r = Base(ent, tipo);
            r.Pontos = pts;
            return r;
        }

        private static EntidadeBruta Multileader(MLeader ml)
        {
            EntidadeBruta b = Base(ml, "MLEADER");
            try
            {
                if (ml.ContentType == ContentType.MTextContent && ml.MText != null)
                {
                    MText m = ml.MText;
                    b.Texto = m.Contents; b.X = m.Location.X; b.Y = m.Location.Y; b.Altura = m.TextHeight;
                    try { Extents3d x = ml.GetContentGeomExtents(); b.Caixa = new[] { x.MinPoint.X, x.MinPoint.Y, x.MaxPoint.X, x.MaxPoint.Y }; }
                    catch (System.Exception) { }
                }
            }
            catch (System.Exception) { }
            try
            {
                foreach (object li in ml.GetLeaderIndexes())
                    foreach (object ll in ml.GetLeaderLineIndexes((int)li))
                    {
                        int linha = (int)ll;
                        List<Ponto> pts = new List<Ponto>();
                        for (int k = 0; k < ml.VerticesCount(linha); k++) pts.Add(P(ml.GetVertex(linha, k)));
                        if (pts.Count >= 2) b.Linhas.Add(pts);
                    }
            }
            catch (System.Exception) { }
            return b;
        }

        public static CorBruta CorDaLayer(LayerTableRecord ltr)
        {
            CorBruta c = CorBrutaDe(ltr.Color);
            if (c.Metodo != "rgb") { c.Metodo = "aci"; c.Aci = ltr.Color.ColorIndex; }
            return c;
        }

        /// <summary>O espaço do modelo inteiro.</summary>
        public static DesenhoBruto Ler(Document doc)
        {
            DesenhoBruto d = new DesenhoBruto();
            d.Arquivo = doc.Name;
            Database db = doc.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId lid in lt)
                {
                    LayerTableRecord ltr = tr.GetObject(lid, OpenMode.ForRead) as LayerTableRecord;
                    if (ltr != null) d.Layers[ltr.Name] = CorDaLayer(ltr);
                }
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                foreach (ObjectId id in ms)
                {
                    Entity ent = null;
                    try
                    {
                        ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if (ent == null) continue;
                        EntidadeBruta b = Bruta(ent, tr, true);
                        if (b != null) d.Entidades.Add(b);
                    }
                    catch (System.Exception e)
                    {
                        d.Avisos.Add((ent != null ? ent.GetType().Name + " " + ent.Handle : id.ToString()) + ": " + e.Message);   // não derruba a leitura
                    }
                }
                tr.Commit();
            }
            return d;
        }

        /// <summary>Entidades escolhidas pelo usuário (Construir do zero) → brutas, com as cores das layers.</summary>
        public static DesenhoBruto LerSelecao(Document doc, IEnumerable<ObjectId> ids)
        {
            DesenhoBruto d = new DesenhoBruto();
            d.Arquivo = doc.Name;
            Database db = doc.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                foreach (ObjectId lid in lt)
                {
                    LayerTableRecord ltr = tr.GetObject(lid, OpenMode.ForRead) as LayerTableRecord;
                    if (ltr != null) d.Layers[ltr.Name] = CorDaLayer(ltr);
                }
                foreach (ObjectId id in ids)
                {
                    try
                    {
                        Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        EntidadeBruta b = ent == null ? null : Bruta(ent, tr, false);
                        if (b != null) d.Entidades.Add(b);
                    }
                    catch (System.Exception e) { d.Avisos.Add(id + ": " + e.Message); }
                }
                tr.Commit();
            }
            return d;
        }
    }
}
