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

        // Um leitor por tipo de entidade, cada um no seu método (TASK-014): no nanoCAD 5, um tipo que a API não tenha (ex.: MLeader)
        // só desliga o SEU leitor (Compat), sem derrubar a leitura do desenho. Ordem importa (AttributeReference é DBText).
        private delegate bool LeitorEntidade(Entity ent, Transaction tr, bool comDefinicao, out EntidadeBruta b);

        private static readonly KeyValuePair<string, LeitorEntidade>[] Leitores =
        {
            new KeyValuePair<string, LeitorEntidade>("MTEXT", LerMText),
            new KeyValuePair<string, LeitorEntidade>("ATTRIB", LerAtributo),
            new KeyValuePair<string, LeitorEntidade>("ATTDEF", LerDefinicaoDeAtributo),
            new KeyValuePair<string, LeitorEntidade>("TEXT", LerTexto),
            new KeyValuePair<string, LeitorEntidade>("MLEADER", LerMultileader),
            new KeyValuePair<string, LeitorEntidade>("INSERT", LerBloco),
            new KeyValuePair<string, LeitorEntidade>("CIRCLE", LerCirculo),
            new KeyValuePair<string, LeitorEntidade>("LINE", LerLinha),
            new KeyValuePair<string, LeitorEntidade>("LWPOLYLINE", LerPolilinha),
            new KeyValuePair<string, LeitorEntidade>("POLYLINE 2D", LerPolilinha2d),
            new KeyValuePair<string, LeitorEntidade>("POLYLINE 3D", LerPolilinha3d),
            new KeyValuePair<string, LeitorEntidade>("LEADER", LerLeader),
        };

        /// <summary>Uma entidade → bruta (null = tipo que o leitor não usa).</summary>
        public static EntidadeBruta Bruta(Entity ent, Transaction tr, bool comDefinicao)
        {
            foreach (KeyValuePair<string, LeitorEntidade> l in Leitores)
            {
                EntidadeBruta b = null;
                LeitorEntidade f = l.Value;
                bool achou = false;
                if (!Compat.Tentar("leitura de " + l.Key, () => { achou = f(ent, tr, comDefinicao, out b); })) continue;
                if (achou) return b;
            }
            return null;
        }

        private static double[] CaixaSegura(Entity e) { return Compat.Tentar("contorno das entidades", () => Caixa(e), null); }

        private static bool LerMText(Entity ent, Transaction tr, bool def, out EntidadeBruta b)
        {
            b = null;
            MText mt = ent as MText;
            if (mt == null) return false;
            b = Base(ent, "MTEXT");
            b.Texto = mt.Contents; b.X = mt.Location.X; b.Y = mt.Location.Y; b.Altura = mt.TextHeight; b.Caixa = CaixaSegura(mt);
            return true;
        }

        private static bool LerAtributo(Entity ent, Transaction tr, bool def, out EntidadeBruta b)
        {
            b = null;
            AttributeReference ar = ent as AttributeReference;
            if (ar == null) return false;
            b = Base(ent, "ATTRIB");
            b.Texto = ar.TextString; b.X = ar.Position.X; b.Y = ar.Position.Y; b.Altura = ar.Height; b.Caixa = CaixaSegura(ar);
            AttributeReference a = ar;
            EntidadeBruta bb = b;
            Compat.Tentar("atributo invisível", () => { bb.Invisivel = a.Invisible; });
            return true;
        }

        private static bool LerDefinicaoDeAtributo(Entity ent, Transaction tr, bool def, out EntidadeBruta b)
        {
            b = null;
            return ent is AttributeDefinition;
        }

        private static bool LerTexto(Entity ent, Transaction tr, bool def, out EntidadeBruta b)
        {
            b = null;
            DBText dt = ent as DBText;
            if (dt == null) return false;
            b = Base(ent, "TEXT");
            b.Texto = dt.TextString; b.X = dt.Position.X; b.Y = dt.Position.Y; b.Altura = dt.Height; b.Caixa = CaixaSegura(dt);
            return true;
        }

        private static bool LerMultileader(Entity ent, Transaction tr, bool def, out EntidadeBruta b)
        {
            b = null;
            MLeader ml = ent as MLeader;
            if (ml == null) return false;
            b = Multileader(ml);
            return true;
        }

        private static bool LerBloco(Entity ent, Transaction tr, bool comDefinicao, out EntidadeBruta b)
        {
            b = null;
            BlockReference br = ent as BlockReference;
            if (br == null) return false;
            b = Base(ent, "INSERT");
            b.Bloco = Compat.Tentar("blocos dinâmicos", () => NomeDoBloco(br, tr), null) ?? br.Name;
            b.X = br.Position.X; b.Y = br.Position.Y;
            foreach (ObjectId aid in br.AttributeCollection)
            {
                AttributeReference a = tr.GetObject(aid, OpenMode.ForRead) as AttributeReference;
                if (a != null) { EntidadeBruta f = Bruta(a, tr, false); if (f != null) b.Filhos.Add(f); }
            }
            if (comDefinicao)
            {
                // como a aplicação: TEXT/MTEXT da DEFINIÇÃO do bloco, com as coordenadas da definição
                BlockTableRecord defn = tr.GetObject(br.BlockTableRecord, OpenMode.ForRead) as BlockTableRecord;
                if (defn != null)
                    foreach (ObjectId id in defn)
                    {
                        Entity e = tr.GetObject(id, OpenMode.ForRead) as Entity;
                        if ((e is DBText && !(e is AttributeDefinition)) || e is MText) { EntidadeBruta f = Bruta(e, tr, false); if (f != null) b.Definicao.Add(f); }
                    }
            }
            return true;
        }

        private static bool LerCirculo(Entity ent, Transaction tr, bool def, out EntidadeBruta b)
        {
            b = null;
            Circle c = ent as Circle;
            if (c == null) return false;
            b = Base(ent, "CIRCLE");
            b.X = c.Center.X; b.Y = c.Center.Y; b.Raio = c.Radius;
            return true;
        }

        private static EntidadeBruta ComPontos(Entity ent, string tipo, List<Ponto> pts)
        {
            EntidadeBruta r = Base(ent, tipo);
            r.Pontos = pts;
            return r;
        }

        private static bool LerLinha(Entity ent, Transaction tr, bool def, out EntidadeBruta b)
        {
            b = null;
            Line ln = ent as Line;
            if (ln == null) return false;
            b = ComPontos(ent, "LINE", new List<Ponto> { P(ln.StartPoint), P(ln.EndPoint) });
            return true;
        }

        private static bool LerPolilinha(Entity ent, Transaction tr, bool def, out EntidadeBruta b)
        {
            b = null;
            Polyline pl = ent as Polyline;
            if (pl == null) return false;
            List<Ponto> pts = new List<Ponto>();
            for (int i = 0; i < pl.NumberOfVertices; i++) { Point2d q = pl.GetPoint2dAt(i); pts.Add(new Ponto(q.X, q.Y)); }
            if (pl.Closed && pts.Count > 0) pts.Add(pts[0]);
            b = ComPontos(ent, "LWPOLYLINE", pts);
            return true;
        }

        private static bool LerPolilinha2d(Entity ent, Transaction tr, bool def, out EntidadeBruta b)
        {
            b = null;
            Polyline2d p2 = ent as Polyline2d;
            if (p2 == null) return false;
            List<Ponto> pts = new List<Ponto>();
            foreach (ObjectId vid in p2) { Vertex2d v = tr.GetObject(vid, OpenMode.ForRead) as Vertex2d; if (v != null) pts.Add(P(v.Position)); }
            if (p2.Closed && pts.Count > 0) pts.Add(pts[0]);
            b = ComPontos(ent, "POLYLINE", pts);
            return true;
        }

        private static bool LerPolilinha3d(Entity ent, Transaction tr, bool def, out EntidadeBruta b)
        {
            b = null;
            Polyline3d p3 = ent as Polyline3d;
            if (p3 == null) return false;
            List<Ponto> pts = new List<Ponto>();
            foreach (ObjectId vid in p3) { PolylineVertex3d v = tr.GetObject(vid, OpenMode.ForRead) as PolylineVertex3d; if (v != null) pts.Add(P(v.Position)); }
            if (p3.Closed && pts.Count > 0) pts.Add(pts[0]);
            b = ComPontos(ent, "POLYLINE", pts);
            return true;
        }

        private static bool LerLeader(Entity ent, Transaction tr, bool def, out EntidadeBruta b)
        {
            b = null;
            Leader ld = ent as Leader;
            if (ld == null) return false;
            List<Ponto> pts = new List<Ponto>();
            for (int i = 0; i < ld.NumVertices; i++) pts.Add(P(ld.VertexAt(i)));
            b = ComPontos(ent, "LEADER", pts);
            return true;
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
