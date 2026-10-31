// Plugin/ExtratorDesenho.cs — lê o DWG aberto para o modelo do núcleo (Desenho): textos (TEXT, MTEXT por parágrafo, atributos),
// blocos de poste (nome contém "poste") e linhas (LINE, LWPOLYLINE, POLYLINE 2D, LEADER). Cor resolvida como o leitor de DXF da aplicação.
using System;
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
        public static string Hex(Color c)
        {
            System.Drawing.Color v = c.ColorValue;
            return Cores.Hex(v.R, v.G, v.B);
        }

        /// <summary>true color → ACI → layer TXT_RRGGBB → cor da layer → preto. ByBlock dentro de bloco = cor do bloco.</summary>
        public static string CorDe(Entity ent, Transaction tr, string corDoBloco)
        {
            Color c = ent.Color;
            if (c.ColorMethod == ColorMethod.ByColor) return Hex(c);
            if (c.ColorMethod == ColorMethod.ByAci) return Cores.AciParaHex(c.ColorIndex);
            if (c.ColorMethod == ColorMethod.ByBlock && corDoBloco != null) return corDoBloco;
            string h = Cores.CorDoNomeDaLayer(ent.Layer);
            if (h != null) return h;
            try
            {
                LayerTableRecord ltr = tr.GetObject(ent.LayerId, OpenMode.ForRead) as LayerTableRecord;
                if (ltr != null)
                {
                    Color lc = ltr.Color;
                    if (lc.ColorMethod == ColorMethod.ByColor) return Hex(lc);
                    return Cores.AciParaHex(lc.ColorIndex);
                }
            }
            catch (System.Exception) { }
            return "#000000";
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

        public static bool EhPoste(string nomeBloco, RegrasProjeto regras)
        {
            string filtro = string.IsNullOrEmpty(regras.FiltroBlocoPoste) ? "poste" : regras.FiltroBlocoPoste;
            return (nomeBloco ?? "").IndexOf(filtro, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static List<TextoDesenho> Textos(Entity ent, Transaction tr, string corDoBloco)
        {
            List<TextoDesenho> r = new List<TextoDesenho>();
            string handle = ent.Handle.ToString();
            MText mt = ent as MText;
            if (mt != null)
            {
                string corBase = CorDe(mt, tr, corDoBloco);
                int parte = 0;
                foreach (ParteTexto p in Cores.DividirMText(mt.Contents, corBase))
                {
                    TextoDesenho t = new TextoDesenho();
                    t.Handle = handle; t.Parte = parte++; t.Texto = p.Texto; t.Cor = p.Cor; t.Layer = mt.Layer;
                    t.X = mt.Location.X; t.Y = mt.Location.Y; t.Altura = mt.TextHeight; t.Tipo = "MTEXT";
                    r.Add(t);
                }
                return r;
            }
            DBText dt = ent as DBText;   // TEXT e ATTRIB (AttributeReference herda de DBText)
            if (dt != null)
            {
                AttributeReference ar = dt as AttributeReference;
                if (ar != null && ar.Invisible) return r;
                string txt = Cores.LimparTextoSimples(dt.TextString);
                if (txt.Length == 0) return r;
                TextoDesenho t = new TextoDesenho();
                t.Handle = handle; t.Texto = txt; t.Cor = CorDe(dt, tr, corDoBloco); t.Layer = dt.Layer;
                t.X = dt.Position.X; t.Y = dt.Position.Y; t.Altura = dt.Height; t.Tipo = ar != null ? "ATTRIB" : "TEXT";
                r.Add(t);
            }
            return r;
        }

        public static LinhaDesenho Linha(Entity ent, Transaction tr)
        {
            List<Ponto> pts = new List<Ponto>();
            bool leader = false;
            Line ln = ent as Line;
            Polyline pl = ent as Polyline;
            Polyline2d p2 = ent as Polyline2d;
            Leader ld = ent as Leader;
            if (ln != null) { pts.Add(P(ln.StartPoint)); pts.Add(P(ln.EndPoint)); }
            else if (pl != null)
            {
                for (int i = 0; i < pl.NumberOfVertices; i++) { Point2d q = pl.GetPoint2dAt(i); pts.Add(new Ponto(q.X, q.Y)); }
                if (pl.Closed && pts.Count > 0) pts.Add(pts[0]);
            }
            else if (p2 != null)
            {
                foreach (ObjectId vid in p2)
                {
                    Vertex2d v = tr.GetObject(vid, OpenMode.ForRead) as Vertex2d;
                    if (v != null) pts.Add(P(v.Position));
                }
                if (p2.Closed && pts.Count > 0) pts.Add(pts[0]);
            }
            else if (ld != null)
            {
                leader = true;
                for (int i = 0; i < ld.NumVertices; i++) pts.Add(P(ld.VertexAt(i)));
            }
            else return null;
            LinhaDesenho l = new LinhaDesenho();
            l.Handle = ent.Handle.ToString(); l.Layer = ent.Layer; l.Cor = CorDe(ent, tr, null); l.Pontos = pts; l.Leader = leader;
            return l;
        }

        private static Ponto P(Point3d p) { return new Ponto(p.X, p.Y); }

        public static PosteDesenho Poste(BlockReference br, Transaction tr)
        {
            PosteDesenho p = new PosteDesenho();
            p.Handle = br.Handle.ToString(); p.X = br.Position.X; p.Y = br.Position.Y; p.Bloco = NomeDoBloco(br, tr); p.Layer = br.Layer;
            return p;
        }

        /// <summary>O espaço do modelo inteiro.</summary>
        public static Desenho Ler(Document doc, RegrasProjeto regras)
        {
            Desenho d = new Desenho();
            d.Arquivo = doc.Name;
            Database db = doc.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                foreach (ObjectId id in ms)
                {
                    Entity ent;
                    try { ent = tr.GetObject(id, OpenMode.ForRead) as Entity; }
                    catch (System.Exception) { continue; }
                    if (ent == null) continue;
                    try { Acrescentar(d, ent, tr, regras); }
                    catch (System.Exception) { /* entidade estranha não derruba a leitura */ }
                }
                tr.Commit();
            }
            return d;
        }

        public static void Acrescentar(Desenho d, Entity ent, Transaction tr, RegrasProjeto regras)
        {
            BlockReference br = ent as BlockReference;
            if (br != null)
            {
                string nome = NomeDoBloco(br, tr);
                string corBloco = CorDe(br, tr, null);
                if (EhPoste(nome, regras)) d.Postes.Add(Poste(br, tr));
                TextoDesenho marca = new TextoDesenho();      // o leitor de DXF da aplicação também lista o nome do bloco
                marca.Handle = br.Handle.ToString(); marca.Texto = "[BLOCO: " + nome + "]"; marca.Cor = corBloco; marca.Layer = br.Layer;
                marca.X = br.Position.X; marca.Y = br.Position.Y; marca.Tipo = "BLOCO";
                d.Textos.Add(marca);
                foreach (ObjectId aid in br.AttributeCollection)
                {
                    AttributeReference ar = tr.GetObject(aid, OpenMode.ForRead) as AttributeReference;
                    if (ar != null) d.Textos.AddRange(Textos(ar, tr, corBloco));
                }
                return;
            }
            if (ent is MText || ent is DBText) { d.Textos.AddRange(Textos(ent, tr, null)); return; }
            LinhaDesenho l = Linha(ent, tr);
            if (l != null && l.Pontos.Count >= 2) d.Linhas.Add(l);
        }
    }
}
