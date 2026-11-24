// Plugin/Marcas.cs — números (P1…, V1…) e linhas de vínculo vão ⇄ poste desenhados como GRÁFICOS TRANSITÓRIOS: aparecem na tela,
// não entram no DWG (sem camada de marcas). Somem ao fechar o desenho; o comando ORC_MARCAS liga/desliga.
using System;
using System.Collections.Generic;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Teigha.Colors;
using Teigha.DatabaseServices;
using Teigha.Geometry;
#if !SEM_TRANSIENTES
using TransientDrawingMode = Teigha.GraphicsInterface.TransientDrawingMode;
using TransientManager = Teigha.GraphicsInterface.TransientManager;
#endif

namespace OrcamentoNanoCAD
{
    public class Marcas
    {
        private readonly List<Entity> _desenhadas = new List<Entity>();

        // cores ACI: azul (poste), verde (I), vermelho (R), magenta (M), ciano (vínculo), amarelo (conferir), vermelho (erro)
        private const short Azul = 5, Verde = 3, Vermelho = 1, Magenta = 6, Ciano = 4, Amarelo = 2, Cinza = 8;

        private static short CorDaOperacao(string op)
        {
            string o = (op ?? "").Replace("*", "").ToUpperInvariant();
            if (o == "I") return Verde;
            if (o == "R") return Vermelho;
            if (o == "M") return Magenta;
            return Cinza;
        }

        private static short CorDoEstado(Elemento e, short normal)
        {
            if (e.Estado == Estados.Erro) return Vermelho;
            if (e.Estado == Estados.Conferir) return Amarelo;
            return normal;
        }

        // Dois jeitos de desenhar (TASK-014): gráficos transitórios (TransientManager: números escritos, molduras) e, onde a API não tem
        // (nanoCAD 5 / build SEM_TRANSIENTES), vetores temporários do Editor (somem no próximo REGEN; sem os números escritos, só
        // linhas e uma cruz em cada rótulo). A troca é automática na primeira falta (Compat).
        private const string RecursoTransientes = "gráficos transitórios (números na tela)";
        private bool _vetores;

        public void Limpar()
        {
            if (_desenhadas.Count == 0 && !_vetores) return;
            List<Entity> l = new List<Entity>(_desenhadas);
            _desenhadas.Clear();
            if (!_vetores && l.Count > 0 && Compat.Tentar(RecursoTransientes, () => ApagarTransientes(l))) { foreach (Entity e in l) e.Dispose(); return; }
            foreach (Entity e in l) e.Dispose();
            HostMgd.EditorInput.Editor ed = Aplicacao.Ed();
            if (ed != null) Compat.Tentar("REGEN pelo plugin", () => ed.Regen());
        }

        private void Add(Entity e, short cor)
        {
            if (!_vetores && Compat.Tentar(RecursoTransientes, () => AdicionarTransiente(e, cor))) { _desenhadas.Add(e); return; }
            _vetores = true;
            Line l = e as Line;
            HostMgd.EditorInput.Editor ed = Aplicacao.Ed();
            if (ed != null && l != null) Compat.Tentar("vetores temporários", () => ed.DrawVector(l.StartPoint, l.EndPoint, cor, false));
            e.Dispose();
        }

        private void Rotulo(string texto, Point3d onde, double h, short cor)
        {
            if (!_vetores && !Compat.Falta(RecursoTransientes))
            {
                DBText t = new DBText();
                t.TextString = texto;
                t.Height = h;
                t.Position = onde;
                Add(t, cor);
                if (!_vetores)
                {
                    // moldura em volta do número (largura aproximada pela quantidade de letras)
                    double w = h * 0.85 * texto.Length + h * 0.6, m = h * 0.3;
                    Polyline caixa = new Polyline();
                    caixa.AddVertexAt(0, new Point2d(onde.X - m, onde.Y - m), 0, 0, 0);
                    caixa.AddVertexAt(1, new Point2d(onde.X - m + w, onde.Y - m), 0, 0, 0);
                    caixa.AddVertexAt(2, new Point2d(onde.X - m + w, onde.Y + h + m), 0, 0, 0);
                    caixa.AddVertexAt(3, new Point2d(onde.X - m, onde.Y + h + m), 0, 0, 0);
                    caixa.Closed = true;
                    Add(caixa, cor);
                    return;
                }
            }
            Add(new Line(new Point3d(onde.X - h, onde.Y, 0), new Point3d(onde.X + h, onde.Y, 0)), cor);
            Add(new Line(new Point3d(onde.X, onde.Y - h, 0), new Point3d(onde.X, onde.Y + h, 0)), cor);
        }

#if SEM_TRANSIENTES
        private static void AdicionarTransiente(Entity e, short cor) { throw new TypeLoadException("build sem TransientManager"); }
        private static void ApagarTransientes(List<Entity> l) { }
#else
        private static void AdicionarTransiente(Entity e, short cor)
        {
            e.Color = Color.FromColorIndex(ColorMethod.ByAci, cor);
            TransientManager.CurrentTransientManager.AddTransient(e, TransientDrawingMode.Main, 128, new IntegerCollection());
        }

        private static void ApagarTransientes(List<Entity> l)
        {
            TransientManager tm = TransientManager.CurrentTransientManager;
            IntegerCollection vps = new IntegerCollection();
            foreach (Entity e in l)
                try { tm.EraseTransient(e, vps); } catch (System.Exception ex) { if (ex is TypeLoadException || ex is MissingMemberException) throw; }
        }
#endif

        public void ContornoRegiao(IList<Ponto> vertices, double passo)
        {
            // Desenha apenas a última aresta enquanto o CAD solicita os pontos.
            if (vertices.Count < 2) return;
            Ponto a = vertices[vertices.Count - 2], b = vertices[vertices.Count - 1];
            Tracejado(new Point3d(a.X, a.Y, 0), new Point3d(b.X, b.Y, 0), passo, 30);
        }

        public void RegiaoFechada(IList<Ponto> vertices, double passo)
        {
            if (vertices == null || vertices.Count < 3) return;
            for (int i = 0; i < vertices.Count; i++)
            {
                Ponto a = vertices[i], b = vertices[(i + 1) % vertices.Count];
                Tracejado(new Point3d(a.X, a.Y, 0), new Point3d(b.X, b.Y, 0), passo, 30);
            }
        }

        private void Tracejado(Point3d a, Point3d b, double passo, short cor)
        {
            double dist = a.DistanceTo(b);
            if (dist <= 0) return;
            int n = Math.Max(1, (int)(dist / passo));
            if (n > 200) n = 200;
            Vector3d d = (b - a) / (2.0 * n);
            for (int i = 0; i < n; i++)
            {
                Point3d p = a + d * (2 * i);
                Add(new Line(p, p + d), cor);
            }
        }

        private static Point3d P(Ponto p) { return new Point3d(p.X, p.Y, 0); }

        public void Mostrar(Obra obra, RegrasProjeto regras, bool analise = false)
        {
            Limpar();
            double h = Math.Max(obra.AlturaTexto, 1e-3) * 1.1;
            if (analise)
            {
                foreach (Elemento e in Relatorios.DoPainel(obra,null))
                {
                    if (!e.TextoX.HasValue || !e.TextoY.HasValue) continue;
                    string op = (e.Operacao ?? "").Replace("*","");
                    short cor = op == "I" ? Vermelho : op == "R" ? Magenta : op == "M" ? Azul : Cinza;
                    MarcarTexto(new Ponto(e.TextoX.Value,e.TextoY.Value), h, e.Numero + " " + e.Operacao + (e.Operacao.StartsWith("*") ? " LV" : ""),cor);
                }
                return;
            }
            Dictionary<string, int> pilha = new Dictionary<string, int>();      // irmãos (P3b, V7b…) empilhados embaixo do principal
            // só o que o painel mostra: na leitura, o que as regras do leitor classificaram (rede existente sem anotação não ganha número)
            foreach (Poste p in obra.Postes)
            {
                if (!Relatorios.Visivel(obra, p)) continue;
                int k = Pilha(pilha, p);
                Rotulo(p.Numero, new Point3d(p.X + h * 0.8, p.Y + h * 0.4 - k * h * 1.6, 0), h, CorDoEstado(p, Azul));
            }
            foreach (Vao v in obra.Vaos)
            {
                if (v.Pontos.Count < 2 || !Relatorios.Visivel(obra, v)) continue;
                Ponto meio = Geometria.Meio(v.Pontos);
                Point3d pm = P(meio);
                // vínculo vão ⇄ poste: tracejado ciano do meio do vão até cada poste ligado
                foreach (Vinculo vi in v.Vinculos)
                {
                    Poste p = obra.PostePorNumero(vi.Poste);
                    if (p != null) Tracejado(pm, new Point3d(p.X, p.Y, 0), h * 0.8, Ciano);
                }
                int kv = Pilha(pilha, v);
                Rotulo(v.Numero, new Point3d(meio.X - h * 0.9, meio.Y - h * 0.5 - kv * h * 1.5, 0), h * 0.9, CorDoEstado(v, CorDaOperacao(v.Operacao)));
            }
            foreach (Avulso a in obra.Avulsos)
                if (a.TextoX.HasValue && a.Estado != Estados.Ok)
                    Rotulo(a.Numero, new Point3d(a.TextoX.Value, a.TextoY.Value - h * 1.6, 0), h * 0.8, CorDoEstado(a, Cinza));
        }

        private static int Pilha(Dictionary<string, int> pilha, Elemento e)
        {
            string raiz = e.IrmaoDe ?? e.Numero ?? "";
            int k;
            pilha.TryGetValue(raiz, out k);
            pilha[raiz] = k + 1;
            return k;
        }

        /// <summary>Destaque passageiro de um elemento (pop-up / clique na lista): círculo em volta.</summary>
        public void MarcarTexto(Ponto p, double h, string rotulo, short cor)
        {
            Rotulo(rotulo,P(p),Math.Max(0.01,h),cor);
        }

        public void Destacar(Obra obra, Elemento e)
        {
            double h = Math.Max(obra.AlturaTexto, 1e-3);
            Point3d c;
            Poste p = e as Poste;
            Vao v = e as Vao;
            if (p != null) c = new Point3d(p.X, p.Y, 0);
            else if (v != null && v.Pontos.Count > 0) c = P(Geometria.Meio(v.Pontos));
            else if (e.TextoX.HasValue) c = new Point3d(e.TextoX.Value, e.TextoY.Value, 0);
            else return;
            Rotulo(e.Numero,c,h,Amarelo);
            if (e.TextoX.HasValue && e.TextoY.HasValue) MarcarTexto(new Ponto(e.TextoX.Value,e.TextoY.Value),h,"Texto",Amarelo);
            Add(new Circle(c, Vector3d.ZAxis, h * 3), Amarelo);
            if (e.OrigemMetros != null) MarcarTexto(new Ponto(e.OrigemMetros.X,e.OrigemMetros.Y),h,"m",Ciano);
            if (e.GuiaUsada != null)                       // a linha guia que ligou a anotação a este elemento
                for (int i = 0; i + 1 < e.GuiaUsada.Count; i++)
                    Add(new Line(P(e.GuiaUsada[i]), P(e.GuiaUsada[i + 1])), Ciano);
            if (v != null)
                for (int i = 0; i + 1 < v.Pontos.Count; i++)
                    Add(new Line(P(v.Pontos[i]), P(v.Pontos[i + 1])), Amarelo);
        }
    
        // ── 0.16.3: montagem no formato do mapa — o que já foi coletado aparece em outra cor (some ao fechar o painel) ──
        private const short Laranja = 30;

        private void Retangulo(double[] c, double folga, short cor)
        {
            if (c == null || c.Length != 4) return;
            double x0 = c[0] - folga, y0 = c[1] - folga, x1 = c[2] + folga, y1 = c[3] + folga;
            Add(new Line(new Point3d(x0, y0, 0), new Point3d(x1, y0, 0)), cor);
            Add(new Line(new Point3d(x1, y0, 0), new Point3d(x1, y1, 0)), cor);
            Add(new Line(new Point3d(x1, y1, 0), new Point3d(x0, y1, 0)), cor);
            Add(new Line(new Point3d(x0, y1, 0), new Point3d(x0, y0, 0)), cor);
        }

        private void Circulo(double x, double y, double r, short cor)
        {
            const int n = 20;                                  // em linhas: também funciona com os vetores temporários do nanoCAD 5
            for (int i = 0; i < n; i++)
            {
                double a = 2 * Math.PI * i / n, b = 2 * Math.PI * (i + 1) / n;
                Add(new Line(new Point3d(x + r * Math.Cos(a), y + r * Math.Sin(a), 0), new Point3d(x + r * Math.Cos(b), y + r * Math.Sin(b), 0)), cor);
            }
        }

        private void Caminho(IList<Ponto> pts, short cor)
        {
            for (int i = 0; pts != null && i + 1 < pts.Count; i++) Add(new Line(P(pts[i]), P(pts[i + 1])), cor);
        }

        /// <summary>Pontos (ciano) e cabos (laranja) já coletados no modo mapa, com as anotações usadas emolduradas; enviados ao OrçaCore em cinza.</summary>
        public void MostrarMapa(Obra obra)
        {
            Limpar();
            if (obra == null) return;
            double h = Math.Max(obra.AlturaTexto, 1e-3);
            foreach (PontoMapa p in obra.PontosMapa)
            {
                short cor = p.Enviado ? Cinza : Ciano;
                Circulo(p.X, p.Y, h * 1.6, cor);
                foreach (double[] c in p.Caixas) Retangulo(c, h * 0.25, cor);
                Rotulo(p.NomeFinal, new Point3d(p.X + h * 1.9, p.Y + h * 0.5, 0), h * 0.9, cor);
            }
            foreach (CaboMapa c in obra.CabosMapa)
            {
                short cor = c.Enviado ? Cinza : Laranja;
                Caminho(c.Pontos, cor);
                foreach (double[] cx in c.Caixas) Retangulo(cx, h * 0.25, cor);
                if (c.Pontos.Count >= 2) { Ponto m = Geometria.Meio(c.Pontos); Rotulo(c.Numero, new Point3d(m.X + h * 0.5, m.Y + h * 0.5, 0), h * 0.9, cor); }
            }
        }

        /// <summary>Destaque (amarelo) do item escolhido no painel, junto com o objeto base: o poste (círculo) ou a linha/trecho do cabo.</summary>
        public void DestacarMapa(Obra obra, string numero)
        {
            if (obra == null || string.IsNullOrEmpty(numero)) return;
            double h = Math.Max(obra.AlturaTexto, 1e-3);
            PontoMapa p = obra.PontosMapa.FirstOrDefault(x => x.Numero == numero);
            if (p != null)
            {
                Circulo(p.X, p.Y, h * 2.6, Amarelo);
                Circulo(p.X, p.Y, h * 3.2, Amarelo);
                foreach (double[] c in p.Caixas) { Retangulo(c, h * 0.5, Amarelo); Add(new Line(new Point3d(p.X, p.Y, 0), new Point3d((c[0] + c[2]) / 2, (c[1] + c[3]) / 2, 0)), Amarelo); }
                Rotulo(p.NomeFinal, new Point3d(p.X + h * 3.4, p.Y - h * 1.8, 0), h, Amarelo);
                return;
            }
            CaboMapa cabo = obra.CabosMapa.FirstOrDefault(x => x.Numero == numero);
            if (cabo == null) return;
            Caminho(cabo.Pontos, Amarelo);
            if (cabo.Pontos.Count >= 2)
            {
                Circulo(cabo.Pontos[0].X, cabo.Pontos[0].Y, h * 0.8, Amarelo);
                Circulo(cabo.Pontos[cabo.Pontos.Count - 1].X, cabo.Pontos[cabo.Pontos.Count - 1].Y, h * 0.8, Amarelo);
            }
            bool temMeio = cabo.Pontos.Count >= 2;
            Ponto meio = temMeio ? Geometria.Meio(cabo.Pontos) : new Ponto(0, 0);
            foreach (double[] c in cabo.Caixas)
            {
                Retangulo(c, h * 0.5, Amarelo);
                if (temMeio) Add(new Line(P(meio), new Point3d((c[0] + c[2]) / 2, (c[1] + c[3]) / 2, 0)), Amarelo);
            }
            if (temMeio) Rotulo(cabo.Numero, new Point3d(meio.X + h, meio.Y - h * 1.8, 0), h, Amarelo);
        }
}
}
