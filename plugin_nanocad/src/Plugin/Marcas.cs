// Plugin/Marcas.cs — números (P1…, V1…) e linhas de vínculo vão ⇄ poste desenhados como GRÁFICOS TRANSITÓRIOS: aparecem na tela,
// não entram no DWG (sem camada de marcas). Somem ao fechar o desenho; o comando ORC_MARCAS liga/desliga.
using System;
using System.Collections.Generic;
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

        public void Mostrar(Obra obra, RegrasProjeto regras)
        {
            Limpar();
            double h = Math.Max(obra.AlturaTexto, 1e-3) * 1.1;
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
            Add(new Circle(c, Vector3d.ZAxis, h * 3), Amarelo);
            if (e.GuiaUsada != null)                       // a linha guia que ligou a anotação a este elemento
                for (int i = 0; i + 1 < e.GuiaUsada.Count; i++)
                    Add(new Line(P(e.GuiaUsada[i]), P(e.GuiaUsada[i + 1])), Ciano);
            if (v != null)
                for (int i = 0; i + 1 < v.Pontos.Count; i++)
                    Add(new Line(P(v.Pontos[i]), P(v.Pontos[i + 1])), Amarelo);
        }
    }
}
