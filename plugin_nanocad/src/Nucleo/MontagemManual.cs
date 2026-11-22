using System;
using System.Collections.Generic;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    public static class MontagemManual
    {
        public static readonly string[] Operacoes = { "I", "R", "M", "*I", "*R", "*M" };

        public static bool Dentro(Ponto p, IList<Ponto> poligono)
        {
            bool dentro = false;
            for (int i = 0, j = poligono.Count - 1; i < poligono.Count; j = i++)
            {
                Ponto a = poligono[j], b = poligono[i];
                if (Geometria.ProjetarSegmento(p, a, b).Distancia <= 1e-8) return true;
                if ((a.Y > p.Y) != (b.Y > p.Y) && p.X < (b.X - a.X) * (p.Y - a.Y) / (b.Y - a.Y) + a.X)
                    dentro = !dentro;
            }
            return dentro;
        }

        private static double Lado(Ponto a, Ponto b, Ponto p)
        {
            return (b.X - a.X) * (p.Y - a.Y) - (b.Y - a.Y) * (p.X - a.X);
        }

        public static bool Cruza(IList<Ponto> linha, IList<Ponto> poligono)
        {
            foreach (Ponto p in linha) if (Dentro(p, poligono)) return true;
            for (int i = 1; i < linha.Count; i++)
            {
                Ponto a = linha[i - 1], b = linha[i];
                for (int j = 0; j < poligono.Count; j++)
                {
                    Ponto c = poligono[j], d = poligono[(j + 1) % poligono.Count];
                    if (Math.Max(a.X, b.X) < Math.Min(c.X, d.X) || Math.Max(c.X, d.X) < Math.Min(a.X, b.X)
                        || Math.Max(a.Y, b.Y) < Math.Min(c.Y, d.Y) || Math.Max(c.Y, d.Y) < Math.Min(a.Y, b.Y))
                        continue;
                    if (Math.Sign(Lado(a, b, c)) * Math.Sign(Lado(a, b, d)) <= 0 &&
                        Math.Sign(Lado(c, d, a)) * Math.Sign(Lado(c, d, b)) <= 0) return true;
                }
            }
            return false;
        }

        public static void Recortar(Desenho desenho, IList<Ponto> poligono, ConfigLeitura cfg = null)
        {
            desenho.Textos.RemoveAll(t => !Dentro(new Ponto(t.X, t.Y), poligono));
            desenho.Linhas.RemoveAll(l => !Cruza(l.Pontos, poligono));
            // Os postes das extremidades de uma linha que cruza a seleção permanecem como topologia.
            var pontas = desenho.Linhas.Where(l => !l.Leader && l.Pontos.Count >= 2)
                .SelectMany(l => new[] { l.Pontos[0], l.Pontos[l.Pontos.Count - 1] }).ToList();
            double tol = (cfg == null ? 1 : cfg.TolPosteAlturas) * LeitorProjeto.AlturaTipica(desenho.Textos);
            desenho.Postes.RemoveAll(p =>
            {
                Ponto ponto = new Ponto(p.X, p.Y);
                return !Dentro(ponto, poligono) && !pontas.Any(q => q.Dist(ponto) <= tol);
            });
        }

        private static string Fonte(Elemento e)
        {
            return !string.IsNullOrEmpty(e.TextoHandle) ? e.TextoHandle + ":" + e.ParteTexto :
                (e.TextoOriginal ?? "") + "|" + new Ponto(e.TextoX ?? 0, e.TextoY ?? 0) + "|" + e.TextoLayer + "|" + e.TextoCor;
        }

        private static bool MesmoGrupo(Elemento a, Elemento b)
        {
            return string.Equals(a.LayerAtual, b.LayerAtual, StringComparison.OrdinalIgnoreCase)
                && string.Equals(a.TextoCor, b.TextoCor, StringComparison.OrdinalIgnoreCase) && a.Operacao == b.Operacao;
        }

        private static string NumeroIrmao(string raiz, IEnumerable<Elemento> lista)
        {
            var numeros = new HashSet<string>(lista.Select(e => e.Numero));
            for (int i = 1; ; i++)
            {
                string candidato = raiz + (i < 26 ? ((char)('a' + i)).ToString() : "b" + i);
                if (!numeros.Contains(candidato)) return candidato;
            }
        }

        private static void AcrescentarAoPoste(Obra obra, Poste raiz, Poste anotacao)
        {
            if (!anotacao.TemAnotacao) return;
            List<Poste> grupo = obra.GrupoDoPoste(raiz);
            string fonte = Fonte(anotacao);
            if (grupo.Any(p => p.TemAnotacao && (Fonte(p) == fonte || p.FontesMontagem.Contains(fonte)))) return;
            Poste destino = grupo.FirstOrDefault(p => p.TemAnotacao && !p.Excluido && !p.Bloqueada && !p.Comentario && !p.EnviadaNaMontagem && MesmoGrupo(p, anotacao));
            if (destino != null)
            {
                destino.FontesMontagem.Add(fonte);
                destino.TextoOriginal += "\n" + anotacao.TextoOriginal;
                // Uma anotação repetida em outra entidade continua representada na mesma linha.
                Construtor.AplicarAtivo(destino, destino.AtivoLido + " " + anotacao.AtivoLido);
                return;
            }
            string numero = raiz.TemAnotacao ? NumeroIrmao(raiz.Numero, grupo.Cast<Elemento>()) : raiz.Numero;
            anotacao.Numero = numero;
            anotacao.Handle = raiz.Handle;
            anotacao.X = raiz.X;
            anotacao.Y = raiz.Y;
            anotacao.Bloco = raiz.Bloco;
            anotacao.Placa = raiz.Placa;
            anotacao.OutrosHandles = raiz.OutrosHandles;
            anotacao.Manual = true;
            anotacao.FontesMontagem.Add(fonte);
            if (raiz.TemAnotacao)
            {
                anotacao.IrmaoDe = raiz.Numero;
                obra.Postes.Add(anotacao);
            }
            else
            {
                anotacao.IrmaoDe = null;
                obra.Postes[obra.Postes.IndexOf(raiz)] = anotacao;
            }
        }

        private static Poste ObterRaiz(Obra obra, Poste origem)
        {
            Poste raiz = obra.Postes.FirstOrDefault(p => p.IrmaoDe == null &&
                (!string.IsNullOrEmpty(origem.Handle) ? p.Handle == origem.Handle || p.OutrosHandles.Contains(origem.Handle) :
                 string.IsNullOrEmpty(p.Handle) && p.Ponto.Dist(origem.Ponto) < 1e-6));
            if (raiz != null) return raiz;
            raiz = new Poste { Numero = "P" + Construtor.Proximo(obra.Postes.Cast<Elemento>(), 'P'),
                Handle = origem.Handle, X = origem.X, Y = origem.Y, Bloco = origem.Bloco, Placa = origem.Placa, Manual = true,
                OutrosHandles = new List<string>(origem.OutrosHandles) };
            obra.Postes.Add(raiz);
            return raiz;
        }

        public static Poste AdicionarPoste(Obra obra, PosteDesenho bloco, List<TextoDesenho> textos, RegrasProjeto regras)
        {
            TextoDesenho primeiro = textos.FirstOrDefault();
            if (bloco == null && primeiro == null) return null;
            Poste origem = bloco != null ? new Poste { Handle = bloco.Handle, X = bloco.X, Y = bloco.Y, Bloco = bloco.Bloco } :
                new Poste { X = primeiro.X, Y = primeiro.Y };
            string numero = ObterRaiz(obra, origem).Numero;
            foreach (TextoLido tl in LeitorProjeto.ClassificarTextos(textos, regras))
            {
                if (string.IsNullOrEmpty((tl.R.Ativo ?? "").Trim()) || tl.R.Entidade == "CABO") continue;
                if (tl.R.Entidade == "" || tl.R.Entidade == "0") tl.R.Entidade = "ESTRUTURA";
                if (!Operacoes.Contains(tl.R.Operacao)) tl.R.Operacao = "I";
                Poste anotacao = new Poste();
                LeitorProjeto.Anotar(anotacao, tl);
                AcrescentarAoPoste(obra, obra.PostePorNumero(numero), anotacao);
            }
            Construtor.Recalcular(obra, regras);
            return obra.PostePorNumero(numero);
        }

        public static List<TextoLido> CabosSelecionados(List<TextoDesenho> textos, RegrasProjeto regras,
            Func<IEnumerable<TextoDesenho>> desenho)
        {
            var lidos = LeitorProjeto.ClassificarTextos(textos, regras);
            var cabos = new List<TextoLido>();
            var medidas = textos.Where(t => Construtor.EhComprimento(t.Texto)).ToList();
            List<TextoDesenho> vizinhos = null;
            var medidasUsadas = new HashSet<TextoDesenho>(lidos.Where(t => t.Comprimento != null).Select(t => t.Comprimento));
            medidas.RemoveAll(medidasUsadas.Contains);
            foreach (TextoLido tl in lidos.Where(t => !Construtor.EhComprimento(t.T.Texto)))
            {
                // O leitor pode ter consumido um texto como parte de outra linha.
                if (string.IsNullOrEmpty((tl.R.Ativo ?? "").Trim())) continue;
                if (!Construtor.TemComprimento(tl.R.Ativo))
                {
                    TextoDesenho medida = medidas.OrderBy(t => t.Centro.Dist(tl.T.Centro)).FirstOrDefault();
                    if (medida == null && desenho != null)
                    {
                        if (vizinhos == null) vizinhos = desenho().ToList();
                        medida = Construtor.ComprimentoVizinho(tl.T, textos, vizinhos);
                        if (medida != null) vizinhos.Remove(medida);
                    }
                    if (medida != null)
                    {
                        medidas.Remove(medida);
                        tl.Comprimento = medida;
                        tl.R.Ativo += " " + medida.Texto.Trim();
                    }
                }
                tl.R.Entidade = "CABO";
                if (!Operacoes.Contains(tl.R.Operacao)) tl.R.Operacao = "I";
                cabos.Add(tl);
            }
            return cabos;
        }

        public static Vao AdicionarVao(Obra obra, LinhaDesenho linha, List<TextoDesenho> textos, Poste a, Poste b,
            RegrasProjeto regras, Func<IEnumerable<TextoDesenho>> desenho)
        {
            var lidos = CabosSelecionados(textos, regras, desenho);
            Vao raiz = obra.Vaos.FirstOrDefault(v => v.IrmaoDe == null && v.LinhaHandle == linha.Handle);
            if (raiz == null) raiz = Construtor.AdicionarVao(obra, linha, new List<TextoDesenho>(), a, b, regras);
            if (raiz.Bloqueada) return raiz;
            raiz.PosteA = a == null ? null : a.Numero;
            raiz.PosteB = b == null ? null : b.Numero;
            foreach (TextoLido tl in lidos)
            {
                Vao candidato = new Vao();
                LeitorProjeto.Anotar(candidato, tl);
                if (obra.Vaos.Any(v => v.LinhaHandle == linha.Handle && v.TemAnotacao && Fonte(v) == Fonte(candidato))) continue;
                if (!raiz.TemAnotacao)
                {
                    LeitorProjeto.Anotar(raiz, tl);
                    continue;
                }
                candidato.Numero = NumeroIrmao(raiz.Numero, obra.Vaos.Cast<Elemento>());
                candidato.IrmaoDe = raiz.Numero;
                candidato.LinhaHandle = raiz.LinhaHandle;
                candidato.Pontos = raiz.Pontos;
                candidato.LinhaLayer = raiz.LinhaLayer;
                candidato.LinhaCor = raiz.LinhaCor;
                candidato.Tipo = raiz.Tipo;
                candidato.PosteA = raiz.PosteA;
                candidato.PosteB = raiz.PosteB;
                candidato.Manual = true;
                obra.Vaos.Add(candidato);
            }
            Construtor.Recalcular(obra, regras);
            return raiz;
        }

        public static void Incorporar(Obra obra, Obra regiao, RegrasProjeto regras)
        {
            if (!obra.Elementos.Any()) obra.AlturaTexto = regiao.AlturaTexto;
            var postes = new Dictionary<string, string>();
            foreach (Poste p in regiao.Postes)
            {
                string anterior = p.Numero;
                Poste raiz = p.IrmaoDe == null ? ObterRaiz(obra, p) : obra.PostePorNumero(postes[p.IrmaoDe]);
                postes[anterior] = raiz.Numero;
                AcrescentarAoPoste(obra, raiz, p);
            }
            var linhas = obra.Vaos.ToLookup(v => v.LinhaHandle ?? "");
            var vaos = new Dictionary<string, string>();
            foreach (Vao v in regiao.Vaos)
            {
                string anterior = v.Numero;
                Vao existente = linhas[v.LinhaHandle ?? ""].FirstOrDefault(a => Fonte(a) == Fonte(v) ||
                    (!a.TemAnotacao && !v.TemAnotacao));
                if (existente != null) { vaos[anterior] = existente.Numero; continue; }
                if (!v.TemAnotacao && linhas[v.LinhaHandle ?? ""].Any())
                {
                    vaos[anterior] = linhas[v.LinhaHandle ?? ""].First().Numero;
                    continue;
                }
                v.PosteA = v.PosteA == null ? null : postes[v.PosteA];
                v.PosteB = v.PosteB == null ? null : postes[v.PosteB];
                v.IrmaoDe = v.IrmaoDe == null ? null : vaos[v.IrmaoDe];
                Vao vazio = v.IrmaoDe == null && v.TemAnotacao ?
                    linhas[v.LinhaHandle ?? ""].FirstOrDefault(a => !a.TemAnotacao && a.IrmaoDe == null) : null;
                v.Numero = vazio != null ? vazio.Numero : v.IrmaoDe != null ?
                    NumeroIrmao(v.IrmaoDe, obra.Vaos.Cast<Elemento>()) : "V" + Construtor.Proximo(obra.Vaos.Cast<Elemento>(), 'V');
                vaos[anterior] = v.Numero;
                v.Manual = true;
                v.VinculoA = v.VinculoB = null;
                if (vazio == null) obra.Vaos.Add(v);
                else obra.Vaos[obra.Vaos.IndexOf(vazio)] = v;
            }
            var fontes = new HashSet<string>(obra.Avulsos.Select(Fonte));
            foreach (Avulso a in regiao.Avulsos)
            {
                if (!fontes.Add(Fonte(a))) continue;
                a.Numero = "L" + Construtor.Proximo(obra.Avulsos.Cast<Elemento>(), 'L');
                a.Manual = true;
                obra.Avulsos.Add(a);
            }
            obra.Ignorados.AddRange(regiao.Ignorados.Where(t => !obra.Ignorados.Any(a =>
                a.Handle == t.Handle && a.Parte == t.Parte && a.Texto == t.Texto && a.X == t.X && a.Y == t.Y)));
            Construtor.Recalcular(obra, regras);
        }

        public static void Editar(Obra obra, Elemento linha, string ativo, string operacao,
            bool comentario, bool congelada, bool bloqueada, RegrasProjeto regras)
        {
            if (!Operacoes.Contains(operacao) || string.IsNullOrEmpty((ativo ?? "").Trim())) return;
            if (!linha.Bloqueada)
            {
                if (ativo != linha.AtivoLido)
                {
                    if (!linha.AtivoEditado) linha.AtivoDoDesenho = linha.AtivoLido;
                    Construtor.AplicarAtivo(linha, ativo);
                }
                linha.Operacao = operacao;
                linha.Comentario = comentario;
                linha.Congelada = congelada || comentario;
            }
            linha.Bloqueada = bloqueada;
            Construtor.Recalcular(obra, regras);
        }
    }
}
