// Nucleo/DesenhoBruto.cs — o DWG "cru", do jeito que o CAD entrega (TASK-004/005/006 do plugin):
//   entidades com o texto ORIGINAL (códigos de MTEXT inclusos), a cor SEM resolver (ByLayer/ByBlock/ACI/RGB), blocos com atributos e
//   textos da definição, linhas, LEADER, MLEADER (texto + linhas de chamada) e o contorno de cada texto.
// A camada do CAD só preenche isto (ExtratorDesenho). Tudo o mais — cor final, parágrafos do MTEXT, postes, redes, guias — é feito
// AQUI (MontadorDesenho), igual para o nanoCAD, para os testes (DXF convertido por tests/oraculo/dxf_para_bruto.py) e para o
// diagnóstico (ORC_DIAGNOSTICO grava este mesmo formato, que vira teste quando o usuário manda).
using System;
using System.Collections.Generic;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    public class CorBruta
    {
        public string Metodo = "layer";   // layer | block | aci | rgb
        public int Aci;
        public int Rgb = -1;              // 0xRRGGBB

        public JObj ParaJson()
        {
            JObj o = new JObj().Set("m", Metodo);
            if (Metodo == "aci") o.Set("aci", Aci);
            if (Metodo == "rgb") o.Set("rgb", Rgb);
            return o;
        }

        public static CorBruta De(JObj o)
        {
            CorBruta c = new CorBruta();
            if (o == null) return c;
            c.Metodo = o.Str("m") ?? "layer";
            c.Aci = (int)(o.Num("aci") ?? 0);
            c.Rgb = (int)(o.Num("rgb") ?? -1);
            return c;
        }

        public static CorBruta DoAci(int aci)
        {
            CorBruta c = new CorBruta();
            if (aci == 256) c.Metodo = "layer";
            else if (aci == 0) c.Metodo = "block";
            else { c.Metodo = "aci"; c.Aci = aci; }
            return c;
        }
    }

    public class EntidadeBruta
    {
        public string Tipo = "";          // TEXT, MTEXT, ATTRIB, INSERT, LINE, LWPOLYLINE, POLYLINE, LEADER, MLEADER, CIRCLE
        public string Handle = "";
        public string Layer = "";
        public CorBruta Cor = new CorBruta();
        public string Texto;              // TEXT/ATTRIB: texto; MTEXT/MLEADER: conteúdo cru (com códigos)
        public double X, Y, Altura, Raio;
        public double[] Caixa;            // [minX, minY, maxX, maxY] do texto, quando o CAD informa
        public string Bloco;              // INSERT: nome efetivo do bloco
        public bool Invisivel;
        public List<Ponto> Pontos = new List<Ponto>();                   // linhas
        public List<List<Ponto>> Linhas = new List<List<Ponto>>();       // MLEADER: linhas de chamada
        public List<EntidadeBruta> Filhos = new List<EntidadeBruta>();   // INSERT: atributos da instância
        public List<EntidadeBruta> Definicao = new List<EntidadeBruta>(); // INSERT: TEXT/MTEXT/ATTDEF da definição (coordenadas do bloco, como a aplicação)

        private static List<object> Pts(List<Ponto> l) { return l.Select(p => (object)new List<object> { Math.Round(p.X, 4), Math.Round(p.Y, 4) }).ToList(); }

        private static List<Ponto> DePts(object v)
        {
            List<Ponto> r = new List<Ponto>();
            List<object> l = v as List<object>;
            if (l == null) return r;
            foreach (object x in l)
            {
                List<object> xy = x as List<object>;
                if (xy != null && xy.Count >= 2) r.Add(new Ponto(Json.ComoNumero(xy[0]) ?? 0, Json.ComoNumero(xy[1]) ?? 0));
            }
            return r;
        }

        public JObj ParaJson()
        {
            JObj o = new JObj().Set("t", Tipo).Set("h", Handle).Set("layer", Layer).Set("cor", Cor.ParaJson());
            if (Texto != null) o.Set("texto", Texto);
            if (X != 0 || Y != 0) { o.Set("x", Math.Round(X, 4)); o.Set("y", Math.Round(Y, 4)); }
            if (Altura > 0) o.Set("altura", Math.Round(Altura, 4));
            if (Raio > 0) o.Set("raio", Math.Round(Raio, 4));
            if (Caixa != null) o.Set("caixa", Caixa.Select(v => (object)Math.Round(v, 4)).ToList());
            if (Bloco != null) o.Set("bloco", Bloco);
            if (Invisivel) o.Set("invisivel", true);
            if (Pontos.Count > 0) o.Set("pontos", Pts(Pontos));
            if (Linhas.Count > 0) o.Set("linhas", Linhas.Select(l => (object)Pts(l)).ToList());
            if (Filhos.Count > 0) o.Set("atributos", Filhos.Select(f => (object)f.ParaJson()).ToList());
            if (Definicao.Count > 0) o.Set("definicao", Definicao.Select(f => (object)f.ParaJson()).ToList());
            return o;
        }

        public static EntidadeBruta De(JObj o)
        {
            EntidadeBruta e = new EntidadeBruta();
            e.Tipo = o.Str("t") ?? "";
            e.Handle = o.Str("h") ?? "";
            e.Layer = o.Str("layer") ?? "";
            e.Cor = CorBruta.De(o.Obj("cor"));
            e.Texto = o.Str("texto");
            e.X = o.Num("x") ?? 0; e.Y = o.Num("y") ?? 0;
            e.Altura = o.Num("altura") ?? 0;
            e.Raio = o.Num("raio") ?? 0;
            List<object> cx = o.Lista("caixa");
            if (cx != null && cx.Count == 4) e.Caixa = cx.Select(v => Json.ComoNumero(v) ?? 0).ToArray();
            e.Bloco = o.Str("bloco");
            e.Invisivel = o.Bool("invisivel", false);
            e.Pontos = DePts(o["pontos"]);
            foreach (object l in o.Lista("linhas") ?? new List<object>()) e.Linhas.Add(DePts(l));
            foreach (JObj f in (o.Lista("atributos") ?? new List<object>()).OfType<JObj>()) e.Filhos.Add(De(f));
            foreach (JObj f in (o.Lista("definicao") ?? new List<object>()).OfType<JObj>()) e.Definicao.Add(De(f));
            return e;
        }
    }

    public class DesenhoBruto
    {
        public string Arquivo = "";
        public Dictionary<string, CorBruta> Layers = new Dictionary<string, CorBruta>(StringComparer.OrdinalIgnoreCase);
        public List<EntidadeBruta> Entidades = new List<EntidadeBruta>();
        public List<string> Avisos = new List<string>();   // entidades que o extrator não conseguiu ler

        public JObj ParaJson()
        {
            JObj layers = new JObj();
            foreach (KeyValuePair<string, CorBruta> kv in Layers.OrderBy(k => k.Key, StringComparer.OrdinalIgnoreCase)) layers.Set(kv.Key, kv.Value.ParaJson());
            return new JObj().Set("formato", "plugin-desenho-bruto").Set("versao", 1).Set("arquivo", Arquivo).Set("layers", layers)
                .Set("entidades", Entidades.Select(e => (object)e.ParaJson()).ToList()).Set("avisos", Avisos.Cast<object>().ToList());
        }

        public static DesenhoBruto De(JObj o)
        {
            if (o != null && o.Obj("bruto") != null) o = o.Obj("bruto");          // aceita o arquivo de diagnóstico inteiro
            DesenhoBruto d = new DesenhoBruto();
            if (o == null) return d;
            d.Arquivo = o.Str("arquivo") ?? "";
            JObj layers = o.Obj("layers");
            if (layers != null) foreach (KeyValuePair<string, object> kv in layers) d.Layers[kv.Key] = CorBruta.De(kv.Value as JObj);
            foreach (JObj e in (o.Lista("entidades") ?? new List<object>()).OfType<JObj>()) d.Entidades.Add(EntidadeBruta.De(e));
            d.Avisos = Json.ListaDeTextos(o["avisos"]);
            return d;
        }
    }

    /// <summary>DesenhoBruto → Desenho (o que o leitor usa). Mesmos critérios do leitor de DXF da aplicação (services/dxf_service.py).</summary>
    public static class MontadorDesenho
    {
        /// <summary>true color → ACI → layer TXT_RRGGBB → cor da layer → preto; ByBlock dentro de bloco = cor do bloco.</summary>
        public static string Cor(CorBruta c, string layer, DesenhoBruto d, string corDoBloco)
        {
            if (c.Metodo == "rgb" && c.Rgb >= 0) return Cores.Hex((c.Rgb >> 16) & 0xFF, (c.Rgb >> 8) & 0xFF, c.Rgb & 0xFF);
            if (c.Metodo == "aci" && c.Aci != 0 && c.Aci != 256) return Cores.AciParaHex(c.Aci);
            if (c.Metodo == "block" && corDoBloco != null) return corDoBloco;
            string h = Cores.CorDoNomeDaLayer(layer);
            if (h != null) return h;
            CorBruta lc;
            if (d.Layers.TryGetValue(layer ?? "", out lc))
            {
                if (lc.Metodo == "rgb" && lc.Rgb >= 0) return Cores.Hex((lc.Rgb >> 16) & 0xFF, (lc.Rgb >> 8) & 0xFF, lc.Rgb & 0xFF);
                return Cores.AciParaHex(Math.Abs(lc.Aci));     // layer desligada tem ACI negativo
            }
            return "#000000";
        }

        private static void Textos(EntidadeBruta e, DesenhoBruto d, string corDoBloco, List<TextoDesenho> saida, string tipo)
        {
            string corBase = Cor(e.Cor, e.Layer, d, corDoBloco);
            List<ParteTexto> partes;
            if (e.Tipo == "MTEXT" || e.Tipo == "MLEADER") partes = Cores.DividirMText(e.Texto, corBase);
            else
            {
                partes = new List<ParteTexto>();
                string t = Cores.LimparTextoSimples(e.Texto);
                if (t.Length > 0) partes.Add(new ParteTexto { Texto = t, Cor = corBase });
            }
            int k = 0;
            foreach (ParteTexto p in partes)
            {
                TextoDesenho t = new TextoDesenho();
                t.Handle = e.Handle; t.Parte = k++; t.Texto = p.Texto; t.Cor = p.Cor; t.Layer = e.Layer;
                t.X = e.X; t.Y = e.Y; t.Altura = e.Altura; t.Tipo = tipo; t.Caixa = e.Caixa;
                foreach (List<Ponto> l in e.Linhas) if (l.Count >= 2) t.Guias.Add(l);
                saida.Add(t);
            }
        }

        public static Desenho Montar(DesenhoBruto b, ConfigLeitura cfg)
        {
            Desenho d = new Desenho();
            d.Arquivo = b.Arquivo;
            foreach (EntidadeBruta e in b.Entidades)
            {
                switch (e.Tipo)
                {
                    case "TEXT":
                    case "MTEXT":
                    case "ATTRIB":
                        if (!e.Invisivel) Textos(e, b, null, d.Textos, e.Tipo);
                        break;
                    case "MLEADER":
                        // texto da linha de chamada múltipla (a aplicação não lê MLEADER do DXF; aqui ele vale como texto, com as linhas como guias)
                        if (!string.IsNullOrEmpty(e.Texto)) Textos(e, b, null, d.Textos, "MLEADER");
                        foreach (List<Ponto> l in e.Linhas) if (l.Count >= 2) d.Linhas.Add(new LinhaDesenho { Handle = e.Handle, Layer = e.Layer, Cor = Cor(e.Cor, e.Layer, b, null), Pontos = l, Leader = true });
                        break;
                    case "INSERT":
                        {
                            string corBloco = Cor(e.Cor, e.Layer, b, null);
                            if (cfg.EhBlocoPoste(e.Bloco, e.Layer)) d.Postes.Add(new PosteDesenho { Handle = e.Handle, X = e.X, Y = e.Y, Bloco = e.Bloco ?? "", Layer = e.Layer });
                            d.Textos.Add(new TextoDesenho { Handle = e.Handle, Texto = "[BLOCO: " + e.Bloco + "]", Cor = corBloco, Layer = e.Layer, X = e.X, Y = e.Y, Tipo = "BLOCO" });
                            foreach (EntidadeBruta def in e.Definicao) Textos(def, b, corBloco, d.Textos, "BLOCO_DEF");
                            foreach (EntidadeBruta a in e.Filhos) if (!a.Invisivel) Textos(a, b, corBloco, d.Textos, "ATTRIB");
                            break;
                        }
                    case "CIRCLE":
                        if (cfg.EhCirculoPoste(e.Layer)) d.Postes.Add(new PosteDesenho { Handle = e.Handle, X = e.X, Y = e.Y, Bloco = "(círculo)", Layer = e.Layer });
                        break;
                    default:
                        if (e.Pontos.Count >= 2)
                            d.Linhas.Add(new LinhaDesenho { Handle = e.Handle, Layer = e.Layer, Cor = Cor(e.Cor, e.Layer, b, null), Pontos = e.Pontos, Leader = e.Tipo == "LEADER" });
                        break;
                }
            }
            return d;
        }
    }
}
