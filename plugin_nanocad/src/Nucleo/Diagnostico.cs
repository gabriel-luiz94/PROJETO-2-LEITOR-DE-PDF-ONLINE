// Nucleo/Diagnostico.cs — ORC_DIAGNOSTICO (TASK-004 do plugin): o que o plugin enxergou no desenho e o que fez com cada texto.
// O arquivo "<desenho>.diagnostico.json" leva o desenho BRUTO inteiro (vira teste em tests/Nucleo.Testes quando o usuário manda),
// um resumo por layer/bloco/tipo e, para cada texto, o resultado das regras do leitor e o destino (poste, vão, avulso ou ignorado + motivo).
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    public static class Diagnostico
    {
        private static void Somar(JObj o, string chave, int n)
        {
            o.Set(chave, (Json.ComoNumero(o[chave]) ?? 0) + n);
        }

        private static JObj Camada(JObj layers, string nome)
        {
            JObj l = layers.Obj(nome ?? "");
            if (l == null) { l = new JObj(); layers.Set(nome ?? "", l); }
            return l;
        }

        public static JObj Resumo(DesenhoBruto b, Desenho d, Obra obra, RegrasProjeto regras)
        {
            ConfigLeitura cfg = regras.Leitura ?? new ConfigLeitura();
            JObj tipos = new JObj(), layers = new JObj(), blocos = new JObj();
            foreach (EntidadeBruta e in b.Entidades)
            {
                Somar(tipos, e.Tipo, 1);
                Somar(Camada(layers, e.Layer), e.Tipo, 1);
                if (e.Tipo == "INSERT") Somar(blocos, e.Bloco ?? "", 1);
            }
            double h = LeitorProjeto.AlturaTipica(d.Textos);
            int mt = 0, bt = 0, guias = 0;
            foreach (LinhaDesenho l in d.Linhas)
            {
                string t = l.Leader ? null : cfg.ClassificarRede(l.Layer, l.Cor);
                if (t == "MT") mt++;
                else if (t == "BT") bt++;
                else if (l.Leader || (cfg.PodeSerGuia(l.Layer) && Geometria.Comprimento(l.Pontos) <= cfg.GuiaMaxAlturas * h)) guias++;
            }
            List<string> avisos = new List<string>();
            if (mt + bt == 0) avisos.Add("nenhuma linha de rede achada: confira as layers da rede em Configurar leitura (hoje: "
                                         + string.Join(", ", cfg.LayersRedeMT.Concat(cfg.LayersRedeBT).ToArray()) + ")");
            if (d.Postes.Count == 0) avisos.Add("nenhum poste achado: confira o nome do bloco de poste em Configurar leitura (hoje: "
                                                + string.Join(", ", cfg.BlocosPoste.ToArray()) + ") — blocos do desenho: "
                                                + string.Join(", ", blocos.Keys.Take(15).ToArray()));
            if (guias == 0) avisos.Add("nenhuma linha guia achada (LEADER, MLEADER ou linha curta fora da rede)");
            int lidos = obra.Elementos.Count(e => e.TemAnotacao);
            if (obra.Ignorados.Count > 0 && lidos == 0) avisos.Add("todas as anotações foram ignoradas pelas regras do leitor: confira se as regras escolhidas são as do projeto");
            return new JObj()
                .Set("entidades_por_tipo", tipos).Set("blocos", blocos).Set("layers", layers)
                .Set("altura_tipica", Math.Round(h, 3))
                .Set("redes_mt", mt).Set("redes_bt", bt).Set("guias", guias).Set("postes", d.Postes.Count)
                .Set("textos", d.Textos.Count(t => t.Tipo != "BLOCO"))
                .Set("obra", new JObj().Set("postes", obra.Postes.Count).Set("vaos", obra.Vaos.Count).Set("avulsos", obra.Avulsos.Count)
                                       .Set("ignorados", obra.Ignorados.Count))
                .Set("avisos", avisos.Cast<object>().ToList())
                .Set("erros_de_leitura", b.Avisos.Cast<object>().ToList());
        }

        public static JObj Gerar(DesenhoBruto b, Desenho d, Obra obra, RegrasProjeto regras, DateTime agora)
        {
            Dictionary<string, string> destino = new Dictionary<string, string>();
            foreach (Elemento e in obra.Elementos)
                if (!string.IsNullOrEmpty(e.TextoHandle))
                    destino[e.TextoHandle + "|" + e.TextoOriginal] = e.Numero + (e.GuiaUsada != null ? " (pela guia)" : "") + (e.Estado != Estados.Ok ? " — " + string.Join("; ", e.Mensagens.ToArray()) : "");
            foreach (TextoIgnorado t in obra.Ignorados) destino[t.Handle + "|" + t.Texto] = "ignorado — " + t.Motivo;

            List<TextoLido> lidos = LeitorProjeto.ClassificarTextos(d.Textos, regras);
            List<object> textos = new List<object>();
            foreach (TextoLido tl in lidos)
            {
                if (tl.T.Tipo == "BLOCO") continue;
                string dest;
                if (!destino.TryGetValue(tl.T.Handle + "|" + tl.T.Texto, out dest)) dest = tl.T.Tipo == "BLOCO_DEF" ? "texto da definição de bloco" : "—";
                textos.Add(new JObj().Set("h", tl.T.Handle).Set("tipo", tl.T.Tipo).Set("texto", tl.T.Texto).Set("layer", tl.T.Layer).Set("cor", tl.T.Cor)
                    .Set("entidade", tl.R.Entidade).Set("operacao", tl.R.Operacao).Set("ativo", tl.R.Ativo).Set("destino", dest));
            }
            return new JObj()
                .Set("formato", "plugin-diagnostico").Set("versao", 1).Set("plugin", ObraArquivo.VersaoPlugin)
                .Set("gerado_em", agora.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture))
                .Set("regras", regras.Nome).Set("leitura", (regras.Leitura ?? new ConfigLeitura()).ParaJson())
                .Set("resumo", Resumo(b, d, obra, regras))
                .Set("textos", textos)
                .Set("bruto", b.ParaJson());
        }

        /// <summary>Linhas curtas para a linha de comando.</summary>
        public static List<string> Linhas(JObj diag)
        {
            JObj r = diag.Obj("resumo");
            List<string> l = new List<string>();
            JObj o = r.Obj("obra");
            l.Add("Textos: " + Json.ComoTexto(r["textos"]) + " · postes (blocos): " + Json.ComoTexto(r["postes"]) + " · redes MT/BT: "
                  + Json.ComoTexto(r["redes_mt"]) + "/" + Json.ComoTexto(r["redes_bt"]) + " · guias: " + Json.ComoTexto(r["guias"]));
            l.Add("Obra: " + Json.ComoTexto(o["postes"]) + " postes, " + Json.ComoTexto(o["vaos"]) + " vãos, " + Json.ComoTexto(o["avulsos"])
                  + " avulsos, " + Json.ComoTexto(o["ignorados"]) + " textos ignorados pelas regras");
            foreach (string a in Json.ListaDeTextos(r["avisos"])) l.Add("⚠ " + a);
            int erros = (r.Lista("erros_de_leitura") ?? new List<object>()).Count;
            if (erros > 0) l.Add("⚠ " + erros + " entidade(s) não puderam ser lidas (detalhes no arquivo)");
            return l;
        }
    }
}
