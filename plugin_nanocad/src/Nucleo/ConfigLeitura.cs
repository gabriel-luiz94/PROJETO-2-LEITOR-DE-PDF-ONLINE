// Nucleo/ConfigLeitura.cs — como achar a REDE, os POSTES e as LINHAS GUIAS no desenho (TASK-006 do plugin).
// Antes era fixo no padrão de Pernambuco (layers REDE_MT/REDE_BT, bloco com "poste" no nome); agora é configurável por projeto
// (tela "Configurar leitura", guardada em %APPDATA%\OrcamentoNanoCAD\leitura_<projeto>.json). O padrão continua o de Pernambuco,
// igual a services/dxf_geometria.py da aplicação.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace OrcamentoNanoCAD.Nucleo
{
    public class ConfigLeitura
    {
        // rede: layer cujo nome CONTÉM um destes trechos (sem diferença de maiúsculas)
        public List<string> LayersRedeMT = new List<string> { "REDE_MT" };
        public List<string> LayersRedeBT = new List<string> { "REDE_BT" };
        public string Regra = "COR_MANDA";
        public object CorMT, CorBT;

        // poste: bloco cujo NOME contém um destes trechos, OU bloco/círculo numa destas layers
        public List<string> BlocosPoste = new List<string> { "poste" };
        public List<string> LayersPoste = new List<string>();
        // guias: vazio = qualquer linha curta fora da rede (como a aplicação); preenchido = só linhas nessas layers. LEADER/MLEADER sempre valem.
        public List<string> LayersGuia = new List<string>();
        public double RaioAlturas = 4;          // texto → rede/poste: até 4 × a altura do texto
        public double GuiaMaxAlturas = 40;      // linha fora da rede vale como guia até 40 × a altura típica
        public double TolPosteAlturas = 2.5;    // poste "na" rede (e blocos juntos = um poste): até 2,5 × a altura típica
        public double GuiaToqueAlturas = 6;     // guia vale quando a ponta chega a até 6 × a altura do texto (do contorno)
        public double RaioReservaAlturas = 12;  // sem guia e sem nada no raio: elemento SEM anotação até 12 × a altura (vira Conferir)
        public string PlacaRegex = "^[A-Z]{1,2}\\d{5,7}$";   // placa do poste (B039571), ligada ao poste mais perto

        /// <summary>O nome CONTÉM o trecho (sem diferença de maiúsculas); "=NOME" exige o nome exato (ex.: "=0" = só a layer 0).</summary>
        public static bool Contem(IEnumerable<string> trechos, string nome)
        {
            string n = (nome ?? "").Trim().ToUpperInvariant();
            foreach (string t in trechos)
            {
                string x = (t ?? "").Trim().ToUpperInvariant();
                if (x.StartsWith("=")) { if (n == x.Substring(1).Trim()) return true; continue; }
                if (x.IndexOfAny(new char[] { '*', '?' }) >= 0)
                {
                    string expressao = Regex.Escape(x).Replace(@"\*", ".*").Replace(@"\?", ".");
                    if (Regex.IsMatch(n, "^" + expressao + "$")) return true;
                }
                else if (x.Length > 0 && n.Contains(x)) return true;
            }
            return false;
        }

        public string TipoPelaLayer(string layer)
        {
            if (Contem(LayersRedeMT, layer)) return "MT";
            if (Contem(LayersRedeBT, layer)) return "BT";
            return null;
        }

        public static string TipoPelaCor(string hex)
        {
            if (string.IsNullOrEmpty(hex) || hex.Length < 7) return null;
            int r, g, b;
            if (!int.TryParse(hex.Substring(1, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out r)
                || !int.TryParse(hex.Substring(3, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out g)
                || !int.TryParse(hex.Substring(5, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out b)) return null;
            if (r >= 170 && g <= 100 && b <= 100) return "MT";
            if (g >= 140 && r <= 120 && b <= 120) return "BT";
            return null;
        }

        private static bool NaFaixa(JObj faixa, string canal, int valor)
        {
            List<object> limites = faixa.Lista(canal);
            return limites != null && limites.Count == 2 && Json.ComoNumero(limites[0]) <= valor && valor <= Json.ComoNumero(limites[1]);
        }

        private static bool CasaCor(object especificacao, string hex, int rgb, bool padrao)
        {
            List<object> lista = especificacao as List<object>;
            if (lista != null) return lista.Count == 0 ? padrao : lista.Any(x => CasaCor(x, hex, rgb, false));
            JObj faixa = especificacao as JObj;
            if (faixa != null)
                return NaFaixa(faixa, "r", rgb >> 16) && NaFaixa(faixa, "g", (rgb >> 8) & 255) && NaFaixa(faixa, "b", rgb & 255);
            string valor = (Json.ComoTexto(especificacao) ?? "").Trim();
            if (valor.Length == 0) return padrao;
            int aci;
            if (int.TryParse(valor, out aci))
            {
                if (aci <= 0 || aci > 255) return false;
                valor = aci == 7 ? "#ffffff" : Cores.AciParaHex(aci);
            }
            return string.Equals(valor, hex, StringComparison.OrdinalIgnoreCase);
        }

        public string CorDoProjeto(string hex)
        {
            int rgb;
            if (hex == null || hex.Length != 7 || !int.TryParse(hex.Substring(1), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out rgb)) return null;
            string padrao = TipoPelaCor(hex);
            if (CasaCor(CorMT, hex, rgb, padrao == "MT")) return "MT";
            if (CasaCor(CorBT, hex, rgb, padrao == "BT")) return "BT";
            return null;
        }

        public string ClassificarRede(string layer, string hex)
        {
            string daLayer = TipoPelaLayer(layer);
            if (Regra == "SO_LAYER") return daLayer;
            string daCor = CorDoProjeto(hex);
            if (Regra == "LAYER_MANDA") return daLayer ?? daCor;
            return daCor == null || daCor == daLayer ? daLayer : null;
        }

        public bool EhBlocoPoste(string nomeBloco, string layer)
        {
            return Contem(BlocosPoste, nomeBloco) || Contem(LayersPoste, layer);
        }

        public bool EhCirculoPoste(string layer) { return LayersPoste.Count > 0 && Contem(LayersPoste, layer); }

        public bool PodeSerGuia(string layer)
        {
            if (TipoPelaLayer(layer) != null) return false;
            return LayersGuia.Count == 0 || Contem(LayersGuia, layer);
        }

        // ── JSON ──
        private static List<object> L(List<string> l) { return l.Cast<object>().ToList(); }

        public JObj ParaJson()
        {
            return new JObj().Set("layers_rede_mt", L(LayersRedeMT)).Set("layers_rede_bt", L(LayersRedeBT)).Set("cor_confirma", Regra != "SO_LAYER")
                .Set("regra", Regra).Set("cor_mt", CorMT).Set("cor_bt", CorBT)
                .Set("blocos_poste", L(BlocosPoste)).Set("layers_poste", L(LayersPoste)).Set("layers_guia", L(LayersGuia))
                .Set("raio_alturas", RaioAlturas).Set("guia_max_alturas", GuiaMaxAlturas).Set("tol_poste_alturas", TolPosteAlturas)
                .Set("guia_toque_alturas", GuiaToqueAlturas).Set("raio_reserva_alturas", RaioReservaAlturas).Set("placa_regex", PlacaRegex);
        }

        public static ConfigLeitura De(JObj o)
        {
            ConfigLeitura c = new ConfigLeitura();
            if (o == null) return c;
            if (o["layers_rede_mt"] is List<object>) c.LayersRedeMT = Json.ListaDeTextos(o["layers_rede_mt"]);
            if (o["layers_rede_bt"] is List<object>) c.LayersRedeBT = Json.ListaDeTextos(o["layers_rede_bt"]);
            string regra = (o.Str("regra") ?? "").Trim().ToUpperInvariant();
            c.Regra = regra == "SO_LAYER" || regra == "LAYER_MANDA" || regra == "COR_MANDA" ? regra : (o.Bool("cor_confirma", true) ? "COR_MANDA" : "SO_LAYER");
            c.CorMT = o["cor_mt"];
            c.CorBT = o["cor_bt"];
            if (o["blocos_poste"] is List<object>) c.BlocosPoste = Json.ListaDeTextos(o["blocos_poste"]);
            if (o["layers_poste"] is List<object>) c.LayersPoste = Json.ListaDeTextos(o["layers_poste"]);
            if (o["layers_guia"] is List<object>) c.LayersGuia = Json.ListaDeTextos(o["layers_guia"]);
            c.RaioAlturas = Positivo(o.Num("raio_alturas"), 4);
            c.GuiaMaxAlturas = Positivo(o.Num("guia_max_alturas"), 40);
            c.TolPosteAlturas = Positivo(o.Num("tol_poste_alturas"), 2.5);
            c.GuiaToqueAlturas = Positivo(o.Num("guia_toque_alturas"), 6);
            c.RaioReservaAlturas = Positivo(o.Num("raio_reserva_alturas"), 12);
            if (o.Has("placa_regex")) c.PlacaRegex = o.Str("placa_regex") ?? "";
            return c;
        }

        private static double Positivo(double? v, double padrao) { return v.HasValue && v.Value > 0 ? v.Value : padrao; }

        public ConfigLeitura Copia() { return De(ParaJson()); }

        /// <summary>"A, B" → [A, B] (para a tela de configuração).</summary>
        public static List<string> Lista(string texto)
        {
            return (texto ?? "").Split(new[] { ',', ';' }).Select(x => x.Trim()).Where(x => x.Length > 0).ToList();
        }
    }
}
