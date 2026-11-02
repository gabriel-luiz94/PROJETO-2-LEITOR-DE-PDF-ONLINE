// Nucleo/ConfigLeitura.cs — como achar a REDE, os POSTES e as LINHAS GUIAS no desenho (TASK-006 do plugin).
// Antes era fixo no padrão de Pernambuco (layers REDE_MT/REDE_BT, bloco com "poste" no nome); agora é configurável por projeto
// (tela "Configurar leitura", guardada em %APPDATA%\OrcamentoNanoCAD\leitura_<projeto>.json). O padrão continua o de Pernambuco,
// igual a services/dxf_geometria.py da aplicação.
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace OrcamentoNanoCAD.Nucleo
{
    public class ConfigLeitura
    {
        // rede: layer cujo nome CONTÉM um destes trechos (sem diferença de maiúsculas)
        public List<string> LayersRedeMT = new List<string> { "REDE_MT" };
        public List<string> LayersRedeBT = new List<string> { "REDE_BT" };
        public bool CorConfirma = true;                         // vermelho = MT, verde = BT; cor que contradiz a layer descarta a linha
        // poste: bloco cujo NOME contém um destes trechos, OU bloco/círculo numa destas layers
        public List<string> BlocosPoste = new List<string> { "poste" };
        public List<string> LayersPoste = new List<string>();
        // guias: vazio = qualquer linha curta fora da rede (como a aplicação); preenchido = só linhas nessas layers. LEADER/MLEADER sempre valem.
        public List<string> LayersGuia = new List<string>();
        public double RaioAlturas = 4;          // texto → rede/poste: até 4 × a altura do texto
        public double GuiaMaxAlturas = 40;      // linha fora da rede vale como guia até 40 × a altura típica
        public double TolPosteAlturas = 1.5;    // poste "na" rede: até 1,5 × a altura típica

        private static bool Contem(IEnumerable<string> trechos, string nome)
        {
            string n = (nome ?? "").ToUpperInvariant();
            foreach (string t in trechos)
            {
                string x = (t ?? "").Trim().ToUpperInvariant();
                if (x.Length > 0 && n.Contains(x)) return true;
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

        /// <summary>A layer diz se é rede; a cor confirma (se ligada). Cor neutra vale a layer; cor contrária descarta.</summary>
        public string ClassificarRede(string layer, string hex)
        {
            string daLayer = TipoPelaLayer(layer);
            if (daLayer == null || !CorConfirma) return daLayer;
            string daCor = TipoPelaCor(hex);
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
            return new JObj().Set("layers_rede_mt", L(LayersRedeMT)).Set("layers_rede_bt", L(LayersRedeBT)).Set("cor_confirma", CorConfirma)
                .Set("blocos_poste", L(BlocosPoste)).Set("layers_poste", L(LayersPoste)).Set("layers_guia", L(LayersGuia))
                .Set("raio_alturas", RaioAlturas).Set("guia_max_alturas", GuiaMaxAlturas).Set("tol_poste_alturas", TolPosteAlturas);
        }

        public static ConfigLeitura De(JObj o)
        {
            ConfigLeitura c = new ConfigLeitura();
            if (o == null) return c;
            if (o["layers_rede_mt"] is List<object>) c.LayersRedeMT = Json.ListaDeTextos(o["layers_rede_mt"]);
            if (o["layers_rede_bt"] is List<object>) c.LayersRedeBT = Json.ListaDeTextos(o["layers_rede_bt"]);
            c.CorConfirma = o.Bool("cor_confirma", true);
            if (o["blocos_poste"] is List<object>) c.BlocosPoste = Json.ListaDeTextos(o["blocos_poste"]);
            if (o["layers_poste"] is List<object>) c.LayersPoste = Json.ListaDeTextos(o["layers_poste"]);
            if (o["layers_guia"] is List<object>) c.LayersGuia = Json.ListaDeTextos(o["layers_guia"]);
            c.RaioAlturas = Positivo(o.Num("raio_alturas"), 4);
            c.GuiaMaxAlturas = Positivo(o.Num("guia_max_alturas"), 40);
            c.TolPosteAlturas = Positivo(o.Num("tol_poste_alturas"), 1.5);
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
