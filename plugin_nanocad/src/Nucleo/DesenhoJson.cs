using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace OrcamentoNanoCAD.Nucleo
{
    // Contrato geométrico independente de leitor-obra. Valores DXF permanecem strings.
    public static class DesenhoJson
    {
        public static JObj LerDxf(TextReader leitor)
        {
            var secoes = new List<object>();
            List<object> registros = null, tags = null;
            bool fim = false;
            string linha;
            while ((linha = leitor.ReadLine()) != null)
            {
                int codigo;
                string valor = leitor.ReadLine();
                if (valor == null || !int.TryParse(linha, NumberStyles.Integer, CultureInfo.InvariantCulture, out codigo))
                    throw new InvalidDataException("Par DXF inválido.");
                if (codigo == 0 && valor.Trim() == "EOF") { if (registros != null) throw new InvalidDataException("Seção aberta."); fim = true; break; }
                if (codigo == 0 && valor.Trim() == "SECTION")
                {
                    if (registros != null || leitor.ReadLine() == null) throw new InvalidDataException("Seção inválida.");
                    string nome = leitor.ReadLine();
                    if (string.IsNullOrEmpty(nome)) throw new InvalidDataException("Seção sem nome.");
                    registros = new List<object>(); tags = null;
                    secoes.Add(new JObj().Set("name", nome.Trim()).Set("records", registros));
                    continue;
                }
                if (registros == null) throw new InvalidDataException("Tag fora de seção.");
                if (codigo == 0 && valor.Trim() == "ENDSEC") { registros = null; tags = null; continue; }
                if (codigo == 0 || tags == null)
                {
                    tags = new List<object>();
                    registros.Add(new JObj().Set("type", codigo == 0 ? valor.Trim() : "data").Set("tags", tags));
                }
                tags.Add(new object[] { codigo, valor });
            }
            if (!fim) throw new InvalidDataException("DXF sem EOF.");
            return new JObj().Set("format", "cad-dxf-tags-v1").Set("sections", secoes);
        }

        public static void EscreverDxf(JObj documento, TextWriter escritor)
        {
            if (documento == null || documento.Str("format") != "cad-dxf-tags-v1" || documento.Lista("sections") == null)
                throw new InvalidDataException("Formato geométrico inválido.");
            var nomes = new HashSet<string>();
            foreach (object item in documento.Lista("sections"))
            {
                JObj secao = item as JObj;
                if (secao == null || secao.Lista("records") == null) throw new InvalidDataException("Seção inválida.");
                string nome = secao.Str("name");
                if (string.IsNullOrEmpty(nome) || !nomes.Add(nome)) throw new InvalidDataException("Nome de seção inválido ou repetido.");
                Par(escritor, 0, "SECTION"); Par(escritor, 2, nome);
                foreach (object obj in secao.Lista("records"))
                {
                    JObj registro = obj as JObj;
                    if (registro == null || registro.Lista("tags") == null) throw new InvalidDataException("Registro inválido.");
                    foreach (object tag in registro.Lista("tags"))
                    {
                        var par = tag as List<object>;
                        if (par == null || par.Count != 2 || !(par[1] is string)) throw new InvalidDataException("Tag inválida.");
                        double? n = Json.ComoNumero(par[0]);
                        if (!n.HasValue || n.Value != Math.Floor(n.Value) || n.Value < 0 || n.Value > 1071) throw new InvalidDataException("Código DXF inválido.");
                        string v = (string)par[1];
                        if (n.Value == 0 && (v.Trim() == "EOF" || v.Trim() == "SECTION" || v.Trim() == "ENDSEC")) throw new InvalidDataException("Delimitador dentro do registro.");
                        Par(escritor, (int)n.Value, v);
                    }
                }
                Par(escritor, 0, "ENDSEC");
            }
            if (!nomes.Contains("ENTITIES")) throw new InvalidDataException("Seção ENTITIES ausente.");
            Par(escritor, 0, "EOF");
        }

        private static void Par(TextWriter w, int codigo, string valor)
        {
            if (valor.IndexOfAny(new[] { '\r', '\n', '\0' }) >= 0) throw new InvalidDataException("Valor DXF contém quebra de linha ou NUL.");
            w.WriteLine(codigo.ToString(CultureInfo.InvariantCulture)); w.WriteLine(valor);
        }
    }
}
