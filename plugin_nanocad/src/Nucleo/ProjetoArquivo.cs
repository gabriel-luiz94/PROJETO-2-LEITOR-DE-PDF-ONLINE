// Nucleo/ProjetoArquivo.cs — o PROJETO DE TRABALHO do OrçaCore (TASK-020 do plugin / TASK-163 do principal).
// O Salvar do plugin grava "<desenho>.obra" na pasta do desenho: o mesmo envelope do .obra.json (dados + cad) mais o bloco
// "projeto_trabalho" {nome, numero_obra, sap, resultado, salvo_em} do programa. O OrçaCore abre esse arquivo em 📁 Projeto → Abrir e, ao
// gravar, PRESERVA a seção "cad" (services/projeto_obra.py). Divisão: linhas (dados) e cad = o desenho (plugin); SAP e resultado = o
// programa; nome e número de obra = os dois. Para não apagar sem aviso o que foi editado no programa, o plugin guarda em
// cad.assinaturaLinhas uma impressão das linhas que gravou; se o arquivo tiver outras linhas, é porque o programa mudou (pergunta antes).
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace OrcamentoNanoCAD.Nucleo
{
    public static class ProjetoArquivo
    {
        public const string Extensao = ".obra";

        /// <summary>"C:\obras\Rua.dwg" → "C:\obras\Rua.obra".</summary>
        public static string Caminho(string desenho)
        {
            return string.IsNullOrEmpty(desenho) ? null : Path.ChangeExtension(desenho, Extensao);
        }

        /// <summary>Impressão das linhas (tabela · operação · ativo), insensível à formatação do JSON de quem gravou.</summary>
        public static string AssinaturaLinhas(JObj dados)
        {
            StringBuilder sb = new StringBuilder();
            foreach (string tabela in new[] { "cabos", "outros" })
                foreach (JObj l in ((dados != null ? dados.Lista(tabela) : null) ?? new List<object>()).OfType<JObj>())
                    sb.Append(tabela[0]).Append('|').Append((l.Str("operacao") ?? "").Trim()).Append('|').Append((l.Str("ativo") ?? "").Trim()).Append('\n');
            using (SHA1 h = SHA1.Create())
                return string.Concat(h.ComputeHash(Encoding.UTF8.GetBytes(sb.ToString())).Select(b => b.ToString("x2")).ToArray());
        }

        /// <summary>null = pode gravar sem perguntar; senão o motivo para perguntar antes de gravar por cima.</summary>
        public static string Conflito(JObj existente)
        {
            if (existente == null) return null;
            JObj cad = existente.Obj("cad");
            if (cad == null)
                return "Já existe um projeto do OrçaCore com esse nome na pasta do desenho (criado no programa). As linhas dele serão substituídas pelas do desenho.";
            string esperada = cad.Str("assinaturaLinhas");
            if (!string.IsNullOrEmpty(esperada) && esperada != AssinaturaLinhas(existente.Obj("dados")))
                return "As linhas do projeto foram alteradas no OrçaCore depois do último salvamento do desenho. Salvar agora grava as linhas do desenho por cima delas.";
            return null;
        }

        /// <summary>O .obra a gravar: envelope do plugin + projeto_trabalho (SAP, resultado e o que o programa guardou ficam; nome e número vêm da obra).</summary>
        public static JObj Montar(Obra obra, JObj existente, DateTime agoraUtc)
        {
            JObj env = ObraArquivo.Envelope(obra, agoraUtc);
            JObj antigo = existente != null ? existente.Obj("projeto_trabalho") : null;
            JObj pt = new JObj();
            if (antigo != null) foreach (string k in antigo.Keys) pt.Set(k, antigo[k]);
            pt.Set("nome", string.IsNullOrEmpty(obra.Nome) ? "obra" : obra.Nome);
            pt.Set("numero_obra", obra.NumeroObra ?? (antigo != null ? antigo.Str("numero_obra") : "") ?? "");
            if (pt["sap"] == null) pt.Set("sap", new JObj());
            pt.Set("salvo_em", agoraUtc.ToString("yyyy-MM-dd'T'HH:mm:ss'+00:00'", CultureInfo.InvariantCulture));
            env.Set("projeto_trabalho", pt);
            JObj cad = env.Obj("cad");
            if (cad != null) cad.Set("assinaturaLinhas", AssinaturaLinhas(env.Obj("dados")));
            return env;
        }

        /// <summary>Nome e número de obra guardados no projeto (o programa pode ter mudado em Propriedades).</summary>
        public static void LerTrabalho(JObj env, Obra obra)
        {
            JObj pt = env != null ? env.Obj("projeto_trabalho") : null;
            if (pt == null) return;
            string nome = pt.Str("nome");
            if (!string.IsNullOrEmpty(nome)) obra.Nome = nome;
            obra.NumeroObra = pt.Str("numero_obra") ?? "";
        }

        /// <summary>Grava sem deixar o arquivo pela metade (temporário + troca).</summary>
        public static void GravarAtomico(string caminho, string texto)
        {
            string tmp = caminho + ".tmp";
            File.WriteAllText(tmp, texto, new UTF8Encoding(false));
            if (File.Exists(caminho)) File.Replace(tmp, caminho, null);
            else File.Move(tmp, caminho);
        }

        public static JObj LerArquivo(string caminho)
        {
            if (string.IsNullOrEmpty(caminho) || !File.Exists(caminho)) return null;
            try { return Json.Ler(File.ReadAllText(caminho, Encoding.UTF8)) as JObj; }
            catch (Exception) { return null; }
        }
    }
}
