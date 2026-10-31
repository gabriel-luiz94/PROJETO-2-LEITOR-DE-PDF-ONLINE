// Plugin/PersistenciaDwg.cs — a obra guardada nos DOIS lugares: dentro do DWG (dicionário de objetos nomeados → Xrecord, em pedaços
// de texto) e no arquivo "<desenho>.obra.json" na pasta do desenho. Os dois têm o MESMO conteúdo (o envelope leitor-obra).
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using HostMgd.ApplicationServices;
using OrcamentoNanoCAD.Nucleo;
using Teigha.DatabaseServices;

namespace OrcamentoNanoCAD
{
    public static class PersistenciaDwg
    {
        public const string Dicionario = "ORCAMENTO_LEITOR";
        public const string Registro = "OBRA";
        private const int Pedaco = 250;

        public static void GravarNoDwg(Document doc, string json)
        {
            Database db = doc.Database;
            using (DocumentLock trava = doc.LockDocument())
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForWrite);
                DBDictionary meu;
                if (nod.Contains(Dicionario)) meu = (DBDictionary)tr.GetObject(nod.GetAt(Dicionario), OpenMode.ForWrite);
                else
                {
                    meu = new DBDictionary();
                    nod.SetAt(Dicionario, meu);
                    tr.AddNewlyCreatedDBObject(meu, true);
                }
                ResultBuffer rb = new ResultBuffer();
                for (int i = 0; i < json.Length; i += Pedaco)
                    rb.Add(new TypedValue((int)DxfCode.Text, json.Substring(i, Math.Min(Pedaco, json.Length - i))));
                Xrecord xr;
                if (meu.Contains(Registro))
                {
                    xr = (Xrecord)tr.GetObject(meu.GetAt(Registro), OpenMode.ForWrite);
                    xr.Data = rb;
                }
                else
                {
                    xr = new Xrecord();
                    xr.Data = rb;
                    meu.SetAt(Registro, xr);
                    tr.AddNewlyCreatedDBObject(xr, true);
                }
                tr.Commit();
            }
        }

        /// <summary>O texto guardado no DWG, ou null.</summary>
        public static string LerDoDwg(Document doc)
        {
            Database db = doc.Database;
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);
                if (!nod.Contains(Dicionario)) return null;
                DBDictionary meu = (DBDictionary)tr.GetObject(nod.GetAt(Dicionario), OpenMode.ForRead);
                if (!meu.Contains(Registro)) return null;
                Xrecord xr = (Xrecord)tr.GetObject(meu.GetAt(Registro), OpenMode.ForRead);
                StringBuilder sb = new StringBuilder();
                ResultBuffer rb = xr.Data;
                if (rb != null) foreach (TypedValue tv in rb) sb.Append(tv.Value as string);
                tr.Commit();
                return sb.Length > 0 ? sb.ToString() : null;
            }
        }

        public static void GravarArquivo(string caminho, string json)
        {
            File.WriteAllText(caminho, json + "\n", new UTF8Encoding(false));
        }

        /// <summary>Resumo da obra salva para a tela inicial: "28/10 14:32 · no DWG e em X.obra.json".</summary>
        public class ObraSalva
        {
            public string Json;
            public string Onde;
            public string Quando;
        }

        public static ObraSalva Procurar(Sessao s)
        {
            string noDwg = null, noArquivo = null;
            try { noDwg = LerDoDwg(s.Doc); } catch (System.Exception) { }
            string arq = s.ArquivoObra;
            try { if (arq != null && File.Exists(arq)) noArquivo = File.ReadAllText(arq, Encoding.UTF8); } catch (System.Exception) { }
            string escolhido = noDwg ?? noArquivo;
            if (escolhido == null) return null;
            ObraSalva r = new ObraSalva();
            r.Json = escolhido;
            List<string> onde = new List<string>();
            if (noDwg != null) onde.Add("no DWG");
            if (noArquivo != null) onde.Add("em " + Path.GetFileName(arq));
            r.Onde = string.Join(" e ", onde.ToArray());
            try
            {
                JObj env = Json.Ler(escolhido) as JObj;
                JObj cad = env != null ? env.Obj("cad") : null;
                r.Quando = cad != null ? cad.Str("salvo_em") : null;
            }
            catch (System.Exception) { }
            return r;
        }
    }
}
