using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OrcamentoNanoCAD.Nucleo;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;

namespace OrcamentoNanoCAD
{
    public partial class Comandos
    {
        [CommandMethod("ORC_SINCRONIZAR_JSON")]
        public void SincronizarJson()
        {
            Rodar("ORC_SINCRONIZAR_JSON", s =>
            {
                JObj recebido;
                using (var janela = new OpenFileDialog { Filter = "Alterações incrementais OrçaCore|*.json" })
                {
                    if (janela.ShowDialog() != DialogResult.OK) return;
                    recebido = Json.Ler(File.ReadAllText(janela.FileName, DesenhoJson.Utf8)) as JObj;
                }
                DesenhoJson.EscreverDxf(recebido, TextWriter.Null);
                DesenhoJsonModelo.GarantirSemProxies(recebido);
                var destino = s.Doc.Database;
                var plano = DesenhoJsonModelo.PlanejarPatch(CapturarDxf(destino), recebido);
                foreach (string diagnostico in plano.Diagnosticos) Aplicacao.Log("ORC_SINCRONIZAR_JSON: " + diagnostico);
                if (plano.Excluidos.Count > 0 && (s.Obra != null || PersistenciaDwg.Procurar(s) != null))
                    throw new InvalidDataException("Exclusões bloqueadas quando existe uma obra OrçaCore associada.");
                if (plano.Novos.Count + plano.Modificados.Count + plano.Excluidos.Count == 0)
                { Aplicacao.Msg("Nenhuma alteração."); return; }
                if (!Confirmar(s.Doc.Editor, "Sincronização experimental: " + plano.Novos.Count + " novos, " +
                    plano.Modificados.Count + " modificados, " + plano.Excluidos.Count + " excluídos?") ||
                    !SalvarBackupDxf(destino)) return;
                // Revalidar após o diálogo/backup antes de escrever.
                plano = DesenhoJsonModelo.PlanejarPatch(CapturarDxf(destino), recebido);
                string temp = CaminhoDxfTemporario();
                try
                {
                    // Reaplica os metadados validados depois que todas as entidades têm handle no destino.
                    var fonteJson = (JObj)Json.Ler(Json.Escrever(recebido, false));
                    var entidadesFonte = DesenhoJsonModelo.Entidades(fonteJson);
                    foreach (string h in plano.XDataNovos.Keys.Concat(plano.Modificados.Where(h =>
                        DesenhoJsonModelo.XDataObservacao(entidadesFonte[h]).Count > 0)))
                        entidadesFonte[h].Lista("tags").RemoveAll(o => Json.ComoNumero(((IList<object>)o)[0]) >= 1000);
                    EscreverDxfNoArquivo(fonteJson, temp);
                    using (var fonte = new Database(false, true))
                    {
                        AbrirDxfTemporario(fonte, temp);
                        using (var leitura = fonte.TransactionManager.StartTransaction())
                        using (var escrita = destino.TransactionManager.StartTransaction())
                        {
                            var de = GuiasDeNota.Modelo(fonte, leitura, OpenMode.ForRead);
                            var para = GuiasDeNota.Modelo(destino, escrita, OpenMode.ForRead);
                            var src = IdsDoModelo(de, leitura, true);
                            var dst = IdsDoModelo(para, escrita, false);
                            if (!new HashSet<string>(src.Keys, StringComparer.OrdinalIgnoreCase).SetEquals(plano.Novos.Concat(plano.Modificados)))
                                throw new InvalidDataException("O CAD alterou os handles da fonte incremental.");
                            foreach (string h in plano.Novos)
                                if (dst.ContainsKey(h)) throw new InvalidDataException("Colisão de handle: " + h);
                            foreach (string h in plano.Modificados)
                                AtualizarGeometria((Entity)escrita.GetObject(dst[h], OpenMode.ForWrite),
                                    (Entity)leitura.GetObject(src[h], OpenMode.ForRead));
                            if (plano.Novos.Count > 0)
                            {
                                using (var ids = new ObjectIdCollection(plano.Novos.Select(h => src[h]).ToArray()))
                                using (var mapa = new IdMapping())
                                {
                                    fonte.WblockCloneObjects(ids, para.ObjectId, mapa, DuplicateRecordCloning.Ignore, false);
                                    var novos = new Dictionary<string, Entity>(StringComparer.OrdinalIgnoreCase);
                                    foreach (string h in plano.Novos)
                                    {
                                        var par = mapa[src[h]];
                                        if (!par.IsCloned || par.Value.IsNull)
                                            throw new InvalidOperationException("Falha na clonagem: " + h);
                                        var entidade = escrita.GetObject(par.Value, OpenMode.ForRead) as Entity;
                                        if (entidade == null || entidade.OwnerId != para.ObjectId)
                                            throw new InvalidOperationException("Entidade fora do Model Space.");
                                        novos.Add(h, entidade);
                                    }
                                    foreach (var xdata in plano.XDataNovos)
                                    {
                                        var entidade = novos[xdata.Key];
                                        var tags = DesenhoJsonModelo.RemapearXDataPatch(xdata.Value, entidade.Handle.ToString(), h =>
                                        {
                                            Entity nova;
                                            if (novos.TryGetValue(h, out nova)) return nova.Handle.ToString();
                                            ObjectId id;
                                            return dst.TryGetValue(h, out id)
                                                ? escrita.GetObject(id, OpenMode.ForRead).Handle.ToString() : null;
                                        });
                                        GravarXDataPatch(destino, escrita, entidade, tags);
                                    }
                                }
                            }
                            foreach (string h in plano.Excluidos)
                                escrita.GetObject(dst[h], OpenMode.ForWrite).Erase();
                            escrita.Commit();
                        }
                    }
                    s.UltimoBruto = null; s.UltimoDesenho = null;
                    Aplicacao.Msg("Sincronização concluída. Exporte uma nova base antes de continuar editando.");
                }
                finally { if (File.Exists(temp)) File.Delete(temp); }
            });
        }

        private static void GravarXDataPatch(Database destino, Transaction escrita, Entity entidade, IList<IList<object>> tags)
        {
            string nome = (string)tags[0][1];
            var apps = (RegAppTable)escrita.GetObject(destino.RegAppTableId, OpenMode.ForRead);
            if (!apps.Has(nome))
            {
                if (!apps.IsWriteEnabled) apps.UpgradeOpen();
                var app = new RegAppTableRecord { Name = nome };
                apps.Add(app);
                escrita.AddNewlyCreatedDBObject(app, true);
            }
            var valores = new List<TypedValue>();
            for (int i = 0; i < tags.Count; i++)
            {
                int codigo = (int)Json.ComoNumero(tags[i][0]).Value;
                string valor = (string)tags[i][1];
                if (codigo == 1010)
                {
                    double x = double.Parse(valor, CultureInfo.InvariantCulture);
                    double y = double.Parse((string)tags[++i][1], CultureInfo.InvariantCulture);
                    double z = double.Parse((string)tags[++i][1], CultureInfo.InvariantCulture);
                    valores.Add(new TypedValue(codigo, new Point3d(x, y, z)));
                }
                else if (codigo == 1070)
                    valores.Add(new TypedValue(codigo, short.Parse(valor, CultureInfo.InvariantCulture)));
                else valores.Add(new TypedValue(codigo, valor));
            }
            if (!entidade.IsWriteEnabled) entidade.UpgradeOpen();
            using (var dados = new ResultBuffer(valores.ToArray())) entidade.XData = dados;
        }
    }
}
