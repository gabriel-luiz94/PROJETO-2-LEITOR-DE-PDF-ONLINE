using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using OrcamentoNanoCAD.Nucleo;
using Teigha.DatabaseServices;
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
                if (plano.Excluidos.Count > 0 && (s.Obra != null || PersistenciaDwg.Procurar(s) != null))
                    throw new InvalidDataException("Exclusões bloqueadas quando existe uma obra OrçaCore associada.");
                if (plano.Novos.Count + plano.Modificados.Count + plano.Excluidos.Count == 0)
                { Aplicacao.Msg("Nenhuma alteração."); return; }
                if (!Confirmar(s.Doc.Editor, "Sincronização experimental: " + plano.Novos.Count + " novos, " +
                    plano.Modificados.Count + " modificados, " + plano.Excluidos.Count + " excluídos?") ||
                    !SalvarBackupDxf(destino)) return;
                // Revalidar após o diálogo/backup antes de escrever.
                plano = DesenhoJsonModelo.PlanejarPatch(CapturarDxf(destino), recebido);
                var registros = DesenhoJsonModelo.Entidades(recebido);
                string temp = CaminhoDxfTemporario();
                try
                {
                    // Retirar apenas os links de observação da fonte temporária; serão gravados
                    // explicitamente com handles do destino, não com handles do banco auxiliar.
                    var fonteJson = (JObj)Json.Ler(Json.Escrever(recebido, false));
                    foreach (var r in DesenhoJsonModelo.Entidades(fonteJson).Values)
                        if (DesenhoJsonModelo.XDataObservacao(r).Count > 0)
                            r.Lista("tags").RemoveAll(o => Json.ComoNumero(((IList<object>)o)[0]) >= 1000);
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
                                    foreach (string h in plano.Novos)
                                    {
                                        var par = mapa[src[h]];
                                        if (!par.IsCloned || par.Value.IsNull)
                                            throw new InvalidOperationException("Falha na clonagem: " + h);
                                        var entidade = escrita.GetObject(par.Value, OpenMode.ForWrite) as Entity;
                                        if (entidade == null || entidade.OwnerId != para.ObjectId)
                                            throw new InvalidOperationException("Entidade fora do Model Space.");
                                        var xd = DesenhoJsonModelo.XDataObservacao(registros[h]);
                                        if (xd.Count == 0) continue;
                                        var apps = (RegAppTable)escrita.GetObject(destino.RegAppTableId, OpenMode.ForRead);
                                        if (!apps.Has("CADSUITE_OBS"))
                                        {
                                            apps.UpgradeOpen();
                                            var app = new RegAppTableRecord { Name = "CADSUITE_OBS" };
                                            apps.Add(app); escrita.AddNewlyCreatedDBObject(app, true);
                                        }
                                        var valores = new List<TypedValue>();
                                        foreach (var tag in xd)
                                        {
                                            int codigo = (int)Json.ComoNumero(tag[0]).Value;
                                            string valor = (string)tag[1];
                                            if (codigo == 1005)
                                                valor = escrita.GetObject(dst[valor.Trim().TrimStart('0')], OpenMode.ForRead).Handle.ToString();
                                            valores.Add(new TypedValue(codigo, valor));
                                        }
                                        using (var dados = new ResultBuffer(valores.ToArray())) entidade.XData = dados;
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
    }
}
