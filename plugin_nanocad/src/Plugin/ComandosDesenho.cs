using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using OrcamentoNanoCAD.Nucleo;
using Teigha.DatabaseServices;
using Teigha.Runtime;

namespace OrcamentoNanoCAD
{
    public partial class Comandos
    {
        private static readonly UTF8Encoding Utf8Dxf = new UTF8Encoding(false, true);
        private static readonly Encoding BytesDxf = Encoding.GetEncoding(28591, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);

        [CommandMethod("CSEXPORTAR")]
        public void ExportarDesenhoJson()
        {
            Rodar("CSEXPORTAR", s =>
            {
                using (var janela = new SaveFileDialog { Filter = "Desenho JSON|*.json", OverwritePrompt = true })
                {
                    if (janela.ShowDialog() != DialogResult.OK) return;
                    JObj desenho = CapturarDxf(s.Doc.Database);
                    DesenhoJsonModelo.RestringirAoModelo(desenho);
                    DesenhoJsonModelo.MarcarOrigem(desenho);
                    File.WriteAllText(janela.FileName, Json.Escrever(desenho, false), Utf8Dxf);
                    Aplicacao.Msg("Geometria do Model Space exportada para JSON.");
                }
            });
        }

        [CommandMethod("CSIMPORTAR")]
        public void ImportarDesenhoJson()
        {
            Rodar("CSIMPORTAR", s => AplicarDesenhoJson(s, false, false));
        }

        [CommandMethod("CSSUBSTITUIR")]
        public void SubstituirDesenhoJson()
        {
            Rodar("CSSUBSTITUIR", s => AplicarDesenhoJson(s, true, false));
        }

        [CommandMethod("CSATUALIZAR")]
        public void AtualizarDesenhoJson()
        {
            Rodar("CSATUALIZAR", s => AplicarDesenhoJson(s, false, true));
        }

        private static void AplicarDesenhoJson(Sessao s, bool substituir, bool atualizar)
        {
            JObj recebido;
            using (var janela = new OpenFileDialog { Filter = "Desenho JSON|*.json" })
            {
                if (janela.ShowDialog() != DialogResult.OK) return;
                recebido = Json.Ler(File.ReadAllText(janela.FileName, Utf8Dxf)) as JObj;
            }

            // Valida a estrutura recebida antes de filtrar ou chamar a API nativa.
            DesenhoJson.EscreverDxf(recebido, TextWriter.Null);
            DesenhoJsonModelo.RestringirAoModelo(recebido);
            var entidades = DesenhoJsonModelo.Entidades(recebido);
            if (entidades.Count == 0) throw new InvalidDataException("O JSON não contém entidades no Model Space.");
            DesenhoJsonModelo.GarantirSemProxies(recebido);

            Database destino = s.Doc.Database;
            JObj modeloAtual = null;
            List<string> alterados = null;
            if (substituir || atualizar)
            {
                if (substituir && (s.Obra != null || PersistenciaDwg.Procurar(s) != null))
                    throw new InvalidOperationException("Este desenho tem uma obra do OrçaCore com referências a handles. Use CSIMPORTAR ou uma cópia sem a obra associada.");
                JObj integral = CapturarDxf(destino);
                modeloAtual = substituir ? (JObj)Json.Ler(Json.Escrever(integral, false)) : integral;
                DesenhoJsonModelo.RestringirAoModelo(modeloAtual);
                if (substituir) DesenhoJsonModelo.GarantirSubstituicaoSemReferencias(integral, modeloAtual);
                if (atualizar)
                {
                    alterados = DesenhoJsonModelo.PlanejarAtualizacao(modeloAtual, recebido);
                    if (alterados.Count == 0) { Aplicacao.Msg("Nenhuma alteração geométrica a aplicar."); return; }
                }
            }
            string temp = CaminhoDxfTemporario();
            try
            {
                string encoding = recebido.Str("encoding");
                if (encoding != "" && encoding != "bytes-iso-8859-1")
                    throw new InvalidDataException("Codificação DXF desconhecida.");
                using (var escritor = new StreamWriter(temp, false, encoding == "" ? Utf8Dxf : BytesDxf))
                    DesenhoJson.EscreverDxf(recebido, escritor);

                using (var fonte = new Database(false, true))
                {
                    AbrirDxfTemporario(fonte, temp);
                    using (var leitura = fonte.TransactionManager.StartTransaction())
                    using (var escrita = destino.TransactionManager.StartTransaction())
                    {
                        var de = GuiasDeNota.Modelo(fonte, leitura, OpenMode.ForRead);
                        var para = GuiasDeNota.Modelo(destino, escrita, OpenMode.ForRead);
                        var importados = IdsDoModelo(de, leitura, true);
                        var anteriores = IdsDoModelo(para, escrita, substituir);
                        if (importados.Count != entidades.Count)
                            throw new InvalidDataException("O nanoCAD não carregou todas as entidades do JSON.");
                        if (modeloAtual != null)
                        {
                            var exportados = DesenhoJsonModelo.Entidades(modeloAtual);
                            if (exportados.Count != anteriores.Count || !new HashSet<string>(anteriores.Keys, StringComparer.OrdinalIgnoreCase).SetEquals(exportados.Keys))
                                throw new InvalidDataException("Os handles do DXF não correspondem ao desenho ativo.");
                        }

                        if (atualizar)
                        {
                            if (!Confirmar(s.Doc.Editor, "Atualizar " + alterados.Count + " entidade(s) do Model Space?") ||
                                !SalvarBackupDxf(destino)) return;
                            foreach (string handle in alterados)
                            {
                                ObjectId anterior, novo;
                                if (!anteriores.TryGetValue(handle, out anterior) || !importados.TryGetValue(handle, out novo))
                                    throw new InvalidDataException("O nanoCAD remapeou o handle " + handle + ".");
                                AtualizarGeometria((Entity)escrita.GetObject(anterior, OpenMode.ForWrite),
                                    (Entity)leitura.GetObject(novo, OpenMode.ForRead));
                            }
                        }
                        else
                        {
                            if (substituir &&
                                (!Confirmar(s.Doc.Editor, "Substituir " + anteriores.Count + " entidade(s) do Model Space por " + importados.Count + "?") ||
                                !SalvarBackupDxf(destino))) return;
                            using (var ids = new ObjectIdCollection(new List<ObjectId>(importados.Values).ToArray()))
                            using (var mapa = new IdMapping())
                            {
                                // Mantém definições importadas separadas quando já existem nomes no destino.
                                fonte.WblockCloneObjects(ids, para.ObjectId, mapa, DuplicateRecordCloning.MangleName, false);
                                foreach (ObjectId id in ids)
                                    if (!mapa[id].IsCloned || mapa[id].Value.IsNull ||
                                        escrita.GetObject(mapa[id].Value, OpenMode.ForRead).OwnerId != para.ObjectId)
                                        throw new InvalidOperationException("A clonagem não entregou todas as entidades.");
                            }
                            if (substituir)
                                foreach (ObjectId id in anteriores.Values)
                                    escrita.GetObject(id, OpenMode.ForWrite).Erase();
                        }
                        escrita.Commit();
                    }
                }
                s.UltimoBruto = null;
                s.UltimoDesenho = null;
                Aplicacao.Msg("Operação concluída no Model Space. Os layouts foram mantidos.");
            }
            finally { if (File.Exists(temp)) File.Delete(temp); }
        }

        private static JObj CapturarDxf(Database banco)
        {
            string temp = CaminhoDxfTemporario();
            try
            {
                DxfDoDesenho.Exportar(banco, temp);
                JObj desenho;
                try
                {
                    using (var leitor = new StreamReader(temp, Utf8Dxf, false))
                        desenho = DesenhoJson.LerDxf(leitor);
                }
                catch (DecoderFallbackException)
                {
                    // O DXF antigo pode estar em uma página de códigos legada. Preserva os bytes.
                    using (var leitor = new StreamReader(temp, BytesDxf, false))
                        desenho = DesenhoJson.LerDxf(leitor).Set("encoding", "bytes-iso-8859-1");
                }
                return desenho;
            }
            finally { if (File.Exists(temp)) File.Delete(temp); }
        }

        private static string CaminhoDxfTemporario()
        {
            return Path.Combine(Path.GetTempPath(), "cad-" + Guid.NewGuid().ToString("N") + ".dxf");
        }

        private static void AbrirDxfTemporario(Database banco, string caminho)
        {
            MethodInfo metodo = typeof(Database).GetMethod("DxfIn", new[] { typeof(string), typeof(string) });
            if (metodo == null) throw new NotSupportedException("A API deste nanoCAD não expõe a leitura DXF necessária.");
            string log = Path.ChangeExtension(caminho, ".log");
            try { metodo.Invoke(banco, new object[] { caminho, log }); }
            catch (TargetInvocationException erro)
            {
                throw new InvalidDataException("O nanoCAD recusou o DXF recebido.", erro.InnerException ?? erro);
            }
            finally { if (File.Exists(log)) File.Delete(log); }
        }

        private static bool SalvarBackupDxf(Database banco)
        {
            using (var janela = new SaveFileDialog { Filter = "Backup DXF|*.dxf", FileName = "backup-cadsuite.dxf", OverwritePrompt = true })
            {
                janela.Title = "Salvar uma cópia do desenho antes da alteração";
                if (janela.ShowDialog() != DialogResult.OK) return false;
                if (File.Exists(janela.FileName)) throw new IOException("Escolha um novo arquivo para o backup.");
                string temporario = Path.Combine(Path.GetDirectoryName(janela.FileName), "backup-" + Guid.NewGuid().ToString("N") + ".dxf");
                try
                {
                    DxfDoDesenho.Exportar(banco, temporario);
                    if (new FileInfo(temporario).Length == 0) throw new IOException("O backup DXF saiu vazio.");
                    File.Move(temporario, janela.FileName);
                    return true;
                }
                finally { if (File.Exists(temporario)) File.Delete(temporario); }
            }
        }

        private static Dictionary<string, ObjectId> IdsDoModelo(BlockTableRecord modelo, Transaction tr, bool bloquearProxy)
        {
            var ids = new Dictionary<string, ObjectId>(StringComparer.OrdinalIgnoreCase);
            foreach (ObjectId id in modelo)
            {
                var entidade = tr.GetObject(id, OpenMode.ForRead) as Entity;
                if (entidade == null) continue;
                if (bloquearProxy && entidade.GetType().Name.IndexOf("Proxy", StringComparison.OrdinalIgnoreCase) >= 0)
                    throw new NotSupportedException("O desenho contém objeto proxy no Model Space. Operação cancelada.");
                ids.Add(entidade.Handle.ToString(), id);
            }
            return ids;
        }

        private static void AtualizarGeometria(Entity destino, Entity origem)
        {
            if (origem.GetType() != destino.GetType())
                throw new InvalidDataException("O tipo nativo da entidade mudou.");
            var l = destino as Line;
            if (l != null)
            {
                var linha = (Line)origem;
                l.StartPoint = linha.StartPoint; l.EndPoint = linha.EndPoint;
                return;
            }
            var c = destino as Circle;
            if (c != null)
            {
                var circulo = (Circle)origem;
                c.Center = circulo.Center; c.Radius = circulo.Radius;
                return;
            }
            var a = destino as Arc;
            if (a != null)
            {
                var arco = (Arc)origem;
                a.Center = arco.Center; a.Radius = arco.Radius;
                a.StartAngle = arco.StartAngle; a.EndAngle = arco.EndAngle;
                return;
            }
            var p = destino as Polyline;
            if (p != null)
            {
                var polilinha = (Polyline)origem;
                if (p.NumberOfVertices != polilinha.NumberOfVertices)
                    throw new InvalidDataException("Uma atualização preserva a quantidade de vértices.");
                p.Elevation = polilinha.Elevation;
                for (int i = 0; i < p.NumberOfVertices; i++)
                {
                    p.SetPointAt(i, polilinha.GetPoint2dAt(i));
                    p.SetBulgeAt(i, polilinha.GetBulgeAt(i));
                    p.SetStartWidthAt(i, polilinha.GetStartWidthAt(i));
                    p.SetEndWidthAt(i, polilinha.GetEndWidthAt(i));
                }
                return;
            }
            throw new NotSupportedException("Não é possível escrever a geometria deste tipo de entidade.");
        }
    }
}
