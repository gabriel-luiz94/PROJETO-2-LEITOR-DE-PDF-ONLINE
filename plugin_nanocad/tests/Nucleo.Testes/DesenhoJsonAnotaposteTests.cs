using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

public class DesenhoJsonAnotaposteTests
{
    private const string DesenhoId = "c39177a2-8b75-48d2-86a7-26ca64843d95";
    private const string Peca = "0\nMTEXT\n5\nAC\n1001\nANOTAPOSTE\n1000\nNOTA1\n1000\nPECA\n1000\nG=a\n";
    private const string Guia = "0\nLWPOLYLINE\n5\nAD\n1001\nANOTAPOSTE\n1000\nNOTA1\n1000\nGUIA\n1000\nA\n1000\nAD\n" +
        "1010\n1\n1020\n2\n1030\n0\n1070\n1\n1005\nAC\n1070\n1\n1005\nAB\n";
    private const string Alvo = "0\nLINE\n5\nAB\n10\n1\n";
    private const string Legado = "0\nMTEXT\n5\nAE\n1\nAnterior\n";

    [Fact]
    public void SubstituiTextoLegadoEPlanejaRemapeamentoDoGrupo()
    {
        var baseDwg = Ler(Alvo + Legado);
        var patch = CriarPatch(baseDwg, Peca + Guia, new[] { "AC", "AD" }, new[] { "AE" });
        var plano = DesenhoJsonModelo.PlanejarPatch(baseDwg, patch);
        Assert.Equal(new[] { "AC", "AD" }, plano.Novos);
        Assert.Equal("AE", Assert.Single(plano.Excluidos));
        var tags = DesenhoJsonModelo.RemapearXDataPatch(plano.XDataNovos["AD"], "F2",
            h => h == "AC" ? "F1" : h == "AB" ? "AB" : null);
        Assert.Equal("F2", tags[4][1]);
        Assert.Equal(new[] { "F1", "AB" }, tags.Where(t => Json.ComoNumero(t[0]) == 1005).Select(t => (string)t[1]));
        Assert.Equal("AD", plano.XDataNovos["AD"][4][1]);
        Assert.Throws<InvalidDataException>(() => DesenhoJsonModelo.RemapearXDataPatch(plano.XDataNovos["AD"], "F2", h => null));
    }

    [Fact]
    public void RecusaGrupoIncompletoEReferenciaSemComprovacaoDeOrigem()
    {
        var baseDwg = Ler(Alvo);
        var incompleto = CriarPatch(baseDwg, Guia, new[] { "AD" }, new string[0]);
        Assert.Contains("todas as peças", Falha(baseDwg, incompleto));

        var patch = CriarPatch(baseDwg, Peca + Guia, new[] { "AC", "AD" }, new string[0]);
        var outraCaptura = Ler(Alvo + Legado);
        Assert.Contains("sem registro de referência", Falha(outraCaptura, patch));
        patch.Obj("cadsuitePatch").Set("references", new List<object> { DesenhoJsonModelo.Entidades(baseDwg)["AB"] });
        Assert.Equal(2, DesenhoJsonModelo.PlanejarPatch(outraCaptura, patch).Novos.Count);
        Assert.Contains("alvo da guia AB", Falha(Ler(Alvo.Replace("10\n1", "10\n3")), patch));
    }

    [Fact]
    public void SubstituiGrupoExistenteMesmoComReferenciasInternasDasEntidadesExcluidas()
    {
        var baseDwg = Ler(Alvo + Peca + Guia);
        var novos = (Peca + Guia).Replace("AC\n", "BC\n").Replace("AD\n", "BD\n");
        var recebido = CriarPatch(baseDwg, novos, new[] { "BC", "BD" }, new[] { "AC", "AD" });
        Assert.Equal(2, DesenhoJsonModelo.PlanejarPatch(baseDwg, recebido).Excluidos.Count);
        var externo = Ler(Alvo + Peca + Guia + "0\nLINE\n5\nAF\n67\n1\n1001\nOUTRO\n1005\n00AC\n");
        recebido.Obj("cadsuitePatch").Set("references", new List<object> { DesenhoJsonModelo.Entidades(baseDwg)["AB"] });
        Assert.Contains("Exclusão bloqueada por referência", Falha(externo, recebido));
        Assert.Contains("AF", DesenhoJsonModelo.Entidades(externo).Keys);
    }

    private static string Falha(JObj atual, JObj patch)
    {
        return Assert.Throws<InvalidDataException>(() => DesenhoJsonModelo.PlanejarPatch(atual, patch)).Message;
    }

    private static JObj CriarPatch(JObj baseDwg, string entidades, string[] novos, string[] excluidos)
    {
        DesenhoJsonModelo.MarcarOrigem(baseDwg);
        var anteriores = DesenhoJsonModelo.Entidades(baseDwg);
        return Ler(entidades).Set("cadsuitePatch", new JObj().Set("version", 2).Set("sourceDrawingId", DesenhoId)
            .Set("sourceModel", baseDwg.Str("sourceModel")).Set("sourceDependencies", baseDwg.Str("sourceDependencies"))
            .Set("added", novos.Cast<object>().ToList()).Set("modified", new List<object>())
            .Set("deleted", excluidos.Cast<object>().ToList()).Set("references", new List<object>())
            .Set("before", excluidos.Select(h => (object)anteriores[h]).ToList()));
    }

    private static JObj Ler(string entidades)
    {
        string dxf = "0\nSECTION\n2\nHEADER\n9\n$FINGERPRINTGUID\n2\n" + DesenhoId + "\n0\nENDSEC\n" +
            "0\nSECTION\n2\nTABLES\n0\nLAYER\n5\n10\n2\n0\n0\nSTYLE\n5\n11\n2\nSTANDARD\n0\nENDSEC\n" +
            "0\nSECTION\n2\nENTITIES\n" + entidades + "0\nENDSEC\n0\nEOF\n";
        return (JObj)Json.Ler(Json.Escrever(DesenhoJson.LerDxf(new StringReader(dxf)), false));
    }
}
