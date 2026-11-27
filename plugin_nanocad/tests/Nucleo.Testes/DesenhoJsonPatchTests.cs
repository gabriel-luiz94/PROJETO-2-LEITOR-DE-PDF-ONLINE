using System.Linq;
using System.IO;
using System.Collections.Generic;
using OrcamentoNanoCAD.Nucleo;
using Xunit;
public class DesenhoJsonPatchTests
{
    private const string DesenhoId = "31ded1fe-5e10-45e6-967c-c715be471b12";
    private static JObj Ler(string entidades)
    {
        return (JObj)Json.Ler(Json.Escrever(DesenhoJson.LerDxf(new StringReader(
            "0\nSECTION\n2\nENTITIES\n" + entidades + "0\nENDSEC\n0\nEOF\n")), false));
    }
    [Fact]
    public void AcrescimoComOrigemExata()
    {
        var a = Ler("0\nLINE\n5\nAB\n10\n1\n");
        DesenhoJsonModelo.MarcarOrigem(a);
        var b = Ler("0\nLINE\n5\nAC\n10\n2\n");
        b.Set("cadsuitePatch", new JObj().Set("version", 1).Set("sourceModel", a.Str("sourceModel"))
            .Set("sourceDependencies", a.Str("sourceDependencies"))
            .Set("added", new List<object>{"AC"}).Set("modified", new List<object>()).Set("deleted", new List<object>())
            .Set("before", new List<object>()));
        Assert.Equal("AC", Assert.Single(DesenhoJsonModelo.PlanejarPatch(a,b).Novos));
        b.Obj("cadsuitePatch").Set("sourceModel", "outra");
        Assert.Throws<InvalidDataException>(() => DesenhoJsonModelo.PlanejarPatch(a,b));
    }

    [Fact]
    public void RecusaGuiaNovaQueApontaParaEntidadeExcluidaNoMesmoPatch()
    {
        var atual = Ler("0\nLINE\n5\nAB\n10\n1\n");
        DesenhoJsonModelo.MarcarOrigem(atual);
        var recebido = Ler("0\nLINE\n5\nAC\n10\n2\n1001\nCADSUITE_OBS\n1000\nOBS1\n1000\nGUIA\n1005\n00ab\n");
        recebido.Set("cadsuitePatch", new JObj().Set("version", 1)
            .Set("sourceModel", atual.Str("sourceModel")).Set("sourceDependencies", atual.Str("sourceDependencies"))
            .Set("added", new List<object> { "AC" }).Set("modified", new List<object>())
            .Set("deleted", new List<object> { "AB" })
            .Set("before", new List<object> { DesenhoJsonModelo.Entidades(atual)["AB"] }));

        var erro = Assert.Throws<InvalidDataException>(() => DesenhoJsonModelo.PlanejarPatch(atual, recebido));
        Assert.Contains("aponta para uma entidade marcada para exclusão", erro.Message);
    }

    [Fact]
    public void PatchV2AceitaObservacaoComMudancasEmEntidadesETabelasAlheias()
    {
        var baseExportada = LerV2("1.000", "1", "1");
        var atual = LerV2("1", "2", "3");
        var recebido = GuiaV2(baseExportada, "1");
        var plano = DesenhoJsonModelo.PlanejarPatch(atual, recebido);
        Assert.Equal("AC", Assert.Single(plano.Novos));
        Assert.Contains("sourceModel", Assert.Single(plano.Diagnosticos));
        Assert.Contains("sourceDependencies", Assert.Single(plano.Diagnosticos));
    }

    [Fact]
    public void PatchV2RecusaOutroDwgAlvoAlteradoEDependenciaUsadaAlterada()
    {
        var baseExportada = LerV2("1", "1", "1");
        var recebido = GuiaV2(baseExportada, "1");
        var outro = LerV2("1", "1", "1", "ea65277f-e52e-4ab2-861b-cf32bbaa90a6");
        Assert.Contains("identificador persistente", Falha(outro, recebido));
        Assert.Contains("alvo da guia AB", Falha(LerV2("2", "1", "1"), recebido));
        Assert.Contains("LAYER 0", Falha(LerV2("1", "1", "1"), GuiaV2(baseExportada, "2")));
        recebido.Obj("cadsuitePatch").Set("references", new List<object>());
        Assert.Contains("cada alvo distinto", Falha(baseExportada, recebido));
    }

    [Fact]
    public void PatchV2ModificaApenasAEntidadeCujoAntesConfere()
    {
        var baseExportada = LerV2("1.00", "1", "1");
        var atual = LerV2("1", "2", "3");
        var recebido = GuiaV2(baseExportada, "1");
        var modificado = DesenhoJsonModelo.Entidades(Ler("0\nLINE\n5\nAB\n8\n0\n10\n4\n"))["AB"];
        foreach (JObj secao in recebido.Lista("sections").OfType<JObj>())
            if (secao.Str("name") == "ENTITIES")
            {
                secao.Lista("records").Clear();
                secao.Lista("records").Add(modificado);
            }
        recebido.Obj("cadsuitePatch").Set("references", new List<object>()).Set("added", new List<object>())
            .Set("modified", new List<object> { "AB" })
            .Set("before", new List<object> { DesenhoJsonModelo.Entidades(baseExportada)["AB"] });
        Assert.Equal("AB", Assert.Single(DesenhoJsonModelo.PlanejarPatch(atual, recebido).Modificados));
        Assert.Contains("Conflito na entidade AB", Falha(LerV2("2", "1", "1"), recebido));
    }

    private static string Falha(JObj atual, JObj recebido)
    {
        return Assert.Throws<InvalidDataException>(() => DesenhoJsonModelo.PlanejarPatch(atual, recebido)).Message;
    }

    private static string Tabelas(string corUsada, string corAlheia)
    {
        return "0\nSECTION\n2\nTABLES\n0\nLAYER\n5\n10\n2\n0\n62\n" + corUsada +
            "\n6\nCONTINUOUS\n0\nLTYPE\n5\n11\n2\nCONTINUOUS\n0\nLAYER\n5\n12\n2\nALHEIA\n62\n" +
            corAlheia + "\n0\nENDSEC\n";
    }

    private static JObj LerV2(string posicaoAlvo, string posicaoAlheia, string corAlheia, string id = DesenhoId)
    {
        string dxf = "0\nSECTION\n2\nHEADER\n9\n$FINGERPRINTGUID\n2\n{" + id + "}\n0\nENDSEC\n" +
            Tabelas("1", corAlheia) + "0\nSECTION\n2\nENTITIES\n0\nLINE\n5\nAB\n8\n0\n10\n" + posicaoAlvo +
            "\n0\nLINE\n5\nAE\n8\nALHEIA\n10\n" + posicaoAlheia + "\n0\nENDSEC\n0\nEOF\n";
        return (JObj)Json.Ler(Json.Escrever(DesenhoJson.LerDxf(new StringReader(dxf)), false));
    }

    private static JObj GuiaV2(JObj baseExportada, string corUsada)
    {
        DesenhoJsonModelo.MarcarOrigem(baseExportada);
        Assert.Equal(DesenhoId, baseExportada.Str("sourceDrawingId"));
        var recebido = (JObj)Json.Ler(Json.Escrever(DesenhoJson.LerDxf(new StringReader(Tabelas(corUsada, "1") +
            "0\nSECTION\n2\nENTITIES\n0\nLINE\n5\nAC\n8\n0\n10\n2\n1001\nCADSUITE_OBS\n1000\nOBS1\n1000\nGUIA\n1005\nAB\n0\nENDSEC\n0\nEOF\n")), false));
        return recebido.Set("cadsuitePatch", new JObj().Set("version", 2)
            .Set("sourceModel", baseExportada.Str("sourceModel")).Set("sourceDependencies", baseExportada.Str("sourceDependencies"))
            .Set("sourceDrawingId", baseExportada.Str("sourceDrawingId"))
            .Set("added", new List<object> { "AC" }).Set("modified", new List<object>())
            .Set("deleted", new List<object>()).Set("before", new List<object>())
            .Set("references", new List<object> { DesenhoJsonModelo.Entidades(baseExportada)["AB"] }));
    }
}
