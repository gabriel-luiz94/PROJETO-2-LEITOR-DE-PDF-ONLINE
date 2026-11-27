using System.IO;
using System.Collections.Generic;
using OrcamentoNanoCAD.Nucleo;
using Xunit;
public class DesenhoJsonPatchTests
{
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
}
