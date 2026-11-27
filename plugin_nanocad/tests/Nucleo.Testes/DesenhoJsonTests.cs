using System.IO;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

public class DesenhoJsonTests
{
    [Fact]
    public void PreservaTagsEHandles()
    {
        string dxf = "0\nSECTION\n2\nENTITIES\n0\nLINE\n5\nAB\n1001\nANOTAPOSTE\n1005\nCD\n10\n123.456\n0\nENDSEC\n0\nEOF\n";
        var json = DesenhoJson.LerDxf(new StringReader(dxf));
        var saida = new StringWriter();
        DesenhoJson.EscreverDxf((JObj)Json.Ler(Json.Escrever(json, false)), saida);
        Assert.Equal(dxf, saida.ToString().Replace("\r\n", "\n"));
    }
    [Fact]
    public void RejeitaDxfTruncado()
    {
        Assert.Throws<InvalidDataException>(() => DesenhoJson.LerDxf(new StringReader("0\nSECTION\n2\nENTITIES\n")));
    }
}
