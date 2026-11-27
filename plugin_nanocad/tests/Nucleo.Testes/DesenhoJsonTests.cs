using System.IO;
using System.Linq;
using System.Text;
using OrcamentoNanoCAD.Nucleo;
using Xunit;

public class DesenhoJsonTests
{
    [Fact]
    public void DxfSemCampoDeCodificacaoUsaUtf8()
    {
        var json = new JObj();
        Assert.Same(DesenhoJson.Utf8, DesenhoJson.CodificacaoDxf(json));
        Assert.Same(DesenhoJson.Utf8, DesenhoJson.CodificacaoDxf(json.Set("encoding", null)));
    }

    [Theory]
    [InlineData("")]
    [InlineData("UTF-8")]
    [InlineData("utf8")]
    public void MarcadoresUtf8PreservamAcentos(string marcador)
    {
        var encoding = DesenhoJson.CodificacaoDxf(new JObj().Set("encoding", marcador));
        Assert.Empty(encoding.GetPreamble());
        Assert.Equal(new byte[] { 0xC3, 0xA7 }, encoding.GetBytes("ç"));
    }

    [Fact]
    public void DxfMantemBytesLegadosERecusaMarcadorDesconhecido()
    {
        var json = new JObj().Set("encoding", "bytes-iso-8859-1");
        Assert.Equal(new byte[] { 0xE7 }, DesenhoJson.CodificacaoDxf(json).GetBytes("ç"));
        Assert.Throws<EncoderFallbackException>(() => DesenhoJson.CodificacaoDxf(json).GetBytes("€"));
        json.Set("encoding", "outro");
        Assert.Throws<InvalidDataException>(() => DesenhoJson.CodificacaoDxf(json));
    }

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

    [Fact]
    public void ExportaSoAEntidadeDoModeloEMantemDependencias()
    {
        string original = "0\nSECTION\n2\nBLOCKS\n0\nBLOCK\n2\n*Paper_Space0\n0\nLINE\n5\n32\n0\nENDBLK\n" +
            "0\nBLOCK\n2\nPOSTE\n0\nCIRCLE\n5\n33\n0\nENDBLK\n0\nENDSEC\n" +
            "0\nSECTION\n2\nENTITIES\n0\nPOLYLINE\n5\nA1\n67\n1\n0\nVERTEX\n5\nA2\n0\nSEQEND\n5\nA3\n" +
            "0\nINSERT\n5\nB1\n410\nModel\n2\nPOSTE\n0\nATTRIB\n5\nB2\n0\nSEQEND\n5\nB3\n" +
            "0\nLINE\n5\nC1\n410\nLayout1\n0\nENDSEC\n" +
            "0\nSECTION\n2\nOBJECTS\n0\nXRECORD\n5\n34\n1000\nInformação\n0\nENDSEC\n0\nEOF\n";
        var json = Ler(original);
        DesenhoJsonModelo.RestringirAoModelo(json);
        Assert.Equal("B1", Assert.Single(DesenhoJsonModelo.Entidades(json)).Key);
        var dxf = new StringWriter();
        DesenhoJson.EscreverDxf(json, dxf);
        Assert.Contains("\nB2\n", dxf.ToString().Replace("\r\n", "\n"));
        Assert.Contains("Informação", dxf.ToString());
        Assert.Contains("CIRCLE", dxf.ToString());
        Assert.DoesNotContain("\nA2\n", dxf.ToString().Replace("\r\n", "\n"));
        Assert.DoesNotContain("\n32\n", dxf.ToString().Replace("\r\n", "\n"));
    }

    [Fact]
    public void AtualizacaoNaoTrocaHandleNemAceitaMudancaDeMetadadosOuOrigem()
    {
        string linha = "0\nSECTION\n2\nENTITIES\n0\nLINE\n5\nAB\n8\nCABOS\n10\n1.25\n20\n15\n11\n20\n21\n30\n" +
            "1001\nANOTAPOSTE\n1005\nAA\n0\nENDSEC\n0\nEOF\n";
        var atual = Ler(linha);
        DesenhoJsonModelo.MarcarOrigem(atual);
        string json = Json.Escrever(atual, false);
        var atualizado = (JObj)Json.Ler(json.Replace("\"1.25\"", "\"2.75\""));
        Assert.Equal(new[] { "AB" }, DesenhoJsonModelo.PlanejarAtualizacao(atual, atualizado));
        var metadados = (JObj)Json.Ler(json.Replace("\"AA\"", "\"BB\""));
        Assert.Throws<InvalidDataException>(() => DesenhoJsonModelo.PlanejarAtualizacao(atual, metadados));
        var outroHandle = (JObj)Json.Ler(json.Replace("\"AB\"", "\"AC\""));
        Assert.Throws<InvalidDataException>(() => DesenhoJsonModelo.PlanejarAtualizacao(atual, outroHandle));
        Assert.Throws<InvalidDataException>(() => DesenhoJsonModelo.PlanejarAtualizacao(Ler(linha.Replace("1.25", "3.50")), atualizado));
    }

    [Fact]
    public void SubstituicaoRecusaReferenciaDeLayoutOuDicionarioAoModelo()
    {
        string cabecalho = "0\nSECTION\n2\nENTITIES\n0\nLINE\n5\nAB\n0\nLINE\n5\nAC\n67\n1\n1005\nAB\n0\nENDSEC\n";
        string original = cabecalho + "0\nSECTION\n2\nOBJECTS\n0\nXRECORD\n5\nB2\n330\nAB\n0\nENDSEC\n0\nEOF\n";
        var modelo = Ler(original);
        DesenhoJsonModelo.RestringirAoModelo(modelo);
        Assert.Single(DesenhoJsonModelo.Entidades(modelo));
        Assert.Throws<InvalidDataException>(() => DesenhoJsonModelo.GarantirSubstituicaoSemReferencias(Ler(original), modelo));
        Assert.Throws<InvalidDataException>(() => DesenhoJsonModelo.GarantirSubstituicaoSemReferencias(Ler(cabecalho + "0\nEOF\n"), modelo));
        DesenhoJsonModelo.GarantirSubstituicaoSemReferencias(Ler(original.Replace("1005\nAB", "1005\nAC").Replace("330\nAB", "330\nAC")), modelo);
    }

    [Fact]
    public void RejeitaProxiesERegistroDisfarcado()
    {
        var proxy = Ler("0\nSECTION\n2\nENTITIES\n0\nACAD_PROXY_ENTITY\n5\nAB\n0\nENDSEC\n0\nEOF\n");
        Assert.Throws<System.NotSupportedException>(() => DesenhoJsonModelo.GarantirSemProxies(proxy));
        var registros = ((JObj)proxy.Lista("sections").Single()).Lista("records");
        ((JObj)registros.Single()).Set("type", "LINE");
        Assert.Throws<InvalidDataException>(() => DesenhoJson.EscreverDxf(proxy, TextWriter.Null));
    }

    [Fact]
    public void ContaEntidadesR12SemHandleEIgnoraComentarioESubordinados()
    {
        const string entidades = "999\nComentário inicial\n0\nLINE\n8\n0\n0\nPOLYLINE\n8\n0\n0\nVERTEX\n8\n0\n0\nSEQEND\n";
        var r12 = DesenhoComEntidades("AC1009", entidades);
        Assert.Equal(2, DesenhoJsonModelo.ContarParaImportacao(r12));
        Assert.Contains("Handle ausente em LINE", Assert.Throws<InvalidDataException>(() => DesenhoJsonModelo.Entidades(r12)).Message);
        var moderno = DesenhoComEntidades("AC1027", entidades);
        Assert.Contains("Handle ausente em LINE", Assert.Throws<InvalidDataException>(() => DesenhoJsonModelo.ContarParaImportacao(moderno)).Message);
    }

    [Fact]
    public void ImportacaoDiferenciaHandleInvalidoDeRepetido()
    {
        var invalido = DesenhoComEntidades("AC1009", "0\nLINE\n5\n0\n");
        Assert.Contains("Handle inválido em LINE", Assert.Throws<InvalidDataException>(() => DesenhoJsonModelo.ContarParaImportacao(invalido)).Message);
        var repetido = DesenhoComEntidades("AC1009", "0\nLINE\n5\nAB\n0\nCIRCLE\n5\n00ab\n");
        Assert.Contains("Handle repetido (AB) em CIRCLE", Assert.Throws<InvalidDataException>(() => DesenhoJsonModelo.ContarParaImportacao(repetido)).Message);
    }

    private static JObj DesenhoComEntidades(string versao, string entidades)
    {
        return Ler("0\nSECTION\n2\nHEADER\n9\n$ACADVER\n1\n" + versao + "\n0\nENDSEC\n" +
            "0\nSECTION\n2\nENTITIES\n" + entidades + "0\nENDSEC\n0\nEOF\n");
    }

    private static JObj Ler(string dxf)
    {
        return DesenhoJson.LerDxf(new StringReader(dxf));
    }
}
