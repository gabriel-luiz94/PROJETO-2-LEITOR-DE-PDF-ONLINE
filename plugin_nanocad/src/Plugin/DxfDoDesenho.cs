using System;
using System.Globalization;
using System.Reflection;
using HostMgd.ApplicationServices;
using Teigha.DatabaseServices;

namespace OrcamentoNanoCAD
{
    internal static class DxfDoDesenho
    {
        public static bool ArquivoAtual()
        {
            try
            {
                object alteracoes = Application.GetSystemVariable("DBMOD");
                return alteracoes != null && Convert.ToInt32(alteracoes, CultureInfo.InvariantCulture) == 0;
            }
            catch (Exception erro)
            {
                Aplicacao.Log("Consulta de DBMOD indisponível. Gerando DXF do desenho ativo. " + erro.Message);
                return false;
            }
        }

        public static void Exportar(Database banco, string caminho)
        {
            // Consulta a API instalada porque as sobrecargas variam entre gerações do nanoCAD.
            Type tipo = typeof(Database);
            MethodInfo metodo = tipo.GetMethod("DxfOut", new[] { typeof(string), typeof(int), typeof(bool) });
            // Precisão 16 escreve DXF ASCII. O booleano é a opção de miniatura.
            object[] argumentos = { caminho, 16, false };
            if (metodo == null)
            {
                metodo = tipo.GetMethod("DxfOut", new[] { typeof(string), typeof(int), typeof(DwgVersion) });
                argumentos = new object[] { caminho, 16, DwgVersion.AC1024 };
            }
            if (metodo == null)
            {
                metodo = tipo.GetMethod("DxfOut", new[] { typeof(string), typeof(int) });
                argumentos = new object[] { caminho, 16 };
            }
            if (metodo == null)
                throw new NotSupportedException("A API deste nanoCAD não disponibilizou a exportação DXF do desenho ativo.");
            try { metodo.Invoke(banco, argumentos); }
            catch (TargetInvocationException erro)
            {
                Exception causa = erro.InnerException ?? erro;
                throw new InvalidOperationException("Falha ao exportar o DXF do desenho ativo. " + causa.Message, causa);
            }
        }
    }
}
