// Plugin/Compat.cs — convivência com versões do nanoCAD com API menor (nanoCAD 5, TASK-014 do plugin).
// A DLL do nanoCAD 5 é compilada contra as declarações de src/ApiNanoCAD5 (o que o plugin usa da API do nanoCAD 26). No nanoCAD 5,
// um tipo/membro que não exista lá dá TypeLoadException/MissingMemberException quando o .NET compila (JIT) o método que o usa.
// Por isso cada recurso "opcional" fica num método (ou lambda) próprio, chamado por Compat.Tentar: se faltar, o recurso é desligado
// (com um aviso no log e na linha de comando, uma vez) e o resto do plugin segue.
using System;
using System.Collections.Generic;

namespace OrcamentoNanoCAD
{
    public static class Compat
    {
        private static readonly Dictionary<string, string> _faltando = new Dictionary<string, string>();

        /// <summary>Recursos que este nanoCAD não tem (nome → erro), para o ORC_AJUDA e o diagnóstico.</summary>
        public static IDictionary<string, string> Faltando { get { return _faltando; } }

        public static bool Falta(string recurso) { lock (_faltando) return _faltando.ContainsKey(recurso); }

        public static void Registrar(string recurso, Exception e)
        {
            bool novo;
            lock (_faltando) { novo = !_faltando.ContainsKey(recurso); if (novo) _faltando[recurso] = e.GetType().Name + ": " + e.Message; }
            if (!novo) return;
            Aplicacao.Log("compatibilidade: este nanoCAD não tem '" + recurso + "' (" + e.GetType().Name + ": " + e.Message + ") — recurso desligado");
            try { Aplicacao.Msg("Aviso: este nanoCAD não oferece '" + recurso + "' para plugins; o plugin segue sem esse recurso."); }
            catch (Exception) { }
        }

        private static bool DeApi(Exception e) { return e is TypeLoadException || e is MissingMemberException || e is EntryPointNotFoundException; }

        /// <summary>Executa; se a API não tiver o necessário, registra e devolve false (outras exceções seguem normalmente).</summary>
        public static bool Tentar(string recurso, Action acao)
        {
            if (Falta(recurso)) return false;
            try { acao(); return true; }
            catch (Exception e)
            {
                if (!DeApi(e)) throw;
                Registrar(recurso, e);
                return false;
            }
        }

        public static T Tentar<T>(string recurso, Func<T> f, T padrao)
        {
            if (Falta(recurso)) return padrao;
            try { return f(); }
            catch (Exception e)
            {
                if (!DeApi(e)) throw;
                Registrar(recurso, e);
                return padrao;
            }
        }

        /// <summary>Texto para o ORC_AJUDA.</summary>
        public static string Resumo()
        {
            lock (_faltando)
                return _faltando.Count == 0 ? "API do nanoCAD completa para o plugin"
                    : "recursos indisponíveis neste nanoCAD: " + string.Join(", ", new List<string>(_faltando.Keys).ToArray());
        }
    }
}
