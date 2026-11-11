// GerarApiNanoCAD5 — gera src/ApiNanoCAD5/hostmgd.cs e hostdbmgd.cs: declarações (sem corpo de verdade) de TUDO que o plugin usa da API
// do nanoCAD, com os mesmos nomes, namespaces e assinaturas da API (lidos do nanoCAD 26). Compiladas para .NET 4, viram as referências
// para gerar a DLL do nanoCAD 5 aqui (não existe pacote da API do nanoCAD 5). Em tempo de execução, o nanoCAD 5 usa as DLLs dele
// (hostmgd/hostdbmgd, sem assinatura forte): o que existir lá com o mesmo nome e assinatura funciona; o que não existir só falha no
// trecho isolado que o usa (ver Plugin/Compat.cs).
//   uso: dotnet run --project ferramentas/GerarApiNanoCAD5 -- <pasta plugin_nanocad>
using System.Reflection;
using System.Text;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;

[AttributeUsage(AttributeTargets.Assembly)]
class ApiNanoCADAttribute : Attribute { public readonly string V; public ApiNanoCADAttribute(string v) { V = v; } }
[AttributeUsage(AttributeTargets.Assembly)]
class WinDesktopRefAttribute : Attribute { public readonly string V; public WinDesktopRefAttribute(string v) { V = v; } }

static class Program
{
    static readonly SymbolDisplayFormat Nome = new SymbolDisplayFormat(
        globalNamespaceStyle: SymbolDisplayGlobalNamespaceStyle.Included,
        typeQualificationStyle: SymbolDisplayTypeQualificationStyle.NameAndContainingTypesAndNamespaces,
        genericsOptions: SymbolDisplayGenericsOptions.IncludeTypeParameters,
        miscellaneousOptions: SymbolDisplayMiscellaneousOptions.UseSpecialTypes | SymbolDisplayMiscellaneousOptions.EscapeKeywordIdentifiers);

    static bool DaApi(ISymbol s) => s != null && s.ContainingAssembly != null && (s.ContainingAssembly.Name == "hostmgd" || s.ContainingAssembly.Name == "hostdbmgd");

    // tipo → membros usados
    static readonly Dictionary<INamedTypeSymbol, HashSet<ISymbol>> Usados = new(SymbolEqualityComparer.Default);

    static int Main(string[] args)
    {
        string raiz = Path.GetFullPath(args.Length > 0 ? args[0] : ".");
        string api = typeof(Program).Assembly.GetCustomAttribute<ApiNanoCADAttribute>()!.V;
        string wd = typeof(Program).Assembly.GetCustomAttribute<WinDesktopRefAttribute>()!.V;
        string netRef = Directory.GetDirectories(Path.Combine(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(wd))!)!)!, "microsoft.netcore.app.ref"))
            .OrderByDescending(d => d).Select(d => Path.Combine(d, "ref", "net6.0")).First(Directory.Exists);

        var arquivos = Directory.GetFiles(Path.Combine(raiz, "src", "Nucleo"), "*.cs").Concat(Directory.GetFiles(Path.Combine(raiz, "src", "Plugin"), "*.cs"));
        var opcoes = new CSharpParseOptions(LanguageVersion.CSharp5, preprocessorSymbols: new[] { "NANOCAD5" });
        var arvores = arquivos.Select(f => CSharpSyntaxTree.ParseText(File.ReadAllText(f), opcoes, f)).ToList();
        var refs = Directory.GetFiles(netRef, "*.dll").Concat(Directory.GetFiles(wd, "*.dll")).Concat(new[] { Path.Combine(api, "hostmgd.dll"), Path.Combine(api, "hostdbmgd.dll") })
            .Select(f => MetadataReference.CreateFromFile(f));
        var comp = CSharpCompilation.Create("OrcamentoNanoCAD", arvores, refs, new CSharpCompilationOptions(OutputKind.DynamicallyLinkedLibrary, allowUnsafe: false));
        var erros = comp.GetDiagnostics().Where(d => d.Severity == DiagnosticSeverity.Error).ToList();
        foreach (var e in erros.Take(30)) Console.Error.WriteLine(e);
        if (erros.Count > 0) { Console.Error.WriteLine(erros.Count + " erro(s) compilando o plugin contra a API do nanoCAD 26."); return 1; }

        foreach (var t in arvores)
        {
            var m = comp.GetSemanticModel(t);
            foreach (var n in t.GetRoot().DescendantNodes())
            {
                if (n is ForEachStatementSyntax fe)
                {
                    var info = m.GetForEachStatementInfo(fe);
                    Membro(info.GetEnumeratorMethod); Membro(info.MoveNextMethod); Membro(info.CurrentProperty); Membro(info.DisposeMethod);
                }
                if (n is UsingStatementSyntax || n is AttributeSyntax || n is IdentifierNameSyntax || n is GenericNameSyntax || n is MemberAccessExpressionSyntax
                    || n is ObjectCreationExpressionSyntax || n is QualifiedNameSyntax || n is ElementAccessExpressionSyntax || n is BinaryExpressionSyntax
                    || n is CastExpressionSyntax || n is AssignmentExpressionSyntax || n is PrefixUnaryExpressionSyntax)
                {
                    var si = m.GetSymbolInfo(n);
                    Simbolo(si.Symbol);
                    foreach (var c in si.CandidateSymbols) Simbolo(c);
                    var ti = m.GetTypeInfo(n);
                    Tipo(ti.Type); Tipo(ti.ConvertedType);
                    var conv = m.GetConversion(n);
                    if (conv.IsUserDefined) Membro(conv.MethodSymbol);
                }
                if (n is ClassDeclarationSyntax cd)
                {
                    var s = m.GetDeclaredSymbol(cd);
                    if (s != null) foreach (var i in s.AllInterfaces) { Tipo(i); foreach (var mm in i.GetMembers()) Membro(mm); }
                }
            }
        }
        // fecho: bases, interfaces da API, tipos das assinaturas
        bool mudou = true;
        while (mudou)
        {
            int antes = Usados.Count + Usados.Values.Sum(v => v.Count);
            foreach (var t in Usados.Keys.ToList())
            {
                Tipo(t.BaseType);
                foreach (var i in t.Interfaces) Tipo(i);
                if (t.ContainingType != null) Tipo(t.ContainingType);
                if (t.TypeKind == TypeKind.Delegate && t.DelegateInvokeMethod != null) Assinatura(t.DelegateInvokeMethod);
                if (t.TypeKind == TypeKind.Interface) foreach (var mm in t.GetMembers()) Membro(mm);
                foreach (var mm in Usados[t].ToList()) Assinatura(mm);
            }
            mudou = Usados.Count + Usados.Values.Sum(v => v.Count) != antes;
        }

        string saida = Path.Combine(raiz, "src", "ApiNanoCAD5");
        Directory.CreateDirectory(saida);
        foreach (string asm in new[] { "hostdbmgd", "hostmgd" })
        {
            var sb = new StringBuilder();
            sb.AppendLine("// GERADO por ferramentas/GerarApiNanoCAD5 — não editar à mão. Só as declarações que o plugin usa da API do nanoCAD (" + asm + "),");
            sb.AppendLine("// para compilar a DLL do nanoCAD 5 (.NET 4). Em tempo de execução valem as DLLs do próprio nanoCAD.");
            sb.AppendLine("#pragma warning disable 67, 108, 114, 169, 649, 659, 660, 661, 809");
            foreach (var g in Usados.Keys.Where(t => t.ContainingAssembly.Name == asm && t.ContainingType == null).GroupBy(t => t.ContainingNamespace.ToDisplayString()).OrderBy(g => g.Key))
            {
                sb.AppendLine("namespace " + g.Key + "\n{");
                foreach (var t in g.OrderBy(t => t.Name)) EscreverTipo(sb, t, "    ");
                sb.AppendLine("}");
            }
            if (asm == "hostdbmgd") sb.AppendLine("namespace OrcamentoNanoCAD.Api5 { internal static class Vazio { } }");
            Directory.CreateDirectory(Path.Combine(saida, asm));
            File.WriteAllText(Path.Combine(saida, asm, asm + ".cs"), sb.ToString());
            Console.WriteLine(asm + ": " + Usados.Keys.Count(t => t.ContainingAssembly.Name == asm) + " tipos, "
                              + Usados.Where(k => k.Key.ContainingAssembly.Name == asm).Sum(k => k.Value.Count) + " membros");
        }
        return 0;
    }

    static void Simbolo(ISymbol s)
    {
        if (s == null) return;
        if (s is INamedTypeSymbol nt) { Tipo(nt); return; }
        if (s is IMethodSymbol || s is IPropertySymbol || s is IFieldSymbol || s is IEventSymbol) Membro(s);
        if (s is IAliasSymbol al) Simbolo(al.Target);
    }

    static void Tipo(ITypeSymbol t)
    {
        if (t == null) return;
        if (t is IArrayTypeSymbol a) { Tipo(a.ElementType); return; }
        if (t is INamedTypeSymbol n)
        {
            foreach (var ta in n.TypeArguments) Tipo(ta);
            n = n.OriginalDefinition;
            if (!DaApi(n)) return;
            if (!Usados.ContainsKey(n)) Usados[n] = new HashSet<ISymbol>(SymbolEqualityComparer.Default);
        }
    }

    static void Membro(ISymbol s)
    {
        if (s == null) return;
        s = s.OriginalDefinition;
        if (s is IMethodSymbol ms && ms.MethodKind == MethodKind.PropertyGet || s is IMethodSymbol ms2 && ms2.MethodKind == MethodKind.PropertySet)
            s = ((IMethodSymbol)s).AssociatedSymbol!;
        if (s is IMethodSymbol red && red.ReducedFrom != null) s = red.ReducedFrom;
        if (!DaApi(s)) return;
        var t = s.ContainingType;
        Tipo(t);
        Usados[t].Add(s);
    }

    static void Assinatura(ISymbol s)
    {
        switch (s)
        {
            case IMethodSymbol m: Tipo(m.ReturnType); foreach (var p in m.Parameters) Tipo(p.Type); break;
            case IPropertySymbol p: Tipo(p.Type); foreach (var q in p.Parameters) Tipo(q.Type); break;
            case IFieldSymbol f: Tipo(f.Type); break;
            case IEventSymbol e: Tipo(e.Type); break;
        }
    }

    static string N(ITypeSymbol t) => t.ToDisplayString(Nome);

    static string Acesso(ISymbol s) => s.DeclaredAccessibility switch
    {
        Accessibility.Public => "public ",
        Accessibility.Protected => "protected ",
        Accessibility.ProtectedOrInternal => "protected ",
        _ => "public ",
    };

    static void EscreverTipo(StringBuilder sb, INamedTypeSymbol t, string ind)
    {
        var usados = Usados[t];
        string nome = t.Name;
        if (t.TypeKind == TypeKind.Enum)
        {
            if (t.GetAttributes().Any(a => a.AttributeClass?.Name == "FlagsAttribute")) sb.AppendLine(ind + "[System.Flags]");
            sb.AppendLine(ind + "public enum " + nome + " : " + N(t.EnumUnderlyingType!) + "\n" + ind + "{");
            foreach (var f in t.GetMembers().OfType<IFieldSymbol>().Where(f => f.HasConstantValue))
                sb.AppendLine(ind + "    " + Esc(f.Name) + " = " + Valor(f.ConstantValue, t.EnumUnderlyingType!) + ",");
            sb.AppendLine(ind + "}");
            return;
        }
        if (t.TypeKind == TypeKind.Delegate)
        {
            var inv = t.DelegateInvokeMethod!;
            sb.AppendLine(ind + "public delegate " + N(inv.ReturnType) + " " + nome + "(" + Params(inv, false) + ");");
            return;
        }
        var bases = new List<string>();
        if (t.TypeKind == TypeKind.Class && t.BaseType != null && t.BaseType.SpecialType != SpecialType.System_Object)
        {
            if (DaApi(t.BaseType) || EhDoSistema(t.BaseType)) bases.Add(N(t.BaseType));
        }
        bool disposable = false, enumeravel = false;
        foreach (var i in t.Interfaces)
        {
            if (DaApi(i)) { if (t.TypeKind == TypeKind.Interface) bases.Add(N(i)); }      // classes: interfaces internas da API não fazem falta
            else if (i.ToDisplayString() == "System.IDisposable" && t.TypeKind != TypeKind.Interface) { bases.Add("System.IDisposable"); disposable = true; }
            else if (i.ToDisplayString() == "System.Collections.IEnumerable" && t.TypeKind != TypeKind.Interface) { bases.Add("System.Collections.IEnumerable"); enumeravel = true; }
        }
        string tipo = t.TypeKind == TypeKind.Interface ? "interface" : t.TypeKind == TypeKind.Struct ? "struct" : "class";
        string mods = t.TypeKind == TypeKind.Class ? (t.IsStatic ? "static " : t.IsAbstract ? "abstract " : t.IsSealed ? "sealed " : "") : "";
        if (t.BaseType != null && t.BaseType.ToDisplayString() == "System.Attribute") sb.AppendLine(ind + "[System.AttributeUsage(System.AttributeTargets.All, AllowMultiple = true)]");
        sb.AppendLine(ind + "public " + mods + tipo + " " + nome + (bases.Count > 0 ? " : " + string.Join(", ", bases) : "") + "\n" + ind + "{");
        string i2 = ind + "    ";
        bool interfaceT = t.TypeKind == TypeKind.Interface;
        bool temCtorSemParametro = false;
        foreach (var s in usados.OrderBy(s => s.Kind).ThenBy(s => s.Name))
        {
            switch (s)
            {
                case IFieldSymbol f:
                    if (f.IsConst) sb.AppendLine(i2 + "public const " + N(f.Type) + " " + Esc(f.Name) + " = " + Valor(f.ConstantValue, f.Type) + ";");
                    else sb.AppendLine(i2 + Acesso(f) + (f.IsStatic ? "static " : "") + (f.IsReadOnly ? "readonly " : "") + N(f.Type) + " " + Esc(f.Name) + ";");
                    break;
                case IPropertySymbol p:
                {
                    string acc = (p.GetMethod != null ? "get" + Corpo(interfaceT) + " " : "") + (p.SetMethod != null ? "set" + (interfaceT ? ";" : " { }") : "");
                    string mod = interfaceT ? "" : Acesso(p) + (p.IsStatic ? "static " : "") + ((p.IsVirtual || p.IsOverride || p.IsAbstract) && !t.IsSealed && t.TypeKind == TypeKind.Class ? "virtual " : "");
                    string nomeP = p.IsIndexer ? "this[" + Params(p.Parameters, false) + "]" : Esc(p.Name);
                    sb.AppendLine(i2 + mod + N(p.Type) + " " + nomeP + " { " + acc + " }");
                    break;
                }
                case IEventSymbol e:
                    sb.AppendLine(i2 + (interfaceT ? "" : Acesso(e) + (e.IsStatic ? "static " : "")) + "event " + N(e.Type) + " " + Esc(e.Name) + ";");
                    break;
                case IMethodSymbol m:
                    if (m.MethodKind == MethodKind.Constructor)
                    {
                        if (m.Parameters.Length == 0) temCtorSemParametro = true;
                        sb.AppendLine(i2 + Acesso(m) + nome + "(" + Params(m, true) + ")" + BaseCtor(t) + " { }");
                    }
                    else if (m.MethodKind == MethodKind.UserDefinedOperator || m.MethodKind == MethodKind.Conversion)
                    {
                        string op = m.MethodKind == MethodKind.Conversion
                            ? (m.Name == "op_Implicit" ? "implicit" : "explicit") + " operator " + N(m.ReturnType)
                            : N(m.ReturnType) + " operator " + Operador(m.Name);
                        sb.AppendLine(i2 + "public static " + op + "(" + Params(m, false) + ") { throw null; }");
                        // C# exige os operadores aos pares
                        string par = m.Name == "op_Equality" ? "op_Inequality" : m.Name == "op_Inequality" ? "op_Equality" : null;
                        if (par != null && !usados.OfType<IMethodSymbol>().Any(x => x.Name == par))
                            sb.AppendLine(i2 + "public static " + N(m.ReturnType) + " operator " + Operador(par) + "(" + Params(m, false) + ") { throw null; }");
                    }
                    else if (m.MethodKind == MethodKind.Ordinary || m.MethodKind == MethodKind.EventAdd)
                    {
                        string mod = interfaceT ? "" : Acesso(m) + (m.IsStatic ? "static " : "") + ((m.IsVirtual || m.IsOverride || m.IsAbstract) && !t.IsSealed && t.TypeKind == TypeKind.Class ? "virtual " : "");
                        string corpo = interfaceT ? ";" : " { throw null; }";
                        sb.AppendLine(i2 + mod + N(m.ReturnType) + " " + Esc(m.Name) + "(" + Params(m, true) + ")" + corpo);
                    }
                    break;
            }
        }
        if (t.TypeKind == TypeKind.Class && !t.IsStatic && !temCtorSemParametro) sb.AppendLine(i2 + "protected " + nome + "()" + BaseCtor(t) + " { }");
        if (disposable && !usados.OfType<IMethodSymbol>().Any(m => m.Name == "Dispose" && m.Parameters.Length == 0) && !HerdaUsado(t, "Dispose"))
            sb.AppendLine(i2 + "void System.IDisposable.Dispose() { }");
        if (enumeravel && !usados.OfType<IMethodSymbol>().Any(m => m.Name == "GetEnumerator" && m.Parameters.Length == 0 && m.ReturnType.ToDisplayString() == "System.Collections.IEnumerator"))
            sb.AppendLine(i2 + "System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() { throw null; }");
        foreach (var n in Usados.Keys.Where(x => SymbolEqualityComparer.Default.Equals(x.ContainingType, t)).OrderBy(x => x.Name)) EscreverTipo(sb, n, i2);
        sb.AppendLine(ind + "}");
    }

    static bool HerdaUsado(INamedTypeSymbol t, string nome)
    {
        for (var b = t.BaseType; b != null; b = b.BaseType)
            if (Usados.TryGetValue(b, out var u) && u.OfType<IMethodSymbol>().Any(m => m.Name == nome && m.Parameters.Length == 0)) return true;
        return false;
    }

    static bool EhDoSistema(INamedTypeSymbol t) => t.ContainingNamespace.ToDisplayString().StartsWith("System");

    static string BaseCtor(INamedTypeSymbol t) => "";

    static string Corpo(bool interfaceT) => interfaceT ? ";" : " { throw null; }";

    static string Params(IMethodSymbol m, bool padroes) => Params(m.Parameters, padroes);

    static string Params(System.Collections.Immutable.ImmutableArray<IParameterSymbol> ps, bool padroes)
    {
        return string.Join(", ", ps.Select(p =>
        {
            string r = p.RefKind == RefKind.Out ? "out " : p.RefKind == RefKind.Ref ? "ref " : "";
            if (p.IsParams) r = "params " + r;
            string s = r + N(p.Type) + " " + Esc(p.Name);
            if (padroes && p.HasExplicitDefaultValue) s += " = " + Valor(p.ExplicitDefaultValue, p.Type);
            return s;
        }));
    }

    static string Valor(object v, ITypeSymbol t)
    {
        if (v == null) return t.IsValueType ? "default(" + N(t) + ")" : "null";
        if (t.TypeKind == TypeKind.Enum && t is INamedTypeSymbol en && SymbolEqualityComparer.Default.Equals(en.EnumUnderlyingType, t) == false && !(v is Enum))
            return "(" + N(t) + ")(" + Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture) + ")";
        if (v is bool b) return b ? "true" : "false";
        if (v is string s) return SymbolDisplay.FormatLiteral(s, true);
        if (v is char c) return SymbolDisplay.FormatLiteral(c, true);
        if (v is double d) return d.ToString("R", System.Globalization.CultureInfo.InvariantCulture) + "d";
        if (v is float f) return f.ToString("R", System.Globalization.CultureInfo.InvariantCulture) + "f";
        string num = Convert.ToString(v, System.Globalization.CultureInfo.InvariantCulture)!;
        if (v is uint) num += "u"; else if (v is long) num += "L"; else if (v is ulong) num += "UL";
        return "unchecked((" + N(t) + ")(" + num + "))";
    }

    static string Esc(string n) => SyntaxFacts.GetKeywordKind(n) != SyntaxKind.None ? "@" + n : n;

    static string Operador(string n) => n switch
    {
        "op_Addition" => "+", "op_Subtraction" => "-", "op_Multiply" => "*", "op_Division" => "/", "op_Equality" => "==", "op_Inequality" => "!=",
        "op_UnaryNegation" => "-", "op_UnaryPlus" => "+", "op_LessThan" => "<", "op_GreaterThan" => ">", "op_LessThanOrEqual" => "<=", "op_GreaterThanOrEqual" => ">=",
        "op_Modulus" => "%", "op_BitwiseAnd" => "&", "op_BitwiseOr" => "|", "op_ExclusiveOr" => "^", "op_LogicalNot" => "!", "op_True" => "true", "op_False" => "false",
        _ => throw new NotSupportedException(n),
    };
}
