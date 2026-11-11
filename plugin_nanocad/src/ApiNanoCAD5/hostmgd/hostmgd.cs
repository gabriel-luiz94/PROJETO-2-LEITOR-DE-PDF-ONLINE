// GERADO por ferramentas/GerarApiNanoCAD5 — não editar à mão. Só as declarações que o plugin usa da API do nanoCAD (hostmgd),
// para compilar a DLL do nanoCAD 5 (.NET 4). Em tempo de execução valem as DLLs do próprio nanoCAD.
#pragma warning disable 67, 108, 114, 169, 649, 659, 660, 661, 809
namespace HostMgd.ApplicationServices
{
    public sealed class Application
    {
        public static global::System.Windows.Forms.DialogResult ShowModalDialog(global::System.Windows.Forms.Form formToShow) { throw null; }
        public static void ShowModelessDialog(global::System.Windows.Forms.Form formToShow) { throw null; }
        public static global::HostMgd.ApplicationServices.DocumentCollection DocumentManager { get { throw null; }  }
        protected Application() { }
    }
    public sealed class Document : global::Teigha.Runtime.DisposableWrapper
    {
        public global::HostMgd.ApplicationServices.DocumentLock LockDocument() { throw null; }
        public void SendStringToExecute(string command, bool activate, bool wrapUpInactiveDoc, bool echoCommand) { throw null; }
        public string CommandInProgress { get { throw null; }  }
        public global::Teigha.DatabaseServices.Database Database { get { throw null; }  }
        public global::HostMgd.EditorInput.Editor Editor { get { throw null; }  }
        public string Name { get { throw null; }  }
        public global::System.Collections.Hashtable UserData { get { throw null; }  }
        protected Document() { }
    }
    public sealed class DocumentCollection : global::System.MarshalByRefObject, System.IDisposable
    {
        public global::HostMgd.ApplicationServices.Document MdiActiveDocument { get { throw null; } set { } }
        protected DocumentCollection() { }
        void System.IDisposable.Dispose() { }
    }
    public sealed class DocumentLock : System.IDisposable
    {
        protected DocumentLock() { }
        void System.IDisposable.Dispose() { }
    }
}
namespace HostMgd.EditorInput
{
    public sealed class Editor : global::System.MarshalByRefObject, System.IDisposable
    {
        public void DrawVector(global::Teigha.Geometry.Point3d from, global::Teigha.Geometry.Point3d to, int color, bool drawHighlighted) { throw null; }
        public global::Teigha.DatabaseServices.ViewTableRecord GetCurrentView() { throw null; }
        public global::HostMgd.EditorInput.PromptEntityResult GetEntity(global::HostMgd.EditorInput.PromptEntityOptions options) { throw null; }
        public global::HostMgd.EditorInput.PromptEntityResult GetEntity(string message) { throw null; }
        public global::HostMgd.EditorInput.PromptResult GetKeywords(global::HostMgd.EditorInput.PromptKeywordOptions options) { throw null; }
        public global::HostMgd.EditorInput.PromptSelectionResult GetSelection(global::HostMgd.EditorInput.PromptSelectionOptions options, global::HostMgd.EditorInput.SelectionFilter filter) { throw null; }
        public global::HostMgd.EditorInput.PromptResult GetString(global::HostMgd.EditorInput.PromptStringOptions options) { throw null; }
        public void Regen() { throw null; }
        public void SetCurrentView(global::Teigha.DatabaseServices.ViewTableRecord value) { throw null; }
        public void WriteMessage(string message) { throw null; }
        protected Editor() { }
        void System.IDisposable.Dispose() { }
    }
    public sealed class KeywordCollection : System.Collections.IEnumerable
    {
        public string Default { get { throw null; } set { } }
        protected KeywordCollection() { }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() { throw null; }
    }
    public abstract class PromptEditorOptions : global::HostMgd.EditorInput.PromptOptions
    {
        protected PromptEditorOptions() { }
    }
    public sealed class PromptEntityOptions : global::HostMgd.EditorInput.PromptEditorOptions
    {
        public PromptEntityOptions(string message) { }
        public bool AllowNone { get { throw null; } set { } }
        protected PromptEntityOptions() { }
    }
    public class PromptEntityResult : global::HostMgd.EditorInput.PromptResult
    {
        public global::Teigha.DatabaseServices.ObjectId ObjectId { get { throw null; }  }
        public global::Teigha.Geometry.Point3d PickedPoint { get { throw null; }  }
        protected PromptEntityResult() { }
    }
    public sealed class PromptKeywordOptions : global::HostMgd.EditorInput.PromptEditorOptions
    {
        public PromptKeywordOptions(string messageAndKeywords, string globalKeywords) { }
        public bool AllowNone { get { throw null; } set { } }
        protected PromptKeywordOptions() { }
    }
    public abstract class PromptOptions
    {
        public global::HostMgd.EditorInput.KeywordCollection Keywords { get { throw null; }  }
        protected PromptOptions() { }
    }
    public class PromptResult
    {
        public global::HostMgd.EditorInput.PromptStatus Status { get { throw null; }  }
        public string StringResult { get { throw null; } set { } }
        protected PromptResult() { }
    }
    public sealed class PromptSelectionOptions
    {
        public PromptSelectionOptions() { }
        public string MessageForAdding { get { throw null; } set { } }
    }
    public sealed class PromptSelectionResult
    {
        public global::HostMgd.EditorInput.PromptStatus Status { get { throw null; }  }
        public global::HostMgd.EditorInput.SelectionSet Value { get { throw null; }  }
        protected PromptSelectionResult() { }
    }
    public enum PromptStatus : int
    {
        Cancel = unchecked((int)(-5002)),
        Error = unchecked((int)(-5001)),
        Keyword = unchecked((int)(-5005)),
        Modeless = unchecked((int)(5027)),
        None = unchecked((int)(5000)),
        OK = unchecked((int)(5100)),
        Other = unchecked((int)(5028)),
    }
    public sealed class PromptStringOptions : global::HostMgd.EditorInput.PromptEditorOptions
    {
        public PromptStringOptions(string message) { }
        public bool AllowSpaces { get { throw null; } set { } }
        protected PromptStringOptions() { }
    }
    public class SelectionFilter
    {
        public SelectionFilter(global::Teigha.DatabaseServices.TypedValue[] value) { }
        protected SelectionFilter() { }
    }
    public class SelectionSet
    {
        public virtual global::Teigha.DatabaseServices.ObjectId[] GetObjectIds() { throw null; }
        protected SelectionSet() { }
    }
}
