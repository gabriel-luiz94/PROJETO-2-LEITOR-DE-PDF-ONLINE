// GERADO por ferramentas/GerarApiNanoCAD5 — não editar à mão. Só as declarações que o plugin usa da API do nanoCAD (hostdbmgd),
// para compilar a DLL do nanoCAD 5 (.NET 4). Em tempo de execução valem as DLLs do próprio nanoCAD.
#pragma warning disable 67, 108, 114, 169, 649, 659, 660, 661, 809
namespace Teigha.Colors
{
    public sealed class Color : global::Teigha.Runtime.DisposableWrapper
    {
        public static global::Teigha.Colors.Color FromColorIndex(global::Teigha.Colors.ColorMethod colorMethod, short colorIndex) { throw null; }
        public static bool operator ==(global::Teigha.Colors.Color a, global::Teigha.Colors.Color b) { throw null; }
        public static bool operator !=(global::Teigha.Colors.Color a, global::Teigha.Colors.Color b) { throw null; }
        public short ColorIndex { get { throw null; }  }
        public global::Teigha.Colors.ColorMethod ColorMethod { get { throw null; }  }
        public global::System.Drawing.Color ColorValue { get { throw null; }  }
        protected Color() { }
    }
    public enum ColorMethod : int
    {
        ByLayer = unchecked((int)(192)),
        ByBlock = unchecked((int)(193)),
        ByColor = unchecked((int)(194)),
        ByAci = unchecked((int)(195)),
        ByPen = unchecked((int)(196)),
        Foreground = unchecked((int)(197)),
        LayerOff = unchecked((int)(198)),
        LayerFrozen = unchecked((int)(199)),
        None = unchecked((int)(200)),
    }
}
namespace Teigha.DatabaseServices
{
    public abstract class AbstractViewTableRecord : global::Teigha.DatabaseServices.SymbolTableRecord
    {
        public global::Teigha.Geometry.Point2d CenterPoint { get { throw null; } set { } }
        public double Height { get { throw null; } set { } }
        public double Width { get { throw null; } set { } }
        protected AbstractViewTableRecord() { }
    }
    public sealed class AttributeCollection : System.Collections.IEnumerable
    {
        public global::System.Collections.IEnumerator GetEnumerator() { throw null; }
        protected AttributeCollection() { }
    }
    public class AttributeDefinition : global::Teigha.DatabaseServices.DBText
    {
        protected AttributeDefinition() { }
    }
    public class AttributeReference : global::Teigha.DatabaseServices.DBText
    {
        public bool Invisible { get { throw null; } set { } }
        protected AttributeReference() { }
    }
    public class BlockReference : global::Teigha.DatabaseServices.Entity
    {
        public global::Teigha.DatabaseServices.AttributeCollection AttributeCollection { get { throw null; }  }
        public virtual global::Teigha.DatabaseServices.ObjectId BlockTableRecord { get { throw null; } set { } }
        public global::Teigha.DatabaseServices.ObjectId DynamicBlockTableRecord { get { throw null; }  }
        public bool IsDynamicBlock { get { throw null; }  }
        public string Name { get { throw null; }  }
        public global::Teigha.Geometry.Point3d Position { get { throw null; } set { } }
        protected BlockReference() { }
    }
    public class BlockTable : global::Teigha.DatabaseServices.SymbolTable, System.Collections.IEnumerable
    {
        protected BlockTable() { }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() { throw null; }
    }
    public class BlockTableRecord : global::Teigha.DatabaseServices.SymbolTableRecord, System.Collections.IEnumerable
    {
        public static readonly string ModelSpace;
        public virtual global::Teigha.DatabaseServices.BlockTableRecordEnumerator GetEnumerator() { throw null; }
        protected BlockTableRecord() { }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() { throw null; }
    }
    public sealed class BlockTableRecordEnumerator : global::Teigha.Runtime.DisposableWrapper
    {
        public bool MoveNext() { throw null; }
        public global::Teigha.DatabaseServices.ObjectId Current { get { throw null; }  }
        protected BlockTableRecordEnumerator() { }
    }
    public class Circle : global::Teigha.DatabaseServices.Curve
    {
        public Circle(global::Teigha.Geometry.Point3d center, global::Teigha.Geometry.Vector3d normal, double radius) { }
        public global::Teigha.Geometry.Point3d Center { get { throw null; } set { } }
        public double Radius { get { throw null; } set { } }
        protected Circle() { }
    }
    public enum ContentType : int
    {
        NoneContent = unchecked((int)(0)),
        BlockContent = unchecked((int)(1)),
        MTextContent = unchecked((int)(2)),
        ToleranceContent = unchecked((int)(3)),
    }
    public abstract class Curve : global::Teigha.DatabaseServices.Entity
    {
        protected Curve() { }
    }
    public sealed class Database : global::Teigha.Runtime.RXObject
    {
        public global::Teigha.DatabaseServices.ObjectId GetObjectId(bool createIfNotFound, global::Teigha.DatabaseServices.Handle objHandle, int identifier) { throw null; }
        public global::Teigha.DatabaseServices.ObjectId BlockTableId { get { throw null; }  }
        public global::Teigha.DatabaseServices.ObjectId LayerTableId { get { throw null; }  }
        public global::Teigha.DatabaseServices.ObjectId NamedObjectsDictionaryId { get { throw null; }  }
        public global::Teigha.DatabaseServices.TransactionManager TransactionManager { get { throw null; }  }
        protected Database() { }
    }
    public class DBDictionary : global::Teigha.DatabaseServices.DBObject
    {
        public DBDictionary() { }
        public bool Contains(string entryName) { throw null; }
        public virtual global::Teigha.DatabaseServices.ObjectId GetAt(string entryName) { throw null; }
        public global::Teigha.DatabaseServices.ObjectId SetAt(string searchKey, global::Teigha.DatabaseServices.DBObject newValue) { throw null; }
    }
    public abstract class DBObject : global::Teigha.GraphicsInterface.Drawable
    {
        public void UpgradeOpen() { throw null; }
        public global::Teigha.DatabaseServices.Handle Handle { get { throw null; }  }
        protected DBObject() { }
    }
    public class DBText : global::Teigha.DatabaseServices.Entity
    {
        public DBText() { }
        public double Height { get { throw null; } set { } }
        public global::Teigha.Geometry.Point3d Position { get { throw null; } set { } }
        public string TextString { get { throw null; } set { } }
    }
    public enum DxfCode : int
    {
        Alpha = unchecked((int)(440)),
        Angle = unchecked((int)(50)),
        ArbitraryHandle = unchecked((int)(320)),
        AttributePrompt = unchecked((int)(3)),
        AttributeTag = unchecked((int)(2)),
        BinaryChunk = unchecked((int)(310)),
        BlockName = unchecked((int)(2)),
        Bool = unchecked((int)(290)),
        CircleSides = unchecked((int)(72)),
        CLShapeName = unchecked((int)(4)),
        CLShapeText = unchecked((int)(9)),
        Color = unchecked((int)(62)),
        ColorName = unchecked((int)(430)),
        ColorRgb = unchecked((int)(420)),
        Comment = unchecked((int)(999)),
        ControlString = unchecked((int)(102)),
        DashLength = unchecked((int)(49)),
        Description = unchecked((int)(3)),
        DimBlk1 = unchecked((int)(6)),
        DimBlk2 = unchecked((int)(7)),
        DimensionAlternativePrefixSuffix = unchecked((int)(4)),
        DimensionBlock = unchecked((int)(5)),
        DimPostString = unchecked((int)(3)),
        DimStyleName = unchecked((int)(3)),
        DimVarHandle = unchecked((int)(105)),
        Elevation = unchecked((int)(38)),
        EmbeddedObjectStart = unchecked((int)(101)),
        End = unchecked((int)(-1)),
        ExtendedDataAsciiString = unchecked((int)(1000)),
        ExtendedDataBinaryChunk = unchecked((int)(1004)),
        ExtendedDataControlString = unchecked((int)(1002)),
        ExtendedDataDist = unchecked((int)(1041)),
        ExtendedDataHandle = unchecked((int)(1005)),
        ExtendedDataInteger16 = unchecked((int)(1070)),
        ExtendedDataInteger32 = unchecked((int)(1071)),
        ExtendedDataLayerName = unchecked((int)(1003)),
        ExtendedDataReal = unchecked((int)(1040)),
        ExtendedDataRegAppName = unchecked((int)(1001)),
        ExtendedDataScale = unchecked((int)(1042)),
        ExtendedDataWorldXCoordinate = unchecked((int)(1011)),
        ExtendedDataWorldXDir = unchecked((int)(1013)),
        ExtendedDataWorldXDisp = unchecked((int)(1012)),
        ExtendedDataWorldYCoordinate = unchecked((int)(1021)),
        ExtendedDataWorldYDir = unchecked((int)(1023)),
        ExtendedDataWorldYDisp = unchecked((int)(1022)),
        ExtendedDataWorldZCoordinate = unchecked((int)(1031)),
        ExtendedDataWorldZDir = unchecked((int)(1033)),
        ExtendedDataWorldZDisp = unchecked((int)(1032)),
        ExtendedDataXCoordinate = unchecked((int)(1010)),
        ExtendedDataYCoordinate = unchecked((int)(1020)),
        ExtendedDataZCoordinate = unchecked((int)(1030)),
        ExtendedInt16 = unchecked((int)(400)),
        FirstEntityId = unchecked((int)(-2)),
        GradientAngle = unchecked((int)(460)),
        GradientColCount = unchecked((int)(453)),
        GradientColVal = unchecked((int)(463)),
        GradientName = unchecked((int)(470)),
        GradientObjType = unchecked((int)(450)),
        GradientPatType = unchecked((int)(451)),
        GradientShift = unchecked((int)(461)),
        GradientTintType = unchecked((int)(452)),
        GradientTintVal = unchecked((int)(462)),
        Handle = unchecked((int)(5)),
        HardOwnershipId = unchecked((int)(360)),
        HardPointerId = unchecked((int)(340)),
        HasSubentities = unchecked((int)(66)),
        HeaderId = unchecked((int)(-2)),
        Int16 = unchecked((int)(70)),
        Int32 = unchecked((int)(90)),
        Int8 = unchecked((int)(280)),
        Invalid = unchecked((int)(-9999)),
        LayerLinetype = unchecked((int)(61)),
        LayerName = unchecked((int)(8)),
        LayoutName = unchecked((int)(410)),
        LinetypeAlign = unchecked((int)(72)),
        LinetypeElement = unchecked((int)(49)),
        LinetypeName = unchecked((int)(6)),
        LinetypePdc = unchecked((int)(73)),
        LinetypeProse = unchecked((int)(3)),
        LinetypeScale = unchecked((int)(48)),
        LineWeight = unchecked((int)(370)),
        MlineOffset = unchecked((int)(49)),
        MlineStyleName = unchecked((int)(2)),
        NormalX = unchecked((int)(210)),
        NormalY = unchecked((int)(220)),
        NormalZ = unchecked((int)(230)),
        Operator = unchecked((int)(-4)),
        PixelScale = unchecked((int)(47)),
        PlotStyleNameId = unchecked((int)(390)),
        PlotStyleNameType = unchecked((int)(380)),
        PReactors = unchecked((int)(-5)),
        Real = unchecked((int)(40)),
        RegAppFlags = unchecked((int)(71)),
        RenderMode = unchecked((int)(281)),
        ShapeName = unchecked((int)(2)),
        ShapeScale = unchecked((int)(46)),
        ShapeXOffset = unchecked((int)(44)),
        ShapeYOffset = unchecked((int)(45)),
        SoftOwnershipId = unchecked((int)(350)),
        SoftPointerId = unchecked((int)(330)),
        Start = unchecked((int)(0)),
        Subclass = unchecked((int)(100)),
        SymbolTableName = unchecked((int)(2)),
        SymbolTableRecordComments = unchecked((int)(4)),
        SymbolTableRecordName = unchecked((int)(2)),
        Text = unchecked((int)(1)),
        TextBigFontFile = unchecked((int)(4)),
        TextFontFile = unchecked((int)(3)),
        TextStyleName = unchecked((int)(7)),
        Thickness = unchecked((int)(39)),
        TxtSize = unchecked((int)(40)),
        TxtStyleFlags = unchecked((int)(71)),
        TxtStylePSize = unchecked((int)(42)),
        TxtStyleXScale = unchecked((int)(41)),
        UcsOrg = unchecked((int)(110)),
        UcsOrientationX = unchecked((int)(111)),
        UcsOrientationY = unchecked((int)(112)),
        ViewBackClip = unchecked((int)(44)),
        ViewFrontClip = unchecked((int)(43)),
        ViewHeight = unchecked((int)(45)),
        ViewLensLength = unchecked((int)(42)),
        ViewMode = unchecked((int)(71)),
        ViewportActive = unchecked((int)(68)),
        ViewportAspect = unchecked((int)(41)),
        ViewportGrid = unchecked((int)(76)),
        ViewportHeight = unchecked((int)(40)),
        ViewportIcon = unchecked((int)(74)),
        ViewportNumber = unchecked((int)(69)),
        ViewportSnap = unchecked((int)(75)),
        ViewportSnapAngle = unchecked((int)(50)),
        ViewportSnapPair = unchecked((int)(78)),
        ViewportSnapStyle = unchecked((int)(77)),
        ViewportTwist = unchecked((int)(51)),
        ViewportVisibility = unchecked((int)(67)),
        ViewportZoom = unchecked((int)(73)),
        ViewWidth = unchecked((int)(41)),
        Visibility = unchecked((int)(60)),
        XCoordinate = unchecked((int)(10)),
        XDataStart = unchecked((int)(-3)),
        XDictionary = unchecked((int)(-6)),
        XInt16 = unchecked((int)(170)),
        XReal = unchecked((int)(140)),
        XRefPath = unchecked((int)(1)),
        XTextString = unchecked((int)(300)),
        XXInt16 = unchecked((int)(270)),
        YCoordinate = unchecked((int)(20)),
        ZCoordinate = unchecked((int)(30)),
    }
    public abstract class Entity : global::Teigha.DatabaseServices.DBObject
    {
        public global::Teigha.Colors.Color Color { get { throw null; } set { } }
        public global::Teigha.DatabaseServices.Extents3d GeometricExtents { get { throw null; }  }
        public string Layer { get { throw null; } set { } }
        protected Entity() { }
    }
    public struct Extents3d
    {
        public global::Teigha.Geometry.Point3d MaxPoint { get { throw null; }  }
        public global::Teigha.Geometry.Point3d MinPoint { get { throw null; }  }
    }
    public struct Handle
    {
        public Handle(long v) { }
        public string ToString() { throw null; }
    }
    public sealed class LayerTable : global::Teigha.DatabaseServices.SymbolTable
    {
        public global::Teigha.DatabaseServices.SymbolTableEnumerator GetEnumerator() { throw null; }
        protected LayerTable() { }
    }
    public class LayerTableRecord : global::Teigha.DatabaseServices.SymbolTableRecord
    {
        public LayerTableRecord() { }
        public global::Teigha.Colors.Color Color { get { throw null; } set { } }
    }
    public class Leader : global::Teigha.DatabaseServices.Curve
    {
        public virtual global::Teigha.Geometry.Point3d VertexAt(int value) { throw null; }
        public virtual int NumVertices { get { throw null; }  }
        protected Leader() { }
    }
    public class Line : global::Teigha.DatabaseServices.Curve
    {
        public Line(global::Teigha.Geometry.Point3d stPoint, global::Teigha.Geometry.Point3d endPoint) { }
        public virtual global::Teigha.Geometry.Point3d EndPoint { get { throw null; } set { } }
        public virtual global::Teigha.Geometry.Point3d StartPoint { get { throw null; } set { } }
        protected Line() { }
    }
    public class MLeader : global::Teigha.DatabaseServices.Entity
    {
        public global::Teigha.DatabaseServices.Extents3d GetContentGeomExtents() { throw null; }
        public global::System.Collections.ArrayList GetLeaderIndexes() { throw null; }
        public global::System.Collections.ArrayList GetLeaderLineIndexes(int leaderIndex) { throw null; }
        public global::Teigha.Geometry.Point3d GetVertex(int leaderLineIndex, int index) { throw null; }
        public int VerticesCount(int leaderLineIndex) { throw null; }
        public global::Teigha.DatabaseServices.ContentType ContentType { get { throw null; } set { } }
        public global::Teigha.DatabaseServices.MText MText { get { throw null; } set { } }
        protected MLeader() { }
    }
    public class MText : global::Teigha.DatabaseServices.Entity
    {
        public string Contents { get { throw null; } set { } }
        public global::Teigha.Geometry.Point3d Location { get { throw null; } set { } }
        public double TextHeight { get { throw null; } set { } }
        protected MText() { }
    }
    public struct ObjectId
    {
        public string ToString() { throw null; }
        public bool IsErased { get { throw null; }  }
        public bool IsNull { get { throw null; }  }
    }
    public enum OpenMode : int
    {
        ForRead = unchecked((int)(0)),
        ForWrite = unchecked((int)(1)),
        ForNotify = unchecked((int)(2)),
    }
    public class Polyline : global::Teigha.DatabaseServices.Curve
    {
        public Polyline() { }
        public void AddVertexAt(int index, global::Teigha.Geometry.Point2d pt, double bulge, double startWidth, double endWidth) { throw null; }
        public global::Teigha.Geometry.Point2d GetPoint2dAt(int index) { throw null; }
        public bool Closed { get { throw null; } set { } }
        public int NumberOfVertices { get { throw null; }  }
    }
    public class Polyline2d : global::Teigha.DatabaseServices.Curve, System.Collections.IEnumerable
    {
        public virtual global::System.Collections.IEnumerator GetEnumerator() { throw null; }
        public bool Closed { get { throw null; } set { } }
        protected Polyline2d() { }
    }
    public class Polyline3d : global::Teigha.DatabaseServices.Curve, System.Collections.IEnumerable
    {
        public virtual global::System.Collections.IEnumerator GetEnumerator() { throw null; }
        public bool Closed { get { throw null; } set { } }
        protected Polyline3d() { }
    }
    public class PolylineVertex3d : global::Teigha.DatabaseServices.Vertex
    {
        public global::Teigha.Geometry.Point3d Position { get { throw null; } set { } }
        protected PolylineVertex3d() { }
    }
    public sealed class ResultBuffer : global::Teigha.Runtime.DisposableWrapper, System.Collections.IEnumerable
    {
        public ResultBuffer() { }
        public void Add(global::Teigha.DatabaseServices.TypedValue value) { throw null; }
        public global::Teigha.DatabaseServices.ResultBufferEnumerator GetEnumerator() { throw null; }
        public static bool operator !=(global::Teigha.DatabaseServices.ResultBuffer a, global::Teigha.DatabaseServices.ResultBuffer b) { throw null; }
        public static bool operator ==(global::Teigha.DatabaseServices.ResultBuffer a, global::Teigha.DatabaseServices.ResultBuffer b) { throw null; }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() { throw null; }
    }
    public sealed class ResultBufferEnumerator
    {
        public bool MoveNext() { throw null; }
        public global::Teigha.DatabaseServices.TypedValue Current { get { throw null; }  }
        protected ResultBufferEnumerator() { }
    }
    public abstract class SymbolTable : global::Teigha.DatabaseServices.DBObject, System.Collections.IEnumerable
    {
        public global::Teigha.DatabaseServices.ObjectId Add(global::Teigha.DatabaseServices.SymbolTableRecord value) { throw null; }
        public bool Has(string key) { throw null; }
        public global::Teigha.DatabaseServices.ObjectId this[string key] { get { throw null; }  }
        protected SymbolTable() { }
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() { throw null; }
    }
    public class SymbolTableEnumerator : global::Teigha.Runtime.DisposableWrapper
    {
        public virtual bool MoveNext() { throw null; }
        public virtual global::Teigha.DatabaseServices.ObjectId Current { get { throw null; }  }
        protected SymbolTableEnumerator() { }
    }
    public abstract class SymbolTableRecord : global::Teigha.DatabaseServices.DBObject
    {
        public string Name { get { throw null; } set { } }
        protected SymbolTableRecord() { }
    }
    public class Transaction : global::Teigha.Runtime.RXObject
    {
        public virtual void AddNewlyCreatedDBObject(global::Teigha.DatabaseServices.DBObject obj, bool add) { throw null; }
        public virtual void Commit() { throw null; }
        public virtual global::Teigha.DatabaseServices.DBObject GetObject(global::Teigha.DatabaseServices.ObjectId id, global::Teigha.DatabaseServices.OpenMode mode) { throw null; }
        protected Transaction() { }
    }
    public class TransactionManager : global::Teigha.Runtime.RXObject
    {
        public virtual global::Teigha.DatabaseServices.Transaction StartTransaction() { throw null; }
        protected TransactionManager() { }
    }
    public struct TypedValue
    {
        public TypedValue(int typeCode, object value) { }
        public object Value { get { throw null; }  }
    }
    public class Vertex : global::Teigha.DatabaseServices.Entity
    {
        protected Vertex() { }
    }
    public class Vertex2d : global::Teigha.DatabaseServices.Vertex
    {
        public global::Teigha.Geometry.Point3d Position { get { throw null; } set { } }
        protected Vertex2d() { }
    }
    public class ViewTableRecord : global::Teigha.DatabaseServices.AbstractViewTableRecord
    {
        protected ViewTableRecord() { }
    }
    public sealed class Xrecord : global::Teigha.DatabaseServices.DBObject
    {
        public Xrecord() { }
        public global::Teigha.DatabaseServices.ResultBuffer Data { get { throw null; } set { } }
    }
}
namespace Teigha.Geometry
{
    public class IntegerCollection
    {
        public IntegerCollection() { }
    }
    public struct Point2d
    {
        public Point2d(double x, double y) { }
        public double X { get { throw null; }  }
        public double Y { get { throw null; }  }
    }
    public struct Point3d
    {
        public Point3d(double x, double y, double z) { }
        public double DistanceTo(global::Teigha.Geometry.Point3d point) { throw null; }
        public static global::Teigha.Geometry.Point3d operator +(global::Teigha.Geometry.Point3d a, global::Teigha.Geometry.Vector3d vector) { throw null; }
        public static global::Teigha.Geometry.Vector3d operator -(global::Teigha.Geometry.Point3d a, global::Teigha.Geometry.Point3d b) { throw null; }
        public double X { get { throw null; }  }
        public double Y { get { throw null; }  }
    }
    public struct Vector3d
    {
        public static global::Teigha.Geometry.Vector3d operator /(global::Teigha.Geometry.Vector3d a, double factor) { throw null; }
        public static global::Teigha.Geometry.Vector3d operator *(global::Teigha.Geometry.Vector3d a, double factor) { throw null; }
        public static global::Teigha.Geometry.Vector3d ZAxis { get { throw null; }  }
    }
}
namespace Teigha.GraphicsInterface
{
    public abstract class Drawable : global::Teigha.Runtime.RXObject
    {
        protected Drawable() { }
    }
    public enum TransientDrawingMode : int
    {
        Contrast = unchecked((int)(5)),
        DirectShortTerm = unchecked((int)(2)),
        DirectTopmost = unchecked((int)(4)),
        Highlight = unchecked((int)(3)),
        Main = unchecked((int)(0)),
        Sprite = unchecked((int)(1)),
        Count = unchecked((int)(6)),
    }
    public class TransientManager : global::Teigha.Runtime.DisposableWrapper
    {
        public bool AddTransient(global::Teigha.GraphicsInterface.Drawable added, global::Teigha.GraphicsInterface.TransientDrawingMode mode, int subDrawingMode, global::Teigha.Geometry.IntegerCollection viewportNumbers) { throw null; }
        public bool EraseTransient(global::Teigha.GraphicsInterface.Drawable erased, global::Teigha.Geometry.IntegerCollection viewportNumbers) { throw null; }
        public static global::Teigha.GraphicsInterface.TransientManager CurrentTransientManager { get { throw null; } set { } }
        protected TransientManager() { }
    }
}
namespace Teigha.Runtime
{
    [System.Flags]
    public enum CommandFlags : int
    {
        Modal = unchecked((int)(0)),
        Transparent = unchecked((int)(1)),
        UsePickSet = unchecked((int)(2)),
        Redraw = unchecked((int)(4)),
        NoPerspective = unchecked((int)(8)),
        NoMultiple = unchecked((int)(16)),
        NoTileMode = unchecked((int)(32)),
        NoPaperSpace = unchecked((int)(64)),
        PlotOnly = unchecked((int)(128)),
        NoOem = unchecked((int)(256)),
        Undefined = unchecked((int)(512)),
        InProgress = unchecked((int)(1024)),
        Defun = unchecked((int)(2048)),
        NoNewStack = unchecked((int)(65536)),
        NoInternalLock = unchecked((int)(131072)),
        DocReadLock = unchecked((int)(524288)),
        DocExclusiveLock = unchecked((int)(1048576)),
        Session = unchecked((int)(2097152)),
        Interruptible = unchecked((int)(4194304)),
        NoHistory = unchecked((int)(8388608)),
        NoUndoMarker = unchecked((int)(16777216)),
        NoBlockEditor = unchecked((int)(33554432)),
        NoActionRecording = unchecked((int)(67108864)),
        ActionMacro = unchecked((int)(134217728)),
    }
    [System.AttributeUsage(System.AttributeTargets.All, AllowMultiple = true)]
    public sealed class CommandMethodAttribute : global::System.Attribute
    {
        public CommandMethodAttribute(string globalName) { }
        protected CommandMethodAttribute() { }
    }
    public abstract class DisposableWrapper : global::System.MarshalByRefObject, System.IDisposable
    {
        public virtual void Dispose() { throw null; }
        public static bool operator ==(global::Teigha.Runtime.DisposableWrapper a, global::Teigha.Runtime.DisposableWrapper b) { throw null; }
        public static bool operator !=(global::Teigha.Runtime.DisposableWrapper a, global::Teigha.Runtime.DisposableWrapper b) { throw null; }
        protected DisposableWrapper() { }
    }
    [System.AttributeUsage(System.AttributeTargets.All, AllowMultiple = true)]
    public sealed class ExtensionApplicationAttribute : global::System.Attribute
    {
        public ExtensionApplicationAttribute(global::System.Type entryPointType) { }
        protected ExtensionApplicationAttribute() { }
    }
    public interface ICommandLineCallable
    {
        global::System.Type ContextMenuExtensionType { get;  }
        global::Teigha.Runtime.CommandFlags Flags { get;  }
        string GlobalName { get;  }
        string GroupName { get;  }
        string HelpFileName { get;  }
        string HelpTopic { get;  }
        string LocalizedNameId { get;  }
    }
    public interface IExtensionApplication
    {
        void Initialize();
        void Terminate();
    }
    public abstract class RXObject : global::Teigha.Runtime.DisposableWrapper
    {
        protected RXObject() { }
    }
}
namespace OrcamentoNanoCAD.Api5 { internal static class Vazio { } }
