using System;
using System.Collections.Generic;
using Teigha.DatabaseServices;
using Teigha.GraphicsInterface;
using Teigha.Runtime;

namespace OrcamentoNanoCAD
{
    // Filtra o desenho da anotação na viewport. O estado da layer no arquivo não é alterado.
    internal sealed class VisibilidadeDeNotas : IDisposable
    {
#if !NANOCAD5
        private Ocultar _filtro;
        private RXClass _classe;

        private sealed class Ocultar : DrawableOverrule
        {
            public Ocultar(ObjectId[] ids) { SetIdFilter(ids); }
            public override bool WorldDraw(Drawable drawable, WorldDraw worldDraw) { return true; }
            public override void ViewportDraw(Drawable drawable, ViewportDraw viewportDraw) { }
        }
#endif
        public bool Aplicar(ObjectId[] ids)
        {
#if NANOCAD5
            return false;
#else
            Dispose();
            if (ids.Length == 0) return true;
            var classe = RXObject.GetClass(typeof(MText));
            var filtro = new Ocultar(ids);
            try { Overrule.AddOverrule(classe, filtro, true); }
            catch { filtro.Dispose(); throw; }
            _classe = classe;
            _filtro = filtro;
            Overrule.Overruling = true;
            return true;
#endif
        }

        public void Dispose()
        {
#if !NANOCAD5
            if (_filtro == null) return;
            Overrule.RemoveOverrule(_classe, _filtro);
            _filtro.Dispose();
            _filtro = null;
#endif
        }
    }
}
