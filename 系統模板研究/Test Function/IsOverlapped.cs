using System; // 要使用Exception需引用此參考
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class IsOverlapped : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            
            return Result.Succeeded;
        }
        private static IEnumerable<Solid> GetSolids(Element element)
        {
            GeometryElement geometry = element.get_Geometry(new Options{ ComputeReferences = true, IncludeNonVisibleObjects = true});
            if (geometry == null)
            {
                return Enumerable.Empty<Solid>();
            }
            else
            {
                return GetSolids(geometry).Where(x => x.Volume > 0);
            }
        }

        private static IEnumerable<Solid> GetSolids(IEnumerable<GeometryObject> geometryElement)
        {
            foreach (var geometry in geometryElement)
            {
                Solid solid = geometry as Solid;
                if (solid != null)
                    yield return solid;

                GeometryInstance instance = geometry as GeometryInstance;
                if (instance != null)
                    foreach (Solid instanceSolid in GetSolids(instance.GetInstanceGeometry()))
                        yield return instanceSolid;

                GeometryElement element = geometry as GeometryElement;
                if (element != null)
                    foreach (Solid elementSolid in GetSolids(element))
                        yield return elementSolid;
            }
        }
    }
}