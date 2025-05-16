using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Windows;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class GetPointCoordinate : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            Reference r = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            Element elem = doc.GetElement(r);
            LocationPoint location = elem.Location as LocationPoint;
            MessageBox.Show(location.ToString());

            return Result.Succeeded;
        }
    }
}
