using System;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using Autodesk.Revit.DB.Structure;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class Test : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            //double wallHeight = 5000 * 0.003281;
            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("4F")) as Level;

            Line line1 = Line.CreateBound(new XYZ(0, 0, 0),new XYZ(20, 0 , 0));

            Line line2 = Line.CreateBound(new XYZ(20, 0, 0), new XYZ(100, 0, 0));

            Line line3 = Line.CreateBound(new XYZ(20, 0, 0), new XYZ(-80, 0, 0));

            Line line4 = Line.CreateBound(new XYZ(-80, 0, 0), new XYZ(20, 0, 0));

            MessageBox.Show(line1.Direction.ToString());
            MessageBox.Show(line2.Direction.ToString());
            MessageBox.Show(line3.Direction.ToString());
            MessageBox.Show(line4.Direction.ToString());

            return Result.Succeeded;
        }
    }
}
