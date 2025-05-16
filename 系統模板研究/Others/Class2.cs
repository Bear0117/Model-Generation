using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Diagnostics.Contracts;
using Teigha.Runtime;
using Aspose.Cells.Charts;
using System;
//using Exception = System.Exception;
//using Teigha.DatabaseServices;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class Class_2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("1F")) as Level;

            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(Level));

            Reference r1 = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            Element elem1 = doc.GetElement(r1);

            Reference r2 = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            Element elem2 = doc.GetElement(r2);

            Solid wallSolid = GetPrimarySolid(doc, elem1);
            Solid beamSolid = GetPrimarySolid(doc, elem2);

            Solid intersectionSolid = BooleanOperationsUtils.ExecuteBooleanOperation(wallSolid, beamSolid, BooleanOperationsType.Intersect);

            List<XYZ> intersectionPoints = new List<XYZ>();
            foreach (Face face in intersectionSolid.Faces)
            {
                foreach (EdgeArray edgeArray in face.EdgeLoops)
                {
                    foreach (Edge edge in edgeArray)
                    {
                        XYZ startPt = edge.AsCurve().GetEndPoint(0);
                        XYZ endPt = edge.AsCurve().GetEndPoint(1);
                        intersectionPoints.Add(startPt);
                        intersectionPoints.Add(endPt);
                    }
                }
            }
            MessageBox.Show(intersectionPoints.Count.ToString());



            return Result.Succeeded;
        }
        public Solid GetPrimarySolid(Document doc, Element elem)
        {
            Solid mainSolid = null;
            double maxVolume = 0;


            Options opt = new Options();
            GeometryElement geomElem = elem.get_Geometry(opt);

            foreach (GeometryObject geomObj in geomElem)
            {
                if (geomObj is Solid)
                {
                    Solid solid = geomObj as Solid;
                    if (solid.Volume > maxVolume)
                    {
                        mainSolid = solid;
                        maxVolume = solid.Volume;
                    }
                }

                else if (geomObj is GeometryInstance)
                {
                    GeometryInstance geomInst = geomObj as GeometryInstance;
                    GeometryElement geomElemInst = geomInst.GetInstanceGeometry();
                    foreach (GeometryObject geomObjInst in geomElemInst)
                    {
                        if (geomObjInst is Solid)
                        {
                            Solid solid = geomObjInst as Solid;
                            if (solid.Volume > maxVolume)
                            {
                                mainSolid = solid;
                                maxVolume = solid.Volume;
                            }
                        }
                    }
                }
            }
            return mainSolid;
        }


    }
}