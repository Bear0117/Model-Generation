using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using Autodesk.Revit.DB.IFC;
using System;
using System.Linq;
using System.IO;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class GetWallProfile : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            View view = doc.ActiveView;

            Autodesk.Revit.Creation.Application creapp = app.Create;
            Autodesk.Revit.Creation.Document credoc = doc.Create;

            Reference r = uidoc.Selection.PickObject(ObjectType.Element, "Select a wall");
            Element e = uidoc.Document.GetElement(r);
            Wall wall = e as Wall;

            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Wall Profile");

                // Get the external wall face for the profile
                IList<Reference> sideFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior);
                Element e2 = doc.GetElement(sideFaces[0]);
                Face face = e2.GetGeometryObjectFromReference(sideFaces[0]) as Face;

                // The normal of the wall external face.
                XYZ normal = face.ComputeNormal(new UV(0, 0));

                // Offset curve copies for visibility.
                Transform offset = Transform.CreateTranslation(5 * normal);

                // If the curve loop direction is counter-
                // clockwise, change its color to RED.
                Color colorRed = new Color(255, 0, 0);

                // Get edge loops as curve loops.
                IList<CurveLoop> curveLoops = face.GetEdgesAsCurveLoops();

                // ExporterIFCUtils class can also be used for
                // non-IFC purposes. The SortCurveLoops method
                // sorts curve loops (edge loops) so that the
                // outer loops come first.
                //IList<IList<CurveLoop>> curveLoopLoop = ExporterIFCUtils.SortCurveLoops(curveLoops);

                //foreach (IList<CurveLoop> curveLoops2 in curveLoopLoop)
                //{
                //    foreach (CurveLoop curveLoop2 in curveLoops2)
                //    {
                //        // Check if curve loop is counter-clockwise.
                //        bool isCCW = curveLoop2.IsCounterclockwise(normal);
                //        CurveArray curves = creapp.NewCurveArray();

                //        foreach (Curve curve in curveLoop2)
                //        {
                //            curves.Append(curve.CreateTransformed(offset));
                //        }

                //        // Create model lines for an curve loop.

                //        // this was obsolete, now it's missing
                //        //Plane plane = creapp.NewPlane(curves);

                //        // this is not offseted, and I couldn't figure out how to offset it
                //        Plane plane1 = curveLoop2.GetPlane();

                //        // so I created a new one
                //        Plane plane2 = GeometryUtil.CreatePlaneByXYVectorsContainingPoint(
                //        plane1.XVec, plane1.YVec, plane1.Origin + 5 * normal);

                //        SketchPlane sketchPlane = SketchPlane.Create(doc, plane2);
                //        ModelCurveArray curveElements = credoc.NewModelCurveArray(curves, sketchPlane);

                //        if (isCCW)
                //        {
                //            foreach (ModelCurve mcurve in curveElements)
                //            {
                //                OverrideGraphicSettings overrides = view.GetElementOverrides(mcurve.Id);
                //                overrides.SetProjectionLineColor(colorRed);
                //                view.SetElementOverrides(mcurve.Id, overrides);
                //            }
                //        }
                //    }
                //}
                tx.Commit();
            }
            return Result.Succeeded;
        }
    }
}
