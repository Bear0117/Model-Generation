using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class WallDimensionsl : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            Face face1 = null;
            Face face2 = null;
            Reference reference = uiDoc.Selection.PickObject(Autodesk.Revit.UI.Selection.ObjectType.Element);
            Element elem = doc.GetElement(reference.ElementId);
            Options options = new Options();
            options.View = uiDoc.ActiveView;
            options.ComputeReferences = true;
            GeometryElement geometryElement = elem.get_Geometry(options);
            foreach (GeometryObject gObj in geometryElement)
            {
                Solid solid = gObj as Solid;
                foreach (Face face in solid.Faces)
                {
                    XYZ normal = face.ComputeNormal(new UV(0, 0));
                    if (Math.Abs(normal.X) > 0.1)
                    {
                        if (normal.X > 0.1)
                        {
                            face1 = face;
                        }
                        else
                        {
                            face2 = face;
                        }
                    }
                }
            }

            if (face1 != null && face2 != null)
            {
                Transaction tran = new Transaction(doc, "Create Dimension");
                tran.Start();
                XYZ p1 = face1.Evaluate(new UV(0, 0));
                XYZ p2 = face2.Project(p1).XYZPoint;
                p1 = new XYZ(p1.X, p1.Y + 10.0, p1.Z);
                p2 = new XYZ(p2.X, p2.Y + 10.0, p2.Z);
                Line line = Line.CreateBound(p1, p2);

                ReferenceArray referenceArray = new ReferenceArray();
                MessageBox.Show(face1.Reference.ToString() + face2.Reference.ToString());
                referenceArray.Append(face1.Reference);
                referenceArray.Append(face2.Reference);
                doc.Create.NewDimension(uiDoc.ActiveView, line, referenceArray);
                tran.Commit();
            }

            return Result.Succeeded;
        }
    }
}

