 using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateModelLine_Test : IExternalCommand
    {
        Document doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;

            Reference refer = uidoc.Selection.PickObject(ObjectType.PointOnElement, "Select a CAD Layer");
            Element elem = doc.GetElement(refer);
            GeometryObject geoObj = elem.GetGeometryObjectFromReference(refer);

            Category targetCategory = null;
            ElementId graphicsStyleId = null;

            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                graphicsStyleId = geoObj.GraphicsStyleId;
                if (doc.GetElement(geoObj.GraphicsStyleId) is GraphicsStyle gs)
                {
                    targetCategory = gs.GraphicsStyleCategory;
                    String name = gs.GraphicsStyleCategory.Name;
                    TaskDialog.Show("1", name);
                }
            }

            // Hide the selected layer.
            Transaction trans = new Transaction(doc, "隱藏圖層");
            trans.Start();
            if (targetCategory != null)
            {
                ElementId elementId = targetCategory.Id;
                doc.ActiveView.SetCategoryHidden(elementId, true);
            }
            trans.Commit();

            //Draw the Model Line
            TransactionGroup transGroup = new TransactionGroup(doc, "繪製模型線");
            transGroup.Start();
            CurveArray curveArray = new CurveArray();

            int numPolyline = 0;

            // To verify the type of elements.
            GeometryElement geoElem = elem.get_Geometry(new Options());
            foreach (GeometryObject gObj in geoElem)
            {
                GeometryInstance geomInstance = gObj as GeometryInstance;
                Transform transform = geomInstance.Transform;
                if (null != geomInstance)
                {
                    foreach (GeometryObject insObj in geomInstance.SymbolGeometry)
                    {
                        if (insObj.GraphicsStyleId.IntegerValue != graphicsStyleId.IntegerValue)
                            continue;

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.Line")
                        {
                            Line line = insObj as Line;

                            XYZ normal = XYZ.BasisZ;
                            XYZ point = line.GetEndPoint(0);
                            point = transform.OfPoint(point);
                            curveArray.Append(TransformLine(transform, line));
                            CreateModelCurveArray(curveArray, normal, point);
                            numPolyline++;

                        }

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> points = polyLine.GetCoordinates();

                            for (int i = 0; i < points.Count - 1; i++)
                            {
                                if (points[i].DistanceTo(points[i + 1]) < 1 * 0.003281)
                                {
                                    continue;
                                }
                                Line line = Line.CreateBound(points[i], points[i + 1]);
                                line = TransformLine(transform, line);
                                curveArray.Append(line);
                                numPolyline++;
                            }
                            XYZ normal = XYZ.BasisZ;
                            XYZ point = points.First();
                            point = transform.OfPoint(point);

                            CreateModelCurveArray(curveArray, normal, point);
                            
                        }
                    }
                }
            }
            transGroup.Assimilate();
            return Result.Succeeded;
        }
        private void CreateModelCurveArray(CurveArray curveArray, XYZ normal, XYZ point)
        {
            if (curveArray.Size > 0)
            {
                Transaction transaction2 = new Transaction(doc);
                transaction2.Start("繪製模型線");
                try
                {
                    SketchPlane modelSketch = SketchPlane.Create(doc, Plane.CreateByNormalAndOrigin(normal, point));
                    ModelCurveArray modelLine = doc.Create.NewModelCurveArray(curveArray, modelSketch);

                    Category tCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
                    Category nCat = doc.Settings.Categories.NewSubcategory(tCat, "MyLine");

                    // 更改線條顏色
                    //Category tCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
                    //Category nCat = doc.Settings.Categories.NewSubcategory(tCat, "MyLine");
                    //nCat.LineColor = new Color(255, 0, 0);
                }
                catch
                { }
                transaction2.Commit();
                curveArray.Clear();
            }
        }
        private Line TransformLine(Transform transform, Line line)
        {
            XYZ startPoint = transform.OfPoint(line.GetEndPoint(0));
            XYZ endPoint = transform.OfPoint(line.GetEndPoint(1));
            Line newLine = Line.CreateBound(startPoint, endPoint);
            return newLine;
        }
    }
}
