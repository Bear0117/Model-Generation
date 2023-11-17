using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class TwoLineToDoor : IExternalCommand
    {
        Autodesk.Revit.ApplicationServices.Application app;
        Document doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("1F")) as Level;
            for (int i = 0; i < 10000; i++)
            {
                IList<XYZ> points = new List<XYZ>();
                for (int j = 0; j < 2; j++)
                {
                    Reference r = uidoc.Selection.PickObject(ObjectType.PointOnElement);
                    Element elem = doc.GetElement(r);
                    LocationCurve locationCurve = elem.Location as LocationCurve;
                    Line locationLine = locationCurve.Curve as Line;
                    points.Add(locationLine.GetEndPoint(0));
                    points.Add(locationLine.GetEndPoint(1));
                }
                Line Line1 = Line.CreateBound(points[0], points[1]);
                Line Line2 = Line.CreateBound(points[2], points[3]);
                Line BreakLine = Line.CreateBound(points[0], points[2]);
                XYZ midPoint1 = midPoint(points[0], points[3]);
                XYZ midPoint2 = midPoint(points[1], points[2]);
                Line STDLine1 = Line.CreateBound(points[0], points[1]);
                Line STDLine2 = Line.CreateBound(points[1], points[3]);
                //1mm = 0.03937 inch
                if (Line1.Length == Line2.Length)
                {
                    if (midPoint1.DistanceTo(midPoint2) < 150 / 304.78)
                    {
                        midPoint1 = midPoint(points[0], points[2]);
                        midPoint2 = midPoint(points[1], points[3]);
                    }
                }
                else if (Line1.Length > Line2.Length)
                {
                    double offset = 150 * 0.03937 / 2;
                    //XYZ XoffsetCoor = new XYZ(offset, 0, 0);
                    //XYZ YoffsetCoor = new XYZ(0, offset, 0); 
                    if (Line1.Length > Line2.Length)
                    {
                        if (Line1.GetEndPoint(0).X.ToString() != Line1.GetEndPoint(1).X.ToString())
                        {
                            XYZ dir = Line1.GetEndPoint(0) - Line2.GetEndPoint(0);
                            if (dir.Y.ToString().Contains("-"))
                            {
                                midPoint1 = points[2] + new XYZ(0, offset, 0) * (-1);
                                midPoint2 = points[3] + new XYZ(0, offset, 0) * (-1);
                            }
                            else
                            {
                                midPoint1 = points[2] + new XYZ(0, offset, 0);
                                midPoint2 = points[3] + new XYZ(0, offset, 0);
                            }
                        }
                        else
                        {
                            XYZ dir = Line1.GetEndPoint(0) - Line2.GetEndPoint(0);
                            if (dir.X.ToString().Contains("-"))
                            {
                                midPoint1 = points[2] + new XYZ(0, 0, offset) * (-1);
                                midPoint2 = points[3] + new XYZ(0, 0, offset) * (-1);
                            }
                            else
                            {
                                midPoint1 = points[2] + new XYZ(0, 0, offset);
                                midPoint2 = points[3] + new XYZ(0, 0, offset);
                            }
                        }
                    }
                    else
                    {
                        if (Line1.GetEndPoint(0).X.ToString() != Line1.GetEndPoint(1).X.ToString())
                        {
                            XYZ dir = Line2.GetEndPoint(0) - Line1.GetEndPoint(0);
                            if (dir.Y.ToString().Contains("-"))
                            {
                                midPoint1 = points[0] + new XYZ(0, offset, 0) * (-1);
                                midPoint2 = points[1] + new XYZ(0, offset, 0) * (-1);
                            }
                            else
                            {
                                midPoint1 = points[0] + new XYZ(0, offset, 0);
                                midPoint2 = points[1] + new XYZ(0, offset, 0);
                            }
                        }
                        else
                        {
                            XYZ dir = Line2.GetEndPoint(0) - Line1.GetEndPoint(0);
                            if (dir.X.ToString().Contains("-"))
                            {
                                midPoint1 = points[0] + new XYZ(0, 0, offset) * (-1);
                                midPoint2 = points[1] + new XYZ(0, 0, offset) * (-1);
                            }
                            else
                            {
                                midPoint1 = points[0] + new XYZ(0, 0, offset);
                                midPoint2 = points[1] + new XYZ(0, 0, offset);
                            }
                        }
                    }
                    midPoint1 = midPoint(points[0], points[3]);
                    midPoint2 = midPoint(points[1], points[2]);
                    if (STDLine1.Length == STDLine2.Length)
                    {
                        midPoint1 = midPoint(points[0], points[1]);
                        midPoint2 = midPoint(points[2], points[3]);
                    }

                }
                foreach (XYZ p in points)
                {
                    //TaskDialog.Show("線條資訊", p.ToString(), TaskDialogCommonButtons.Ok);
                }
                Line wallLine = Line.CreateBound(midPoint1, midPoint2);
                if (BreakLine.Length > 80)
                    break;
                Transaction t1 = new Transaction(doc, "創建");
                t1.Start();
                Wall wall = Wall.Create(doc, wallLine, level.Id, true);
                CreateOpening(doc, wall);
                Parameter wallHeight = wall.get_Parameter(BuiltInParameter.WALL_ATTR_HEIGHT_PARAM);
                //wallHeight.Set(3300 * 0.003281);
                //Parameter wallHeight = wall.get_Parameter(BuiltInParameter.WALL_ATTR_WIDTH_PARAM);
                //wallHeight.Set(wallHeightSet);
                //double wallHeightSet = (input by User)
                t1.Commit();
            }
            return Result.Succeeded;
        }
        private XYZ midPoint(XYZ a, XYZ b)
        {
            XYZ c = (a + b) / 2;
            return c;
        }

        private Line TransformLine(Transform transform, Line line)
        {
            XYZ startPoint = transform.OfPoint(line.GetEndPoint(0));
            XYZ endPoint = transform.OfPoint(line.GetEndPoint(1));
            Line newLine = Line.CreateBound(startPoint, endPoint);
            return newLine;
        }
        private void CreateModelCurveArray(CurveArray curveArray, XYZ normal, XYZ point, Line line)
        {
            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("1F")) as Level;
            if (curveArray.Size > 0)
            {
                Transaction t1 = new Transaction(doc, "創建");
                t1.Start();
                try
                {
                    SketchPlane modelSketch = SketchPlane.Create(doc, Plane.CreateByNormalAndOrigin(normal, point));
                    ModelCurveArray modelLine = doc.Create.NewModelCurveArray(curveArray, modelSketch);
                    //Wall wall = Wall.Create(doc, line, level.Id, true);
                    //Parameter wallHeight = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                    //wallHeight.Set(50);


                    //TaskDialog.Show("圖層資訊", "線條種類：" +"co co ", TaskDialogCommonButtons.Ok);
                }
                catch
                {

                }
                t1.Commit();
                curveArray.Clear();
            }
        }
        private void CreateOpening(Document RevitDoc, Wall wall)
        {
            LocationCurve locationCurve = wall.Location as LocationCurve;
            Line location = locationCurve.Curve as Line;
            XYZ startPoint = location.GetEndPoint(0);
            XYZ endPoint = location.GetEndPoint(1);
            //Transaction t1 = new Transaction((RevitDoc), "Set Wall Height");
            //t1.Start();          
            //t1.Commit();
            using (Transaction transaction = new Transaction(RevitDoc))
            {
                transaction.Start("Create Opening on wall");
                Parameter wallHeightParameter = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                wallHeightParameter.Set(4350 * 0.00328);
                double wallHeight = wallHeightParameter.AsDouble();
                //XYZ delta = new XYZ(0, 0, wallHeight) / 3;
                Opening opening = RevitDoc.Create.NewOpening(wall, startPoint, endPoint + new XYZ(0, 0, 213 * 0.00328));
                transaction.Commit();
            }

        }
    }
}