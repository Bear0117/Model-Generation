using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class CreateWallModelline : IExternalCommand
    {
        Document doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;

            CurveArray curveArray = new CurveArray();
            XYZ point_1 = new XYZ(0, 0, 0);
            int numLine = 0;

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
                    MessageBox.Show(name.ToString());
                }
            }

            // Decide gridline size by users.
            ModelingParam.Initialize();
            double gridline_size = ModelingParam.parameters.General.GridSize * 10; // unit: mm

            // Hide the selected layer.
            Transaction trans = new Transaction(doc, "隱藏圖層");
            trans.Start();
            if (targetCategory != null)
            {
                ElementId elementId = targetCategory.Id;
                doc.ActiveView.SetCategoryHidden(elementId, true);
            }
            trans.Commit();


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
                            if (line.GetEndPoint(0).DistanceTo(line.GetEndPoint(1)) > CentimetersToUnits(1))
                            {
                                Line newLine = Line.CreateBound(line.GetEndPoint(0), line.GetEndPoint(1));
                                XYZ normal = XYZ.BasisZ;
                                XYZ point = line.GetEndPoint(0);
                                point = Algorithm.RoundPoint(transform.OfPoint(point), gridline_size);
                                curveArray.Append(TransformLine(transform, newLine, gridline_size));
                                numLine++;
                            }
                            else continue;
                        }

                        else if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> points = polyLine.GetCoordinates();

                            for (int j = 0; j < points.Count - 1; j++)
                            {
                                if (points[j].DistanceTo(points[j + 1]) < CentimetersToUnits(1))
                                {
                                    continue;
                                }
                                Line line = Line.CreateBound(points[j], points[j + 1]);
                                line = TransformLine(transform, line, gridline_size);
                                curveArray.Append(line);

                            }
                            XYZ normal = XYZ.BasisZ;
                            XYZ point = points.First();
                            point = transform.OfPoint(point);
                            point_1 = point;
                            numLine++;
                        }
                        else continue;
                    }
                }
            }

            CurveArray curveArray_new = new CurveArray();
            IList<Curve> curves = GetCurveList(curveArray);
            IList<Curve> hCurves = new List<Curve>();
            IList<Curve> vCurves = new List<Curve>();

            // Classify the horizontal and vertical lines. 
            foreach (Curve c in curves)
            {
                if (Math.Abs(c.GetEndPoint(0).Y - c.GetEndPoint(1).Y) < CentimetersToUnits(0.1))
                {
                    hCurves.Add(c);
                }
                else
                {
                    vCurves.Add(c);
                }
            }

            // Merge the horizontal and vertical lines.
            IList<Curve> hCurves_merged = MergeCurves(hCurves, true, gridline_size);
            IList<Curve> vCurves_merged = MergeCurves(vCurves, false, gridline_size);

            // Add all the merged vertical lines into "result" CurvrArray.
            foreach (Curve c in hCurves_merged)
            {
                curveArray_new.Append(c);
            }

            // Add all the merged vertical lines into "result" CurvrArray.
            foreach (Curve c in vCurves_merged)
            {
                curveArray_new.Append(c);
            }

            // Create model lines according to all the curves in the "curveArray_new".
            CreateModelCurveArray(curveArray_new, XYZ.BasisZ, point_1);

            return Result.Succeeded;
        }

        public IList<Curve> MergeCurves_new(IList<Curve> curveList, bool horizontal)
        {
            List<Curve> result = new List<Curve>();

            if (horizontal)
            {
                List<Curve> canMerged = new List<Curve>();
                foreach (Curve c1 in curveList)
                {
                    bool isMerged = false;
                    foreach (Curve c2 in curveList)
                    {
                        if (c1.Id == c2.Id) continue;

                        if (c1.Distance(c2.GetEndPoint(0)) < CentimetersToUnits(0.1) ||
                                c1.Distance(c2.GetEndPoint(1)) < CentimetersToUnits(0.1))
                        {
                            canMerged.Add(c1);
                            isMerged = true;
                            break;
                        }
                    }
                    if (!isMerged)
                    {
                        result.Add(c1);
                    }
                }

                bool conti_h = true;
                while (conti_h)
                {
                    int num = 0;
                    for (int i = 0; i < canMerged.Count(); i++)
                    {
                        Line c1 = canMerged[i] as Line;
                        for (int j = 0; j < canMerged.Count(); j++)
                        {
                            if (i >= j)
                            {
                                continue;
                            }
                            Line c2 = canMerged[j] as Line;
                            if (c1.Distance(c2.GetEndPoint(0)) < CentimetersToUnits(0.1) ||
                                c1.Distance(c2.GetEndPoint(1)) < CentimetersToUnits(0.1))
                            {
                                double Y = c1.GetEndPoint(0).Y;
                                double Z = c1.GetEndPoint(0).Z;
                                Curve cur1 = canMerged[i];
                                Curve cur2 = canMerged[j];
                                canMerged.Remove(cur1);
                                canMerged.Remove(cur2);
                                List<double> coors = new List<double>
                                {
                                cur1.GetEndPoint(0).X,
                                cur1.GetEndPoint(1).X,
                                cur2.GetEndPoint(0).X,
                                cur2.GetEndPoint(1).X
                                };
                                XYZ startPt = new XYZ(coors.Min(), Y, Z);
                                XYZ endPt = new XYZ(coors.Max(), Y, Z);
                                if (startPt.DistanceTo(endPt) > 0)
                                {
                                    canMerged.Add(Line.CreateBound(startPt, endPt) as Curve);
                                    num++;
                                }
                                break;
                            }

                        }
                        if (num != 0)
                        {
                            break;
                        }
                    }
                    if (num == 0)
                    {
                        conti_h = false;
                    }
                }
                foreach (Curve c in canMerged)
                {
                    result.Add(c);
                } 
            }

            else
            {
                List<Curve> canMerged = new List<Curve>();

                foreach (Curve c1 in curveList)
                {
                    bool isMerged = false;
                    foreach (Curve c2 in curveList)
                    {
                        if (c1.Id == c2.Id) continue;

                        if (c1.Distance(c2.GetEndPoint(0)) < CentimetersToUnits(0.1) ||
                                c1.Distance(c2.GetEndPoint(1)) < CentimetersToUnits(0.1))
                        {
                            canMerged.Add(c1);
                            isMerged = true;
                            break;
                        }
                    }
                    if (!isMerged)
                    {
                        result.Add(c1);
                    }
                }

                bool conti_h = true;
                while (conti_h)
                {
                    int num = 0;
                    for (int i = 0; i < canMerged.Count(); i++)
                    {
                        Line c1 = canMerged[i] as Line;
                        for (int j = 0; j < canMerged.Count(); j++)
                        {
                            if (i >= j)
                            {
                                continue;
                            }
                            Line c2 = canMerged[j] as Line;
                            
                            if (c1.Distance(c2.GetEndPoint(0)) < CentimetersToUnits(0.1) ||
                                c1.Distance(c2.GetEndPoint(1)) < CentimetersToUnits(0.1))
                            {
                                double X = c1.GetEndPoint(0).X;
                                double Z = c1.GetEndPoint(0).Z;
                                Curve cur1 = canMerged[i];
                                Curve cur2 = canMerged[j];
                                canMerged.Remove(cur1);
                                canMerged.Remove(cur2);
                                List<double> coors = new List<double>
                                {
                                cur1.GetEndPoint(0).Y,
                                cur1.GetEndPoint(1).Y,
                                cur2.GetEndPoint(0).Y,
                                cur2.GetEndPoint(1).Y
                                };
                                XYZ startPt = new XYZ(X, coors.Min(), Z);
                                XYZ endPt = new XYZ(X, coors.Max(), Z);
                                if (startPt.DistanceTo(endPt) > 0)
                                {
                                    canMerged.Add(Line.CreateBound(startPt, endPt) as Curve);
                                    num++;
                                }
                                break;
                            }
                        }
                        if (num != 0)
                        {
                            break;
                        }
                    }
                    if (num == 0)
                    {
                        conti_h = false;
                    }
                }
                foreach (Curve c in canMerged)
                {
                    result.Add(c);
                }
            }
            return result;
        }

        public IList<Curve> MergeCurves(IList<Curve> curveList, bool horizontal, double gridline_size)
        {
            if (horizontal)
            {
                bool conti_h = true;
                while (conti_h)
                {
                    int num = 0;
                    for (int i = 0; i < curveList.Count(); i++)
                    {
                        Line c1 = curveList[i] as Line;
                        for (int j = 0; j < curveList.Count(); j++)
                        {
                            if (i >= j)
                            {
                                continue;
                            }
                            Line c2 = curveList[j] as Line;
                            if (c1.Distance(c2.GetEndPoint(0)) < CentimetersToUnits(0.1) ||
                                c1.Distance(c2.GetEndPoint(1)) < CentimetersToUnits(0.1))
                            {
                                //MessageBox.Show("h");
                                double Y = c1.GetEndPoint(0).Y;
                                double Z = c1.GetEndPoint(0).Z;
                                Curve cur1 = curveList[i];
                                Curve cur2 = curveList[j];
                                curveList.Remove(cur1);
                                curveList.Remove(cur2);
                                List<double> coors = new List<double>
                                {
                                cur1.GetEndPoint(0).X,
                                cur1.GetEndPoint(1).X,
                                cur2.GetEndPoint(0).X,
                                cur2.GetEndPoint(1).X
                                };
                                XYZ startPt = new XYZ(coors.Min(), Y, Z);
                                XYZ endPt = new XYZ(coors.Max(), Y, Z);
                                if (startPt.DistanceTo(endPt) > 0)
                                {
                                    curveList.Add(Line.CreateBound(Algorithm.RoundPoint(startPt, gridline_size), Algorithm.RoundPoint(endPt, gridline_size)) as Curve);
                                    num++;
                                }
                                break;
                            }
                        }
                        if (num != 0)
                        {
                            break;
                        }
                    }
                    if (num == 0)
                    {
                        conti_h = false;
                    }
                }
            }
            else
            {
                bool conti_v = true;
                while (conti_v)
                {
                    int num = 0;
                    for (int i = 0; i < curveList.Count(); i++)
                    {
                        Line c1 = curveList[i] as Line;
                        for (int j = 0; j < curveList.Count(); j++)
                        {
                            if (i >= j)
                            {
                                continue;
                            }
                            Line c2 = curveList[j] as Line;
                            if (c1.Distance(c2.GetEndPoint(0)) < CentimetersToUnits(0.1) ||
                                c1.Distance(c2.GetEndPoint(1)) < CentimetersToUnits(0.1))
                            {
                                double X = curveList[i].GetEndPoint(0).X;
                                double Z = curveList[i].GetEndPoint(0).Z;
                                Curve cur1 = curveList[i];
                                Curve cur2 = curveList[j];
                                curveList.Remove(cur1);
                                curveList.Remove(cur2);
                                List<double> coors = new List<double>
                                {
                                cur1.GetEndPoint(0).Y,
                                cur1.GetEndPoint(1).Y,
                                cur2.GetEndPoint(0).Y,
                                cur2.GetEndPoint(1).Y
                                };

                                XYZ startPt = new XYZ(X, coors.Min(), Z);
                                XYZ endPt = new XYZ(X, coors.Max(), Z);
                                if (startPt.DistanceTo(endPt) > 0)
                                {
                                    curveList.Add(Line.CreateBound(Algorithm.RoundPoint(startPt, gridline_size), Algorithm.RoundPoint(endPt, gridline_size)) as Curve);
                                    num++;
                                }
                                break;
                            }
                        }
                        if (num != 0)
                        {
                            break;
                        }
                    }
                    if (num == 0)
                    {
                        conti_v = false;
                    }
                }

            }
            return curveList;
        }

        public double UnitsToCentimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
        }

        public double CentimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Centimeters);
        }

        public Line RoundLine(Line line, double gridline_size)
        {
            XYZ point_A = Algorithm.RoundPoint(line.GetEndPoint(0), gridline_size);
            XYZ point_B = Algorithm.RoundPoint(line.GetEndPoint(1), gridline_size);
            Line line_new = Line.CreateBound(point_A, point_B);
            return line_new;
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

        private IList<Curve> GetCurveList(CurveArray curveArray)
        {
            IList<Curve> list = new List<Curve>();
            foreach (Curve curve in curveArray)
            {
                list.Add(curve);
            }
            return list;
        }

        private Line TransformLine(Transform transform, Line line, double gridline_size)
        {
            XYZ startPoint = transform.OfPoint(line.GetEndPoint(0));
            XYZ endPoint = transform.OfPoint(line.GetEndPoint(1));
            Line newLine = Line.CreateBound(Algorithm.RoundPoint(startPoint, gridline_size), Algorithm.RoundPoint(endPoint, gridline_size));
            return newLine;
        }

    }
}
