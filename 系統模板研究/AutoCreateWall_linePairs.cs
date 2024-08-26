using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Teigha.DatabaseServices;
using Outline = Autodesk.Revit.DB.Outline;
using Line = Autodesk.Revit.DB.Line;
using Transaction = Autodesk.Revit.DB.Transaction;
using Solid = Autodesk.Revit.DB.Solid;
using Curve = Autodesk.Revit.DB.Curve;
using Aspose.Cells;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateWall_linePairs : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            
            // To know the information of selected layer
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
                    // Get the name of the CAD layer which is selected (Column).
                    String name = gs.GraphicsStyleCategory.Name;
                }
            }

            XYZ point_1 = new XYZ(0, 0, 0);
            Transaction tstart = new Transaction(doc, "Get wall line info!");
            tstart.Start();
            GeometryElement geoElem = elem.get_Geometry(new Options());

            CurveArray curveArray_h = new CurveArray();
            CurveArray curveArray_v = new CurveArray();
            double gridline_size = new double();
            gridline_size = 5;

            foreach (GeometryObject gObj in geoElem)
            {
                GeometryInstance geomInstance = gObj as GeometryInstance;
                Transform transform = geomInstance.Transform;//座標轉換=
                if (null != geomInstance)
                {
                    foreach (GeometryObject insObj in geomInstance.SymbolGeometry)
                    {
                        if (insObj.GraphicsStyleId.IntegerValue != graphicsStyleId.IntegerValue)
                        {
                            continue;
                        }

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.Line")
                        {
                            Line line = insObj as Line;
                            if (line.GetEndPoint(0).DistanceTo(line.GetEndPoint(1)) > CentimetersToUnits(1))
                            {
                                Line newLine = Line.CreateBound(line.GetEndPoint(0), line.GetEndPoint(1));
                                XYZ point = line.GetEndPoint(0);
                                point = Algorithm.RoundPoint(transform.OfPoint(point), gridline_size);
                                point_1 = point;
                                if (Math.Abs(line.GetEndPoint(0).X - line.GetEndPoint(1).X) < CentimetersToUnits(1))
                                    curveArray_v.Append(TransformLine(transform, newLine, gridline_size));
                                else
                                    curveArray_h.Append(TransformLine(transform, newLine, gridline_size));
                            }
                            else continue;
                        }

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
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
                                //line = TransformLine(transform, line, gridline_size);
                                if (Math.Abs(line.GetEndPoint(0).X - line.GetEndPoint(1).X) < CentimetersToUnits(1))
                                    curveArray_v.Append(TransformLine(transform, line, gridline_size));
                                else
                                    curveArray_h.Append(TransformLine(transform, line, gridline_size));
                            }
                            XYZ normal = XYZ.BasisZ;
                            XYZ point = points.First();
                            point = transform.OfPoint(point);
                            point_1 = point; // 一定要用線上的點嗎?
                        }
                    }
                }
                MessageBox.Show("Nice Work!");
            }
            tstart.Commit();
            MessageBox.Show(curveArray_h.Size.ToString());
            MessageBox.Show(curveArray_v.Size.ToString());

            // Default "1F"
            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("1F")) as Level;

            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(Level));


            CurveArray central_lines = FindPairs(curveArray_v);
            MessageBox.Show(central_lines.Size.ToString());
            CreateModelCurveArray(doc, central_lines, XYZ.BasisZ, point_1);

            //while (true)
            //{
            //    double modelLineElevation = 0; // default
            //    IList<XYZ> points = new List<XYZ>();
            //    for (int j = 0; j < 2; j++)
            //    {
            //        Reference r = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            //        Element elem = doc.GetElement(r);
            //        level = doc.GetElement(elem.LevelId) as Level;
            //        LocationCurve locationCurve = elem.Location as LocationCurve;
            //        Line locationLine = locationCurve.Curve as Line;
            //        points.Add(locationLine.GetEndPoint(0));
            //        points.Add(locationLine.GetEndPoint(1));
            //        modelLineElevation = locationLine.GetEndPoint(0).Z;
            //    }

            //    //Stop building the wall.
            //    if (points[0].ToString() == points[2].ToString())
            //    {
            //        break;
            //    }

            //    double levelHeight = CentimetersToUnits(330); // default
            //    List<double> levelDistancesList = new List<double>();

            //    // retrieve all the Level elements in the document
            //    IEnumerable<Level> levels = collector.Cast<Level>();

            //    // loop through all the levels
            //    foreach (Level level_1 in levels)
            //    {
            //        // get the level elevation.
            //        double levelElevation = level_1.Elevation;
            //        if (Math.Abs(levelElevation - modelLineElevation) < CentimetersToUnits(0.1))
            //        {
            //            level = level_1;
            //        }
            //        if (levelElevation - modelLineElevation > CentimetersToUnits(0.1))
            //        {
            //            levelDistancesList.Add(levelElevation);
            //        }
            //    }

            //    if (levelDistancesList.Count() > 0)
            //    {
            //        levelHeight = levelDistancesList.Min() - level.Elevation;
            //    }
            //    else
            //    {
            //        LevelHeightForm form1 = new LevelHeightForm(doc);
            //        form1.ShowDialog();
            //        levelHeight = CentimetersToUnits(form1.levelHeight);
            //    }

            //    // To calculate the  central line of two lines. 
            //    //Line Line1 = Line.CreateBound(Algorithm.RoundPoint(points[0], gridline_size), Algorithm.RoundPoint(points[1], gridline_size));
            //    //Line Line2 = Line.CreateBound(Algorithm.RoundPoint(points[2], gridline_size), Algorithm.RoundPoint(points[3], gridline_size));
            //    Line Line1 = Line.CreateBound(points[0], points[1]);
            //    Line Line2 = Line.CreateBound(points[2], points[3]);
            //    XYZ midPoint1 = MidPoint(points[0], points[3]);
            //    XYZ midPoint2 = MidPoint(points[1], points[2]);

            //    double width = GetWallWidth(Line1, Line2); // Unit : mm

            //    Line wallLine = null;
            //    XYZ unitZ = new XYZ(0, 0, 1);
            //    XYZ vector_line1 = points[0] - points[1];
            //    XYZ offset = Cross(unitZ, vector_line1) / Line1.Length * CentimetersToUnits(width / 10) / 2;
            //    if (Dot(offset, points[2] - points[1]) > 0)
            //    {
            //        midPoint1 = points[0] + offset;
            //        midPoint2 = points[1] + offset;
            //        wallLine = Line.CreateBound(midPoint1, midPoint2);
            //    }
            //    else
            //    {
            //        midPoint1 = points[0] - offset;
            //        midPoint2 = points[1] - offset;
            //        wallLine = Line.CreateBound(midPoint1, midPoint2);
            //    }

            //    FilteredElementCollector colWall = new FilteredElementCollector(doc)
            //        .WhereElementIsNotElementType()
            //        .OfCategory(BuiltInCategory.OST_Walls)
            //        .OfClass(typeof(Wall));
            //    IList<Wall> walls = new List<Wall>();
            //    foreach (Wall w in colWall)
            //    {
            //        walls.Add(w);
            //    }

            //    // Create Wall.
            //    Transaction t1 = new Transaction(doc, "Create Wall");
            //    t1.Start();

            //    Wall wall = Wall.Create(doc, wallLine, level.Id, true);
            //    // = WallLocationLine.WallCenterline;
            //    Parameter wallHeightP = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
            //    wallHeightP.Set(levelHeight);
            //    WallUtils.DisallowWallJoinAtEnd(wall, 0);
            //    WallUtils.DisallowWallJoinAtEnd(wall, 1);
            //    t1.Commit();

            //    ChangeWallType(doc, wall, Line1, Line2);

            //    List<double> depthList = new List<double>();

            //    //Get the element by elementID and get the boundingbox and outline of this element.
            //    Element elementWall = wall as Element;
            //    BoundingBoxXYZ bbxyzElement = elementWall.get_BoundingBox(null);
            //    //Outline outline = new Outline(bbxyzElement.Min, bbxyzElement.Max);
            //    Outline newOutline = ReducedOutline(bbxyzElement);

            //    //Create a filter to get all the intersection elements with wall.
            //    BoundingBoxIntersectsFilter filterW = new BoundingBoxIntersectsFilter(newOutline);

            //    //Create a filter to get StructuralFraming (which include beam and column) and Slabs.
            //    ElementCategoryFilter filterSt = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);
            //    ElementCategoryFilter filterSl = new ElementCategoryFilter(BuiltInCategory.OST_Floors);
            //    LogicalOrFilter filterS = new LogicalOrFilter(filterSt, filterSl);

            //    //Combine two filter.
            //    LogicalAndFilter filter = new LogicalAndFilter(filterS, filterW);

            //    //A list to store the intersected elements.
            //    IList<Element> inter = new FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements();
            //    //MessageBox.Show(inter.Count.ToString());

            //    for (int i = 0; i < inter.Count; i++)
            //    {
            //        if (inter[i] != null)
            //        {
            //            string elementName = inter[i].Category.Name;
            //            if (elementName == "結構構架")
            //            {
            //                // Find the depth
            //                BoundingBoxXYZ bbxyzBeam = inter[i].get_BoundingBox(null);
            //                depthList.Add(Math.Abs(bbxyzBeam.Min.Z - bbxyzBeam.Max.Z));

            //                // Join
            //                Transaction joinAndSwitch = new Transaction(doc, "Join and Switch");
            //                joinAndSwitch.Start();
            //                RunJoinGeometryAndSwitch(doc, wall, inter[i]);
            //                joinAndSwitch.Commit();
            //            }
            //            else if (elementName == "樓板")
            //            {
            //                BoundingBoxXYZ bbxyzSlab = inter[i].get_BoundingBox(null);
            //                depthList.Add(Math.Abs(bbxyzSlab.Min.Z - bbxyzSlab.Max.Z));

            //                //Transaction join = new Transaction(doc, "Join");
            //                //join.Start();
            //                //RunJoinGeometry(doc, wall, inter[i]);
            //                //join.Commit();
            //            }
            //            else continue;
            //        }
            //    }

            //    List<double> faceZ = new List<double>();
            //    List<Solid> list_solid = GetElementSolidList(wall);
            //    FaceArray faceArray;
            //    foreach (Solid solid in list_solid)
            //    {
            //        faceArray = solid.Faces;
            //        foreach (Face face in faceArray)
            //        {
            //            BoundingBoxUV boxUV = face.GetBoundingBox();
            //            XYZ min = face.Evaluate(boxUV.Min);
            //            XYZ max = face.Evaluate(boxUV.Max);
            //            faceZ.Add(min.Z);
            //            faceZ.Add(max.Z);
            //            //MessageBox.Show(((face.GetBoundingBox().Max + face.GetBoundingBox().Min) / 2).ToString());
            //        }
            //        break;
            //        //MessageBox.Show(faceArray.Size.ToString());
            //    }

            //    Transaction t2 = new Transaction(doc, "Adjust Wall Height");
            //    t2.Start();
            //    wallHeightP.Set(faceZ.Max() - level.Elevation);
            //    t2.Commit();

            //    //ChangeWallType(doc, wall, Line1, Line2);
            //}
            return Result.Succeeded;
        }

        public CurveArray FindPairs(CurveArray curves)
        {
            List<Tuple<Line, Line, double>> pairedLines = new List<Tuple<Line, Line, double>>();
            CurveArray centralLines = new CurveArray();
            // 轉換CurveArray到List<Line>，確保所有曲線都是Line類型
            List<Line> lines = new List<Line>();
            foreach (Curve curve in curves)
            {
                if (curve is Line line)
                {
                    lines.Add(line);
                }
            }

            // 循環比較每一對線
            for (int i = 0; i < lines.Count; i++)
            {
                Line mainLine = lines[i];
                double mainYMin = Math.Min(mainLine.GetEndPoint(0).Y, mainLine.GetEndPoint(1).Y);
                double mainYMax = Math.Max(mainLine.GetEndPoint(0).Y, mainLine.GetEndPoint(1).Y);

                for (int j = 0; j < lines.Count; j++)
                {
                    if (i == j)
                    {
                        continue;
                    }
                    Line compareLine = lines[j];
                    double lineYMin = Math.Min(compareLine.GetEndPoint(0).Y, compareLine.GetEndPoint(1).Y);
                    double lineYMax = Math.Max(compareLine.GetEndPoint(0).Y, compareLine.GetEndPoint(1).Y);

                    // 檢查Y軸投影是否重疊
                    if (mainYMax - CentimetersToUnits(1) > lineYMin && lineYMax - CentimetersToUnits(1) > mainYMin)
                    {
                        // 檢查X軸距離
                        double distance = Math.Abs(mainLine.GetEndPoint(0).X - compareLine.GetEndPoint(0).X);
                        if (Math.Abs(UnitsToCentimeters(distance)) - 10 < 0.01 || Math.Abs(UnitsToCentimeters(distance)) - 15 < 0.01 || Math.Abs(UnitsToCentimeters(distance)) - 30 < 0.01)
                        {
                            Line wallLine = null;
                            double width = GetWallWidth(mainLine, compareLine);
                            XYZ unitZ = new XYZ(0, 0, 1);
                            XYZ vector_line1 = mainLine.GetEndPoint(0) - mainLine.GetEndPoint(1);
                            XYZ offset = Cross(unitZ, vector_line1) / mainLine.Length * CentimetersToUnits(width / 10) / 2;
                            if (Dot(offset, compareLine.GetEndPoint(0) - mainLine.GetEndPoint(0)) > 0)
                            {
                                XYZ midPoint1 = mainLine.GetEndPoint(0) + offset;
                                XYZ midPoint2 = mainLine.GetEndPoint(1) + offset;
                                wallLine = Line.CreateBound(midPoint1, midPoint2);
                            }
                            else
                            {
                                XYZ midPoint1 = mainLine.GetEndPoint(0) - offset;
                                XYZ midPoint2 = mainLine.GetEndPoint(1) - offset;
                                wallLine = Line.CreateBound(midPoint1, midPoint2);
                            }
                            centralLines.Append(wallLine);
                            // 如果條件符合，則將配對和距離加入列表
                            //pairedLines.Add(new Tuple<Line, Line, double>(mainLine, compareLine, distance));
                        }
                    }
                }
            }

            return centralLines;
        }

        private XYZ MidPoint(XYZ a, XYZ b)
        {
            XYZ c = (a + b) / 2;
            return c;
        }

        private void CreateModelCurveArray(Document doc, CurveArray curveArray, XYZ normal, XYZ point)
        {
            if (curveArray.Size > 0)
            {
                using (Transaction transaction2 = new Transaction(doc, "繪製模型線"))
                {
                    transaction2.Start();
                    try
                    {
                        SketchPlane modelSketch = SketchPlane.Create(doc, Plane.CreateByNormalAndOrigin(normal, point));
                        ModelCurveArray modelLine = doc.Create.NewModelCurveArray(curveArray, modelSketch);

                        // 處理類別和顏色設定
                        Category tCat = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Lines);
                        Category nCat = tCat.SubCategories.get_Item("MyLine");

                        if (nCat == null)
                        {
                            nCat = doc.Settings.Categories.NewSubcategory(tCat, "MyLine");
                            if (nCat != null)
                            {
                                nCat.LineColor = new Color(255, 0, 0);  // 設定顏色為紅色
                            }
                            else
                            {
                                throw new InvalidOperationException("Failed to create a new subcategory for lines.");
                            }
                        }

                        transaction2.Commit();
                    }
                    catch (Exception ex)
                    {
                        transaction2.RollBack();
                        MessageBox.Show($"Error: {ex.Message}");
                        throw;  // 向上拋出異常，以便呼叫者處理
                    }
                }
            }
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

        void RunJoinGeometry(Document doc, Element e1, Element e2)
        {
            if (!JoinGeometryUtils.AreElementsJoined(doc, e1, e2))
                JoinGeometryUtils.JoinGeometry(doc, e1, e2);
            else
            {
                JoinGeometryUtils.UnjoinGeometry(doc, e1, e2);
                JoinGeometryUtils.JoinGeometry(doc, e1, e2);
            }
        }

        // 執行接合並改變接合順序
        void RunJoinGeometryAndSwitch(Document doc, Element e1, Element e2)
        {
            if (!JoinGeometryUtils.AreElementsJoined(doc, e1, e2))
            {
                JoinGeometryUtils.JoinGeometry(doc, e1, e2);
                JoinGeometryUtils.SwitchJoinOrder(doc, e1, e2);
            }
            else
            {
                JoinGeometryUtils.UnjoinGeometry(doc, e1, e2);
                if (!JoinGeometryUtils.AreElementsJoined(doc, e1, e2))
                {
                    JoinGeometryUtils.JoinGeometry(doc, e1, e2);
                }
                JoinGeometryUtils.SwitchJoinOrder(doc, e1, e2);
            }
        }

        public static List<WallType> GetWallTypes(Autodesk.Revit.DB.Document doc)
        {
            List<WallType> oWallTypes = new List<WallType>();
            try
            {
                FilteredElementCollector collector
                    = new FilteredElementCollector(doc);

                FilteredElementIterator itor = collector
                    .OfClass(typeof(HostObjAttributes))
                    .GetElementIterator();

                // Reset the iterator
                itor.Reset();

                // Iterate through each family
                while (itor.MoveNext())
                {
                    Autodesk.Revit.DB.HostObjAttributes oSystemFamilies =
                    itor.Current as Autodesk.Revit.DB.HostObjAttributes;

                    if (oSystemFamilies == null) continue;

                    // Get the family's category
                    Category oCategory = oSystemFamilies.Category;
                    //TaskDialog.Show("1", oCategory.Name);

                    // Process if the category is found
                    if (oCategory != null)
                    {
                        if (oCategory.Name == "牆")
                        {
                            WallType oWallType = oSystemFamilies as WallType;
                            if (oWallType != null) oWallTypes.Add(oWallType);
                        }
                    }
                } //while itor.NextMove()

                return oWallTypes;
            }
            catch (Exception)
            {
                //MessageBox.Show( ex.Message );
                return oWallTypes = new List<WallType>();
            }
        } //GetWallTypes

        public double CentimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Centimeters);
        }

        public double UnitsToCentimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
        }

        private XYZ Cross(XYZ a, XYZ b)
        {

            XYZ c = new XYZ(a.Y * b.Z - a.Z * b.Y, a.Z * b.X - a.X * b.Z, a.X * b.Y - a.Y * b.X);
            return c;
        }

        private double Dot(XYZ a, XYZ b)
        {

            double c = a.X * b.X + a.Y * b.Y;
            return c;
        }

        public int GetWallWidth(Line line1, Line line2)
        {
            Line line = line1.Clone() as Line;
            XYZ ponit = line2.GetEndPoint(0);
            line.MakeUnbound();
            double width = line.Distance(ponit);
            int wallWidth = (int)Math.Round(UnitsToCentimeters(width) * 10);
            return wallWidth;
        }

        public Outline ReducedOutline(BoundingBoxXYZ bbxyz)
        {
            XYZ diagnalVector = (bbxyz.Max - bbxyz.Min) / 1000;
            Outline newOutline = new Outline(bbxyz.Min + diagnalVector, bbxyz.Max - diagnalVector);
            return newOutline;
        }

        public void ChangeWallType(Document doc, Wall wall, Line line1, Line line2)
        {
            WallType wallType = null;
            int width = GetWallWidth(line1, line2);
            String wallTypeName = "Generic - " + width.ToString() + "mm";

            FilteredElementCollector Collector = new FilteredElementCollector(doc);
            List<WallType> familySymbolList = Collector.OfClass(typeof(WallType))
            .OfCategory(BuiltInCategory.OST_Walls)
            .Cast<WallType>().ToList();
            Boolean IsWallTypeExist = false;
            foreach (WallType fs in familySymbolList)
            {
                if (fs.Name != wallTypeName)
                {
                    continue;
                }
                else
                {
                    IsWallTypeExist = true;
                    break;
                }
            }
            if (!IsWallTypeExist)
            {
                CreateWallType(doc, wallTypeName);
            }

            try
            {
                wallType = new FilteredElementCollector(doc)
                    .OfClass(typeof(WallType))
                    .First<Element>(x => x.Name.Equals(wallTypeName)) as WallType;
            }
            catch (Exception ex)
            {
                TaskDialog td = new TaskDialog("error")
                {
                    Title = "error",
                    AllowCancellation = true,
                    MainInstruction = "error",
                    MainContent = "Error" + ex.Message,
                    CommonButtons = TaskDialogCommonButtons.Close
                };
                td.Show();
                Debug.Print(ex.Message);
            }
            if (wallType != null)
            {
                Transaction t = new Transaction(doc, "Edit Type");
                t.Start();
                try
                {
                    wall.WallType = wallType;
                }
                catch (Exception ex)
                {
                    TaskDialog td = new TaskDialog("error")
                    {
                        Title = "error",
                        AllowCancellation = true,
                        MainInstruction = "error",
                        MainContent = "Error" + ex.Message,
                        CommonButtons = TaskDialogCommonButtons.Close
                    };
                    td.Show();
                    Debug.Print(ex.Message);
                }
                t.Commit();
            }

        }
        public void CreateWallType(Document doc, string wallTypeName)
        {
            string pattern = @"- (\d+)mm";
            double thickness = CentimetersToUnits(15);
            Match match = Regex.Match(wallTypeName, pattern);
            if (match.Success)
            {
                string tn = match.Groups[1].Value;
                thickness = Algorithm.MillimetersToUnits(double.Parse(tn));
            }
            Transaction tran = new Transaction(doc, "Create WallType!");
            tran.Start();

            WallType wallType = new FilteredElementCollector(doc)
                        .OfClass(typeof(WallType))
                        .FirstOrDefault(q => q.Name == "default") as WallType; // Replace "default" with the name of your wall type

            if (wallType != null)
            {
                // Duplicate the wall type and replace wallTypeName with your new wall type's name
                WallType newWallType = wallType.Duplicate(wallTypeName) as WallType;

                // Get the CompoundStructure of the new wall type
                CompoundStructure cs = newWallType.GetCompoundStructure();

                // Change the thickness of the first layer
                if (cs != null && cs.LayerCount > 0)
                {
                    cs.SetLayerWidth(0, thickness); // Change 0.2 to your desired thickness
                    newWallType.SetCompoundStructure(cs);
                }
            }
            else
            {
                //MessageBox.Show("The specified wall type could not be found.");
            }

            tran.Commit();
        }

        public List<Wall> GetWalls(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> Walls = collector.OfClass(typeof(Wall)).ToElements();
            List<Wall> ListWalls = new List<Wall>();
            foreach (Wall w in Walls)
            {
                ListWalls.Add(w);
            }
            return ListWalls;
        }

        static public List<Solid> GetElementSolidList(Element elem, Options opt = null, bool useOriginGeom4FamilyInstance = false)
        {
            if (null == elem)
            {
                return null;
            }
            if (null == opt)
                opt = new Options();
            opt.IncludeNonVisibleObjects = false;
            opt.DetailLevel = ViewDetailLevel.Medium;

            GeometryElement gElem;
            List<Solid> solidList = new List<Solid>();
            try
            {
                if (useOriginGeom4FamilyInstance && elem is FamilyInstance)
                {
                    // we transform the geometry to instance coordinate to reflect actual geometry 
                    FamilyInstance fInst = elem as FamilyInstance;
                    gElem = fInst.GetOriginalGeometry(opt);
                    Transform trf = fInst.GetTransform();
                    if (!trf.IsIdentity)
                        gElem = gElem.GetTransformed(trf);
                }
                else
                    gElem = elem.get_Geometry(opt);
                if (null == gElem)
                {
                    return null;
                }
                IEnumerator<GeometryObject> gIter = gElem.GetEnumerator();
                gIter.Reset();
                while (gIter.MoveNext())
                {
                    solidList.AddRange(getSolids(gIter.Current));
                }
            }
            catch (Exception ex)
            {
                // In Revit, sometime get the geometry will failed.
                string error = ex.Message;
            }
            return solidList;
        }

        static public List<Solid> getSolids(GeometryObject gObj)
        {
            List<Solid> solids = new List<Solid>();
            if (gObj is Solid) // already solid
            {
                Solid solid = gObj as Solid;
                if (solid.Faces.Size > 0 && Math.Abs(solid.Volume) > 0) // skip invalid solid
                    solids.Add(gObj as Solid);
            }
            else if (gObj is GeometryInstance) // find solids from GeometryInstance
            {
                IEnumerator<GeometryObject> gIter2 = (gObj as GeometryInstance).GetInstanceGeometry().GetEnumerator();
                gIter2.Reset();
                while (gIter2.MoveNext())
                {
                    solids.AddRange(getSolids(gIter2.Current));
                }
            }
            else if (gObj is GeometryElement) // find solids from GeometryElement
            {
                IEnumerator<GeometryObject> gIter2 = (gObj as GeometryElement).GetEnumerator();
                gIter2.Reset();
                while (gIter2.MoveNext())
                {
                    solids.AddRange(getSolids(gIter2.Current));
                }
            }
            return solids;
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
