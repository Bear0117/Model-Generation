using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Documents;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateWalls : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            
            // To know the information of selected layer
            Reference refer = uidoc.Selection.PickObject(ObjectType.PointOnElement, "Select a CAD Layer");
            Element elem = doc.GetElement(refer);
            GeometryObject geoObj = elem.GetGeometryObjectFromReference(refer);
            ElementId graphicsStyleId = null;
            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                graphicsStyleId = geoObj.GraphicsStyleId;
            }

            // Initialize Parameters.
            ModelingParam.Initialize();
            double gridline_size = ModelingParam.parameters.General.GridSize * 10; // unit: mm
            double levelHeight = ModelingParam.parameters.General.LevelHeight; // unit: cm
            int[] wallWidths = ModelingParam.parameters.WallParam.WallWidths; // unit: mm

            // Get Level
            Level level = doc.GetElement(elem.LevelId) as Level;

            // Get lines
            Transaction tstart = new Transaction(doc, "Get wall line info!");
            tstart.Start();
            GeometryElement geoElem = elem.get_Geometry(new Options());

            CurveArray curveArray_h = new CurveArray();
            CurveArray curveArray_v = new CurveArray();
           

            foreach (GeometryObject gObj in geoElem)
            {
                GeometryInstance geomInstance = gObj as GeometryInstance;
                Transform transform = geomInstance.Transform;//座標轉換
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
                                if (Math.Abs(newLine.GetEndPoint(0).X - newLine.GetEndPoint(1).X) < CentimetersToUnits(1))
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
                                if (Math.Abs(line.GetEndPoint(0).X - line.GetEndPoint(1).X) < CentimetersToUnits(1))
                                    curveArray_v.Append(TransformLine(transform, line, gridline_size));
                                else
                                    curveArray_h.Append(TransformLine(transform, line, gridline_size));
                            }
                            XYZ normal = XYZ.BasisZ;
                            XYZ point = points.First();
                            point = transform.OfPoint(point);
                        }
                    }
                }
            }
            tstart.Commit();

            // Find Central Line.
            List<Tuple<Line, int>> central_lines_v = FindPairs(curveArray_v, true, wallWidths);
            List<Tuple<Line, int>> central_lines_h = FindPairs(curveArray_h, false, wallWidths);

            // Merge collinear Central Line.
            List<Tuple<Line, int>> merge_Lines_v = MergeCollinearLines(central_lines_v);
            List<Tuple<Line, int>> merge_Lines_h = MergeCollinearLines(central_lines_h);

            // Combine two List and process intersected lines.
            merge_Lines_v.AddRange(merge_Lines_h);
            List<Tuple<Line, int>> finalLines = AdjustLines(merge_Lines_v);

            // Create Walls.
            foreach (Tuple<Line, int> wallLines in finalLines)
            {
                // Create Wall.
                Transaction t1 = new Transaction(doc, "Create Wall");
                t1.Start();
                Wall wall = Wall.Create(doc, wallLines.Item1, level.Id, true);
                Parameter param = wall.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT);
                if (param != null && param.IsReadOnly == false)
                {
                    param.Set(0);  // 0 for structural and 1 for non-structural.
                }
                Parameter wallHeightP = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                wallHeightP.Set(CentimetersToUnits(levelHeight));
                WallUtils.DisallowWallJoinAtEnd(wall, 0);
                WallUtils.DisallowWallJoinAtEnd(wall, 1);
                t1.Commit();

                ChangeWallType(doc, wall, wallLines.Item2);

                List<double> depthList = new List<double>();

                //Get the element by elementID and get the boundingbox and outline of this element.
                Element elementWall = wall as Element;
                BoundingBoxXYZ bbxyzElement = elementWall.get_BoundingBox(null);
                //Outline outline = new Outline(bbxyzElement.Min, bbxyzElement.Max);
                Outline newOutline = ReducedOutline(bbxyzElement);

                //Create a filter to get all the intersection elements with wall.
                BoundingBoxIntersectsFilter filterW = new BoundingBoxIntersectsFilter(newOutline);

                //Create a filter to get StructuralFraming (which include beam and column) and Slabs.
                ElementCategoryFilter filterSt = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);
                ElementCategoryFilter filterSl = new ElementCategoryFilter(BuiltInCategory.OST_Floors);
                LogicalOrFilter filterS = new LogicalOrFilter(filterSt, filterSl);

                //Combine two filter.
                LogicalAndFilter filter = new LogicalAndFilter(filterS, filterW);

                //A list to store the intersected elements.
                IList<Element> inter = new FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements();

                for (int i = 0; i < inter.Count; i++)
                {
                    if (inter[i] != null)
                    {
                        string elementName = inter[i].Category.Name;
                        if (elementName == "結構構架")
                        {
                            // Find the depth
                            BoundingBoxXYZ bbxyzBeam = inter[i].get_BoundingBox(null);
                            depthList.Add(Math.Abs(bbxyzBeam.Min.Z - bbxyzBeam.Max.Z));

                            // Join
                            Transaction join = new Transaction(doc, "Join");
                            join.Start();
                            RunJoinGeometry(doc, wall, inter[i]);
                            join.Commit();
                        }
                        else if (elementName == "樓板")
                        {
                            BoundingBoxXYZ bbxyzSlab = inter[i].get_BoundingBox(null);
                            depthList.Add(Math.Abs(bbxyzSlab.Min.Z - bbxyzSlab.Max.Z));

                            Transaction join = new Transaction(doc, "Join");
                            join.Start();
                            RunJoinGeometry(doc, wall, inter[i]);
                            join.Commit();
                        }
                        else continue;
                    }
                }

                List<double> faceZ = new List<double>();
                List<Solid> list_solid = GetElementSolidList(wall);
                FaceArray faceArray;
                foreach (Solid solid in list_solid)
                {
                    faceArray = solid.Faces;
                    foreach (Face face in faceArray)
                    {
                        BoundingBoxUV boxUV = face.GetBoundingBox();
                        XYZ min = face.Evaluate(boxUV.Min);
                        XYZ max = face.Evaluate(boxUV.Max);
                        faceZ.Add(min.Z);
                        faceZ.Add(max.Z);
                    }
                    break;
                }

                Transaction t2 = new Transaction(doc, "Adjust Wall Height");
                t2.Start();
                wallHeightP.Set(faceZ.Max() - level.Elevation);
                t2.Commit();
            }
            return Result.Succeeded;
        }

        public List<Tuple<Line, int>> AdjustLines(List<Tuple<Line, int>> lines)
        {

            // 先遍歷到的線會退縮，由於組合垂直與水平線段的時候，水平線段是加到垂直線斷之後，所以垂直線會退縮。
            {
                // intersectionsExist = false;
                for (int i = 0; i < lines.Count; i++)
                {
                    for (int j = i + 1; j < lines.Count; j++)
                    {
                        IntersectionResultArray results;
                        SetComparisonResult compareResult = Line.CreateBound(lines[i].Item1.GetEndPoint(0), lines[i].Item1.GetEndPoint(1))
                            .Intersect(Line.CreateBound(lines[j].Item1.GetEndPoint(0), lines[j].Item1.GetEndPoint(1)), out results);

                        if (compareResult == SetComparisonResult.Overlap)
                        {
                            // intersectionsExist = true;
                            // Process the intersection
                            if (results != null && results.Size > 0)
                            {
                                XYZ intersection = results.get_Item(0).XYZPoint;
                                XYZ start = lines[i].Item1.GetEndPoint(0);
                                XYZ end = lines[i].Item1.GetEndPoint(1);
                                double distanceToStart = intersection.DistanceTo(start);
                                double distanceToEnd = intersection.DistanceTo(end);
                                XYZ fianl = new XYZ();
                                Line line_fianl;
                                // Adjust the lines based on their 'double' values
                                if (Math.Abs(start.X - end.X) < 0.001) // Vertical line
                                {
                                    if (distanceToStart < distanceToEnd)
                                    {
                                        fianl = 2 * intersection - start;
                                        if (fianl.DistanceTo(end) < CentimetersToUnits(0.1)) continue;
                                        line_fianl = Line.CreateBound(fianl, end);
                                        lines[i] = new Tuple<Line, int>(line_fianl, lines[i].Item2);
                                    }
                                    else
                                    {
                                        fianl = 2 * intersection - end;
                                        if (fianl.DistanceTo(start) < CentimetersToUnits(0.1)) continue;
                                        line_fianl = Line.CreateBound(fianl, start);
                                        lines[i] = new Tuple<Line, int>(line_fianl, lines[i].Item2);
                                    }
                                }
                                else if (Math.Abs(start.Y - end.Y) < 0.001) // Horizontal line
                                {
                                    XYZ start_2 = lines[j].Item1.GetEndPoint(0);
                                    XYZ end_2 = lines[j].Item1.GetEndPoint(1);
                                    double distanceToStart_2 = intersection.DistanceTo(start_2);
                                    double distanceToEnd_2 = intersection.DistanceTo(end_2);
                                    XYZ fianl_2 = new XYZ();

                                    if (distanceToStart_2 < distanceToEnd_2)
                                    {
                                        fianl_2 = 2 * intersection - start_2;
                                        if (fianl_2.DistanceTo(end_2) < CentimetersToUnits(0.1)) continue;
                                        line_fianl = Line.CreateBound(fianl_2, end_2);
                                        lines[j] = new Tuple<Line, int>(line_fianl, lines[j].Item2);
                                    }
                                    else
                                    {
                                        fianl_2 = 2 * intersection - end_2;
                                        if (fianl_2.DistanceTo(start_2) < CentimetersToUnits(0.1)) continue;
                                        line_fianl = Line.CreateBound(fianl_2, start_2);
                                        lines[j] = new Tuple<Line, int>(line_fianl, lines[j].Item2);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return lines;
        }

        public List<Tuple<Line, int>> MergeCollinearLines(List<Tuple<Line, int>> linePairs)
        {
            bool merged = true;
            while (merged)
            {
                merged = false;
                for (int i = 0; i < linePairs.Count; i++)
                {
                    Tuple<Line, int> pair1 = linePairs[i];
                    for (int j = i + 1; j < linePairs.Count; j++)
                    {
                        Tuple<Line, int> pair2 = linePairs[j];
                        if (AreCollinear(pair1.Item1, pair2.Item1) && AreOverlapping(pair1.Item1, pair2.Item1) && pair1.Item2 == pair2.Item2)
                        {
                            Line mergedLine = MergeLines(pair1.Item1, pair2.Item1);
                            Tuple<Line, int> newPair = new Tuple<Line, int>(mergedLine, pair1.Item2); // Use the same double value
                            linePairs[i] = newPair; // Replace the line in the first pair
                            linePairs.RemoveAt(j);  // Remove the second line
                            merged = true;
                            break;
                        }
                    }
                    if (merged) break; // For directly going to the next while loop.
                }
            }
            return linePairs;
        }

        private bool AreCollinear(Line line1, Line line2)
        {
            // 獲取兩條線的方向向量
            XYZ direction1 = (line1.GetEndPoint(1) - line1.GetEndPoint(0)).Normalize();
            XYZ direction2 = (line2.GetEndPoint(1) - line2.GetEndPoint(0)).Normalize();

            // 檢查方向是否平行或反向
            if (!direction1.IsAlmostEqualTo(direction2) && !direction1.IsAlmostEqualTo(-direction2))
                return false;

            // 檢查線 2 的一個端點是否位於線 1 的延長線上
            XYZ pointOnLine2 = line2.GetEndPoint(0); // 任取線 2 的一個端點
            XYZ pointOnLine1 = line1.GetEndPoint(0); // 線 1 的基準點

            // 向量從 line1 的起點到 line2 的某點
            XYZ vectorToPoint = pointOnLine2 - pointOnLine1;

            // 檢查向量是否平行於線 1 的方向（即共線）
            return vectorToPoint.CrossProduct(direction1).IsZeroLength();
        }


        private bool AreOverlapping(Line line1, Line line2)
        {
            // 判斷是否為垂直線
            bool isVertical = Math.Abs(line1.GetEndPoint(0).X - line1.GetEndPoint(1).X) < CentimetersToUnits(1);

            if (isVertical)
            {
                // 垂直線檢查：比較 Y 範圍
                double line1YMin = Math.Min(line1.GetEndPoint(0).Y, line1.GetEndPoint(1).Y);
                double line1YMax = Math.Max(line1.GetEndPoint(0).Y, line1.GetEndPoint(1).Y);
                double line2YMin = Math.Min(line2.GetEndPoint(0).Y, line2.GetEndPoint(1).Y);
                double line2YMax = Math.Max(line2.GetEndPoint(0).Y, line2.GetEndPoint(1).Y);

                return line1YMax > line2YMin - CentimetersToUnits(1) && line2YMax > line1YMin - CentimetersToUnits(1);
            }
            else
            {
                // 水平線檢查：比較 X 範圍
                double line1XMin = Math.Min(line1.GetEndPoint(0).X, line1.GetEndPoint(1).X);
                double line1XMax = Math.Max(line1.GetEndPoint(0).X, line1.GetEndPoint(1).X);
                double line2XMin = Math.Min(line2.GetEndPoint(0).X, line2.GetEndPoint(1).X);
                double line2XMax = Math.Max(line2.GetEndPoint(0).X, line2.GetEndPoint(1).X);

                return line1XMax > line2XMin - CentimetersToUnits(1) && line2XMax > line1XMin - CentimetersToUnits(1);
            }
        }


        private Line MergeLines(Line line1, Line line2)
        {
            List<XYZ> points = new List<XYZ> { line1.GetEndPoint(0), line1.GetEndPoint(1), line2.GetEndPoint(0), line2.GetEndPoint(1) };
            XYZ minPoint = points.OrderBy(p => p.X).ThenBy(p => p.Y).First();
            XYZ maxPoint = points.OrderByDescending(p => p.X).ThenByDescending(p => p.Y).First();

            return Line.CreateBound(minPoint, maxPoint);
        }

        public List<Tuple<Line, int>> FindPairs(CurveArray curves, bool vertical, int[] wallWidths)
        {
            List<Tuple<Line, int>> pairedLines = new List<Tuple<Line, int>>();
            // 轉換CurveArray到List<Line>，確保所有曲線都是Line類型
            List<Line> lines = new List<Line>();
            foreach (Curve curve in curves)
            {
                if (curve is Line line)
                {
                    lines.Add(line);
                }
            }
            if (vertical == true)
            {
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
                            if (wallWidths.Any(targetDistance => Math.Abs(Algorithm.UnitsToMillimeters(distance) - targetDistance) < 0.01))
                            {
                                Line wallLine = null;
                                double width = GetWallWidth(mainLine, compareLine);
                                int width_int = ((int)width);
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
                                Tuple<Line, int> paired = new Tuple<Line, int>(wallLine, width_int);
                                pairedLines.Add(paired);
                                // 如果條件符合，則將配對和距離加入列表
                                //pairedLines.Add(new Tuple<Line, Line, double>(mainLine, compareLine, distance));
                            }
                        }
                    }
                }

            }
            else
            {
                // 循環比較每一對線
                for (int i = 0; i < lines.Count; i++)
                {
                    Line mainLine = lines[i];
                    double mainXMin = Math.Min(mainLine.GetEndPoint(0).X, mainLine.GetEndPoint(1).X);
                    double mainXMax = Math.Max(mainLine.GetEndPoint(0).X, mainLine.GetEndPoint(1).X);

                    for (int j = 0; j < lines.Count; j++)
                    {
                        if (i == j)
                        {
                            continue;
                        }
                        Line compareLine = lines[j];
                        double lineXMin = Math.Min(compareLine.GetEndPoint(0).X, compareLine.GetEndPoint(1).X);
                        double lineXMax = Math.Max(compareLine.GetEndPoint(0).X, compareLine.GetEndPoint(1).X);

                        // 檢查X軸投影是否重疊
                        if (mainXMax - CentimetersToUnits(1) > lineXMin && lineXMax - CentimetersToUnits(1) > mainXMin)
                        {
                            // 檢查Y軸距離
                            double distance = Math.Abs(mainLine.GetEndPoint(0).Y - compareLine.GetEndPoint(0).Y);
                            if (wallWidths.Any(targetDistance => Math.Abs(Algorithm.UnitsToMillimeters(distance) - targetDistance) < 0.01))
                            {
                                Line wallLine = null;
                                double width = GetWallWidth(mainLine, compareLine); // Unit: mm
                                int width_int = ((int)width);
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
                                Tuple<Line, int> paired = new Tuple<Line, int>(wallLine, width_int);
                                pairedLines.Add(paired);
                                // 如果條件符合，則將配對和距離加入列表
                                //pairedLines.Add(new Tuple<Line, Line, double>(mainLine, compareLine, distance));
                            }
                        }
                    }
                }
            }

            return pairedLines;
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

        public void ChangeWallType(Document doc, Wall wall, int width)
        {
            WallType wallType = null;
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
