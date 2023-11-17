using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System;
using Curve = Autodesk.Revit.DB.Curve;
using System.Windows.Forms.VisualStyles;
using System.Diagnostics;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class TwoLineToWall_P : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("1F")) as Level;

            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(Level));

            while (true)
            {
                double modelLineElevation = 0; // default
                IList<XYZ> points = new List<XYZ>();
                for (int j = 0; j < 2; j++)
                {
                    Reference r = uidoc.Selection.PickObject(ObjectType.PointOnElement);
                    Element elem = doc.GetElement(r);
                    LocationCurve locationCurve = elem.Location as LocationCurve;
                    Line locationLine = locationCurve.Curve as Line;
                    points.Add(locationLine.GetEndPoint(0));
                    points.Add(locationLine.GetEndPoint(1));
                    modelLineElevation = locationLine.GetEndPoint(0).Z;
                }

                //Stop building the wall.
                if (points[0].ToString() == points[2].ToString())
                {
                    break;
                }

                double levelHeight = CentimetersToUnits(330); // default
                List<double> levelDistancesList = new List<double>();

                // retrieve all the Level elements in the document
                IEnumerable<Level> levels = collector.Cast<Level>();

                // loop through all the levels
                foreach (Level level_1 in levels)
                {
                    // get the level elevation.
                    double levelElevation = level_1.Elevation;
                    if (Math.Abs(levelElevation - modelLineElevation) < CentimetersToUnits(0.1))
                    {
                        level = level_1;
                    }
                    if (levelElevation - modelLineElevation > CentimetersToUnits(0.1))
                    {
                        levelDistancesList.Add(levelElevation);
                    }
                }

                if (levelDistancesList.Count() > 0)
                {
                    levelHeight = levelDistancesList.Min() - level.Elevation;
                }
                else
                {
                    LevelHeightForm form = new LevelHeightForm(doc);
                    form.ShowDialog();
                    levelHeight = CentimetersToUnits(form.levelHeight);
                }

                // To calculate the  central line of two lines. 
                Line Line1 = Line.CreateBound(points[0], points[1]);
                Line Line2 = Line.CreateBound(points[2], points[3]);
                XYZ midPoint1 = MidPoint(points[0], points[3]);
                XYZ midPoint2 = MidPoint(points[1], points[2]);

                double width = GetWallWidth(Line1, Line2); // Unit : mm

                Line wallLine = null;
                XYZ unitZ = new XYZ(0, 0, 1);
                XYZ vector_line1 = points[0] - points[1];
                XYZ offset = unitZ.CrossProduct(vector_line1) / Line1.Length * CentimetersToUnits(width / 10) / 2;
                if (Dot(offset, points[2] - points[1]) > 0)
                {
                    midPoint1 = points[0] + offset;
                    midPoint2 = points[1] + offset;
                    wallLine = Line.CreateBound(midPoint1, midPoint2);
                }
                else
                {
                    midPoint1 = points[0] - offset;
                    midPoint2 = points[1] - offset;
                    wallLine = Line.CreateBound(midPoint1, midPoint2);
                }

                FilteredElementCollector colWall = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .OfClass(typeof(Wall));
                IList<Wall> walls = new List<Wall>();
                foreach (Wall w in colWall.Cast<Wall>())
                {
                    walls.Add(w);
                }

                // Create Wall.
                Transaction t1 = new Transaction(doc, "Create Wall");
                t1.Start();
                Wall wall = Wall.Create(doc, wallLine, level.Id, true);
                Parameter wallHeightP = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                wallHeightP.Set(levelHeight);
                WallUtils.DisallowWallJoinAtEnd(wall, 0);
                WallUtils.DisallowWallJoinAtEnd(wall, 1);
                t1.Commit();
                ChangeWallType(doc, wall, Line1, Line2);


                List<double> depthList = new List<double>();

                //Get the element by elementID and get the boundingbox and outline of this element.
                Element elementWall = wall as Element;
                BoundingBoxXYZ bbxyzElement = elementWall.get_BoundingBox(null);
                Outline outline = new Outline(bbxyzElement.Min, bbxyzElement.Max);
                Outline newOutline = ReducedOutline_s(bbxyzElement);

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
                            Transaction joinAndSwitch = new Transaction(doc, "Join and Switch");
                            joinAndSwitch.Start();
                            RunJoinGeometryAndSwitch(doc, wall, inter[i]);
                            joinAndSwitch.Commit();
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

                XYZ lineAxis = wallLine.GetEndPoint(0) - wallLine.GetEndPoint(1);
                List<double> faceZ = new List<double>();
                List<Solid> list_solid = GetElementSolidList(wall);
                List<Face> list_vFace = new List<Face>();
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
                        if (Math.Abs(face.ComputeNormal(face.GetBoundingBox().Max).CrossProduct(lineAxis).GetLength()) < CentimetersToUnits(0.1))
                        {
                            list_vFace.Add(face);
                        }
                    }
                    break;
                }

                // Delete the original wall
                Transaction t_deletewall = new Transaction(doc, "Delete Wall");
                t_deletewall.Start();
                doc.Delete(wall.Id);
                t_deletewall.Commit();

                List<Line> lines = LineSeperatedByFaces(wallLine, list_vFace);

                foreach (Line line_new in lines)
                {
                    Transaction t1_new = new Transaction(doc, "Create Wall");
                    t1_new.Start();
                    Wall wall_new = Wall.Create(doc, line_new, level.Id, true);
                    Parameter wallHeightP_new = wall_new.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                    wallHeightP_new.Set(levelHeight);
                    WallUtils.DisallowWallJoinAtEnd(wall_new, 0);
                    WallUtils.DisallowWallJoinAtEnd(wall_new, 1);
                    t1_new.Commit();
                    ChangeWallType(doc, wall_new, Line1, Line2);

                    //Get the element by elementID and get the boundingbox and outline of this element.
                    Element elementWall_new = wall_new as Element;
                    BoundingBoxXYZ bbxyzElement_new = elementWall_new.get_BoundingBox(null);
                    Outline outline_new = new Outline(bbxyzElement_new.Min, bbxyzElement_new.Max);
                    Outline newOutline_new = ReducedOutline_s(bbxyzElement_new);

                    //Create a filter to get all the intersection elements with wall.
                    BoundingBoxIntersectsFilter filterW_new = new BoundingBoxIntersectsFilter(newOutline_new);

                    //Create a filter to get StructuralFraming (which include beam and column) and Slabs.
                    ElementCategoryFilter filterSt_new = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);
                    ElementCategoryFilter filterSl_new = new ElementCategoryFilter(BuiltInCategory.OST_Floors);
                    LogicalOrFilter filterS_new = new LogicalOrFilter(filterSt_new, filterSl_new);

                    //Combine two filter.
                    LogicalAndFilter filter_new = new LogicalAndFilter(filterS_new, filterW_new);

                    //A list to store the intersected elements.
                    IList<Element> inter_new = new FilteredElementCollector(doc).WherePasses(filter_new).WhereElementIsNotElementType().ToElements();

                    List<double> bottom_z = new List<double>();
                    List<double> depthList_new = new List<double>();
                    for (int i = 0; i < inter_new.Count; i++)
                    {
                        //MessageBox.Show(bottom_z.Count.ToString());
                        if (inter_new[i] != null)
                        {
                            string elementName = inter_new[i].Category.Name;
                            if (elementName == "結構構架")
                            {
                                // Find the depth
                                BoundingBoxXYZ bbxyzBeam = inter_new[i].get_BoundingBox(null);
                                depthList_new.Add(Math.Abs(bbxyzBeam.Min.Z - bbxyzBeam.Max.Z));
                                bottom_z.Add(bbxyzBeam.Min.Z);

                                // Join
                                Transaction joinAndSwitch = new Transaction(doc, "Join and Switch");
                                joinAndSwitch.Start();
                                RunJoinGeometryAndSwitch(doc, wall_new, inter_new[i]);
                                joinAndSwitch.Commit();
                            }
                            else if (elementName == "樓板")
                            {
                                BoundingBoxXYZ bbxyzSlab = inter_new[i].get_BoundingBox(null);
                                depthList_new.Add(Math.Abs(bbxyzSlab.Min.Z - bbxyzSlab.Max.Z));
                                bottom_z.Add(bbxyzSlab.Min.Z);

                                Transaction join = new Transaction(doc, "Join");
                                join.Start();
                                RunJoinGeometry(doc, wall_new, inter_new[i]);
                                join.Commit();
                            }
                            else continue;
                        }
                    }
                    bottom_z.Sort();

                    Face wallTopFace = null;
                    double topZ = double.MinValue;

                    Solid wall_solid = GetPrimarySolid(doc, wall_new as Element);
                    foreach (Face face in wall_solid.Faces)
                    {
                        XYZ normal = face.ComputeNormal(new UV(0, 0));
                        if (normal.Z > 0)
                        {
                            Mesh mesh = face.Triangulate();
                            foreach (XYZ point in mesh.Vertices)
                            {
                                if (point.Z > topZ)
                                {
                                    topZ = point.Z;
                                    wallTopFace = face;
                                }
                            }
                        }
                    }

                    bool isWallH = true;
                    if (Math.Abs(line_new.GetEndPoint(0).Y - line_new.GetEndPoint(1).Y) > CentimetersToUnits(1))
                    {
                        isWallH = false;
                    }

                    //MessageBox.Show(wall_solid.Faces.Size.ToString());

                    if (wall_solid.Faces.Size !=6)
                    {
                        (Line wallLine_top, Line line1, Line line2) = GetTopWallLine(doc, wallTopFace, isWallH);

                        for (int i = 0; i < inter_new.Count; i++)
                        {
                            if (inter_new[i] != null)
                            {
                                string elementName = inter_new[i].Category.Name;
                                if (elementName == "結構構架")
                                {
                                    Transaction joinAndSwitch = new Transaction(doc, "Join and Switch");
                                    joinAndSwitch.Start();
                                    UnJoinGeometry(doc, wall_new, inter_new[i]);
                                    joinAndSwitch.Commit();
                                }
                                else if (elementName == "樓板")
                                {
                                    Transaction join = new Transaction(doc, "Join");
                                    join.Start();
                                    UnJoinGeometry(doc, wall_new, inter_new[i]);
                                    join.Commit();
                                }
                                else continue;
                            }
                        }



                        Transaction t2 = new Transaction(doc, "Create Top Wall!");
                        t2.Start();
                        wallHeightP_new.Set(bottom_z[0] - level.Elevation);
                        WallUtils.DisallowWallJoinAtEnd(wall_new, 0);
                        WallUtils.DisallowWallJoinAtEnd(wall_new, 1);

                        Wall wall_top = Wall.Create(doc, wallLine_top, level.Id, true);
                        Parameter wallHeightP_top = wall_top.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                        wall_top.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).Set(bottom_z[0] - level.Elevation);
                        Level nextLevel = FindNextHigherLevel(doc, level);
                        if (bottom_z.Count != 1)
                        {
                            wall_top.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).Set(bottom_z[1] - bottom_z[0]);
                        }
                        else
                        {
                            wall_top.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).Set(nextLevel.Elevation - bottom_z[0]);
                        }
                        WallUtils.DisallowWallJoinAtEnd(wall_top, 0);
                        WallUtils.DisallowWallJoinAtEnd(wall_top, 1);
                        t2.Commit();
                        ChangeWallType(doc, wall_top, line1, line2);
                    }
                    else
                    {
                        for (int i = 0; i < inter_new.Count; i++)
                        {
                            if (inter_new[i] != null)
                            {
                                string elementName = inter_new[i].Category.Name;
                                if (elementName == "結構構架")
                                {
                                    Transaction joinAndSwitch = new Transaction(doc, "Join and Switch");
                                    joinAndSwitch.Start();
                                    UnJoinGeometry(doc, wall_new, inter_new[i]);
                                    joinAndSwitch.Commit();
                                }
                                else if (elementName == "樓板")
                                {
                                    Transaction join = new Transaction(doc, "Join");
                                    join.Start();
                                    UnJoinGeometry(doc, wall_new, inter_new[i]);
                                    join.Commit();
                                }
                                else continue;
                            }
                        } 

                        Transaction t2 = new Transaction(doc, "Create Top Wall!");
                        t2.Start();
                        if (bottom_z.Count != 0)
                        {
                            wallHeightP_new.Set(bottom_z[0] - level.Elevation);
                        }
                        else
                        {
                            wallHeightP_new.Set(FindNextHigherLevel(doc, level).Elevation - level.Elevation);
                        }
                        t2.Commit();
                    }
                }
            }
            return Result.Succeeded;
        }

        public (Line, Line, Line) GetTopWallLine(Document doc, Face face, bool isHorizontal)
        {
            // 检查面是否为矩形
            if (!(face is PlanarFace planarFace))
                throw new InvalidOperationException("The provided face is not planar.");

            EdgeArrayArray edgeLoops = face.EdgeLoops; // 获取面的边界循环
            if (edgeLoops.Size != 1) // 确保面是简单的矩形
                throw new InvalidOperationException("The face does not have a single outer loop.");

            // 对于矩形，只有一个EdgeLoop
            EdgeArray edgeArray = edgeLoops.get_Item(0);

            //Line1, Line2
            Line line1 = null;
            Line line2 = null;

            // 找到水平边和垂直边
            Edge horizontalEdge1 = null, horizontalEdge2 = null, verticalEdge1 = null, verticalEdge2 = null;
            foreach (Edge edge in edgeArray)
            {
                XYZ point1 = edge.AsCurve().GetEndPoint(0);
                XYZ point2 = edge.AsCurve().GetEndPoint(1);

                // 检查水平还是垂直边
                if (Math.Abs(point1.X - point2.X) < CentimetersToUnits(1)) // 垂直边
                {
                    if (verticalEdge1 == null)
                    {
                        verticalEdge1 = edge;

                    }
                    else
                    {
                        verticalEdge2 = edge;

                    }
                }
                else if (Math.Abs(point1.Y - point2.Y) < CentimetersToUnits(1)) // 水平边
                {
                    if (horizontalEdge1 == null)
                    {
                        horizontalEdge1 = edge;

                    }
                    else
                    {
                        horizontalEdge2 = edge;

                    }
                }
            }

            if (isHorizontal)
            {
                line1 = horizontalEdge1.AsCurve() as Line;
                line2 = horizontalEdge2.AsCurve() as Line;
            }
            else
            {
                line1 = verticalEdge1.AsCurve() as Line;
                line2 = verticalEdge2.AsCurve() as Line;
            }



            // 获取中点
            XYZ midPoint1, midPoint2;
            if (isHorizontal)
            {
                (midPoint1, midPoint2) = GetMidPoints(horizontalEdge1, horizontalEdge2);
                //midPoint1 = GetEdgeMidpoint(horizontalEdge1);
                //midPoint2 = GetEdgeMidpoint(horizontalEdge2);
            }
            else
            {
                (midPoint1, midPoint2) = GetMidPoints(verticalEdge1, verticalEdge2);
                //midPoint1 = GetEdgeMidpoint(verticalEdge1);
                //midPoint2 = GetEdgeMidpoint(verticalEdge2);
            }

            Line wallLineTop = Line.CreateBound(midPoint1, midPoint2);

            return (wallLineTop, line1, line2);

        }

        public (XYZ point1, XYZ point2) GetMidPoints(Edge edge1, Edge edge2)
        {
            XYZ point1 = edge1.AsCurve().GetEndPoint(0);
            XYZ point2 = edge2.AsCurve().GetEndPoint(0);
            XYZ point3 = edge1.AsCurve().GetEndPoint(1);
            XYZ point4 = edge2.AsCurve().GetEndPoint(1);
            XYZ midpoint1 = (point1 + point2) / 2;
            XYZ midpoint2 = (point3 + point4) / 2;
            if (midpoint1.DistanceTo(midpoint2) < CentimetersToUnits(1))
            {
                midpoint1 = (point1 + point4) / 2;
                midpoint2 = (point2 + point3) / 2;
            }
            return (midpoint1, midpoint2);
        }

        private static XYZ GetEdgeMidpoint(Edge edge)
        {
            XYZ endPoint0 = edge.AsCurve().GetEndPoint(0);
            XYZ endPoint1 = edge.AsCurve().GetEndPoint(1);
            return (endPoint0 + endPoint1) / 2; // 返回两点的平均位置作为中点
        }

        public Level FindNextHigherLevel(Document doc, Level currentLevel)
        {
            // 获取所有楼层
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> levels = collector.OfClass(typeof(Level)).ToElements();

            // 将楼层转换为Level类型并排序
            List<Level> sortedLevels = levels
                .Cast<Level>()
                .Where(lvl => lvl.Elevation > currentLevel.Elevation) // 高于当前楼层的楼层
                .OrderBy(lvl => lvl.Elevation) // 按高程排序
                .ToList();

            // 返回紧接在当前楼层之上的下一个楼层，如果存在的话
            return sortedLevels.FirstOrDefault(); // 如果没有更高的楼层，将返回null
        }

        public List<XYZ> RemoveDuplicatePoints(List<XYZ> points, double tolerance)
        {
            List<XYZ> distinctPoints = new List<XYZ>();

            foreach (XYZ pt in points)
            {
                bool isDuplicate = false;

                foreach (XYZ distinctPt in distinctPoints)
                {
                    if (IsApproximatelyEqual(pt, distinctPt, tolerance))
                    {
                        isDuplicate = true;
                        break;
                    }
                }

                if (!isDuplicate)
                {
                    distinctPoints.Add(pt);
                }
            }

            return distinctPoints;
        }

        public bool IsApproximatelyEqual(XYZ point1, XYZ point2, double tolerance)
        {
            double dist = point1.DistanceTo(point2);
            return dist < tolerance;
        }

        public List<XYZ> GetRectangleVertices(Face face)
        {
            HashSet<XYZ> uniqueVertices = new HashSet<XYZ>(new XYZEqualityComparer());
            if (face != null)
            {
                EdgeArrayArray edgeLoops = face.EdgeLoops;
                foreach (EdgeArray edgeArray in edgeLoops)
                {
                    foreach (Edge edge in edgeArray)
                    {
                        Curve curve = edge.AsCurve();
                        uniqueVertices.Add(curve.GetEndPoint(0));
                        uniqueVertices.Add(curve.GetEndPoint(1));
                    }
                }
            }

            if (uniqueVertices.Count != 4)
            {
                MessageBox.Show("uniqueVertices.Count != 4");
                //throw new InvalidOperationException("Expected 4 vertices for a rectangle.");
            }

            return uniqueVertices.ToList();
        }

        private class XYZEqualityComparer : IEqualityComparer<XYZ>
        {
            private const double Tolerance = 0.0001;

            public bool Equals(XYZ x, XYZ y)
            {
                return x.IsAlmostEqualTo(y, Tolerance);
            }

            public int GetHashCode(XYZ obj)
            {
                // Simplistic hash code generation. For more robust scenarios, consider a more involved method.
                return obj.X.GetHashCode() ^ obj.Y.GetHashCode() ^ obj.Z.GetHashCode();
            }
        }


        public List<Curve> GetRectangleEdges(List<XYZ> vertices)
        {
            List<XYZ> vertices_new = new List<XYZ>();
            foreach (XYZ point in vertices)
            {
                vertices_new.Add(new XYZ(point.X, point.Y, 0));
            }

            if (vertices_new.Count != 4)
            {
                MessageBox.Show("vertices_new.Count != 4");
                //throw new ArgumentException("Expected 4 vertices for a rectangle.");
            }

            // Here we sort the vertices based on their X and Y coordinates.
            // This assumes a standard orientation. Adjustments may be needed for different orientations.
            vertices_new.Sort((v1, v2) => v1.X != v2.X ? v1.X.CompareTo(v2.X) : v1.Y.CompareTo(v2.Y));

            List<Curve> edges = new List<Curve>();
            for (int i = 0; i < 4; i++)
            {
                // Connect each vertex to the next, with the last vertex connecting back to the first
                edges.Add(Line.CreateBound(vertices_new[i], vertices_new[(i + 1) % 4]));
            }

            return edges;
        }

        public bool CanProjectPointOntoFace(XYZ point, Face face)
        {
            IntersectionResult result = face.Project(point);
            return result != null;
        }

        private XYZ MidPoint(XYZ a, XYZ b)
        {
            XYZ c = (a + b) / 2;
            return c;
        }

        void UnJoinGeometry(Document doc, Element e1, Element e2)
        {
            if (!JoinGeometryUtils.AreElementsJoined(doc, e1, e2)) { }
            else
            {
                JoinGeometryUtils.UnjoinGeometry(doc, e1, e2);
            }
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
                if ((!JoinGeometryUtils.AreElementsJoined(doc, e1, e2)))
                {
                    JoinGeometryUtils.JoinGeometry(doc, e1, e2);
                }
                JoinGeometryUtils.SwitchJoinOrder(doc, e1, e2);
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

        public Outline ReducedOutline_s(BoundingBoxXYZ bbxyz)
        {
            XYZ diagnalVector = (bbxyz.Max - bbxyz.Min) / 100;
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
            catch (System.Exception ex)
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
                catch (System.Exception ex)
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


        public static List<Solid> GetElementSolidList(Element elem, Options opt = null, bool useOriginGeom4FamilyInstance = false)
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
                    Autodesk.Revit.DB.Transform trf = fInst.GetTransform();
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
                    solidList.AddRange(GetSolids(gIter.Current));
                }
            }
            catch (System.Exception ex)
            {
                // In Revit, sometime get the geometry will failed.
                string error = ex.Message;
            }
            return solidList;
        }

        public static List<Solid> GetSolids(GeometryObject gObj)
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
                    solids.AddRange(GetSolids(gIter2.Current));
                }
            }
            else if (gObj is GeometryElement) // find solids from GeometryElement
            {
                IEnumerator<GeometryObject> gIter2 = (gObj as GeometryElement).GetEnumerator();
                gIter2.Reset();
                while (gIter2.MoveNext())
                {
                    solids.AddRange(GetSolids(gIter2.Current));
                }
            }
            return solids;
        }

        public static XYZ ProjectPointOntoLine(XYZ point, Line line)
        {
            XYZ lineStart = line.GetEndPoint(0);
            XYZ lineEnd = line.GetEndPoint(1);
            XYZ lineDirection = (lineEnd - lineStart).Normalize();
            XYZ pointVector = point - lineStart;
            double dotProduct = pointVector.DotProduct(lineDirection);
            XYZ projectedVector = dotProduct * lineDirection;
            return lineStart + projectedVector;
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
                MessageBox.Show("The specified wall type could not be found.");
            }

            tran.Commit();
        }

        public class XyzEqualityComparer : IEqualityComparer<XYZ>
        {
            public bool Equals(XYZ x, XYZ y)
            {
                return x.IsAlmostEqualTo(y);
            }

            public int GetHashCode(XYZ obj)
            {
                unchecked // Overflow is fine, just wrap
                {
                    int hash = 17;
                    hash = hash * 23 + obj.X.GetHashCode();
                    hash = hash * 23 + obj.Y.GetHashCode();
                    hash = hash * 23 + obj.Z.GetHashCode();
                    return hash;
                }
            }
        }

        public List<XYZ> RemoveDuplicates(List<XYZ> listWithDuplicates)
        {
            return listWithDuplicates.Distinct(new XyzEqualityComparer()).ToList();
        }

        public bool HaveSameXorY(XYZ point1, XYZ point2)
        {
            int trueCount = 0;

            const double tolerance = 0.0001;

            bool sameX = Math.Abs(point1.X - point2.X) < tolerance;
            if (sameX) trueCount++;
            bool sameY = Math.Abs(point1.Y - point2.Y) < tolerance;
            if (sameY) trueCount++;
            return trueCount == 1;
        }



        public Line GetLongestLineFromFourPoints(List<XYZ> points)
        {
            double maxDistance = double.MinValue;
            XYZ startPoint = null;
            XYZ endPoint = null;

            for (int i = 0; i < points.Count - 1; i++)
            {
                for (int j = i + 1; j < points.Count; j++)
                {
                    double currentDistance = points[i].DistanceTo(points[j]);
                    if (currentDistance > maxDistance)
                    {
                        maxDistance = currentDistance;
                        startPoint = points[i];
                        endPoint = points[j];
                    }
                }
            }

            if (startPoint != null && endPoint != null)
            {
                return Line.CreateBound(startPoint, endPoint);
            }

            return null; // This shouldn't happen with valid points.
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

        public List<Line> LineSeperatedByFaces(Line line, List<Face> list_face)
        {
            //Boolean xline;
            //if(Math.Abs(line.GetEndPoint(0).Y - line.GetEndPoint(1).Y) < CentimetersToUnits(0.1))
            //{
            //    xline = true;
            //}
            //else
            //{
            //    xline = false;
            //}

            List<XYZ> pointsOnLine = new List<XYZ>();
            if (true)
            {
                foreach (Face face in list_face)
                {
                    IList<CurveLoop> faceEdges = face.GetEdgesAsCurveLoops().ToList();
                    CurveLoop curveLoop = faceEdges[0];
                    foreach (Curve curve in curveLoop)
                    {
                        XYZ point = curve.GetEndPoint(0);
                        XYZ pointOnLine = ProjectPointOntoLine(point, line);
                        pointsOnLine.Add(pointOnLine);
                        break;
                    }
                }
            }

            List<XYZ> pointsOnLine_sort = pointsOnLine.OrderBy(p => p.DistanceTo(line.GetEndPoint(0))).ToList();
            List<Line> lines = new List<Line>();
            for (int i = 0; i < pointsOnLine_sort.Count() - 1; i++)
            {
                if (pointsOnLine_sort[i].DistanceTo(pointsOnLine_sort[i + 1]) > CentimetersToUnits(0.1))
                {
                    Line line_s = Line.CreateBound(pointsOnLine_sort[i], pointsOnLine_sort[i + 1]);
                    if (line_s.Length > CentimetersToUnits(1))
                    {
                        lines.Add(line_s);
                    }
                }
            }
            return lines;
        }
    }
    public class RectangleCornerFinder
    {
        private readonly List<XYZ> _points;
        private readonly double _tolerance;

        public RectangleCornerFinder(List<XYZ> points, double tolerance = 0.0001)
        {
            _points = points ?? throw new ArgumentNullException(nameof(points));
            _tolerance = tolerance;
        }

        public List<XYZ> GetRectangleCorners()
        {
            List<XYZ> corners = new List<XYZ>();

            foreach (var point in _points)
            {
                List<double> distances = new List<double>();
                foreach (var other in _points)
                {
                    if (point.IsAlmostEqualTo(other, _tolerance)) continue; // 跳過自己
                    distances.Add(point.DistanceTo(other));
                }

                // 如果這個點的最大距離大於其他點的最大距離，那麼它可能是一個角落
                if (Math.Abs(distances.Max() - distances.OrderByDescending(d => d).First()) < _tolerance)
                {
                    corners.Add(point);
                }
            }

            // 確保只有四個點被選為角落
            if (corners.Count != 4)
            {
                throw new InvalidOperationException("找到的角點數量不是4個");
            }

            return corners;
        }
    }
}