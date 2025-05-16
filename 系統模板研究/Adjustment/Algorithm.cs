using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.DB;

namespace Modeling
{
    public class Algorithm
    {
        #region XYZ method

        public static XYZ RoundPoint(XYZ point, double gridSize)
        {
            XYZ newPoint = new XYZ(
                MillimetersToUnits(Math.Round(UnitsToMillimeters(point.X) / gridSize) * gridSize),
                MillimetersToUnits(Math.Round(UnitsToMillimeters(point.Y) / gridSize) * gridSize),
                MillimetersToUnits(Math.Round(UnitsToMillimeters(point.Z) / gridSize) * gridSize)
                );
            return newPoint;
        }

        public static double UnitsToCentimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
        }

        public static double CentimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Centimeters);
        }

        public static double MillimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Millimeters);
        }
        public static double UnitsToMillimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters);
        }

        public static double MmToFoot(double length) { return length / 304.8; }
        public static double FootToMm(double length) { return length * 304.8; }


        public static Curve ExtendCrv(Curve crv, double ratio)
        {
            double pstart = crv.GetEndParameter(0);
            double pend = crv.GetEndParameter(1);
            double pdelta = ratio * (pend - pstart);

            crv.MakeUnbound();
            crv.MakeBound(pstart - pdelta, pend + pdelta);
            return crv;
        }

        public static CurveArray AlignCrv(List<Curve> polylines)
        {
            List<Curve> imagelines = polylines.ToList();
            int lineNum = imagelines.Count;
            CurveArray polygon = new CurveArray();
            polygon.Append(imagelines[0]);
            imagelines.RemoveAt(0);
            while (polygon.Size < lineNum)
            {
                XYZ endPt = polygon.get_Item(polygon.Size - 1).GetEndPoint(1);
                for (int i = 0; i < imagelines.Count; i++)
                {
                    if (imagelines[i].GetEndPoint(0).IsAlmostEqualTo(endPt))
                    {
                        polygon.Append(imagelines[i]);
                        imagelines.Remove(imagelines[i]);
                        break;
                    }
                    if (imagelines[i].GetEndPoint(1).IsAlmostEqualTo(endPt))
                    {
                        polygon.Append(imagelines[i].CreateReversed());
                        imagelines.Remove(imagelines[i]);
                        break;
                    }
                }
            }
            return polygon;
        }



        /// <summary>
        /// Calculate the clockwise angle from vec1 to vec2
        /// </summary>
        /// <param name="vec1"></param>
        /// <param name="vec2"></param>
        /// <returns></returns>
        public static double AngleTo2PI(XYZ vec1, XYZ vec2)
        {
            double dot = vec1.X * vec2.X + vec1.Y * vec2.Y;    // dot product between [x1, y1] and [x2, y2]
            double det = vec1.X * vec2.Y - vec1.Y * vec2.X;    // determinant
            double angle = Math.Atan2(det, dot);  // Atan2(y, x) or atan2(sin, cos)
            return angle;
        }

        /// <summary>
        /// Return XYZ after axis system rotation
        /// </summary>
        /// <param name="pt"></param>
        /// <param name="angle"></param>
        /// <returns></returns>
        public static XYZ PtAxisRotation2D(XYZ pt, double angle)
        {
            double Xtrans = pt.X * Math.Cos(angle) + pt.Y * Math.Sin(angle);
            double Ytrans = pt.Y * Math.Cos(angle) - pt.X * Math.Sin(angle);
            return new XYZ(Xtrans, Ytrans, pt.Z);
        }

        /// <summary>
        /// Check if a point is on a line
        /// </summary>
        /// <param name="pt"></param>
        /// <param name="line"></param>
        /// <returns></returns>
        public static bool IsPtOnLine(XYZ pt, Line line)
        {
            XYZ ptStart = line.GetEndPoint(0);
            XYZ ptEnd = line.GetEndPoint(1);
            XYZ vec1 = (ptStart - pt).Normalize();
            XYZ vec2 = (ptEnd - pt).Normalize();
            if (!vec1.IsAlmostEqualTo(vec2) || pt.IsAlmostEqualTo(ptStart) || pt.IsAlmostEqualTo(ptEnd)) { return true; }
            else { return false; }
        }
        #endregion

        #region Curve method
        /// <summary>
        /// Check if two lines are perpendicular to each other.
        /// </summary>
        /// <param name="line1"></param>
        /// <param name="line2"></param>
        /// <returns></returns>
        public static bool IsPerpendicular(Curve line1, Curve line2)
        {
            XYZ vec1 = line1.GetEndPoint(1) - line1.GetEndPoint(0);
            XYZ vec2 = line2.GetEndPoint(1) - line2.GetEndPoint(0);
            if (vec1.DotProduct(vec2) == 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        
        // Check the parallel lines
        public static bool IsParallel(Curve line1, Curve line2)
        {
            if (line1.IsBound && line2.IsBound)
            {
                XYZ line1_Direction = (line1.GetEndPoint(1) - line1.GetEndPoint(0)).Normalize();
                XYZ line2_Direction = (line2.GetEndPoint(1) - line2.GetEndPoint(0)).Normalize();
                if (line1_Direction.IsAlmostEqualTo(line2_Direction) || line1_Direction.Negate().IsAlmostEqualTo(line2_Direction))
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }

            else
            {
                return false;
            }


        }


        /// <summary>
        /// Check if two curves are strictly intersected
        /// </summary>
        /// <param name="crv1"></param>
        /// <param name="crv2"></param>
        /// <returns></returns>
        public static bool IsIntersected(Curve crv1, Curve crv2)
        {
            // Can be safely apply to lines
            // Line segment can only have 4 comparison results: Disjoint, subset, overlap, equal
            SetComparisonResult result = crv1.Intersect(crv2, out _);
            if (result == SetComparisonResult.Overlap
                || result == SetComparisonResult.Subset
                || result == SetComparisonResult.Superset
                || result == SetComparisonResult.Equal)
            { return true; }
            else { return false; }
        }


        /// <summary>
        /// Check if two lines are almost joined.
        /// </summary>
        /// <param name="crv1"></param>
        /// <param name="crv2"></param>
        /// <returns></returns>
        public static bool IsAlmostJoined(Curve line1, Curve line2)
        {
            double radius = MmToFoot(50);
            XYZ ptStart = line1.GetEndPoint(0);
            XYZ ptEnd = line1.GetEndPoint(1);
            XYZ xAxis = new XYZ(1, 0, 0);   // The x axis to define the arc plane. Must be normalized
            XYZ yAxis = new XYZ(0, 1, 0);   // The y axis to define the arc plane. Must be normalized
            Curve knob1 = Arc.Create(ptStart, radius, 0, 2 * Math.PI, xAxis, yAxis);
            Curve knob2 = Arc.Create(ptEnd, radius, 0, 2 * Math.PI, xAxis, yAxis);
            SetComparisonResult result1 = knob1.Intersect(line2, out _);
            SetComparisonResult result2 = knob2.Intersect(line2, out _);
            if (result1 == SetComparisonResult.Overlap || result2 == SetComparisonResult.Overlap)
            { return true; }
            else { return false; }
        }




        /// <summary>
        /// Recreate a line in replace of the joining/overlapping lines
        /// </summary>
        /// <param name="lines"></param>
        /// <returns></returns>
        public static Curve FuseLines(List<Curve> lines)
        {
            double Z = lines[0].GetEndPoint(0).Z;
            List<XYZ> pts = new List<XYZ>();
            foreach (Curve line in lines)
            {
                pts.Add(line.GetEndPoint(0));
                pts.Add(line.GetEndPoint(1));
            }
            double Xmin = double.PositiveInfinity;
            double Xmax = double.NegativeInfinity;
            double Ymin = double.PositiveInfinity;
            double Ymax = double.NegativeInfinity;
            foreach (XYZ pt in pts)
            {
                if (pt.X < Xmin) { Xmin = pt.X; }
                if (pt.X > Xmax) { Xmax = pt.X; }
                if (pt.Y < Ymin) { Ymin = pt.Y; }
                if (pt.Y > Ymax) { Ymax = pt.Y; }
            }
            return Line.CreateBound(new XYZ(Xmin, Ymin, Z), new XYZ(Xmax, Ymax, Z));
        }

        public static bool AreOverlapping(Curve curve1, Curve curve2)
        {
            Line line1 = (Line)curve1;
            Line line2 = (Line)curve2;

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

        /// <summary>
        /// Cluster line segments if they were all piled upon a single line.
        /// </summary>
        /// <param name="crvs"></param>
        /// <returns></returns>
        public static List<List<Curve>> ClusterByParallel(List<Curve> crvs)
        {
            List<Curve> curveArray_List_copy = new List<Curve>(); // 複製得到的模型
            foreach (Curve curves in crvs)
            {
                curveArray_List_copy.Add(curves);
            }

            List<Curve> NotMatchCurveModel = new List<Curve>(); // 存放不匹配的梁線     
            List<List<Curve>> CurveModelList_List = new List<List<Curve>>();//存放模型List的List
            const double NormBeamWidth = 1000; //梁的最大寬度(mm) 

            while (crvs.Count > 0)
            {
                // 存放距離
                List<double> distanceList = new List<double>();
                // 存放對應距離的 CADModel
                List<Curve> curve_B_List = new List<Curve>();

                Curve Curve_A = crvs[0]; // A為欲被配對之模型線
                crvs.Remove(Curve_A); // 將A從原數組中去除

                if (crvs.Count >= 1)
                {
                    foreach (Curve Curve_B in crvs)
                    {
                        // 若梁的2個線不等長，最大誤差為50mm。方向為絕對值（sin120° = sin60°）
                        if (IsParallel(Curve_A, Curve_B)
                            && Math.Abs(Curve_A.Length - Curve_B.Length) < 50 * 0.003281
                            && AreOverlapping(Curve_A, Curve_B)) // 0.164 foot = 50mm * 0.003281
                        {
                            Curve Curve_A_copy = Curve_A.Clone();
                            //Curve_A_copy.MakeUnbound();
                            double distance = Curve_A_copy.Distance(Curve_B.GetEndPoint(0));
                            distanceList.Add(distance);
                            curve_B_List.Add(Curve_B);
                        }
                    }


                    if (distanceList.Count != 0 && curve_B_List.Count != 0)
                    {
                        double distanceTwoLine = distanceList.Min();
                        // 篩選不正常的寬度。若不正常則將CadModel_B放回數組
                        if (distanceTwoLine * 304.8 < NormBeamWidth && distanceTwoLine > 50 * 0.003281
                            && Curve_A.Length > 1 * 0.003281) // 梁寬介於10與"正常梁寬"之間的兩組模型線，為了避免配對到距離太遠的線
                        {
                            //將配對到的B模型線從原數組中去除
                            Curve Curve_shortDistance = curve_B_List[distanceList.IndexOf(distanceTwoLine)];
                            crvs.Remove(Curve_shortDistance);
                            // 1對配對好的梁的模型裝成一個數組
                            List<Curve> curveModels = new List<Curve>
                            {
                                Curve_A,
                                Curve_shortDistance
                            };
                            CurveModelList_List.Add(curveModels);
                            //TaskDialog.Show("1", CadModel_A.location.ToString() + "\n" + CadModel_shortDistance.location.ToString());

                        }
                    }
                    else
                    {
                        NotMatchCurveModel.Add(Curve_A);
                    }

                }
                else
                {
                    NotMatchCurveModel.Add(Curve_A);
                }

            }
            return CurveModelList_List;
            //return clusters;
        }

        /*
        /// <summary>
        /// Cluster lines if they are almost joined at end point.
        /// </summary>
        /// <param name="lines"></param>
        /// <returns></returns>
        public static List<List<Line>> ClusterByKnob(List<Line> lines)
        {
            List<List<Line>> clusters = new List<List<Line>> { };
            clusters.Add(new List<Line> { lines[0] });
            for (int i = 1; i < lines.Count; i++)
            {
                if (null == lines[i]) { continue; }
                foreach (List<Line> cluster in clusters)
                {
                    if (IsLineAlmostJoinedLines(lines[i], cluster))
                    {
                        cluster.Add(lines[i]);
                        goto a;
                    }
                }
                clusters.Add(new List<Line> { lines[i] });
            a:
                continue;
            }
            return clusters;
        }
        */

        /// <summary>
        /// Get non-duplicated points of a bunch of curves.
        /// </summary>
        /// <param name="crvs"></param>
        /// <returns></returns>
        public static List<XYZ> GetPtsOfCrvs(List<Curve> crvs)
        {
            List<XYZ> pts = new List<XYZ> { };
            foreach (Curve crv in crvs)
            {
                XYZ ptStart = crv.GetEndPoint(0);
                XYZ ptEnd = crv.GetEndPoint(1);
                pts.Add(ptStart);
                pts.Add(ptEnd);
            }
            for (int i = 0; i < pts.Count; i++)
            {
                for (int j = pts.Count - 1; j > i; j--)
                {
                    if (pts[i].IsAlmostEqualTo(pts[j]))
                    {
                        pts.RemoveAt(j);
                    }
                }
            }
            //Debug.Print("Vertices in all: " + pts.Count.ToString());
            return pts;
        }

        /// <summary>
        /// Calculate the distance of double lines
        /// </summary>
        /// <param name="line1"></param>
        /// <param name="line2"></param>
        /// <returns></returns>
        public static double LineSpacing(Line line1, Line line2)
        {
            XYZ midPt = line1.Evaluate(0.5, true);
            Line target = line2.Clone() as Line;
            target.MakeUnbound();
            double spacing = target.Distance(midPt);
            return spacing;
        }

        /// <summary>
        /// Generate axis by offset wall boundary
        /// </summary>
        /// <param name="line1"></param>
        /// <param name="line2"></param>
        /// <returns></returns>
        public static Line GenerateAxis(Line line1, Line line2)
        {
            Curve baseline = line1.Clone();
            Curve targetline = line2.Clone();
            if (line1.Length < line2.Length)
            {
                baseline = line2.Clone();
                targetline = line1.Clone();
            }
            targetline.MakeUnbound();
            XYZ midPt = baseline.Evaluate(0.5, true);
            XYZ midPt_proj = targetline.Project(midPt).XYZPoint;
            XYZ vec = (midPt_proj - midPt) / 2;
            double offset = vec.GetLength() / 2.0;
            //Debug.Print(offset.ToString());
            if (offset != 0)
            {
                Line axis = Line.CreateBound(baseline.GetEndPoint(0) + vec, baseline.GetEndPoint(1) + vec);
                return axis;
            }
            else
            {
                
                return null;
            }
            //Line axis = baseline.CreateOffset(offset, vec.Normalize()) as Line;
        }


        #endregion

        #region Region method

        /// <summary>
        /// Check if a bunch of lines enclose a rectangle.
        /// </summary>
        /// <param name="lines"></param>
        /// <returns></returns>
        public static bool IsRectangle(List<Curve> lines)
        {
            if (lines.Count() == 4)
            {
                if (GetPtsOfCrvs(lines).Count() == lines.Count())
                {
                    CurveArray edges = AlignCrv(lines);
                    if (IsPerpendicular(edges.get_Item(0), edges.get_Item(1)) &&
                        IsPerpendicular(edges.get_Item(1), edges.get_Item(2)))
                    { return true; }
                    else { return false; }
                }
                else { return false; }
            }
            else { return false; }
        }

        /// <summary>
        /// Return the bounding box of curves. 
        /// The box has the minimum area with axis in align with the curve direction.
        /// </summary>
        /// <param name="crvs"></param>
        /// <returns></returns>
        public static List<Curve> CreateBoundingBox2D(List<Curve> crvs)
        {
            // There can be a bounding box of an arc
            // but it is ambiguous to define the deflection of an arc-block
            // which is not the case as a door-block or window-block
            if (crvs.Count <= 1) { return null; }

            // Tolerance is to avoid generating boxes too small
            double tolerance = 0.001;
            List<XYZ> pts = GetPtsOfCrvs(crvs);
            double ZAxis = pts[0].Z;
            List<double> processions = new List<double> { };
            foreach (Curve crv in crvs)
            {
                // The Arc features no deflection of the door block
                if (crv.GetType().ToString() == "Autodesk.Revit.DB.Line")
                {
                    double angle = XYZ.BasisX.AngleTo(crv.GetEndPoint(1) - crv.GetEndPoint(0));
                    if (angle > Math.PI / 2)
                    {
                        angle = Math.PI - angle;
                    }
                    if (!processions.Contains(angle))
                    {
                        processions.Add(angle);
                    }
                }

            }
            //Debug.Print("Deflections in all: " + processions.Count.ToString());

            double area = double.PositiveInfinity;  // Mark the minimum bounding box area
            double deflection = 0;  // Mark the corresponding deflection angle
            double X0 = 0;
            double X1 = 0;
            double Y0 = 0;
            double Y1 = 0;
            foreach (double angle in processions)
            {
                double Xmin = double.PositiveInfinity;
                double Xmax = double.NegativeInfinity;
                double Ymin = double.PositiveInfinity;
                double Ymax = double.NegativeInfinity;
                foreach (XYZ pt in pts)
                {
                    double Xtrans = PtAxisRotation2D(pt, angle).X;
                    double Ytrans = PtAxisRotation2D(pt, angle).Y;
                    if (Xtrans < Xmin) { Xmin = Xtrans; }
                    if (Xtrans > Xmax) { Xmax = Xtrans; }
                    if (Ytrans < Ymin) { Ymin = Ytrans; }
                    if (Ytrans > Ymax) { Ymax = Ytrans; }
                }
                if (((Xmax - Xmin) * (Ymax - Ymin)) < area)
                {
                    area = (Xmax - Xmin) * (Ymax - Ymin);
                    deflection = angle;
                    X0 = Xmin;
                    X1 = Xmax;
                    Y0 = Ymin;
                    Y1 = Ymax;
                }
            }

            if (X1 - X0 < tolerance || Y1 - Y0 < tolerance)
            {
                Debug.Print("WARNING! Bounding box too small to be generated! ");
                return null;
            }

            else
            {
                // Inverse transformation
                XYZ pt1 = PtAxisRotation2D(new XYZ(X0, Y0, ZAxis), -deflection);
                XYZ pt2 = PtAxisRotation2D(new XYZ(X1, Y0, ZAxis), -deflection);
                XYZ pt3 = PtAxisRotation2D(new XYZ(X1, Y1, ZAxis), -deflection);
                XYZ pt4 = PtAxisRotation2D(new XYZ(X0, Y1, ZAxis), -deflection);
                Curve crv1 = Line.CreateBound(pt1, pt2) as Curve;
                Curve crv2 = Line.CreateBound(pt2, pt3) as Curve;
                Curve crv3 = Line.CreateBound(pt3, pt4) as Curve;
                Curve crv4 = Line.CreateBound(pt4, pt1) as Curve;
                List<Curve> boundingBox = new List<Curve> { crv1, crv2, crv3, crv4 };
                return boundingBox;
            }
        }

        // Center point of list of lines
        // Need upgrade to polygon center point method
        public static XYZ GetCenterPt(List<Curve> lines)
        {
            double ptSum_X = 0;
            double ptSum_Y = 0;
            double ptSum_Z = lines[0].GetEndPoint(0).Z;
            foreach (Line line in lines.Cast<Line>())
            {
                ptSum_X += line.GetEndPoint(0).X;
                ptSum_X += line.GetEndPoint(1).X;
                ptSum_Y += line.GetEndPoint(0).Y;
                ptSum_Y += line.GetEndPoint(1).Y;
            }
            XYZ centerPt = new XYZ(ptSum_X / lines.Count / 2, ptSum_Y / lines.Count / 2, ptSum_Z);
            return centerPt;
        }

        // Retrieve the width and depth of a rectangle
        public static Tuple<double, double, double> GetSizeOfRectangle(List<Line> lines)
        {
            List<double> rotations = new List<double> { };  // in radian
            List<double> lengths = new List<double> { };  // in milimeter
            foreach (Line line in lines)
            {
                XYZ vec = line.GetEndPoint(1) - line.GetEndPoint(0);
                double angle = AngleTo2PI(vec, XYZ.BasisX);
                //Debug.Print("Iteration angle is " + angle.ToString());
                rotations.Add(angle);
                lengths.Add( FootToMm(line.Length));
            }
            int baseEdgeId = rotations.IndexOf(rotations.Min());
            double width = lengths[baseEdgeId];
            double depth;
            if (width == lengths.Min()) { depth = lengths.Max(); }
            else { depth = lengths.Min(); }

            return Tuple.Create(Math.Round(width, 2), Math.Round(depth, 2), rotations.Min());
            // clockwise rotation in radian measure
            // x pointing right and y down as is common for computer graphics
            // this will mean you get a positive sign for clockwise angles
        }

        // 
        public static CurveArray RectifyPolygon(List<Line> lines)
        {
            CurveArray boundary = new CurveArray();
            List<XYZ> vertices = new List<XYZ>() { };
            vertices.Add(lines[0].GetEndPoint(0));
            foreach (Line line in lines)
            {
                XYZ ptStart = line.GetEndPoint(0);
                XYZ ptEnd = line.GetEndPoint(1);
                if (vertices.Last().IsAlmostEqualTo(ptStart))
                {
                    vertices.Add(ptEnd);
                    continue;
                }
                if (vertices.Last().IsAlmostEqualTo(ptEnd))
                {
                    vertices.Add(ptStart);
                    continue;
                }
            }
            /*
            Debug.Print("number of vertices: " + vertices.Count());
            foreach (XYZ pt in vertices)
            {
                Debug.Print(Util.PrintXYZ(pt));
            }
            */
            for (int i = 0; i < lines.Count; i++)
            {
                boundary.Append(Line.CreateBound(vertices[i], vertices[i + 1]));
            }
            return boundary;
        }


        public static List<Curve> CenterLinesOfBox(List<Curve> box)
        {
            XYZ centPt = GetCenterPt(box);
            List<Curve> centerLines = new List<Curve>();
            foreach (Line edge in box.Cast<Line>())
            {
                centerLines.Add(Line.CreateBound(centPt, edge.Evaluate(0.5, true)));
            }
            return centerLines;
        }

        public static Tuple<double, double, double> GetSizeOfFootprint(List<Line> lines)
        {
            if (lines is null)
            {
                throw new ArgumentNullException(nameof(lines));
            }

            return null;
        }
        #endregion

        // Here collects obsolete methods
        #region Trashbin
        /// <summary>
        /// Bubble sort algorithm. BUG FIXATION NEEDED!
        /// </summary>
        /// <param name="pts"></param>
        /// <returns></returns>
        public static List<XYZ> BubbleSort(List<XYZ> pts)
        {
            double threshold = 0.01;
            for (int i = 0; i < pts.Count(); i++)
            {
                for (int j = 0; j < pts.Count() - 1; j++)
                {
                    if (pts[j].X > pts[j + 1].X + threshold)
                    {
                        var ptTemp = pts[j];
                        pts[j] = pts[j + 1];
                        pts[j + 1] = ptTemp;
                    }
                }
            }
            for (int i = 0; i < pts.Count(); i++)
            {
                for (int j = 0; j < pts.Count() - 1; j++)
                {
                    if (pts[j].Y > pts[j + 1].Y + threshold)
                    {
                        var ptTemp = pts[j];
                        pts[j] = pts[j + 1];
                        pts[j + 1] = ptTemp;
                    }
                }
            }
            return pts;
        }
        #endregion
    }
}
