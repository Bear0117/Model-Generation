using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Management;
using System.Net;
using ClipperLib;
using System.IO;
using System.Windows;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class GapDetection : IExternalCommand
    {
        Document doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;

            ModelingParam.Initialize();
            double MaxdistMm = ModelingParam.parameters.General.Gap * 10; // unit: mm

            Dictionary<Face, ElementId> facesWithId = GetAllFaceElementMap(doc);
            //MessageBox.Show(facesWithId.Count.ToString());
            //MessageBox.Show(facesWithId.
            Dictionary<string, Dictionary<Face, ElementId>> faces_grouped = GroupFacesByNormal(facesWithId);
            //List<Face> faces = RunCheck(faces_grouped);

            List<(Face, ElementId)> faces = RunCheck(faces_grouped, MaxdistMm);
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string csvFilePath = Path.Combine(desktopPath, "ProblemFaces.csv");

            ExportPairsToCsv(doc, faces, csvFilePath);  
            //MessageBox.Show(faces.Count.ToString());

            //PaintFacesInRed painter = new PaintFacesInRed(doc, facesWithId, faces);
            //painter.Execute();


            return Result.Succeeded;
        }
        // The end of the main code.

        public void ExportPairsToCsv(Document doc, List<(Face face, ElementId elementId)> problemFaces, string csvFilePath)
        {
            // 若要避免覆蓋，這裡可判斷 File.Exists(csvFilePath)，看是要append還是 overwrite
            using (var sw = new StreamWriter(csvFilePath, false /*不append，直接覆蓋*/))
            {
                // 1. 寫標題列
                sw.WriteLine("Pairs, ElementId 1,ElementId 2");

                // 2. 逐兩筆輸出
                //    i = 0,2,4,6,...
                for (int i = 0; i < problemFaces.Count; i += 2)
                {
                    // 確保不會超過陣列範圍
                    if (i + 1 >= problemFaces.Count)
                        break;

                    var (face1, elem1) = problemFaces[i];
                    var (face2, elem2) = problemFaces[i + 1];

                    // Pair 編號
                    int pairIndex = (i / 2) + 1;

                    // FaceId 我們通常用「StableRepresentation」或至少用 face.Reference 的 Hash
                    //string faceId1 = GetFaceIdentifier(doc, face1);
                    //string faceId2 = GetFaceIdentifier(doc, face2);

                    // ElementId => elem.Id.IntegerValue
                    string elemId1 = elem1.IntegerValue.ToString() ?? "null";
                    string elemId2 = elem2.IntegerValue.ToString() ?? "null";

                    // 寫一行 CSV
                    // 注意格式中若含有逗號，可考慮加引號或使用 CSV 轉義
                    sw.WriteLine($"{pairIndex},{elemId1}, {elemId2}");
                }
            }
            MessageBox.Show("File Saved to Desktop!");
        }


        /// <summary>
        /// Collect all the Faces of column, beam, wall and slab in project.
        /// </summary>
        /// <param name="doc">Revit Document</param>
        /// <returns>Return a Dictionary，Key are Faces，Value are ElementIds.</returns>
        public Dictionary<Face, ElementId> GetAllFaceElementMap(Document doc)
        {
            Dictionary<Face, ElementId> faceElementMap = new Dictionary<Face, ElementId>();

            FilteredElementCollector columnCollector = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_Columns) ;
            FilteredElementCollector framingCollector = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).OfCategory(BuiltInCategory.OST_StructuralFraming);
            FilteredElementCollector floorCollector = new FilteredElementCollector(doc).OfClass(typeof(Floor));
            FilteredElementCollector wallCollector = new FilteredElementCollector(doc).OfClass(typeof(Wall));
            List<Element> allElements = columnCollector.Union(framingCollector).Union(floorCollector).Union(wallCollector).Where(e => e.OwnerViewId == ElementId.InvalidElementId).ToList();

            Options geomOptions = new Options
            {
                ComputeReferences = true,
                DetailLevel = ViewDetailLevel.Fine,
                IncludeNonVisibleObjects = false
            };

            //MessageBox.Show(allElements.Count().ToString());

            foreach (Element elem in allElements)
            {
                GeometryElement geomElem = elem.get_Geometry(geomOptions);
                if (geomElem == null) continue;

                foreach (GeometryObject geomObj in geomElem)
                {
                    if (geomObj is GeometryInstance geomInstance)
                    {
                        GeometryElement instGeomElem = geomInstance.GetSymbolGeometry(geomInstance.Transform);
                        if (instGeomElem != null)
                        {
                            foreach (GeometryObject instObj in instGeomElem)
                            {
                                if (instObj is Solid instSolid && instSolid.Volume > 0)
                                {
                                    foreach (Face face in instSolid.Faces)
                                    {
                                        faceElementMap[face] = elem.Id;
                                    }
                                }
                            }
                        }
                    }
                    else if (geomObj is Solid solid && solid.Volume > 0)
                    {
                        foreach (Face face in solid.Faces)
                        {
                            faceElementMap[face] = elem.Id;
                        }
                    }
                }
            }

            return faceElementMap;
        }

        /// <summary>
        /// 將 Dictionary<Face, ElementId> 依據 Face 的法向量分成 7 組 (±X, ±Y, ±Z, Others)
        /// </summary>
        public Dictionary<string, Dictionary<Face, ElementId>> GroupFacesByNormal(Dictionary<Face, ElementId> faceElementMap, double tolerance = 0.01)
        {
            // 建立七組 Dictionary<Face, ElementId>，用來存放不同方向的 face
            Dictionary<string, Dictionary<Face, ElementId>> faceNormalGroups = new Dictionary<string, Dictionary<Face, ElementId>>()
        {
            { "PositiveX", new Dictionary<Face, ElementId>() },
            { "NegativeX", new Dictionary<Face, ElementId>() },
            { "PositiveY", new Dictionary<Face, ElementId>() },
            { "NegativeY", new Dictionary<Face, ElementId>() },
            { "PositiveZ", new Dictionary<Face, ElementId>() },
            { "NegativeZ", new Dictionary<Face, ElementId>() },
            { "Others",    new Dictionary<Face, ElementId>() },
        };

            foreach (KeyValuePair<Face, ElementId> kvp in faceElementMap)
            {
                Face face = kvp.Key;
                ElementId eid = kvp.Value;

                // 取得法向量
                string directionKey = GetDirectionKey(face, tolerance);

                // 放進對應分組
                faceNormalGroups[directionKey].Add(face, eid);
            }

            return faceNormalGroups;
        }

        /// <summary>
        /// 根據 Face (若是 PlanarFace) 的法向量，判斷是 ±X, ±Y, ±Z 或 Others
        /// </summary>
        private string GetDirectionKey(Face face, double tolerance)
        {
            if (!(face is PlanarFace planarFace))
            {
                return "Others";
            }

            XYZ normal = planarFace.FaceNormal;
            normal = normal.Normalize();

            if (IsAlmostEqualTo(normal, XYZ.BasisX, tolerance))
                return "PositiveX";
            if (IsAlmostEqualTo(normal, -XYZ.BasisX, tolerance))
                return "NegativeX";
            if (IsAlmostEqualTo(normal, XYZ.BasisY, tolerance))
                return "PositiveY";
            if (IsAlmostEqualTo(normal, -XYZ.BasisY, tolerance))
                return "NegativeY";
            if (IsAlmostEqualTo(normal, XYZ.BasisZ, tolerance))
                return "PositiveZ";
            if (IsAlmostEqualTo(normal, -XYZ.BasisZ, tolerance))
                return "NegativeZ";
            return "Others";
        }

        /// <summary>
        /// 判斷兩個向量是否在指定誤差範圍內近似 (Revit 提供的 XYZ.IsAlmostEqualTo 也可用)
        /// </summary>
        private bool IsAlmostEqualTo(XYZ vecA, XYZ vecB, double tolerance)
        {
            XYZ diff = vecA - vecB;
            return diff.GetLength() < tolerance;
        }

        public List<(Face face, ElementId elementId)> RunCheck(Dictionary<string, Dictionary<Face, ElementId>> faces_grouped, double MaxDistanceMm)
        {
            // 1. 準備回傳的清單: (Face face, Element element)
            List<(Face face, ElementId elementId)> result = new List<(Face face, ElementId elementId)>();

            double MinDistanceMm = 2.5;   // mm
            double MinAreaThreshold = 0.05;  // mm² (若對應 1 cm²，請注意單位換算)

            // 要比對的法向量分組 (正對/負對)
            List<(string groupA, string groupB)> pairsToCheck = new List<(string groupA, string groupB)>
    {
        ("PositiveX", "NegativeX"),
        ("PositiveY", "NegativeY"),
        ("PositiveZ", "NegativeZ")
    };

            foreach ((string groupAKey, string groupBKey) in pairsToCheck)
            {
                if (!faces_grouped.ContainsKey(groupAKey) || !faces_grouped.ContainsKey(groupBKey))
                    continue;

                Dictionary<Face, ElementId> groupA = faces_grouped[groupAKey];
                Dictionary<Face, ElementId> groupB = faces_grouped[groupBKey];

                // 兩兩比對
                foreach (var kvpA in groupA)
                {
                    Face faceA = kvpA.Key;
                    ElementId idA = kvpA.Value;
                    //Element elementA = doc.GetElement(idA); // 轉成 Element

                    foreach (var kvpB in groupB)
                    {
                        Face faceB = kvpB.Key;
                        ElementId idB = kvpB.Value;
                        //Element elementB = doc.GetElement(idB);

                        if (idA == idB) // 同一元素，跳過
                            continue;

                        // (1) 計算投影後的面積 (mm²)
                        double projectedArea = CalculateProjectedIntersectionArea_new(faceA, faceB);
                        if (projectedArea < MinAreaThreshold)
                            continue;

                        // (2) 計算兩面之間的距離 (mm)
                        double distance = ComputeDistanceBetweenFaces(faceA, faceB);
                        if (distance >= MinDistanceMm && distance <= MaxDistanceMm)
                        {
                            // --> 代表面A 與 面B 疑似有問題
                            // 將 (faceA, elementA), (faceB, elementB) 都加到結果中
                            result.Add((faceA, idA));
                            result.Add((faceB, idB));
                        }
                    }
                }
            }

            return result; // 回傳 (Face, Element) 的清單
        }

        /// <summary>
        /// 若確認兩個面 (Face A, Face B) 都是平面，計算它們之間的「垂直距離」
        /// (假設它們近似平行)。回傳值單位為 mm。
        /// </summary>
        public double ComputeDistanceBetweenFaces(Face faceA, Face faceB)
        {
            // 1. 強制轉型為 PlanarFace (若失敗則拋出例外)
            PlanarFace pA = faceA as PlanarFace;
            PlanarFace pB = faceB as PlanarFace;
            if (pA == null || pB == null)
                throw new InvalidOperationException("One of the faces is not PlanarFace.");

            // 2. 取得 faceA 的平面方程: normalA, originA
            XYZ normalA = pA.FaceNormal.Normalize();
            XYZ originA = pA.Origin;
            // plane A: normalA · (X - originA) = 0

            // 3. 取得 faceB 的代表點 (面積中心或其他)
            XYZ centerB = pB.Origin;

            //CreateColumns(originA);
            //CreateColumns(centerB);

            // 4. 計算 "centerB" 到 planeA 的 signed distance
            double offsetB = (centerB - originA).DotProduct(normalA);

            // 垂直距離取絕對值
            double distanceInInternal = Math.Abs(offsetB);

            // 5. 轉換成 mm (依你的工具方法)
            return Algorithm.UnitsToMillimeters(distanceInInternal);
        }

        /// <summary>
        /// 假設這兩個面 (faceA, faceB) 都是平行的平面 (PlanarFace)，可能為多邊形。
        /// 計算它們重疊部分的面積 (mm²)。
        /// </summary>
        public double CalculateProjectedIntersectionArea_new(Face faceA, Face faceB)
        {
            // 1. 轉型檢查：確保都是平面
            PlanarFace pA = faceA as PlanarFace;
            PlanarFace pB = faceB as PlanarFace;
            if (pA == null || pB == null) return 0;

            // 2. 確認兩面法向量平行 (或反向)
            XYZ normalA = pA.FaceNormal.Normalize();
            XYZ normalB = pB.FaceNormal.Normalize();
            double dot = normalA.DotProduct(normalB);
            if (Math.Abs(Math.Abs(dot) - 1.0) > 0.01) return 0;

            // 3. 取得 Face A 的平面資訊 (planeA)
            //    - 原點 pA.Origin
            //    - 法向量 normalA
            //    - XVec, YVec 由 Revit 自動生成
            Plane planeA = Plane.CreateByNormalAndOrigin(normalA, pA.Origin);

            // 4. 取得兩個面的 2D 多邊形 (在 planeA 平面上)
            //    - faceAPolygon2D: Face A 本身的 2D 多邊形
            //    - faceBPolygon2D: Face B 投影到 planeA 之後的 2D 多邊形
            List<List<UV>> faceA2DLoops = GetFace2DPolygonLoops(pA, planeA);
            List<List<UV>> faceB2DLoops = GetFace2DPolygonLoops(pB, planeA);

            // 若面有多個邊界 (外環+內孔)，會有多個 Loop
            // 這裡你可以先把所有 Loop 合併或分別處理
            // 以下假設只有一個主要外環 (簡化示例)

            // 5. 計算多邊形交集面積 (2D)
            //    - 這裡只示範呼叫假函式
            //    - 實際需做多邊形布林運算 (intersection)
            double totalIntersectionArea = 0.0;
            if (faceA2DLoops.Count > 0 && faceB2DLoops.Count > 0)
            {
                // 這裡示範最簡單: 只拿第 0 個 loop 來做相交
                List<UV> polyA = faceA2DLoops[0];
                List<UV> polyB = faceB2DLoops[0];

                // 呼叫你自己的 2D 多邊形交集函式
                // (例如整合 Clipper、或自寫演算法)
                totalIntersectionArea = ComputePolygonIntersectionArea(polyA, polyB);
            }

            return totalIntersectionArea;
        }

        /// <summary>
        /// 將一個 PlanarFace 的邊界投影到 planeA (即 Face A 的參考平面)，
        /// 並以 UV 座標 (u, v) 表示。
        /// 回傳一組多邊形迴路 (可能包含外圈 & 內孔)。
        /// </summary>
        private List<List<UV>> GetFace2DPolygonLoops(PlanarFace face, Plane planeA)
        {
            List<List<UV>> polygonLoops2D = new List<List<UV>>();

            // 取得該面的所有邊界 Loop (外框、內孔等)
            // 有些狀況下可以使用 face.GetEdgesAsCurveLoops()，也可用 GetEdgesAsCurveLoops(FaceLoopType) 來分外內
            IList<CurveLoop> loops = face.GetEdgesAsCurveLoops();

            foreach (CurveLoop loop in loops)
            {
                List<UV> poly2D = new List<UV>();
                // 走訪 Loop 中的每條曲線
                foreach (Curve c in loop)
                {
                    // 離散化 (Tessellate) 或用更精準方法取多段
                    IList<XYZ> pts = c.Tessellate();

                    foreach (XYZ pt in pts)
                    {
                        // 投影此 pt 到 planeA (若 planeA == faceA 的 plane 而 face 與 planeA 平行，
                        //   其實只是把 pt 轉到 planeA 的 2D 坐標系)
                        UV uv = ProjectPointToPlaneUV(pt, planeA);
                        poly2D.Add(uv);
                    }
                }
                polygonLoops2D.Add(poly2D);
            }
            return polygonLoops2D;
        }

        /// <summary>
        /// 將 3D 點 (pt) 在指定 plane 上表達成 2D UV 座標。
        /// 假設 planeA 的 Normal, XVec, YVec 皆有效。
        /// </summary>
        private UV ProjectPointToPlaneUV(XYZ pt, Plane planeA)
        {
            // 1) 做正交投影 => p' (3D)
            //   p' = pt - [(pt - planeA.Origin) · planeA.Normal] * planeA.Normal
            XYZ origin = planeA.Origin;
            XYZ normal = planeA.Normal;
            XYZ vec = pt - origin;
            double dist = vec.DotProduct(normal);
            XYZ projected3D = pt - dist * normal;

            // 2) 轉成 2D => (u, v)
            double u = (projected3D - origin).DotProduct(planeA.XVec);
            double v = (projected3D - origin).DotProduct(planeA.YVec);
            return new UV(u, v);
        }

        /// <summary>
        /// 計算兩個 2D 多邊形 polyA, polyB 的交集面積 (mm²)。
        /// 實際上可用第三方庫 (Clipper, etc.)
        /// 這裡僅示範個假函式。
        /// </summary>
        /// <summary>
        /// 使用 Clipper 來計算 polyA 與 polyB 之間的「相交面積」(square units)。
        /// 注意：Clipper 需要整數座標，所以需先做縮放 (scaleFactor)。
        /// </summary>
        public double ComputePolygonIntersectionArea(List<UV> polyA, List<UV> polyB)
        {
            // (1) 決定縮放倍數
            //     數值越大，整數化誤差越小，但也要避免超過 long 的範圍。
            double scaleFactor = 1000000.0;

            // (2) 將 UV 多邊形轉成 Clipper 需要的 IntPoint
            List<IntPoint> clipPolyA = ConvertToClipperPolygon(polyA, scaleFactor);
            List<IntPoint> clipPolyB = ConvertToClipperPolygon(polyB, scaleFactor);

            //MessageBox.Show(clipPolyA.Count.ToString());
            //MessageBox.Show(clipPolyB.Count.ToString());

            // (3) 用 Clipper 執行多邊形「交集」(Intersection)
            Clipper clipper = new Clipper();
            // polyA 當 Subject
            clipper.AddPolygon(clipPolyA, PolyType.ptSubject);
            // polyB 當 Clip
            clipper.AddPolygon(clipPolyB, PolyType.ptClip);

            List<List<IntPoint>> solution = new List<List<IntPoint>>();
            // 執行交集，填充規則可用 NonZero 或 EvenOdd，看需求
            clipper.Execute(ClipType.ctIntersection, solution, PolyFillType.pftNonZero, PolyFillType.pftNonZero);

            // (4) 對交集結果中的所有多邊形，計算總面積 (在縮放前的整數座標系)
            double totalAreaScaled = 0.0;
            foreach (List<IntPoint> poly in solution)
            {
                // Clipper.Area(poly) 回傳多邊形面積，可能為負值(方向)，所以取絕對值
                double area = Clipper.Area(poly);
                totalAreaScaled += Math.Abs(area);
            }

            // (5) 將面積除以 (scaleFactor^2) 回復到原本座標尺度
            double realArea = totalAreaScaled / (scaleFactor * scaleFactor);
            return realArea;
        }

        /// <summary>
        /// 將 2D UV 座標 (double) 轉成 Clipper 使用的整數 IntPoint，並應用指定縮放。
        /// </summary>
        private List<IntPoint> ConvertToClipperPolygon(List<UV> polygon, double scaleFactor)
        {
            var result = new List<IntPoint>(polygon.Count);
            foreach (UV uv in polygon)
            {
                // 乘上縮放後，再四捨五入轉成 long
                long x = (long)Math.Round(uv.U * scaleFactor);
                long y = (long)Math.Round(uv.V * scaleFactor);

                result.Add(new IntPoint(x, y));
            }
            return result;
        }



        const double ScalingFactor = 10000000000000;
        public static double CalculateProjectedIntersectionArea(Face face1, Face face2)
        {
            Mesh mesh1 = face1.Triangulate();
            Mesh mesh2 = face2.Triangulate();
            string plane = DetermineProjectionPlane(face1.ComputeNormal(UV.Zero));
            List<IntPoint> poly1 = ProjectVerticesToPlane(mesh1, plane);
            List<IntPoint> poly2 = ProjectVerticesToPlane(mesh2, plane);

            Clipper clipper = new Clipper();
            clipper.AddPolygon(EnsureClockwise(poly1), PolyType.ptSubject);
            clipper.AddPolygon(EnsureClockwise(poly2), PolyType.ptClip);

            List<List<IntPoint>> solution = new List<List<IntPoint>>();
            clipper.Execute(ClipType.ctIntersection, solution, PolyFillType.pftNonZero, PolyFillType.pftNonZero);

            List<List<IntPoint>> outerPolygons = new List<List<IntPoint>>();
            List<List<IntPoint>> holePolygons = new List<List<IntPoint>>();

            foreach (var polyg in solution)
            {
                if (Clipper.Orientation(polyg))
                {
                    outerPolygons.Add(polyg);
                }
                else
                {
                    holePolygons.Add(polyg);
                }
            }

            List<List<XYZ>> intersectionPolygons = new List<List<XYZ>>();
            List<List<double>> edgeLengthsList = new List<List<double>>();

            double totalArea = 0;
            if (outerPolygons.Count > 0)
            {
                XYZ origin = GetRepresentativePoint(face1);
                foreach (var polygon in outerPolygons)
                {
                    List<XYZ> vertices = ProjectVerticesFromPlane(polygon, plane, origin);
                    List<double> edgeLengths = new List<double>();

                    for (int i = 0; i < polygon.Count; i++)
                    {
                        IntPoint currentPoint = polygon[i];
                        IntPoint nextPoint = polygon[(i + 1) % polygon.Count];

                        double length = Math.Sqrt(Math.Pow(currentPoint.X - nextPoint.X, 2) + Math.Pow(currentPoint.Y - nextPoint.Y, 2)) / ScalingFactor;
                        edgeLengths.Add(length);
                    }

                    intersectionPolygons.Add(vertices);
                    edgeLengthsList.Add(edgeLengths);

                    totalArea += Math.Abs(Clipper.Area(polygon)) / (ScalingFactor * ScalingFactor);
                }

                foreach (var hole in holePolygons)
                {
                    totalArea -= Math.Abs(Clipper.Area(hole)) / (ScalingFactor * ScalingFactor);
                }
            }

            return totalArea;
        }
        static private XYZ GetRepresentativePoint(Face face)
        {
            BoundingBoxUV bbox = face.GetBoundingBox();
            UV centerUV = new UV((bbox.Min.U + bbox.Max.U) / 2, (bbox.Min.V + bbox.Max.V) / 2);
            XYZ centerPoint = face.Evaluate(centerUV);
            return centerPoint;
        }
        private static List<IntPoint> ProjectVerticesToPlane(Mesh mesh, string plane)
        {
            List<IntPoint> projectedVertices = new List<IntPoint>();
            foreach (XYZ vertex in mesh.Vertices)
            {
                long x = (long)(vertex.X * ScalingFactor);
                long y = (long)(vertex.Y * ScalingFactor);
                long z = (long)(vertex.Z * ScalingFactor);

                switch (plane)
                {
                    case "XY":
                        projectedVertices.Add(new IntPoint(x, y));
                        break;
                    case "YZ":
                        projectedVertices.Add(new IntPoint(y, z));
                        break;
                    case "ZX":
                        projectedVertices.Add(new IntPoint(z, x));
                        break;
                    default:
                        throw new ArgumentException("Invalid plane specified. Use 'XY', 'YZ', or 'ZX'.");
                }
            }
            return projectedVertices;
        }
        private static List<XYZ> ProjectVerticesFromPlane(List<IntPoint> points, string plane, XYZ origin)
        {
            List<XYZ> vertices = new List<XYZ>();
            foreach (IntPoint point in points)
            {
                XYZ vertex;
                switch (plane)
                {
                    case "XY":
                        vertex = new XYZ(point.X / ScalingFactor, point.Y / ScalingFactor, origin.Z);
                        break;
                    case "YZ":
                        vertex = new XYZ(origin.X, point.X / ScalingFactor, point.Y / ScalingFactor);
                        break;
                    case "ZX":
                        vertex = new XYZ(point.Y / ScalingFactor, origin.Y, point.X / ScalingFactor);
                        break;
                    default:
                        throw new ArgumentException("Invalid plane specified. Use 'XY', 'YZ', or 'ZX'.");
                }
                vertices.Add(vertex);
            }
            return vertices;
        }
        private static List<IntPoint> EnsureClockwise(List<IntPoint> polygon)
        {
            if (!IsClockwise(polygon))
            {
                polygon.Reverse();
            }
            return polygon;
        }
        public static string DetermineProjectionPlane(XYZ normal)
        {
            double absX = Math.Abs(normal.X);
            double absY = Math.Abs(normal.Y);
            double absZ = Math.Abs(normal.Z);

            if (absX >= absY && absX >= absZ)
                return "YZ";  // X最大，YZ
            else if (absY >= absX && absY >= absZ)
                return "ZX";  // 
            else
                return "XY";  // Z，XY
        }
        private static bool IsClockwise(List<IntPoint> polygon)
        {
            double sum = 0;
            for (int i = 0; i < polygon.Count; i++)
            {
                IntPoint pt1 = polygon[i];
                IntPoint pt2 = polygon[(i + 1) % polygon.Count];
                sum += (pt2.X - pt1.X) * (pt2.Y + pt1.Y);
            }
            return sum > 0;
        }
    }

    public class PaintFacesInRed
    {
        private Document _doc;
        private Dictionary<Face, ElementId> _faceElementMap;
        private List<Face> _finalFaces;

        // 你想要輸出的 CSV 路徑
        private string _csvPath = @"C:\Temp\CannotPaintFaces.csv";

        public PaintFacesInRed(Document doc,
                               Dictionary<Face, ElementId> faceElementMap,
                               List<Face> finalFaces)
        {
            _doc = doc;
            _faceElementMap = faceElementMap;
            _finalFaces = finalFaces;
        }

        public void Execute()
        {
            // 1. 準備紅色材料 (若不存在就自動建立)
            ElementId redMatId = GetOrCreateRedMaterial(_doc, "MyRedMaterial", new Autodesk.Revit.DB.Color(255, 0, 0));

            // 2. 開啟 Transaction
            using (Transaction trans = new Transaction(_doc, "Paint Problematic Faces in Red"))
            {
                trans.Start();

                // 3. 準備 CSV 輸出 (標題列)
                if (!File.Exists(_csvPath))
                {
                    File.WriteAllText(_csvPath, "FaceHash,ElementId,Reason\r\n");
                }

                // 4. 逐一對每個 face 嘗試 Paint
                foreach (Face face in _finalFaces)
                {
                    if (!_faceElementMap.TryGetValue(face, out ElementId elemId))
                    {
                        // 找不到對應 Element => 寫入 CSV
                        AppendCsvRow(face, ElementId.InvalidElementId, "No ElementId in map");
                        continue;
                    }

                    // 取得 Reference 以便 Paint
                    Reference faceRef = face.Reference;
                    if (faceRef == null)
                    {
                        // 某些面可能沒有 Reference
                        AppendCsvRow(face, elemId, "face.Reference is null");
                        continue;
                    }

                    try
                    {
                        // 執行 Paint
                        _doc.Paint(elemId, face, redMatId);
                    }
                    catch (Exception ex)
                    {
                        // 可能拋出例外：例如該元素類型不允許上漆
                        AppendCsvRow(face, elemId, ex.Message);
                    }
                }

                trans.Commit();
            }
        }

        /// <summary>
        /// 建立或取得名為 materialName 的材料, 並將其設為指定顏色
        /// </summary>
        private ElementId GetOrCreateRedMaterial(Document doc, string materialName, Autodesk.Revit.DB.Color color)
        {
            // 先嘗試找同名材料
            FilteredElementCollector matCollector = new FilteredElementCollector(doc)
                .OfClass(typeof(Material));

            foreach (Material mat in matCollector)
            {
                if (mat.Name.Equals(materialName, StringComparison.OrdinalIgnoreCase))
                {
                    // 設定顏色後直接回傳
                    using (Transaction t = new Transaction(doc, "Update Material Color"))
                    {
                        t.Start();
                        mat.Color = color;
                        t.Commit();
                    }
                    return mat.Id;
                }
            }

            // 若找不到, 則建立一個新的
            Material newMat = null;
            using (Transaction t = new Transaction(doc, "Create Red Material"))
            {
                t.Start();
                ElementId newMatId = Material.Create(doc, materialName);
                newMat = doc.GetElement(newMatId) as Material;
                if (newMat != null)
                {
                    newMat.Color = color;
                }
                t.Commit();
            }

            return newMat?.Id ?? ElementId.InvalidElementId;
        }

        /// <summary>
        /// 在 CSV 中新增一列
        /// </summary>
        private void AppendCsvRow(Face face, ElementId eId, string reason)
        {
            // 你可以在這裡放更多資訊 (例如 face 的中心點, 法向量等)
            string faceInfo = face.GetHashCode().ToString();
            string elemIdInfo = eId.IntegerValue.ToString();
            string row = $"{faceInfo},{elemIdInfo},{reason}\r\n";

            File.AppendAllText(_csvPath, row);
        }
    }
}
