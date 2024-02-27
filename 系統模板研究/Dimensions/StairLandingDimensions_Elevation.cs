using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Windows.Media.Animation;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Aspose.Cells.Charts;
using System.Net;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class StairLandingDimensions_Elevation : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Select stair to dimension
            Reference pickRef = uiapp.ActiveUIDocument.Selection.PickObject(ObjectType.Element, "Select stair to dimension");
            Element selectedElem = doc.GetElement(pickRef);

            if (selectedElem is Stairs)
            {
                Stairs selectedStair = selectedElem as Stairs;

                ICollection<ElementId> stairLandingsICollectionId = selectedStair.GetStairsLandings();
                // MessageBox.Show("平台:" + stairLandingsICollectionId.Count.ToString());

                //選到的樓梯裡的平台
                IList<Element> stairLandingsList = new List<Element>();
                foreach (ElementId stairLandingsId in stairLandingsICollectionId)
                {
                    Element elem = doc.GetElement(stairLandingsId);
                    if (elem != null)
                    {
                        stairLandingsList.Add(elem);
                    }
                }

                //建立標駐所需要的參考
                foreach (StairsLanding stairLanding in stairLandingsList)
                {
                    List<Face> stairLandingTopFace = GetFace(stairLanding).Item1;
                    List<Face> stairLandingBottomFace = GetFace(stairLanding).Item2;

                    List<Line> stairLandingTopLine = GetLine(stairLandingTopFace);
                    stairLandingTopLine = MergeLines(stairLandingTopLine, 0.00001);

                    List<Line> stairLandingBottomLine = GetLine(stairLandingBottomFace);
                    stairLandingBottomLine = MergeLines(stairLandingBottomLine, 0.00001);

                    Transaction tran = new Transaction(doc, "Create Dimension");
                    tran.Start();

                    foreach (Line line in stairLandingTopLine)
                    {
                        if (line.GetEndPoint(0).Y == line.GetEndPoint(1).Y)
                        {
                            // 創建一個平面
                            Plane plane = Plane.CreateByNormalAndOrigin(uidoc.ActiveView.ViewDirection, new XYZ(0, line.GetEndPoint(0).Y, 0));
                            SketchPlane sketchPlane = SketchPlane.Create(doc, plane);
                            Plane sketchPlanePlane = sketchPlane.GetPlane();
                            double parameter1 = sketchPlanePlane.Normal.DotProduct(line.GetEndPoint(0) - sketchPlanePlane.Origin);
                            double parameter2 = sketchPlanePlane.Normal.DotProduct(line.GetEndPoint(1) - sketchPlanePlane.Origin);
                            if (Math.Abs(parameter1) < 0.0001 && Math.Abs(parameter2) < 0.0001)
                            {
                                // 在 Family Editor 中創建一個 ModelCurve
                                ModelCurve modelCurve = doc.Create.NewModelCurve(line, sketchPlane);

                                XYZ offsetStartPoint = GetOffsetByStairLandingOrientation(modelCurve.GeometryCurve.GetEndPoint(0), 0.05);
                                XYZ offsetEndPoint = GetOffsetByStairLandingOrientation(modelCurve.GeometryCurve.GetEndPoint(1), 0.05);
                                Line offsetLine = Line.CreateBound(offsetStartPoint, offsetEndPoint);

                                ReferenceArray referenceArray = new ReferenceArray();
                                // MessageBox.Show(modelCurve.GeometryCurve.GetEndPointReference(0) + "\n" + modelCurve.GeometryCurve.GetEndPointReference(1));

                                referenceArray.Append(modelCurve.GeometryCurve.GetEndPointReference(0));
                                referenceArray.Append(modelCurve.GeometryCurve.GetEndPointReference(1));
                                // MessageBox.Show("參考數量:"+referenceArray.Size.ToString());
                                doc.Create.NewDimension(uidoc.ActiveView, offsetLine, referenceArray);
                            }
                        }
                    }

                    foreach (Line line in stairLandingBottomLine)
                    {
                        if (line.GetEndPoint(0).Y == line.GetEndPoint(1).Y)
                        {
                            // 創建一個平面
                            Plane plane = Plane.CreateByNormalAndOrigin(uidoc.ActiveView.ViewDirection, new XYZ(0, line.GetEndPoint(0).Y, 0));
                            SketchPlane sketchPlane = SketchPlane.Create(doc, plane);
                            Plane sketchPlanePlane = sketchPlane.GetPlane();
                            double parameter1 = sketchPlanePlane.Normal.DotProduct(line.GetEndPoint(0) - sketchPlanePlane.Origin);
                            double parameter2 = sketchPlanePlane.Normal.DotProduct(line.GetEndPoint(1) - sketchPlanePlane.Origin);
                            if (Math.Abs(parameter1) < 0.0001 && Math.Abs(parameter2) < 0.0001)
                            {
                                // 在 Family Editor 中創建一個 ModelCurve
                                ModelCurve modelCurve = doc.Create.NewModelCurve(line, sketchPlane);

                                XYZ offsetStartPoint = GetOffsetByStairLandingOrientation(modelCurve.GeometryCurve.GetEndPoint(0), -0.05);
                                XYZ offsetEndPoint = GetOffsetByStairLandingOrientation(modelCurve.GeometryCurve.GetEndPoint(1), -0.05);
                                Line offsetLine = Line.CreateBound(offsetStartPoint, offsetEndPoint);

                                ReferenceArray referenceArray = new ReferenceArray();
                                // MessageBox.Show(modelCurve.GeometryCurve.GetEndPointReference(0) + "\n" + modelCurve.GeometryCurve.GetEndPointReference(1));

                                referenceArray.Append(modelCurve.GeometryCurve.GetEndPointReference(0));
                                referenceArray.Append(modelCurve.GeometryCurve.GetEndPointReference(1));
                                // MessageBox.Show("參考數量:"+referenceArray.Size.ToString());
                                doc.Create.NewDimension(uidoc.ActiveView, offsetLine, referenceArray);
                            }
                        }
                    }
                    tran.Commit();
                }

                MessageBox.Show("建立標註完成");
            }

            else
            {
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        public static double UnitsToCentimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
        }
        //public Line GreateDimensionLine(List<Edge> sortedEdges, double value)
        //{
        //    Edge edge = sortedEdges[0];
        //    Line line = edge.AsCurve() as Line;
        //    XYZ lineDirection = (line.GetEndPoint(1) - line.GetEndPoint(0)).Normalize();
        //    XYZ offsetVector = lineDirection * value;

        //    // 偏移线的起点和终点
        //    XYZ offsetStartPoint = line.GetEndPoint(0) + offsetVector;
        //    XYZ offsetEndPoint = line.GetEndPoint(1) + offsetVector;

        //    Line dimensionLine = Line.CreateBound(offsetStartPoint, offsetEndPoint);
        //    return dimensionLine;
        //}

        private (List<Face>,List<Face>) GetFace(Element element)
        {
            GeometryElement geometryElement = element.get_Geometry(new Options());
            List<Face> topFace = new List<Face>();
            List<Face> bottomFace = new List<Face>();

            // 編歷Geometry對象以獲取尺寸信息
            foreach (GeometryObject geomObj in geometryElement)
            {
                // 抓取在實體上的幾何參數
                if (geomObj is Solid solid)
                {
                    // 抓取幾何面
                    FaceArray faces = solid.Faces;
                    // MessageBox.Show("平台面數量:" + faces.Size.ToString());

                    foreach (Face face in faces)
                    {
                        XYZ targetFaceNormal = face.ComputeNormal(UV.Zero);
                        XYZ zDirection = XYZ.BasisZ;
                        double dotProduct = zDirection.DotProduct(targetFaceNormal);
                        //Zdirection.DotProduct(targetFaceNormal))  == 1 或是 == -1

                        if (Math.Abs(dotProduct - 1.0) < 1e-9 )
                        {
                            topFace.Add(face);
                        }

                        if (Math.Abs(dotProduct + 1.0) < 1e-9)
                        {
                            bottomFace.Add(face);
                        }
                    }
                }
            }
            return (topFace, bottomFace);
        }

        private List<Line> GetLine(List<Face> faces)
        {
            List<Line> lineList = new List<Line>();
            foreach (Face face in faces)
            {
                EdgeArrayArray edgeArrays = face.EdgeLoops;
                EdgeArray edges = edgeArrays.get_Item(0);

                foreach (Edge edge in edges)
                {
                    Line line = edge.AsCurve() as Line;
                    lineList.Add(line);
                }
            }
            // MessageBox.Show("線數量:"+lineList.Count.ToString());
            return lineList;
        }
        public List<Line> MergeLines(List<Line> lines, double tolerance)
        {
            List<Line> mergedLines = new List<Line>();

            if (lines.Count > 0)
            {
                Line currentLine = lines[0];

                for (int i = 1; i < lines.Count; i++)
                {
                    Line nextLine = lines[i];

                    // 判斷相鄰線段是否在同一直線上
                    if (AreLinesOnSameLine(currentLine, nextLine, tolerance))
                    {
                        // 合併相鄰線段
                        currentLine = MergeTwoLines(currentLine, nextLine);
                    }
                    else
                    {
                        // 將合併後的線段添加到結果列表中
                        mergedLines.Add(currentLine);

                        // 將當前線段設置為下一條線段，準備進行下一輪比較
                        currentLine = nextLine;
                    }
                }

                // 將最後一條線段添加到結果列表中
                mergedLines.Add(currentLine);

                // 再次比較最後一條線段與第一條線段
                if (AreLinesOnSameLine(currentLine, lines[0], tolerance))
                {
                    // 合併相鄰線段
                    mergedLines[0] = MergeTwoLines(currentLine, lines[0]);
                    mergedLines.Remove(currentLine);
                }
            }

            //StringBuilder coordinatesStringBuilder = new StringBuilder();
            //foreach (Line line in mergedLines)
            //{
            //    XYZ startPoint = line.GetEndPoint(0);
            //    XYZ endPoint = line.GetEndPoint(1);

            //    // 將座標添加到 StringBuilder
            //    coordinatesStringBuilder.AppendLine($"SP: {startPoint}, \n EP: {endPoint}");
            //}
            //// 顯示所有座標在 TaskDialog 中
            //TaskDialog.Show("Line Coordinates", coordinatesStringBuilder.ToString());

            return mergedLines;
        }

        private bool AreLinesOnSameLine(Line line1, Line line2, double tolerance)
        {
            // 判斷兩條線段是否在同一直線上，這裡使用 XYZ 誤差作為 tolerance
            return line1.Direction.IsAlmostEqualTo(line2.Direction) &&
                   Math.Abs(line1.GetEndPoint(1).DistanceTo(line2.Origin)) < tolerance;
        }

        private Line MergeTwoLines(Line line1, Line line2)
        {
            // 合併兩條線段，這裡使用兩條線段的端點來創建一條新的線段
            XYZ startPoint = line1.GetEndPoint(0);
            XYZ endPoint = line2.GetEndPoint(1);

            return Line.CreateBound(startPoint, endPoint);
        }

        private XYZ GetOffsetByStairLandingOrientation(XYZ point/*, XYZ orientation*/, double value)
        {
            XYZ newVector = point.Multiply(value);
            XYZ returnPoint = point.Add(newVector);

            return returnPoint;
        }
    }
}
