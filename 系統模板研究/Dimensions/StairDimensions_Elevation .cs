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
// using Aspose.Pdf.LogicalStructure;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class StairDimensions_Elevation : IExternalCommand
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
                ICollection<ElementId> stairRunsICollectionId = selectedStair.GetStairsRuns();

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

                //選到的樓梯裡的梯段
                IList<Element> stairRunsList = new List<Element>();
                foreach (ElementId stairRunsId in stairRunsICollectionId)
                {
                    Element elem = doc.GetElement(stairRunsId);
                    if (elem != null)
                    {
                        stairRunsList.Add(elem);
                    }
                }

                List<Line> stairHorrizontalLines = new List<Line>();
                List<Line> stairVerticalLines = new List<Line>();
                List<Line> stairSlopeLines = new List<Line>();
                List<Line> stairAllLines = new List<Line>();
                
                //抓取平台邊緣線並分類
                foreach (StairsLanding stairLanding in stairLandingsList)
                {
                    ViewSection sectionView = uidoc.ActiveView as ViewSection;

                    List<Face> stairLandingFace = GetFace(stairLanding, sectionView.ViewDirection);

                    List<Line> stairLandingHorrizontalLine = GetAndClassifyLine(stairLandingFace).Item1;
                    List<Line> stairLandingVerticalLine = GetAndClassifyLine(stairLandingFace).Item2;
                    List<Line> stairLandingSlopeLine = GetAndClassifyLine(stairLandingFace).Item3;

                    stairHorrizontalLines = AddLine(stairHorrizontalLines, stairLandingHorrizontalLine);
                    stairVerticalLines = AddLine(stairVerticalLines, stairLandingVerticalLine);
                    stairSlopeLines = AddLine(stairSlopeLines, stairLandingSlopeLine);
                }
                
                foreach (StairsRun stairRun in stairRunsList)
                {
                    ViewSection sectionView = uidoc.ActiveView as ViewSection;

                    List<Face> stairRunFace = GetFace(stairRun, sectionView.ViewDirection);

                    List<Line> stairRunHorrizontalLine = GetAndClassifyLine(stairRunFace).Item1;
                    List<Line> stairRunVerticalLine = GetAndClassifyLine(stairRunFace).Item2;
                    List<Line> stairRunSlopeLine = GetAndClassifyLine(stairRunFace).Item3;

                    stairHorrizontalLines = AddLine(stairHorrizontalLines, stairRunHorrizontalLine);
                    stairVerticalLines = AddLine(stairVerticalLines, stairRunVerticalLine);
                    stairSlopeLines = AddLine(stairSlopeLines, stairRunSlopeLine);
                }

                stairHorrizontalLines = AlignAndSortLines(stairHorrizontalLines, XYZ.BasisX, point => point.X);
                stairHorrizontalLines = AlignAndSortLines(stairHorrizontalLines, XYZ.BasisX, point => point.Z);
                stairVerticalLines = AlignAndSortLines(stairVerticalLines, XYZ.BasisX, point => point.X);
                stairVerticalLines = AlignAndSortLines(stairVerticalLines, XYZ.BasisX, point => point.Z);
                stairSlopeLines = AlignAndSortLines(stairSlopeLines, XYZ.BasisX, point => point.Z);
                stairSlopeLines = AlignAndSortLines(stairSlopeLines, XYZ.BasisX, point => point.X);
                stairSlopeLines = AlignAndSortLines(stairSlopeLines, XYZ.BasisX, point => point.Y);

                //StringBuilder coordinatesStringBuilder = new StringBuilder();
                //foreach (Line line in stairHorrizontalLines)
                //{
                //    XYZ startPoint = line.GetEndPoint(0);
                //    XYZ endPoint = line.GetEndPoint(1);

                //    // 將座標添加到 StringBuilder
                //    coordinatesStringBuilder.AppendLine($"SP: {startPoint}, \n EP: {endPoint}");
                //}
                //// 顯示所有座標在 TaskDialog 中
                //TaskDialog.Show("Line Coordinates", coordinatesStringBuilder.ToString());


                stairHorrizontalLines = RemoveSameLines(stairHorrizontalLines);
                stairVerticalLines = RemoveSameLines(stairVerticalLines);
                stairSlopeLines = RemoveSameLines(stairSlopeLines);
                
                stairHorrizontalLines = MergeLines(stairHorrizontalLines, 0.00001);
                stairVerticalLines = MergeLines(stairVerticalLines, 0.00001);
                stairSlopeLines = MergeLines(stairSlopeLines, 0.00001);
                
                stairAllLines = AddLine(stairAllLines, stairHorrizontalLines);
                //stairAllLines = AddLine(stairAllLines, stairVerticalLines);
                stairAllLines = AddLine(stairAllLines, stairSlopeLines);
                

                Transaction tran = new Transaction(doc, "Create Dimension");
                tran.Start();

                foreach (Line line in stairAllLines)
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

                            XYZ offsetStartPoint = GetOffsetByStairOrientation(modelCurve.GeometryCurve.GetEndPoint(0), 0.05);
                            XYZ offsetEndPoint = GetOffsetByStairOrientation(modelCurve.GeometryCurve.GetEndPoint(1), 0.05);
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

                MessageBox.Show("建立標註完成");
            }

            else
            {
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        private double UnitsToCentimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
        }
        
        private List<Face> GetFace(Element element, XYZ viewDirection)
        {
            GeometryElement geometryElement = element.get_Geometry(new Options());
            List<Face> newFace = new List<Face>();

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
                        if (targetFaceNormal.DotProduct(viewDirection) > 0)
                        {
                            newFace.Add(face);
                        }
                    }
                }
            }
            return newFace;
        }
        private (List<Line>, List<Line>, List<Line>) GetAndClassifyLine(List<Face> faces)
        {
            List<Line> verticalLineList = new List<Line>();
            List<Line> horrizontalLineList = new List<Line>();
            List<Line> slopeLineList = new List<Line>();

            foreach (Face face in faces)
            {
                EdgeArrayArray edgeArrays = face.EdgeLoops;
                EdgeArray edges = edgeArrays.get_Item(0);

                foreach (Edge edge in edges)
                {
                    Line line = edge.AsCurve() as Line;
                    if (LineDirection(line.Direction, XYZ.BasisY) || LineDirection(line.Direction, XYZ.BasisX))
                    {
                        horrizontalLineList.Add(line);
                    }
                    else if (LineDirection(line.Direction, XYZ.BasisZ))
                    {
                        verticalLineList.Add(line);
                    }
                    else
                    {
                        slopeLineList.Add(line);
                    }
                }
            }
            return (horrizontalLineList, verticalLineList, slopeLineList);
        }
        bool LineDirection(XYZ vector, XYZ referenceDirection)
        {
            double dotProduct = vector.DotProduct(referenceDirection);

            // 如果兩個向量的點積為1（或者-1），表示它們平行
            // 這裡可以考慮一個小的閾值，因為在實際應用中可能不會精確為1（或者-1）
            double tolerance = 1e-9;
            return Math.Abs(Math.Abs(dotProduct) - 1) < tolerance;
        }
        private List<Line> AddLine(List<Line> allLines, List<Line> classifyLines)
        {
            foreach (Line line in classifyLines)
            {
                allLines.Add(line);
            }
            return allLines;
        }

        private List<Line> AlignAndSortLines(List<Line> lines, XYZ targetDirection, Func<XYZ, double> sortingCriteria)
        {
            List<Line> newLine = new List<Line>();

            // 1. 修改斜線方向
            foreach (Line line in lines)
            {
                XYZ currentDirection = line.Direction;
                newLine.Add(line);

                // 檢查方向是否與目標方向相反
                if (currentDirection.DotProduct(targetDirection) < 0)
                {
                    newLine.Remove(line);
                    // 反轉方向
                    Line flippedLine = Line.CreateBound(line.GetEndPoint(1), line.GetEndPoint(0));
                    newLine.Add(flippedLine);
                }
            }

            // 2. 排序
            newLine = newLine.OrderBy(line => sortingCriteria(line.GetEndPoint(0))).ToList();
            return newLine;
        }
        private List<Line> RemoveSameLines(List<Line> lines)
        {
            for (int i = 0; i < lines.Count - 1; i++)
            {
                if (lines[i].GetEndPoint(0).IsAlmostEqualTo(lines[i + 1].GetEndPoint(0)) &&
                   lines[i].GetEndPoint(1).IsAlmostEqualTo(lines[i + 1].GetEndPoint(1)))
                {
                    lines.Remove(lines[i]);
                }
            }
            return lines;
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
                        Line newLine = MergeTwoLines(currentLine, nextLine);
                        mergedLines.Remove(currentLine);
                        mergedLines.Remove(nextLine);
                        currentLine = newLine;
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
 
        private XYZ GetOffsetByStairOrientation(XYZ point/*, XYZ orientation*/, double value)
        {
            XYZ newVector = point.Multiply(value);
            XYZ returnPoint = point.Add(newVector);

            return returnPoint;
        }


    }
}

