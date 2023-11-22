using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI.Selection;
using System.Windows;
using System.Windows.Documents;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateStairs_Test2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            StairPlacement stairPlacement = new StairPlacement(uidoc);
            Level levelBottom = stairPlacement.GetBottomFloor();
            Level levelTop = stairPlacement.GetTopFloor(levelBottom);

            // 抓梯段的圖層
            Reference refer_run = uidoc.Selection.PickObject(ObjectType.PointOnElement, "Select a CAD Layer");
            Element elem_run = doc.GetElement(refer_run);
            GeometryObject geoObj_run = elem_run.GetGeometryObjectFromReference(refer_run);
            Category targetCategory_run = null;
            ElementId graphicsStyleId_run = null;
            var selectCADLayer_run = SelectCADLayer(doc, geoObj_run, targetCategory_run, graphicsStyleId_run);
            targetCategory_run = selectCADLayer_run.Item1;
            graphicsStyleId_run = selectCADLayer_run.Item2;

            // 抓梯段裡的所有線
            List<Line> allrunLines = new List<Line>();
            GeometryElement geoElem_run = elem_run.get_Geometry(new Options());
            foreach (GeometryObject gObj_run in geoElem_run)
            {
                GeometryInstance geomInstance_run = gObj_run as GeometryInstance;

                if (null != geomInstance_run)
                {
                    foreach (GeometryObject insObj_run in geomInstance_run.SymbolGeometry)
                    {
                        if (insObj_run.GraphicsStyleId.IntegerValue != graphicsStyleId_run.IntegerValue)
                            continue;

                        if (insObj_run.GetType().ToString() == "Autodesk.Revit.DB.Line")
                        {
                            Line line = insObj_run as Line;
                            if (Math.Abs(line.GetEndPoint(0).Y - line.GetEndPoint(1).Y) < CentimetersToUnits(1))
                            {
                                continue;
                            }
                            allrunLines.Add(line);
                        }

                        if (insObj_run.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj_run as PolyLine;
                            IList<XYZ> points = polyLine.GetCoordinates();
                            Line line_1 = Line.CreateBound(Algorithm.RoundPoint(points[0], 5), Algorithm.RoundPoint(points[1], 5));
                            Line line_2 = Line.CreateBound(Algorithm.RoundPoint(points[1], 5), Algorithm.RoundPoint(points[2], 5));
                            allrunLines.Add(line_1);
                            allrunLines.Add(line_2);
                        }
                    }
                    //MessageBox.Show("所有選到的線總共有:" + allriserCurves.Count.ToString());
                }
            }

            // 將線調整方線並依序排列
            List<Line> newrunLines = new List<Line>();
            newrunLines = TargetDirection(allrunLines, newrunLines);

            // 將梯段分為兩邊
            List<Line> runLines_1 = new List<Line>();
            List<Line> runLines_2 = new List<Line>();
            var allrunLines_classify = ClassifyLine(runLines_1, runLines_2, newrunLines);
            runLines_1 = allrunLines_classify.Item1;
            runLines_2 = allrunLines_classify.Item2;

            //建置樓梯
            ElementId newStairsId = null;
            using (StairsEditScope newStairsScope = new StairsEditScope(doc, "New Run"))
            {
                newStairsId = newStairsScope.Start(levelBottom.Id, levelTop.Id);

                using (Transaction stairsTrans = new Transaction(doc, "Add Runs to Stairs"))
                {
                    stairsTrans.Start();

                    var createRuns1 = GetRunsParameter(runLines_1);
                    XYZ path_Start1 = createRuns1.Item4;
                    XYZ path_End1 = createRuns1.Item5;

                    var createRuns2 = GetRunsParameter(runLines_2);
                    XYZ path_Start2 = createRuns2.Item4;
                    XYZ path_End2 = createRuns2.Item5;

                    //Add a straight run
                    Line locationLine1 = Line.CreateBound(new XYZ(path_Start1.X, path_Start1.Y, levelBottom.Elevation), new XYZ(path_End1.X, path_End1.Y, levelBottom.Elevation));
                    StairsRun newRun1 = StairsRun.CreateStraightRun(doc, newStairsId, locationLine1, StairsRunJustification.Center);

                    newRun1.ActualRunWidth = createRuns1.Item6;
                    //newRun1.TopElevation = levelBottom.Elevation + createRuns1.Item7 * CentimetersToUnits(15);

                    //Line locationLine2 = Line.CreateBound(new XYZ(path_Start2.X, path_Start2.Y, newRun1.TopElevation), new XYZ(path_End2.X, path_End2.Y, newRun1.TopElevation));
                    //StairsRun newRun2 = StairsRun.CreateStraightRun(doc, newStairsId, locationLine2, StairsRunJustification.Center);

                    //newRun2.ActualRunWidth = createRuns2.Item6;
                    //newRun2.TopElevation = levelTop.Elevation;

                    //IList<ElementId> newLanding = StairsLanding.CreateAutomaticLanding(doc, newRun1.Id, newRun2.Id);
                    stairsTrans.Commit();
                }
                // A failure preprocessor is to handle possible failures during the edit mode commitment process.
                newStairsScope.Commit(new StairsFailurePreprocessor());
            }
            return Result.Succeeded;
        }

        class StairsFailurePreprocessor : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                // Use default failure processing
                return FailureProcessingResult.Continue;
            }
        }
        private void DeleteElement(Autodesk.Revit.DB.Document document, ElementId elemId)
        {
            // 將指定元件以及所有與該元件相關聯的元件刪除，並將刪除後所有的元件存到到容器中
            ICollection<Autodesk.Revit.DB.ElementId> deletedIdSet = document.Delete(elemId);

            // 可利用上述容器來查看刪除的數量，若數量為0，則刪除失敗，提供錯誤訊息
            if (deletedIdSet.Count == 0)
            {
                throw new Exception("選取的元件刪除失敗");
            }
        }

        public XYZ ConvertFeetToCentimeters(XYZ xyz)
        {
            // 将 XYZ 坐标从英尺转换为公分
            double factor = UnitUtils.ConvertFromInternalUnits(1.0, UnitTypeId.Centimeters);
            return new XYZ(xyz.X * factor, xyz.Y * factor, xyz.Z * factor);
        }

        public double CentimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Centimeters);
        }

        public class FloorCollector
        {
            private readonly Document _doc;

            public FloorCollector(Document doc)
            {
                _doc = doc;
            }

            public List<Level> GetFloors()
            {
                // 創建一個 FilteredElementCollector 對象，以檢索 Level 元件
                FilteredElementCollector collector = new FilteredElementCollector(_doc)
                    .OfClass(typeof(Level));

                // 使用 LINQ 查詢來篩選出 Floor 樓層元件
                List<Level> floors = collector.Cast<Level>()
                    .Where(level => level.Elevation != 0) // 排除 Ground Floor 或其他特殊樓層
                    .OrderBy(level => level.Elevation) // 根據樓層高度排序
                    .ToList();

                return floors;
            }
        }

        public (Category, ElementId) SelectCADLayer(Document doc, GeometryObject geoObj, Category targetCategory, ElementId graphicsStyleId)
        {
            //選取CAD圖
            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                graphicsStyleId = geoObj.GraphicsStyleId;
                if (doc.GetElement(graphicsStyleId) is GraphicsStyle gs)
                {
                    targetCategory = gs.GraphicsStyleCategory;

                    // Get the name of the CAD layer which is selected (Stair).
                    String name = gs.GraphicsStyleCategory.Name;
                    //MessageBox.Show(name.ToString());
                }
            }
            return (targetCategory, graphicsStyleId);
        }

        public List<Line> TargetDirection(List<Line> allLines, List<Line> newLines)
        {
            // Define the target direction as Y axis positive direction
            XYZ targetDirection = XYZ.BasisY;

            foreach (Line line1 in allLines)
            {
                XYZ startPoint = Algorithm.RoundPoint(line1.GetEndPoint(0), 5);
                XYZ endPoint = Algorithm.RoundPoint(line1.GetEndPoint(1), 5);

                // Calculate the vector from start point to end point
                XYZ vector = endPoint - startPoint;

                // Check if the start point needs adjustment to align with the target direction
                double dotProduct_stair = vector.DotProduct(targetDirection);
                if (dotProduct_stair < 0) // Adjust the start point
                {
                    //allCurves.Remove(line1);
                    XYZ newStartPoint_stair = endPoint;
                    XYZ newEndPoint_stair = startPoint;
                    Line newLine_stair = Line.CreateBound(newStartPoint_stair, newEndPoint_stair);
                    newLines.Add(newLine_stair);
                }
                else
                {
                    newLines.Add(line1);
                }
            }
            //MessageBox.Show("變更完的線總共有:" + newstairCurves.Count.ToString());

            // 將線按照順序排列
            double epsilon = 0.1; // 定義你的誤差範圍

            newLines.Sort((line1, line2) =>
            {
                var diffY = Math.Abs(line1.GetEndPoint(0).Y - line2.GetEndPoint(0).Y);
                var diffX = Math.Abs(line1.GetEndPoint(0).X - line2.GetEndPoint(0).X);

                if (diffY <= epsilon)
                {
                    if (diffX <= epsilon) return 0;  // 在誤差範圍內，視為相等
                    return line1.GetEndPoint(0).X.CompareTo(line2.GetEndPoint(0).X);
                }

                return line1.GetEndPoint(0).Y.CompareTo(line2.GetEndPoint(0).Y);
            });

            return newLines;
        }

        public (List<Line>, List<Line>) ClassifyLine(List<Line> lines_1, List<Line> lines_2, List<Line> newLines)
        {
            for (int i = 0; i < newLines.Count(); i++)
            {
                if (Math.Abs(newLines[0].GetEndPoint(0).Y - newLines[i].GetEndPoint(0).Y) < 0.001)
                {
                    lines_1.Add(newLines[i]);
                }
                else
                {
                    lines_2.Add(newLines[i]);
                }
            }
            lines_2.Reverse();
            return (lines_1, lines_2);
        }

        public (IList<Curve>, IList<Curve>, IList<Curve>, XYZ, XYZ, Double, int) GetRunsParameter(List<Line> runLines)
        {
            // 梯段第一條和最後條線的點
            XYZ firstriserLineStartPoint = null;
            XYZ firstriserLineEndPoint = null;
            XYZ lastriserLineStartPoint = null;
            XYZ lastriserLineEndPoint = null;

            //樓梯邊界曲線
            IList<Curve> bdryCurves = new List<Curve>();
            //樓梯踏板曲線
            IList<Curve> riserCurves = new List<Curve>();
            //樓梯路徑曲線
            IList<Curve> pathCurves = new List<Curve>();

            //run
            firstriserLineStartPoint = runLines[0].GetEndPoint(0);
            firstriserLineEndPoint = runLines[0].GetEndPoint(1);
            lastriserLineStartPoint = runLines[runLines.Count - 1].GetEndPoint(0);
            lastriserLineEndPoint = runLines[runLines.Count - 1].GetEndPoint(1);

            // runwidth
            Double actualRunWidth = Math.Abs(firstriserLineStartPoint.Y - firstriserLineEndPoint.Y);

            // boundaries
            bdryCurves.Add(Line.CreateBound(firstriserLineStartPoint, lastriserLineStartPoint));
            bdryCurves.Add(Line.CreateBound(firstriserLineEndPoint, lastriserLineEndPoint));

            // risers
            int riserCount = 0;
            foreach (Curve curve in runLines)
            {
                riserCurves.Add(curve);
                riserCount++;
            }

            // path
            XYZ pathEnd_0 = (firstriserLineStartPoint + firstriserLineEndPoint) / 2.0;
            XYZ pathEnd_1 = (lastriserLineStartPoint + lastriserLineEndPoint) / 2.0;
            pathCurves.Add(Line.CreateBound(pathEnd_1, pathEnd_0));

            //level
            Level runlevel;

            return (bdryCurves, riserCurves, pathCurves, pathEnd_0, pathEnd_1, actualRunWidth, riserCount);
        }

        public CurveLoop GetLandingsParameter(List<Line> landLines)
        {
            // 平台邊界線的點
            XYZ firstlandingLineStartPoint = null;
            XYZ firstlandingLineEndPoint = null;
            XYZ lastlandingLineStartPoint = null;
            XYZ lastlandingLineEndPoint = null;

            // 樓梯平台閉合曲線
            CurveLoop landingLoop = new CurveLoop();

            //第一個平台
            firstlandingLineStartPoint = landLines[0].GetEndPoint(0);
            firstlandingLineEndPoint = landLines[0].GetEndPoint(1);
            lastlandingLineStartPoint = landLines[1].GetEndPoint(0);
            lastlandingLineEndPoint = landLines[1].GetEndPoint(1);

            Line curve_1 = Line.CreateBound(firstlandingLineStartPoint, firstlandingLineEndPoint);
            Line curve_2 = Line.CreateBound(firstlandingLineEndPoint, lastlandingLineEndPoint);
            Line curve_3 = Line.CreateBound(lastlandingLineEndPoint, lastlandingLineStartPoint);
            Line curve_4 = Line.CreateBound(lastlandingLineStartPoint, firstlandingLineStartPoint);

            landingLoop.Append(curve_1);
            landingLoop.Append(curve_2);
            landingLoop.Append(curve_3);
            landingLoop.Append(curve_4);

            //level
            Level landinglevel;

            return landingLoop;
        }
        public class StairParameter
        {
            public int ActualRisersNumber { get; set; }
            public int DesiredRisersNumber { get; set; }
            public double ActualRisersHight { get; set; }
            public bool EndsWithRiser { get; set; }
            public double ActualRunWidth { get; set; }
            public int ActualTreadsNumber { get; set; }
            public double ActualTreadsDepth { get; set; }
            public double BaseElevation { get; set; }
            public double TopElevation { get; set; }
            public double Height { get; set; }
            public StairsRunJustification LocationLineJustification { get; set; }
            public StairsRunStyle StairsRunStyle { get; }

            public StairParameter() { }
        }

    }

}
