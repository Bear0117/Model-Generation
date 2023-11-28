using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Aspose.Cells.Charts;
using System.Net;
using Aspose.Cells;
using System.Xml.Linq;
using System.Security.Policy;
using System.Windows.Media;
using Transform = Autodesk.Revit.DB.Transform;
using Document = Autodesk.Revit.DB.Document;
using Aspose.Pdf;
using System.Data.Common;
using System.Windows.Documents;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateStairs_ClickCADLayer : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            ///123123132
            // Default "1F"
            //Level levelBottom = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("4F")) as Level;
            //Level levelTop = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("5F")) as Level;
            //FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(Level));

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
                            Line newLine = Line.CreateBound(Algorithm.RoundPoint(line.GetEndPoint(0), 5), Algorithm.RoundPoint(line.GetEndPoint(1), 5));
                            allrunLines.Add(newLine);
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
            newrunLines = ArrangeLines(allrunLines, newrunLines);

            // 將梯段分為兩邊
            List<Line> runLines_1 = new List<Line>();
            List<Line> runLines_2 = new List<Line>();
            var allrunLines_classify = ClassifyLine(runLines_1, runLines_2, newrunLines);
            runLines_1 = allrunLines_classify.Item1;
            runLines_2 = allrunLines_classify.Item2;


            // 抓平台的圖層
            Reference refer_land = uidoc.Selection.PickObject(ObjectType.PointOnElement, "Select a CAD Layer");
            Element elem_land = doc.GetElement(refer_land);
            GeometryObject geoObj_land = elem_land.GetGeometryObjectFromReference(refer_land);
            Category targetCategory_land = null;
            ElementId graphicsStyleId_land = null;
            var selectCADLayer_land = SelectCADLayer(doc, geoObj_land, targetCategory_land, graphicsStyleId_land);
            targetCategory_land = selectCADLayer_land.Item1;
            graphicsStyleId_land = selectCADLayer_land.Item2;

            // 抓平台裡的所有線
            List<Line> alllandLines = new List<Line>();
            CurveLoop landingLoopPoly = new CurveLoop();
            GeometryElement geoElem_land = elem_land.get_Geometry(new Options());
            foreach (GeometryObject gObj_land in geoElem_land)
            {
                GeometryInstance geomInstance_land = gObj_land as GeometryInstance;

                if (null != geomInstance_land)
                {
                    foreach (GeometryObject insObj_land in geomInstance_land.SymbolGeometry)
                    {
                        if (insObj_land.GraphicsStyleId.IntegerValue != graphicsStyleId_land.IntegerValue)
                            continue;

                        if (insObj_land.GetType().ToString() == "Autodesk.Revit.DB.Line")
                        {
                            Line line = insObj_land as Line;
                            if (Math.Abs(line.GetEndPoint(0).Y - line.GetEndPoint(1).Y) < CentimetersToUnits(1))
                            {
                                continue;
                            }
                            Line newLine = Line.CreateBound(Algorithm.RoundPoint(line.GetEndPoint(0), 5), Algorithm.RoundPoint(line.GetEndPoint(1), 5));
                            alllandLines.Add(newLine);
                        }

                        if (insObj_land.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj_land as PolyLine;
                            List<XYZ> points_list = new List<XYZ>(polyLine.GetCoordinates());
                            //CurveLoop prof = new CurveLoop() as CurveLoop;

                            for (int i = 0; i < points_list.Count - 1; i++)
                            {
                                if (points_list[i].DistanceTo(points_list[i + 1]) < CentimetersToUnits(0.1))
                                {
                                    continue;
                                }
                                Line line = Line.CreateBound(Algorithm.RoundPoint(points_list[i], 5), Algorithm.RoundPoint(points_list[i + 1], 5));
                                //line = TransformLine(transform, line);
                                landingLoopPoly.Append(line);

                            }
                        }
                    }
                    //MessageBox.Show("所有選到的線總共有:" + allriserCurves.Count.ToString());
                }
            }

            // 將線調整方線並依序排列
            List<Line> newlandLines = new List<Line>();
            newlandLines = ArrangeLines(alllandLines, newlandLines);

            // 將平台分為兩邊
            List<Line> landLines_1 = new List<Line>();
            List<Line> landLines_2 = new List<Line>();
            var alllandLines_classify = ClassifyLine(landLines_1, landLines_2, newlandLines);
            landLines_1 = alllandLines_classify.Item1;
            landLines_2 = alllandLines_classify.Item2;

            //建置樓梯
            ElementId newStairsId = null;
            using (StairsEditScope newStairsScope = new StairsEditScope(doc, "New Run"))
            {
                newStairsId = newStairsScope.Start(levelBottom.Id, levelTop.Id);

                using (Transaction stairsTrans = new Transaction(doc, "Add Runs to Stairs"))
                {
                    stairsTrans.Start();

                    Element stair = doc.GetElement(newStairsId) as Element;
                    stair.LookupParameter("所需梯級數").Set(newrunLines.Count);
                    stair.LookupParameter("實際級深").Set(Math.Abs(runLines_1[0].GetEndPoint(0).X - runLines_1[1].GetEndPoint(0).X));

                    var createRuns1 = GetRunsParameter(runLines_1);
                    IList<Curve> bdryCurves1 = createRuns1.Item1;
                    IList<Curve> riserCurves1 = createRuns1.Item2;
                    IList<Curve> pathCurves1 = createRuns1.Item3;

                    //CurveLoop landingLoop1 = GetLandingsParameter(landLines_1);

                    var createRuns2 = GetRunsParameter(runLines_2);
                    IList<Curve> bdryCurves2 = createRuns2.Item1;
                    IList<Curve> riserCurves2 = createRuns2.Item2;
                    IList<Curve> pathCurves2 = createRuns2.Item3;

                    //CurveLoop landingLoop3 = GetLandingsParameter(landLines_2);


                    //StairsLanding newLanding1 = StairsLanding.CreateSketchedLanding(doc, newStairsId, landingLoop1, CentimetersToUnits(20));


                    StairsRun newRun1 = StairsRun.CreateSketchedRun(doc, newStairsId, CentimetersToUnits(20), bdryCurves1, riserCurves1, pathCurves1);
                    //newRun1.ActualRunWidth = Math.Abs(runLines_1[0].GetEndPoint(0).Y - runLines_1[0].GetEndPoint(1).Y);
                    //newRun1.LookupParameter("以豎板結束").Set(0);
                    //newRun1.LookupParameter("突沿長度").Set(2);

                    StairsLanding newLanding2 = StairsLanding.CreateSketchedLanding(doc, newStairsId, landingLoopPoly, newRun1.TopElevation);


                    StairsRun newRun2 = StairsRun.CreateSketchedRun(doc, newStairsId, newRun1.TopElevation, bdryCurves2, riserCurves2, pathCurves2);
                    //newRun2.ActualRunWidth = Math.Abs(runLines_2[0].GetEndPoint(0).Y - runLines_2[0].GetEndPoint(1).Y);
                    //newRun2.LookupParameter("從豎板開始").Set(0);
                    //newRun2.LookupParameter("突沿長度").Set(2);


                    //StairsLanding newLanding3 = StairsLanding.CreateSketchedLanding(doc, newStairsId, landingLoop3, newRun2.TopElevation);

                    stairsTrans.Commit();
                }

                FilteredElementCollector staircollector = new FilteredElementCollector(doc);
                ElementCategoryFilter stairFilter = new ElementCategoryFilter(BuiltInCategory.OST_Stairs);
                IList<Element> stairList = staircollector.WherePasses(stairFilter).WhereElementIsNotElementType().ToElements();

                // A failure preprocessor is to handle possible failures during the edit mode commitment process.
                newStairsScope.Commit(new StairsFailurePreprocessor());
            }


            //篩選所有的樓梯扶手
            FilteredElementCollector railingcollector = new FilteredElementCollector(doc);
            ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRailing);
            IList<Element> raillist = railingcollector.WherePasses(filter).WhereElementIsNotElementType().ToElements();

            //FilteredElementCollector stringcollector = new FilteredElementCollector(doc);
            //ElementCategoryFilter stringfilter = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);
            //IList<Element> stringlist = railingcollector.WherePasses(filter).WhereElementIsNotElementType().ToElements();

            //DeleteStairRailings(doc);
            // 開始一個transaction，每個改變模型的動作都需在transaction中進行
            Transaction railingtrans = new Transaction(doc);
            railingtrans.Start("刪除元件");

            //MessageBox.Show(raillist.Count.ToString());

            // 刪除選取的元件
            foreach (Element elem in raillist)
            {
                ElementId elemId = elem.Id;
                DeleteElement(doc, elemId);
            }
            //foreach (Element elem in stringlist)
            //{
            //    ElementId elemId = elem.Id;
            //    DeleteElement(doc, elemId);
            //}
            railingtrans.Commit();

            return Result.Succeeded;
        }
        // The end of the main code.


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

        class StairsFailurePreprocessor : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor failuresAccessor)
            {
                // Use default failure processing
                return FailureProcessingResult.Continue;
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

        public List<Line> ArrangeLines(List<Line> allLines, List<Line> newLines)
        {
            // 將線的起點方向改為同一邊
            // Define the target direction as Y axis positive direction
            //if (樓梯方向是朝X放向的)
            //{
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
                double diffY = Math.Abs(line1.GetEndPoint(0).Y - line2.GetEndPoint(0).Y);
                double diffX = Math.Abs(line1.GetEndPoint(0).X - line2.GetEndPoint(0).X);

                if (diffY <= epsilon)
                {
                    if (diffX <= epsilon) return 0;  // 在誤差範圍內，視為相等
                    return line1.GetEndPoint(0).X.CompareTo(line2.GetEndPoint(0).X);
                }

                return line1.GetEndPoint(0).Y.CompareTo(line2.GetEndPoint(0).Y);
            }
            );
            //}
            //else
            //{
            //    XYZ targetDirection = XYZ.BasisX;

            //    foreach (Line line1 in allLines)
            //    {
            //        XYZ startPoint = Algorithm.RoundPoint(line1.GetEndPoint(0), 5);
            //        XYZ endPoint = Algorithm.RoundPoint(line1.GetEndPoint(1), 5);

            //        // Calculate the vector from start point to end point
            //        XYZ vector = endPoint - startPoint;

            //        // Check if the start point needs adjustment to align with the target direction
            //        double dotProduct_stair = vector.DotProduct(targetDirection);
            //        if (dotProduct_stair < 0) // Adjust the start point
            //        {
            //            //allCurves.Remove(line1);
            //            XYZ newStartPoint_stair = endPoint;
            //            XYZ newEndPoint_stair = startPoint;
            //            Line newLine_stair = Line.CreateBound(newStartPoint_stair, newEndPoint_stair);
            //            newLines.Add(newLine_stair);
            //        }
            //        else
            //        {
            //            newLines.Add(line1);
            //        }
            //    }
            //    //MessageBox.Show("變更完的線總共有:" + newstairCurves.Count.ToString());

            //    // 將線按照順序排列
            //    double epsilon = 0.1; // 定義你的誤差範圍

            //    newLines.Sort((line1, line2) =>
            //    {
            //        var diffY = Math.Abs(line1.GetEndPoint(0).Y - line2.GetEndPoint(0).Y);
            //        var diffX = Math.Abs(line1.GetEndPoint(0).X - line2.GetEndPoint(0).X);

            //        if (diffX <= epsilon)
            //        {
            //            if (diffY <= epsilon) return 0;  // 在誤差範圍內，視為相等
            //            return line1.GetEndPoint(0).Y.CompareTo(line2.GetEndPoint(0).Y);
            //        }

            //        return line1.GetEndPoint(0).X.CompareTo(line2.GetEndPoint(0).X);
            //    }
            //    );
            //}
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

        public (IList<Curve>, IList<Curve>, IList<Curve>) GetRunsParameter(List<Line> runLines)
        {
            // 梯段第一條和最後條線的點
            XYZ riserfirstLineStartPoint = null;
            XYZ riserfirstLineEndPoint = null;
            XYZ riserlastLineStartPoint = null;
            XYZ riserlastLineEndPoint = null;

            //樓梯邊界曲線
            IList<Curve> bdryCurves = new List<Curve>();
            //樓梯踏板曲線
            IList<Curve> riserCurves = new List<Curve>();
            //樓梯路徑曲線
            IList<Curve> pathCurves = new List<Curve>();

            //Run
            riserfirstLineStartPoint = runLines[0].GetEndPoint(0);
            riserfirstLineEndPoint = runLines[0].GetEndPoint(1);
            riserlastLineStartPoint = runLines[runLines.Count - 1].GetEndPoint(0);
            riserlastLineEndPoint = runLines[runLines.Count - 1].GetEndPoint(1);

            // boundaries
            bdryCurves.Add(Line.CreateBound(riserfirstLineStartPoint, riserlastLineStartPoint));
            bdryCurves.Add(Line.CreateBound(riserfirstLineEndPoint, riserlastLineEndPoint));

            // risers
            foreach (Curve curve in runLines)
            {
                riserCurves.Add(curve);
            }

            // path
            XYZ pathEnd_0 = (riserfirstLineStartPoint + riserfirstLineEndPoint) / 2.0;
            XYZ pathEnd_1 = (riserlastLineStartPoint + riserlastLineEndPoint) / 2.0;
            pathCurves.Add(Line.CreateBound(pathEnd_1, pathEnd_0));

            return (bdryCurves, riserCurves, pathCurves);
        }

        public CurveLoop GetLandingsParameter(List<Line> landLines)
        {
            // 平台邊界線的點
            XYZ landingfirstLineStartPoint = null;
            XYZ landingfirstLineEndPoint = null;
            XYZ landinglastLineStartPoint = null;
            XYZ landinglastLineEndPoint = null;

            // 樓梯平台閉合曲線
            CurveLoop landingLoop = new CurveLoop();

            //第一個平台
            landingfirstLineStartPoint = landLines[0].GetEndPoint(0);
            landingfirstLineEndPoint = landLines[0].GetEndPoint(1);
            landinglastLineStartPoint = landLines[1].GetEndPoint(0);
            landinglastLineEndPoint = landLines[1].GetEndPoint(1);

            Line curve1 = Line.CreateBound(landingfirstLineStartPoint, landingfirstLineEndPoint);
            Line curve2 = Line.CreateBound(landingfirstLineEndPoint, landinglastLineEndPoint);
            Line curve3 = Line.CreateBound(landinglastLineEndPoint, landinglastLineStartPoint);
            Line curve4 = Line.CreateBound(landinglastLineStartPoint, landingfirstLineStartPoint);

            landingLoop.Append(curve1);
            landingLoop.Append(curve2);
            landingLoop.Append(curve3);
            landingLoop.Append(curve4);

            return landingLoop;
        }
    }
    public class StairPlacement
    {
        private readonly Document _doc;
        private readonly UIDocument _uiDoc;

        public StairPlacement(UIDocument uiDoc)
        {
            _uiDoc = uiDoc;
            _doc = _uiDoc.Document;
        }

        public Level GetBottomFloor()
        {
            // 獲取用戶選擇的元件（在這裡假設是匯入的CAD圖元件）
            Reference refer = _uiDoc.Selection.PickObject(ObjectType.Element);
            Element selectedElement = _doc.GetElement(refer);
            GeometryObject geoObj = selectedElement.GetGeometryObjectFromReference(refer);
            Category targetCategory = null;
            ElementId graphicsStyleId = null;

            // 判斷梯段或平台圖層*未完成*
            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                graphicsStyleId = geoObj.GraphicsStyleId;
                //if (_doc.GetElement(geoObj.GraphicsStyleId) is GraphicsStyle gs)
                //{
                //    targetCategory = gs.GraphicsStyleCategory;
                //    // Get the name of the CAD layer which is selected (Column).
                //    String name = gs.GraphicsStyleCategory.Name;
                //}
            }

            Level level = _doc.GetElement(selectedElement.LevelId) as Level;

            // 獲取選擇的元件所在的樓層
            Level bottomFloor = _doc.GetElement(selectedElement.LevelId) as Level;

            return bottomFloor;
        }

        public Level GetTopFloor(Level bottomFloor)
        {
            // 獲取所有樓層
            FilteredElementCollector collector = new FilteredElementCollector(_doc)
                .OfClass(typeof(Level));

            // 找到位於 bottomFloor 上方的最近的樓層
            Level topFloor = null;
            double maxElevation = double.MaxValue;

            foreach (Level level in collector)
            {
                double elevation = level.Elevation;

                // 排除掉與 bottomFloor 相同或在它以下的樓層
                if (elevation > bottomFloor.Elevation && elevation < maxElevation)
                {
                    topFloor = level;
                    maxElevation = elevation;
                }
            }

            return topFloor;
        }
    }

    public class StairParameter
    {
        public int ActualRisersNumber { get; set; }
        public int DesiredRisersNumber { get; set; }
        public double ActualRisersHight { get; set; }
        public double ActualRunWidth { get; set; }
        public int ActualTreadsNumber { get; set; }
        public double ActualTreadsDepth { get; set; }
        public double BaseElevation { get; set; }
        public double TopElevation { get; set; }
        public double Height { get; set; }

        public StairParameter()
        {
            ActualRisersNumber = 0;
            DesiredRisersNumber = 0;
            ActualRisersHight = 0;
            ActualRunWidth = 0;
            ActualTreadsNumber = 0;
            ActualTreadsDepth = 0;
            BaseElevation = 0;
            TopElevation = 0;
            Height = 0;
        }
    }

    public class LandingParameter
    {
        public Double Thickness { get; set; }
        public LandingParameter()
        {
            Thickness = Algorithm.CentimetersToUnits(15);
        }
    }
}



