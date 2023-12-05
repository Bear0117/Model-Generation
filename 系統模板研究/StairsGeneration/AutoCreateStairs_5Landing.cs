using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.ExternalService;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data.Common;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Policy;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Media;
using System.Windows.Documents;
using System.Windows.Automation.Peers;
using System.Xml.Linq;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Pdf.Operators;
using Aspose.Pdf;
using Teigha.DatabaseServices;
using Teigha.Runtime;
using Teigha.Geometry;
using Line = Autodesk.Revit.DB.Line;
using Curve = Autodesk.Revit.DB.Curve;
using Transform = Autodesk.Revit.DB.Transform;
using Document = Autodesk.Revit.DB.Document;
using static System.Net.Mime.MediaTypeNames;
using System.Collections;
using System.Drawing;
using Transaction = Autodesk.Revit.DB.Transaction;


namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateStairs_5Landing : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

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
            GeometryElement geoElem_run = elem_run.get_Geometry(new Options());
            Category targetCategory_run = null;
            ElementId graphicsStyleId_run = null;
            string riser_layer_name = "STAIR_RISER"; //Default
            if (geoObj_run.GraphicsStyleId != ElementId.InvalidElementId)
            {
                graphicsStyleId_run = geoObj_run.GraphicsStyleId;
                using (GraphicsStyle gs_run = doc.GetElement(geoObj_run.GraphicsStyleId) as GraphicsStyle)
                {
                    if (gs_run != null)
                    {
                        targetCategory_run = gs_run.GraphicsStyleCategory;
                        riser_layer_name = gs_run.GraphicsStyleCategory.Name;
                    }
                }

                //if (doc.GetElement(graphicsStyleId_run) is GraphicsStyle gs)
                //{
                //    targetCategory_run = gs.GraphicsStyleCategory;

                //    // Get the name of the CAD layer which is selected (Stair).
                //    String name = gs.GraphicsStyleCategory.Name;
                //    //MessageBox.Show(name.ToString());
                //}
            }
            if (geoElem_run == null || graphicsStyleId_run == null)
            {
                message = "GeometryElement or ID does not exist！";
                return Result.Failed;
            }

            // 抓梯段裡的所有線
            List<Line> allrunLines = new List<Line>();
            List<Line> verticalLines = new List<Line>();
            List<Line> horizontalLines = new List<Line>();
            foreach (GeometryObject gObj_run in geoElem_run)
            {
                GeometryInstance geomInstance_run = gObj_run as GeometryInstance;

                if (null != geomInstance_run)
                {
                    foreach (GeometryObject insObj_run in geomInstance_run.SymbolGeometry)
                    {
                        if (insObj_run.GraphicsStyleId.IntegerValue != graphicsStyleId_run.IntegerValue)
                        {
                            continue;
                        }

                        if (insObj_run.GetType().ToString() == "Autodesk.Revit.DB.Line")
                        {
                            Line line = insObj_run as Line;
                            line = Line.CreateBound(Algorithm.RoundPoint(line.GetEndPoint(0), 5), Algorithm.RoundPoint(line.GetEndPoint(1), 5));
                            //if (Math.Abs(line.GetEndPoint(0).Y - line.GetEndPoint(1).Y) < CentimetersToUnits(1))
                            //{
                            //    continue;
                            //}
                            //Line newLine = Line.CreateBound(Algorithm.RoundPoint(line.GetEndPoint(0), 5), Algorithm.RoundPoint(line.GetEndPoint(1), 5));
                            //allrunLines.Add(newLine);
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

                    //判斷線是垂直或水平
                    foreach (Line line in allrunLines)
                    {
                        // 判断线的方向
                        if (IsVertical(line)) // 自定義函數，用于检查线是否垂直
                        {
                            verticalLines.Add(line);
                        }
                        else if (IsHorizontal(line)) // 自定義函數，用于检查线是否水平
                        {
                            horizontalLines.Add(line);
                        }
                    }
                    //MessageBox.Show("所有選到的線總共有:" + allriserCurves.Count.ToString());
                }
            }

            //將垂直梯段調整方線並依序排列
            List<Line> newverticalrunLines = new List<Line>();
            newverticalrunLines = ArrangeVerticalLines(verticalLines, newverticalrunLines);

            // 將垂直梯段分為兩邊
            List<Line> verticalrunLines_1 = new List<Line>();
            List<Line> verticalrunLines_2 = new List<Line>();
            (List<Line>, List<Line>) verticalrunLines_classify = ClassifyVerticalLine(verticalrunLines_1, verticalrunLines_2, newverticalrunLines);
            verticalrunLines_1 = verticalrunLines_classify.Item1;
            verticalrunLines_2 = verticalrunLines_classify.Item2;


            //將水平梯段調整方線並依序排列
            List<Line> newhorizontalrunLines = new List<Line>();
            newhorizontalrunLines = ArrangeHorizontalLines(horizontalLines, newhorizontalrunLines);

            // 將水平梯段分為兩邊
            List<Line> horizontalrunLines_1 = new List<Line>();
            List<Line> horizontalrunLines_2 = new List<Line>();
            (List<Line>, List<Line>) horizontalrunLines_classify = ClassifyHorizontalLine(horizontalrunLines_1, horizontalrunLines_2, newhorizontalrunLines);
            horizontalrunLines_1 = horizontalrunLines_classify.Item1;
            horizontalrunLines_2 = horizontalrunLines_classify.Item2;

            StringBuilder stv_1 = new StringBuilder();
            StringBuilder stv_2 = new StringBuilder();
            StringBuilder sth_1 = new StringBuilder();
            StringBuilder sth_2 = new StringBuilder();

            MessageBox.Show(verticalrunLines_1.Count.ToString());
            foreach (Line line in verticalrunLines_1)
            {
                stv_1.Append(ConvertFeetToCentimeters(line.GetEndPoint(0)).ToString() + "\n");
            }
            TaskDialog.Show("XYZ", stv_1.ToString());

            MessageBox.Show(verticalrunLines_2.Count.ToString());
            foreach (Line line in verticalrunLines_2)
            {
                stv_2.Append(ConvertFeetToCentimeters(line.GetEndPoint(0)).ToString() + "\n");
            }
            TaskDialog.Show("XYZ", stv_2.ToString());

            MessageBox.Show(horizontalrunLines_1.Count.ToString());
            foreach (Line line in horizontalrunLines_1)
            {
                sth_1.Append(ConvertFeetToCentimeters(line.GetEndPoint(0)).ToString() + "\n");
            }
            TaskDialog.Show("XYZ", sth_1.ToString());

            MessageBox.Show(horizontalrunLines_2.Count.ToString());
            foreach (Line line in horizontalrunLines_2)
            {
                sth_2.Append(ConvertFeetToCentimeters(line.GetEndPoint(0)).ToString() + "\n");
            }
            TaskDialog.Show("XYZ", sth_2.ToString());

            //抓取路徑圖層
            Reference refer_path = uidoc.Selection.PickObject(ObjectType.PointOnElement, "Select a CAD Layer");
            Element elem_path = doc.GetElement(refer_path);
            GeometryObject geoObj_path = elem_path.GetGeometryObjectFromReference(refer_path);
            GeometryElement geoElem_path = elem_path.get_Geometry(new Options());
            // Category targetCategory_path = null;
            ElementId graphicsStyleId_path = null;
            string path = GetCADPath(elem_path.GetTypeId(), doc);
            if (geoObj_path.GraphicsStyleId != ElementId.InvalidElementId)
            {
                graphicsStyleId_path = geoObj_path.GraphicsStyleId;
                using (GraphicsStyle gs_path = doc.GetElement(geoObj_path.GraphicsStyleId) as GraphicsStyle)
                {
                    if (gs_path != null)
                    {
                        targetCategory_run = gs_path.GraphicsStyleCategory;
                        //path = gs_path.GraphicsStyleCategory.Name;
                    }
                }


            }
            if (geoElem_path == null || graphicsStyleId_path == null)
            {
                message = "GeometryElement or ID does not exist！";
                return Result.Failed;
            }

            //路徑裡的所有線及點
            List<Line> allpathLines = new List<Line>();
            IList<XYZ> allpathpoints = new List<XYZ>();
            foreach (GeometryObject gObj_path in geoElem_path)
            {
                GeometryInstance geomInstance_path = gObj_path as GeometryInstance;

                if (null != geomInstance_path)
                {
                    foreach (GeometryObject insObj_path in geomInstance_path.SymbolGeometry)
                    {
                        if (insObj_path.GraphicsStyleId.IntegerValue != graphicsStyleId_path.IntegerValue)
                        {
                            continue;
                        }

                        if (insObj_path.GetType().ToString() == "Autodesk.Revit.DB.Line")
                        {
                            Line line = insObj_path as Line;

                            if (Math.Abs(line.GetEndPoint(0).Y - line.GetEndPoint(1).Y) < CentimetersToUnits(1))
                            {
                                continue;
                            }
                            Line newLine = Line.CreateBound(Algorithm.RoundPoint(line.GetEndPoint(0), 5), Algorithm.RoundPoint(line.GetEndPoint(1), 5));
                            allpathLines.Add(newLine);
                        }
                        if (insObj_path.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj_path as PolyLine;
                            allpathpoints = polyLine.GetCoordinates();

                            //MessageBox.Show(allpathLines.Count.ToString());
                            for (int i = 0; i < allpathpoints.Count - 1; i++)
                            {
                                Line line = Line.CreateBound(Algorithm.RoundPoint(allpathpoints[i], 5), Algorithm.RoundPoint(allpathpoints[i + 1], 5));
                                allpathLines.Add(line);
                            }
                        }
                    }
                    //MessageBox.Show("點總共有:" + allpathpoints.Count.ToString());
                    //MessageBox.Show("所有選到的線總共有:" + allpathLines.Count.ToString());
                }
            }

            //StringBuilder st_pathpoints = new StringBuilder();
            //foreach (Line line in allpathLines)
            //{
            //    //pathPoints.Add(line.GetEndPoint(0));
            //    //pathPoints.Add(line.GetEndPoint(1));
            //    st_pathpoints.Append(ConvertFeetToCentimeters(line.GetEndPoint(0)).ToString() + "\n");
            //    st_pathpoints.Append(ConvertFeetToCentimeters(line.GetEndPoint(1)).ToString() + "\n");
            //}
            //MessageBox.Show(st_pathpoints.ToString());

            // 判斷線的方向
            List<Line> vhpathLines = new List<Line>();
            List<XYZ> vhpathpoints = new List<XYZ>();
            StringBuilder st_vhpathpoints = new StringBuilder();
            foreach (Line line in allpathLines)
            {
                if (IsVertical(line) || IsHorizontal(line))
                {
                    vhpathLines.Add(line);
                }
            }
            //MessageBox.Show(vhpathLines.Count.ToString());

            foreach (Line line in vhpathLines)

            {
                vhpathpoints.Add(line.GetEndPoint(0));
                vhpathpoints.Add(line.GetEndPoint(1));
                st_vhpathpoints.Append(ConvertFeetToCentimeters(line.GetEndPoint(0)).ToString() + "\n");
                st_vhpathpoints.Append(ConvertFeetToCentimeters(line.GetEndPoint(1)).ToString() + "\n");
            }
            //MessageBox.Show(vhpathpoints.Count.ToString());
            //MessageBox.Show(st_vhpathpoints.ToString());

            //判斷線離UP的距離
            XYZ up = GetNearestLine(vhpathpoints, path);
            //MessageBox.Show(up.ToString());

            List<List<Line>> finalrun = new List<List<Line>>();
            List<List<Line>> runlines_list = new List<List<Line>>
            {
                verticalrunLines_1,
                verticalrunLines_2,
                horizontalrunLines_1,
                horizontalrunLines_2
            };

            int count = 0;

            while (runlines_list.Count() > 0)
            {
                count++;

                // 初始化最近的線段和清單
                Line closestLine_run = null;
                List<Line> closestList_run = null;
                double minDistance_run = double.MaxValue;

                // 定義一個幫助函數，用於計算線段的中點距離 point 的距離
                double CalculateDistanceToXYZ(Line line)
                {
                    XYZ midPoint = (line.GetEndPoint(0) + line.GetEndPoint(1)) / 2;
                    return up.DistanceTo(midPoint);
                }

                // 遍歷每個清單，找到最近的線段和清單
                foreach (List<Line> line_list in runlines_list)
                {
                    if (line_list != null)
                    {
                        foreach (Line line in line_list)
                        {
                            double distance = CalculateDistanceToXYZ(line);
                            if (distance < minDistance_run)
                            {
                                minDistance_run = distance;
                                closestList_run = line_list;
                                closestLine_run = line;
                            }
                        }
                    }
                }

                if (closestList_run != null && up != null)
                {
                    closestList_run = SortLines(closestList_run, up);
                    //MessageBox.Show(ConvertFeetToCentimeters(closestList_run[0].GetEndPoint(0)).ToString());
                    finalrun.Add(closestList_run);
                    up = closestList_run[closestList_run.Count - 1].GetEndPoint(1);
                    runlines_list.Remove(closestList_run);
                }

                if (count > 100)
                {
                    break;
                }
            }


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
            List<CurveLoop> curvelooplist_Landing = new List<CurveLoop>();

            GeometryElement geoElem_land = elem_land.get_Geometry(new Options());

            int polyLine_count = 0;
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
                            polyLine_count++;
                            PolyLine polyLine = insObj_land as PolyLine;
                            List<XYZ> points_list = new List<XYZ>(polyLine.GetCoordinates());
                            CurveLoop landingLoopPoly = new CurveLoop();
                            //MessageBox.Show(points_list.Count.ToString());

                            for (int i = 0; i < points_list.Count - 1; i++)
                            {
                                if (points_list[i].DistanceTo(points_list[i + 1]) < CentimetersToUnits(0.1))
                                {
                                    continue;
                                }

                                Line line = Line.CreateBound(Algorithm.RoundPoint(points_list[i], 5), Algorithm.RoundPoint(points_list[i + 1], 5));
                                landingLoopPoly.Append(line);
                            }

                            //StringBuilder sb = new StringBuilder();
                            //foreach (Line line in landingLoopPoly)
                            //{
                            //    sb.Append(line.GetEndPoint(0).ToString() + ", " + line.GetEndPoint(0).ToString() + "\n");
                            //}
                            //MessageBox.Show(sb.ToString());


                            curvelooplist_Landing.Add(landingLoopPoly);
                        }
                    }
                    //MessageBox.Show("所有選到的線總共有:" + allriserCurves.Count.ToString());
                }
            }

            //MessageBox.Show(curvelooplist_Landing.Count.ToString());

            //平台排列
            List<CurveLoop> finallanding = new List<CurveLoop>();
            //List<CurveLoop>=curvelooplist_Landing
            List<List<Line>> finalrun_copy = new List<List<Line>>();

            foreach (List<Line> l in finalrun)
            {
                finalrun_copy.Add(l);
            }

            List<CurveLoop> curvelooplist_Landing_copy = new List<CurveLoop>();
            foreach (CurveLoop c in curvelooplist_Landing)
            {
                curvelooplist_Landing_copy.Add(c);
            }

            for (int i = 0; i < curvelooplist_Landing_copy.Count; i++)
            {
                //MessageBox.Show(finalrun_copy.Count.ToString());
                Line line1 = finalrun_copy[i][0];
                Line line2 = finalrun_copy[i][finalrun_copy[i].Count - 1];
                CurveLoop closestCurveLoop = null;
                double minDistance = double.MaxValue;

                //Line1
                foreach (CurveLoop curveLoop in curvelooplist_Landing)
                {
                    foreach (Curve curve in curveLoop)
                    {
                        if (curve is Line line)
                        {
                            // 計算第一個Line和每個CurveLoop中的每個Line之間的距離
                            double distance = CalculateDistanceToXYZ(line, line1);

                            // 如果距離更近，更新最近的CurveLoop和最小距離
                            if (distance < minDistance)
                            {
                                minDistance = distance;
                                closestCurveLoop = curveLoop;

                            }
                        }
                    }
                }

                finallanding.Add(closestCurveLoop);
                curvelooplist_Landing.Remove(closestCurveLoop);

                //MessageBox.Show(curvelooplist_Landing_copy.Count.ToString());
                if (curvelooplist_Landing.Count <= 0)
                {
                    break;
                }

                //Line2
                minDistance = double.MaxValue;

                foreach (CurveLoop curveLoop in curvelooplist_Landing)
                {
                    foreach (Curve curve in curveLoop)
                    {
                        if (curve is Line line)
                        {
                            double distance = CalculateDistanceToXYZ(line, line2);

                            if (distance < minDistance)
                            {
                                minDistance = distance;
                                closestCurveLoop = curveLoop;
                            }
                        }
                    }
                }

                finallanding.Add(closestCurveLoop);
                curvelooplist_Landing.Remove(closestCurveLoop);

                //MessageBox.Show(curvelooplist_Landing_copy.Count.ToString());
                if (curvelooplist_Landing.Count <= 0)
                {
                    break;
                }

            }
            //MessageBox.Show(finallanding.Count.ToString());



            //建置樓梯
            ElementId newStairsId = null;
            using (StairsEditScope newStairsScope = new StairsEditScope(doc, "New Run"))
            {
                // Level levenew = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("123")) as Level;
                newStairsId = newStairsScope.Start(levelBottom.Id, levelTop.Id);

                using (Transaction stairsTrans = new Transaction(doc, "Add Runs to Stairs"))
                {
                    stairsTrans.Start();

                    Element stair = doc.GetElement(newStairsId) as Element;
                    stair.LookupParameter("所需梯級數").Set(allrunLines.Count);
                    stair.LookupParameter("實際級深").Set(Math.Abs(finalrun[0][0].GetEndPoint(0).X - finalrun[0][1].GetEndPoint(0).X));

                    (IList<Curve>, IList<Curve>, IList<Curve>) createRuns1 = GetRunsParameter(finalrun[0]);
                    IList<Curve> bdryCurves1 = createRuns1.Item1;
                    IList<Curve> riserCurves1 = createRuns1.Item2;
                    IList<Curve> pathCurves1 = createRuns1.Item3;

                    //CurveLoop landingLoop1 = GetLandingsParameter(landLines_1);

                    (IList<Curve>, IList<Curve>, IList<Curve>) createRuns2 = GetRunsParameter(finalrun[1]);
                    IList<Curve> bdryCurves2 = createRuns2.Item1;
                    IList<Curve> riserCurves2 = createRuns2.Item2;
                    IList<Curve> pathCurves2 = createRuns2.Item3;

                    //CurveLoop landingLoop3 = GetLandingsParameter(landLines_2);

                    //var createRuns3 = GetRunsParameter(finalrun[2]);
                    //IList<Curve> bdryCurves3 = createRuns3.Item1;
                    //IList<Curve> riserCurves3 = createRuns3.Item2;
                    //IList<Curve> pathCurves3 = createRuns3.Item3;


                    StairsLanding newLanding1 = StairsLanding.CreateSketchedLanding(doc, newStairsId, finallanding[0], Algorithm.CentimetersToUnits(20));

                    StairsRun newRun1 = StairsRun.CreateSketchedRun(doc, newStairsId, Algorithm.CentimetersToUnits(20), bdryCurves1, riserCurves1, pathCurves1);
                    //newRun1.ActualRunWidth = Math.Abs(runLines_1[0].GetEndPoint(0).Y - runLines_1[0].GetEndPoint(1).Y);
                    //newRun1.LookupParameter("以豎板結束").Set(0);
                    //newRun1.LookupParameter("突沿長度").Set(2);

                    StairsLanding newLanding2 = StairsLanding.CreateSketchedLanding(doc, newStairsId, finallanding[1], newRun1.TopElevation);
                    StairsLanding newLanding3 = StairsLanding.CreateSketchedLanding(doc, newStairsId, finallanding[2], newLanding2.BaseElevation);


                    StairsRun newRun2 = StairsRun.CreateSketchedRun(doc, newStairsId, newRun1.TopElevation, bdryCurves2, riserCurves2, pathCurves2);
                    //newRun2.ActualRunWidth = Math.Abs(runLines_2[0].GetEndPoint(0).Y - runLines_2[0].GetEndPoint(1).Y);
                    //newRun2.LookupParameter("從豎板開始").Set(0);
                    //newRun2.LookupParameter("突沿長度").Set(2);

                    StairsLanding newLanding4 = StairsLanding.CreateSketchedLanding(doc, newStairsId, finallanding[3], newRun2.TopElevation);
                    StairsLanding newLanding5 = StairsLanding.CreateSketchedLanding(doc, newStairsId, finallanding[4], newLanding4.BaseElevation);

                    //StairsRun newRun3 = StairsRun.CreateSketchedRun(doc, newStairsId, newRun2.TopElevation, bdryCurves3, riserCurves3, pathCurves3);

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



        double CalculateDistanceToXYZ(Line curveLine, Line riserLine)
        {
            XYZ curveMidPoint = (curveLine.GetEndPoint(0) + curveLine.GetEndPoint(1)) / 2;
            XYZ firstLineMidPoint = (riserLine.GetEndPoint(0) + riserLine.GetEndPoint(1)) / 2;
            return firstLineMidPoint.DistanceTo(curveMidPoint);
        }

        public List<Line> FindClosestListAndLineToXYZ(List<List<Line>> line_list_list, XYZ point)
        {
            // 初始化最近的線段和清單
            Line closestLine = null;
            List<Line> closestList = null;
            double minDistance = double.MaxValue;

            // 定義一個幫助函數，用於計算線段的中點距離 point 的距離
            double CalculateDistanceToXYZ(Line line)
            {
                XYZ midPoint = (line.GetEndPoint(0) + line.GetEndPoint(1)) / 2;
                return point.DistanceTo(midPoint);
            }

            // 遍歷每個清單，找到最近的線段和清單
            foreach (List<Line> line_list in line_list_list)
            {
                if (line_list != null)
                {
                    foreach (Line line in line_list)
                    {
                        double distance = CalculateDistanceToXYZ(line);
                        if (distance < minDistance)
                        {
                            minDistance = distance;
                            closestList = line_list;
                            closestLine = line;
                        }
                    }
                }
            }
            //MessageBox.Show(closestLine.GetEndPoint(0).ToString());
            return closestList;
        }

        public List<Line> SortLines(List<Line> lines, XYZ point)
        {
            List<Line> lines_dup = lines;
            // 複製原始的線段清單，以確保不影響原始資料
            List<Line> newLines = new List<Line>();

            // 建立 KD-Tree 或 Quadtree 以加速查找最近線段的過程
            int count = 0;
            while (lines_dup.Count > 0)
            {
                double minDistance = double.MaxValue;
                Line closestLine = null;

                foreach (Line line in lines_dup)
                {
                    XYZ midPoint = (line.GetEndPoint(0) + line.GetEndPoint(1)) / 2;
                    double distance = point.DistanceTo(midPoint);

                    if (distance < minDistance)
                    {
                        minDistance = distance;
                        closestLine = line;
                    }
                }

                if (closestLine != null)
                {
                    newLines.Add(closestLine);
                    lines_dup.Remove(closestLine);
                }
                count++;
                if (count > 100)
                {
                    break;
                }
            }

            return newLines;
        }



        private void DeleteElement(Autodesk.Revit.DB.Document document, ElementId elemId)
        {
            // 將指定元件以及所有與該元件相關聯的元件刪除，並將刪除後所有的元件存到到容器中
            ICollection<Autodesk.Revit.DB.ElementId> deletedIdSet = document.Delete(elemId);

            // 可利用上述容器來查看刪除的數量，若數量為0，則刪除失敗，提供錯誤訊息
            if (deletedIdSet.Count == 0)
            {
                throw new System.Exception("選取的元件刪除失敗");
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

        private bool IsVertical(Line line)
        {
            double angle = line.Direction.AngleTo(XYZ.BasisY);

            // 如果線的方向與Y軸方向的夾角接近90度或接近270度，則認為是垂直線
            return Math.Abs(angle) < 0.01 || Math.Abs(angle - Math.PI) < 0.01;

        }
        private bool IsHorizontal(Line line)
        {
            double angle = line.Direction.AngleTo(XYZ.BasisX);

            // 如果线的方向与X轴方向的夹角接近0度，则认为是水平线
            return Math.Abs(angle) < 0.01 || Math.Abs(angle - Math.PI) < 0.01;
        }


        public List<Line> ArrangeVerticalLines(List<Line> allLines, List<Line> newLines)
        {
            // 將線的起點方向改為同一邊
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
            return newLines;
        }
        public List<Line> ArrangeHorizontalLines(List<Line> allLines, List<Line> newLines)
        {
            // 將線的起點方向改為同一邊
            // Define the target direction as Y axis positive direction
            XYZ targetDirection = XYZ.BasisX;

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

            //還要修改
            double epsilon = 0.1; // 定義你的誤差範圍

            newLines.Sort((line1, line2) =>
            {
                double diffY = Math.Abs(line1.GetEndPoint(0).Y - line2.GetEndPoint(0).Y);
                double diffX = Math.Abs(line1.GetEndPoint(0).X - line2.GetEndPoint(0).X);

                if (diffX <= epsilon)
                {
                    if (diffY <= epsilon) return 0;  // 在誤差範圍內，視為相等
                    return line1.GetEndPoint(0).Y.CompareTo(line2.GetEndPoint(0).Y);
                }

                return line1.GetEndPoint(0).X.CompareTo(line2.GetEndPoint(0).X);
            }
            );
            return newLines;
        }

        public (List<Line>, List<Line>) ClassifyVerticalLine(List<Line> lines_1, List<Line> lines_2, List<Line> newLines)
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
        public (List<Line>, List<Line>) ClassifyHorizontalLine(List<Line> lines_1, List<Line> lines_2, List<Line> newLines)
        {
            for (int i = 0; i < newLines.Count(); i++)
            {
                if (Math.Abs(newLines[0].GetEndPoint(0).X - newLines[i].GetEndPoint(0).X) < 0.001)
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

        private string GetParameterInformation(Parameter para, Document document)
        {
            string defName = para.Definition.Name;
            switch (para.StorageType)
            {
                case StorageType.Double:
                    return defName + ":" + para.AsValueString();

                case StorageType.ElementId:
                    ElementId id = para.AsElementId();
                    if (id.IntegerValue >= 0)
                        return defName + ":" + document.GetElement(id).Name;
                    //在2014以前取得元件的方法為get_Element()，而在2014方法更新為GetElement()
                    else
                        return defName + ":" + id.IntegerValue.ToString();

                case StorageType.Integer:
                    // if (ParameterType.YesNo == para.Definition.ParameterType) //在2022以前的做法
                    if (SpecTypeId.Boolean.YesNo == para.Definition.GetDataType()) //在2023以後的做法
                    {
                        if (para.AsInteger() == 0)
                            return defName + ":" + "False";
                        else
                            return defName + ":" + "True";
                    }
                    else
                    {
                        return defName + ":" + para.AsInteger().ToString();
                    }

                case StorageType.String:
                    return defName + ":" + para.AsString();

                default:
                    return "未公開的參數";
            }
        }
        public (IList<Curve>, IList<Curve>, IList<Curve>) GetRunsParameter(List<Line> runLines)
        {
            // 梯段第一條和最後條線的點
            XYZ riserfirstLineStartPoint;
            XYZ riserfirstLineEndPoint;
            XYZ riserlastLineStartPoint;
            XYZ riserlastLineEndPoint;

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
            XYZ landingfirstLineStartPoint;
            XYZ landingfirstLineEndPoint;
            XYZ landinglastLineStartPoint;
            XYZ landinglastLineEndPoint;

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

        public XYZ GetNearestLine(List<XYZ> points, String path)
        {
            List<XYZ> ups = new List<XYZ>();
            XYZ up_coor = new XYZ(0, 0, 0);

            using (new Services())
            {
                using (Database database = new Database(false, false))
                {
                    database.ReadDwgFile(path, FileShare.Read, true, "");
                    using (Teigha.DatabaseServices.Transaction trans1 = database.TransactionManager.StartTransaction())
                    {
                        using (BlockTable table = (BlockTable)database.BlockTableId.GetObject(OpenMode.ForRead))
                        {
                            using (SymbolTableEnumerator enumerator = table.GetEnumerator())
                            {
                                enumerator.MoveNext();
                                {
                                    using (BlockTableRecord record = (BlockTableRecord)enumerator.Current.GetObject(OpenMode.ForRead))
                                    {

                                        //Start from here
                                        foreach (ObjectId id in record)
                                        {
                                            bool isUpFound = false;
                                            Entity entity2 = (Entity)id.GetObject(OpenMode.ForRead, false, false);
                                            switch (entity2.GetRXClass().Name)
                                            {
                                                case "AcDbText":
                                                    Teigha.DatabaseServices.DBText text = (Teigha.DatabaseServices.DBText)entity2;
                                                    XYZ DBtext_Position = ConverCADPointToRevitPoint(text.Position);
                                                    if (text.Layer == "STAIR_PATH" && text.TextString.Contains("UP"))
                                                    {
                                                        up_coor = PointMilimeterToUnit(DBtext_Position);
                                                        isUpFound = true;
                                                        ups.Add(DBtext_Position);
                                                        //MessageBox.Show(text.TextString);
                                                        //MessageBox.Show(up_coor.ToString());
                                                    }
                                                    break;


                                                case "AcDbMText":
                                                    Teigha.DatabaseServices.MText text_m = (Teigha.DatabaseServices.MText)entity2;
                                                    XYZ Mtext_Position = ConverCADPointToRevitPoint(text_m.Location);
                                                    if (text_m.Layer == "STAIR_PATH" && text_m.Text.Contains("UP"))
                                                    {
                                                        up_coor = PointMilimeterToUnit(Mtext_Position);
                                                        isUpFound = true;
                                                        ups.Add(Mtext_Position);
                                                        //MessageBox.Show(text_m.Text);
                                                        //MessageBox.Show(up_coor.ToString());
                                                    }
                                                    break;
                                            }
                                            if (isUpFound)
                                            {
                                                break;
                                            }

                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if (points == null || points.Count == 0)
            {
                throw new ArgumentException("The points list is empty or null.");
            }

            double minDistance = double.MaxValue;
            XYZ closestPoint = null;

            foreach (XYZ point in points)
            {
                double distance = point.DistanceTo(up_coor);
                if (distance < minDistance)
                {
                    minDistance = distance;
                    closestPoint = point;
                }
            }

            if (closestPoint != null)
            {
                //MessageBox.Show(closestPoint.ToString());
                return closestPoint;
            }
            else
            {
                throw new InvalidOperationException("No closest point found.");
            }
        }


        public XYZ ConverCADPointToRevitPoint(Point3d point)
        {
            return new XYZ(point.X, point.Y, point.Z);
        }
        public string GetCADPath(ElementId cadLinkTypeID, Document revitDoc)
        {
            CADLinkType cadLinkType = revitDoc.GetElement(cadLinkTypeID) as CADLinkType;
            return ModelPathUtils.ConvertModelPathToUserVisiblePath(cadLinkType.GetExternalFileReference().GetAbsolutePath());
        }
        public XYZ PointMilimeterToUnit(XYZ point)
        {
            XYZ newPoint = new XYZ(
                CentimetersToUnits(point.X),
                CentimetersToUnits(point.Y),
                CentimetersToUnits(point.Z)
                );
            return newPoint;
        }
        public List<Autodesk.Revit.DB.Line> CurveLoopToLineList(CurveLoop curveLoop)
        {
            List<Autodesk.Revit.DB.Line> lineList = new List<Autodesk.Revit.DB.Line>();
            List<Autodesk.Revit.DB.Curve> curveList = curveLoop.ToList();
            foreach (Autodesk.Revit.DB.Curve curve in curveList)
            {
                lineList.Add(Autodesk.Revit.DB.Line.CreateBound(curve.GetEndPoint(0), curve.GetEndPoint(1)));
            }
            return lineList;
        }
        public bool IsInsideOutline(XYZ TargetPoint, CurveLoop curveloop)
        {
            bool result = true;
            int insertCount = 0;
            List<Autodesk.Revit.DB.Line> lines = CurveLoopToLineList(curveloop);
            Autodesk.Revit.DB.Line rayLine = Autodesk.Revit.DB.Line.CreateBound(TargetPoint, TargetPoint.Add(new XYZ(1, 0, 0) * 100000000));

            foreach (Autodesk.Revit.DB.Line areaLine in lines)
            {
                SetComparisonResult interResult = areaLine.Intersect(rayLine, out IntersectionResultArray resultArray);
                IntersectionResult insPoint = resultArray?.get_Item(0);
                if (insPoint != null)
                {
                    insertCount++;
                }
            }

            // To varify the point is inside the outline or not.
            if (insertCount % 2 == 0) //even
            {
                result = false;
                return result;
            }
            else
            {
                return result;
            }
        }

    }
    //public class StairPlacement
    //{
    //    private readonly Document _doc;
    //    private readonly UIDocument _uiDoc;

    //    public StairPlacement(UIDocument uiDoc)
    //    {
    //        _uiDoc = uiDoc;
    //        _doc = _uiDoc.Document;
    //    }

    //    public Level GetBottomFloor()
    //    {
    //        // 獲取用戶選擇的元件（在這裡假設是匯入的CAD圖元件）
    //        Reference refer = _uiDoc.Selection.PickObject(ObjectType.Element);
    //        Element selectedElement = _doc.GetElement(refer);
    //        GeometryObject geoObj = selectedElement.GetGeometryObjectFromReference(refer);
    //        Category targetCategory = null;
    //        ElementId graphicsStyleId = null;

    //        // 判斷梯段或平台圖層*未完成*
    //        if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
    //        {
    //            graphicsStyleId = geoObj.GraphicsStyleId;
    //            //if (_doc.GetElement(geoObj.GraphicsStyleId) is GraphicsStyle gs)
    //            //{
    //            //    targetCategory = gs.GraphicsStyleCategory;
    //            //    // Get the name of the CAD layer which is selected (Column).
    //            //    String name = gs.GraphicsStyleCategory.Name;
    //            //}
    //        }

    //        Level level = _doc.GetElement(selectedElement.LevelId) as Level;

    //        // 獲取選擇的元件所在的樓層
    //        Level bottomFloor = _doc.GetElement(selectedElement.LevelId) as Level;

    //        return bottomFloor;
    //    }

    //    public Level GetTopFloor(Level bottomFloor)
    //    {
    //        // 獲取所有樓層
    //        FilteredElementCollector collector = new FilteredElementCollector(_doc)
    //            .OfClass(typeof(Level));

    //        // 找到位於 bottomFloor 上方的最近的樓層
    //        Level topFloor = null;
    //        double maxElevation = double.MaxValue;

    //        foreach (Level level in collector)
    //        {
    //            double elevation = level.Elevation;

    //            // 排除掉與 bottomFloor 相同或在它以下的樓層
    //            if (elevation > bottomFloor.Elevation && elevation < maxElevation)
    //            {
    //                topFloor = level;
    //                maxElevation = elevation;
    //            }
    //        }

    //        return topFloor;
    //    }
    //}

    //public class StairParameter
    //{
    //    public int ActualRisersNumber { get; set; }
    //    public int DesiredRisersNumber { get; set; }
    //    public double ActualRisersHight { get; set; }
    //    public double ActualRunWidth { get; set; }
    //    public int ActualTreadsNumber { get; set; }
    //    public double ActualTreadsDepth { get; set; }
    //    public double BaseElevation { get; set; }
    //    public double TopElevation { get; set; }
    //    public double Height { get; set; }

    //    public StairParameter()
    //    {
    //        ActualRisersNumber = 0;
    //        DesiredRisersNumber = 0;
    //        ActualRisersHight = 0;
    //        ActualRunWidth = 0;
    //        ActualTreadsNumber = 0;
    //        ActualTreadsDepth = 0;
    //        BaseElevation = 0;
    //        TopElevation = 0;
    //        Height = 0;
    //    }
    //}

    //public class LandingParameter
    //{
    //    public Double Thickness { get; set; }
    //    public LandingParameter()
    //    {
    //        Thickness = Algorithm.CentimetersToUnits(15);
    //    }
    //}
}