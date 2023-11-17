using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
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


namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateStairs_ClickLine : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            //Reference stairlevel = uidoc.Selection.PickObject(ObjectType.Element, "Pick a level.");
            //Element elem_stairlevel = doc.GetElement(stairlevel);
            //Level levelBottom = doc.GetElement(elem_stairlevel.LevelId) as Level;

            // Default "1F"
            Level levelBottom = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("4F")) as Level;
            Level levelTop = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("5F")) as Level;

            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(Level));

            ElementId newStairsId = null;

            using (StairsEditScope newStairsScope = new StairsEditScope(doc, "New Stairs"))
            {
                newStairsId = newStairsScope.Start(levelBottom.Id, levelTop.Id);

                using (Transaction stairsTrans = new Transaction(doc, "Add Runs and Landings to Stairs"))
                {
                    stairsTrans.Start();

                    //第一個樓梯平台建置
                    //樓梯平台閉合曲線
                    CurveLoop landingLoop1 = new CurveLoop();

                    //樓梯平台點
                    IList<XYZ> landingpoints1 = new List<XYZ>();

                    //樓梯梯段第一條被選到的線
                    Line landingfirstline1 = null;

                    //樓梯梯段線數量
                    int landinglinecount1 = 0;

                    // 樓梯平台第一條和最後條線的點
                    XYZ landingfirstLineStartPoint1 = null;
                    XYZ landingfirstLineEndPoint1 = null;
                    XYZ landinglastLineStartPoint1 = null;
                    XYZ landinglastLineEndPoint1 = null;


                    //開始選樓梯平台線
                    while (true)
                    {
                        Reference pickedlandingRef1 = uidoc.Selection.PickObject(ObjectType.Element, "Pick a stair model line.");
                        Element elem_landing1 = doc.GetElement(pickedlandingRef1);
                        //levelBottom = doc.GetElement(elem_landing.LevelId) as Level;
                        LocationCurve landinglocationCurve1 = elem_landing1.Location as LocationCurve;
                        Line landinglocationLine1 = landinglocationCurve1.Curve as Line;

                        if (landingfirstline1 != null && landingfirstline1 == landinglocationLine1)
                        {
                            break;
                        }

                        landingfirstline1 = landinglocationLine1;

                        XYZ landingstartpoint1 = landinglocationLine1.GetEndPoint(0);
                        XYZ landingendpoint1 = landinglocationLine1.GetEndPoint(1);


                        landingpoints1.Add(landingstartpoint1);
                        landingpoints1.Add(landingendpoint1);

                        landinglinecount1++;

                        ////輸出所選到平台線的點座標
                        //TaskDialog.Show(" Landing Line point Information",
                        //                     $"Total selected lines: {landinglinecount} " +
                        //                     $"\nLandingStartPoint={(startpoint)}," +
                        //                     $"\nLandingEndPoint={(endpoint)}");

                    }

                    //Add a landing between the runs
                    landingfirstLineStartPoint1 = landingpoints1[0];
                    landingfirstLineEndPoint1 = landingpoints1[1];
                    landinglastLineStartPoint1 = landingpoints1[landinglinecount1 * 2 - 2];
                    landinglastLineEndPoint1 = landingpoints1[landinglinecount1 * 2 - 1];

                    Line curve_1 = Line.CreateBound(landingfirstLineStartPoint1, landingfirstLineEndPoint1);
                    Line curve_2 = Line.CreateBound(landingfirstLineEndPoint1, landinglastLineStartPoint1);
                    Line curve_3 = Line.CreateBound(landinglastLineStartPoint1, landinglastLineEndPoint1);
                    Line curve_4 = Line.CreateBound(landinglastLineEndPoint1, landingfirstLineStartPoint1);

                    landingLoop1.Append(curve_1);
                    landingLoop1.Append(curve_2);
                    landingLoop1.Append(curve_3);
                    landingLoop1.Append(curve_4);

                    StairsLanding newLanding1 = StairsLanding.CreateSketchedLanding(doc, newStairsId, landingLoop1, levelBottom.Elevation);

                    //第一個樓梯梯段建置
                    //樓梯邊界曲線
                    IList<Curve> bdryCurves1 = new List<Curve>();

                    //樓梯梯段曲線
                    IList<Curve> riserCurves1 = new List<Curve>();

                    //樓梯路徑曲線
                    IList<Curve> pathCurves1 = new List<Curve>();

                    //樓梯梯段點
                    IList<XYZ> riserpoints1 = new List<XYZ>();

                    //樓梯梯段第一條被選到的線
                    Line stairfirstline1 = null;

                    //樓梯梯段線數量
                    int stairlinecount1 = 0;

                    // 樓梯第一條和最後條線的點
                    XYZ stairfirstLineStartPoint1 = null;
                    XYZ stairfirstLineEndPoint1 = null;
                    XYZ stairlastLineStartPoint1 = null;
                    XYZ stairlastLineEndPoint1 = null;

                    //開始選樓梯梯段線
                    while (true)
                    {
                        Reference pickedstairRef1 = uidoc.Selection.PickObject(ObjectType.Element, "Pick a stair model line.");
                        Element elem_stair1 = doc.GetElement(pickedstairRef1);
                        //levelBottom = doc.GetElement(elem_stair.LevelId) as Level;
                        LocationCurve stairlocationCurve1 = elem_stair1.Location as LocationCurve;
                        Line stairlocationLine1 = stairlocationCurve1.Curve as Line;

                        if (stairfirstline1 != null && stairfirstline1 == stairlocationLine1)
                        {
                            break;
                        }

                        stairfirstline1 = stairlocationLine1;

                        XYZ stairstartpoint1 = stairlocationLine1.GetEndPoint(0);
                        XYZ stairendpoint1 = stairlocationLine1.GetEndPoint(1);

                        riserpoints1.Add(stairstartpoint1);
                        riserpoints1.Add(stairendpoint1);
                        riserCurves1.Add(stairlocationLine1);

                        stairlinecount1++;

                        ////輸出樓梯所有點的座標
                        //TaskDialog.Show(" Line point Information",
                        //                  $"Total selected lines: {stairlinecount} " +
                        //                  $"\nStairStartPoint={(startpoint)}," +
                        //                  $"\nStairEndPoint={(endpoint)}");

                    }

                    //樓梯第一條線和最後一條線的點座標
                    stairfirstLineStartPoint1 = riserpoints1[0];
                    stairfirstLineEndPoint1 = riserpoints1[1];
                    stairlastLineStartPoint1 = riserpoints1[stairlinecount1 * 2 - 2];
                    stairlastLineEndPoint1 = riserpoints1[stairlinecount1 * 2 - 1];

                    ////輸出樓梯第一條線和最後一條線的點座標
                    //TaskDialog.Show("Stair Line point Information",
                    //                         $"Total selected lines: {stairlinecount} " +
                    //                         $"\nstairfirstLineStartPoint={(stairfirstLineStartPoint)}," +
                    //                         $"\nstairfirstLineStartPoint={(stairfirstLineStartPoint)}"+
                    //                         $"\nstairlastLineStartPoint={(stairlastLineStartPoint)}," +
                    //                         $"\nstairlastLineEndPoint={(stairlastLineEndPoint)}," );

                    // boundaries
                    if (stairfirstLineStartPoint1.X == stairlastLineStartPoint1.X)
                    {
                        bdryCurves1.Add(Line.CreateBound(stairfirstLineStartPoint1, stairlastLineStartPoint1));
                        bdryCurves1.Add(Line.CreateBound(stairfirstLineEndPoint1, stairlastLineEndPoint1));
                    }
                    else
                    {
                        bdryCurves1.Add(Line.CreateBound(stairfirstLineStartPoint1, stairlastLineEndPoint1));
                        bdryCurves1.Add(Line.CreateBound(stairfirstLineEndPoint1, stairlastLineStartPoint1));
                    }

                    //stairs path curves
                    XYZ pathEnd1_0 = (stairfirstLineStartPoint1 + stairfirstLineEndPoint1) / 2.0;
                    XYZ pathEnd1_1 = (stairlastLineStartPoint1 + stairlastLineEndPoint1) / 2.0;
                    pathCurves1.Add(Line.CreateBound(pathEnd1_0, pathEnd1_1));

                    StairsRun newRun1 = StairsRun.CreateSketchedRun(doc, newStairsId, levelBottom.Elevation, bdryCurves1, riserCurves1, pathCurves1);


                    //第二個樓梯平台建置
                    //樓梯平台閉合曲線
                    CurveLoop landingLoop2 = new CurveLoop();

                    //樓梯平台點
                    IList<XYZ> landingpoints2 = new List<XYZ>();

                    //樓梯梯段第一條被選到的線
                    Line landingfirstline2 = null;

                    //樓梯梯段線數量
                    int landinglinecount2 = 0;

                    // 樓梯第一條和最後條線的點
                    XYZ landingfirstLineStartPoint2 = null;
                    XYZ landingfirstLineEndPoint2 = null;
                    XYZ landinglastLineStartPoint2 = null;
                    XYZ landinglastLineEndPoint2 = null;


                    //開始選樓梯平台線
                    while (true)
                    {
                        Reference pickedlandingRef2 = uidoc.Selection.PickObject(ObjectType.Element, "Pick a stair model line.");
                        Element elem_landing2 = doc.GetElement(pickedlandingRef2);
                        //levelBottom = doc.GetElement(elem_landing.LevelId) as Level;
                        LocationCurve landinglocationCurve2 = elem_landing2.Location as LocationCurve;
                        Line landinglocationLine2 = landinglocationCurve2.Curve as Line;

                        if (landingfirstline2 != null && landingfirstline2 == landinglocationLine2)
                        {
                            break;
                        }

                        landingfirstline2 = landinglocationLine2;

                        XYZ landingstartpoint2 = landinglocationLine2.GetEndPoint(0);
                        XYZ landingendpoint2 = landinglocationLine2.GetEndPoint(1);


                        landingpoints2.Add(landingstartpoint2);
                        landingpoints2.Add(landingendpoint2);

                        landinglinecount2++;

                        ////輸出所選到平台線的點座標
                        //TaskDialog.Show(" Landing Line point Information",
                        //                     $"Total selected lines: {landinglinecount} " +
                        //                     $"\nLandingStartPoint={(startpoint)}," +
                        //                     $"\nLandingEndPoint={(endpoint)}");

                    }

                    //Add a landing between the runs
                    landingfirstLineStartPoint2 = landingpoints2[0];
                    landingfirstLineEndPoint2 = landingpoints2[1];
                    landinglastLineStartPoint2 = landingpoints2[landinglinecount2 * 2 - 2];
                    landinglastLineEndPoint2 = landingpoints2[landinglinecount2 * 2 - 1];

                    //if(landingfirstLineStartPoint2.ToString()=)
                    Line curve2_1 = Line.CreateBound(landingfirstLineStartPoint2, landingfirstLineEndPoint2);
                    Line curve2_2 = Line.CreateBound(landingfirstLineEndPoint2, landinglastLineStartPoint2);
                    Line curve2_3 = Line.CreateBound(landinglastLineStartPoint2, landinglastLineEndPoint2);
                    Line curve2_4 = Line.CreateBound(landinglastLineEndPoint2, landingfirstLineStartPoint2);

                    landingLoop2.Append(curve2_1);
                    landingLoop2.Append(curve2_2);
                    landingLoop2.Append(curve2_3);
                    landingLoop2.Append(curve2_4);

                    StairsLanding newLanding2 = StairsLanding.CreateSketchedLanding(doc, newStairsId, landingLoop2, newRun1.TopElevation);

                    //第二個樓梯梯段建置
                    //樓梯邊界曲線
                    IList<Curve> bdryCurves2 = new List<Curve>();

                    //樓梯梯段曲線
                    IList<Curve> riserCurves2 = new List<Curve>();

                    //樓梯路徑曲線
                    IList<Curve> pathCurves2 = new List<Curve>();

                    //樓梯梯段點
                    IList<XYZ> riserpoints2 = new List<XYZ>();

                    //樓梯梯段第一條被選到的線
                    Line stairfirstline2 = null;

                    //樓梯梯段線數量
                    int stairlinecount2 = 0;

                    // 樓梯第一條和最後條線的點
                    XYZ stairfirstLineStartPoint2 = null;
                    XYZ stairfirstLineEndPoint2 = null;
                    XYZ stairlastLineStartPoint2 = null;
                    XYZ stairlastLineEndPoint2 = null;

                    //開始選第二個樓梯梯段線
                    while (true)
                    {
                        Reference pickedstairRef2 = uidoc.Selection.PickObject(ObjectType.Element, "Pick a stair model line.");
                        Element elem_stair2 = doc.GetElement(pickedstairRef2);
                        //levelBottom = doc.GetElement(elem_stair.LevelId) as Level;
                        LocationCurve stairlocationCurve2 = elem_stair2.Location as LocationCurve;
                        Line stairlocationLine2 = stairlocationCurve2.Curve as Line;

                        if (stairfirstline2 != null && stairfirstline2 == stairlocationLine2)
                        {
                            break;
                        }

                        stairfirstline2 = stairlocationLine2;

                        XYZ startpoint = stairlocationLine2.GetEndPoint(0);
                        XYZ endpoint = stairlocationLine2.GetEndPoint(1);

                        riserpoints2.Add(startpoint);
                        riserpoints2.Add(endpoint);
                        riserCurves2.Add(stairlocationLine2);

                        stairlinecount2++;

                        ////輸出樓梯所有點的座標
                        //TaskDialog.Show(" Line point Information",
                        //                  $"Total selected lines: {stairlinecount} " +
                        //                  $"\nStairStartPoint={(startpoint)}," +
                        //                  $"\nStairEndPoint={(endpoint)}");

                    }

                    //樓梯第一條線和最後一條線的點座標
                    stairfirstLineStartPoint2 = riserpoints2[0];
                    stairfirstLineEndPoint2 = riserpoints2[1];
                    stairlastLineStartPoint2 = riserpoints2[stairlinecount2 * 2 - 2];
                    stairlastLineEndPoint2 = riserpoints2[stairlinecount2 * 2 - 1];

                    ////輸出樓梯第一條線和最後一條線的點座標
                    //TaskDialog.Show("Stair Line point Information",
                    //                         $"Total selected lines: {stairlinecount} " +
                    //                         $"\nstairfirstLineStartPoint={(stairfirstLineStartPoint)}," +
                    //                         $"\nstairfirstLineStartPoint={(stairfirstLineStartPoint)}"+
                    //                         $"\nstairlastLineStartPoint={(stairlastLineStartPoint)}," +
                    //                         $"\nstairlastLineEndPoint={(stairlastLineEndPoint)}," );

                    // boundaries
                    if (stairfirstLineStartPoint2.X == stairlastLineStartPoint2.X)
                    {
                        bdryCurves2.Add(Line.CreateBound(stairfirstLineStartPoint2, stairlastLineStartPoint2));
                        bdryCurves2.Add(Line.CreateBound(stairfirstLineEndPoint2, stairlastLineEndPoint2));
                    }
                    else
                    {
                        bdryCurves2.Add(Line.CreateBound(stairfirstLineStartPoint2, stairlastLineEndPoint2));
                        bdryCurves2.Add(Line.CreateBound(stairfirstLineEndPoint2, stairlastLineStartPoint2));
                    }

                    //stairs path curves
                    XYZ pathEnd2_0 = (stairfirstLineStartPoint2 + stairfirstLineEndPoint2) / 2.0;
                    XYZ pathEnd2_1 = (stairlastLineStartPoint2 + stairlastLineEndPoint2) / 2.0;
                    pathCurves2.Add(Line.CreateBound(pathEnd2_0, pathEnd2_1));

                    StairsRun newRun = StairsRun.CreateSketchedRun(doc, newStairsId, newRun1.TopElevation, bdryCurves2, riserCurves2, pathCurves2);

                    stairsTrans.Commit();
                }

                // A failure preprocessor is to handle possible failures during the edit mode commitment process.
                newStairsScope.Commit(new StairsFailurePreprocessor());
            }

            //篩選所有的樓梯扶手
            FilteredElementCollector railingcollector = new FilteredElementCollector(doc);
            ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRailing);
            IList<Element> raillist = railingcollector.WherePasses(filter).WhereElementIsNotElementType().ToElements();

            // 開始一個transaction，每個改變模型的動作都需在transaction中進行
            Transaction railingtrans = new Transaction(doc);
            railingtrans.Start("刪除元件");

            MessageBox.Show(raillist.Count.ToString());


            // 刪除選取的元件
            foreach (Element elem in raillist)
            {
                ElementId elemId = elem.Id;
                DeleteElement(doc, elemId);
            }

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
        public List<CurveLoop> GetAllOutlines(Document doc, GeometryObject geoObj, GeometryElement geoElem)
        {
            if (doc is null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            if (geoObj is null)
            {
                throw new ArgumentNullException(nameof(geoObj));
            }

            if (geoElem is null)
            {
                throw new ArgumentNullException(nameof(geoElem));
            }

            ElementId graphicsStyleId = null;
            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                graphicsStyleId = geoObj.GraphicsStyleId;
            }

            List<CurveLoop> allOutlines = new List<CurveLoop>();
            foreach (GeometryObject gObj in geoElem)
            {
                GeometryInstance geomInstance = gObj as GeometryInstance;
                Transform transform = geomInstance.Transform;
                if (null != geomInstance)
                {
                    foreach (GeometryObject insObj in geomInstance.SymbolGeometry)
                    {
                        if (insObj.GraphicsStyleId.IntegerValue != graphicsStyleId.IntegerValue)
                            continue;

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.Line")
                            continue;

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> points = polyLine.GetCoordinates();
                            CurveLoop prof = new CurveLoop() as CurveLoop;

                            for (int i = 0; i < points.Count - 1; i++)
                            {
                                if (points[i].DistanceTo(points[i + 1]) > Algorithm.CentimetersToUnits(0.1))
                                {
                                    Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(points[i], points[i + 1]);
                                    line = TransformLine(transform, line);
                                    prof.Append(line);
                                }
                            }
                            allOutlines.Add(prof);
                        }
                    }
                }
            }
            return allOutlines;
        }
        private Autodesk.Revit.DB.Line TransformLine(Transform transform, Autodesk.Revit.DB.Line line)
        {
            XYZ startPoint = transform.OfPoint(line.GetEndPoint(0));
            XYZ endPoint = transform.OfPoint(line.GetEndPoint(1));
            Autodesk.Revit.DB.Line newLine = Autodesk.Revit.DB.Line.CreateBound(startPoint, endPoint);
            return newLine;
        }
        public class StairModel
        {
            public StairModel()
            {
                ActualTreadsNumber = 10;
                TopElevation = "FL4";
                BaseElevation = "FL5";
                DesiredRisersNumber = 10;
                CurveLoop = null;
                Location = new XYZ(0, 0, 0);
                Thickness = 15;
            }

            public int ActualTreadsNumber { get; }
            public string TopElevation { get; }
            public string BaseElevation { get; }
            public int DesiredRisersNumber { get; set; }
            public CurveLoop CurveLoop { get; set; }
            public XYZ Location { get; set; }
            public double Thickness { get; set; }

        }
        public CurveLoop RoundCurveLoop(CurveLoop curveLoop, double gridline_size)
        {
            List<Autodesk.Revit.DB.Curve> curveList = curveLoop.ToList();
            CurveLoop curevLoop_new = new CurveLoop();
            List<Autodesk.Revit.DB.Curve> curveList_new = new List<Autodesk.Revit.DB.Curve>();
            foreach (Autodesk.Revit.DB.Curve curve in curveList)
            {
                XYZ point1 = Algorithm.RoundPoint(curve.GetEndPoint(0), gridline_size);
                XYZ point2 = Algorithm.RoundPoint(curve.GetEndPoint(1), gridline_size);
                Autodesk.Revit.DB.Curve newCurve = Autodesk.Revit.DB.Line.CreateBound(point1, point2) as Autodesk.Revit.DB.Curve;
                curevLoop_new.Append(newCurve);
            }
            return curevLoop_new;
        }
        public List<StairModel> GetLocation(List<StairModel> StairModels)
        {
            foreach (StairModel stair in StairModels)
            {
                List<double> x = new List<double>();
                List<double> y = new List<double>();
                double z = 0;
                foreach (Autodesk.Revit.DB.Curve c in stair.CurveLoop)
                {
                    x.Add(c.GetEndPoint(0).X);
                    x.Add(c.GetEndPoint(1).X);
                    y.Add(c.GetEndPoint(0).Y);
                    y.Add(c.GetEndPoint(1).Y);
                }
                XYZ midpoint = new XYZ((x.Max() + x.Min()) / 2, (y.Max() + y.Min()) / 2, z);//為啥麼是中點
                x.Clear();
                y.Clear();
                stair.Location = midpoint;
            }
            return StairModels;
        }
        public string GetCADPath(ElementId cadLinkTypeID, Document revitDoc)
        {
            CADLinkType cadLinkType = revitDoc.GetElement(cadLinkTypeID) as CADLinkType;
            return ModelPathUtils.ConvertModelPathToUserVisiblePath(cadLinkType.GetExternalFileReference().GetAbsolutePath());
        }
        public class LineComparer : IComparer<Line>
        {
            public int Compare(Line line1, Line line2)
            {
                XYZ startPoint1 = line1.GetEndPoint(0);
                XYZ startPoint2 = line2.GetEndPoint(0);

                // 根据 X 轴坐标进行排序，如果 X 轴坐标相同，则使用 Y 轴坐标进行排序
                if (startPoint1.X != startPoint2.X)
                    return startPoint1.X.CompareTo(startPoint2.X);
                else if (startPoint1.Y != startPoint2.Y)
                    return startPoint1.Y.CompareTo(startPoint2.Y);
                else
                    return startPoint1.Z.CompareTo(startPoint2.Z);
            }
        }
        public XYZ ConvertFeetToCentimeters(XYZ xyz)
        {
            // 将 XYZ 坐标从英尺转换为公分
            double factor = UnitUtils.ConvertFromInternalUnits(1.0, UnitTypeId.Centimeters);
            return new XYZ(xyz.X * factor, xyz.Y * factor, xyz.Z * factor);
        }

    }
}
