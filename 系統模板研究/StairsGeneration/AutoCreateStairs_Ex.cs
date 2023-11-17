using Aspose.Cells.Charts;
using Aspose.Pdf.Facades;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using Teigha.Geometry;
using Autodesk.Revit.DB.Mechanical;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateStairs_Ex : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            //定義樓梯建置之樓層
            Level levelBottom = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("4F")) as Level;
            Level levelTop = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("5F")) as Level;

            //建置樓梯
            CreateStairs1(doc, levelBottom, levelTop);

            //篩選所有的樓梯扶手
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ElementCategoryFilter filter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRailing);
            IList<Element> raillist = collector.WherePasses(filter).WhereElementIsNotElementType().ToElements();

            // 開始一個transaction，每個改變模型的動作都需在transaction中進行
            Transaction trans = new Transaction(doc);
            trans.Start("刪除元件");

            // 刪除選取的元件
            foreach (Element elem in raillist)
            {
                ElementId elemId = elem.Id;
                deleteElement(doc, elemId);
            }

            trans.Commit();

            return Result.Succeeded;
        }
        // The end of the main code.

        private ElementId CreateStairs1(Document document, Level levelBottom, Level levelTop)
        {
            ElementId newStairsId = null;

            using (StairsEditScope newStairsScope = new StairsEditScope(document, "New Stairs"))
            {
                newStairsId = newStairsScope.Start(levelBottom.Id, levelTop.Id);

                using (Transaction stairsTrans = new Transaction(document, "Add Runs and Landings to Stairs"))
                {
                    stairsTrans.Start();

                    // Create a sketched run for the stairs
                    IList<Curve> bdryCurves = new List<Curve>();
                    IList<Curve> riserCurves = new List<Curve>();
                    IList<Curve> pathCurves = new List<Curve>();
                    XYZ pnt1 = new XYZ(0, 0, 0);
                    XYZ pnt2 = new XYZ(15, 0, 0);
                    XYZ pnt3 = new XYZ(0, 10, 0);
                    XYZ pnt4 = new XYZ(15, 10, 0);

                    // boundaries       
                    bdryCurves.Add(Line.CreateBound(pnt1, pnt2));
                    bdryCurves.Add(Line.CreateBound(pnt3, pnt4));

                    // riser curves
                    const int riserNum = 20;
                    for (int ii = 0; ii <= riserNum; ii++)
                    {
                        XYZ end0 = (pnt1 + pnt2) * ii / (double)riserNum;
                        XYZ end1 = (pnt3 + pnt4) * ii / (double)riserNum;
                        XYZ end2 = new XYZ(end1.X, 10, 0);
                        riserCurves.Add(Line.CreateBound(end0, end2));
                    }

                    //stairs path curves
                    XYZ pathEnd0 = (pnt1 + pnt3) / 2.0;
                    XYZ pathEnd1 = (pnt2 + pnt4) / 2.0;
                    pathCurves.Add(Line.CreateBound(pathEnd0, pathEnd1));

                    StairsRun newRun1 = StairsRun.CreateSketchedRun(document, newStairsId, levelBottom.Elevation, bdryCurves, riserCurves, pathCurves);

                    // Add a straight run
                    Line locationLine = Line.CreateBound(new XYZ(20, -5, newRun1.TopElevation), new XYZ(35, -5, newRun1.TopElevation));
                    StairsRun newRun2 = StairsRun.CreateStraightRun(document, newStairsId, locationLine, StairsRunJustification.Center);
                    newRun2.ActualRunWidth = 10;

                    // Add a landing between the runs
                    CurveLoop landingLoop = new CurveLoop();
                    XYZ p1 = new XYZ(15, 10, 0);
                    XYZ p2 = new XYZ(20, 10, 0);
                    XYZ p3 = new XYZ(20, -10, 0);
                    XYZ p4 = new XYZ(15, -10, 0);
                    Line curve_1 = Line.CreateBound(p1, p2);
                    Line curve_2 = Line.CreateBound(p2, p3);
                    Line curve_3 = Line.CreateBound(p3, p4);
                    Line curve_4 = Line.CreateBound(p4, p1);

                    landingLoop.Append(curve_1);
                    landingLoop.Append(curve_2);
                    landingLoop.Append(curve_3);
                    landingLoop.Append(curve_4);
                    StairsLanding newLanding = StairsLanding.CreateSketchedLanding(document, newStairsId, landingLoop, newRun1.TopElevation);

                    stairsTrans.Commit();
                }
                // A failure preprocessor is to handle possible failures during the edit mode commitment process.
                newStairsScope.Commit(new StairsFailurePreprocessor());
            }

            return newStairsId;
        }

        private ElementId CreateStairs(Document document, Level levelBottom, Level levelTop)
        {
            ElementId newStairsId = null;

            using (StairsEditScope newStairsScope = new StairsEditScope(document, "New Stairs"))
            {
                newStairsId = newStairsScope.Start(levelBottom.Id, levelTop.Id);

                using (Transaction stairsTrans = new Transaction(document, "Add Runs and Landings to Stairs"))
                {
                    stairsTrans.Start();

                    // Add a landing1 between the runs
                    CurveLoop landingLoop1 = new CurveLoop();
                    XYZ p1_1 = new XYZ(0, 0, 0);
                    XYZ p1_2 = new XYZ(Algorithm.CentimetersToUnits(198), 0, 0);
                    XYZ p1_3 = new XYZ(Algorithm.CentimetersToUnits(198), Algorithm.CentimetersToUnits(122), 0);
                    XYZ p1_4 = new XYZ(0, Algorithm.CentimetersToUnits(122), 0);
                    //XYZ n1 = new XYZ(Algorithm.CentimetersToUnits(2), 0, 0);

                    Line curve1_1 = Line.CreateBound(p1_1, p1_2);
                    Line curve1_2 = Line.CreateBound(p1_2, p1_3);
                    Line curve1_3 = Line.CreateBound(p1_3, p1_4);
                    Line curve1_4 = Line.CreateBound(p1_4, p1_1);

                    landingLoop1.Append(curve1_1);
                    landingLoop1.Append(curve1_2);
                    landingLoop1.Append(curve1_3);
                    landingLoop1.Append(curve1_4);

                    StairsLanding newLanding1 = StairsLanding.CreateSketchedLanding(document, newStairsId, landingLoop1, levelBottom.Elevation);


                    // Create a sketched run for the stairs
                    //邊界曲線
                    IList<Curve> bdryCurves1 = new List<Curve>();
                    //踏板曲線
                    IList<Curve> riserCurves1 = new List<Curve>();
                    //路徑曲線
                    IList<Curve> pathCurves1 = new List<Curve>();

                    XYZ pnt1_1 = new XYZ(Algorithm.CentimetersToUnits(198), 0, 0);
                    XYZ pnt1_2 = new XYZ(Algorithm.CentimetersToUnits(458), 0, 0);
                    XYZ pnt1_3 = new XYZ(Algorithm.CentimetersToUnits(198), Algorithm.CentimetersToUnits(122), 0);
                    XYZ pnt1_4 = new XYZ(Algorithm.CentimetersToUnits(458), Algorithm.CentimetersToUnits(122), 0);

                    // boundaries       
                    bdryCurves1.Add(Line.CreateBound(pnt1_1, pnt1_2));
                    bdryCurves1.Add(Line.CreateBound(pnt1_3, pnt1_4));

                    // riser curves
                    const int riserNum1 = 10;
                    for (int ii = 0; ii <= riserNum1; ii++)
                    {
                        XYZ end1_0 = pnt1_1 + (pnt1_2 - pnt1_1) * ii / (double)riserNum1;
                        XYZ end1_1 = pnt1_1 + (pnt1_4 - pnt1_1) * ii / (double)riserNum1;
                        XYZ end1_2 = new XYZ(end1_1.X, Algorithm.CentimetersToUnits(122), 0);
                        riserCurves1.Add(Line.CreateBound(end1_0, end1_2));
                    }

                    //stairs path curves
                    XYZ pathEnd1_0 = (pnt1_1 + pnt1_3) / 2.0;
                    XYZ pathEnd1_1 = (pnt1_2 + pnt1_4) / 2.0;
                    pathCurves1.Add(Line.CreateBound(pathEnd1_0, pathEnd1_1));

                    StairsRun newRun1 = StairsRun.CreateSketchedRun(document, newStairsId, levelBottom.Elevation, bdryCurves1, riserCurves1, pathCurves1);


                    //// Add a straight run
                    //Line locationLine = Line.CreateBound(new XYZ(20, -5, newRun1.TopElevation), new XYZ(35, -5, newRun1.TopElevation));
                    //StairsRun newRun2 = StairsRun.CreateStraightRun(document, newStairsId, locationLine, StairsRunJustification.Center);
                    //newRun2.ActualRunWidth = 10;


                    //Add a landing2 between the runs
                    CurveLoop landingLoop2 = new CurveLoop();
                    XYZ p2_1 = new XYZ(Algorithm.CentimetersToUnits(458), 0, 0);
                    XYZ p2_2 = new XYZ(Algorithm.CentimetersToUnits(582), 0, 0);
                    XYZ p2_3 = new XYZ(Algorithm.CentimetersToUnits(582), Algorithm.CentimetersToUnits(250), 0);
                    XYZ p2_4 = new XYZ(Algorithm.CentimetersToUnits(458), Algorithm.CentimetersToUnits(250), 0);

                    Line curve2_1 = Line.CreateBound(p2_1, p2_2);
                    Line curve2_2 = Line.CreateBound(p2_2, p2_3);
                    Line curve2_3 = Line.CreateBound(p2_3, p2_4);
                    Line curve2_4 = Line.CreateBound(p2_4, p2_1);

                    landingLoop2.Append(curve2_1);
                    landingLoop2.Append(curve2_2);
                    landingLoop2.Append(curve2_3);
                    landingLoop2.Append(curve2_4);

                    StairsLanding newLanding2 = StairsLanding.CreateSketchedLanding(document, newStairsId, landingLoop2, newRun1.TopElevation);


                    // Create a sketched run for the stairs
                    //邊界曲線
                    IList<Curve> bdryCurves2 = new List<Curve>();
                    //踏板曲線
                    IList<Curve> riserCurves2 = new List<Curve>();
                    //路徑曲線
                    IList<Curve> pathCurves2 = new List<Curve>();
                    XYZ pnt2_1 = new XYZ(Algorithm.CentimetersToUnits(458), Algorithm.CentimetersToUnits(128), 0);
                    XYZ pnt2_2 = new XYZ(Algorithm.CentimetersToUnits(198), Algorithm.CentimetersToUnits(128), 0);
                    XYZ pnt2_3 = new XYZ(Algorithm.CentimetersToUnits(458), Algorithm.CentimetersToUnits(250), 0);
                    XYZ pnt2_4 = new XYZ(Algorithm.CentimetersToUnits(198), Algorithm.CentimetersToUnits(250), 0);

                    // boundaries       
                    bdryCurves2.Add(Line.CreateBound(pnt2_1, pnt2_2));
                    bdryCurves2.Add(Line.CreateBound(pnt2_3, pnt2_4));

                    // riser curves
                    const int riserNum2 = 10;
                    for (int ii = 0; ii <= riserNum2; ii++)
                    {
                        XYZ end2_0 = pnt2_2 + (pnt2_1 - pnt2_2) * ii / (double)riserNum2;
                        XYZ end2_1 = pnt2_2 + (pnt2_3 - pnt2_4) * ii / (double)riserNum2;
                        XYZ end2_2 = new XYZ(end2_1.X, Algorithm.CentimetersToUnits(250), 0);
                        riserCurves2.Add(Line.CreateBound(end2_0, end2_2));
                    }

                    //stairs path curves
                    XYZ pathEnd0 = (pnt2_1 + pnt2_3) / 2.0;
                    XYZ pathEnd1 = (pnt2_2 + pnt2_4) / 2.0;
                    pathCurves2.Add(Line.CreateBound(pathEnd1, pathEnd0));

                    StairsRun newRun2 = StairsRun.CreateSketchedRun(document, newStairsId, newRun1.TopElevation, bdryCurves2, riserCurves2, pathCurves2);

                    stairsTrans.Commit();

                }

                // A failure preprocessor is to handle possible failures during the edit mode commitment process.
                newStairsScope.Commit(new StairsFailurePreprocessor());
            }

            return newStairsId;
        }
        private void deleteElement(Autodesk.Revit.DB.Document document, ElementId elemId)
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
    }
}
