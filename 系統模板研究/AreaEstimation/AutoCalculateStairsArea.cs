using Autodesk.Revit.UI.Selection;
using System.Text;
using System.Windows.Forms;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Architecture;
using System.Collections.Generic;
using System;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class AutoCalculateStairsArea : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            //Reference stairRef = uidoc.Selection.PickObject(ObjectType.Element, "選擇一個樓梯");
            //ICollection<ElementId> ids = uidoc.Selection.GetElementIds();

            //Stairs
            FilteredElementCollector stairsCollector = new FilteredElementCollector(doc);
            ElementCategoryFilter stairsFilter = new ElementCategoryFilter(BuiltInCategory.OST_Stairs);
            IList<Element> stairList = stairsCollector.WherePasses(stairsFilter).WhereElementIsNotElementType().ToElements();

            StringBuilder stairsParameterst = new StringBuilder();

            double stairsActualNumRisersdou = 0;
            double stairsActualTreadDepthdou = 0;
            double stairsActualRiserHightdou = 0;
            double stairsHightdou = 0;
            double stairPathLengthdou = 0;
            double structuralDepthdou = 0;
            double stairsLandingThicknessdou = 0;

            foreach (Element stair in stairList)
            {
                foreach (Parameter para in stair.Parameters)
                {
                    string defName = para.Definition.Name;
                    if (defName == "實際梯級數")
                    {
                        string stairsActualNumRisers = defName + ":" + para.AsValueString();
                        stairsParameterst.AppendLine(stairsActualNumRisers);

                        stairsActualNumRisersdou = Convert.ToInt32(para.AsValueString());
                    }
                    if (defName == "實際級深")
                    {
                        string stairsActualTreadDepth = defName + ":" + para.AsValueString();
                        stairsParameterst.AppendLine(stairsActualTreadDepth);

                        stairsActualTreadDepthdou = Convert.ToInt32(para.AsValueString());
                    }
                    if (defName == "實際級高")
                    {
                        string stairsActualRiserHight = defName + ":" + para.AsValueString();
                        stairsParameterst.AppendLine(stairsActualRiserHight);

                        stairsActualRiserHightdou = Convert.ToInt32(para.AsValueString());
                    }
                    if (defName == "所需樓梯高度")
                    {
                        string stairsHight = defName + ":" + para.AsValueString();
                        stairsParameterst.AppendLine(stairsHight);

                        stairsHightdou = Convert.ToInt32(para.AsValueString());
                    }
                }
            }

            // StairsRuns
            FilteredElementCollector stairsRunsCollector = new FilteredElementCollector(doc);
            ElementCategoryFilter stairsRunsFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsRuns);
            IList<Element> stairRunsList = stairsRunsCollector.WherePasses(stairsRunsFilter).WhereElementIsNotElementType().ToElements();
            foreach (StairsRun stairRun in stairRunsList)
            {
                CurveLoop stairPathCurveLoop = null;
                stairPathCurveLoop = stairRun.GetStairsPath();
                foreach (Curve stairPathCurve in stairPathCurveLoop)
                {
                    string stairPathLength = "梯段路徑:" + UnitsToCentimeters(stairPathCurve.Length).ToString();
                    stairsParameterst.AppendLine(stairPathLength);

                    stairPathLengthdou = Convert.ToInt32(UnitsToCentimeters(stairPathCurve.Length).ToString());
                }

                StairsRunType stairsRunType = doc.GetElement(stairRun.GetTypeId()) as StairsRunType;
                string structuralDepth = "梯段結構深度:" + UnitsToCentimeters(stairsRunType.StructuralDepth).ToString();
                stairsParameterst.AppendLine(structuralDepth);

                structuralDepthdou = Convert.ToInt32(UnitsToCentimeters(stairsRunType.StructuralDepth).ToString());
            }

            // StairsLandings
            FilteredElementCollector stairsLandingsCollector = new FilteredElementCollector(doc);
            ElementCategoryFilter stairsLandingsFilter = new ElementCategoryFilter(BuiltInCategory.OST_StairsLandings);
            IList<Element> stairLandingsList = stairsLandingsCollector.WherePasses(stairsLandingsFilter).WhereElementIsNotElementType().ToElements();
            foreach (StairsLanding stairLanding in stairLandingsList)
            {
                string stairsLandingThickness = "樓梯平台厚度:" + stairLanding.get_Parameter(BuiltInParameter.STAIRS_LANDING_THICKNESS).AsValueString();
                stairsParameterst.AppendLine(stairsLandingThickness);

                stairsLandingThicknessdou = Convert.ToInt32(stairLanding.get_Parameter(BuiltInParameter.STAIRS_LANDING_THICKNESS).AsValueString());
            }

            Reference paramerterRef = uidoc.Selection.PickObject(ObjectType.Element);

            XYZ staitsLandingWidthendpoint = uidoc.Selection.PickPoint();
            XYZ staitsLandingWidthcurrentPoint = uidoc.Selection.PickPoint();
            Line staitsLandingWidthline = Line.CreateBound(staitsLandingWidthendpoint, staitsLandingWidthcurrentPoint);
            double staitsLandingWidthdou = UnitsToCentimeters(staitsLandingWidthline.Length);
            string staitsLandingWidth = "平台寬度" + staitsLandingWidthdou.ToString();
            stairsParameterst.AppendLine(staitsLandingWidth);

            XYZ stairsLandingLengthendPoint = uidoc.Selection.PickPoint();
            XYZ stairsLandingLengthcurrentPoint = uidoc.Selection.PickPoint();
            Line staitsLandingLengthline = Line.CreateBound(stairsLandingLengthendPoint, stairsLandingLengthcurrentPoint);
            double staitsLandingLengthdou = UnitsToCentimeters(staitsLandingLengthline.Length);
            string staitsLandingLength = "平台長度" + staitsLandingLengthdou.ToString();
            stairsParameterst.AppendLine(staitsLandingLength);

            XYZ staitsRunsWidthendpoint = uidoc.Selection.PickPoint();
            XYZ staitsRunsWidthcurrentPoint = uidoc.Selection.PickPoint();
            Line staitsRunsWidthline = Line.CreateBound(staitsRunsWidthendpoint, staitsRunsWidthcurrentPoint);
            double staitsRunsWidthdou = UnitsToCentimeters(staitsRunsWidthline.Length);
            string staitsRunsWidth = "梯段實際寬度" + staitsRunsWidthdou.ToString();
            stairsParameterst.AppendLine(staitsRunsWidth);

            double staitsRunsPathhypotenusedou = Math.Sqrt(Math.Pow(stairPathLengthdou, 2) + Math.Pow(stairsHightdou / 2, 2));
            string staitsRunsPathhypotenuse = "梯段路徑斜面長度:" + staitsRunsPathhypotenusedou.ToString();
            stairsParameterst.AppendLine(staitsRunsPathhypotenuse);

            MessageBox.Show(stairsParameterst.ToString(), "Revit", MessageBoxButtons.OK);

            StringBuilder area = new StringBuilder();

            double stairsRunsPathArea = staitsRunsWidthdou * (staitsRunsPathhypotenusedou * 2);
            double stairsRunsThreadsArea = staitsRunsWidthdou * (stairPathLengthdou * 2);
            double stairsRunsRisersArea = staitsRunsWidthdou * (stairsActualNumRisersdou * stairsActualRiserHightdou);
            double stairsLandingsArea = staitsLandingWidthdou * staitsLandingLengthdou;
            double total = stairsRunsPathArea + stairsRunsThreadsArea + stairsRunsRisersArea + stairsLandingsArea;

            string stairsRunsPathAreast = "樓梯路徑斜面面積=" + stairsRunsPathArea.ToString();
            area.AppendLine(stairsRunsPathAreast);
            string stairsRunsRisersAreast = "樓梯梯段立板面積=" + stairsRunsRisersArea.ToString();
            area.AppendLine(stairsRunsRisersAreast);
            string stairsRunsThreadsAreast = "樓梯梯段踏板面積=" + stairsRunsThreadsArea.ToString();
            area.AppendLine(stairsRunsThreadsAreast);
            string stairsLandingsAreast = "樓梯平台面積=" + stairsLandingsArea.ToString();
            area.AppendLine(stairsLandingsAreast);
            string totalst = "樓梯總面積=" + total.ToString();
            area.AppendLine(totalst);

            MessageBox.Show(area.ToString(), "Revit", MessageBoxButtons.OK);

            return Result.Succeeded;
        }



        public static double UnitsToCentimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
        }
    }
}