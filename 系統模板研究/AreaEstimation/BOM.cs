using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Forms;
using System.Collections.Generic;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.ApplicationServices;
using Application = Autodesk.Revit.ApplicationServices.Application;
using System;
using Autodesk.Revit.UI.Selection;
using System.IO;
using System.Text;
using System.Linq;
using Aspose.Cells.Charts;
using Aspose.Pdf.Facades;
using System.Windows.Media.Animation;
using Teigha.DatabaseServices;
using Solid = Autodesk.Revit.DB.Solid;
using Line = Autodesk.Revit.DB.Line;
using Face = Autodesk.Revit.DB.Face;
using Curve = Autodesk.Revit.DB.Curve;
using System.Windows.Media;
using Dimension = Autodesk.Revit.DB.Dimension;
using System.Drawing.Printing;
using Aspose.Pdf;


[Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
public class BOM : IExternalCommand
{
    public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
    {

        this.Run(commandData.Application);

        MessageBox.Show("BOM表匯出成功!");
        return Result.Succeeded;
    }
    class BOMData
    {
        public string 名稱 { get; set; }
        public double 實際梯段寬度 { get; set; }
        public double 實際級高 { get; set; }
        public double 實際級深 { get; set; }
        public int 實際梯級數 { get; set; }
        public double 所有梯段路徑長度 { get; set; }
        public double 突沿長度 { get; set; }
        public double 所有平台底部高度 { get; set; }
        //public double 中間平台底部高度 { get; set; }
        //public double 結束平台底部高度 { get; set; }
        public double 所有梯段底部斜長度 { get; set; }
        //public double 第一梯段底部斜長度 { get; set; }
        //public double 第二梯段底部斜長度 { get; set; }
        public double 所有平台頂面面積 { get; set; }
        public double 所有平台底面面積 { get; set; }
        //public double 中間平台頂面面積 { get; set; }
        //public double 中間平台底面面積 { get; set; }
        //public double 結束平台底面面積 { get; set; }
        public double 平台總面積 { get; set; }
        public double 踏板面積 { get; set; }
        public double 立板面積 { get; set; }
        public double 所有梯段斜面面積 { get; set; }
        public double 總面積 { get; set; }

        public BOMData()
        {
            名稱 = null;
            實際梯段寬度 = 0;
            實際級高 = 0;
            實際級深 = 0;
            實際梯級數 = 0;
            所有梯段路徑長度 = 0;
            突沿長度 = 0;

            所有平台底部高度 = 0;
            // 中間平台底部高度 = 0;
            // 結束平台底部高度 = 0;

            所有平台頂面面積 = 0;
            所有平台底面面積 = 0;
            // 中間平台頂面面積
            // 中間平台底面面積
            // 結束平台底面面積

            所有梯段底部斜長度 = (所有平台底部高度) / 實際級高 * (Math.Sqrt(Math.Pow(實際級高, 2) + Math.Pow(實際級深, 2)));
            // 第一梯段底部斜長度
            // 第二梯段底部斜長度

            踏板面積 = (實際梯段寬度 * 所有梯段路徑長度 * 2) / 10000;
            立板面積 = (實際梯段寬度 * 實際級高 * 實際梯級數) / 10000;
            所有梯段斜面面積 = 所有梯段底部斜長度 * 實際梯段寬度;
            平台總面積 = 所有平台頂面面積 + 所有平台底面面積;

            總面積 = 踏板面積 + 立板面積 + 所有梯段斜面面積 + 平台總面積;
        }
    }


    void WriteToCSV(List<BOMData> data)
    {
        SaveFileDialog dlg = new SaveFileDialog();
        dlg.FileName = "StairParameter"; // 預設檔案名稱
        dlg.DefaultExt = ".csv"; // 預設檔案副檔名
        dlg.Filter = "CSV documents (.csv)|*.csv"; // 檔案篩選器
        dlg.ShowDialog();

        using (var file = new StreamWriter(dlg.FileName, false, Encoding.UTF8))
        {

            foreach (var item in data)
            {
                //file.WriteLine($"{"名稱"},{"實際梯段寬度(RW)"},{"實際級高(RH)"},{"實際級深(RD)"},{"實際梯級數(RN)"},{"梯段路徑長度(RPL)"}" +
                //          $",{"平台長度(LL)"},{"平台寬度(LW)"}"/);

                //file.WriteLine($"{item.名稱},{item.實際梯段寬度},{item.實際級高},{item.實際級深},{item.實際梯級數},{item.梯段路徑長度}" +
                //               $",{item.平台長度},{item.平台寬度}");

                file.WriteLine($"{"名稱"},{item.名稱}");

                file.WriteLine();

                file.WriteLine($"{"樓梯參數"}");
                file.WriteLine($"{"1. 實際梯段寬度"},{item.實際梯段寬度}");
                file.WriteLine($"{"2. 實際級高"},{item.實際級高}");
                file.WriteLine($"{"3. 實際級深"},{item.實際級深}");
                file.WriteLine($"{"4. 實際梯級數"},{item.實際梯級數}");
                file.WriteLine($"{"5. 所有梯段路徑長度"},{item.所有梯段路徑長度}");
                file.WriteLine($"{"6. 所有平台底部高度"},{item.所有平台底部高度}");
                // file.WriteLine($"{"1. 中間平台底部高度"},{item.中間平台底部高度}");
                // file.WriteLine($"{"2. 結束平台底部高度"},{item.結束平台底部高度}");
                file.WriteLine($"{"7. 所有梯段底部斜長度"},{item.所有梯段底部斜長度}");

                file.WriteLine();

                file.WriteLine($"{"面積計算"}");
                file.WriteLine($"{"1.所有平台頂面面積"},{item.所有平台頂面面積}");
                file.WriteLine($"{"2.所有平台底面面積"},{item.所有平台底面面積}");
                file.WriteLine($"{"3. 踏板面積 = 實際梯段寬度 * 所有梯段路徑長度"},{item.踏板面積}");
                file.WriteLine($"{"4. 立板面積 = 實際梯段寬度 * 實際級高 * 實際梯級數 "},{item.立板面積}");
                file.WriteLine($"{"5. 所有梯段斜面面積 = 所有梯段底部斜長度 * 實際梯段寬度  "},{item.所有梯段斜面面積}");
                file.WriteLine($"{"6. 平台總面積 = 所有平台頂面面積 + 所有平台底面面積"},{item.平台總面積}");
                file.WriteLine($"{"7. 總面積 = 踏板面積 + 立板面積 + 梯段總斜面積 + 平台總面積 "},{item.總面積},{"單位:m^2"}");
                // 添加空行分隔每个 BOM 条目
            }

        }
    }
    public void Run(UIApplication uiapp)
    {
        UIDocument uidoc;
        uidoc = uiapp.ActiveUIDocument;
        Autodesk.Revit.DB.Document doc = uidoc.Document;

        //所有所需參數
        string panelName = null;
        string panelCode = null;
        double stairRunsWidth = 0;
        double stairActualRiserHight = 0;
        double stairActualTreadDepth = 0;
        int stairActualNumRisers = 0;
        double stairPathLength = 0;
        double stairRunBottomLength = 0;
        double stairLandingThick = 0;

        double stairLandingBottomLevel = 0;
        // double m_stairLandingBottomLevel = 0;
        // double e_stairLandingBottomLevel = 0;

        double threadsArea = 0;
        double risersArea = 0;
        double stairRunBottomArea = 0;
        double landing_allTopFaceArea = 0;
        double landing_allBottomFaceArea = 0;

        double landingArea1 = 0; // 中間平台的底+頂面積
        double landingArea2 = 0; // 結束平的的底面積

        double alllandingArea = 0;
        double totalArea = 0;


        //中間平台頂面面積 = 0;
        //中間平台底面面積 = 0;
        //結束平台底面面積 = 0;

        //點選樓梯
        Reference paramerterRef = uidoc.Selection.PickObject(ObjectType.Element);
        ElementId selectedElementId = paramerterRef.ElementId;
        Element selectedElement = doc.GetElement(selectedElementId);
        string salectedCategoryName = selectedElement.Category.Name;

        //ElementId stairId = selectedElement.LevelId;
        //Level stairLevel = doc.GetElement(stairId) as Level;
        //double stairElevation = stairLevel.Elevation;

        IList<Element> stairList = new List<Element>();
        IList<Element> stairRunsList = new List<Element>();
        IList<Element> stairLandingsList = new List<Element>();


        if (selectedElement != null && selectedElement is Stairs)
        {
            Stairs stair = (Stairs)selectedElement;

            Element panelElement1 = doc.GetElement(stair.Id);
            FamilyInstance familyInstance = panelElement1 as FamilyInstance;
            ElementId pickedtypeId = panelElement1.GetTypeId();
            ElementType family = doc.GetElement(pickedtypeId) as ElementType;

            panelName = panelElement1.Name.ToString();
            panelCode = family.FamilyName.ToString();

            //樓梯可以抓到的參數
            stairActualRiserHight = (double)UnitUtils.Convert(stair.LookupParameter("實際級高").AsDouble(), UnitTypeId.Feet, UnitTypeId.Centimeters);
            stairActualTreadDepth = (double)UnitUtils.Convert(stair.LookupParameter("實際級深").AsDouble(), UnitTypeId.Feet, UnitTypeId.Centimeters);
            string stairActualNumRisersst = stair.LookupParameter("實際梯級數").AsValueString();
            stairActualNumRisers = Convert.ToInt32(stairActualNumRisersst);


            ICollection<ElementId> stairLandingsICollectionId = stair.GetStairsLandings();
            //MessageBox.Show("平台:" + stairLandingsICollectionId.Count.ToString());
            ICollection<ElementId> stairRunsICollectionId = stair.GetStairsRuns();
            //MessageBox.Show("梯段:" + stairRunsICollectionId.Count.ToString());

            //選到的樓梯裡的梯段
            foreach (ElementId stairsRunsId in stairRunsICollectionId)
            {
                Element elem = doc.GetElement(stairsRunsId);
                if (elem != null)
                {
                    stairRunsList.Add(elem);
                }
            }

            //抓取梯段路徑長度參數
            int stairsRunsCount = 0;
            foreach (StairsRun stairRun in stairRunsList)
            {
                stairsRunsCount++;
                CurveLoop stairPathCurveLoop = null;
                stairPathCurveLoop = stairRun.GetStairsPath();
                foreach (Curve stairPathCurve in stairPathCurveLoop)
                {
                    stairPathLength += UnitsToCentimeters(stairPathCurve.Length);
                }
            }

            List<XYZ> runPoints = new List<XYZ>();
            Face riserTopFace = null;

            foreach (Element stairRun in stairRunsList)
            {
                GeometryElement geometryElement = stairRun.get_Geometry(new Options());

                // 編歷Geometry對象以獲取尺寸信息
                foreach (GeometryObject geomObj in geometryElement)
                {
                    // 抓取在實體上的幾何參數
                    if (geomObj is Solid solid)
                    {
                        // 抓取幾何面
                        FaceArray faces = solid.Faces;
                        //MessageBox.Show("梯段面數量:" + faces.Size.ToString());
                        foreach (Face face in faces)
                        {
                            XYZ targetFaceNormal = face.ComputeNormal(UV.Zero);
                            XYZ Zdirection = XYZ.BasisZ;
                            double dotProduct = Zdirection.DotProduct(targetFaceNormal);
                            //Zdirection.DotProduct(targetFaceNormal))  == 1 或是 == -1

                            if (Math.Abs(dotProduct - 1.0) < 1e-9)
                            {
                                riserTopFace = face;
                                //MessageBox.Show("抓到頂面");

                                stairRunsWidth = 0.0;
                                Edge longestEdge = null;

                                foreach (EdgeArray edgeArray in riserTopFace.EdgeLoops)
                                {
                                    foreach (Edge edge in edgeArray)
                                    {
                                        // 取得邊的幾何曲線
                                        Curve curve = edge.AsCurve();

                                        // 檢查邊的長度
                                        double edgeLength = curve.Length;
                                        if (edgeLength > stairRunsWidth)
                                        {
                                            stairRunsWidth = UnitsToCentimeters(edgeLength);
                                            longestEdge = edge;
                                        }
                                    }
                                }

                                if (longestEdge != null)
                                {
                                    // 現在 longestEdge 包含最長的邊
                                    //MessageBox.Show("最長的邊長度:" + stairRunsWidth.ToString());
                                }
                                else
                                {
                                    //MessageBox.Show("未找到邊");
                                }
                                break;
                            }
                        }
                    }
                }
            }


            //選到的樓梯裡的平台
            foreach (ElementId stairLandingsId in stairLandingsICollectionId)
            {
                Element elem = doc.GetElement(stairLandingsId);
                if (elem != null)
                {
                    stairLandingsList.Add(elem);
                }

            }

            foreach (StairsLanding stairsLanding in stairLandingsList)
            {
                // 須減平台厚度
                stairLandingThick = UnitsToCentimeters(stairsLanding.Thickness);
                stairLandingBottomLevel = UnitsToCentimeters(stairsLanding.BaseElevation) - stairLandingThick;
            }


            //foreach (StairsLanding stairLanding in stairLandingsList)
            //{
            //    if (stairLanding.BaseElevation == stairElevation)
            //    {
            //        stairLandingsList.Remove(stairLanding);
            //    }
            //}

            //平台的參數
            foreach (StairsLanding stairLanding in stairLandingsList)
            {
                // 獲取樓梯平台的Geometry對象
                GeometryElement geometryElement = stairLanding.get_Geometry(new Options());

                List<Line> yLines = new List<Line>(); // 存 Y 相同的點的，其 X 最大和最小值相連的線
                List<Line> xLines = new List<Line>(); // 存 X 相同的點的，其 Y 最大和最小值相連的線

                Face topFace = null;
                Face bottomFace = null;

                // List<double> allTopFaceArea = new List<double>();
                // List<double> allBottomFaceArea = new List<double>();

                XYZ direction = new XYZ(0, 0, 0);
                List<CurveLoop> topOutlines = new List<CurveLoop>();
                List<CurveLoop> bottomOutlines = new List<CurveLoop>();

                if (UnitsToCentimeters(stairLanding.BaseElevation) != stairActualRiserHight) 
                {
                    // 編歷Geometry對象以獲取尺寸信息
                    foreach (GeometryObject geomObj in geometryElement)
                    {
                        GeometryInstance geomInstance = geomObj as GeometryInstance;
                        // Autodesk.Revit.DB.Transform transform = geomInstance.Transform;
                        bool isValid = true;

                        // 抓取在實體上的幾何參數
                        if (geomObj is Solid solid)
                        {
                            // 抓取幾何面
                            FaceArray faces = solid.Faces;
                            //MessageBox.Show("平台面數量:" + faces.Size.ToString());
                            foreach (Face face in faces)
                            {
                                XYZ targetFaceNormal = face.ComputeNormal(UV.Zero);
                                XYZ Zdirection = XYZ.BasisZ;
                                double dotProduct = Zdirection.DotProduct(targetFaceNormal);
                                //Zdirection.DotProduct(targetFaceNormal))  == 1 或是 == -1

                                if (Math.Abs(dotProduct - 1.0) < 1e-9)
                                {
                                    topFace = face;
                                    //MessageBox.Show("這是頂面");
                                    // 計算頂面的面積
                                    //double topFaceArea = SquareUnitToSquareCentimeter(topFace.Area);
                                    //allTopFaceArea.Add(topFaceArea);
                                    // MessageBox.Show("頂面面積: " + topFaceArea.ToString());

                                    // 获取 topFace 的 CurveLoop
                                    // CurveLoop curveLoop = GetCurveLoopFromFace(topFace);
                                    PolyLine polyLine = ConvertFaceEdgesToPolyLine(topFace);

                                    List<XYZ> points_list = new List<XYZ>(polyLine.GetCoordinates());
                                    //foreach (XYZ point in points_list)
                                    //{
                                    //    MessageBox.Show("TopFacePoints:" + point.ToString());
                                    //}

                                    CurveLoop prof = new CurveLoop() as CurveLoop;

                                    for (int i = 0; i < points_list.Count - 1; i++)
                                    {

                                        if (points_list[i].DistanceTo(points_list[i + 1]) < CentimetersToUnits(0.1))
                                        {
                                            continue;
                                        }
                                        XYZ line_direction = (points_list[i] - points_list[i + 1]).Normalize();
                                        if (line_direction == direction)
                                        {
                                            prof = new CurveLoop() as CurveLoop;
                                            isValid = false;
                                            break;
                                        }
                                        Line line = Line.CreateBound(points_list[i], points_list[i + 1]);
                                        //line = TransformLine(transform, line);
                                        prof.Append(line);
                                    }

                                    if (isValid)
                                    {
                                        topOutlines.Add(prof);
                                    }

                                    List<CurveLoop> landingCurves = new List<CurveLoop>();

                                    foreach (CurveLoop topOutLine in topOutlines)
                                    {
                                        foreach (Curve c in topOutLine)
                                        {
                                            MessageBox.Show("StartPoint:" + c.GetEndPoint(0).ToString() + "\n" +
                                                            "EndPoint:" + c.GetEndPoint(1).ToString());
                                        }
                                        landingCurves = GetRecSlabsFromPolySlab(topOutLine, 5);
                                        MessageBox.Show(landingCurves.Count.ToString());
                                    }

                                    List<double> allTopFaceLength = new List<double>();
                                    List<double> allTopFaceWidth = new List<double>();

                                    double topFaceLength = 0;
                                    double topFaceWidth = 0;

                                    foreach (CurveLoop curveLoop in landingCurves)
                                    {
                                        List<double> curveLengths = new List<double>();
                                        foreach (Curve curve in curveLoop)
                                        {
                                            double curveLength = Math.Abs(Math.Sqrt(Math.Pow(curve.GetEndPoint(0).X,2)+ Math.Pow(curve.GetEndPoint(0).Y, 2)) - 
                                                                          Math.Sqrt(Math.Pow(curve.GetEndPoint(1).X, 2) + Math.Pow(curve.GetEndPoint(1).Y, 2)));
                                            curveLengths.Add(curveLength);
                                        }
                                        topFaceLength = curveLengths[0];
                                        topFaceWidth = curveLengths[1];
                                        allTopFaceLength.Add(topFaceLength);
                                        allTopFaceWidth.Add(topFaceWidth);
                                    }
                                    
                                    List<double> landingArea = new List<double>();
                                    for(int i = 0; i < allTopFaceLength.Count; i++)
                                    {
                                        double lA = 0;
                                        lA += allTopFaceLength[i]* allTopFaceWidth[i];
                                    }

                                }

                                else if (Math.Abs(dotProduct + 1.0) < 1e-9)
                                {
                                    bottomFace = face;
                                    // MessageBox.Show("這是底面");
                                    // 計算底面的面積
                                    // double bottomFaceArea = SquareUnitToSquareCentimeter(bottomFace.Area);
                                    // allBottomFaceArea.Add(bottomFaceArea);
                                    // MessageBox.Show("底面面積: " + bottomFaceArea.ToString());

                                    //获取 topFace 的 CurveLoop
                                    // CurveLoop curveLoop = GetCurveLoopFromFace(bottomFace);
                                    PolyLine polyLine = ConvertFaceEdgesToPolyLine(bottomFace);

                                    List<XYZ> points_list = new List<XYZ>(polyLine.GetCoordinates());
                                    foreach (XYZ point in points_list)
                                    {
                                        MessageBox.Show(point.ToString());
                                    }

                                    CurveLoop prof = new CurveLoop() as CurveLoop;

                                    for (int i = 0; i < points_list.Count - 1; i++)
                                    {

                                        if (points_list[i].DistanceTo(points_list[i + 1]) < CentimetersToUnits(0.1))
                                        {
                                            continue;
                                        }
                                        XYZ line_direction = (points_list[i] - points_list[i + 1]).Normalize();
                                        if (line_direction == direction)
                                        {
                                            prof = new CurveLoop() as CurveLoop;
                                            isValid = false;
                                            break;
                                        }
                                        Line line = Line.CreateBound(points_list[i], points_list[i + 1]);
                                        //line = TransformLine(transform, line);
                                        prof.Append(line);

                                    }

                                    if (isValid)
                                    {
                                        bottomOutlines.Add(prof);
                                    }

                                    foreach (CurveLoop c in bottomOutlines)
                                    {
                                        List<CurveLoop> landingCurves = GetRecSlabsFromPolySlab(c, 5);
                                        MessageBox.Show(landingCurves.Count.ToString());
                                    }

                                    List<double> allBottomFaceLength = new List<double>();
                                    List<double> allBottomFaceWidth = new List<double>();
                                    double bottomFaceLength = 0;
                                    double bottomFaceWidth = 0;

                                }
                            }
                        }

                        //// 創建一個字典，用於存儲不同 Z 座標的點列表
                        //Dictionary<double, List<XYZ>> pointDictionary_Top = GroupPointsByZ(points_Top);
                        //Dictionary<double, List<XYZ>> pointDictionary_Bottom = GroupPointsByZ(points_Bottom);

                        //List<Line> verticalLine_Top = LandingBoundaryLines(pointDictionary_Top).Item1;
                        //List<Line> horrizontalLine_Top = LandingBoundaryLines(pointDictionary_Top).Item2;

                        //List<Line> verticalLine_Bottom = LandingBoundaryLines(pointDictionary_Bottom).Item1;
                        //List<Line> horrizontalLine_Bottom = LandingBoundaryLines(pointDictionary_Bottom).Item2;
                    }
                }

                //for (int i = 0; i < allTopFaceArea.Count; i++)
                //{

                //    landing_allTopFaceArea += allTopFaceArea[i];
                //}

                //for (int i = 0; i < allBottomFaceArea.Count; i++)
                //{
                //    landing_allBottomFaceArea += allBottomFaceArea[i];
                //}


            }
        }

        stairRunBottomLength = stairLandingBottomLevel / stairActualRiserHight * (Math.Sqrt(Math.Pow(stairActualRiserHight, 2) + Math.Pow(stairActualTreadDepth, 2)));

        landing_allTopFaceArea = landing_allTopFaceArea / 10000;
        landing_allBottomFaceArea = landing_allBottomFaceArea / 10000;
        alllandingArea = (double)(landing_allTopFaceArea + landing_allBottomFaceArea);
        threadsArea = (double)(stairRunsWidth * stairPathLength) / 10000;
        risersArea = (double)(stairRunsWidth * stairActualRiserHight * stairActualNumRisers) / 10000;

        stairRunBottomArea = (double)(stairRunsWidth * stairRunBottomLength) / 10000;

        totalArea = (double)(threadsArea + risersArea + stairRunBottomArea + alllandingArea);

        MessageBox.Show(panelName + "\n" + panelCode + "\n" +
            stairActualRiserHight.ToString() + "\n" + stairActualTreadDepth.ToString() + "\n" +
            stairActualNumRisers.ToString() + "\n" + stairRunBottomLength.ToString() + "\n" +
            alllandingArea.ToString() + "\n" + threadsArea.ToString() + "\n" + risersArea.ToString() + "\n" +
            stairRunBottomArea.ToString() + "\n" + totalArea.ToString());


        //BOM表
        List<BOMData> BOMDatas = new List<BOMData>();
        Dictionary<string, BOMData> keyValuePairs = new Dictionary<string, BOMData>();

        BOMData bOMData = new BOMData()
        {

            名稱 = panelName,
            實際梯段寬度 = stairRunsWidth,
            實際級高 = stairActualRiserHight,
            實際級深 = stairActualTreadDepth,
            實際梯級數 = stairActualNumRisers,
            所有梯段路徑長度 = stairPathLength,
            所有平台底部高度 = stairLandingBottomLevel,
            所有梯段底部斜長度 = stairRunBottomLength,

            所有平台頂面面積 = landing_allTopFaceArea,
            所有平台底面面積 = landing_allBottomFaceArea,
            踏板面積 = threadsArea,
            立板面積 = risersArea,
            所有梯段斜面面積 = stairRunBottomArea,
            平台總面積 = alllandingArea,
            總面積 = totalArea
        };
        keyValuePairs[panelName] = bOMData;


        foreach (BOMData bOMData1 in keyValuePairs.Values)
        {
            BOMDatas.Add(bOMData1);
        }

        WriteToCSV(BOMDatas);
    }


    private Autodesk.Revit.DB.Line TransformLine(Autodesk.Revit.DB.Transform transform, Autodesk.Revit.DB.Line line)
    {
        XYZ startPoint = transform.OfPoint(line.GetEndPoint(0));
        XYZ endPoint = transform.OfPoint(line.GetEndPoint(1));
        Autodesk.Revit.DB.Line newLine = Autodesk.Revit.DB.Line.CreateBound(startPoint, endPoint);
        return newLine;
    }
    private List<Curve> SortCurvesClockwise(List<Curve> curves)
    {
        // 计算曲线的中心点
        XYZ centerPoint = XYZ.Zero;
        foreach (Curve curve in curves)
        {
            XYZ startPoint = curve.GetEndPoint(0);
            XYZ endPoint = curve.GetEndPoint(1);
            centerPoint = centerPoint.Add(startPoint.Add(endPoint).Divide(2.0));
        }
        centerPoint = centerPoint.Divide(curves.Count);

        // 使用法向量方向进行排序
        List<Curve> sortedCurves = curves.OrderBy(c =>
        {
            XYZ startPoint = c.GetEndPoint(0);
            XYZ endPoint = c.GetEndPoint(1);
            XYZ midPoint = startPoint.Add(endPoint).Divide(2.0);

            // 计算中点相对于中心点的法向量
            XYZ normalVector = midPoint.Subtract(centerPoint).Normalize();

            // 计算法向量相对于X轴的角度
            double angle = Math.Atan2(normalVector.Y, normalVector.X);

            // 将角度映射到[0, 2π)范围内
            if (angle < 0)
                angle += 2 * Math.PI;

            return angle;
        }).ToList();

        return sortedCurves;
    }

    // Function to calculate the centroid of a list of curves
    private XYZ GetCentroid(List<Curve> curves)
    {
        double totalArea = 0;
        XYZ centroid = XYZ.Zero;

        foreach (Curve curve in curves)
        {
            XYZ curveCentroid = curve.Evaluate(0.5, true);
            double curveArea = curve.Length; // You may need to use a more accurate method to calculate area

            centroid += curveCentroid * curveArea;
            totalArea += curveArea;
        }

        if (totalArea > 0)
        {
            centroid /= totalArea;
        }

        return centroid;
    }

    // Function to compare curves based on their angle relative to a reference point (centroid)
    private int CompareCurvesByAngle(Curve c1, Curve c2, XYZ referencePoint)
    {
        XYZ midpoint1 = c1.Evaluate(0.5, true);
        XYZ midpoint2 = c2.Evaluate(0.5, true);

        double angle1 = Math.Atan2(midpoint1.Y - referencePoint.Y, midpoint1.X - referencePoint.X);
        double angle2 = Math.Atan2(midpoint2.Y - referencePoint.Y, midpoint2.X - referencePoint.X);

        return angle1.CompareTo(angle2);
    }
    private List<Curve> SortCurvesCounterClockwise(List<Curve> curves)
    {
        // Calculate the centroid of the curves to determine the sorting order
        XYZ centroid = GetCentroid(curves);

        // Sort the curves based on their angle relative to the centroid
        curves.Sort((c1, c2) => CompareCurvesByAngle(c1, c2, centroid));

        return curves;
    }


    //將Face邊緣轉換成PolyLine
    public PolyLine ConvertFaceEdgesToPolyLine(Face face)
    {
        List<Curve> curves = new List<Curve>();
        EdgeArrayArray edgeArrayArray = face.EdgeLoops;

        // 这里简单地选择第一个边界（Loop）
        EdgeArray edgeArray = edgeArrayArray.get_Item(0);

        CurveLoop curveLoop = new CurveLoop();

        foreach (Edge edge in edgeArray)
        {
            // 获取 Edge 的 Curve，并添加到 CurveLoop 中
            Curve curve = edge.AsCurve();
            curves.Add(curve);
        }

        curves = SortCurvesCounterClockwise(curves);

        foreach (Curve cc in curves)
        {
            MessageBox.Show("起點:" + cc.GetEndPoint(0).ToString() + "\n" +
                            "終點:" + cc.GetEndPoint(1).ToString());
        }

        foreach (Curve curve in curves)
        {
            curveLoop.Append(curve);
        }

        
        List<XYZ> points = new List<XYZ>();
        foreach (Curve curve in curveLoop)
        {
            points.Add(curve.GetEndPoint(0));
            points.Add(curve.GetEndPoint(1));
        }
        PolyLine polyline = PolyLine.Create(points);

        return polyline;
    }

    public CurveLoop GetCurveLoopFromFace(Face face)
    {
        List<Curve> curves = new List<Curve>();
        EdgeArrayArray edgeArrayArray = face.EdgeLoops;

        // 这里简单地选择第一个边界（Loop）
        EdgeArray edgeArray = edgeArrayArray.get_Item(0);

        CurveLoop curveLoop = new CurveLoop();

        foreach (Edge edge in edgeArray)
        {
            // 获取 Edge 的 Curve，并添加到 CurveLoop 中
            Curve curve = edge.AsCurve();
            curves.Add(curve);

        }

        curves = SortCurvesCounterClockwise(curves);
        foreach (Curve c in curves)
        {
            curveLoop.Append(c);
        }

        return curveLoop;
    }

    public List<CurveLoop> GetRecSlabsFromPolySlab(CurveLoop curveLoop, double gridline_size)
    {
        // Initiate some parameters.
        List<double> xCoorList = new List<double>();
        List<double> yCoorList = new List<double>();
        double zCoor = 0;

        foreach (Autodesk.Revit.DB.Curve curve in curveLoop.ToList())
        {
            xCoorList.Add(Modeling.Algorithm.RoundPoint(curve.GetEndPoint(0), gridline_size).X);
            xCoorList.Add(Modeling.Algorithm.RoundPoint(curve.GetEndPoint(1), gridline_size).X);
            yCoorList.Add(Modeling.Algorithm.RoundPoint(curve.GetEndPoint(0), gridline_size).Y);
            yCoorList.Add(Modeling.Algorithm.RoundPoint(curve.GetEndPoint(1), gridline_size).Y);
            zCoor = curve.GetEndPoint(0).Z;
        }

        // Sort the X and Y coordinate values.
        xCoorList.Sort();
        yCoorList.Sort();
        //string xCoordinatesString = "X Coordinates: " + string.Join(", ", xCoorList);
        //string yCoordinatesString = "Y Coordinates: " + string.Join(", ", yCoorList);
        //// Display the information in MessageBox.
        //MessageBox.Show(xCoordinatesString + Environment.NewLine + yCoordinatesString, "Coordinates Information");

        // Get the distinct values of sorted values.
        List<double> xSorted = GetDistinctList(xCoorList);
        List<double> ySorted = GetDistinctList(yCoorList);
        //string xSortedString = "X Sorted: " + string.Join(", ", xSorted);
        //string ySortedString = "Y Sorted: " + string.Join(", ", ySorted);
        //// Display the information in MessageBox.
        //MessageBox.Show(xSortedString + Environment.NewLine + ySortedString, "Coordinates Information");

        //foreach(double v in xSorted)
        //{
        //    MessageBox.Show(v.ToString());
        //}
        //foreach(double x in ySorted)
        //{
        //    MessageBox.Show(x.ToString());
        //}

        List<CurveLoop> profiles = new List<CurveLoop>();

        // Create small rectangles which consist of the polygon.
        List<List<XYZ>> recsPoints = new List<List<XYZ>>();
        for (int i = 0; i < xSorted.Count - 1; i++)
        {
            for (int j = 0; j < ySorted.Count - 1; j++)
            {
                CurveLoop profile = new CurveLoop();
                XYZ point1 = new XYZ(xSorted[i], ySorted[j], zCoor);
                XYZ point2 = new XYZ(xSorted[i + 1], ySorted[j], zCoor);
                XYZ point3 = new XYZ(xSorted[i + 1], ySorted[j + 1], zCoor);
                XYZ point4 = new XYZ(xSorted[i], ySorted[j + 1], zCoor);
                if (point1.DistanceTo(point2) < CentimetersToUnits(0.1)
                    || point2.DistanceTo(point3) < CentimetersToUnits(0.1)
                    || point3.DistanceTo(point4) < CentimetersToUnits(0.1)
                    || point4.DistanceTo(point1) < CentimetersToUnits(0.1))
                {
                    continue;
                }
                Autodesk.Revit.DB.Line line1 = Autodesk.Revit.DB.Line.CreateBound(point1, point2);
                Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(point2, point3);
                Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateBound(point3, point4);
                Autodesk.Revit.DB.Line line4 = Autodesk.Revit.DB.Line.CreateBound(point4, point1);

                profile.Append(line1);
                profile.Append(line2);
                profile.Append(line3);
                profile.Append(line4);

                // Using midpoint to chech whether the rectangle is in the outline or not.
                XYZ midPoint = (point1 + point3) / 2;
                if (IsInsideOutline(midPoint, curveLoop))
                {
                    profiles.Add(profile);
                }
            }
        }

        bool isAllSlabMerge = false;
        //recsPoints = SortPointsList(recsPoints);
        while (!isAllSlabMerge)
        {
            List<CurveLoop> curveloops = new List<CurveLoop>();
            int count = 0;
            for (int i = 0; i < profiles.Count; i++)
            {
                for (int j = 0; j < profiles.Count; j++)
                {
                    if (j <= i)
                    {
                        continue;
                    }
                    if (IsMergeable(profiles[i], profiles[j]))
                    {
                        curveloops.Add(MergeTwoRectangles(profiles[i], profiles[j]));
                        for (int k = 0; k < profiles.Count; k++)
                        {
                            if (k == i || k == j)
                            {
                                continue;
                            }
                            else
                            {
                                curveloops.Add(profiles[k]);
                            }
                        }
                        profiles = curveloops;
                        count = 1;
                        break;
                    }
                }
                if (count == 1)
                {
                    break;
                }
            }

            if (count == 0)
                isAllSlabMerge = true;
        }

        foreach (CurveLoop profile in profiles)
        {
            foreach (Curve c in profile)
            {
                MessageBox.Show("Startpoint1:" + c.GetEndPoint(0) + "\n" +
                                "Endpoint1:" + c.GetEndPoint(1));
            }
        }
        return profiles;
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
    public bool IsMergeable(CurveLoop a, CurveLoop b)
    {
        List<XYZ> list_a = new List<XYZ>();
        List<XYZ> list_b = new List<XYZ>();

        foreach (Autodesk.Revit.DB.Curve curve in a.ToList())
        {

            list_a.Add(curve.GetEndPoint(0));
        }

        foreach (Autodesk.Revit.DB.Curve curve in b.ToList())
        {
            list_b.Add(curve.GetEndPoint(0));
        }


        int count = 0;
        for (int i = 0; i < list_a.Count; i++)
        {
            for (int j = 0; j < list_b.Count; j++)
            {

                if (list_a[i].DistanceTo(list_b[j]) < CentimetersToUnits(0.1))
                {
                    count++;
                }
            }
        }

        if (count == 2)
        {
            return true;
        }
        else
            return false;
    }
    public CurveLoop MergeTwoRectangles(CurveLoop a, CurveLoop b)
    {
        List<XYZ> list_a = new List<XYZ>();
        List<XYZ> list_b = new List<XYZ>();

        foreach (Autodesk.Revit.DB.Curve curve in a.ToList())
        {
            list_a.Add(curve.GetEndPoint(0));
        }

        foreach (Autodesk.Revit.DB.Curve curve in b.ToList())
        {
            list_a.Add(curve.GetEndPoint(0));
        }

        List<XYZ> c = list_a.Concat(list_b).ToList();
        List<XYZ> d = new List<XYZ>();
        int count = 0;
        for (int i = 0; i < c.Count; i++)
        {
            for (int j = 0; j < c.Count; j++)
            {
                if (c[i].DistanceTo(c[j]) > CentimetersToUnits(0.1))
                {
                    count++;
                }
            }
            if (count == 7)
            {
                d.Add(c[i]);
            }
            count = 0;
        }

        return ListToCurveLooop(d);
    }
    public CurveLoop ListToCurveLooop(List<XYZ> c)
    {
        List<double> xCoor = new List<double>();
        List<double> yCoor = new List<double>();
        double zCoor = 0;
        foreach (XYZ point in c)
        {
            xCoor.Add(point.X);
            yCoor.Add(point.Y);
            zCoor = point.Z;
        }

        xCoor.Sort();
        yCoor.Sort();

        List<double> xSorted = GetDistinctList(xCoor);
        List<double> ySorted = GetDistinctList(yCoor);
        CurveLoop profileLoop = new CurveLoop();
        XYZ point1 = new XYZ(xSorted[0], ySorted[0], zCoor);
        XYZ point2 = new XYZ(xSorted[1], ySorted[0], zCoor);
        XYZ point3 = new XYZ(xSorted[1], ySorted[1], zCoor);
        XYZ point4 = new XYZ(xSorted[0], ySorted[1], zCoor);
        Autodesk.Revit.DB.Line line1 = Autodesk.Revit.DB.Line.CreateBound(point1, point2);
        Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(point2, point3);
        Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateBound(point3, point4);
        Autodesk.Revit.DB.Line line4 = Autodesk.Revit.DB.Line.CreateBound(point4, point1);
        profileLoop.Append(line1);
        profileLoop.Append(line2);
        profileLoop.Append(line3);
        profileLoop.Append(line4);
        return profileLoop;
    }
    public List<double> GetDistinctList(List<double> list)
    {
        List<double> list_new = new List<double>();
        for (int i = 0; i < list.Count - 1; ++i)
        {
            if (Math.Abs(list[i] - list[i + 1]) > CentimetersToUnits(0.1))
            {
                list_new.Add(list[i]);
            }
            if (i == list.Count - 2)
            {
                list_new.Add(list[i + 1]);
            }
        }
        return list_new;
    }

    public static double CentimetersToUnits(double value)
    {
        return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Centimeters);
    }

    public static double UnitsToCentimeters(double value)
    {
        return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
    }

    public XYZ XYZUnitsToCentimeters(XYZ point)
    {
        XYZ newPoint = new XYZ(
            UnitsToCentimeters(point.X),
            UnitsToCentimeters(point.Y),
            UnitsToCentimeters(point.Z)
            );
        return newPoint;
    }

    // 將平方英吋轉換為平方單位
    public double SquareUnitToSquareCentimeter(double squareInch)
    {
        double conversionFactor = 929.0304;
        return Math.Round(squareInch * conversionFactor);
    }

    public List<XYZ> RemoveDuplicatePoints(List<XYZ> points, double tolerance)
    {
        // 新建一个列表来保存唯一的点
        List<XYZ> uniquePoints = new List<XYZ>();

        // 遍历所有点
        foreach (XYZ point in points)
        {
            // 假设点是唯一的
            bool isDuplicate = false;

            // 与uniquePoints中已有的点进行距离比较
            foreach (XYZ uniquePoint in uniquePoints)
            {
                // 如果距离小于指定的容差，则认为这是一个重复的点
                if (point.DistanceTo(uniquePoint) < tolerance)
                {
                    isDuplicate = true;
                    break;
                }
            }

            // 如果点是唯一的，则添加到uniquePoints列表中
            if (!isDuplicate)
            {
                uniquePoints.Add(point);
            }
        }

        // 返回包含唯一点的列表
        return uniquePoints;
    }

    public Dictionary<double, List<XYZ>> GroupPointsByZ(List<XYZ> points) //將相同Z座標的點存到同一個List
    {
        Dictionary<double, List<XYZ>> groupedPoints = new Dictionary<double, List<XYZ>>();

        foreach (XYZ point in points)
        {
            double z = point.Z;

            if (!groupedPoints.ContainsKey(z))
            {
                groupedPoints[z] = new List<XYZ>();
            }

            groupedPoints[z].Add(point);
        }

        // 現在，pointDictionary 中的每個鍵值對應到一個 Z 座標，相應的值是具有相同 Z 座標的點的列表
        foreach (KeyValuePair<double, List<XYZ>> entry in groupedPoints)
        {
            double zCoordinate = entry.Key;
            List<XYZ> pointsWithSameZ = entry.Value;

            // 在這裡對 pointsWithSameZ 做任何你需要的處理
            MessageBox.Show($" Z 座標為 {UnitsToCentimeters(zCoordinate)} 的點數量: {pointsWithSameZ.Count}");
            foreach (XYZ p in pointsWithSameZ)
            {
                MessageBox.Show($"座標: {XYZUnitsToCentimeters(p)}");
            }
        }

        return groupedPoints;
    }


    public Dictionary<double, List<XYZ>> GroupPointsByY(List<XYZ> points) //將相同Y座標的點存到同一個List
    {
        Dictionary<double, List<XYZ>> groupedPoints = new Dictionary<double, List<XYZ>>();

        foreach (XYZ point in points)
        {
            double y = point.Y;

            if (!groupedPoints.ContainsKey(y))
            {
                groupedPoints[y] = new List<XYZ>();
            }

            groupedPoints[y].Add(point);
        }

        return groupedPoints;
    }


    public Dictionary<double, List<XYZ>> GroupPointsByX(List<XYZ> points) //將相同 X 座標的點存到同一個List
    {
        Dictionary<double, List<XYZ>> groupedPoints = new Dictionary<double, List<XYZ>>();

        foreach (XYZ point in points)
        {
            double x = point.X;

            if (!groupedPoints.ContainsKey(x))
            {
                groupedPoints[x] = new List<XYZ>();
            }

            groupedPoints[x].Add(point);
        }

        return groupedPoints;
    }


    public List<Line> xMaxAndMinPointsInList(List<XYZ> coordinateGroup) // 在相同 Y 座標裡的點找出 X 最大及最小的點連線
    {
        List<Line> horrizontalLines = new List<Line>();

        // 獲取 Y 相同的点的 X 最大和最小值
        if (coordinateGroup.Count > 1)
        {
            List<XYZ> arrangePoints_X = coordinateGroup.OrderBy(point => point.X).ToList();

            XYZ minXPoint = arrangePoints_X.First();
            XYZ maxXPoint = arrangePoints_X.Last();

            // 使用 LINQ 查询来获取 X 坐标最小的点
            minXPoint = arrangePoints_X.First(point => point.X == arrangePoints_X.Min(p => p.X));
            // 使用 LINQ 查询来获取 X 坐标最大的点
            maxXPoint = arrangePoints_X.First(point => point.X == arrangePoints_X.Max(p => p.X));


            Line yLine = Line.CreateBound(minXPoint, maxXPoint);
            horrizontalLines.Add(yLine);

        }

        MessageBox.Show("水平方向的線:" + horrizontalLines.Count.ToString());

        foreach (Line line in horrizontalLines)
        {
            MessageBox.Show(XYZUnitsToCentimeters(line.GetEndPoint(0)).ToString() + XYZUnitsToCentimeters(line.GetEndPoint(1)).ToString());
        }

        return horrizontalLines;
    }


    public List<Line> yMaxAndMinPointsInList(List<XYZ> coordinateGroup) // 在相同 X 座標裡的點找出 Y 最大及最小的點連線
    {
        List<Line> verticalLines = new List<Line>();

        // 獲取 X 相同的点的 Y 最大和最小值
        if (coordinateGroup.Count > 1)
        {
            List<XYZ> arrangePoints_Y = coordinateGroup.OrderBy(point => point.Y).ToList();

            XYZ minXPoint = arrangePoints_Y.First();
            XYZ maxXPoint = arrangePoints_Y.Last();

            // 使用 LINQ 查询来获取 X 坐标最小的点
            minXPoint = arrangePoints_Y.First(point => point.Y == arrangePoints_Y.Min(p => p.Y));
            // 使用 LINQ 查询来获取 X 坐标最大的点
            maxXPoint = arrangePoints_Y.First(point => point.Y == arrangePoints_Y.Max(p => p.Y));

            Line xLine = Line.CreateBound(minXPoint, maxXPoint);
            verticalLines.Add(xLine);
        }

        MessageBox.Show("垂直方向的線:" + verticalLines.Count.ToString());

        foreach (Line line in verticalLines)
        {
            MessageBox.Show(XYZUnitsToCentimeters(line.GetEndPoint(0)).ToString() + XYZUnitsToCentimeters(line.GetEndPoint(1)).ToString());
        }

        return verticalLines;
    }

    public (List<Line>, List<Line>) LandingBoundaryLines(Dictionary<double, List<XYZ>> pointDictionary)
    {
        List<Line> horrizontalLines = new List<Line>(); // 存 Y 相同的點的，其 X 最大和最小值相連的線
        List<Line> verticalLines = new List<Line>(); // 存 X 相同的點的，其 Y 最大和最小值相連的線

        foreach (KeyValuePair<double, List<XYZ>> z_kvp in pointDictionary)
        {
            double zCoordinate = z_kvp.Key; // Z 座標
            List<XYZ> z_coordinateGroup = z_kvp.Value; // 相同 Z 座標裡的點座標

            // 输出現在所選到的 Z 座標
            MessageBox.Show("Z 坐标: " + UnitsToCentimeters(zCoordinate));

            // 將相同Y、Z座標的點放進List中
            Dictionary<double, List<XYZ>> yCoordinatePointsDict = GroupPointsByY(z_coordinateGroup);

            foreach (KeyValuePair<double, List<XYZ>> y_kvp in yCoordinatePointsDict)
            {
                double yCoordinate = y_kvp.Key; // Y 座標
                List<XYZ> y_coordinateGroup = y_kvp.Value; // 相同 Y 座標裡的點座標

                // 输出現在所選到的 Y 座標
                MessageBox.Show("Y 坐标: " + UnitsToCentimeters(yCoordinate));
                MessageBox.Show("相同Y、Z座標的點共有: " + y_coordinateGroup.Count.ToString());

                //// 編歷在當前 Y 座標裡的座標點並輸出
                //foreach (XYZ point in y_coordinateGroup)
                //{
                //    MessageBox.Show("坐标点: (" + UnitsToCentimeters(point.X) + ", " + UnitsToCentimeters(point.Y) + ", " + UnitsToCentimeters(point.Z) + ")");
                //}

                horrizontalLines = xMaxAndMinPointsInList(y_coordinateGroup); // 在當前 Z 列表中 Y 座標相同的點的 X 最大值和最小值相連成線
            }

            // 將相同X、Z座標的點放進List中
            Dictionary<double, List<XYZ>> xCoordinatePointsDict = GroupPointsByX(z_coordinateGroup);
            foreach (KeyValuePair<double, List<XYZ>> x_kvp in xCoordinatePointsDict)
            {
                double xCoordinate = x_kvp.Key; // X 座標
                List<XYZ> x_coordinateGroup = x_kvp.Value; // 相同 X 座標裡的點座標

                // 输出現在所選到的 X 座標
                MessageBox.Show(" X 坐标: " + UnitsToCentimeters(xCoordinate));
                MessageBox.Show("相同、Z座標的點共有: " + x_coordinateGroup.Count.ToString());

                //// 編歷在當前 X 座標裡的座標點並輸出
                //foreach (XYZ point in x_coordinateGroup)
                //{
                //    MessageBox.Show("坐标点: (" + UnitsToCentimeters(point.X) + ", " + UnitsToCentimeters(point.Y) + ", " + UnitsToCentimeters(point.Z) + ")");
                //}

                verticalLines = yMaxAndMinPointsInList(x_coordinateGroup); // 在當前 Z 列表中 X 座標相同的點的 Y 最大值和最小值相連成線
            }
        }
        return (horrizontalLines, verticalLines);
    }
    private double GetFaceHeight(Face face)
    {
        // 获取面的 BoundingBox
        BoundingBoxUV boundingBoxUV = face.GetBoundingBox();
        double faceHeight = boundingBoxUV.Max.V - boundingBoxUV.Min.V;

        // 将高度从内部单位转换为可读的单位（例如，英尺）
        return faceHeight;
    }

}

