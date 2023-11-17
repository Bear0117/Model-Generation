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
//using static Autodesk.Revit.DB.SpecTypeId;

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
        public double 梯段路徑長度 { get; set; }
        public double 突沿長度 { get; set; }
        public double 平台長度 { get; set; }
        public double 平台寬度 { get; set; }
        public double 平台面積 { get; set; }
        public double 踏板面積 { get; set; }
        public double 立板面積 { get; set; }
        public double 路徑斜面積 { get; set; }
        public double 總面積 { get; set; }

        public BOMData()
        {
            名稱 = null;
            實際梯段寬度 = 0;
            實際級高 = 0;
            實際級深 = 0;
            實際梯級數 = 0;
            梯段路徑長度 = 0;
            突沿長度 = 0;

            平台長度 = 0;
            平台寬度 = 0;

            平台面積 = (平台長度 * 平台寬度) / 10000;
            踏板面積 = (實際梯段寬度 * 梯段路徑長度 * 2) / 10000;
            立板面積 = (實際梯段寬度 * 實際級高 * 實際梯級數) / 10000;
            路徑斜面積 = 實際梯段寬度 * (Math.Sqrt(Math.Pow(梯段路徑長度, 2) + Math.Pow(((實際級高 * 實際梯級數) / 2), 2))) / 1000;
            總面積 = 平台面積 + 踏板面積 + 立板面積 + 路徑斜面積;
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
                file.WriteLine($"{"5. 梯段路徑長度"},{item.梯段路徑長度}");
                file.WriteLine($"{"6. 平台長度"},{item.平台長度}");
                file.WriteLine($"{"7. 平台寬度"},{item.平台寬度},{"單位:cm"}");

                file.WriteLine();

                file.WriteLine($"{"面積計算"}");
                file.WriteLine($"{"1. 平台面積 = LL * LW "},{item.平台面積}");
                file.WriteLine($"{"2. 踏板面積 = RW * RPL * 2"},{item.踏板面積}");
                file.WriteLine($"{"3. 立板面積 = RW * RH * RN "},{item.立板面積}");
                file.WriteLine($"{"4. 路徑斜面積 = RW * Sqrt ((RPL ^ 2 + ((RH * RN) / 2) ^ 2 ) "},{item.路徑斜面積}");
                file.WriteLine($"{"5. 總面積 = 平台面積 + 踏板面積 + 立板面積 + 路徑斜面積 "},{item.總面積},{"單位:m^2"}");
                // 添加空行分隔每个 BOM 条目
            }

        }
    }
    public void Run(UIApplication uiapp)
    {
        UIDocument uidoc;
        uidoc = uiapp.ActiveUIDocument;
        Document doc = uidoc.Document;

        //所有所需參數
        string panelName = null;
        string panelCode = null;
        double stairRunsWidth = 0;
        double stairActualRiserHight = 0;
        double stairActualTreadDepth = 0;
        int stairActualNumRisers = 0;
        double stairPathLength = 0;
        double stairNosingLength = 0;

        double stairLandingLength = 0;
        double stairLandingWidth = 0;

        double landingArea = 0;
        double threadsArea = 0;
        double risersArea = 0;
        double pathArea = 0;
        double totalArea = 0;



        //點選樓梯
        Reference paramerterRef = uidoc.Selection.PickObject(ObjectType.Element);
        ElementId selectedElementId = paramerterRef.ElementId;
        Element selectedElement = doc.GetElement(selectedElementId);
        string salectedCategoryName = selectedElement.Category.Name;

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

            //須更新
            double stairPathhypotenuseLength = Math.Sqrt(Math.Pow(stairPathLength, 2) + Math.Pow((stairActualRiserHight * stairActualNumRisers) / 2, 2));

            ICollection<ElementId> stairLandingsICollectionId = stair.GetStairsLandings();
            MessageBox.Show("平台:" + stairLandingsICollectionId.Count.ToString());
            ICollection<ElementId> stairRunsICollectionId = stair.GetStairsRuns();
            MessageBox.Show("梯段:" + stairRunsICollectionId.Count.ToString());

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
                    stairPathLength = Convert.ToInt32(UnitsToCentimeters(stairPathCurve.Length).ToString());
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

            //平台的參數
            foreach (Element stairLanding in stairLandingsList)
            {
                // 獲取樓梯平台的Geometry對象
                GeometryElement geometryElement = stairLanding.get_Geometry(new Options());

                List<XYZ> points = new List<XYZ>(); //平台所有的點
                //List<Line> linesList_X = new List<Line>();
                //List<Line> linesList_Y = new List<Line>();

                List<Line> yLines = new List<Line>(); // 存 Y 相同的點的，其 X 最大和最小值相連的線
                List<Line> xLines = new List<Line>(); // 存 X 相同的點的，其 Y 最大和最小值相連的線


                // 編歷Geometry對象以獲取尺寸信息
                foreach (GeometryObject geomObj in geometryElement)
                {
                    // 抓取在實體上的幾何參數
                    if (geomObj is Solid solid)
                    {
                        // 抓取幾何面
                        FaceArray faces = solid.Faces;
                        MessageBox.Show("平台面數量:" + faces.Size.ToString());
                        foreach (Face face in faces)
                        {
                            // MessageBox.Show("face");

                            // 幾何面的邊緣
                            EdgeArrayArray egdearrayarray = face.EdgeLoops;
                            foreach (EdgeArray edgearray in egdearrayarray)
                            {
                                // 將邊緣線的點存到points裡
                                foreach (Edge edge in edgearray)
                                {
                                    Curve curve = edge.AsCurve();
                                    points.Add(Modeling.Algorithm.RoundPoint(curve.GetEndPoint(0), 5));
                                    points.Add(Modeling.Algorithm.RoundPoint(curve.GetEndPoint(1), 5));
                                }
                            }
                        }
                        // 將points裡相同的點移除
                        points = RemoveDuplicatePoints(points, 0.001);
                        MessageBox.Show("在平台中不重複的點數量:" + points.Count.ToString());
                        //foreach (XYZ point in points)
                        //{
                        //    MessageBox.Show("點座標:" + point.ToString());
                        //}

                        // 創建一个字典，用於儲存每個 Z 坐座標對應的點列表
                        Dictionary<double, List<XYZ>> zCoordinatePointsDict = GroupPointsByZ(points);

                        foreach (var kvp in zCoordinatePointsDict)
                        {
                            double zCoordinate = kvp.Key; // Z 座標
                            List<XYZ> coordinateGroup = kvp.Value; // 相同 Z 座標裡的點座標

                            // 输出現在所選到的 Z 座標
                            MessageBox.Show("Z 坐标: " + zCoordinate);

                            //// 編歷在當前 Z 座標裡的座標點並輸出
                            //foreach (XYZ point in coordinateGroup)
                            //{
                            //    MessageBox.Show("坐标点: (" + point.X + ", " + point.Y + ", " + point.Z + ")");
                            //}

                            yLines = yMaxAndMinPointsInList(coordinateGroup); // 在當前 Z 列表中 Y 座標相同的點的 X 最大值和最小值相連成線
                            xLines = xMaxAndMinPointsInList(coordinateGroup); // 在當前 Z 列表中 X 座標相同的點的 Y 最大值和最小值相連成線

                        }
                    }
                }
                break;
            }




            landingArea = (double)(stairLandingWidth * stairLandingLength) / 10000;
            threadsArea = (double)(stairRunsWidth * stairPathLength * 2) / 10000;
            risersArea = (double)(stairRunsWidth * stairActualRiserHight * stairActualNumRisers) / 10000;
            pathArea = (double)(stairRunsWidth * stairPathhypotenuseLength * 2) / 10000;
            totalArea = (double)(landingArea + threadsArea + risersArea + pathArea);

            MessageBox.Show(panelName + "\n" + panelCode + "\n" +
                stairActualRiserHight.ToString() + "\n" + stairActualTreadDepth.ToString() + "\n" + stairNosingLength.ToString() + "\n" + stairActualNumRisers.ToString() + "\n" + stairPathhypotenuseLength.ToString() + "\n" +
                landingArea.ToString() + "\n" + threadsArea.ToString() + "\n" + risersArea.ToString() + "\n" + pathArea.ToString() + "\n" + totalArea.ToString());
        }

        //BOM表
        List<BOMData> BOMDatas = new List<BOMData>();
        Dictionary<string, BOMData> keyValuePairs = new Dictionary<string, BOMData>();


        //foreach (Element stair in stairList)
        //{
        //    Element panelElement1 = doc.GetElement(stair.Id);
        //    FamilyInstance familyInstance = panelElement1 as FamilyInstance;
        //    ElementId pickedtypeid = panelElement1.GetTypeId();
        //    ElementType family = doc.GetElement(pickedtypeid) as ElementType;

        //    panelName = panelElement1.Name.ToString();
        //    panelCode = family.FamilyName.ToString();

        //    stairActualRiserHight = (double)UnitUtils.Convert(stair.LookupParameter("實際級高").AsDouble(), UnitTypeId.Feet, UnitTypeId.Centimeters);
        //    stairActualTreadDepth = (double)UnitUtils.Convert(stair.LookupParameter("實際級深").AsDouble(), UnitTypeId.Feet, UnitTypeId.Centimeters);
        //    string stairActualNumRisersst = stair.LookupParameter("實際梯級數").AsValueString();
        //    stairActualNumRisers = Convert.ToInt32(stairActualNumRisersst);
        //    double stairPathhypotenuseLength = Math.Sqrt(Math.Pow(stairPathLength, 2) + Math.Pow((stairActualRiserHight * stairActualNumRisers) / 2, 2));


        //    //StairsRunType stairsRunType = doc.GetElement(stair.GetTypeId()) as StairsRunType;
        //    //stairNosingLength = UnitsToCentimeters(stairsRunType.NosingLength);

        //    //double stairNosingLength = (double)UnitUtils.Convert(stair.get_Parameter(BuiltInParameter.STAIRS_TRISERTYPE_NOSING_LENGTH).AsDouble(), UnitTypeId.Feet, UnitTypeId.Centimeters);

        //    landingArea = (double)(stairLandingWidth * stairLandingLength) / 10000;
        //    threadsArea = (double)(stairRunsWidth * stairPathLength * 2) / 10000;
        //    risersArea = (double)(stairRunsWidth * stairActualRiserHight * stairActualNumRisers) / 10000;
        //    pathArea = (double)(stairRunsWidth * stairPathhypotenuseLength * 2) / 10000;
        //    totalArea = (double)(landingArea+ threadsArea+ risersArea+ pathArea);

        //    MessageBox.Show(panelName + "\n" + panelCode + "\n" +
        //        stairActualRiserHight.ToString() + "\n" + stairActualTreadDepth.ToString() + "\n" + stairNosingLength.ToString() + "\n" + stairActualNumRisers.ToString() + "\n" + stairPathhypotenuseLength.ToString() + "\n" +
        //        landingArea.ToString() + "\n" + threadsArea.ToString() + "\n" + risersArea.ToString() + "\n" + pathArea.ToString() + "\n" + totalArea.ToString());
        //}

        BOMData bOMData = new BOMData()
        {
            名稱 = panelName,
            實際梯段寬度 = stairRunsWidth,
            實際級高 = stairActualRiserHight,
            實際級深 = stairActualTreadDepth,
            實際梯級數 = stairActualNumRisers,
            梯段路徑長度 = stairPathLength,
            突沿長度 = stairNosingLength,

            平台長度 = stairLandingLength,
            平台寬度 = stairLandingWidth,

            平台面積 = landingArea,
            踏板面積 = threadsArea,
            立板面積 = risersArea,
            路徑斜面積 = pathArea,
            總面積 = totalArea
        };
        keyValuePairs[panelName] = bOMData;


        foreach (BOMData bOMData1 in keyValuePairs.Values)
        {
            BOMDatas.Add(bOMData1);
        }

        WriteToCSV(BOMDatas);

    }
    public static double UnitsToCentimeters(double value)
    {
        return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
    }

    public List<XYZ> RemoveDuplicatePoints(List<XYZ> points, double tolerance)
    {
        // 新建一个列表来保存唯一的点
        List<XYZ> uniquePoints = new List<XYZ>();

        // 遍历所有点
        foreach (var point in points)
        {
            // 假设点是唯一的
            bool isDuplicate = false;

            // 与uniquePoints中已有的点进行距离比较
            foreach (var uniquePoint in uniquePoints)
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

    public Dictionary<double, List<XYZ>> GroupPointsByZ(List<XYZ> points)
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

        return groupedPoints;


        //Dictionary<double, List<XYZ>> zCoordinatePointsDict = new Dictionary<double, List<XYZ>>();

        //// 編歷 points 列表，根据 Z 座標將點分组儲存到字典中
        //foreach (XYZ point in points)
        //{
        //    double z = point.Z;
        //    if (!zCoordinatePointsDict.ContainsKey(z))
        //    {
        //        zCoordinatePointsDict[z] = new List<XYZ>();
        //    }
        //    zCoordinatePointsDict[z].Add(point);
        //}
    }

    public List<Line> yMaxAndMinPointsInList(List<XYZ> coordinateGroup)
    {
        // 按 Y 坐标值对点进行分组
        var groupedPoints = coordinateGroup.GroupBy(point => point.Y);

        List<Line> yLines = new List<Line>();

        foreach (var group in groupedPoints)
        {
            List<XYZ> pointsWithSameY = group.ToList();

            // 连接相同 Y 坐标的点成线
            for (int i = 0; i < pointsWithSameY.Count - 1; i++)
            {
                XYZ startPoint = pointsWithSameY[i];
                XYZ endPoint = pointsWithSameY[i + 1];
                Line yLine = Line.CreateBound(startPoint, endPoint);
                yLines.Add(yLine);
            }
        }

        MessageBox.Show("水平方向的線:" + yLines.Count.ToString());

        foreach (Line line in yLines)
        {
            MessageBox.Show("(" + line.GetEndPoint(0).ToString() + ")" + "(" + line.GetEndPoint(1).ToString() + ")");
        }

        return yLines;

        //// 將當前分组的座標點，根據 Y 和 X 座標值分別儲存
        //List<XYZ> ySortedPoints = coordinateGroup.OrderBy(point => point.Y).ToList(); // Y 相同的點
        //List<Line> yLines = new List<Line>(); // 存 Y 相同的點的，其 X 最大和最小值相連的線
        //MessageBox.Show("in");
        //// 獲取 Y 相同的点的 X 最大和最小值
        //if (ySortedPoints.Count > 1)
        //{
        //    XYZ minXPoint = ySortedPoints.First();
        //    XYZ maxXPoint = ySortedPoints.Last();

        //    // 使用 LINQ 查询来获取 X 坐标最小的点
        //    minXPoint = ySortedPoints.First(point => point.X == ySortedPoints.Min(p => p.X));
        //    // 使用 LINQ 查询来获取 X 坐标最大的点
        //    maxXPoint = ySortedPoints.First(point => point.X == ySortedPoints.Max(p => p.X));


        //    Line yLine = Line.CreateBound(minXPoint, maxXPoint);
        //    yLines.Add(yLine);
        //}
        //MessageBox.Show("水平方向的線:" + yLines.Count.ToString());

        //foreach (Line line in yLines)
        //{
        //    MessageBox.Show("(" + line.GetEndPoint(0).ToString() + ")" + "(" + line.GetEndPoint(1).ToString() + ")");
        //}

        //return yLines;

    }


    public List<Line> xMaxAndMinPointsInList(List<XYZ> coordinateGroup)
    {
        // 按 Y 坐标值对点进行分组
        var groupedPoints = coordinateGroup.GroupBy(point => point.X);

        List<Line> xLines = new List<Line>();

        foreach (var group in groupedPoints)
        {
            List<XYZ> pointsWithSameX = group.ToList();

            // 连接相同 Y 坐标的点成线
            for (int i = 0; i < pointsWithSameX.Count - 1; i++)
            {
                XYZ startPoint = pointsWithSameX[i];
                XYZ endPoint = pointsWithSameX[i + 1];
                Line xLine = Line.CreateBound(startPoint, endPoint);
                xLines.Add(xLine);
            }
        }

        MessageBox.Show("水平方向的線:" + xLines.Count.ToString());

        foreach (Line line in xLines)
        {
            MessageBox.Show("(" + line.GetEndPoint(0).ToString() + ")" + "(" + line.GetEndPoint(1).ToString() + ")");
        }

        return xLines;




        //// 將當前分组的座標點，根據 Y 和 X 座標值分別儲存
        //List<XYZ> xSortedPoints = coordinateGroup.OrderBy(point => point.X).ToList(); // X 相同的點
        //List<Line> xLines = new List<Line>(); // 存 X 相同的點的，其 Y 最大和最小值相連的線

        //// 获取 X 相同的点的 Y 最大和最小值
        //if (xSortedPoints.Count > 1)
        //{
        //    XYZ minYPoint = xSortedPoints.First();
        //    XYZ maxYPoint = xSortedPoints.Last();

        //    // 使用 LINQ 查询来获取 X 坐标最小的点
        //    minYPoint = xSortedPoints.First(point => point.Y == xSortedPoints.Min(p => p.Y));
        //    // 使用 LINQ 查询来获取 X 坐标最大的点
        //    maxYPoint = xSortedPoints.First(point => point.Y == xSortedPoints.Max(p => p.Y));

        //    Line xLine = Line.CreateBound(minYPoint, maxYPoint);
        //    xLines.Add(xLine);
        //}
        //MessageBox.Show("垂直方向的線:" + xLines.Count.ToString());

        //foreach (Line line in xLines)
        //{
        //    MessageBox.Show(line.GetEndPoint(0).ToString() + line.GetEndPoint(1).ToString());
        //}

        //return xLines;

    }








}