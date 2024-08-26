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
    class StairLandingDimensions_FloorPlane : IExternalCommand
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

                double stairBaseElevation = selectedStair.BaseElevation;

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
                    //Element sta = stairLandingsList[0];
                    //StairsLanding stairLanding = sta as StairsLanding;
                    List<Face> stairLandingFace = GetFace(stairLanding);

                    List<Line> stairLandingedge = GetLine(stairLandingFace);
                    stairLandingedge = MergeLines(stairLandingedge, 0.00001);

                    Transaction tran = new Transaction(doc, "Create Dimension");
                    tran.Start();

                    // 創建一個平面
                    double landingHeight = stairLanding.BaseElevation + stairBaseElevation;
                    Plane plane = Plane.CreateByNormalAndOrigin(uidoc.ActiveView.ViewDirection, new XYZ(0, 0, landingHeight));
                    SketchPlane sketchPlane = SketchPlane.Create(doc, plane);

                    foreach (Line line in stairLandingedge)
                    {
                        Plane sketchPlanePlane = sketchPlane.GetPlane();
                        double parameter1 = sketchPlanePlane.Normal.DotProduct(line.GetEndPoint(0) - sketchPlanePlane.Origin);
                        double parameter2 = sketchPlanePlane.Normal.DotProduct(line.GetEndPoint(1) - sketchPlanePlane.Origin);
                        if (Math.Abs(parameter1) < 0.0001 && Math.Abs(parameter2) < 0.0001)
                        {
                            // 在 Family Editor 中創建一個 ModelCurve
                            ModelCurve modelCurve = doc.Create.NewModelCurve(line, sketchPlane);

                            ReferenceArray referenceArray = new ReferenceArray();
                            // MessageBox.Show(modelCurve.GeometryCurve.GetEndPointReference(0) + "\n" + modelCurve.GeometryCurve.GetEndPointReference(1));

                            referenceArray.Append(modelCurve.GeometryCurve.GetEndPointReference(0));
                            referenceArray.Append(modelCurve.GeometryCurve.GetEndPointReference(1));
                            // MessageBox.Show("參考數量:" + referenceArray.Size.ToString());
                            doc.Create.NewDimension(uidoc.ActiveView, line, referenceArray);
                            
                            // MessageBox.Show("Right");
                        }
                        else
                        {
                            MessageBox.Show(line.GetEndPoint(0) + "\n" + line.GetEndPoint(1) + "\n" + landingHeight);
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
       
        private List<Face> GetFace(Element element)
        {
            GeometryElement geometryElement = element.get_Geometry(new Options());
            List<Face> returnFace = new List<Face>();

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

                        if (Math.Abs(dotProduct - 1.0) < 1e-9 /*|| Math.Abs(dotProduct + 1.0) < 1e-9*/)
                        {
                            // MessageBox.Show("in");
                            returnFace.Add(face);
                        }
                    }
                }
            }
            // MessageBox.Show("頂部平面數量:" + returnFace.Count.ToString());
            return returnFace;
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

            //StringBuilder coordinatesStringBuilder = new StringBuilder();
            //foreach (Line line in lineList)
            //{
            //    XYZ startPoint = line.GetEndPoint(0);
            //    XYZ endPoint = line.GetEndPoint(1);

            //    // 將座標添加到 StringBuilder
            //    coordinatesStringBuilder.AppendLine($"SP: {startPoint}, \n EP: {endPoint}");
            //}
            //// 顯示所有座標在 TaskDialog 中
            //TaskDialog.Show("Line Coordinates", coordinatesStringBuilder.ToString());

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

    }
}
