using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class Sample : IExternalCommand
    {
        Autodesk.Revit.ApplicationServices.Application app;
        Document doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            app = commandData.Application.Application;
            doc = uidoc.Document;

            Reference r = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            string ss = r.ConvertToStableRepresentation(doc);

            Element elem = doc.GetElement(r);
            GeometryElement geoElem = elem.get_Geometry(new Options());
            GeometryObject geoObj = elem.GetGeometryObjectFromReference(r);

            //獲取選中的cad圖層
            Category targetCategory = null;
            ElementId graphicsStyleId = null;

            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                graphicsStyleId = geoObj.GraphicsStyleId;
                GraphicsStyle gs = doc.GetElement(geoObj.GraphicsStyleId) as GraphicsStyle;
                if (gs != null)
                {
                    targetCategory = gs.GraphicsStyleCategory;
                    var name = gs.GraphicsStyleCategory.Name;
                }
            }

            double NormWallWidth = 360 ;//unit：mm　（1 mm = 0.00328 ft）

            List<CADModel> curveArray_List = getCurveArray(doc, geoElem, graphicsStyleId);
            List<CADModel> curveArray_List_copy = new List<CADModel>();

            // 复制得到的模型
            foreach (CADModel OrginCADModle in curveArray_List)
            {
                curveArray_List_copy.Add(OrginCADModle);
            }

            int LineNumber = curveArray_List.Count;//取得的模型的线的总数量

            List<CADModel> NotMatchCadModel = new List<CADModel>();//存放不匹配的梁的相关线      
            List<List<CADModel>> CADModelList_List = new List<List<CADModel>>();//存放模型数组的数组

            //筛选模型
            while (curveArray_List.Count > 0)
            {
                //存放距离
                List<double> distanceList = new List<double>();
                //存放对应距离的CADModel
                List<CADModel> cADModel_B_List = new List<CADModel>();

                CADModel CadModel_A = curveArray_List[0];
                curveArray_List.Remove(CadModel_A);//去除取出的梁的二段线段之一

                if (curveArray_List.Count >= 1)
                {
                    foreach (CADModel CadModel_B in curveArray_List)
                    {
                        //梁的2个段线非同一长度最大误差为50mm，方向为绝对值（然而sin120°=sin60°）
                        if ((float)Math.Abs(CadModel_A.Rotation) == (float)Math.Abs(CadModel_B.Rotation) 
                            && Math.Abs(CadModel_A.Length - CadModel_B.Length) * 304.87 < 12321231)
                        {
                            //TaskDialog.Show("線條資訊", "兩線等長", TaskDialogCommonButtons.Ok);
                            double distance = CadModel_A.Location.DistanceTo(CadModel_B.Location);
                            distanceList.Add(distance);
                            cADModel_B_List.Add(CadModel_B);
                        }
                    }


                    if (distanceList.Count != 0 && cADModel_B_List.Count != 0)
                    {
                        double distanceTwoLine = distanceList.Min();
                        //筛选不正常的宽度,如发现不正常，将CadModel_B继续放入数组
                        if (distanceTwoLine * 304.87 < NormWallWidth && distanceTwoLine > 0)
                        {
                            TaskDialog.Show("1", (distanceTwoLine * 304.8).ToString());

                            CADModel CadModel_shortDistance = cADModel_B_List[distanceList.IndexOf(distanceTwoLine)];
                            curveArray_List.Remove(CadModel_shortDistance);
                            //1对梁的模型装入数组
                            List<CADModel> cADModels = new List<CADModel>();
                            cADModels.Add(CadModel_A);
                            cADModels.Add(CadModel_shortDistance);
                            CADModelList_List.Add(cADModels);
                            //TaskDialog.Show("1", CadModel_A.location.ToString() + "\n" + CadModel_shortDistance.location.ToString());

                        }
                    }
                    else
                    {
                        NotMatchCadModel.Add(CadModel_A);
                    }

                }
                else
                {
                    NotMatchCadModel.Add(CadModel_A);
                }

            }
            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("1F")) as Level;
            int tranNumber = 0;//用于改变事务的ID
                               //生成梁
            foreach (var cadModelList in CADModelList_List)
            {
                CADModel cADModel_A = cadModelList[0];
                CADModel cADModel_B = cadModelList[1];

                //TaskDialog.Show("1", cADModel_A.location.ToString() + "\n" + cADModel_B.location.ToString());

                XYZ cADModel_A_StratPoint = cADModel_A.CurveArray.get_Item(0).GetEndPoint(0);
                XYZ cADModel_A_EndPoint = cADModel_A.CurveArray.get_Item(0).GetEndPoint(1);
                XYZ cADModel_B_StratPoint = cADModel_B.CurveArray.get_Item(0).GetEndPoint(0);
                XYZ cADModel_B_EndPoint = cADModel_B.CurveArray.get_Item(0).GetEndPoint(1);

                XYZ ChangeXYZ = new XYZ();

                double LineLength = (GetMiddlePoint(cADModel_A_StratPoint, cADModel_B_StratPoint)).DistanceTo(GetMiddlePoint(cADModel_A_EndPoint, cADModel_B_EndPoint));
                if (LineLength < 0.00328)//梁的2段线起点非同一端。2段线非同一长度，又非同一端的，中间点的误差选择为1mm
                {
                    //TaskDialog.Show("線條資訊", "兩線等長", TaskDialogCommonButtons.Ok);
                    ChangeXYZ = cADModel_B_StratPoint;
                    cADModel_B_StratPoint = cADModel_B_EndPoint;
                    cADModel_B_EndPoint = ChangeXYZ;
                }
                //if ( ! (GetMiddlePoint(cADModel_A_StratPoint, cADModel_B_StratPoint)).Equals(GetMiddlePoint(cADModel_A_EndPoint, cADModel_B_EndPoint)))
                {
                    //TaskDialog.Show("線條資訊", "兩線等長", TaskDialogCommonButtons.Ok);
                    Curve curve = Line.CreateBound((GetMiddlePoint(cADModel_A_StratPoint, cADModel_B_StratPoint)), GetMiddlePoint(cADModel_A_EndPoint, cADModel_B_EndPoint));
                    //TaskDialog.Show("線條資訊", "線條種類：" + curve.Length, TaskDialogCommonButtons.Ok);
                    Transaction t1 = new Transaction(doc, "創建");
                    t1.Start();
                    Wall wall = Wall.Create(doc, curve, level.Id, true);

                    t1.Commit();
                }

                


                //double distance = cADModel_A.location.DistanceTo(cADModel_B.location);

                //distance = Math.Round(distance * 304.8, 1);//作为梁_b的参数
                //WidthList.Add(distance);//梁宽度集合

                //string beamName = "ZBIM矩形梁 " + (float)(distance) + "*" + (float)(600) + "mm";//类型名 宽度*高度
                //if (!familSymbol_exists(beamName, "ZBIM - 矩形梁", doc))
                //{
                //    MakeBeamType(beamName, "ZBIM - 矩形梁");
                //    EditBeamType(beamName, (float)(distance), (float)(600));
                //}

                //用于数据显示和选择，已注释
                #region
                //List<string> columnTypes = new List<string>();
                //columnTypes = getBeamTypes(doc);
                //bool repeat = false;
                //foreach (string context in columnTypes)
                //{
                //    if (context == beamName)
                //    {
                //        repeat = true;
                //        break;
                //    }
                //}
                //if (!repeat)
                //{
                //    columnTypes.Add(beamName);
                //}
                #endregion

                //using (Transaction transaction = new Transaction(doc))
                //{
                //    transaction.Start("Beadm Strart Bulid" + tranNumber.ToString());
                //    FilteredElementCollector collector = new FilteredElementCollector(doc);
                //    collector.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralFraming);


                //    foreach (FamilySymbol beamType in collector)
                //    {
                //        if (beamType.Name == beamName)
                //        {
                //            if (!beamType.IsActive)
                //            {
                //                beamType.Activate();
                //            }
                //            FamilyInstance beamInstance = doc.Create.NewFamilyInstance(curve, beamType, level, StructuralType.Beam);
                //            var Elevation = beamInstance.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);

                //            break;
                //        }
                //    }



                //    transaction.Commit();
                //}


            }


            return Result.Succeeded;
        }



        private XYZ GetMiddlePoint(XYZ startPoint, XYZ endPoint)
        {
            XYZ MiddlePoint = new XYZ((startPoint.X + endPoint.X) / 2, (startPoint.Y + endPoint.Y) / 2, (startPoint.Z + endPoint.Z) / 2);
            return MiddlePoint;
        }
        private Line TransformLine(Transform transform, Line line)
        {
            XYZ startPoint = transform.OfPoint(line.GetEndPoint(0));
            XYZ endPoint = transform.OfPoint(line.GetEndPoint(1));
            Line newLine = Line.CreateBound(startPoint, endPoint);
            return newLine;
        }
        private List<CADModel> getCurveArray(Document doc, GeometryElement geoElem, ElementId graphicsStyleId)
        {
            List<CADModel> curveArray_List = new List<CADModel>();
            TransactionGroup transGroup = new TransactionGroup(doc, "绘制模型线");
            transGroup.Start();
            //判断元素类型
            foreach (GeometryObject gObj in geoElem)
            {
                GeometryInstance geomInstance = gObj as GeometryInstance;
                //坐标转换。如果选择的是“自动-中心到中心”，或者移动了importInstance，需要进行坐标转换
                Transform transform = geomInstance.Transform;
                if (null != geomInstance)
                {
                    foreach (GeometryObject insObj in geomInstance.SymbolGeometry)//取几何得类别
                    {
                        if (insObj.GraphicsStyleId.IntegerValue != graphicsStyleId.IntegerValue)
                            continue;
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.NurbSpline")
                        {//不需要
                        }
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.Line")
                        {
                            //TaskDialog.Show("線條資訊", "兩線等長", TaskDialogCommonButtons.Ok);
                            Line line = insObj as Line;
                            XYZ normal = XYZ.BasisZ;
                            XYZ point = line.GetEndPoint(0);
                            point = transform.OfPoint(point);

                            Line newLine = TransformLine(transform, line);

                            CurveArray curveArray = new CurveArray();
                            curveArray.Append(TransformLine(transform, line));

                            XYZ startPoint = newLine.GetEndPoint(0);
                            XYZ endPoint = newLine.GetEndPoint(1);
                            XYZ MiddlePoint = GetMiddlePoint(startPoint, endPoint);
                            double angle = (startPoint.Y - endPoint.Y) / startPoint.DistanceTo(endPoint);
                            double rotation = Math.Asin(angle);

                            CADModel cADModel = new CADModel();
                            cADModel.CurveArray = curveArray;
                            cADModel.Length = newLine.Length;
                            cADModel.Shape = "矩形梁";
                            cADModel.Width = 300 / 304.8;
                            cADModel.Location = MiddlePoint;
                            cADModel.Rotation = rotation;

                            curveArray_List.Add(cADModel);
                        }
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.Arc")
                        {//不需要
                        }
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")//对于连续的折线
                        {

                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> points = polyLine.GetCoordinates();


                            for (int i = 0; i < points.Count - 1; i++)
                            {
                                if (points[i] != points[i + 1])
                                {
                                    //TaskDialog.Show("線條資訊", "兩線等長", TaskDialogCommonButtons.Ok);
                                    Line line = Line.CreateBound(points[i], points[i + 1]);
                                    line = TransformLine(transform, line);
                                    Line newLine = line;
                                    CurveArray curveArray = new CurveArray();
                                    curveArray.Append(newLine);

                                    XYZ startPoint = newLine.GetEndPoint(0);
                                    XYZ endPoint = newLine.GetEndPoint(1);
                                    XYZ MiddlePoint = GetMiddlePoint(startPoint, endPoint);
                                    double angle = (startPoint.Y - endPoint.Y) / startPoint.DistanceTo(endPoint);
                                    double rotation = Math.Asin(angle);

                                    CADModel cADModel = new CADModel();
                                    cADModel.CurveArray = curveArray;
                                    cADModel.Length = newLine.Length;
                                    cADModel.Shape = "矩形梁";
                                    cADModel.Width = 300 / 304.8;
                                    cADModel.Location = MiddlePoint;
                                    cADModel.Rotation = rotation;

                                    curveArray_List.Add(cADModel);
                                }
                            }
                        }
                    }
                }
            }
            transGroup.Assimilate();
            return curveArray_List;
        }
    }
}
