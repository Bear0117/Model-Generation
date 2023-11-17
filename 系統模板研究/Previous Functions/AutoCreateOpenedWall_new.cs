using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;
using Line = Autodesk.Revit.DB.Line;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateOpenedWall_new : IExternalCommand
    {
        Autodesk.Revit.ApplicationServices.Application app;
        Document doc;
        string layername = null;

        //取得cad的路徑位置
        public string GetCADPath(ElementId cadLinkTypeID, Document revitDoc)
        {
            CADLinkType cadLinkType = revitDoc.GetElement(cadLinkTypeID) as CADLinkType;
            return ModelPathUtils.ConvertModelPathToUserVisiblePath(cadLinkType.GetExternalFileReference().GetAbsolutePath());
        }

        //執行檔案
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;


            //抓取正確位置的文字
            Reference reference = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            Element elem2 = doc.GetElement(reference);
            GeometryElement geoElem = elem2.get_Geometry(new Options());
            GeometryObject geoObj = elem2.GetGeometryObjectFromReference(reference);


            string levelN = "1F";
            List<BuiltInParameter> list = GetBuiltInParametersByElement(elem2);

            foreach (BuiltInParameter param in list)
            {
                if (param == BuiltInParameter.IMPORT_BASE_LEVEL)
                {
                    Parameter para = elem2.get_Parameter(param);
                    ElementId id = para.AsElementId();
                    levelN = doc.GetElement(id).Name;
                }
            }
            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals(levelN)) as Level;

            //建立CAD連結路徑
            string path = GetCADPath(elem2.GetTypeId(), doc);

            //抓取圖層名稱
            Category targetCategory1 = null;
            ElementId graphicsStyleId1 = null;

            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                graphicsStyleId1 = geoObj.GraphicsStyleId;
                GraphicsStyle gs1 = doc.GetElement(geoObj.GraphicsStyleId) as GraphicsStyle;
                if (gs1 != null)
                {
                    targetCategory1 = gs1.GraphicsStyleCategory;
                    layername = gs1.GraphicsStyleCategory.Name;

                    TaskDialog.Show("Revit", layername);
                }
            }

            //Cad內部線級轉換
            List<CADModel> cadcurveArray = GetCurveArray(geoElem, graphicsStyleId1);
            List<Autodesk.Revit.DB.Curve> curves = new List<Autodesk.Revit.DB.Curve>();

            foreach (CADModel a in cadcurveArray)
            {
                foreach (Autodesk.Revit.DB.Curve curve in a.CurveArray)
                {
                    curves.Add(curve);
                }
            }

            //從Cad內抓出了幾條線

            //TaskDialog.Show("revit", " 總共抓出了" + curves.Count.ToString() + "條線");

            //從Cad內抓出多少文字

            List<CADTextModel> text = GetCADTextInfoparing(path, 0);

            //TaskDialog.Show("revit", " 總共抓出了" + text.Count.ToString() + "組文字");


            //foreach(CADTextModel a in text)
            //{
            //    TaskDialog.Show("revit",a.HText.ToString());
            //    TaskDialog.Show("revit", a.WText.ToString());

            //}

            //將cad內部的線分組成block
            List<List<Autodesk.Revit.DB.Curve>> doorBlocks = new List<List<Autodesk.Revit.DB.Curve>>();
            List<List<Autodesk.Revit.DB.Curve>> Clusters = Algorithm.ClusterByParallel(curves);


            //  牆最後抓取出的中線
            List<Wallmodel> openingwall = new List<Wallmodel>();

            //建立族群用的寬度類
            List<double> widthtype = new List<double>();


            //將cad內組成的block做調整 抓出isbound的線級
            foreach (List<Autodesk.Revit.DB.Curve> clustergroup in Clusters)
            {
                int count = 0;
                if (Algorithm.CreateBoundingBox2D(clustergroup) != null)
                {
                    List<Wallmodel> comparewall = new List<Wallmodel>();
                    List<Line> compareline = new List<Line>();
                    List<Autodesk.Revit.DB.Curve> finalcluster = Algorithm.CreateBoundingBox2D(clustergroup);

                    double firstwidth = 0;

                    //用平行判斷cluster內的中線
                    for (int i = 0; i < finalcluster.Count; i++)
                    {
                        for (int j = i + 1; j < finalcluster.Count; j++)
                        {
                            if (Algorithm.IsParallel(finalcluster[i], finalcluster[j]))
                            {
                                Wallmodel axses = new Wallmodel
                                {
                                    Wallaxes = Algorithm.GenerateAxis(finalcluster[i] as Line, finalcluster[j] as Line),
                                    Width = GetWallWidth(finalcluster[i] as Line, finalcluster[j] as Line)
                                };
                                axses.Midpoint = GetMiddlePoint(axses.Wallaxes.GetEndPoint(0), axses.Wallaxes.GetEndPoint(1));
                                //MessageBox.Show(axses.midpoint.ToString());
                                if (axses.Width != firstwidth)
                                {
                                    widthtype.Add(axses.Width);
                                    firstwidth = axses.Width;
                                }

                                comparewall.Add(axses);
                                compareline.Add(axses.Wallaxes as Line);



                            }
                        }
                    }


                    // 抓取出最後的中線
                    int finalnumber = 0;
                    if (compareline[1].Length > compareline[0].Length)
                    {

                        finalnumber = 1;
                    }

                    openingwall.Add(comparewall[finalnumber]);

                }
            }

            //TaskDialog.Show("revit", openingwall.Count.ToString());



            foreach (Wallmodel wall in openingwall)
            {
                double comparedistance = 0;
                double distanceBetween = 10000000;

                if (text.Count >= 1)
                {
                    foreach (CADTextModel walltext in text)
                    {
                        XYZ mid_new = new XYZ(wall.Midpoint.X, wall.Midpoint.Y, 0) * 0.1;
                        comparedistance = mid_new.DistanceTo(walltext.Location);
                        //comparedistance = wall.midpoint.DistanceTo(walltext.Location * 0.003281);
                        //MessageBox.Show(mid_new.ToString() + "and" + (walltext.Location).ToString());
                        //break;
                        if (comparedistance < distanceBetween)
                        {
                            distanceBetween = comparedistance;
                            wall.HText = walltext.HText;
                            wall.WText = walltext.WText;
                            //MessageBox.Show(distanceBetween.ToString());
                            //MessageBox.Show(wall.WText.ToString());
                        }
                        //MessageBox.Show(distanceBetween.ToString());
                    }
                }
                //break;
            }

            //在模型裡面繪製中軸線以檢視
            //Autodesk.Revit.DB.Transaction t3 = new Autodesk.Revit.DB.Transaction(doc, "drawing line");

            //t3.Start();
            //foreach (wallmodel walls in openingwall)
            //{

            //    if (walls.Wallaxes.IsBound)
            //    {
            //        SketchPlane modelSketch = SketchPlane.Create(doc, Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero));//這裡要改一下原點座標，要改成General一點
            //        ModelCurve modelLine = doc.Create.NewModelCurve(walls.Wallaxes, modelSketch);
            //    }
            //}

            //t3.Commit();

            //牆的高度有待調整
            double wallHeight = MillimetersToUnits(3300);

            //創建一道新的牆面
            foreach (Wallmodel openwall in openingwall)
            {
                //MessageBox.Show(openwall.width.ToString());
                if (Math.Abs(openwall.Width - 209) < 1 || Math.Abs(openwall.Width - 358) < 1)
                {
                    //MessageBox.Show("1");
                    continue;
                }
                //MessageBox.Show("Create!");
                //先為牆建立一種族群
                ElementId id = CreatWallType(doc, openwall.Width.ToString(), MillimetersToUnits(openwall.Width));

                //建立牆
                Autodesk.Revit.DB.Transaction t1 = new Autodesk.Revit.DB.Transaction(doc, "Create Wall");
                t1.Start();

                Wall wall = Wall.Create(doc, openwall.Wallaxes, id, level.Id, wallHeight, 0, true, true);
                WallUtils.DisallowWallJoinAtEnd(wall, 0);
                WallUtils.DisallowWallJoinAtEnd(wall, 1);

                t1.Commit();

                List<double> depthList = new List<double>();

                //Get the element by elementID and get the boundingbox and outline of this element.
                Element elementWall = wall as Element;
                BoundingBoxXYZ bbxyzElement = elementWall.get_BoundingBox(null);
                Outline outline = new Outline(bbxyzElement.Min, bbxyzElement.Max);

                //Create a filter to get all the intersection elements with wall.
                BoundingBoxIntersectsFilter filterW = new BoundingBoxIntersectsFilter(outline);

                //Create a filter to get StructuralFraming (which include beam and column) and Slabs.
                ElementCategoryFilter filterSt = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);
                ElementCategoryFilter filterSl = new ElementCategoryFilter(BuiltInCategory.OST_Floors);
                LogicalOrFilter filterS = new LogicalOrFilter(filterSt, filterSl);

                //Combine two filter.
                LogicalAndFilter filter = new LogicalAndFilter(filterS, filterW);

                //A list to store the intersected elements.
                IList<Element> inter = new FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements();
                //MessageBox.Show(inter.Count.ToString());

                for (int i = 0; i < inter.Count; i++)
                {
                    if (inter[i] != null)
                    {
                        //MessageBox.Show(inter[i].ToString());
                        string elementName = inter[i].Category.Name;
                        if (elementName == "結構構架")
                        {
                            //MessageBox.Show("結構構架");
                            BoundingBoxXYZ bbxyzBeam = inter[i].get_BoundingBox(null);
                            depthList.Add(Math.Abs(bbxyzBeam.Min.Z - bbxyzBeam.Max.Z));
                            //MessageBox.Show(bbxyzBeam.Min.ToString());
                            //MessageBox.Show(bbxyzBeam.Max.ToString());
                        }
                        else if (elementName == "樓板")
                        {
                            //MessageBox.Show("樓板");
                            BoundingBoxXYZ bbxyzSlab = inter[i].get_BoundingBox(null);
                            depthList.Add(Math.Abs(bbxyzSlab.Min.Z - bbxyzSlab.Max.Z));
                        }
                        else continue;
                    }

                }


                double h = Convert.ToDouble(openwall.HText);
                double fl = Convert.ToDouble(openwall.WText);
                //開口牆建立
                Autodesk.Revit.DB.Transaction t3 = new Autodesk.Revit.DB.Transaction(doc, "Create Wall");
                t3.Start();
                //CreateOpening(doc, wall, MillimetersToUnits(h), MillimetersToUnits(fl));
                //MessageBox.Show(h.ToString() + "+" +fl.ToString());
                CreateOpening(doc, wall, h, fl);
                t3.Commit();

                Autodesk.Revit.DB.Transaction t2 = new Autodesk.Revit.DB.Transaction(doc, "Adjust Wall Height");
                t2.Start();
                //MessageBox.Show(depthList.Count.ToString());
                //MessageBox.Show(depthList.Max().ToString());
                Parameter wallHeightP = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                if (depthList.Count() != 0)
                {
                    wallHeightP.Set(CentimetersToUnits(330) - depthList.Min());
                }
                else
                {
                    wallHeightP.Set(CentimetersToUnits(330));
                }

                t2.Commit();




                //TaskDialog.Show("reivt", h.ToString());
                //TaskDialog.Show("reivt", fl.ToString());



                //break;
            }

            return Result.Succeeded;
        }


        //建立一個牆的類型(從既有的類型庫複製性質後建立)
        public static ElementId CreatWallType(Document doc, string wallTypeName, double width)
        {

            ElementId wallTypeId = null;
            FilteredElementCollector Col = new FilteredElementCollector(doc);
            List<Element> familySymbolList = Col.OfClass(typeof(WallType)).ToList();
            WallType baseWallType = null;
            WallType newWallType = null;
            using (Autodesk.Revit.DB.Transaction transaction = new Autodesk.Revit.DB.Transaction(doc))
            {
                if (transaction.Start("創建新牆族群") == TransactionStatus.Started)
                {

                    foreach (WallType item in familySymbolList)
                    {
                        if (item.Name == wallTypeName)
                        {
                            //TaskDialog.Show("建立墙", "牆的類型已經存在");
                            transaction.Commit();
                            return item.Id;
                        }
                    }

                    try
                    {
                        //獲得一個牆的類性並重新命名
                        foreach (WallType item in familySymbolList)
                        {
                            if (item.Name == "Generic - 200mm")
                            {
                                baseWallType = item;
                                break;
                            }
                        }

                        newWallType = baseWallType.Duplicate(wallTypeName) as WallType;
                        doc.Regenerate();
                        //改變厚度

                        CompoundStructure wallTypeStructure = newWallType.GetCompoundStructure();
                        double wallThickness = wallTypeStructure.GetWidth();//得到厚度
                        int startIndex = wallTypeStructure.GetFirstCoreLayerIndex();


                        wallTypeStructure.SetLayerWidth(startIndex, width);

                        newWallType.SetCompoundStructure(wallTypeStructure);//修改后設定

                        if (TransactionStatus.Committed != transaction.Commit())
                        {
                            TaskDialog.Show("創建新墙類型", "提交失败！");
                        }
                    }
                    catch
                    {
                        transaction.RollBack();
                        throw;
                    }
                }
            }
            wallTypeId = newWallType.Id;
            return wallTypeId;
        }

        //取得線級組件
        private List<CADModel> GetCurveArray(GeometryElement geoElem, ElementId graphicsStyleId)
        {
            List<CADModel> curveArray_List = new List<CADModel>();
            // 判斷元素類型
            foreach (GeometryObject gObj in geoElem)
            {
                GeometryInstance geomInstance = gObj as GeometryInstance;
                // 座標轉換
                Transform transform = geomInstance.Transform;
                if (null != geomInstance)
                {
                    foreach (GeometryObject insObj in geomInstance.SymbolGeometry) // 取得幾何類別
                    {
                        if (insObj.GraphicsStyleId.IntegerValue != graphicsStyleId.IntegerValue)
                            continue;

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.Line")
                        {
                            Autodesk.Revit.DB.Line line = insObj as Autodesk.Revit.DB.Line;
                            //XYZ normal = XYZ.BasisZ;
                            XYZ point = line.GetEndPoint(0);
                            point = transform.OfPoint(point);

                            Autodesk.Revit.DB.Line newLine = TransformLine(transform, line);

                            CurveArray curveArray = new CurveArray();
                            curveArray.Append(newLine);

                            XYZ startPoint = newLine.GetEndPoint(0);
                            XYZ endPoint = newLine.GetEndPoint(1);
                            XYZ MiddlePoint = GetMiddlePoint(startPoint, endPoint);
                            double angle = (startPoint.Y - endPoint.Y) / startPoint.DistanceTo(endPoint);
                            double rotation = Math.Asin(angle);

                            CADModel cADModel = new CADModel
                            {
                                CurveArray = curveArray,
                                Length = newLine.Length,
                                Shape = "牆",
                                Width = 300 / 304.8,
                                Location = MiddlePoint,
                                Rotation = rotation
                            };
                            curveArray_List.Add(cADModel);

                        }

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")//对于连续的折线
                        {

                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> points = polyLine.GetCoordinates();


                            for (int i = 0; i < points.Count - 1; i++)
                            {
                                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(points[i], points[i + 1]);
                                line = TransformLine(transform, line);
                                Autodesk.Revit.DB.Line newLine = line;
                                CurveArray curveArray = new CurveArray();
                                curveArray.Append(newLine);

                                XYZ startPoint = newLine.GetEndPoint(0);
                                XYZ endPoint = newLine.GetEndPoint(1);
                                XYZ MiddlePoint = GetMiddlePoint(startPoint, endPoint);
                                double angle = (startPoint.Y - endPoint.Y) / startPoint.DistanceTo(endPoint);
                                double rotation = Math.Asin(angle);

                                CADModel cADModel = new CADModel
                                {
                                    CurveArray = curveArray,
                                    Length = newLine.Length,
                                    Shape = "牆",
                                    Width = 300 / 304.8,
                                    Location = MiddlePoint,
                                    Rotation = rotation
                                };
                                curveArray_List.Add(cADModel);

                            }
                        }
                    }
                }
            }



            return curveArray_List;
        }

        //取得文字組件
        public List<CADTextModel> GetCADTextInfoparing(string dwgFile, double underbeam)
        {
            List<CADTextModel> CADModels = new List<CADTextModel>();
            List<CADText> HTEXT = new List<CADText>();
            List<CADText> FLTEXT = new List<CADText>();
            List<ObjectId> allObjectId = new List<ObjectId>();

            using (new Services())
            {
                using (Database database = new Database(false, false))
                {
                    database.ReadDwgFile(dwgFile, FileShare.Read, true, "");
                    using (Teigha.DatabaseServices.Transaction trans = database.TransactionManager.StartTransaction())
                    {
                        using (BlockTable table = (BlockTable)database.BlockTableId.GetObject(OpenMode.ForRead))
                        {
                            using (SymbolTableEnumerator enumerator = table.GetEnumerator())
                            {
                                while (enumerator.MoveNext())
                                {
                                    using (BlockTableRecord record = (BlockTableRecord)enumerator.Current.GetObject(OpenMode.ForRead))
                                    {
                                        foreach (ObjectId id in record)
                                        {
                                            Entity entity = (Entity)id.GetObject(OpenMode.ForRead, false, false);
                                            CADText fltext = new CADText();
                                            CADText htext = new CADText();

                                            if (entity.Layer == layername)
                                            {
                                                if (entity.GetRXClass().Name == "AcDbText")
                                                {
                                                    Teigha.DatabaseServices.DBText text = (Teigha.DatabaseServices.DBText)entity;
                                                    if (text.TextString.Contains("h"))
                                                    {
                                                        htext.Text = Regex.Replace(text.TextString, "[^0-9]", "");
                                                        htext.Location = ConverCADPointToRevitPoint(text.Position);
                                                        HTEXT.Add(htext);
                                                    }

                                                    else if (text.TextString.Contains("FL"))
                                                    {
                                                        if (text.TextString.Contains("樑下"))
                                                        {
                                                            fltext.Text = underbeam.ToString();
                                                            fltext.Location = ConverCADPointToRevitPoint(text.Position);
                                                            FLTEXT.Add(fltext);
                                                        }
                                                        else
                                                        {
                                                            fltext.Text = Regex.Replace(text.TextString, "[^0-9]", "");
                                                            fltext.Location = ConverCADPointToRevitPoint(text.Position);
                                                            FLTEXT.Add(fltext);
                                                        }
                                                    }
                                                }

                                                if (entity.GetRXClass().Name == "AcDbMText")
                                                {
                                                    Teigha.DatabaseServices.MText mText = (Teigha.DatabaseServices.MText)entity;

                                                    if (mText.Text.Contains("h"))
                                                    {
                                                        htext.Text = Regex.Replace(mText.Text, "[^0-9]", "");
                                                        htext.Location = ConverCADPointToRevitPoint(mText.Location);
                                                        HTEXT.Add(htext);
                                                    }

                                                    else if (mText.Text.Contains("FL"))
                                                    {
                                                        if (mText.Text.Contains("樑下"))
                                                        {
                                                            fltext.Text = underbeam.ToString();
                                                            fltext.Location = ConverCADPointToRevitPoint(mText.Location);
                                                            FLTEXT.Add(fltext);

                                                        }

                                                        else
                                                        {
                                                            fltext.Text = Regex.Replace(mText.Text, "[^0-9]", "");
                                                            fltext.Location = ConverCADPointToRevitPoint(mText.Location);
                                                            FLTEXT.Add(fltext);
                                                        }
                                                    }
                                                }
                                            }
                                            //TaskDialog.Show("revit", entity.Layer);
                                            //break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            //文字抓取檢查用

            //int count1 = 0;
            //foreach (CADText height in HTEXT)
            //{
            //    count1++;
            //    TaskDialog.Show("REVIT", "第" + count1.ToString() + "項為" + height.Text);
            //}
            //int count = 0;
            //foreach (CADText height in FLTEXT)
            //{
            //    count++;
            //    TaskDialog.Show("REVIT", "第" + count.ToString() + "項為" + height.Text);
            //}



            //文字分組配對
            foreach (CADText height in HTEXT)
            {
                //存放配對組合
                CADTextModel comparemodel = new CADTextModel();
                double comparedistance = 0;
                double distanceBetween = 1000000000000000000;

                foreach (CADText floorheight in FLTEXT)
                {

                    comparedistance = height.Location.DistanceTo(floorheight.Location);
                    if (comparedistance < distanceBetween)
                    {
                        distanceBetween = comparedistance;
                        comparemodel.HText = height.Text;
                        comparemodel.WText = floorheight.Text;
                        comparemodel.Location = MidPoint(height.Location, floorheight.Location);
                    }

                }

                CADModels.Add(comparemodel);

            }

            //文字抓取檢查用

            //int count2 = 0;
            //foreach ( CADTextModel text  in CADModels)
            //{
            //    count2++;
            //    TaskDialog.Show("REVIT", "第" + count1.ToString() + "項為" + text.HText);
            //    TaskDialog.Show("REVIT", "第" + count1.ToString() + "項為" + text.WText);
            //}


            return CADModels;
        }




        //函式庫 數學運算
        private XYZ GetMiddlePoint(XYZ startPoint, XYZ endPoint)
        {
            XYZ MiddlePoint = new XYZ((startPoint.X + endPoint.X) / 2, (startPoint.Y + endPoint.Y) / 2, (startPoint.Z + endPoint.Z) / 2);
            return MiddlePoint;
        }
        public XYZ Cross(XYZ a, XYZ b)
        {

            XYZ c = new XYZ(a.Y * b.Z - a.Z * b.Y, a.Z * b.X - a.X * b.Z, a.X * b.Y - a.Y * b.X);
            return c;
        }
        private double Dot(XYZ a, XYZ b)
        {

            double c = a.X * b.X + a.Y * b.Y;
            return c;

        }
        public XYZ MidPoint(XYZ a, XYZ b)
        {
            XYZ c = (a + b) / 2;
            return c;
        }
        double MillimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Millimeters);
        }




        //函式庫 幾何轉換&運算
        private Line TransformLine(Transform transform, Line line)
        {
            XYZ startPoint = transform.OfPoint(line.GetEndPoint(0));
            XYZ endPoint = transform.OfPoint(line.GetEndPoint(1));
            Line newLine = Line.CreateBound(startPoint, endPoint);
            return newLine;
        }
        int GetWallWidth(Autodesk.Revit.DB.Line line1, Autodesk.Revit.DB.Line line2)
        {
            XYZ ponit1_1 = line1.GetEndPoint(0);
            XYZ ponit1_2 = line1.GetEndPoint(1);
            XYZ ponit2_1 = line2.GetEndPoint(0);
            XYZ ponit2_2 = line2.GetEndPoint(1);
            line1.MakeUnbound();
            double width = line1.Distance(ponit2_2);
            int wallWidth = (int)Math.Round(width / 0.003281);

            return wallWidth;
        }
        public static XYZ ConverCADPointToRevitPoint(Point3d point)
        {
            double MillimetersToUnits(double value)
            {
                return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Millimeters);
            }

            return new XYZ(MillimetersToUnits(point.X), MillimetersToUnits(point.Y), MillimetersToUnits(point.Z));


        }
        private void CreateOpening(Document RevitDoc, Wall wall, double h, double fl)
        {
            LocationCurve locationCurve = wall.Location as LocationCurve;
            Autodesk.Revit.DB.Line location = locationCurve.Curve as Autodesk.Revit.DB.Line;
            XYZ startPoint = location.GetEndPoint(0);
            XYZ endPoint = location.GetEndPoint(1);

            //Reference refwall = new Reference(wall);


            //FamilySymbol windowType = null;
            //foreach (FamilySymbol symbol in new FilteredElementCollector(doc)
            //    .OfClass(typeof(FamilySymbol))
            //    .OfCategory(BuiltInCategory.OST_Windows))
            //{
            //    if (symbol.Name == "0406 x 0610mm")
            //    {
            //        windowType = symbol;
            //        break;
            //    }
            //}


            // The bottom is less than 15cm means it is kerb.
            if (fl < 15)
            {
                XYZ bottomCoor = new XYZ(0, 0, 0);
                XYZ topCoor = new XYZ(0, 0, CentimetersToUnits(fl + h));
                Opening opening = RevitDoc.Create.NewOpening(wall, startPoint + bottomCoor, endPoint + topCoor);
            }
            else
            {
                XYZ bottomCoor = new XYZ(0, 0, CentimetersToUnits(fl));
                XYZ topCoor = new XYZ(0, 0, CentimetersToUnits(fl + h));
                Opening opening = RevitDoc.Create.NewOpening(wall, startPoint + bottomCoor, endPoint + topCoor);
            }



        }

        public double CentimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Centimeters);
        }


        //函式庫 組件定義 
        public class CADTextModel
        {
            private string Htext;

            private string Wtext;

            private XYZ location;

            private double distant;


            public string HText
            {
                get
                {
                    return Htext;
                }

                set
                {
                    Htext = value;
                }
            }

            public string WText
            {
                get
                {
                    return Wtext;
                }

                set
                {
                    Wtext = value;
                }
            }

            public double Distant
            {
                get
                {
                    return distant;
                }

                set
                {
                    distant = value;
                }
            }

            public XYZ Location
            {
                get
                {
                    return location;
                }

                set
                {
                    location = value;
                }
            }


        }
        public class CADText
        {
            private string context;

            private XYZ textlocation;

            public string Text
            {
                get
                {
                    return context;
                }

                set
                {
                    context = value;
                }
            }

            public XYZ Location
            {
                get
                {
                    return textlocation;
                }

                set
                {
                    textlocation = value;
                }
            }
        }


        static List<BuiltInParameter> GetBuiltInParametersByElement(Element element)
        {
            List<BuiltInParameter> bips = new List<BuiltInParameter>();

            foreach (BuiltInParameter bip in BuiltInParameter.GetValues(typeof(BuiltInParameter)))
            {
                Parameter p = element.get_Parameter(bip);

                if (p != null)
                {
                    bips.Add(bip);
                }
            }
            return bips;
        }


        public class Wallmodel
        {

            private string Htext;

            private string Wtext;
            public double Width { get; set; }
            public XYZ Midpoint { get; set; }

            private Autodesk.Revit.DB.Curve wallaxes;

            public Autodesk.Revit.DB.Curve Wallaxes
            {
                get
                {
                    return wallaxes;
                }

                set
                {
                    wallaxes = value;
                }
            }

            public string HText
            {
                get
                {
                    return Htext;
                }

                set
                {
                    Htext = value;
                }
            }

            public string WText
            {
                get
                {
                    return Wtext;
                }

                set
                {
                    Wtext = value;
                }
            }

        }

    }
}
