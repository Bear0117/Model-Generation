using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Media;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;
using Line = Autodesk.Revit.DB.Line;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateOpenedWall : IExternalCommand
    {
        // Initialize Layer Name.
        string layername = null;

        // Get CAD path.
        public string GetCADPath(Document revitDoc, ElementId cadLinkTypeID)
        {
            CADLinkType cadLinkType = revitDoc.GetElement(cadLinkTypeID) as CADLinkType;
            return ModelPathUtils.ConvertModelPathToUserVisiblePath(cadLinkType.GetExternalFileReference().GetAbsolutePath());
        }

        // Main code.
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            //Call the WinForm.
            UIApplication uiApp = commandData.Application;
            Document doc_ui = uiApp.ActiveUIDocument.Document;

            string level_Name = "1F";
            double kerbHeight;
            OpenWallForm form = new OpenWallForm(doc_ui);
            form.ShowDialog();
            level_Name = form.levelName;
            kerbHeight = form.KerbHeight;
            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals(level_Name)) as Level;

            // Select the Opened-wall layer.
            Reference reference = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            Element elem2 = doc.GetElement(reference);
            GeometryElement geoElem = elem2.get_Geometry(new Options());
            GeometryObject geoObj = elem2.GetGeometryObjectFromReference(reference);

            // Link the CAD drawing.
            string path = GetCADPath(doc, elem2.GetTypeId());

            // Grab the name of the layer.
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
                }
            }

            // Get all the curve of Opened-wall layer.
            List<CADModel> cadCurveArray = GetCurveArray(geoElem, graphicsStyleId1);
            List<Autodesk.Revit.DB.Curve> curves = new List<Autodesk.Revit.DB.Curve>();

            foreach (CADModel a in cadCurveArray)
            {
                foreach (Autodesk.Revit.DB.Curve curve in a.CurveArray)
                {
                    curves.Add(curve);
                }
            }

            List<CADTextModel> text = GetCADTextInfoparing(path);

            // Pair the lines which compose opened-walls.
            List<List<Autodesk.Revit.DB.Curve>> clusters = Algorithm.ClusterByParallel(curves);

            //  This List is to store the center line of walls.
            List<Wallmodel> openingwall = new List<Wallmodel>();

            // This List is to store wall widths.
            List<double> widthtype = new List<double>();


            foreach (List<Autodesk.Revit.DB.Curve> clusterGroup in clusters)
            {
                if (Algorithm.CreateBoundingBox2D(clusterGroup) != null)
                {
                    List<Wallmodel> compareWall = new List<Wallmodel>();
                    List<Line> compareLine = new List<Line>();
                    List<Autodesk.Revit.DB.Curve> finalcluster = Algorithm.CreateBoundingBox2D(clusterGroup);

                    double firstwidth = 0;

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
                                if (axses.Width != firstwidth)
                                {
                                    widthtype.Add(axses.Width);
                                    firstwidth = axses.Width;
                                }

                                compareWall.Add(axses);
                                compareLine.Add(axses.Wallaxes as Line);
                            }
                        }
                    }


                    // Get the middle line of the wall.
                    int finalnumber = 0;
                    if (compareLine[1].Length > compareLine[0].Length)
                    {
                        finalnumber = 1;
                    }
                    openingwall.Add(compareWall[finalnumber]);
                }
            }

            // Pair the wall with specific text by distance.
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
                        if (comparedistance < distanceBetween)
                        {
                            distanceBetween = comparedistance;
                            wall.HText = walltext.HText;
                            wall.WText = walltext.WText;
                        }
                    }
                }
            }

            // Transform the wall line coordinate.
            //foreach (Wallmodel walls in openingwall)
            //{
            //    if (walls.Wallaxes.IsBound)
            //    {
            //        GeometryInstance geomInstance = geoObj as GeometryInstance;
            //        Autodesk.Revit.DB.Transform transform = geomInstance.Transform;
            //        XYZ point = walls.Wallaxes.GetEndPoint(0);
            //        point = transform.OfPoint(point);
            //        walls.Wallaxes = TransformLine(transform, walls.Wallaxes as Line);


            //        //SketchPlane modelSketch = SketchPlane.Create(doc, Autodesk.Revit.DB.Plane.CreateByNormalAndOrigin(XYZ.BasisZ, XYZ.Zero));
            //        //ModelCurve modelLine = doc.Create.NewModelCurve(walls.Wallaxes, modelSketch);
            //    }
            //}

            // To get wall height.
            double levelHeight = CentimetersToUnits(330); // default
            List<double> levelDistancesList = new List<double>();

            // retrieve all the Level elements in the document
            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(Level));
            IEnumerable<Level> levels = collector.Cast<Level>();

            // loop through all the levels
            foreach (Level level_1 in levels)
            {
                // get the level elevation.
                double levelElevation = level_1.Elevation;
                if (Math.Abs(levelElevation - level.Elevation) < CentimetersToUnits(0.1))
                {
                    level = level_1;
                }
                if (levelElevation - level.Elevation > CentimetersToUnits(0.1))
                {
                    levelDistancesList.Add(levelElevation);
                }
            }

            if (levelDistancesList.Count() > 0)
            {
                levelHeight = levelDistancesList.Min() - level.Elevation;
            }
            else
            {
                LevelHeightForm form1 = new LevelHeightForm(doc);
                form1.ShowDialog();
                levelHeight = CentimetersToUnits(form1.levelHeight);
            }

            // Create walls.
            foreach (Wallmodel openwall in openingwall)
            {
                // To avoid create wall with two lines in label circles. 
                if (Math.Abs(openwall.Width - 209) < 1 || Math.Abs(openwall.Width - 358) < 1)
                {
                    continue;
                }

                ElementId id = CreatWallType(doc, openwall.Width.ToString(), MillimetersToUnits(openwall.Width));

                Autodesk.Revit.DB.Transaction t1 = new Autodesk.Revit.DB.Transaction(doc, "Create Wall");
                t1.Start();
                Wall wall = Wall.Create(doc, openwall.Wallaxes, id, level.Id, levelHeight, 0, true, true);
                WallUtils.DisallowWallJoinAtEnd(wall, 0);
                WallUtils.DisallowWallJoinAtEnd(wall, 1);
                t1.Commit();

                List<double> depthList = new List<double>();

                //Get the element by elementID and get the boundingbox and outline of this element.
                Element elementWall = wall as Element;
                BoundingBoxXYZ bbxyzElement = elementWall.get_BoundingBox(null);
                Outline outline = new Outline(bbxyzElement.Min, bbxyzElement.Max);
                Outline newOutline = ReducedOutline(bbxyzElement);

                //Create a filter to get all the intersection elements with wall.
                BoundingBoxIntersectsFilter filterW = new BoundingBoxIntersectsFilter(newOutline);

                //Create a filter to get StructuralFraming (which include beam and column) and Slabs.
                ElementCategoryFilter filterSt = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);
                ElementCategoryFilter filterSl = new ElementCategoryFilter(BuiltInCategory.OST_Floors);
                LogicalOrFilter filterS = new LogicalOrFilter(filterSt, filterSl);

                //Combine two filter.
                LogicalAndFilter filter = new LogicalAndFilter(filterS, filterW);

                //A list to store the intersected elements.
                IList<Element> inter = new FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements();

                for (int i = 0; i < inter.Count; i++)
                {
                    if (inter[i] != null)
                    {
                        //MessageBox.Show(inter[i].ToString());
                        string elementName = inter[i].Category.Name;
                        if (elementName == "結構構架")
                        {
                            BoundingBoxXYZ bbxyzBeam = inter[i].get_BoundingBox(null);
                            depthList.Add(Math.Abs(bbxyzBeam.Min.Z - bbxyzBeam.Max.Z));
                        }
                        else if (elementName == "樓板")
                        {
                            BoundingBoxXYZ bbxyzSlab = inter[i].get_BoundingBox(null);
                            depthList.Add(Math.Abs(bbxyzSlab.Min.Z - bbxyzSlab.Max.Z));
                        }
                        else continue;
                    }
                }

                // Create Opened-wall
                double h = Convert.ToDouble(openwall.HText);
                double fl = Convert.ToDouble(openwall.WText);
                CreateWindow(doc, wall, h, fl, kerbHeight);

                // Adjust wall height.
                Autodesk.Revit.DB.Transaction t2 = new Autodesk.Revit.DB.Transaction(doc, "Adjust Wall Height");
                t2.Start();
                Parameter wallHeightP = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                if (depthList.Count() != 0)
                {
                    wallHeightP.Set(levelHeight - depthList.Max());
                }
                else
                {
                    wallHeightP.Set(levelHeight);
                }
                t2.Commit();
            }
            return Result.Succeeded;
        } // The end of main code. 

        // Create a Walltype by duplicating exist Walltype in Revit Project.
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

        // Get all the CurveArray
        private List<CADModel> GetCurveArray(GeometryElement geoElem, ElementId graphicsStyleId)
        {
            List<CADModel> curveArray_List = new List<CADModel>();
            // 判斷元素類型
            foreach (GeometryObject gObj in geoElem)
            {
                GeometryInstance geomInstance = gObj as GeometryInstance;
                // 座標轉換
                Autodesk.Revit.DB.Transform transform = geomInstance.Transform;
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

        // Get all the text in the CAD drawing.
        public List<CADTextModel> GetCADTextInfoparing(string dwgFile)
        {
            List<CADTextModel> CADModels = new List<CADTextModel>();
            List<CADText> HTEXT = new List<CADText>();
            List<CADText> FLTEXT = new List<CADText>();

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
                                                        fltext.Text = Regex.Replace(text.TextString, "[^0-9]", "");
                                                        fltext.Location = ConverCADPointToRevitPoint(text.Position);
                                                        FLTEXT.Add(fltext);
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
                                                        fltext.Text = Regex.Replace(mText.Text, "[^0-9]", "");
                                                        fltext.Location = ConverCADPointToRevitPoint(mText.Location);
                                                        FLTEXT.Add(fltext);
                                                    }
                                                }
                                            }
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




        // Functions : Mathematical Operation.
        private XYZ GetMiddlePoint(XYZ startPoint, XYZ endPoint)
        {
            XYZ MiddlePoint = new XYZ((startPoint.X + endPoint.X) / 2, (startPoint.Y + endPoint.Y) / 2, (startPoint.Z + endPoint.Z) / 2);
            return MiddlePoint;
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

        // Function : Transformation
        public double CentimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Centimeters);
        }

        public double UnitsToCentimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
        }

        private Line TransformLine(Autodesk.Revit.DB.Transform transform, Line line)
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

        public Outline ReducedOutline(BoundingBoxXYZ bbxyz)
        {
            XYZ diagnalVector = (bbxyz.Max - bbxyz.Min) / 1000;
            Outline newOutline = new Outline(bbxyz.Min + diagnalVector, bbxyz.Max - diagnalVector);
            return newOutline;
        }

        public static XYZ ConverCADPointToRevitPoint(Point3d point)
        {
            double MillimetersToUnits(double value)
            {
                return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Millimeters);
            }

            return new XYZ(MillimetersToUnits(point.X), MillimetersToUnits(point.Y), MillimetersToUnits(point.Z));


        }

        private void CreateWindow(Document doc, Wall wall, double h, double fl, double kerbHeight)
        {
            // Neglect the kerbs 
            if (fl < kerbHeight)
            {
                fl = 0;
            }

            // Get the dimension of windows.
            LocationCurve locationCurve = wall.Location as LocationCurve;
            Autodesk.Revit.DB.Line location = locationCurve.Curve as Autodesk.Revit.DB.Line;
            XYZ startPoint = location.GetEndPoint(0);
            XYZ endPoint = location.GetEndPoint(1);
            String windowSize = Math.Round(UnitsToCentimeters(startPoint.DistanceTo(endPoint))).ToString() + "x" + h.ToString();

            // To collect all the window tpye in Revit Project.
            FilteredElementCollector Collector = new FilteredElementCollector(doc);
            List<FamilySymbol> familySymbolList = Collector.OfClass(typeof(FamilySymbol))
            .OfCategory(BuiltInCategory.OST_Windows)
            .Cast<FamilySymbol>().ToList();

            // To verify whether the window type is exist or not. If not, create a new one.
            Boolean IsColumnTypeExist = false;
            foreach (FamilySymbol fs in familySymbolList)
            {
                if (fs.Name != windowSize)
                {
                    continue;
                }
                else
                {
                    IsColumnTypeExist = true;
                    break;
                }
            }
            if (!IsColumnTypeExist)
            {
                using (Autodesk.Revit.DB.Transaction t_createNewColumnType = new Autodesk.Revit.DB.Transaction(doc, "Ｃreate New Window Type"))
                {
                    try
                    {
                        t_createNewColumnType.Start();

                        FamilySymbol default_FamilySymbol = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol))
                        .OfCategory(BuiltInCategory.OST_Windows)
                        .Cast<FamilySymbol>()
                        .FirstOrDefault(q => q.Name == "default") as FamilySymbol;

                        FamilySymbol newFamSym = default_FamilySymbol.Duplicate(windowSize) as FamilySymbol;
                        // set the radius to a new value:
                        IList<Parameter> pars = newFamSym.GetParameters("高度");
                        pars[0].Set(CentimetersToUnits(h));
                        IList<Parameter> pars_2 = newFamSym.GetParameters("寬度");
                        pars_2[0].Set(startPoint.DistanceTo(endPoint));

                        t_createNewColumnType.Commit();
                    }
                    catch (System.Exception ex)
                    {
                        TaskDialog td = new TaskDialog("error")
                        {
                            Title = "error",
                            AllowCancellation = true,
                            MainInstruction = "error",
                            MainContent = "Error" + ex.Message,
                            CommonButtons = TaskDialogCommonButtons.Close
                        };
                        td.Show();

                        Debug.Print(ex.Message);
                        t_createNewColumnType.RollBack();
                    }
                }
            }

            // Get the window0 type
            FamilySymbol windowType = null;
            foreach (FamilySymbol symbol in new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .Cast<FamilySymbol>())
            {
                if (symbol.Name == windowSize)
                {
                    windowType = symbol;
                    break;
                }
            }

            // Create window.
            Autodesk.Revit.DB.Transaction tw = new Autodesk.Revit.DB.Transaction(doc, "Create window on wall");
            tw.Start();
            Level level = doc.GetElement(wall.LevelId) as Level;
            XYZ wallMidpoint = (startPoint + endPoint) / 2 + new XYZ(0, 0, CentimetersToUnits(fl) + level.Elevation);
            FamilyInstance window = doc.Create.NewFamilyInstance(wallMidpoint, windowType, wall as Element, StructuralType.NonStructural);
            tw.Commit();
        }

        // Definition of CADTextModel and CADText 
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
