using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;
using Line = Autodesk.Revit.DB.Line;
using Curve = Autodesk.Revit.DB.Curve;
using Transaction = Autodesk.Revit.DB.Transaction;
using Exception = System.Exception;
using System.Windows;
using static System.Net.Mime.MediaTypeNames;
using static Autodesk.Revit.DB.SpecTypeId;
using System.Xml.Linq;
using System.Threading;
using System.Text;
// 2024

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateWindows : IExternalCommand
    {
        string layername = null;

        // Get CAD path
        public string GetCADPath(ElementId cadLinkTypeID, Document revitDoc)
        {
            CADLinkType cadLinkType = revitDoc.GetElement(cadLinkTypeID) as CADLinkType;
            return ModelPathUtils.ConvertModelPathToUserVisiblePath(cadLinkType.GetExternalFileReference().GetAbsolutePath());
        }

        //執行檔案
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //123123 38958yu93
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Get the reference of Picked object.
            Autodesk.Revit.DB.Reference reference = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            Element elem2 = doc.GetElement(reference);
            GeometryElement geoElem = elem2.get_Geometry(new Options());
            GeometryObject geoObj = elem2.GetGeometryObjectFromReference(reference);

            // Initialize Parameters.
            ModelingParam.Initialize();
            double gridline_size = ModelingParam.parameters.General.GridSize * 10; // unit: mm
            double kerbHeight = ModelingParam.parameters.OpeningParam.KerbHeight;

            // Link the CAD with path
            string path = GetCADPath(elem2.GetTypeId(), doc);

            // Get layer name
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

            // Get all the CurveArray in CAD and put them into a list.
            List<CADModel> cadcurveArray = GetCurveArray(geoElem, graphicsStyleId1, gridline_size);

            // Store all the curve in CurveArray into a Curve list.
            List<Curve> curves = new List<Curve>();
            foreach (CADModel a in cadcurveArray)
            {
                foreach (Curve curve in a.CurveArray)
                {
                    curves.Add(curve);
                }
            }

            (List<CADTextModel> text, List<CADText> others) = GetCADTextInfoparing(path);

            // Pair two line as open.
            List<List<Curve>> Clusters = Algorithm.ClusterByParallel(curves);
            //MessageBox.Show(Clusters.Count().ToString());
            // To store Wallmodel.
            List<Wallmodel> openingwall = new List<Wallmodel>();

            // To store width.
            List<double> widthtype = new List<double>();


            foreach (List<Curve> clustergroup in Clusters)
            {
                Curve curve1 = clustergroup[0];
                Curve curve2 = clustergroup[1];

                // step 1: 取得兩條線的起點和終點
                XYZ start1 = curve1.GetEndPoint(0);
                XYZ end1 = curve1.GetEndPoint(1);
                XYZ start2 = curve2.GetEndPoint(0);
                XYZ end2 = curve2.GetEndPoint(1);

                bool invalid = false;
                foreach (CADTextModel walltext in text)
                {
                    if (AutoCreateBeams.IsPointInsideRectangleOnXYPlane(start1, start2, end1, end2, walltext.Location * 10))
                    {
                        invalid = true;
                        break;
                    }
                }
                if (invalid) continue;

                // step 2: 計算兩種配對的總距離
                double distA = start1.DistanceTo(start2) + end1.DistanceTo(end2);
                double distB = start1.DistanceTo(end2) + end1.DistanceTo(start2);

                // step 3: 根據較小距離決定哪種端點配對
                XYZ pairedStart1, pairedStart2, pairedEnd1, pairedEnd2;
                if (distA <= distB)
                {
                    // 配對 A：start1↔start2, end1↔end2
                    pairedStart1 = start1;
                    pairedStart2 = start2;
                    pairedEnd1 = end1;
                    pairedEnd2 = end2;
                }
                else
                {
                    // 配對 B：start1↔end2, end1↔start2
                    pairedStart1 = start1;
                    pairedStart2 = end2;
                    pairedEnd1 = end1;
                    pairedEnd2 = start2;
                }

                // step 4: 計算中心線兩端的中點
                XYZ centerStart = (pairedStart1 + pairedStart2) / 2;
                XYZ centerEnd = (pairedEnd1 + pairedEnd2) / 2.0;
                Line midLine = Line.CreateBound(centerStart, centerEnd);

                Wallmodel axses = new Wallmodel
                {
                    Wallaxes = midLine,
                    Width = GetWallWidth(curve1 as Line, curve2 as Line)
                };
                axses.Midpoint = GetMiddlePoint(axses.Wallaxes.GetEndPoint(0), axses.Wallaxes.GetEndPoint(1));
                openingwall.Add(axses);
            }



            Transform transform = null;
            foreach (GeometryObject gObj in geoElem)
            {
                GeometryInstance geomInstance = gObj as GeometryInstance;
                transform = geomInstance.Transform;
            }


            //MessageBox.Show(openingwall.Count.ToString());
            foreach (Wallmodel wall in openingwall)
            {
                double comparedistance;
                double distanceBetween = double.MaxValue;
                XYZ mid_new = new XYZ(wall.Midpoint.X, wall.Midpoint.Y, 0) * 0.1;
                CADTextModel bestWallText = null;

                if (text.Count >= 1)
                {
                    foreach (CADTextModel walltext in text)
                    {
                        comparedistance = mid_new.DistanceTo(transform.OfPoint(walltext.Location * 10) / 10);
                        if (comparedistance < distanceBetween)
                        {
                            distanceBetween = comparedistance;
                            bestWallText = walltext;
                            //// WText is the floor height of window.
                            //wall.HText = walltext.HText;
                            //wall.WText = walltext.WText;
                        }
                    }
                    // 離開內層迴圈後，如果成功找到最相近的 CADTextModel
                    if (bestWallText != null)
                    {
                        // 將該 CADTextModel 的內容填到目前 wall
                        wall.HText = bestWallText.HText;
                        wall.WText = bestWallText.WText;

                        // 將該配對文字的 num_paired 加 1，紀錄它又被配對了一次
                        bestWallText.NumPaired++;
                    }
                }
            }

            List<(ElementId, string)> errorData = new List<(ElementId, string)>();
            // 處理重複配對
            foreach (CADTextModel walltext in text)
            {
                if (walltext.NumPaired != 1)
                {
                    FamilySymbol default_column = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .OfCategory(BuiltInCategory.OST_Columns)
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(q => q.Name == "default") as FamilySymbol;
                    using (Transaction tx = new Transaction(doc))
                    {
                        try
                        {
                            tx.Start("createColumn");
                            if (!default_column.IsActive)
                            {
                                default_column.Activate();
                            }
                            Level level = doc.GetElement(elem2.LevelId) as Level; //elem2是CAD圖

                            FamilyInstance familyInstance = doc.Create.NewFamilyInstance(transform.OfPoint(walltext.Location * 10), default_column, level, StructuralType.Column);
                            tx.Commit();

                            string error_meesage = $"This label(\"FL{walltext.WText}\" and \"h={walltext.HText}\") has been paired \"{walltext.NumPaired}\" time(s).";
                            errorData.Add((familyInstance.Id, error_meesage));
                        }
                        catch (Exception ex)
                        {
                            TaskDialog td = new TaskDialog("error") { Title = "error", AllowCancellation = true, MainInstruction = "error", MainContent = "Error" + ex.Message, CommonButtons = TaskDialogCommonButtons.Close };
                            td.Show();
                            Debug.Print(ex.Message);
                            tx.RollBack();
                        }
                    }

                }
            }

            // 處理文字解析警告
            foreach (CADText tx in others)
            {
                if (tx.Text.Contains("W="))
                {
                    ElementId id = null;
                    string error_meesage = $"Warning: This label (\"{tx.Text})\" does not contain FL and h but contains W.";
                    errorData.Add((id, error_meesage));
                }
                else
                {
                    FamilySymbol default_column = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .OfCategory(BuiltInCategory.OST_Columns)
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(q => q.Name == "default") as FamilySymbol;
                    using (Transaction transaction = new Transaction(doc))
                    {
                        try
                        {
                            transaction.Start("createColumn");
                            if (!default_column.IsActive) default_column.Activate();
                            Level level = doc.GetElement(elem2.LevelId) as Level;

                            FamilyInstance familyInstance = doc.Create.NewFamilyInstance(transform.OfPoint(tx.Location * 10), default_column, level, StructuralType.Column);
                            transaction.Commit();

                            string error_meesage = $"This label(\"{tx.Text}\") is in the wrong format.";
                            errorData.Add((familyInstance.Id, error_meesage));
                        }
                        catch (Exception ex)
                        {
                            TaskDialog td = new TaskDialog("error") { Title = "error", AllowCancellation = true, MainInstruction = "error", MainContent = "Error" + ex.Message, CommonButtons = TaskDialogCommonButtons.Close };
                            td.Show();
                            Debug.Print(ex.Message);
                            transaction.RollBack();
                        }
                    }
                }
            }
            errorData = errorData.OrderBy(item => item.Item2.StartsWith("W", StringComparison.OrdinalIgnoreCase) ? 1 : 0).ThenBy(item => item.Item2, StringComparer.OrdinalIgnoreCase).ToList();
            if (errorData != null) ExportChecklistToCsv(errorData);

            // Create openings.
            foreach (Wallmodel openwall in openingwall)
            {
                double h = Convert.ToDouble(openwall.HText);
                double fl = Convert.ToDouble(openwall.WText);
                CreateOpening(doc, openwall.Wallaxes, h, fl, kerbHeight);
            }
            return Result.Succeeded;
        }
        private void CreateOpening(Document doc, Autodesk.Revit.DB.Curve wallLine, double h, double fl, double kerbHeight)
        {

            if (fl < kerbHeight)
            {
                h += fl;
                fl = 0;
            }

            // WallLine Equal To OpenLine.
            XYZ startPoint = wallLine.GetEndPoint(0);
            XYZ endPoint = wallLine.GetEndPoint(1);
            XYZ openPoint = (startPoint + endPoint) / 2;
            bool hOpen;
            if (Math.Abs(startPoint.Y - endPoint.Y) < CentimetersToUnits(1))
            {
                hOpen = true;
            }
            else
            {
                hOpen = false;
            }

            FilteredElementCollector walls = new FilteredElementCollector(doc)
            .OfCategory(BuiltInCategory.OST_Walls)
            .WhereElementIsNotElementType();

            // To ensure open on the correct wall,the horizontal distance between the wall and open should be 0.
            double hdist;

            // Find the wall closest to the point
            Wall closestWall = null;
            double closestDistance = double.MaxValue;
            foreach (Element elem in walls)
            {
                bool hWall;
                if (!(elem is Wall wall)) continue;
                LocationCurve location = wall.Location as LocationCurve;
                XYZ midPoint = (location.Curve.GetEndPoint(0) + location.Curve.GetEndPoint(1)) / 2;
                if (Math.Abs(location.Curve.GetEndPoint(0).Y - location.Curve.GetEndPoint(1).Y) < CentimetersToUnits(1))
                {
                    hWall = true;
                }
                else
                {
                    hWall = false;
                }

                if (hWall && hOpen)
                {
                    hdist = Math.Abs(startPoint.Y - location.Curve.GetEndPoint(0).Y);
                }
                else if (!hWall && !hOpen)
                {
                    hdist = Math.Abs(startPoint.X - location.Curve.GetEndPoint(0).X);
                }
                else
                {
                    continue;
                }

                // get the location curve of the wall
                LocationCurve locationCurve = wall.Location as LocationCurve;

                // get the location line of the wall
                Line locationLine = locationCurve.Curve as Line;

                if (hOpen == hWall)
                {
                    double distance = openPoint.DistanceTo(midPoint);
                    if (distance < closestDistance)
                    {
                        if (locationCurve == null || Math.Abs(locationLine.Distance(startPoint)) < CentimetersToUnits(1))
                        {
                            closestWall = wall;
                            closestDistance = distance;
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
            }

            string windowSize = (Math.Round(Algorithm.UnitsToMillimeters(startPoint.DistanceTo(endPoint))) / 10).ToString() + "x" + h.ToString();

            // To collect all the window tpye in Revit Project.
            FilteredElementCollector Collector = new FilteredElementCollector(doc);
            List<FamilySymbol> familySymbolList = Collector.OfClass(typeof(FamilySymbol))
            .OfCategory(BuiltInCategory.OST_Windows)
            .Cast<FamilySymbol>().ToList();

            // To verify whether the window type is exist or not. If not, create a new one.
            bool IsWindowTypeExist = false;
            foreach (FamilySymbol fs in familySymbolList)
            {
                if (fs.Name != windowSize)
                {
                    continue;
                }
                else
                {
                    IsWindowTypeExist = true;
                    break;
                }
            }
            if (!IsWindowTypeExist)
            {
                using (Transaction t_createNewColumnType = new Transaction(doc, "Ｃreate New Window Type"))
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
            Transaction tw = new Transaction(doc, "Create window on wall");
            tw.Start();
            if (closestWall != null)
            {
                Level level = doc.GetElement(closestWall.LevelId) as Level;
                XYZ wallMidpoint = (startPoint + endPoint) / 2 + new XYZ(0, 0, Algorithm.CentimetersToUnits(fl));
                FamilyInstance window = doc.Create.NewFamilyInstance(wallMidpoint, windowType, closestWall, level, StructuralType.NonStructural);
                Parameter levelParam = window.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
                levelParam?.Set(level.Id);
            }
            tw.Commit();
        }

        // Get all the CurveArray in CAD and put them into a list.
        private List<CADModel> GetCurveArray(GeometryElement geoElem, ElementId graphicsStyleId, double gridline_size)
        {
            List<CADModel> curveArray_List = new List<CADModel>();

            foreach (GeometryObject gObj in geoElem)
            {
                GeometryInstance geomInstance = gObj as GeometryInstance;
                // Coordination Transform.
                Transform transform = geomInstance.Transform;
                if (null != geomInstance)
                {
                    foreach (GeometryObject insObj in geomInstance.SymbolGeometry)
                    {
                        if (insObj.GraphicsStyleId.IntegerValue != graphicsStyleId.IntegerValue)
                            continue;

                        // For straight line.
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.Line")
                        {
                            Line line = insObj as Line;
                            XYZ point = line.GetEndPoint(0);
                            _ = transform.OfPoint(point);

                            Line newLine = TransformLine(transform, line);
                            Line newLine_rounded = RoundLine(newLine, gridline_size);

                            CurveArray curveArray = new CurveArray();
                            curveArray.Append(newLine_rounded);

                            XYZ startPoint = newLine_rounded.GetEndPoint(0);
                            XYZ endPoint = newLine_rounded.GetEndPoint(1);
                            XYZ MiddlePoint = GetMiddlePoint(startPoint, endPoint);
                            double angle = (startPoint.Y - endPoint.Y) / startPoint.DistanceTo(endPoint);
                            double rotation = Math.Asin(angle);

                            CADModel cADModel = new CADModel
                            {
                                CurveArray = curveArray,
                                Length = newLine_rounded.Length,
                                Shape = "牆",
                                Width = UnitsToMillimeters(300),
                                Location = MiddlePoint,
                                Rotation = rotation
                            };
                            curveArray_List.Add(cADModel);

                        }

                        // For polyline.
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> points = polyLine.GetCoordinates();

                            for (int i = 0; i < points.Count - 1; i++)
                            {
                                Line line = Line.CreateBound(points[i], points[i + 1]);
                                line = TransformLine(transform, line);
                                Line newLine = line;
                                Line newLine_rounded = RoundLine(newLine, gridline_size);
                                CurveArray curveArray = new CurveArray();
                                curveArray.Append(newLine_rounded);

                                XYZ startPoint = newLine_rounded.GetEndPoint(0);
                                XYZ endPoint = newLine_rounded.GetEndPoint(1);
                                XYZ MiddlePoint = GetMiddlePoint(startPoint, endPoint);
                                double angle = (startPoint.Y - endPoint.Y) / startPoint.DistanceTo(endPoint);
                                double rotation = Math.Asin(angle);

                                CADModel cADModel = new CADModel
                                {
                                    CurveArray = curveArray,
                                    Length = newLine_rounded.Length,
                                    Shape = "牆",
                                    Width = UnitsToMillimeters(300),
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

        // Get the text in CAD.
        public (List<CADTextModel>, List<CADText>) GetCADTextInfoparing(string dwgFile)
        {
            List<CADTextModel> CADModels = new List<CADTextModel>();
            List<CADText> HTEXT = new List<CADText>();
            List<CADText> FLTEXT = new List<CADText>();
            List<CADText> OTHERS = new List<CADText>();
            //List<ObjectId> allObjectId = new List<ObjectId>();

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
                                            CADText others = new CADText();

                                            if (entity.Layer == layername)
                                            {
                                                if (entity.GetRXClass().Name == "AcDbText")
                                                {
                                                    Teigha.DatabaseServices.DBText text = (Teigha.DatabaseServices.DBText)entity;
                                                    if (text.TextString.IndexOf("h", StringComparison.OrdinalIgnoreCase) >= 0)
                                                    {
                                                        try
                                                        {
                                                            htext.Text = ExtractNumber(text.TextString);
                                                            htext.Location = ConverCADPointToRevitPoint(text.Position);
                                                            HTEXT.Add(htext);
                                                            if (!IsExactlyOneHAndDigits(text.TextString))
                                                            {
                                                                throw new Exception($"This label(\"{text.TextString}\")  is in the wrong format.");
                                                            }
                                                        }
                                                        catch(Exception)
                                                        {
                                                            others.Text = text.TextString;
                                                            others.Location = ConverCADPointToRevitPoint(text.Position);
                                                            OTHERS.Add(others);
                                                        }
                                                    }
                                                    else if (text.TextString.Contains("FL"))
                                                    {
                                                        try
                                                        {
                                                            fltext.Text = ExtractNumber(text.TextString);
                                                            fltext.Location = ConverCADPointToRevitPoint(text.Position);
                                                            FLTEXT.Add(fltext);
                                                            if (!IsExactlyOneFOneLAndDigits(text.TextString))
                                                            {
                                                                throw new Exception($"This label(\"{text.TextString}\")  is in the wrong format.");
                                                            }
                                                        }
                                                        catch (Exception)
                                                        {
                                                            others.Text = text.TextString;
                                                            others.Location = ConverCADPointToRevitPoint(text.Position);
                                                            OTHERS.Add(others);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        others.Text = text.TextString;
                                                        others.Location = ConverCADPointToRevitPoint(text.Position);
                                                        OTHERS.Add(others);
                                                    }

                                                }

                                                if (entity.GetRXClass().Name == "AcDbMText")
                                                {
                                                    MText mText = (MText)entity;

                                                    if (mText.Text.IndexOf("h", StringComparison.OrdinalIgnoreCase) >= 0)
                                                    {
                                                        try
                                                        {
                                                            htext.Text = ExtractNumber(mText.Text);
                                                            htext.Location = ConverCADPointToRevitPoint(mText.Location);
                                                            HTEXT.Add(htext);
                                                            if (!IsExactlyOneHAndDigits(mText.Text))
                                                            {
                                                                throw new Exception($"This label(\"{mText.Text}\")  is in the wrong format.");
                                                            }
                                                        }
                                                        catch (Exception)
                                                        {
                                                            others.Text = mText.Text;
                                                            others.Location = ConverCADPointToRevitPoint(mText.Location);
                                                            OTHERS.Add(others);
                                                        }
                                                    }
                                                    else if (mText.Text.Contains("FL"))
                                                    {
                                                        try
                                                        {
                                                            fltext.Text = ExtractNumber(mText.Text);
                                                            fltext.Location = ConverCADPointToRevitPoint(mText.Location);
                                                            FLTEXT.Add(fltext);
                                                            if (!IsExactlyOneFOneLAndDigits(mText.Text))
                                                            {
                                                                throw new Exception($"This label(\"{mText.Text}\")  is in the wrong format.");
                                                            }
                                                        }
                                                        catch (Exception)
                                                        {
                                                            others.Text = mText.Text;
                                                            others.Location = ConverCADPointToRevitPoint(mText.Location);
                                                            OTHERS.Add(others);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        others.Text = mText.Text;
                                                        others.Location = ConverCADPointToRevitPoint(mText.Location);
                                                        OTHERS.Add(others);
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

            // Pair height text with FL text.
            foreach (CADText height in HTEXT)
            {
                // Store compared text.
                CADTextModel comparemodel = new CADTextModel();
                double comparedistance;
                double distanceBetween = double.MaxValue;

                foreach (CADText floorheight in FLTEXT)
                {
                    comparedistance = height.Location.DistanceTo(floorheight.Location);
                    if (comparedistance < distanceBetween)
                    {
                        distanceBetween = comparedistance;
                        comparemodel.HText = height.Text;
                        comparemodel.WText = floorheight.Text;
                        //comparemodel.Location = MidPoint(height.Location, floorheight.Location);
                        comparemodel.Location = floorheight.Location;
                    }
                }
                CADModels.Add(comparemodel);
            }

            return (CADModels, OTHERS);
        }

        // Functions of operation.
        bool IsExactlyOneFOneLAndDigits(string input)
        {
            string pattern = @"^FL[0-9.]+$"; ;
            return Regex.IsMatch(input, pattern);
        }

        bool IsExactlyOneHAndDigits(string input)
        {
            string pattern = @"^[hH]=[0-9.]+$";
            return Regex.IsMatch(input, pattern);
        }

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

        private Line TransformLine(Transform transform, Line line)
        {
            XYZ startPoint = transform.OfPoint(line.GetEndPoint(0));
            XYZ endPoint = transform.OfPoint(line.GetEndPoint(1));
            Line newLine = Line.CreateBound(startPoint, endPoint);
            return newLine;
        }

        public int GetWallWidth(Line line1, Line line2)
        {
            XYZ ponit2_2 = line2.GetEndPoint(1);
            line1.MakeUnbound();
            double width = line1.Distance(ponit2_2);
            int wallWidth = (int)Math.Round(UnitsToMillimeters(width));

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

        public double CentimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Centimeters);
        }

        public double UnitsToMillimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters);
        }

        static string ExtractNumber(string str)
        {
            // 正則表達式中的 \d 代表數字，點號代表小數點，加號表示一個或多個
            Regex regex = new Regex(@"\d+\.?\d*");
            Match match = regex.Match(str);

            if (match.Success)
            {
                return match.Value;
            }
            else
            {
                return "";
            }
        }

        public Line RoundLine(Line line, double gridline_size)
        {
            XYZ newPoint0 = Algorithm.RoundPoint(line.GetEndPoint(0), gridline_size);
            XYZ newPoint1 = Algorithm.RoundPoint(line.GetEndPoint(1), gridline_size);
            Line newLine = Line.CreateBound(newPoint0, newPoint1);
            return newLine;
        }

        public static void ExportChecklistToCsv(IEnumerable<(ElementId ElementId, string ErrorMessage)> data)
        {
            // Determine folder path on user's desktop
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string folderPath = Path.Combine(desktopPath, "Checklist_Modeling");

            // Ensure folder exists (create if not)
            Directory.CreateDirectory(folderPath);

            // Create the CSV file path
            string filePath = Path.Combine(folderPath, "Checklist_Window.csv");

            // Write to CSV
            // - Overwrite file if it already exists (set append to 'false')
            // - Use UTF-8 encoding
            using (StreamWriter writer = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                // Write CSV header
                writer.WriteLine("ElementId,Error Message");

                // Write each row
                foreach ((ElementId ElementId, string ErrorMessage) pair in data)
                {
                    // ElementId is typically an integer; use pair.ElementId.IntegerValue
                    if(pair.ElementId != null)
                    {
                        writer.WriteLine($"{pair.ElementId.IntegerValue},{pair.ErrorMessage}");
                    }
                    else writer.WriteLine($"{pair.ElementId},{pair.ErrorMessage}");
                }
            }
        }


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

    public class CADText
    {
        private string context;

        private XYZ textlocation;

        public List<string> WTexts { get; set; }

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

    public class CADTextModel
    {
        public CADTextModel()
        {
            NumPaired = 0;
        }
        public Level Level { get; set; }

        public int NumPaired { get; set; }

        public string HText { get; set; }

        public string WText { get; set; }

        public double Distant { get; set; }

        public XYZ Location { get; set; }
    }
}

