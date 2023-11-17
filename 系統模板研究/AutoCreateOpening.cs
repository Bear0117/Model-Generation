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
using System.Windows;
using static System.Net.Mime.MediaTypeNames;
using static Autodesk.Revit.DB.SpecTypeId;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateOpening : IExternalCommand
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

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Get the reference of Picked object.
            Autodesk.Revit.DB.Reference reference = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            Element elem2 = doc.GetElement(reference);
            GeometryElement geoElem = elem2.get_Geometry(new Options());
            GeometryObject geoObj = elem2.GetGeometryObjectFromReference(reference);

            // CAll the WinForm.
            UIApplication uiApp = commandData.Application;
            Document doc_ui = uiApp.ActiveUIDocument.Document;
            string levelN = "1F";
            double kerbHeight;
            double gridline_size;
            OpenWallForm form = new OpenWallForm(doc_ui);
            form.ShowDialog();
            levelN = form.levelName;
            kerbHeight = form.KerbHeight;
            gridline_size = form.gridsize;
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
                    //TaskDialog.Show("Revit", layername);
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

            List<CADTextModel> text = GetCADTextInfoparing(path, 0);

            // Pair two line as open.
            List<List<Curve>> doorBlocks = new List<List<Curve>>();
            List<List<Curve>> Clusters = Algorithm.ClusterByParallel(curves);

            // To store Wallmodel.
            List<Wallmodel> openingwall = new List<Wallmodel>();

            // To store width.
            List<double> widthtype = new List<double>();


            // Create 2D BoundingBox to get the middle line of open.
            foreach (List<Curve> clustergroup in Clusters)
            {
                if (Algorithm.CreateBoundingBox2D(clustergroup) != null)
                {
                    List<Wallmodel> comparewall = new List<Wallmodel>();
                    List<Line> compareline = new List<Line>();
                    List<Curve> finalcluster = Algorithm.CreateBoundingBox2D(clustergroup);

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
                                comparewall.Add(axses);
                                compareline.Add(axses.Wallaxes as Line);
                            }
                        }
                    }

                    // To find the final open middle line.
                    int finalnumber = 0;
                    if (compareline[1].Length > compareline[0].Length)
                    {
                        finalnumber = 1;
                    }
                    openingwall.Add(comparewall[finalnumber]);
                }
            } 

            Transform transform = null;
            foreach (GeometryObject gObj in geoElem)
            {
                GeometryInstance geomInstance = gObj as GeometryInstance;
                transform = geomInstance.Transform;
            }

            foreach (Wallmodel wall in openingwall)
            {
                double comparedistance = 0;
                double distanceBetween = double.MaxValue;

                if (text.Count >= 1)
                {
                    foreach (CADTextModel walltext in text)
                    {
                        XYZ mid_new = new XYZ(wall.Midpoint.X, wall.Midpoint.Y, 0) * 0.1;
                        comparedistance = mid_new.DistanceTo(transform.OfPoint(walltext.Location * 10) /10);
                        if (comparedistance < distanceBetween)
                        {
                            distanceBetween = comparedistance;
                            wall.HText = walltext.HText;
                            wall.WText = walltext.WText;
                        }
                    }
                }
            }
            

            // Get the level height.
            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(Level));
            double levelHeight = CentimetersToUnits(330); // default
            double bottomEleation = level.Elevation;

            List<double> levelDistancesList = new List<double>();

            // retrieve all the Level elements in the document
            IEnumerable<Level> levels = collector.Cast<Level>();

            // loop through all the levels
            foreach (Level level_1 in levels)
            {
                // get the level elevation.
                double levelElevation = level_1.Elevation;
                if (Math.Abs(levelElevation - bottomEleation) < CentimetersToUnits(0.1))
                {
                    level = level_1;
                }
                if (levelElevation - bottomEleation > CentimetersToUnits(0.1))
                {
                    levelDistancesList.Add(levelElevation);
                }
            }

            // Create openings.
            foreach (Wallmodel openwall in openingwall)
            {
                // THIS PROGRAM SHOULD BE CLARIFIED AND UPDATED.
                if (Math.Abs(openwall.Width - 209) < 1 || Math.Abs(openwall.Width - 358) < 1)
                {
                    continue;
                }

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
            if (Math.Abs(startPoint.Y - endPoint.Y) < CentimetersToUnits(0.001))
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
            foreach (Wall wall in walls)
            {
                bool hWall;
                LocationCurve location = wall.Location as LocationCurve;
                XYZ midPoint = (location.Curve.GetEndPoint(0) + location.Curve.GetEndPoint(1)) / 2;
                if (Math.Abs(location.Curve.GetEndPoint(0).Y - location.Curve.GetEndPoint(1).Y) < CentimetersToUnits(0.001))
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

            string windowSize = Math.Round(UnitsToCentimeters(startPoint.DistanceTo(endPoint))).ToString() + "x" + h.ToString();

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
            if(closestWall != null)
            {
                Level level = doc.GetElement(closestWall.LevelId) as Level;
                XYZ wallMidpoint = (startPoint + endPoint) / 2 + new XYZ(0, 0, CentimetersToUnits(fl));
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
        public List<CADTextModel> GetCADTextInfoparing(string dwgFile, double underbeam)
        {
            List<CADTextModel> CADModels = new List<CADTextModel>();
            List<CADText> HTEXT = new List<CADText>();
            List<CADText> FLTEXT = new List<CADText>();
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

                                            if (entity.Layer == layername)
                                            {
                                                if (entity.GetRXClass().Name == "AcDbText")
                                                {
                                                    Teigha.DatabaseServices.DBText text = (Teigha.DatabaseServices.DBText)entity;
                                                    if (text.TextString.Contains("h"))
                                                    {
                                                        htext.Text = ExtractNumber(text.TextString);
                                                        htext.Location = ConverCADPointToRevitPoint(text.Position);
                                                        HTEXT.Add(htext);
                                                    }

                                                    if (text.TextString.Contains("H"))
                                                    {
                                                        htext.Text = ExtractNumber(text.TextString);
                                                        htext.Location = ConverCADPointToRevitPoint(text.Position);
                                                        HTEXT.Add(htext);
                                                    }

                                                    else if (text.TextString.Contains("FL"))
                                                    {
                                                        Regex regex = new Regex(@"FL(\d+)");
                                                        Match match = regex.Match(text.TextString);
                                                        if (match.Success)
                                                        {
                                                            string fl = match.Groups[1].Value;
                                                            fltext.Location = ConverCADPointToRevitPoint(text.Position);
                                                            fltext.Text = fl;
                                                            FLTEXT.Add(fltext);
                                                        }
                                                    }
                                                    else continue;
                                                }

                                                if (entity.GetRXClass().Name == "AcDbMText")
                                                {
                                                    Teigha.DatabaseServices.MText mText = (Teigha.DatabaseServices.MText)entity;

                                                    if (mText.Text.Contains("h"))
                                                    {
                                                        htext.Text = ExtractNumber(mText.Text);
                                                        htext.Location = ConverCADPointToRevitPoint(mText.Location);
                                                        HTEXT.Add(htext);
                                                    }

                                                    if (mText.Text.Contains("H"))
                                                    {
                                                        htext.Text = ExtractNumber(mText.Text);
                                                        htext.Location = ConverCADPointToRevitPoint(mText.Location);
                                                        HTEXT.Add(htext);
                                                    }

                                                    else if (mText.Text.Contains("FL"))
                                                    {
                                                        Regex regex = new Regex(@"FL(\d+)");
                                                        Match match = regex.Match(mText.Text);
                                                        if (match.Success)
                                                        {
                                                            string fl = match.Groups[1].Value;
                                                            fltext.Location = ConverCADPointToRevitPoint(mText.Location);
                                                            fltext.Text = fl;
                                                            FLTEXT.Add(fltext);
                                                        }
                                                    }
                                                    else continue;
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

            // To check the text 
            {
                //    int count1 = 0;
                //    foreach (CADText height in HTEXT)
                //    {
                //        count1++;
                //        TaskDialog.Show("REVIT", "第" + count1.ToString() + "項為" + height.Text);
                //    }
                //    int count = 0;
                //    foreach (CADText height in FLTEXT)
                //    {
                //        count++;
                //        TaskDialog.Show("REVIT", "第" + count.ToString() + "項為" + height.Text);
                //    }
            }
            //MessageBox.Show(HTEXT.Count().ToString());

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
                        comparemodel.Location = MidPoint(height.Location, floorheight.Location);
                    }
                }
                CADModels.Add(comparemodel);
            }


            // Check the paired text.
            {
                //int count2 = 0;
                //foreach ( CADTextModel text  in CADModels)
                //{
                //    count2++;
                //    TaskDialog.Show("REVIT", "第" + count1.ToString() + "項為" + text.HText);
                //    TaskDialog.Show("REVIT", "第" + count1.ToString() + "項為" + text.WText);
                //}
            }
            return CADModels;
        }

        // Functions of operation.
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

        //public XYZ RoundPoint(XYZ point)
        //{
        //    XYZ newPoint = new XYZ(
        //        CentimetersToUnits(Math.Round(UnitsToCentimeters(point.X) * 2) / 2),
        //        CentimetersToUnits(Math.Round(UnitsToCentimeters(point.Y) * 2) / 2),
        //        CentimetersToUnits(Math.Round(UnitsToCentimeters(point.Z) * 2) / 2)
        //        );
        //    return newPoint;
        //}

        public double UnitsToCentimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
        }

        public XYZ PointCentimeterToUnit(XYZ point)
        {
            XYZ newPoint = new XYZ(
                CentimetersToUnits(point.X),
                CentimetersToUnits(point.Y),
                CentimetersToUnits(point.Z)
                );
            return newPoint;
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
    }
}
