using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Teigha.DatabaseServices;
using Teigha.Geometry;
using Teigha.Runtime;
using Line = Autodesk.Revit.DB.Line;
using Curve = Autodesk.Revit.DB.Curve;
using Transaction = Autodesk.Revit.DB.Transaction;
using Arc = Autodesk.Revit.DB.Arc;
using Aspose.Pdf.Drawing;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateOpening_old : IExternalCommand
    {
        string layername = null;

        // Get CAD path
        public string GetCADPath(ElementId cadLinkTypeID, Document revitDoc)
        {
            CADLinkType cadLinkType = revitDoc.GetElement(cadLinkTypeID) as CADLinkType;
            return ModelPathUtils.ConvertModelPathToUserVisiblePath(cadLinkType.GetExternalFileReference().GetAbsolutePath());
        }
        double gridSize;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Get the reference of Picked object.
            Reference reference = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            Element elem2 = doc.GetElement(reference);
            GeometryElement geoElem = elem2.get_Geometry(new Options());
            GeometryObject geoObj = elem2.GetGeometryObjectFromReference(reference);

            // CAll the WinForm.
            UIApplication uiApp = commandData.Application;
            Document doc_ui = uiApp.ActiveUIDocument.Document;
            string levelN = "1F";
            double kerbHeight;
            
            OpenWallForm form = new OpenWallForm(doc_ui);
            form.ShowDialog();
            levelN = form.levelName;
            kerbHeight = form.KerbHeight;
            gridSize = form.gridsize;
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

                    TaskDialog.Show("Revit", layername);
                }
            }

            // Get all the CurveArray in CAD and put them into a list.
            List<CADModel> cadcurveArray = GetCurveArray(geoElem, graphicsStyleId1);

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

            // Paired the open with text
            foreach (Wallmodel wall in openingwall)
            {
                double comparedistance = double.MinValue;
                double distanceBetween = double.MaxValue;

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
            Transaction t3 = new Transaction(doc, "Create Openings");
            t3.Start();
            List<Arc> arc_list = GetArcInCAD(geoElem,graphicsStyleId1);

            bool isLabelLine = false;
            foreach (Wallmodel openwall in openingwall)
            {
                foreach(Arc arc in arc_list)
                {
                    double radius = arc.Radius;
                    XYZ center = arc.Center;
                    if (openwall.Midpoint.DistanceTo(center) <= radius)
                    {
                        isLabelLine = true;
                        break;
                    }
                }

                if (!isLabelLine)
                {
                    double h = Convert.ToDouble(openwall.HText);
                    double fl = Convert.ToDouble(openwall.WText);
                    CreateOpening(doc, openwall.Wallaxes, h, fl, kerbHeight);
                }
                
                // THIS PROGRAM SHOULD BE CLARIFIED AND UPDATED.
                //if (Math.Abs(openwall.Width - 209) < 1 || Math.Abs(openwall.Width - 358) < 1)
                //{
                //    continue;
                //}

                // create a WallType.
                //ElementId id = CreatWallType(doc, openwall.Width.ToString(), MillimetersToUnits(openwall.Width));

                
            }
            t3.Commit();

            return Result.Succeeded;
        }

        //Create opening.
        private void CreateOpening(Document doc, Curve wallLine, double h, double fl, double kerbHeight)
        {
            //if (fl < kerbHeight)
            //{
            //    h += fl;
            //    fl = 0;
            //}

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
            foreach (Wall wall in walls.Cast<Wall>())
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
                //LocationCurve locationCurve = wall.Location as LocationCurve;

                // get the location line of the wall
                Line locationLine = location.Curve as Line;

                if (hOpen == hWall)
                {
                    double distance = openPoint.DistanceTo(midPoint);
                    if (distance < closestDistance)
                    {

                        if (Math.Abs(locationLine.Distance(startPoint)) < CentimetersToUnits(1))
                        {
                            //MessageBox.Show(Math.Abs(locationLine.Distance(startPoint)).ToString());
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

            // The bottom is less than 15cm means it is kerb.

            if (closestWall != null)
            {

                XYZ bottomCoor = new XYZ(0, 0, CentimetersToUnits(fl));
                XYZ topCoor = new XYZ(0, 0, CentimetersToUnits(fl + h));
                Opening opening = doc.Create.NewOpening(closestWall, startPoint + bottomCoor, endPoint + topCoor);
            }
        }

        // Get all the CurveArray in CAD and put them into a list.
        private List<CADModel> GetCurveArray(GeometryElement geoElem, ElementId graphicsStyleId)
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
                            Line newLine_rounded = RoundLine(newLine, gridSize);

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
                                Line newLine_rounded = RoundLine(newLine,gridSize);
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

        private List<Arc> GetArcInCAD(GeometryElement geoElem, ElementId graphicsStyleId)
        {
            List<Arc> arc_List = new List<Arc>();

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

                        // For Arc.
                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.Arc")
                        {
                            Arc arc = insObj as Arc;

                            // transform Arc
                            XYZ startPoint = transform.OfPoint(arc.GetEndPoint(0));
                            XYZ endPoint = transform.OfPoint(arc.GetEndPoint(1));
                            XYZ midPoint = transform.OfPoint(arc.Evaluate(0.5, true));
                            Arc arc_tr = Arc.Create(startPoint, endPoint, midPoint);

                            arc_List.Add(arc_tr);
                        }

                    }
                }
            }
            return arc_List;
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
                                                        htext.Text = Regex.Replace(text.TextString, "[^0-9]", "");
                                                        htext.Location = ConverCADPointToRevitPoint(text.Position);
                                                        HTEXT.Add(htext);
                                                    }

                                                    else if (text.TextString.Contains("FL"))
                                                    {
                                                        fltext.Location = ConverCADPointToRevitPoint(text.Position);

                                                        string pattern = @"FL(\d+)";
                                                        Match match = Regex.Match(text.TextString, pattern);
                                                        if (match.Success)
                                                        {
                                                            string numbersString = match.Groups[1].Value;
                                                            fltext.Text = numbersString;
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
                                                        htext.Text = Regex.Replace(mText.Text, "[^0-9]", "");
                                                        htext.Location = ConverCADPointToRevitPoint(mText.Location);
                                                        HTEXT.Add(htext);
                                                    }

                                                    else if (mText.Text.Contains("FL"))
                                                    {
                                                        fltext.Location = ConverCADPointToRevitPoint(mText.Location);

                                                        string pattern = @"FL(\d+)";
                                                        Match match = Regex.Match(mText.Text, pattern);
                                                        if (match.Success)
                                                        {
                                                            string numbersString = match.Groups[1].Value;
                                                            fltext.Text = numbersString;
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

        public Line RoundLine(Line line, double gs)
        {
            XYZ newPoint0 = RoundPoint(line.GetEndPoint(0), gs);
            XYZ newPoint1 = RoundPoint(line.GetEndPoint(1), gs);
            Line newLine = Line.CreateBound(newPoint0, newPoint1);
            return newLine;
        }

        public XYZ RoundPoint(XYZ point, double gridSize)
        {
            XYZ newPoint = new XYZ(
                CentimetersToUnits(Math.Round(UnitsToCentimeters(point.X) / gridSize) * gridSize),
                CentimetersToUnits(Math.Round(UnitsToCentimeters(point.Y) / gridSize) * gridSize),
                CentimetersToUnits(Math.Round(UnitsToCentimeters(point.Z) / gridSize) * gridSize)
                );
            return newPoint;
        }

        public double UnitsToCentimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
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
