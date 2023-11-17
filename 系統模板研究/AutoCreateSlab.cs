using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;
using System.IO;
using Teigha.DatabaseServices;
using Teigha.Runtime;
using Teigha.Geometry;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Windows;
using Line = Autodesk.Revit.DB.Line;
using Transaction = Autodesk.Revit.DB.Transaction;
using static System.Net.Mime.MediaTypeNames;
using Aspose.Cells;
using Aspose.Pdf.Text;
using System.Windows.Media;
using Transform = Autodesk.Revit.DB.Transform;
using Curve = Autodesk.Revit.DB.Curve;
using Floor = Autodesk.Revit.DB.Floor;
using Aspose.Cells.Charts;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(RegenerationOption.Manual)]
    public class AutoCreateSlabs : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            Reference refer = uidoc.Selection.PickObject(ObjectType.PointOnElement, "Select a CAD Layer");
            Element elem = doc.GetElement(refer);
            GeometryObject geoObj = elem.GetGeometryObjectFromReference(refer);
            GeometryElement geoElem = elem.get_Geometry(new Options());
            Category targetCategory = null;
            ElementId graphicsStyleId = null;

            string layer_name = "S-Slab"; //Default
            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                graphicsStyleId = geoObj.GraphicsStyleId;
                using (GraphicsStyle gs = doc.GetElement(geoObj.GraphicsStyleId) as GraphicsStyle)
                {
                    if (gs != null)
                    {
                        targetCategory = gs.GraphicsStyleCategory;
                        layer_name = gs.GraphicsStyleCategory.Name;
                    }
                }
            }

            if (geoElem == null || graphicsStyleId == null)
            {
                message = "GeometryElement or ID does not exist！";
                return Result.Failed;
            }

            // Decide gridline size by users.
            double gridline_size;
            GridlineForm form = new GridlineForm();
            form.ShowDialog();
            gridline_size = form.gridlineSize;

            // Get the level of CAD drawing.
            Level level = doc.GetElement(elem.LevelId) as Level;

            // Start a TransactionGroup.
            TransactionGroup transGroup = new TransactionGroup(doc, "Start");
            transGroup.Start();

            // Get all the outline in layer of slab and copy them.
            List<CurveLoop> allOutlines = GetAllOutlines(doc, geoObj, geoElem);

            // 將outline分類成樓板或開口
            HashSet<CurveLoop> determinedOpenings = new HashSet<CurveLoop>();

            List<CurveLoop> slabOutlines = new List<CurveLoop>();
            List<CurveLoop> openOutlines = new List<CurveLoop>();

            foreach (CurveLoop outline_A in allOutlines)
            {
                if (determinedOpenings.Contains(outline_A))
                    continue;

                XYZ point_A = GetPolygonMidPoint(outline_A);
                bool isOpen = false;

                foreach (CurveLoop outline_B in allOutlines)
                {
                    if (outline_A == outline_B || determinedOpenings.Contains(outline_B))
                        continue;

                    if (IsInsideOutline(point_A, outline_B))
                    {
                        openOutlines.Add(outline_A);
                        determinedOpenings.Add(outline_A);
                        isOpen = true;
                        break;
                    }
                }

                if (!isOpen)
                {
                    slabOutlines.Add(outline_A);
                }
            }

            //Create SlabModel to store information of each slab.
            List<SlabModel> slabModels = new List<SlabModel>();
            foreach (CurveLoop curveLoop in slabOutlines)
            {
                SlabModel s = new SlabModel
                {
                    CurveLoop = RoundCurveLoop(curveLoop, gridline_size)
                };
                slabModels.Add(s);
                foreach (CurveLoop open in openOutlines)
                {
                    XYZ point_A = GetPolygonMidPoint(open);
                    if (IsInsideOutline(point_A, curveLoop))
                    {
                        s.Open = ConvertCurveLoopToCurveArray(RoundCurveLoop(open, gridline_size));
                        break;
                    }
                }
            }
            MessageBox.Show(slabOutlines.Count.ToString());
            MessageBox.Show(openOutlines.Count.ToString());





            // 設定樓板位置座標
            slabModels = GetLocation(slabModels);

            // Pair slabs with labels by distances.
            string path = GetCADPath(elem.GetTypeId(), doc);
            List<SlabModel> slabModels_paired = GetPairedLabel(slabModels, path, layer_name, geoElem);

            // To create rectangular slabs

            
        
            //List<SlabModel> finalRecSlab = new List<SlabModel>();
            //foreach (SlabModel slabOutline in slabModels_paired)
            //{
            //    finalRecSlab.Add(slabOutline);
            //}
            CreateFloor(doc, slabModels_paired, level, openOutlines);

            transGroup.Assimilate();
            return Result.Succeeded;
        }

        public CurveArray ConvertCurveLoopToCurveArray(CurveLoop curveLoop)
        {
            CurveArray curveArray = new CurveArray();
            foreach (Curve curve in curveLoop)
            {
                curveArray.Append(curve);
            }
            return curveArray;
        }


        public List<XYZ> RemoveNearDuplicates(List<XYZ> points, double tolerance = 0.01)
        {
            List<XYZ> distinctPoints = new List<XYZ>();

            foreach (var point in points)
            {
                bool isNearDuplicate = false;

                foreach (var distinctPoint in distinctPoints)
                {
                    if (point.DistanceTo(distinctPoint) < tolerance)
                    {
                        isNearDuplicate = true;
                        break;
                    }
                }

                if (!isNearDuplicate)
                {
                    distinctPoints.Add(point);
                }
            }

            return distinctPoints;
        }

        public CurveLoop RoundCurveLoop(CurveLoop curveLoop, double gridline_size)
        {
            List<Autodesk.Revit.DB.Curve> curveList = curveLoop.ToList();
            CurveLoop curevLoop_new = new CurveLoop();
            List<Autodesk.Revit.DB.Curve> curveList_new = new List<Autodesk.Revit.DB.Curve>();
            //List<XYZ> points = new List<XYZ>();
            foreach (Autodesk.Revit.DB.Curve curve in curveList)
            {

                XYZ point1 = Algorithm.RoundPoint(curve.GetEndPoint(0), gridline_size);
                XYZ point2 = Algorithm.RoundPoint(curve.GetEndPoint(1), gridline_size);
                if (point1.DistanceTo(point2) > CentimetersToUnits(0.0001))
                {
                    Autodesk.Revit.DB.Curve newCurve = Autodesk.Revit.DB.Line.CreateBound(point1, point2) as Autodesk.Revit.DB.Curve;
                    curevLoop_new.Append(newCurve);
                }
            }
            return curevLoop_new;
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

        public void CreateFloor(Autodesk.Revit.DB.Document doc, List<SlabModel> slabs, Level level, List<CurveLoop> opens)
        {
            foreach (SlabModel c in slabs)
            {
                Floor floor1 = null;
                using (Autodesk.Revit.DB.Transaction tx = new Autodesk.Revit.DB.Transaction(doc))
                {
                    try
                    {

                        tx.Start("Create Slab.");
                        string familyName = "RC slab(" + c.Thickness.ToString() + ")";

                        //MessageBox.Show(familyName);
                        FilteredElementCollector Collector = new FilteredElementCollector(doc);
                        List<FloorType> familySymbolList = Collector.OfClass(typeof(FloorType))
                        .OfCategory(BuiltInCategory.OST_Floors)
                        .Cast<FloorType>().ToList();
                        Boolean IsSlabTypeExist = false;
                        foreach (FloorType fs in familySymbolList)
                        {
                            if (fs.Name != familyName)
                            {
                                continue;
                            }
                            else
                            {
                                IsSlabTypeExist = true;
                                break;
                            }
                        }

                        if (!IsSlabTypeExist)
                        {
                            CreateSlabType(doc, familyName);
                        }

                        FloorType floorType = new FilteredElementCollector(doc)
                        .OfClass(typeof(FloorType))
                        .OfCategory(BuiltInCategory.OST_Floors)
                        .Cast<FloorType>()
                        .FirstOrDefault(q => q.Name == familyName) as FloorType;

                        ElementId floorTypeId = floorType.Id;
                        if (slabs != null)
                        {
                            // 檢查CurveLoop是否是開放的
                            if (c.CurveLoop.IsOpen())
                            {
                                throw new InvalidOperationException("The CurveLoop is open and cannot be used to create a floor.");
                            }



                            floor1 = Floor.Create(doc, new List<CurveLoop> { c.CurveLoop }, floorTypeId, level.Id);
                            Parameter distanceToLevelParam = floor1.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                            distanceToLevelParam.Set(CentimetersToUnits(c.Elevation));

                            
                            
                        }
                        tx.Commit();



                        
                    }
                    catch (System.Exception ex)
                    {
                        TaskDialog td = new TaskDialog("error")
                        {
                            Title = "error1",
                            AllowCancellation = true,
                            MainInstruction = "error",
                            MainContent = "Error" + ex.Message,
                            CommonButtons = TaskDialogCommonButtons.Close
                        };
                        td.Show();
                        Debug.Print(ex.Message);
                        tx.RollBack();
                    }
                }

                Transaction tf = new Transaction(doc, "Create Open");
                tf.Start();
                foreach (CurveLoop open in opens)
                {
                    XYZ point_A = GetPolygonMidPoint(open);
                    if (IsInsideOutline(point_A, c.CurveLoop))
                    {
                        Opening opening = doc.Create.NewOpening(floor1, ConvertCurveLoopToCurveArray(open), true);
                        break;
                    }
                }
                tf.Commit();


                //break;////////////////////////////////////////
            }
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

        public double UnitsToCentimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
        }

        public double CentimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Centimeters);
        }

        public List<CurveLoop> GetAllOutlines(Document doc, GeometryObject geoObj, GeometryElement geoElem)
        {
            XYZ direction = new XYZ(0, 0, 0);

            if (doc is null)
            {
                throw new ArgumentNullException(nameof(doc));
            }

            if (geoObj is null)
            {
                throw new ArgumentNullException(nameof(geoObj));
            }

            if (geoElem is null)
            {
                throw new ArgumentNullException(nameof(geoElem));
            }

            int count = 0;

            ElementId graphicsStyleId = null;
            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                graphicsStyleId = geoObj.GraphicsStyleId;
            }

            List<CurveLoop> allOutlines = new List<CurveLoop>();
            foreach (GeometryObject gObj in geoElem)
            {
                GeometryInstance geomInstance = gObj as GeometryInstance;
                Transform transform = geomInstance.Transform;
                if (null != geomInstance)
                {
                    foreach (GeometryObject insObj in geomInstance.SymbolGeometry)
                    {
                        bool isValid = true;
                        if (insObj.GraphicsStyleId.IntegerValue != graphicsStyleId.IntegerValue)
                            continue;

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.Line")
                            continue;

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            count++;
                            PolyLine polyLine = insObj as PolyLine;
                            List<XYZ> points_list = new List<XYZ>(polyLine.GetCoordinates());
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
                                line = TransformLine(transform, line);
                                prof.Append(line);

                            }

                            if (isValid)
                            {
                                allOutlines.Add(prof);
                            }
                        }
                    }
                }
            }
            return allOutlines;
        }

        private Autodesk.Revit.DB.Line TransformLine(Transform transform, Autodesk.Revit.DB.Line line)
        {
            XYZ startPoint = transform.OfPoint(line.GetEndPoint(0));
            XYZ endPoint = transform.OfPoint(line.GetEndPoint(1));
            Autodesk.Revit.DB.Line newLine = Autodesk.Revit.DB.Line.CreateBound(startPoint, endPoint);
            return newLine;
        }
        public XYZ GetPolygonMidPoint(CurveLoop curveLoop)
        {
            // Initiate some parameters.
            List<double> xCoorList = new List<double>();
            List<double> yCoorList = new List<double>();
            double zCoor = 0;

            // 
            foreach (Autodesk.Revit.DB.Curve curve in curveLoop.ToList())
            {
                xCoorList.Add(curve.GetEndPoint(0).X);
                xCoorList.Add(curve.GetEndPoint(1).X);
                yCoorList.Add(curve.GetEndPoint(0).Y);
                yCoorList.Add(curve.GetEndPoint(1).Y);
                zCoor = curve.GetEndPoint(0).Z;
            }

            // Sort the X and Y coordinate values.
            xCoorList.Sort();
            yCoorList.Sort();

            // Get the distinct values of sorted values.
            List<double> xSorted = GetDistinctList(xCoorList);
            List<double> ySorted = GetDistinctList(yCoorList);

            XYZ max = new XYZ(xSorted.Max(), ySorted.Max(), zCoor);
            XYZ min = new XYZ(xSorted.Min(), ySorted.Min(), zCoor);
            //XYZ midPoint = (max + min) / 2;
            XYZ midPoint = max + (min - max) / 100;

            return midPoint;
        }

        public XYZ GetPointInCurveLoop(CurveLoop curveLoop)
        {
            // Initiate some parameters.
            List<double> xCoorList = new List<double>();
            List<double> yCoorList = new List<double>();
            double zCoor = 0;

            // 
            foreach (Autodesk.Revit.DB.Curve curve in curveLoop.ToList())
            {
                xCoorList.Add(curve.GetEndPoint(0).X);
                xCoorList.Add(curve.GetEndPoint(1).X);
                yCoorList.Add(curve.GetEndPoint(0).Y);
                yCoorList.Add(curve.GetEndPoint(1).Y);
                zCoor = curve.GetEndPoint(0).Z;
            }

            // Sort the X and Y coordinate values.
            xCoorList.Sort();
            yCoorList.Sort();

            // Get the distinct values of sorted values.
            List<double> xSorted = GetDistinctList(xCoorList);
            List<double> ySorted = GetDistinctList(yCoorList);

            XYZ max = new XYZ(xSorted.Max(), ySorted.Max(), zCoor);
            XYZ min = new XYZ(xSorted.Min(), ySorted.Min(), zCoor);
            //XYZ midPoint = (max + min) / 2;
            XYZ midPoint = min + (max - min) / 100;

            return midPoint;
        }

        public XYZ GetPointOnCurveLoop(CurveLoop curveLoop)
        {
            List<Autodesk.Revit.DB.Curve> list_curve = curveLoop.ToList();
            return list_curve[0].GetEndPoint(0);
        }

        public List<SlabModel> GetLocation(List<SlabModel> SlabModels)
        {
            foreach (SlabModel s in SlabModels)
            {
                List<double> x = new List<double>();
                List<double> y = new List<double>();
                double z = 0;
                foreach (Autodesk.Revit.DB.Curve c in s.CurveLoop)
                {
                    x.Add(c.GetEndPoint(0).X);
                    x.Add(c.GetEndPoint(1).X);
                    y.Add(c.GetEndPoint(0).Y);
                    y.Add(c.GetEndPoint(1).Y);
                }
                XYZ midpoint = new XYZ((x.Max() + x.Min()) / 2, (y.Max() + y.Min()) / 2, z);
                x.Clear();
                y.Clear();
                s.Location = midpoint;
            }
            return SlabModels;
        }

        public List<SlabModel> GetPairedLabel(List<SlabModel> SlabModels, String path, String layer_name, GeometryElement geoElem)
        {
            Transform transform = null;
            foreach (GeometryObject gObj in geoElem)
            {
                GeometryInstance geomInstance = gObj as GeometryInstance;
                transform = geomInstance.Transform;
            }

            if (transform == null)
            {
                MessageBox.Show("Empty!");
            }

            //或許可以改成先把標籤的資訊與座標取出，存成一個List之後再進行配對。
            foreach (SlabModel slabModel in SlabModels)
            {
                double distanceBetweenTB = double.MaxValue;
                //double distanceBetweenTB_th = double.MaxValue;
                //double distanceBetweenTB_e = double.MaxValue;
                XYZ location_slab = slabModel.Location;

                using (new Services())
                {
                    using (Database database = new Database(false, false))
                    {
                        database.ReadDwgFile(path, FileShare.Read, true, "");
                        using (Teigha.DatabaseServices.Transaction trans1 = database.TransactionManager.StartTransaction())
                        {
                            using (BlockTable table = (BlockTable)database.BlockTableId.GetObject(OpenMode.ForRead))
                            {
                                using (SymbolTableEnumerator enumerator = table.GetEnumerator())
                                {
                                    enumerator.MoveNext();
                                    {
                                        using (BlockTableRecord record = (BlockTableRecord)enumerator.Current.GetObject(OpenMode.ForRead))
                                        {
                                            int count_l = 0;
                                            foreach (ObjectId id in record)
                                            {
                                                Entity entity2 = (Entity)id.GetObject(OpenMode.ForRead, false, false);
                                                switch (entity2.GetRXClass().Name)
                                                {
                                                    case "AcDbText":
                                                        Teigha.DatabaseServices.DBText text = (Teigha.DatabaseServices.DBText)entity2;
                                                        if (text.Layer == layer_name)
                                                        {
                                                            List<Autodesk.Revit.DB.Curve> s = slabModel.CurveLoop.ToList();
                                                            double z = s[0].GetEndPoint(0).Z;
                                                            XYZ loaction_label = ConverCADPointToRevitPoint(text.Position);
                                                            loaction_label = PointCentimeterToUnit(loaction_label);
                                                            loaction_label = new XYZ(loaction_label.X, loaction_label.Y, location_slab.Z);
                                                            XYZ loaction_label_new = transform.OfPoint(loaction_label);

                                                            if (IsInsideOutline(loaction_label_new, slabModel.CurveLoop))
                                                            {
                                                                string pattern_th = @"(\d+)\s*[+-±](\d+)";
                                                                Match match_th = Regex.Match(text.TextString, pattern_th);
                                                                if (match_th.Success)
                                                                {
                                                                    string c = match_th.Groups[1].Value;
                                                                    double thickness = double.Parse(c) / 10;
                                                                    if (thickness != 0)
                                                                    {
                                                                        slabModel.Thickness = thickness;
                                                                        count_l++;
                                                                    }
                                                                }

                                                                if (text.TextString.Contains("±")) { continue; }
                                                                else if (text.TextString.Contains("+"))
                                                                {
                                                                    string pattern = @"(\d+)\s*\+(\d+)";
                                                                    Match match = Regex.Match(text.TextString, pattern);
                                                                    if (match.Success)
                                                                    {
                                                                        string numbersString = match.Groups[2].Value;
                                                                        double elevation = double.Parse(numbersString) / 10;
                                                                        slabModel.Elevation = elevation; // mm to cm.
                                                                    }
                                                                }
                                                                else if (text.TextString.Contains("-"))
                                                                {

                                                                    string pattern = @"(\d+)\s*\-(\d+)";
                                                                    Match match = Regex.Match(text.TextString, pattern);
                                                                    if (match.Success)
                                                                    {
                                                                        string numbersString = match.Groups[2].Value;
                                                                        double elevation = double.Parse(numbersString) * (-1) / 10;
                                                                        slabModel.Elevation = elevation;
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    string c = text.TextString.ToString();
                                                                    double thickness = double.Parse(c) / 10;
                                                                    slabModel.Thickness = thickness;
                                                                    count_l++;
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                        break;

                                                    case "AcDbMText":
                                                        Teigha.DatabaseServices.MText text_m = (Teigha.DatabaseServices.MText)entity2;
                                                        if (text_m.Layer == layer_name)
                                                        {

                                                            List<Autodesk.Revit.DB.Curve> s = slabModel.CurveLoop.ToList();
                                                            double z = s[0].GetEndPoint(0).Z;
                                                            XYZ loaction_label = ConverCADPointToRevitPoint(text_m.Location);
                                                            loaction_label = PointCentimeterToUnit(loaction_label);
                                                            loaction_label = new XYZ(loaction_label.X, loaction_label.Y, location_slab.Z);

                                                            XYZ loaction_label_new = transform.OfPoint(loaction_label);
                                                            if (IsInsideOutline(loaction_label_new, slabModel.CurveLoop))
                                                            {
                                                                string pattern_th = @"(\d+)\s*[+-±](\d+)";
                                                                Match match_th = Regex.Match(text_m.Text, pattern_th);
                                                                if (match_th.Success)
                                                                {
                                                                    string c = match_th.Groups[1].Value;
                                                                    double thickness = double.Parse(c) / 10;
                                                                    if (thickness != 0)
                                                                    {
                                                                        slabModel.Thickness = thickness;
                                                                        count_l++;
                                                                    }

                                                                }

                                                                if (text_m.Text.Contains("±")) { continue; }
                                                                else if (text_m.Text.Contains("+"))
                                                                {
                                                                    string pattern = @"(\d+)\s*\+(\d+)";
                                                                    Match match = Regex.Match(text_m.Text, pattern);
                                                                    if (match.Success)
                                                                    {
                                                                        //MessageBox.Show("in");
                                                                        string numbersString = match.Groups[2].Value;
                                                                        double elevation = double.Parse(numbersString) / 10;
                                                                        slabModel.Elevation = elevation; // mm to cm.
                                                                    }
                                                                }
                                                                else if (text_m.Text.Contains("-"))
                                                                {
                                                                    string pattern = @"(\d+)\s*\-(\d+)";
                                                                    Match match = Regex.Match(text_m.Text, pattern);
                                                                    if (match.Success)
                                                                    {
                                                                        string numbersString = match.Groups[2].Value;
                                                                        double elevation = double.Parse(numbersString) * (-1) / 10;
                                                                        slabModel.Elevation = elevation;
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    string c = text_m.Text.ToString();
                                                                    double thickness = double.Parse(c) / 10;
                                                                    slabModel.Thickness = thickness;
                                                                    count_l++;
                                                                    break;
                                                                }
                                                            }
                                                        }
                                                        break;
                                                }
                                            }

                                            if (count_l == 0)
                                            {
                                                foreach (ObjectId id in record)
                                                {
                                                    Entity entity2 = (Entity)id.GetObject(OpenMode.ForRead, false, false);
                                                    switch (entity2.GetRXClass().Name)
                                                    {
                                                        case "AcDbText":
                                                            Teigha.DatabaseServices.DBText text = (Teigha.DatabaseServices.DBText)entity2;
                                                            if (text.Layer == layer_name)
                                                            {
                                                                List<Autodesk.Revit.DB.Curve> s = slabModel.CurveLoop.ToList();
                                                                double z = s[0].GetEndPoint(0).Z;
                                                                XYZ loaction_label = ConverCADPointToRevitPoint(text.Position);
                                                                loaction_label = PointCentimeterToUnit(loaction_label);
                                                                loaction_label = new XYZ(loaction_label.X, loaction_label.Y, location_slab.Z);

                                                                XYZ loaction_label_new = transform.OfPoint(loaction_label);


                                                                double distance_cad_text = loaction_label_new.DistanceTo(location_slab);


                                                                if (distance_cad_text < distanceBetweenTB)
                                                                {

                                                                    distanceBetweenTB = distance_cad_text;
                                                                    string pattern_th = @"(\d+)\s*[+-±](\d+)"; ;
                                                                    Match match_th = Regex.Match(text.TextString, pattern_th);

                                                                    if (Regex.IsMatch(text.TextString, @"[±+-]"))
                                                                    {
                                                                        string c = match_th.Groups[1].Value;
                                                                        string d = match_th.Groups[2].Value;
                                                                        double thickness = double.Parse(c) / 10;
                                                                        double elevation = double.Parse(d) / 10;
                                                                        if (thickness != 0)
                                                                        {
                                                                            slabModel.Thickness = thickness;
                                                                        }
                                                                        slabModel.Elevation = elevation;
                                                                        if (text.TextString.Contains("±")) { continue; }
                                                                        else if (text.TextString.Contains("+")) { continue; }
                                                                        else if (text.TextString.Contains("-"))
                                                                        {
                                                                            slabModel.Elevation = slabModel.Elevation * (-1);
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        string th = text.TextString;
                                                                        double thickness = double.Parse(th) / 10;
                                                                        if (thickness != 0)
                                                                        {
                                                                            slabModel.Thickness = thickness;
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            break;

                                                        case "AcDbMText":
                                                            Teigha.DatabaseServices.MText text_m = (Teigha.DatabaseServices.MText)entity2;
                                                            if (text_m.Layer == layer_name)
                                                            {
                                                                List<Autodesk.Revit.DB.Curve> s = slabModel.CurveLoop.ToList();
                                                                double z = s[0].GetEndPoint(0).Z;
                                                                XYZ loaction_label = ConverCADPointToRevitPoint(text_m.Location);
                                                                loaction_label = PointCentimeterToUnit(loaction_label);
                                                                loaction_label = new XYZ(loaction_label.X, loaction_label.Y, location_slab.Z);
                                                                XYZ loaction_label_new = transform.OfPoint(loaction_label);

                                                                double distance_cad_text = loaction_label_new.DistanceTo(location_slab);


                                                                if (distance_cad_text < distanceBetweenTB)
                                                                {
                                                                    distanceBetweenTB = distance_cad_text;
                                                                    string pattern_th = @"(\d+)\s*[+-±]\s*(\d+)"; ;
                                                                    Match match_th = Regex.Match(text_m.Text, pattern_th);
                                                                    if (Regex.IsMatch(text_m.Text, @"[±+-]"))
                                                                    {
                                                                        string c = match_th.Groups[1].Value;
                                                                        string d = match_th.Groups[2].Value;
                                                                        double thickness = double.Parse(c) / 10;
                                                                        double elevation = double.Parse(d) / 10;
                                                                        if (thickness != 0)
                                                                        {
                                                                            slabModel.Thickness = thickness;
                                                                        }
                                                                        slabModel.Elevation = elevation;
                                                                        if (text_m.Text.Contains("±")) { continue; }
                                                                        else if (text_m.Text.Contains("+")) { continue; }
                                                                        else if (text_m.Text.Contains("-"))
                                                                        {
                                                                            slabModel.Elevation = slabModel.Elevation * (-1);
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        string th = text_m.Text;
                                                                        double thickness = double.Parse(th) / 10;
                                                                        if (thickness != 0)
                                                                        {
                                                                            slabModel.Thickness = thickness;
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            break;
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
            return SlabModels;
        }

        public XYZ ConverCADPointToRevitPoint(Point3d point)
        {
            return new XYZ(point.X, point.Y, point.Z);
        }

        public string GetCADPath(ElementId cadLinkTypeID, Document revitDoc)
        {
            CADLinkType cadLinkType = revitDoc.GetElement(cadLinkTypeID) as CADLinkType;
            return ModelPathUtils.ConvertModelPathToUserVisiblePath(cadLinkType.GetExternalFileReference().GetAbsolutePath());
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

        public void CreateSlabType(Document doc, string name)
        {
            string pattern = @"(\d+)";
            double thickness = CentimetersToUnits(15);
            Match match = Regex.Match(name, pattern);
            if (match.Success)
            {
                string tn = match.Groups[1].Value;
                thickness = CentimetersToUnits(double.Parse(tn));
            }

            // Filter for the floor type
            FloorType floorType = new FilteredElementCollector(doc)
                .OfClass(typeof(FloorType))
                .FirstOrDefault(q => q.Name == "default") as FloorType; // Replace "default" with the name of your floor type

            if (floorType != null)
            {
                // Duplicate the floor type
                FloorType newFloorType = floorType.Duplicate(name) as FloorType;

                // Get the CompoundStructure of the new floor type
                CompoundStructure cs = newFloorType.GetCompoundStructure();

                // Change the thickness of the first layer
                if (cs != null && cs.LayerCount > 0)
                {
                    cs.SetLayerWidth(0, thickness); // Change 0.2 to your desired thickness
                    newFloorType.SetCompoundStructure(cs);
                }
            }
            else
            {
                MessageBox.Show("The specified floor type could not be found.");
            }
        }

    }
}