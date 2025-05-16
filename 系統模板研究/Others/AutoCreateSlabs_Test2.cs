using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;
using System.Diagnostics;
using Autodesk.Revit.DB.Structure;
using System.Windows;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(RegenerationOption.Manual)]
    public class AutoCreateSlabs_Test2 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            Reference refer = uidoc.Selection.PickObject(ObjectType.PointOnElement, "Select a CAD Layer");
            Element elem = doc.GetElement(refer);
            GeometryObject geoObj = elem.GetGeometryObjectFromReference(refer);
            GeometryElement geoElem = elem.get_Geometry(new Options());

            // Get the level of CAD drawing.
            Level level = doc.GetElement(elem.LevelId) as Level;

            // Start a TransactionGroup.
            TransactionGroup transGroup = new TransactionGroup(doc, "Start");
            transGroup.Start();

            // Get all the outline in layer of slab and copy them.
            List<CurveLoop> allOutlines = GetAllOutlines(doc, geoObj, geoElem);
            
            List<CurveLoop> allOutlines_copy = new List<CurveLoop>();
            foreach (CurveLoop loop in allOutlines)
            {
                allOutlines_copy.Add(loop);
            }
            

            // Classify the CurveLoops into opens and slabs.
            List<CurveLoop> slabOutlines = new List<CurveLoop>();
            List<CurveLoop> openOutlines = new List<CurveLoop>();
            for (int i = 0; i < allOutlines.Count(); i++)
            {
                //if (allOutlines[i].GetExactLength() < CentimetersToUnits(0))
                //{
                //    continue;
                //}
                CurveLoop outline_A = allOutlines[i];



                XYZ point_A = GetPolygonMidPoint(outline_A);
                allOutlines_copy.Remove(outline_A);
                Boolean isOpen = false;

                

                foreach (CurveLoop outline in allOutlines_copy)
                {
                    if (IsInsideOutline(point_A, outline))
                    {
                        openOutlines.Add(outline_A);
                        isOpen = true;
                        break;
                    }
                }
                if (!isOpen)
                {
                    slabOutlines.Add(outline_A);
                    allOutlines_copy.Add(outline_A);
                }
            }
            MessageBox.Show(slabOutlines.Count.ToString());
            MessageBox.Show(openOutlines.Count.ToString());
            // To create rectangular slabs
            List<CurveLoop> finalRecSlab = new List<CurveLoop>();
            foreach (CurveLoop slabOutline in slabOutlines)
            {
                if (openOutlines.Count() > 0)
                {
                    foreach (CurveLoop openOutline in openOutlines)
                    {
                        XYZ MidPoint_Open = GetPolygonMidPoint(openOutline);
                        if (IsInsideOutline(MidPoint_Open, slabOutline))
                        {
                            List<double> xCoorList = new List<double>();
                            List<double> yCoorList = new List<double>();

                            double zCoor = 0;

                            foreach (Curve curve in slabOutline.ToList())
                            {
                                xCoorList.Add(RoundPoint(curve.GetEndPoint(0)).X);
                                xCoorList.Add(RoundPoint(curve.GetEndPoint(1)).X);
                                yCoorList.Add(RoundPoint(curve.GetEndPoint(0)).Y);
                                yCoorList.Add(RoundPoint(curve.GetEndPoint(1)).Y);
                                zCoor = curve.GetEndPoint(0).Z;
                            }
                            foreach (Curve curve in openOutline.ToList())
                            {
                                xCoorList.Add(RoundPoint(curve.GetEndPoint(0)).X);
                                xCoorList.Add(RoundPoint(curve.GetEndPoint(1)).X);
                                yCoorList.Add(RoundPoint(curve.GetEndPoint(0)).Y);
                                yCoorList.Add(RoundPoint(curve.GetEndPoint(1)).Y);
                            }

                            xCoorList.Sort();
                            yCoorList.Sort();

                            List<double> xSorted = GetDistinctList(xCoorList);
                            List<double> ySorted = GetDistinctList(yCoorList);

                            List<CurveLoop> profile3 = new List<CurveLoop>();

                            //List<List<XYZ>> recsPoints = new List<List<XYZ>>();

                            for (int i = 0; i < xSorted.Count - 1; i++)
                            {
                                for (int j = 0; j < ySorted.Count - 1; j++)
                                {

                                    CurveLoop profileLoop = new CurveLoop();
                                    //List<XYZ> recPoints = new List<XYZ>();
                                    XYZ point1 = RoundPoint(new XYZ(xSorted[i], ySorted[j], zCoor));
                                    XYZ point2 = RoundPoint(new XYZ(xSorted[i + 1], ySorted[j], zCoor));
                                    XYZ point3 = RoundPoint(new XYZ(xSorted[i + 1], ySorted[j + 1], zCoor));
                                    XYZ point4 = RoundPoint(new XYZ(xSorted[i], ySorted[j + 1], zCoor));
                                    Line line1 = Line.CreateBound(point1, point2);
                                    Line line2 = Line.CreateBound(point2, point3);
                                    Line line3 = Line.CreateBound(point3, point4);
                                    Line line4 = Line.CreateBound(point4, point1);
                                    profileLoop.Append(line1);
                                    profileLoop.Append(line2);
                                    profileLoop.Append(line3);
                                    profileLoop.Append(line4);

                                    XYZ midPoint = (point1 + point3) / 2;
                                    if (IsInsideOutline(midPoint, slabOutline))
                                    {
                                        if (IsInsideOutline(midPoint, openOutline))
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            profile3.Add(profileLoop);
                                        }
                                    }
                                }
                            }
                            bool isAllSlabMerge = false;
                            while (!isAllSlabMerge)
                            {
                                List<CurveLoop> profile3_new = new List<CurveLoop>();
                                List<List<XYZ>> recsPoints_copy = new List<List<XYZ>>();
                                int count = 0;
                                for (int i = 0; i < profile3.Count; i++)
                                {
                                    for (int j = 0; j < profile3.Count; j++)
                                    {
                                        if (j <= i)
                                        {
                                            continue;
                                        }
                                        if (IsMergeable(profile3[i], profile3[j]))
                                        {
                                            profile3_new.Add(MergeTwoRectangles(profile3[i], profile3[j]));
                                            for (int k = 0; k < profile3.Count; k++)
                                            {
                                                if (k == i || k == j)
                                                {
                                                    continue;
                                                }
                                                else
                                                {
                                                    profile3_new.Add(profile3[k]);
                                                }
                                            }
                                            profile3 = profile3_new;
                                            count = 1;
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

                            foreach (CurveLoop c in profile3)
                            {
                                finalRecSlab.Add(c);
                            }
                            break;
                        }
                    }
                }
                else
                {
                    List<CurveLoop> recOutlines_s = GetRecSlabsFromPolySlab(slabOutline);
                    foreach (CurveLoop c in recOutlines_s)
                    {
                        //MessageBox.Show(finalRecSlab.Count.ToString());
                        finalRecSlab.Add(c);
                    }
                }
            }

            // Create slabs according to the outline in  "FinalRecSlab".
            CreateFloor(doc, finalRecSlab, level);

            transGroup.Assimilate();
            return Result.Succeeded;
        }

        public List<Line> CurveLoopToLineList(CurveLoop curveLoop)
        {
            List<Line> lineList = new List<Line>();
            List<Curve> curveList = curveLoop.ToList();
            foreach (Curve curve in curveList)
            {
                lineList.Add(Line.CreateBound(curve.GetEndPoint(0), curve.GetEndPoint(1)));
            }
            return lineList;
        }

        public void CreateFloor(Autodesk.Revit.DB.Document doc, List<CurveLoop> slabs, Level level)
        {
            foreach (CurveLoop c in slabs)
            {
                using (Transaction tx = new Transaction(doc))
                {
                    try
                    {
                        tx.Start("Create Slab.");
                        ElementId floorTypeId = Floor.GetDefaultFloorType(doc, false);
                        if (slabs != null)
                        {
                            Floor floor1 = Floor.Create(doc, new List<CurveLoop> { c }, floorTypeId, level.Id);
                        }
                        tx.Commit();
                    }
                    catch (Exception ex)
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
                        tx.RollBack();
                    }
                }
            }
        }

        public bool IsInsideOutline(XYZ TargetPoint, CurveLoop curveloop)
        {
            bool result = true;
            int insertCount = 0;
            List<Line> lines = CurveLoopToLineList(curveloop);
            Line rayLine = Line.CreateBound(TargetPoint, TargetPoint.Add(new XYZ(1, 0, 0) * 100000000));

            foreach (Line areaLine in lines)
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
                return result = false;
            }
            else
                return result;

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

        public bool IsMergeable(CurveLoop a, CurveLoop b)
        {
            List<XYZ> list_a = new List<XYZ>();
            List<XYZ> list_b = new List<XYZ>();

            foreach (Curve curve in a.ToList())
            {

                list_a.Add(curve.GetEndPoint(0));
            }

            foreach (Curve curve in b.ToList())
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

            foreach (Curve curve in a.ToList())
            {
                list_a.Add(curve.GetEndPoint(0));
            }

            foreach (Curve curve in b.ToList())
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
            Line line1 = Line.CreateBound(point1, point2);
            Line line2 = Line.CreateBound(point2, point3);
            Line line3 = Line.CreateBound(point3, point4);
            Line line4 = Line.CreateBound(point4, point1);
            profileLoop.Append(line1);
            profileLoop.Append(line2);
            profileLoop.Append(line3);
            profileLoop.Append(line4);
            return profileLoop;
        }
        public XYZ RoundPoint(XYZ point)
        {
            XYZ newPoint = new XYZ(
                Math.Round(point.X * 1000) / 1000,
                Math.Round(point.Y * 1000) / 1000,
                Math.Round(point.Z * 1000) / 1000
                );
            return newPoint;
        }

        public double CentimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Centimeters);
        }

        public List<CurveLoop> GetAllOutlines(Document doc, GeometryObject geoObj, GeometryElement geoElem)
        {
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
                        if (insObj.GraphicsStyleId.IntegerValue != graphicsStyleId.IntegerValue)
                            continue;

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.Line")
                            continue;

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> points = polyLine.GetCoordinates();
                            CurveLoop prof = new CurveLoop() as CurveLoop;

                            for (int i = 0; i < points.Count - 1; i++)
                            {
                                Line line = Line.CreateBound(points[i], points[i + 1]);
                                line = TransformLine(transform, line);
                                prof.Append(line);
                            }
                            allOutlines.Add(prof);
                        }
                    }
                }
            }
            return allOutlines;
        }

        private Line TransformLine(Transform transform, Line line)
        {
            XYZ startPoint = transform.OfPoint(line.GetEndPoint(0));
            XYZ endPoint = transform.OfPoint(line.GetEndPoint(1));
            Autodesk.Revit.DB.Line newLine = Autodesk.Revit.DB.Line.CreateBound(startPoint, endPoint);
            return newLine;
        }

        public List<CurveLoop> GetRecSlabsFromPolySlab(CurveLoop curveLoop)
        {
            // Initiate some parameters.
            List<double> xCoorList = new List<double>();
            List<double> yCoorList = new List<double>();
            double zCoor = 0;

            foreach (Curve curve in curveLoop.ToList())
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
                    Line line1 = Line.CreateBound(point1, point2);
                    Line line2 = Line.CreateBound(point2, point3);
                    Line line3 = Line.CreateBound(point3, point4);
                    Line line4 = Line.CreateBound(point4, point1);

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
            return profiles;
        }

        public XYZ GetPolygonMidPoint(CurveLoop curveLoop)
        {
            // Initiate some parameters.
            List<double> xCoorList = new List<double>();
            List<double> yCoorList = new List<double>();
            double zCoor = 0;

            // 
            foreach (Curve curve in curveLoop.ToList())
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
            XYZ midPoint = (max + min) / 2;

            return midPoint;
        }


    }

    public class Polygon
    {
        public Polygon()
        {
            Slab = null;
            SubSlab = new List<CurveLoop>();
            Open = new List<CurveLoop>();
        }
        /// <summary>
        /// CurveLoop
        /// </summary>
        public CurveLoop Slab { get; set; }

        /// <summary>
        /// SubSbab
        /// </summary>
        public List<CurveLoop> SubSlab { get; set; }

        /// <summary>
        /// Open
        /// </summary>
        public List<CurveLoop> Open { get; set; }

    }
}
