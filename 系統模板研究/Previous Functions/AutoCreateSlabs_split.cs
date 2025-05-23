﻿using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;
using System.Diagnostics;
using Autodesk.Revit.DB.Structure;
using System.Windows;
using System.IO;
using Teigha.DatabaseServices;
using Teigha.Runtime;
using Teigha.Geometry;
using System.Text.RegularExpressions;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(RegenerationOption.Manual)]
    public class AutoCreateSlabs_split : IExternalCommand
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
            ModelingParam.Initialize();
            double gridline_size = ModelingParam.parameters.General.GridSize * 10; // unit: mm

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

            //MessageBox.Show(allOutlines.Count.ToString());

            // Classify the CurveLoops into opens and slabs.
            List<CurveLoop> slabOutlines = new List<CurveLoop>();
            List<CurveLoop> openOutlines = new List<CurveLoop>();
            for (int i = 0; i < allOutlines.Count(); i++)
            {
                CurveLoop outline_A = allOutlines[i];
                XYZ point_A = GetPolygonMidPoint(outline_A);
                allOutlines_copy.Remove(outline_A);
                Boolean isOpen = false;

                // Do not consider the label's curveLoop.
                //if (allOutlines[i].GetExactLength() < CentimetersToUnits(162))
                //{
                //    continue;
                //}   
                //if (allOutlines[i].I)

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

            //Create SlabModel to store information of each slab.
            List<SlabModel> slabModels = new List<SlabModel>();
            foreach (CurveLoop curveLoop in slabOutlines)
            {
                SlabModel s = new SlabModel
                {
                    CurveLoop = curveLoop
                };
                slabModels.Add(s);
            }
            //MessageBox.Show(slabModels.Count.ToString());
            //MessageBox.Show(openOutlines.Count.ToString());

            slabModels = GetLocation(slabModels);
            //CreateColumn(doc, slabModels, level);
            //MessageBox.Show(slabModels.Count.ToString());
            //foreach(SlabModel s in slabModels)
            //{
            //    MessageBox.Show(s.Location.ToString());
            //}

            // Pair slabs with labels by distances.
            string path = GetCADPath(elem.GetTypeId(), doc);
            List<SlabModel> slabModels_paired = GetPairedLabel(slabModels, path, layer_name);
            //MessageBox.Show(slabModels_paired.Count.ToString());
            //foreach (SlabModel s in slabModels_paired)
            //{
            //    MessageBox.Show(s.Thickness.ToString());
            //}

            // To create rectangular slabs
            List<SlabModel> finalRecSlab = new List<SlabModel>();

            foreach (SlabModel slabOutline in slabModels_paired)
            {
                if (openOutlines.Count() < 0)
                {
                    foreach (CurveLoop openOutline in openOutlines)
                    {
                        XYZ MidPoint_Open = GetPolygonMidPoint(openOutline);
                        if (IsInsideOutline(MidPoint_Open, slabOutline.CurveLoop))
                        {
                            List<double> xCoorList = new List<double>();
                            List<double> yCoorList = new List<double>();

                            double zCoor = 0;

                            foreach (Autodesk.Revit.DB.Curve curve in slabOutline.CurveLoop.ToList())
                            {
                                xCoorList.Add(Algorithm.RoundPoint(curve.GetEndPoint(0), gridline_size).X);
                                xCoorList.Add(Algorithm.RoundPoint(curve.GetEndPoint(1), gridline_size).X);
                                yCoorList.Add(Algorithm.RoundPoint(curve.GetEndPoint(0), gridline_size).Y);
                                yCoorList.Add(Algorithm.RoundPoint(curve.GetEndPoint(1), gridline_size).Y);
                                zCoor = curve.GetEndPoint(0).Z;
                            }
                            foreach (Autodesk.Revit.DB.Curve curve in openOutline.ToList())
                            {
                                xCoorList.Add(Algorithm.RoundPoint(curve.GetEndPoint(0), gridline_size).X);
                                xCoorList.Add(Algorithm.RoundPoint(curve.GetEndPoint(1), gridline_size).X);
                                yCoorList.Add(Algorithm.RoundPoint(curve.GetEndPoint(0), gridline_size).Y);
                                yCoorList.Add(Algorithm.RoundPoint(curve.GetEndPoint(1), gridline_size).Y);
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
                                    XYZ point1 = Algorithm.RoundPoint(new XYZ(xSorted[i], ySorted[j], zCoor), gridline_size);
                                    XYZ point2 = Algorithm.RoundPoint(new XYZ(xSorted[i + 1], ySorted[j], zCoor), gridline_size);
                                    XYZ point3 = Algorithm.RoundPoint(new XYZ(xSorted[i + 1], ySorted[j + 1], zCoor), gridline_size);
                                    XYZ point4 = Algorithm.RoundPoint(new XYZ(xSorted[i], ySorted[j + 1], zCoor), gridline_size);
                                    Autodesk.Revit.DB.Line line1 = Autodesk.Revit.DB.Line.CreateBound(point1, point2);
                                    Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(point2, point3);
                                    Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateBound(point3, point4);
                                    Autodesk.Revit.DB.Line line4 = Autodesk.Revit.DB.Line.CreateBound(point4, point1);
                                    profileLoop.Append(line1);
                                    profileLoop.Append(line2);
                                    profileLoop.Append(line3);
                                    profileLoop.Append(line4);

                                    XYZ midPoint = (point1 + point3) / 2;
                                    if (IsInsideOutline(midPoint, slabOutline.CurveLoop))
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
                                SlabModel s = new SlabModel
                                {
                                    CurveLoop = c,
                                    Thickness = slabOutline.Thickness,
                                    Elevation = slabOutline.Elevation
                                };
                                finalRecSlab.Add(s);
                            }
                            break;
                        }
                    }
                }
                else
                {
                    List<CurveLoop> recOutlines_s = GetRecSlabsFromPolySlab(slabOutline.CurveLoop, gridline_size);
                    foreach (CurveLoop c in recOutlines_s)
                    {
                        SlabModel slabModel = new SlabModel
                        {
                            CurveLoop = c,
                            Thickness = slabOutline.Thickness,
                            Elevation = slabOutline.Elevation
                        };
                        finalRecSlab.Add(slabModel);
                    }
                }
            }

            CreateFloor(doc, finalRecSlab, level);

            transGroup.Assimilate();
            return Result.Succeeded;
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

        public void CreateFloor(Autodesk.Revit.DB.Document doc, List<SlabModel> slabs, Level level)
        {
            foreach (SlabModel c in slabs)
            {
                using (Autodesk.Revit.DB.Transaction tx = new Autodesk.Revit.DB.Transaction(doc))
                {
                    try
                    {
                        tx.Start("Create Slab.");
                        string familyName = "RC slab(" + c.Thickness.ToString() + ")";
                        FloorType floorType = new FilteredElementCollector(doc)
                        .OfClass(typeof(FloorType))
                        .OfCategory(BuiltInCategory.OST_Floors)
                        .Cast<FloorType>()
                        .FirstOrDefault(q => q.Name == familyName) as FloorType;

                        ElementId floorTypeId = floorType.Id;

                        if (slabs != null)
                        {
                            Floor floor1 = Floor.Create(doc, new List<CurveLoop> { c.CurveLoop }, floorTypeId, level.Id);
                            Parameter distanceToLevelParam = floor1.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                            distanceToLevelParam.Set(CentimetersToUnits(c.Elevation));
                        }
                        tx.Commit();
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
                        tx.RollBack();
                    }
                }
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

        public bool IsMergeable(CurveLoop a, CurveLoop b)
        {
            List<XYZ> list_a = new List<XYZ>();
            List<XYZ> list_b = new List<XYZ>();

            foreach (Autodesk.Revit.DB.Curve curve in a.ToList())
            {

                list_a.Add(curve.GetEndPoint(0));
            }

            foreach (Autodesk.Revit.DB.Curve curve in b.ToList())
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

            foreach (Autodesk.Revit.DB.Curve curve in a.ToList())
            {
                list_a.Add(curve.GetEndPoint(0));
            }

            foreach (Autodesk.Revit.DB.Curve curve in b.ToList())
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
                                if(points[i].DistanceTo(points[i + 1]) > Algorithm.CentimetersToUnits(0.1))
                                {
                                    Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(points[i], points[i + 1]);
                                    line = TransformLine(transform, line);
                                    prof.Append(line);
                                }
                            }
                            allOutlines.Add(prof);
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

        public List<CurveLoop> GetRecSlabsFromPolySlab(CurveLoop curveLoop, double gridline_size)
        {
            // Initiate some parameters.
            List<double> xCoorList = new List<double>();
            List<double> yCoorList = new List<double>();
            double zCoor = 0;

            foreach (Autodesk.Revit.DB.Curve curve in curveLoop.ToList())
            {
                xCoorList.Add(Algorithm.RoundPoint(curve.GetEndPoint(0), gridline_size).X);
                xCoorList.Add(Algorithm.RoundPoint(curve.GetEndPoint(1), gridline_size).X);
                yCoorList.Add(Algorithm.RoundPoint(curve.GetEndPoint(0), gridline_size).Y);
                yCoorList.Add(Algorithm.RoundPoint(curve.GetEndPoint(1), gridline_size).Y);
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
                    Autodesk.Revit.DB.Line line1 = Autodesk.Revit.DB.Line.CreateBound(point1, point2);
                    Autodesk.Revit.DB.Line line2 = Autodesk.Revit.DB.Line.CreateBound(point2, point3);
                    Autodesk.Revit.DB.Line line3 = Autodesk.Revit.DB.Line.CreateBound(point3, point4);
                    Autodesk.Revit.DB.Line line4 = Autodesk.Revit.DB.Line.CreateBound(point4, point1);

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

        public List<SlabModel> GetPairedLabel(List<SlabModel> SlabModels, String path, String layer_name)
        {
            //或許可以改成先把標籤的資訊與座標取出，存成一個List之後再進行配對。
            foreach (SlabModel slabModel in SlabModels)
            {
                double distanceBetweenTB = double.MaxValue;
                double distanceBetweenMB = double.MaxValue;
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
                                                            XYZ loaction_label = PointMilimeterToUnit(ConverCADPointToRevitPoint(text.Position));
                                                            List<Autodesk.Revit.DB.Curve> s = slabModel.CurveLoop.ToList();
                                                            double z = s[0].GetEndPoint(0).Z;
                                                            XYZ loaction_label_new = new XYZ(loaction_label.X, loaction_label.Y, z);
                                                            if (IsInsideOutline(loaction_label_new, slabModel.CurveLoop))
                                                            {
                                                                
                                                                if (text.TextString.Contains("±")) { continue; }
                                                                else if (text.TextString.Contains("+"))
                                                                {
                                                                    string pattern = @"+(\d+)";
                                                                    Match match = Regex.Match(text.TextString, pattern);
                                                                    if (match.Success)
                                                                    {
                                                                        string numbersString = match.Groups[1].Value;
                                                                        double elevation = double.Parse(numbersString) / 10;
                                                                        slabModel.Elevation = elevation; // mm to cm.
                                                                    }
                                                                }
                                                                else if (text.TextString.Contains("-"))
                                                                {
                                                                    
                                                                    string pattern = @"-(\d+)";
                                                                    Match match = Regex.Match(text.TextString, pattern);
                                                                    if (match.Success)
                                                                    {
                                                                        //MessageBox.Show("in");
                                                                        string numbersString = match.Groups[1].Value;
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

                                                            XYZ loaction_label = PointMilimeterToUnit(ConverCADPointToRevitPoint(text_m.Location));
                                                            List<Autodesk.Revit.DB.Curve> s = slabModel.CurveLoop.ToList();
                                                            double z = s[0].GetEndPoint(0).Z;
                                                            XYZ loaction_label_new = new XYZ(loaction_label.X, loaction_label.Y, z);
                                                            if (IsInsideOutline(loaction_label_new, slabModel.CurveLoop))
                                                            {
                                                                
                                                                if (text_m.Text.Contains("±")) { continue; }
                                                                else if (text_m.Text.Contains("+"))
                                                                {
                                                                    string pattern = @"+(\d+)";
                                                                    Match match = Regex.Match(text_m.Text, pattern);
                                                                    if (match.Success)
                                                                    {
                                                                        //MessageBox.Show("in");
                                                                        string numbersString = match.Groups[1].Value;
                                                                        double elevation = double.Parse(numbersString) / 10;
                                                                        slabModel.Elevation = elevation; // mm to cm.
                                                                    }
                                                                }
                                                                else if (text_m.Text.Contains("-"))
                                                                {
                                                                    string pattern = @"-(\d+)";
                                                                    Match match = Regex.Match(text_m.Text, pattern);
                                                                    if (match.Success)
                                                                    {
                                                                        string numbersString = match.Groups[1].Value;
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
                                                                double distance_cad_text = PointMilimeterToUnit(ConverCADPointToRevitPoint(text.Position)).DistanceTo(location_slab);

                                                                if (text.TextString.Contains("±")) { continue; }
                                                                else if (text.TextString.Contains("+"))
                                                                {
                                                                    if (distance_cad_text < distanceBetweenMB)
                                                                    {
                                                                        string pattern = @"+(\d+)";
                                                                        Match match = Regex.Match(text.TextString, pattern);
                                                                        if (match.Success)
                                                                        {
                                                                           
                                                                            string numbersString = match.Groups[1].Value;
                                                                            int elevation = int.Parse(numbersString) / 10;
                                                                            slabModel.Elevation = elevation; // mm to cm.
                                                                        }
                                                                    }
                                                                }
                                                                else if (text.TextString.Contains("-"))
                                                                {
                                                                    if (distance_cad_text < distanceBetweenMB)
                                                                    {
                                                                        string pattern = @"-(\d+)";
                                                                        Match match = Regex.Match(text.TextString, pattern);
                                                                        if (match.Success)
                                                                        {
                                                                            //MessageBox.Show("in");
                                                                            string numbersString = match.Groups[1].Value;
                                                                            int elevation = int.Parse(numbersString) * (-1) / 10;
                                                                            slabModel.Elevation = elevation;

                                                                            //MessageBox.Show(slabModel.Elevation.ToString());

                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    if (distance_cad_text < distanceBetweenTB)
                                                                    {
                                                                        distanceBetweenTB = distance_cad_text;
                                                                        string c = text.TextString;
                                                                        double thickness = double.Parse(c) / 10;
                                                                        slabModel.Thickness = thickness;
                                                                    }
                                                                }


                                                            }
                                                            break;

                                                        case "AcDbMText":
                                                            Teigha.DatabaseServices.MText text_m = (Teigha.DatabaseServices.MText)entity2;
                                                            if (text_m.Layer == layer_name)
                                                            {
                                                                double distance_cad_text = PointMilimeterToUnit(ConverCADPointToRevitPoint(text_m.Location)).DistanceTo(location_slab);

                                                                if (text_m.Text.Contains("±")) { continue; }
                                                                else if (text_m.Text.Contains("+"))
                                                                {
                                                                    if (distance_cad_text < distanceBetweenTB)
                                                                    {
                                                                        string pattern = @"+(\d+)";
                                                                        Match match = Regex.Match(text_m.Text, pattern);
                                                                        if (match.Success)
                                                                        {
                                                                            string numbersString = match.Groups[1].Value;
                                                                            int elevation = int.Parse(numbersString) / 10;
                                                                            slabModel.Elevation = elevation;
                                                                        }
                                                                    }
                                                                }
                                                                else if (text_m.Text.Contains("-"))
                                                                {
                                                                    if (distance_cad_text < distanceBetweenTB)
                                                                    {
                                                                        string pattern = @"-(\d+)";
                                                                        Match match = Regex.Match(text_m.Text, pattern);
                                                                        if (match.Success)
                                                                        {
                                                                            //MessageBox.Show("in");
                                                                            string numbersString = match.Groups[1].Value;
                                                                            int elevation = int.Parse(numbersString) * (-1) / 10;
                                                                            slabModel.Elevation = elevation;
                                                                        }
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    if (distance_cad_text < distanceBetweenTB)
                                                                    {
                                                                        distanceBetweenTB = distance_cad_text;
                                                                        string c = text_m.Text;
                                                                        double thickness = double.Parse(c) / 10;
                                                                        slabModel.Thickness = thickness;
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


        public XYZ PointMilimeterToUnit(XYZ point)
        {
            XYZ newPoint = new XYZ(
                CentimetersToUnits(point.X),
                CentimetersToUnits(point.Y),
                CentimetersToUnits(point.Z)
                );
            return newPoint;
        }

    }
    public class SlabModel
    {
        public SlabModel()
        {
            CurveLoop = null;
            Open = null;
            Thickness = 15.0;
            Location = new XYZ(0, 0, 0);
            Elevation = 0;
        }

        /// <summary>
        /// Thickness
        /// </summary>
        public double Thickness { get; set; }

        /// <summary>
        /// Elevation
        /// </summary>
        public double Elevation { get; set; }

        /// <summary>
        /// CurveLoop
        /// </summary>
        public CurveLoop CurveLoop { get; set; }

        /// <summary>
        /// location
        /// </summary>
        public XYZ Location { get; set; }

        public CurveArray Open{get; set; }

    }
}