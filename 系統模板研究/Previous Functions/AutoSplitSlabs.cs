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
using System.IO;
using Teigha.DatabaseServices;
using Teigha.Runtime;
using Teigha.Geometry;
using System.Windows.Media.Animation;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(RegenerationOption.Manual)]
    public class AutoSplitSlabs : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Autodesk.Revit.DB.Document doc = uidoc.Document;

            

            // Get the level of CAD drawing.
            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("5F")) as Level;

            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            List<Floor> slabs = new List<Floor>();

            
            foreach(ElementId id in ids)
            {
                slabs.Add(doc.GetElement(id) as Floor);
            }

            List<CurveLoop> allOutlines = new List<CurveLoop>();


            foreach (Element slab in slabs)
            {
                GeometryElement geometryElement = slab.get_Geometry(new Options());
                foreach (GeometryObject geometryObject in geometryElement)
                {
                    Autodesk.Revit.DB.Solid solid = geometryObject as Autodesk.Revit.DB.Solid;
                    if (solid != null)
                    {
                        foreach (Autodesk.Revit.DB.Face face in solid.Faces)
                        {
                            BoundingBoxUV bbxuv = face.GetBoundingBox();
                            if (XYZ.BasisZ.DotProduct(face.ComputeNormal(bbxuv.Max)) == 1)
                            {
                                CurveLoop curveLoop = face.GetEdgesAsCurveLoops()[0];
                                allOutlines.Add(curveLoop);
                            }
                        }
                    }
                }
            }


            MessageBox.Show(allOutlines.Count.ToString());

            // Start a TransactionGroup.
            TransactionGroup transGroup = new TransactionGroup(doc, "Start");
            transGroup.Start();

            // Get all the outline in layer of slab and copy them.
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
            // MessageBox.Show(slabModels.Count.ToString()); 
            //MessageBox.Show(openOutlines.Count.ToString());

            slabModels = GetLocation(slabModels);
            //CreateColumn(doc, slabModels, level);
            //MessageBox.Show(slabModels.Count.ToString());
            //foreach(SlabModel s in slabModels)
            //{
            //    MessageBox.Show(s.Location.ToString());
            //}

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

            Autodesk.Revit.DB.Transaction t_d = new Autodesk.Revit.DB.Transaction(doc, "Delete Original Slabs");
            t_d.Start();
            foreach (ElementId id in ids)
            {
                doc.Delete(id);
            }
            t_d.Commit();


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
                if (openOutlines.Count() > 0)
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
                                xCoorList.Add(RoundPoint(curve.GetEndPoint(0)).X);
                                xCoorList.Add(RoundPoint(curve.GetEndPoint(1)).X);
                                yCoorList.Add(RoundPoint(curve.GetEndPoint(0)).Y);
                                yCoorList.Add(RoundPoint(curve.GetEndPoint(1)).Y);
                                zCoor = curve.GetEndPoint(0).Z;
                            }
                            foreach (Autodesk.Revit.DB.Curve curve in openOutline.ToList())
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
                                    Thickness = slabOutline.Thickness
                                };
                                finalRecSlab.Add(s);
                            }
                            break;
                        }
                    }
                }
                else
                {
                    List<CurveLoop> recOutlines_s = GetRecSlabsFromPolySlab(slabOutline.CurveLoop);
                    foreach (CurveLoop c in recOutlines_s)
                    {
                        SlabModel slabModel = new SlabModel
                        {
                            CurveLoop = c,
                            Thickness = slabOutline.Thickness
                        };
                        finalRecSlab.Add(slabModel);
                    }
                }
            }


            //foreach (SlabModel s in finalRecSlab)
            //{
            //    MessageBox.Show(s.Thickness.ToString());
            //}

            // Create slabs according to the outline in  "FinalRecSlab".
            //MessageBox.Show(finalRecSlab.Count.ToString());
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

        public XYZ RoundPoint(XYZ point)
        {

            XYZ newPoint = new XYZ(
                CentimetersToUnits(Math.Round(UnitsToCentimeters(point.X) * 2) / 2),
                CentimetersToUnits(Math.Round(UnitsToCentimeters(point.Y) * 2) / 2),
                CentimetersToUnits(Math.Round(UnitsToCentimeters(point.Z) * 2) / 2)
                );
            return newPoint;
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
                                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(points[i], points[i + 1]);
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

        private Autodesk.Revit.DB.Line TransformLine(Transform transform, Autodesk.Revit.DB.Line line)
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

            foreach (Autodesk.Revit.DB.Curve curve in curveLoop.ToList())
            {
                xCoorList.Add(RoundPoint(curve.GetEndPoint(0)).X);
                xCoorList.Add(RoundPoint(curve.GetEndPoint(1)).X);
                yCoorList.Add(RoundPoint(curve.GetEndPoint(0)).Y);
                yCoorList.Add(RoundPoint(curve.GetEndPoint(1)).Y);
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
            XYZ midPoint = (max + min) / 2;

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

        public void CreateColumn(Document doc, List<SlabModel> slabModels, Level level)
        {
            foreach (SlabModel s in slabModels)
            {
                using (Autodesk.Revit.DB.Transaction tx = new Autodesk.Revit.DB.Transaction(doc))
                {
                    try
                    {
                        tx.Start("createColumn");
                        //if (!default_column.IsActive)
                        //{
                        //    default_column.Activate();
                        //}
                        FamilySymbol default_column = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol))
                        .OfCategory(BuiltInCategory.OST_Columns)
                        .Cast<FamilySymbol>()
                        .FirstOrDefault(q => q.Name == "default") as FamilySymbol;
                        FamilyInstance familyInstance = doc.Create.NewFamilyInstance(s.Location, default_column, level, StructuralType.Column);
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

        public List<SlabModel> GetPairedLabel(List<SlabModel> SlabModels, String path, String layer_name)
        {
            //或許可以改成先把標籤的資訊與座標取出，存成一個List之後再進行配對。
            foreach (SlabModel slabModel in SlabModels)
            {
                double distanceBetweenTB = double.MaxValue;
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
                                                            //double distance_cad_text = PointMilimeterToUnit(ConverCADPointToRevitPoint(text.Position)).DistanceTo(location_slab);
                                                            //if (distance_cad_text < distanceBetweenTB)
                                                            //{
                                                            //    {
                                                            //        if (!text.TextString.Contains("±"))
                                                            //        {
                                                            //            distanceBetweenTB = distance_cad_text;
                                                            //            string c = text.TextString.ToString();
                                                            //            double thickness = double.Parse(c) / 10;
                                                            //            slabModel.Thickness = thickness;

                                                            //        }
                                                            //        else
                                                            //        {
                                                            //            continue;
                                                            //        }

                                                            //    }
                                                            //}

                                                            XYZ loaction_label = PointMilimeterToUnit(ConverCADPointToRevitPoint(text.Position));
                                                            if (IsInsideOutline(location_slab, slabModel.CurveLoop))
                                                            {
                                                                string c = text.TextString.ToString();
                                                                double thickness = double.Parse(c) / 10;
                                                                slabModel.Thickness = thickness;
                                                                count_l++;
                                                                break;
                                                            }
                                                        }
                                                        break;

                                                    case "AcDbMText":
                                                        Teigha.DatabaseServices.MText text_m = (Teigha.DatabaseServices.MText)entity2;
                                                        if (text_m.Layer == layer_name)
                                                        {
                                                            //double distance_cad_text = PointMilimeterToUnit(ConverCADPointToRevitPoint(text_m.Location)).DistanceTo(location_slab);
                                                            //if (distance_cad_text < distanceBetweenTB)
                                                            //{
                                                            //    {
                                                            //        if (!text_m.Text.Contains("±"))
                                                            //        {
                                                            //            distanceBetweenTB = distance_cad_text;
                                                            //            string d = text_m.Text.ToString();
                                                            //            double thickness = double.Parse(d) / 10;
                                                            //            slabModel.Thickness = thickness;
                                                            //            //MessageBox.Show(distanceBetweenTB.ToString());
                                                            //            //MessageBox.Show(PointMilimeterToUnit(ConverCADPointToRevitPoint(text_m.Location)).ToString());
                                                            //                //MessageBox.Show(location_slab.ToString());
                                                            //        }
                                                            //        else
                                                            //        {
                                                            //            continue;
                                                            //        }
                                                            //    }
                                                            //}

                                                            XYZ loaction_label = PointMilimeterToUnit(ConverCADPointToRevitPoint(text_m.Location));
                                                            if (IsInsideOutline(location_slab, slabModel.CurveLoop))
                                                            {
                                                                string c = text_m.Text.ToString();
                                                                double thickness = double.Parse(c) / 10;
                                                                slabModel.Thickness = thickness;
                                                                count_l++;
                                                                break;
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
                                                                if (distance_cad_text < distanceBetweenTB)
                                                                {
                                                                    {
                                                                        if (text.TextString.Contains("±") || text.TextString.Contains("+")
                                                                            || text.TextString.Contains("-"))
                                                                        {
                                                                            // To be coded.
                                                                            continue;

                                                                        }
                                                                        else
                                                                        {
                                                                            distanceBetweenTB = distance_cad_text;
                                                                            string c = text.TextString.ToString();
                                                                            double thickness = double.Parse(c) / 10;
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
                                                                double distance_cad_text = PointMilimeterToUnit(ConverCADPointToRevitPoint(text_m.Location)).DistanceTo(location_slab);
                                                                if (distance_cad_text < distanceBetweenTB)
                                                                {
                                                                    {
                                                                        if (text_m.Text.ToString().Contains("±") || text_m.Text.ToString().Contains("+")
                                                                            || text_m.Text.ToString().Contains("-"))
                                                                        {
                                                                            // To be coded.
                                                                            continue;

                                                                        }
                                                                        else
                                                                        {
                                                                            distanceBetweenTB = distance_cad_text;
                                                                            string c = text_m.Text.ToString();
                                                                            double thickness = double.Parse(c) / 10;
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
}
