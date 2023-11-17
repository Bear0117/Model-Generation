using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class ConvertToFormworkPlanningWall : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            FilteredElementCollector collector = new FilteredElementCollector(doc).OfClass(typeof(Level));
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds(); // 將所有框選到的元件的ID，放入ids這個容器裡。
            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("4F")) as Level;

            double levelHeight = CentimetersToUnits(330); // default
            List<double> levelDistancesList = new List<double>();

            // retrieve all the Level elements in the document
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
                LevelHeightForm form = new LevelHeightForm(doc);
                form.ShowDialog();
                levelHeight = CentimetersToUnits(form.levelHeight);
            }

            foreach (ElementId id in ids)
            {
                Wall wall = doc.GetElement(id) as Wall;

                Transaction t1 = new Transaction(doc, "Create Wall");
                t1.Start();
                Parameter wallHeightP = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                wallHeightP.Set(levelHeight);
                WallUtils.DisallowWallJoinAtEnd(wall, 0);
                WallUtils.DisallowWallJoinAtEnd(wall, 1);
                t1.Commit();
                //ChangeWallType(doc, wall, Line1, Line2);

                List<double> depthList = new List<double>();

                //Get the element by elementID and get the boundingbox and outline of this element.
                Element elementWall = wall as Element;
                BoundingBoxXYZ bbxyzElement = elementWall.get_BoundingBox(null);
                Outline outline = new Outline(bbxyzElement.Min, bbxyzElement.Max);
                Outline newOutline = ReducedOutline_s(bbxyzElement);

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
                        string elementName = inter[i].Category.Name;
                        if (elementName == "結構構架")
                        {
                            // Find the depth
                            BoundingBoxXYZ bbxyzBeam = inter[i].get_BoundingBox(null);
                            depthList.Add(Math.Abs(bbxyzBeam.Min.Z - bbxyzBeam.Max.Z));

                            // Join
                            Transaction joinAndSwitch = new Transaction(doc, "Join and Switch");
                            joinAndSwitch.Start();
                            RunJoinGeometryAndSwitch(doc, wall, inter[i]);
                            joinAndSwitch.Commit();
                        }
                        else if (elementName == "樓板")
                        {
                            BoundingBoxXYZ bbxyzSlab = inter[i].get_BoundingBox(null);
                            depthList.Add(Math.Abs(bbxyzSlab.Min.Z - bbxyzSlab.Max.Z));

                            Transaction join = new Transaction(doc, "Join");
                            join.Start();
                            RunJoinGeometry(doc, wall, inter[i]);
                            join.Commit();
                        }
                        else continue;
                    }
                }

                LocationCurve wallLC = wall.Location as LocationCurve;
                Curve wallCurve = wallLC.Curve;
                Line wallLine = wallCurve as Line;
                XYZ lineAxis = wallLine.GetEndPoint(0) - wallLine.GetEndPoint(1);
                List<double> faceZ = new List<double>();
                List<Solid> list_solid = GetElementSolidList(wall);
                List<Face> list_vFace = new List<Face>();
                FaceArray faceArray;
                foreach (Solid solid in list_solid)
                {
                    faceArray = solid.Faces;
                    foreach (Face face in faceArray)
                    {
                        BoundingBoxUV boxUV = face.GetBoundingBox();
                        XYZ min = face.Evaluate(boxUV.Min);
                        XYZ max = face.Evaluate(boxUV.Max);
                        faceZ.Add(min.Z);
                        faceZ.Add(max.Z);
                        if (Math.Abs(face.ComputeNormal(face.GetBoundingBox().Max).CrossProduct(lineAxis).GetLength()) < CentimetersToUnits(0.1))
                        {
                            list_vFace.Add(face);
                        }
                    }
                    break;
                }

                // Delete the original wall
                Transaction t_deletewall = new Transaction(doc, "Delete Wall");
                t_deletewall.Start();
                doc.Delete(wall.Id);
                t_deletewall.Commit();

                List<Line> lines = LineSeperatedByFaces(wallLine, list_vFace);

                foreach (Line line_new in lines)
                {
                    Transaction t1_new = new Transaction(doc, "Create Wall");
                    t1_new.Start();
                    Wall wall_new = Wall.Create(doc, line_new, level.Id, true);
                    Parameter wallHeightP_new = wall_new.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                    wallHeightP_new.Set(levelHeight);
                    WallUtils.DisallowWallJoinAtEnd(wall_new, 0);
                    WallUtils.DisallowWallJoinAtEnd(wall_new, 1);
                    t1_new.Commit();



                    List<double> depthList_new = new List<double>();
                    //Get the element by elementID and get the boundingbox and outline of this element.
                    Element elementWall_new = wall_new as Element;
                    BoundingBoxXYZ bbxyzElement_new = elementWall_new.get_BoundingBox(null);
                    Outline outline_new = new Outline(bbxyzElement_new.Min, bbxyzElement_new.Max);
                    Outline newOutline_new = ReducedOutline_s(bbxyzElement_new);

                    //Create a filter to get all the intersection elements with wall.
                    BoundingBoxIntersectsFilter filterW_new = new BoundingBoxIntersectsFilter(newOutline_new);

                    //Create a filter to get StructuralFraming (which include beam and column) and Slabs.
                    ElementCategoryFilter filterSt_new = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);
                    ElementCategoryFilter filterSl_new = new ElementCategoryFilter(BuiltInCategory.OST_Floors);
                    LogicalOrFilter filterS_new = new LogicalOrFilter(filterSt_new, filterSl_new);

                    //Combine two filter.
                    LogicalAndFilter filter_new = new LogicalAndFilter(filterS_new, filterW_new);

                    //A list to store the intersected elements.
                    IList<Element> inter_new = new FilteredElementCollector(doc).WherePasses(filter_new).WhereElementIsNotElementType().ToElements();

                    for (int i = 0; i < inter_new.Count; i++)
                    {
                        if (inter_new[i] != null)
                        {
                            string elementName = inter_new[i].Category.Name;
                            if (elementName == "結構構架")
                            {
                                // Find the depth
                                BoundingBoxXYZ bbxyzBeam = inter_new[i].get_BoundingBox(null);
                                depthList_new.Add(Math.Abs(bbxyzBeam.Min.Z - bbxyzBeam.Max.Z));

                                // Join
                                Transaction joinAndSwitch = new Transaction(doc, "Join and Switch");
                                joinAndSwitch.Start();
                                RunJoinGeometryAndSwitch(doc, wall_new, inter_new[i]);
                                joinAndSwitch.Commit();
                            }
                            else if (elementName == "樓板")
                            {
                                BoundingBoxXYZ bbxyzSlab = inter_new[i].get_BoundingBox(null);
                                depthList_new.Add(Math.Abs(bbxyzSlab.Min.Z - bbxyzSlab.Max.Z));
                                Transaction join = new Transaction(doc, "Join");
                                join.Start();
                                RunJoinGeometry(doc, wall_new, inter_new[i]);
                                join.Commit();
                            }
                            else continue;
                        }
                    }

                    // Adjust the wall height.
                    Transaction t2_new = new Transaction(doc, "Adjust Wall Height");
                    t2_new.Start();
                    if (depthList_new.Count == 0)
                    {
                        wallHeightP_new.Set(levelHeight);
                    }
                    else
                    {
                        wallHeightP_new.Set(levelHeight - depthList_new.Max());
                    }
                    t2_new.Commit();


                }
            }
            return Result.Succeeded;
        }

        void RunJoinGeometry(Document doc, Element e1, Element e2)
        {
            if (!JoinGeometryUtils.AreElementsJoined(doc, e1, e2))
                JoinGeometryUtils.JoinGeometry(doc, e1, e2);
            else
            {
                JoinGeometryUtils.UnjoinGeometry(doc, e1, e2);
                JoinGeometryUtils.JoinGeometry(doc, e1, e2);
            }
        }

        // 執行接合並改變接合順序
        void RunJoinGeometryAndSwitch(Document doc, Element e1, Element e2)
        {
            if (!JoinGeometryUtils.AreElementsJoined(doc, e1, e2))
            {
                JoinGeometryUtils.JoinGeometry(doc, e1, e2);
                JoinGeometryUtils.SwitchJoinOrder(doc, e1, e2);
            }
            else
            {
                JoinGeometryUtils.UnjoinGeometry(doc, e1, e2);
                JoinGeometryUtils.JoinGeometry(doc, e1, e2);
                JoinGeometryUtils.SwitchJoinOrder(doc, e1, e2);
            }
        }

        public double CentimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Centimeters);
        }

        public static List<WallType> GetWallTypes(Autodesk.Revit.DB.Document doc)
        {
            List<WallType> oWallTypes = new List<WallType>();
            try
            {
                FilteredElementCollector collector
                    = new FilteredElementCollector(doc);

                FilteredElementIterator itor = collector
                    .OfClass(typeof(HostObjAttributes))
                    .GetElementIterator();

                // Reset the iterator
                itor.Reset();

                // Iterate through each family
                while (itor.MoveNext())
                {
                    Autodesk.Revit.DB.HostObjAttributes oSystemFamilies =
                    itor.Current as Autodesk.Revit.DB.HostObjAttributes;

                    if (oSystemFamilies == null) continue;

                    // Get the family's category
                    Category oCategory = oSystemFamilies.Category;
                    //TaskDialog.Show("1", oCategory.Name);

                    // Process if the category is found
                    if (oCategory != null)
                    {
                        if (oCategory.Name == "牆")
                        {
                            WallType oWallType = oSystemFamilies as WallType;
                            if (oWallType != null) oWallTypes.Add(oWallType);
                        }
                    }
                } //while itor.NextMove()

                return oWallTypes;
            }
            catch (Exception)
            {
                //MessageBox.Show( ex.Message );
                return oWallTypes = new List<WallType>();
            }
        } //GetWallTypes

        public double UnitsToCentimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
        }

        public int GetWallWidth(Line line1, Line line2)
        {
            Line line = line1.Clone() as Line;
            XYZ ponit = line2.GetEndPoint(0);
            line.MakeUnbound();
            double width = line.Distance(ponit);
            int wallWidth = (int)Math.Round(UnitsToCentimeters(width) * 10);
            return wallWidth;
        }

        public Outline ReducedOutline_s(BoundingBoxXYZ bbxyz)
        {
            XYZ diagnalVector = (bbxyz.Max - bbxyz.Min) / 50;
            Outline newOutline = new Outline(bbxyz.Min + diagnalVector, bbxyz.Max - diagnalVector);
            return newOutline;
        }

        public Outline ReducedOutline(BoundingBoxXYZ bbxyz)
        {
            XYZ diagnalVector = (bbxyz.Max - bbxyz.Min) / 50;
            Outline newOutline = new Outline(bbxyz.Min + diagnalVector, bbxyz.Max - diagnalVector);
            return newOutline;
        }

        public void ChangeWallType(Document doc, Wall wall, Line line1, Line line2)
        {
            WallType wallType = null;
            int width = GetWallWidth(line1, line2);
            String wallTypeName = "Generic - " + width.ToString() + "mm";
            try
            {
                wallType = new FilteredElementCollector(doc)
                    .OfClass(typeof(WallType))
                    .First<Element>(x => x.Name.Equals(wallTypeName)) as WallType;
            }
            catch
            {

            }
            if (wallType != null)
            {
                Transaction t = new Transaction(doc, "Edit Type");
                t.Start();
                try
                {
                    wall.WallType = wallType;
                }
                catch
                {

                }
                t.Commit();
            }

        }

        public List<Wall> GetWalls(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> Walls = collector.OfClass(typeof(Wall)).ToElements();
            List<Wall> ListWalls = new List<Wall>();
            foreach (Wall w in Walls.Cast<Wall>())
            {
                ListWalls.Add(w);
            }
            return ListWalls;
        }

        public static List<Solid> GetElementSolidList(Element elem, Options opt = null, bool useOriginGeom4FamilyInstance = false)
        {
            if (null == elem)
            {
                return null;
            }
            if (null == opt)
                opt = new Options();
            opt.IncludeNonVisibleObjects = false;
            opt.DetailLevel = ViewDetailLevel.Medium;

            GeometryElement gElem;
            List<Solid> solidList = new List<Solid>();
            try
            {
                if (useOriginGeom4FamilyInstance && elem is FamilyInstance)
                {
                    // we transform the geometry to instance coordinate to reflect actual geometry 
                    FamilyInstance fInst = elem as FamilyInstance;
                    gElem = fInst.GetOriginalGeometry(opt);
                    Transform trf = fInst.GetTransform();
                    if (!trf.IsIdentity)
                        gElem = gElem.GetTransformed(trf);
                }
                else
                    gElem = elem.get_Geometry(opt);
                if (null == gElem)
                {
                    return null;
                }
                IEnumerator<GeometryObject> gIter = gElem.GetEnumerator();
                gIter.Reset();
                while (gIter.MoveNext())
                {
                    solidList.AddRange(GetSolids(gIter.Current));
                }
            }
            catch (Exception ex)
            {
                // In Revit, sometime get the geometry will failed.
                string error = ex.Message;
            }
            return solidList;
        }

        public static List<Solid> GetSolids(GeometryObject gObj)
        {
            List<Solid> solids = new List<Solid>();
            if (gObj is Solid) // already solid
            {
                Solid solid = gObj as Solid;
                if (solid.Faces.Size > 0 && Math.Abs(solid.Volume) > 0) // skip invalid solid
                    solids.Add(gObj as Solid);
            }
            else if (gObj is GeometryInstance) // find solids from GeometryInstance
            {
                IEnumerator<GeometryObject> gIter2 = (gObj as GeometryInstance).GetInstanceGeometry().GetEnumerator();
                gIter2.Reset();
                while (gIter2.MoveNext())
                {
                    solids.AddRange(GetSolids(gIter2.Current));
                }
            }
            else if (gObj is GeometryElement) // find solids from GeometryElement
            {
                IEnumerator<GeometryObject> gIter2 = (gObj as GeometryElement).GetEnumerator();
                gIter2.Reset();
                while (gIter2.MoveNext())
                {
                    solids.AddRange(GetSolids(gIter2.Current));
                }
            }
            return solids;
        }

        public static XYZ ProjectPointOntoLine(XYZ point, Line line)
        {
            XYZ lineStart = line.GetEndPoint(0);
            XYZ lineEnd = line.GetEndPoint(1);
            XYZ lineDirection = (lineEnd - lineStart).Normalize();
            XYZ pointVector = point - lineStart;
            double dotProduct = pointVector.DotProduct(lineDirection);
            XYZ projectedVector = dotProduct * lineDirection;
            return lineStart + projectedVector;
        }

        public List<Line> LineSeperatedByFaces(Line line, List<Face> list_face)
        {
            //Boolean xline;
            //if(Math.Abs(line.GetEndPoint(0).Y - line.GetEndPoint(1).Y) < CentimetersToUnits(0.1))
            //{
            //    xline = true;
            //}
            //else
            //{
            //    xline = false;
            //}

            List<XYZ> pointsOnLine = new List<XYZ>();
            if (true)
            {
                foreach (Face face in list_face)
                {
                    IList<CurveLoop> faceEdges = face.GetEdgesAsCurveLoops().ToList();
                    CurveLoop curveLoop = faceEdges[0];
                    foreach (Curve curve in curveLoop)
                    {
                        XYZ point = curve.GetEndPoint(0);
                        XYZ pointOnLine = ProjectPointOntoLine(point, line);
                        pointsOnLine.Add(pointOnLine);
                        break;
                    }
                }
            }

            List<XYZ> pointsOnLine_sort = pointsOnLine.OrderBy(p => p.DistanceTo(line.GetEndPoint(0))).ToList();
            List<Line> lines = new List<Line>();
            for (int i = 0; i < pointsOnLine_sort.Count() - 1; i++)
            {
                if (pointsOnLine_sort[i].DistanceTo(pointsOnLine_sort[i + 1]) > CentimetersToUnits(0.1))
                {
                    Line line_s = Line.CreateBound(pointsOnLine_sort[i], pointsOnLine_sort[i + 1]);
                    if (line_s.Length > CentimetersToUnits(1))
                    {
                        lines.Add(line_s);
                    }
                }
            }
            return lines;
        }
    }
}


