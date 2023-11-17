using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AdjustWallHeight : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Default
            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("1F")) as Level;

            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();

            // 取消接合所有結構元件
            using (Transaction tx = new Transaction(doc))
            {
                tx.Start("Unjoin All Elements");
                UnjoinAllElements(doc);
                tx.Commit();
            }

            foreach (ElementId elemId in ids)
            {
                Element elem = doc.GetElement(elemId);
                Wall w = elem as Wall;
                level = doc.GetElement(w.LevelId) as Level;

                double levelHeight = GetLevelHieght(doc, level);
                
                // Adjust the height of wall with level height.
                Transaction t1 = new Transaction(doc, "Adjust Wall Height");
                t1.Start();
                Parameter wallHeightP = w.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                wallHeightP.Set(levelHeight);
                WallUtils.DisallowWallJoinAtEnd(w, 0);
                WallUtils.DisallowWallJoinAtEnd(w, 1);
                t1.Commit();

                List<double> depthList = new List<double>();

                //Get the element by elementID and get the boundingbox and outline of this element.
                Element elementWall = w as Element;
                BoundingBoxXYZ bbxyzElement = elementWall.get_BoundingBox(null);
                Outline outline = new Outline(bbxyzElement.Min, bbxyzElement.Max);
                Outline newOutline = ReducedOutline(bbxyzElement);

                //Create a filter to get all the intersection elements with wall.
                BoundingBoxIntersectsFilter filterW = new BoundingBoxIntersectsFilter(newOutline);

                //Create a filter to get StructuralFraming (which include beam and column) and Slabs.
                ElementCategoryFilter filterSt = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);
                ElementCategoryFilter filterSl = new ElementCategoryFilter(BuiltInCategory.OST_Floors);
                ElementCategoryFilter filterSc = new ElementCategoryFilter(BuiltInCategory.OST_Columns);
                LogicalOrFilter filterStl = new LogicalOrFilter(filterSt, filterSl);
                LogicalOrFilter filterS = new LogicalOrFilter(filterStl, filterSc);

                //Combine two filter.
                LogicalAndFilter filter = new LogicalAndFilter(filterS, filterW);

                //A list to store the intersected elements.
                IList<Element> inter = new FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements();
                //MessageBox.Show(inter.Count.ToString());

                for (int i = 0; i < inter.Count; i++)
                {
                    if (inter[i] != null)
                    {
                        string elementName = inter[i].Category.Name;
                        if (elementName == "結構構架")
                        {
                            // Find the depth
                            BoundingBoxXYZ bbxyzBeam = inter[i].get_BoundingBox(null);
                            depthList.Add(Math.Abs(bbxyzBeam.Min.Z - level.Elevation));
                            MessageBox.Show("beam");

                            // Join
                            //Transaction joinAndSwitch = new Transaction(doc, "Join and Switch");
                            //joinAndSwitch.Start();
                            //RunJoinGeometry(doc, w, inter[i]);
                            //joinAndSwitch.Commit();
                        }
                        else if (elementName == "樓板")
                        {
                            BoundingBoxXYZ bbxyzSlab = inter[i].get_BoundingBox(null);
                            depthList.Add(Math.Abs(bbxyzSlab.Min.Z - level.Elevation));
                            MessageBox.Show("slab");

                            //Transaction join = new Transaction(doc, "Join");
                            //join.Start();
                            //RunJoinGeometry(doc, w, inter[i]);
                            //join.Commit();
                        }
                        else if(elementName == "柱")
                        {
                            MessageBox.Show("column");
                            //Transaction join = new Transaction(doc, "Join");
                            //join.Start();
                            //RunJoinGeometryAndSwitch(doc, w, inter[i]);
                            //join.Commit();
                        }
                        else continue;
                    }
                }


                List<double> faceZ = new List<double>();
                List<Solid> list_solid = GetElementSolidList(w);
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
                        //MessageBox.Show(((face.GetBoundingBox().Max + face.GetBoundingBox().Min) / 2).ToString());
                    }
                    break;
                    //MessageBox.Show(faceArray.Size.ToString());
                }
                Transaction t2 = new Transaction(doc, "Adjust Wall Height");
                t2.Start();
                if (depthList.Count == 0)
                {
                    wallHeightP.Set(levelHeight);
                }
                else
                {
                    wallHeightP.Set(depthList.Max());
                }

                //wallHeightP.Set(faceZ.Max() - level.Elevation);
                MessageBox.Show(depthList.Max().ToString());
                t2.Commit();

                for (int i = 0; i < inter.Count; i++)
                {
                    if (inter[i] != null)
                    {
                        string elementName = inter[i].Category.Name;
                        if (elementName == "結構構架")
                        {
                            // Join
                            //Transaction joinAndSwitch = new Transaction(doc, "Join and Switch");
                            //joinAndSwitch.Start();
                            //RunJoinGeometry(doc, w, inter[i]);
                            //joinAndSwitch.Commit();
                        }
                        else if (elementName == "樓板")
                        {
                            // Join
                            Transaction join = new Transaction(doc, "Join");
                            join.Start();
                            RunJoinGeometry(doc, w, inter[i]);
                            join.Commit();
                        }
                        else if (elementName == "柱")
                        {
                            //Transaction join = new Transaction(doc, "Join");
                            //join.Start();
                            //RunJoinGeometryAndSwitch(doc, w, inter[i]);
                            //join.Commit();
                        }
                        else continue;
                    }
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

        public double UnitsToCentimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
        }

        public double GetLevelHieght(Document doc, Level level)
        {
            double levelHeight = CentimetersToUnits(330); // default
            List<double> levelDistancesList = new List<double>();

            // retrieve all the Level elements in the document
            FilteredElementCollector collector_level = new FilteredElementCollector(doc).OfClass(typeof(Level));
            IEnumerable<Level> levels = collector_level.Cast<Level>();

            // loop through all the levels
            foreach (Level level_1 in levels)
            {
                // get the level elevation.
                double levelElevation = level_1.Elevation;
                if (Math.Abs(levelElevation - level.Elevation) < CentimetersToUnits(0.1))
                {
                    continue;
                }
                if (levelElevation - level.Elevation > CentimetersToUnits(1))
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
            return levelHeight;
        }

        public Outline ReducedOutline(BoundingBoxXYZ bbxyz)
        {
            XYZ diagnalVector = (bbxyz.Max - bbxyz.Min) / 100;
            Outline newOutline = new Outline(bbxyz.Min + diagnalVector, bbxyz.Max - diagnalVector);
            return newOutline;
        }

        static public List<Solid> GetElementSolidList(Element elem, Options opt = null, bool useOriginGeom4FamilyInstance = false)
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


        private void UnjoinAllElements(Document doc)
        {
            List<Element> structuralElements = new List<Element>();

            // 篩選牆
            FilteredElementCollector walls = new FilteredElementCollector(doc).OfClass(typeof(Wall));
            structuralElements.AddRange(walls.ToList());

            // 篩選柱
            FilteredElementCollector columns = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralColumns);
            structuralElements.AddRange(columns.ToList());

            // 篩選樑
            FilteredElementCollector beams = new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance))
                .OfCategory(BuiltInCategory.OST_StructuralFraming);
            structuralElements.AddRange(beams.ToList());

            // 篩選版
            FilteredElementCollector floors = new FilteredElementCollector(doc).OfClass(typeof(Floor));
            structuralElements.AddRange(floors.ToList());


            for (int i = 0; i < structuralElements.Count; i++)
            {
                for (int j = i + 1; j < structuralElements.Count; j++)
                {
                    Element elem1 = structuralElements[i];
                    Element elem2 = structuralElements[j];

                    if (JoinGeometryUtils.AreElementsJoined(doc, elem1, elem2))
                    {
                        JoinGeometryUtils.UnjoinGeometry(doc, elem1, elem2);
                    }
                }
            }
        }


        static public List<Solid> GetSolids(GeometryObject gObj)
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
    }
}
