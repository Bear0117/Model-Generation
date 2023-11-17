using Autodesk.Revit.UI.Selection;
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class CountOverlapped : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("3F")) as Level;
            while (true)
            {
                IList<XYZ> points = new List<XYZ>();
                for (int j = 0; j < 2; j++)
                {
                    Reference r = uidoc.Selection.PickObject(ObjectType.PointOnElement);
                    Element elem = doc.GetElement(r);
                    LocationCurve locationCurve = elem.Location as LocationCurve;
                    Line locationLine = locationCurve.Curve as Line;
                    points.Add(locationLine.GetEndPoint(0));
                    points.Add(locationLine.GetEndPoint(1));
                }

                //Stop building the wall.
                if (points[0].ToString() == points[2].ToString())
                {
                    break;
                }

                // To calculate the  central line of two lines. 
                Line Line1 = Line.CreateBound(points[0], points[1]);
                Line Line2 = Line.CreateBound(points[2], points[3]);
                XYZ midPoint1 = MidPoint(points[0], points[3]);
                XYZ midPoint2 = MidPoint(points[1], points[2]);

                double width = GetWallWidth(Line1, Line2); // Unit : mm

                Line wallLine = null;
                XYZ unitZ = new XYZ(0, 0, 1);
                XYZ vector_line1 = points[0] - points[1];
                XYZ offset = Cross(unitZ, vector_line1) / Line1.Length * CentimetersToUnits(width / 10) / 2;
                if (Dot(offset, points[2] - points[1]) > 0)
                {
                    midPoint1 = points[0] + offset;
                    midPoint2 = points[1] + offset;
                    wallLine = Line.CreateBound(midPoint1, midPoint2);
                }
                else
                {
                    midPoint1 = points[0] - offset;
                    midPoint2 = points[1] - offset;
                    wallLine = Line.CreateBound(midPoint1, midPoint2);
                }

                FilteredElementCollector colWall = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .OfClass(typeof(Wall));
                IList<Wall> walls = new List<Wall>();
                foreach (Wall w in colWall)
                {
                    walls.Add(w);
                }

                // Create Wall.
                Transaction t1 = new Transaction(doc, "Create Wall");
                t1.Start();

                Wall wall = Wall.Create(doc, wallLine, level.Id, true);
                Element elementWall = wall as Element;
                BoundingBoxXYZ bbxyzWall = elementWall.get_BoundingBox(null);

                Parameter heightParam = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                double height = heightParam.AsDouble();
                MessageBox.Show(height.ToString());
                WallUtils.DisallowWallJoinAtEnd(wall, 0);
                WallUtils.DisallowWallJoinAtEnd(wall, 1);

                t1.Commit();

                // Get the geometry of the wall
                Options geomOptions = new Options();
                GeometryElement wallGeomElem = wall.get_Geometry(geomOptions);
                Solid wallSolid = null;

                // Find the first solid in the wall geometry element
                foreach (GeometryObject geomObj in wallGeomElem)
                {
                    Solid solid = geomObj as Solid;
                    if (solid != null && solid.Volume > 0)
                    {
                        wallSolid = solid;
                        break;
                    }
                }

                if (wallSolid != null)
                {
                    // Get all elements in the model
                    FilteredElementCollector elementCollector = new FilteredElementCollector(doc);

                    // Filter out any elements that are not of interest, such as the wall itself
                    elementCollector.Excluding(new ElementId[] { wall.Id });

                    List<Element> overlappingElements = new List<Element>();

                    // Loop through each element and check for intersections with the wall
                    foreach (Element element in elementCollector)
                    {
                        GeometryElement elemGeomElem = element.get_Geometry(geomOptions);

                        // Find the first solid in the element geometry element
                        foreach (GeometryObject geomObj in elemGeomElem)
                        {
                            Solid elemSolid = geomObj as Solid;
                            if (elemSolid != null && elemSolid.Volume > 0)
                            {
                                // Check if the wall and element solids intersect
                                Solid interSolid = BooleanOperationsUtils.ExecuteBooleanOperation(wallSolid, elemSolid, BooleanOperationsType.Intersect);
                                if (Math.Abs(interSolid.Volume) > 0.000001)
                                {
                                    overlappingElements.Add(element);
                                }
                                //if (wallSolid.Intersect(elemSolid) != SetComparisonResult.Disjoint)
                                //{
                                //    overlappingElements.Add(element);
                                //    break;
                                //}

                            }
                        }
                    }

                    // Print the result
                    MessageBox.Show(overlappingElements.Count.ToString());
                }
            }

            


            return Result.Succeeded;
        }
        private XYZ Cross(XYZ a, XYZ b)
        {

            XYZ c = new XYZ(a.Y * b.Z - a.Z * b.Y, a.Z * b.X - a.X * b.Z, a.X * b.Y - a.Y * b.X);
            return c;
        }
        private double Dot(XYZ a, XYZ b)
        {

            double c = a.X * b.X + a.Y * b.Y;
            return c;
        }

        public double CentimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Centimeters);
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

        public double UnitsToCentimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
        }
        private XYZ MidPoint(XYZ a, XYZ b)
        {
            XYZ c = (a + b) / 2;
            return c;
        }

    }
}
