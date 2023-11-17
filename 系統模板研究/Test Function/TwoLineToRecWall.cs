using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.DB.Structure;
using System.Management.Instrumentation;
using System.Diagnostics;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class TwoLineToRecWall : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("1F")) as Level;

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
                    //TaskDialog.Show("Info", "Finished", TaskDialogCommonButtons.Ok);
                    break;
                }
                Line Line1 = Line.CreateBound(points[0], points[1]);
                Line Line2 = Line.CreateBound(points[2], points[3]);
                Line UBLine = Line.CreateBound(points[0], points[1]);
                //UBLine.MakeUnbound();
                double width = GetWallWidth(UBLine, Line2) * 0.003281;
                XYZ midPoint1 = MidPoint(points[0], points[3]);
                XYZ midPoint2 = MidPoint(points[1], points[2]);


                Line BreakLine = Line.CreateBound(points[0], points[2]);
                Line STDLine1 = Line.CreateBound(points[0], points[1]);
                //Line STDLine2 = Line.CreateBound(points[1], points[3]);
                Line wallLine = Line.CreateBound(points[1], points[3]);
                //1mm = 0.03937 inch
                if (Line1.Length - Line2.Length < 1 * 0.003281)
                {
                    //TaskDialog.Show("線條資訊","兩線等長", TaskDialogCommonButtons.Ok);
                    if (midPoint1.DistanceTo(midPoint2) < 5 / 304.78)
                    {
                        //TaskDialog.Show("線條資訊", "兩線中點重疊", TaskDialogCommonButtons.Ok);
                        midPoint1 = MidPoint(points[0], points[2]);
                        midPoint2 = MidPoint(points[1], points[3]);
                    }
                    XYZ UnitZ = new XYZ(0, 0, 1);
                    XYZ Vector2 = points[3] - points[2];
                    XYZ offset = Cross(UnitZ, Vector2) / Line2.Length * width / 2;
                    //TaskDialog.Show("線條資訊", offset.ToString(), TaskDialogCommonButtons.Ok);
                    if (Dot(offset, points[1] - points[2]) > 0)
                    {
                        //MessageBox.Show("1");
                        midPoint1 = points[0] - offset;
                        midPoint2 = points[1] - offset;
                        wallLine = Line.CreateBound(midPoint1, midPoint2);
                    }
                    else
                    {
                        //MessageBox.Show("2");
                        midPoint1 = points[0] + offset;
                        midPoint2 = points[1] + offset;
                        wallLine = Line.CreateBound(midPoint1, midPoint2);
                    }

                }

                else
                {
                    //TaskDialog.Show("線條資訊", "兩線不等長", TaskDialogCommonButtons.Ok);
                    XYZ UnitZ = new XYZ(0, 0, 1);
                    XYZ Vector2 = points[3] - points[2];
                    XYZ offset = Cross(UnitZ, Vector2) / Line2.Length * width / 2;
                    //TaskDialog.Show("線條資訊", offset.ToString(), TaskDialogCommonButtons.Ok);
                    if (Dot(offset, points[1] - points[2]) > 0)
                    {
                        //MessageBox.Show("1");
                        midPoint1 = points[0] - offset;
                        midPoint2 = points[1] - offset;
                        wallLine = Line.CreateBound(midPoint1, midPoint2);
                    }
                    else
                    {
                        //MessageBox.Show("2");
                        midPoint1 = points[0] + offset;
                        midPoint2 = points[1] + offset;
                        wallLine = Line.CreateBound(midPoint1, midPoint2);
                    }

                }

                wallLine = Line.CreateBound(midPoint1, midPoint2);

                FilteredElementCollector colWall = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .OfClass(typeof(Wall));
                IList<Wall> walls = new List<Wall>();
                foreach (Wall w in colWall)
                {
                    walls.Add(w);
                }

                Transaction t1 = new Transaction(doc, "Create Wall");
                t1.Start();
                Wall wall = Wall.Create(doc, wallLine, level.Id, true);
                //Parameter wallHeightP = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                //wallHeightP.Set(2500 * 0.003281);
                WallUtils.DisallowWallJoinAtEnd(wall, 0);
                WallUtils.DisallowWallJoinAtEnd(wall, 1);
                t1.Commit();
                // Get the external wall face for the profile

                Transaction t2 = new Transaction(doc, "Create Wall");
                t2.Start();

                IList<Reference> sideFaces = HostObjectUtils.GetSideFaces(wall, ShellLayerType.Exterior);

                MessageBox.Show(sideFaces.Count().ToString());

                Element e2 = doc.GetElement(sideFaces[0]);

                Face face = e2.GetGeometryObjectFromReference(sideFaces[0]) as Face;

                // The normal of the wall external face.

                XYZ normal = face.ComputeNormal(new UV(0, 0));

                // Offset curve copies for visibility.

                //Transform offset = Transform.CreateTranslation(5 * normal);

                // If the curve loop direction is counter-
                // clockwise, change its color to RED.

                Color colorRed = new Color(255, 0, 0);

                // Get edge loops as curve loops.
                IList<CurveLoop> curveLoops = face.GetEdgesAsCurveLoops();
                foreach (Curve curve in curveLoops[0])
                {
                    MessageBox.Show(curve.GetEndPoint(0).ToString() + curve.GetEndPoint(1).ToString());
                }

                //using (Transaction trans = new Transaction(doc))
                //{
                //    trans.Start("start");
                //    Element e1 = wall;
                //    Element e2 = doc.GetElement(r2.ElementId);

                //    // 判斷選擇順序與優先順序是否相同
                //    // 若是，則直接執行接合
                //    // 若否，則先執行接合後再改變順序
                //    if (RunWithDefaultOrder(e1, e2))
                //        RunJoinGeometry(doc, e1, e2);
                //    else
                //        RunJoinGeometryAndSwitch(doc, e1, e2);

                //    trans.Commit();
                //}

                // ExporterIFCUtils class can also be used for 
                // non-IFC purposes. The SortCurveLoops method 
                // sorts curve loops (edge loops) so that the 
                // outer loops come first.

                //IList<IList<CurveLoop>> curveLoopLoop = ExporterIFCUtils.SortCurveLoops( curveLoops);

                //foreach (IList<CurveLoop> curveLoops2 in curveLoopLoop)
                //{
                //    foreach (CurveLoop curveLoop2 in curveLoops2)
                //    {
                //        // Check if curve loop is counter-clockwise.

                //        bool isCCW = curveLoop2.IsCounterclockwise(
                //          normal);

                //        CurveArray curves = creapp.NewCurveArray();

                //        foreach (Curve curve in curveLoop2)
                //        {
                //            curves.Append(curve.CreateTransformed(offset));
                //        }

                //        // Create model lines for an curve loop.

                //        Plane plane = creapp.NewPlane(curves);

                //        SketchPlane sketchPlane
                //          = SketchPlane.Create(doc, plane);

                //        ModelCurveArray curveElements
                //          = credoc.NewModelCurveArray(curves,
                //            sketchPlane);

                //        if (isCCW)
                //        {
                //            foreach (ModelCurve mcurve in curveElements)
                //            {
                //                OverrideGraphicSettings overrides
                //                  = view.GetElementOverrides(
                //                    mcurve.Id);

                //                overrides.SetProjectionLineColor(
                //                  colorRed);

                //                view.SetElementOverrides(
                //                  mcurve.Id, overrides);
                //            }
                //        }
                //    }
                //}
                t2.Commit();


                List<double> depthList = new List<double>();

                //Get the element by elementID and get the boundingbox and outline of this element.
                Element elementWall = wall as Element;
                BoundingBoxXYZ bbxyzElement = elementWall.get_BoundingBox(null);
                Outline outline = new Outline(bbxyzElement.Min, bbxyzElement.Max);

                //Create a filter to get all the intersection elements with wall.
                BoundingBoxIntersectsFilter filterW = new BoundingBoxIntersectsFilter(outline);

                //Create a filter to get StructuralFraming (which include beam and column) and Slabs.
                ElementCategoryFilter filterSt = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);
                ElementCategoryFilter filterSl = new ElementCategoryFilter(BuiltInCategory.OST_Floors);
                LogicalOrFilter filterS = new LogicalOrFilter(filterSt, filterSl);

                //Combine two filter.
                LogicalAndFilter filter = new LogicalAndFilter(filterS, filterW);

                //A list to store the intersected elements.
                IList<Element> inter = new FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements();
                //MessageBox.Show(inter.Count.ToString());

                for (int i = 0; i < inter.Count; i++)
                {
                    if (inter[i] != null)
                    {
                        //MessageBox.Show(inter[i].ToString());
                        string elementName = inter[i].Category.Name;
                        if (elementName == "結構構架")
                        {
                            //MessageBox.Show("結構構架");
                            BoundingBoxXYZ bbxyzBeam = inter[i].get_BoundingBox(null);
                            depthList.Add(Math.Abs(bbxyzBeam.Min.Z - bbxyzBeam.Max.Z));
                            //MessageBox.Show(bbxyzBeam.Min.ToString());
                            //MessageBox.Show(bbxyzBeam.Max.ToString());
                        }
                        else if (elementName == "樓板")
                        {
                            //MessageBox.Show("樓板");
                            BoundingBoxXYZ bbxyzSlab = inter[i].get_BoundingBox(null);
                            depthList.Add(Math.Abs(bbxyzSlab.Min.Z - bbxyzSlab.Max.Z));
                        }
                        else continue;
                    }

                }
                Transaction t3 = new Transaction(doc, "Create Wall");
                t3.Start();
                //MessageBox.Show(depthList.Count.ToString());
                //MessageBox.Show(depthList.Max().ToString());
                Parameter wallHeightP = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                if (depthList.Count() != 0)
                {
                    wallHeightP.Set(3300 * 0.003281 - depthList.Min());
                }
                else
                {
                    wallHeightP.Set(3300 * 0.003281);
                }

                t3.Commit();

                ChangeWallType(doc, wall, Line1, Line2);
            }
            return Result.Succeeded;
        }

        private XYZ MidPoint(XYZ a, XYZ b)
        {
            XYZ c = (a + b) / 2;
            return c;
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


        //private double GetWallAdjustedHeight(Wall wall)
        //{
        //    double height = 0;
        //    return height;
        //}
        private Line TransformLine(Transform transform, Line line)
        {
            XYZ startPoint = transform.OfPoint(line.GetEndPoint(0));
            XYZ endPoint = transform.OfPoint(line.GetEndPoint(1));
            Line newLine = Line.CreateBound(startPoint, endPoint);
            return newLine;
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

        public int GetWallWidth(Line line1, Line line2)
        {
            XYZ ponit1_1 = line1.GetEndPoint(0);
            XYZ ponit1_2 = line1.GetEndPoint(1);
            XYZ ponit2_1 = line2.GetEndPoint(0);
            XYZ ponit2_2 = line2.GetEndPoint(1);
            line1.MakeUnbound();
            double width = line1.Distance(ponit2_2);
            int wallWidth = (int)Math.Round(width / 0.003281);

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
        public List<Wall> GetWalls(Document doc)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> Walls = collector.OfClass(typeof(Wall)).ToElements();
            List<Wall> ListWalls = new List<Wall>();
            foreach (Wall w in Walls)
            {
                ListWalls.Add(w);
            }
            return ListWalls;
        }
        /// <summary>
        /// 不考慮B1,RF1這種情況，
        /// </summary>
        /// <param name="targetElement"></param>
        /// <returns></returns>
        public static int GetLevel(Element targetElement)
        {
            string level = null;
            int categoryId = targetElement.Category.Id.IntegerValue;

            if (categoryId == (int)BuiltInCategory.OST_Walls)
            {
                level = targetElement.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsValueString();
            }
            if (categoryId == (int)BuiltInCategory.OST_Columns || categoryId == (int)BuiltInCategory.OST_StructuralColumns)
            {
                level = targetElement.get_Parameter(BuiltInParameter.FAMILY_BASE_LEVEL_PARAM).AsValueString();
            }
            else if (categoryId == (int)BuiltInCategory.OST_StructuralFraming)
            {
                level = targetElement.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM).AsValueString();

            }
            else if (categoryId == (int)BuiltInCategory.OST_Floors)
            {
                level = targetElement.get_Parameter(BuiltInParameter.LEVEL_PARAM).AsValueString();
            }
            if (level != null)
            {
                string result = System.Text.RegularExpressions.Regex.Replace(level, @"[^0-9]+", "");
                int levelnumber = int.Parse(result);
                if (categoryId == (int)BuiltInCategory.OST_StructuralFraming || categoryId == (int)BuiltInCategory.OST_Floors)
                {
                    levelnumber--;
                }
                return levelnumber;
            }
            else
                return 0;


        }

        private void DeleteElement(Autodesk.Revit.DB.Document document, ElementId elemId)
        {

            // 將指定元件以及所有與該元件相關聯的元件刪除，並將刪除後所有的元件存到到容器中
            ICollection<Autodesk.Revit.DB.ElementId> deletedIdSet = document.Delete(elemId);
            // 可利用上述容器來查看刪除的數量，若數量為0，則刪除失敗，提供錯誤訊息
            if (deletedIdSet.Count == 0)
            {
                throw new Exception("選取的元件刪除失敗");
            }
        }

        public void RunJoinGeometry(Document doc, Element e1, Element e2)
        {
            if (!JoinGeometryUtils.AreElementsJoined(doc, e1, e2))
            {
                //MessageBox.Show("直接接合");
                JoinGeometryUtils.JoinGeometry(doc, e1, e2);
            }
            else
            {
                //MessageBox.Show("先取消接合再執行接合");
                JoinGeometryUtils.UnjoinGeometry(doc, e1, e2);
                JoinGeometryUtils.JoinGeometry(doc, e1, e2);

            }
        }

        // 執行接合並改變接合順序
        public void RunJoinGeometryAndSwitch(Document doc, Element e1, Element e2)
        {
            if (!JoinGeometryUtils.AreElementsJoined(doc, e1, e2))
            {
                //MessageBox.Show("直接接合再交換順序");
                JoinGeometryUtils.JoinGeometry(doc, e1, e2);
                JoinGeometryUtils.SwitchJoinOrder(doc, e1, e2);

            }
            else
            {
                //MessageBox.Show("先取消接合再執行接合再交換順序");
                JoinGeometryUtils.UnjoinGeometry(doc, e1, e2);
                //MessageBox.Show(JoinGeometryUtils.AreElementsJoined(doc, e1, e2).ToString());
                JoinGeometryUtils.UnjoinGeometry(doc, e1, e2);
                JoinGeometryUtils.UnjoinGeometry(doc, e1, e2);
                JoinGeometryUtils.JoinGeometry(doc, e1, e2);
                JoinGeometryUtils.SwitchJoinOrder(doc, e1, e2);
            }
        }
        public bool RunWithDefaultOrder(Element e1, Element e2)
        {
            int sum = GetOrderOfElement(e1) + GetOrderOfElement(e2);

            if (sum == 3 || sum == 7)
            {
                //MessageBox.Show("Correct order");
                return true;
            }

            else
            {
                //MessageBox.Show("Wrong order");
                return false;

            }

        }

        public int GetOrderOfElement(Element e)
        {
            if (e.Category.Name == "柱" || e.Category.Name == "結構柱")
            {
                //TaskDialog.Show("Info", e.Category.Name);
                return 1;
            }
            else if (e.Category.Name == "結構構架")
            {
                //TaskDialog.Show("Info", e.Category.Name);
                return 2;
            }
            else if (e.Category.Name == "樓板")
            {
                //TaskDialog.Show("Info", e.Category.Name);
                return 3;
            }
            else if (e.Category.Name == "牆")
            {
                //TaskDialog.Show("Info", e.Category.Name);
                return 4;
            }
            else
            {
                TaskDialog.Show("Info", "我是" + e.Category.Name.ToString());
                return 0;
            }
        }

    }
}