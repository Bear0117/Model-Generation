using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class TwoLineToWall : IExternalCommand
    {
        UIDocument document;
        UIApplication application;
        Autodesk.Revit.ApplicationServices.Application app;
        Document doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            application = commandData.Application;
            document = application.ActiveUIDocument;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            //WallEditor wallEditor = new WallEditor(commandData);
            //wallEditor.ShowDialog();

            //ElementId wallTypeId = wallEditor.wallTypeId;

            //String wallHeightValue = wallEditor.wallHeightValue.ToString();
            //double wallHeight = int.Parse(wallHeightValue) * 0.003281;

            double wallHeight = 3200 * 0.003281;
            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("2F")) as Level;


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
                    TaskDialog.Show("Info", "Finished", TaskDialogCommonButtons.Ok);
                    break;
                }
                Line Line1 = Line.CreateBound(points[0], points[1]);
                Line Line2 = Line.CreateBound(points[2], points[3]);
                Line UBLine = Line.CreateBound(points[0], points[1]);
                //UBLine.MakeUnbound();
                double width = GetWallWidth(UBLine, Line2) * 0.003281;
                XYZ midPoint1 = midPoint(points[0], points[3]);
                XYZ midPoint2 = midPoint(points[1], points[2]);
                
                    
                Line BreakLine = Line.CreateBound(points[0], points[2]);
                Line STDLine1 = Line.CreateBound(points[0], points[1]);
                Line STDLine2 = Line.CreateBound(points[1], points[3]);
                Line wallLine = Line.CreateBound(midPoint1, midPoint2);
                //1mm = 0.03937 inch
                if (Line1.Length == Line2.Length)
                {
                    //TaskDialog.Show("線條資訊","兩線等長", TaskDialogCommonButtons.Ok);
                    if (midPoint1.DistanceTo(midPoint2)  < 5 / 304.78)
                    {
                        //TaskDialog.Show("線條資訊", "兩線中點重疊", TaskDialogCommonButtons.Ok);
                        midPoint1 = midPoint(points[0], points[2]);
                        midPoint2 = midPoint(points[1], points[3]);
                        wallLine = Line.CreateBound(midPoint1, midPoint2);                       
                    }
                }
                
                
                else 
                {
                    //TaskDialog.Show("線條資訊", "兩線不等長", TaskDialogCommonButtons.Ok);
                    XYZ UnitZ = new XYZ(0, 0, 1);
                    XYZ Vector2 = points[3] - points[2];
                    XYZ offset = Cross(UnitZ, Vector2)/Line2.Length * width / 2;
                    //TaskDialog.Show("線條資訊", offset.ToString(), TaskDialogCommonButtons.Ok);
                    if (Dot(offset, points[1] - points[2]) > 0)
                    {
                        midPoint1 = points[0] - offset;
                        midPoint2 = points[1] - offset;
                        wallLine = Line.CreateBound(midPoint1, midPoint2);
                    }
                    else
                    {
                        midPoint1 = points[0] + offset;
                        midPoint2 = points[1] + offset;
                        wallLine = Line.CreateBound(midPoint1, midPoint2);
                    }

                }

                foreach (XYZ point in points)
                {
                    //TaskDialog.Show("線條資訊", point.ToString(), TaskDialogCommonButtons.Ok);
                }

                wallLine = Line.CreateBound(midPoint1, midPoint2);
                //TaskDialog.Show("線條資訊", wallLine.Length.ToString(), TaskDialogCommonButtons.Ok);
                if (Math.Abs(wallLine.Length - 1350 * 0.00328) < 1 * 0.00328) 
                {

                }

                //Transaction t2 = new Transaction(doc, "change_wall_thickness");
                //t2.Start();
                //List<WallType> oWallTypes = new List<WallType>();
                //List<WallType> GenericWallTypes = new List<WallType>();
                //oWallTypes = GetWallTypes(doc);
                //TaskDialog.Show("1", oWallTypes[0].ToString());
                ////Get a list of all Generic Wall Types in the document
                //foreach (WallType wt in oWallTypes)
                //{
                //    if (wt.Name.Contains("Generic - 100mm"))
                //    {
                //        GenericWallTypes.Add(wt);
                //    }
                //}
                //CompoundStructure cs = GenericWallTypes[0].GetCompoundStructure();
                //int i = cs.GetFirstCoreLayerIndex();
                //double thickness = 150 * 0.003281;
                //cs.SetLayerWidth(i, thickness);
                //t2.Commit();



                Transaction t1 = new Transaction(doc, "Create Wall");
                t1.Start();
                Wall wall = Wall.Create(doc, wallLine, level.Id, true); 
                WallUtils.DisallowWallJoinAtEnd(wall, 0);
                WallUtils.DisallowWallJoinAtEnd(wall, 1);
                //wall.ChangeTypeId(wallTypeId);
                Parameter wallHeightParameter = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                wallHeightParameter.Set(wallHeight);
                t1.Commit();
                ChangeWallType(doc, wall, Line1, Line2);
            }
            return Result.Succeeded;
        }

        private XYZ midPoint(XYZ a,XYZ b)
        {
            XYZ c = (a + b) / 2;
            return c;
        }

        //public void changeWallType(Document doc, String wallName)
        //{
        //    using (Transaction t = new Transaction(doc, "Change Wall Type"))
        //    {
        //        t.Start();
        //        // the familyinstance you want to change -> e.g. UB-Universal Beam 305x165x40UB
        //        List<Wall> walls = new FilteredElementCollector(doc, doc.ActiveView.Id)
        //       .OfClass(typeof(Wall))
        //       .OfCategory(BuiltInCategory.OST_StructuralFraming)
        //       .Cast<Wall>()
        //       .Where(q => q.Name == "Generic - 100mm").ToList();

        //        // the target familyinstance(familysymbol) what it should be -> e.g. UB-Universal Beam 533x210x92UB         
        //        Wall wallId = new FilteredElementCollector(doc)
        //        .OfClass(typeof(Wall))
        //        .OfCategory(BuiltInCategory.OST_StructuralFraming)
        //        .Cast<Wall>()
        //        .FirstOrDefault(q => q.Id.ToString() == wallName) as Wall;

        //        foreach (Wall wall in walls)
        //        {
        //            wall.ChangeTypeId(doc, walls,) = fs;
        //        }

        //        t.Commit();
        //    }
        //    Wall wall;
        //    double dThickness = 500 / 304.8;
        //    double dHeight = wall.get_Prarmeter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble();
        //    double dOffset = wall.get_Parameter(BuiltInParameter.WALL_BASE_OFFSET).AsDouble();
        //    LocationCurve lc = wall.Location as LocationCurve;
        //    Curve curve = lc.Curve;
        //    using (Transaction trans = new Transaction(doc, "edit"))
        //    {
        //        trans.Start();
        //        try
        //        {
        //            //厚度时类型属性，建议新建一个墙体类型
        //            WallType newType = wall.WallType.Duplicate("NewWallType") as WallType;
        //            CompoundStructure cs = newType.GetCompoundStrycture();
        //            //获取所有层
        //            IList<CompoundStructureLayer> lstLayers = cs.GetLayers();
        //            foreach (CompoundStructureLayer item in lstLayers)
        //            {
        //                if (item.Function == MaterialFunctionAssignment.Structure)
        //                {//这里只考虑有一个结构层，如果有多个就自己算算
        //                    item.Width = dThickness;
        //                    break;
        //                }
        //            }
        //            //修改后要设置一遍
        //            cs.SetLayers(lstLayers);
        //            newType.SetCompoundStructure(cs);
        //            Wall.Create(doc, curve, newType.Id, wall.LevelId, dHeight, dOffset, false, false);
        //            doc.Delete(wall.Id);
        //            trans.Commit();
        //        }
        //        catch
        //        {
        //            trans.RollBack();
        //        }
        //    }
        //}

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
        private XYZ Cross (XYZ a, XYZ b)
        {

            XYZ c = new XYZ(a.Y * b.Z - a.Z * b.Y, a.Z * b.X - a.X * b.Z, a.X * b.Y -a.Y * b.X);
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
            int wallWidth =(int)Math.Round(width / 0.003281);

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
    }
}