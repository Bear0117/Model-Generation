using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateColumns : IExternalCommand
    {
        Document doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;

            // To know the information of selected layer
            Reference refer = uidoc.Selection.PickObject(ObjectType.PointOnElement, "Select a CAD Layer");
            Element elem = doc.GetElement(refer);


            Level level = doc.GetElement(elem.LevelId) as Level;
            
            GeometryObject geoObj = elem.GetGeometryObjectFromReference(refer);
            Category targetCategory = null;
            ElementId graphicsStyleId = null;

            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                graphicsStyleId = geoObj.GraphicsStyleId;
                if (doc.GetElement(geoObj.GraphicsStyleId) is GraphicsStyle gs)
                {
                    targetCategory = gs.GraphicsStyleCategory;
                    // Get the name of the CAD layer which is selected (Column).
                    String name = gs.GraphicsStyleCategory.Name;
                }
            }

            // Initialize Parameters.
            ModelingParam.Initialize();
            double gridline_size = ModelingParam.parameters.General.GridSize * 10; // unit: mm
            double[] columnWidthRange = ModelingParam.parameters.ColumnParam.columnWidthsRange;

            TransactionGroup transGroup = new TransactionGroup(doc, "Grab the lines in column layer");
            transGroup.Start();
            CurveArray curveArray = new CurveArray();
            //curveArray.Append("你要加的東西");
            //List<Curve> curves = new List<Curve>();
            //curves.Add("");

            GeometryElement geoElem = elem.get_Geometry(new Options());

            // Get the "default" FamilySymbol (Architectural Columns).      
            FamilySymbol default_column = new FilteredElementCollector(doc)
            .OfClass(typeof(FamilySymbol))
            .OfCategory(BuiltInCategory.OST_Columns)
            .Cast<FamilySymbol>()
            .FirstOrDefault(q => q.Name == "default") as FamilySymbol;

            // Get the "default" FamilySymbol (Structural Columns).
            {
                //collector.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralColumns);
                //FamilySymbol default_column_s = new FilteredElementCollector(doc)
                //.OfClass(typeof(FamilySymbol))
                //.OfCategory(BuiltInCategory.OST_Columns)
                //.Cast<FamilySymbol>()
                //.FirstOrDefault(q => q.Name == "default") as FamilySymbol;
            }

            foreach (GeometryObject gObj in geoElem)
            {
                GeometryInstance geomInstance = gObj as GeometryInstance; 
                Transform transform = geomInstance.Transform;//座標轉換
                if (null != geomInstance)
                {
                    foreach (GeometryObject insObj in geomInstance.SymbolGeometry)
                    {
                        if (insObj.GraphicsStyleId.IntegerValue != graphicsStyleId.IntegerValue) 
                            continue;

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.Line")
                        {
                            MessageBox.Show("There is a 'line' that cannot construct a rectangular column.");
                            continue;
                        }

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> points = polyLine.GetCoordinates();
                            IList<XYZ> points_t = new List<XYZ>();
                            foreach (XYZ point in points)
                            {
                                XYZ point_t = transform.OfPoint(point);
                                points_t.Add(point_t);
                            }
                            XYZ point_a = Algorithm.RoundPoint(points_t[0], gridline_size);
                            XYZ point_b = Algorithm.RoundPoint(points_t[2], gridline_size);
                            double disX = Math.Abs(Algorithm.UnitsToCentimeters(point_a.X - point_b.X));
                            double disY = Math.Abs(Algorithm.UnitsToCentimeters(point_a.Y - point_b.Y));
                            if (disX < columnWidthRange[0] || disX > columnWidthRange[1] || disY < columnWidthRange[0] ||disY > columnWidthRange[1]) continue;

                            XYZ center = GetMiddlePoint(point_a, point_b);
                            using (Transaction tx = new Transaction(doc))
                            {
                                try
                                {
                                    tx.Start("createColumn");
                                    if (!default_column.IsActive)
                                    {
                                        default_column.Activate();
                                    }

                                    string levelstring = "4F";
                                    Level placeLevel = null;
                                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                                    collector.OfClass(typeof(Level));
                                    foreach (Level level_1 in collector.Cast<Level>())
                                    {
                                        //MessageBox.Show(level.Name.ToString());
                                        if (levelstring == level.Name.ToString())
                                        {
                                            placeLevel = level;
                                        }
                                    }
                                    
                                    FamilyInstance familyInstance = doc.Create.NewFamilyInstance(center, default_column, level, StructuralType.Column);
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
                            // 柱子的斜對角點
                            String columnSize = GetColumnSize(point_a, point_b);
                            ChangeColumnType(doc, columnSize);
                            
                        }
                    }
                }
            }
            transGroup.Assimilate();

            return Result.Succeeded;
        }
        // The end of the main code.

        public void ChangeColumnType(Document doc, String columnSize)
        {
            // 50x30
            char separator = 'x';
            string[] parts = columnSize.Split(separator);
            // Int32.Parse:字串轉整數
            // double.Parse：字串轉浮點數
            double width = double.Parse(parts[0]);
            double depth = double.Parse(parts[1]);
            // Size: 50x30

            FilteredElementCollector Collector = new FilteredElementCollector(doc);
            List<FamilySymbol> familySymbolList = Collector.OfClass(typeof(FamilySymbol))
            .OfCategory(BuiltInCategory.OST_Columns)
            .Cast<FamilySymbol>().ToList();

            Boolean IsColumnTypeExist = false;
            foreach (FamilySymbol fs in familySymbolList)
            {
                if(fs.Name != columnSize)
                {
                    continue;
                }
                else
                {
                    IsColumnTypeExist = true;
                    break;
                }
            }
            if (!IsColumnTypeExist)
            {
                using (Transaction t_createNewColumnType = new Transaction(doc, "Ｃreate New Column Type"))
                {
                    try
                    {
                        t_createNewColumnType.Start("createColumnType");

                        FamilySymbol default_FamilySymbol = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol))
                        .OfCategory(BuiltInCategory.OST_Columns)
                        .Cast<FamilySymbol>()
                        .FirstOrDefault(q => q.Name == "default") as FamilySymbol;

                        FamilySymbol newFamSym = default_FamilySymbol.Duplicate(columnSize) as FamilySymbol;
                        // set the radius to a new value:
                        IList<Parameter> pars = newFamSym.GetParameters("Depth");
                        pars[0].Set(Algorithm.CentimetersToUnits(depth));
                        IList<Parameter> pars_2 = newFamSym.GetParameters("Width");
                        pars_2[0].Set(Algorithm.CentimetersToUnits(width));

                        t_createNewColumnType.Commit();
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
                        t_createNewColumnType.RollBack();
                    }
                }
            }

            using (Transaction t = new Transaction(doc, "Change Column Type"))
            {
                t.Start();
                // the familyinstance you want to change -> e.g. "default" column
                List<FamilyInstance> columns = new FilteredElementCollector(doc, doc.ActiveView.Id)
               .OfClass(typeof(FamilyInstance))
               .OfCategory(BuiltInCategory.OST_Columns)
               .Cast<FamilyInstance>()
               .Where(q => q.Name == "default").ToList();

                // the target familyinstance(familysymbol) what it should be.     
                FamilySymbol newColumns = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Columns)
                .Cast<FamilySymbol>()
                .FirstOrDefault(q => q.Name == columnSize) as FamilySymbol;

                foreach (FamilyInstance column in columns)
                {
                    column.Symbol = newColumns;
                }

                t.Commit();
            }
        }

        private XYZ GetMiddlePoint(XYZ startPoint, XYZ endPoint)
        {
            XYZ MiddlePoint = (startPoint  + endPoint) / 2;
            return MiddlePoint;
        }

        //public string GetColumnSize(XYZ point_a, XYZ point_b)
        //{
        //    String columnSize = Math.Abs(Math.Round(Algorithm.UnitsToMillimeters(point_a.X - point_b.X)) / 10).ToString() + "x"
        //                        + Math.Abs(Math.Round(Algorithm.UnitsToMillimeters(point_a.Y - point_b.Y)) / 10).ToString();
        //    return columnSize;
        //}
        public string GetColumnSize(XYZ point_a, XYZ point_b)
        {
            String columnSize = Math.Abs(Math.Round(Algorithm.UnitsToMillimeters(point_a.X - point_b.X)) / 10).ToString() + "x"
                                + Math.Abs(Math.Round(Algorithm.UnitsToMillimeters(point_a.Y - point_b.Y)) / 10).ToString();
            return columnSize;
        }
    }
}
