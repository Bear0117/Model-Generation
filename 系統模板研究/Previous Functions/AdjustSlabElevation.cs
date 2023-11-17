using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AdjustSlabElevation : IExternalCommand
    {
        //Document doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();

            foreach(ElementId id in ids)
            {
                Floor floor = doc.GetElement(id) as Floor;
                Parameter distanceToLevelParam = floor.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM);
                
                double distanceToLevel = distanceToLevelParam.AsDouble();
                MessageBox.Show(distanceToLevel.ToString());
            }
            return Result.Succeeded;
        }
        // The end of the main code.

        public void ChangeColumnType(Document doc, String columnSize)
        {
            char separator = 'x';
            string[] parts = columnSize.Split(separator);
            double width = int.Parse(parts[0]);
            double depth = int.Parse(parts[1]);

            FilteredElementCollector Collector = new FilteredElementCollector(doc);
            List<FamilySymbol> familySymbolList = Collector.OfClass(typeof(FamilySymbol))
            .OfCategory(BuiltInCategory.OST_Columns)
            .Cast<FamilySymbol>().ToList();

            Boolean IsColumnTypeExist = false;
            foreach (FamilySymbol fs in familySymbolList)
            {
                if (fs.Name != columnSize)
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
                        pars[0].Set(CentimetersToUnits(depth));
                        IList<Parameter> pars_2 = newFamSym.GetParameters("Width");
                        pars_2[0].Set(CentimetersToUnits(width));

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

        public XYZ RoundPoint(XYZ point)
        {

            XYZ newPoint = new XYZ(
                CentimetersToUnits(Math.Round(UnitsToCentimeters(point.X) * 2) / 2),
                CentimetersToUnits(Math.Round(UnitsToCentimeters(point.Y) * 2) / 2),
                CentimetersToUnits(Math.Round(UnitsToCentimeters(point.Z) * 2) / 2)
                );
            return newPoint;
        }

        private XYZ GetMiddlePoint(XYZ startPoint, XYZ endPoint)
        {
            XYZ MiddlePoint = (startPoint + endPoint) / 2;
            return MiddlePoint;
        }

        public string GetColumnSize(XYZ point_a, XYZ point_b)
        {
            String columnSize = Math.Abs(Math.Round(UnitsToCentimeters(point_a.X - point_b.X))).ToString() + "x"
                                + Math.Abs(Math.Round(UnitsToCentimeters(point_a.Y - point_b.Y))).ToString();
            return columnSize;
        }

        public double UnitsToCentimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
        }

        public double CentimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Centimeters);
        }
    }
}
