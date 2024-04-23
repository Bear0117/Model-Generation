using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using Autodesk.Revit.Attributes;
using System.Windows.Controls;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Windows;
using MessageBox = System.Windows.MessageBox;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Security.Cryptography;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Windows.Forms;
using System.Windows.Input;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateColumns_Grid : IExternalCommand
    {
        Document doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // 任务对话框展示结果数量
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";  // 只显示JSON文件
                openFileDialog.Title = "Select a JSON file";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // 获取选择的文件路径
                    string filePath = openFileDialog.FileName;

                    // 可以添加代码来处理矩阵，如显示内容等
                    TaskDialog.Show("CSV Read", "CSV file has been read successfully.");
                }
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string jsonFilePath = Path.Combine(desktopPath, "Predict_Result.json");
            string jsonContent = File.ReadAllText(jsonFilePath);
            JObject data = JObject.Parse(jsonContent);

            // 提取 1F 和 2F 的矩陣
            var floors = data["DatasetConstruction"];
            List<int> matrix1F = new List<int>();
            List<int> matrix2F = new List<int>();

            foreach (var floor in floors)
            {
                if (floor["Floor"].ToString() == "1F")
                {
                    matrix1F = FlattenMatrix(floor["ColumnLayout"]);
                }
                else if (floor["Floor"].ToString() == "2F")
                {
                    matrix2F = FlattenMatrix(floor["ColumnLayout"]);
                }
            }

            // 合併矩陣
            List<int> combinedMatrix = new List<int>(matrix1F);
            combinedMatrix.AddRange(matrix2F);

            string displayText = string.Join(", ", combinedMatrix);

            Document doc = commandData.Application.ActiveUIDocument.Document;

            // 筛选所有柱线
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType();
            
            Level level;
            foreach (Element elem in collector)
            {
                level = doc.GetElement(elem.LevelId) as Level;
                break;
            }
            
            List<Curve> verticalLines = new List<Curve>();
            List<Curve> horizontalLines = new List<Curve>();

            int count_verti_grid = 0;
            int count_horiz_grid = 0;
            foreach (Autodesk.Revit.DB.Grid grid in collector)
            {
                Curve curve = grid.Curve;
                if (Math.Abs(curve.GetEndPoint(0).X - curve.GetEndPoint(1).X) < 1e-6)
                {
                    verticalLines.Add(curve);
                    count_verti_grid++;
                }
                else if (Math.Abs(curve.GetEndPoint(0).Y - curve.GetEndPoint(1).Y) < 1e-6)
                {
                    horizontalLines.Add(curve);
                    count_horiz_grid++;
                }
            }

            // 计算交点
            List<XYZ> intersections = new List<XYZ>();
            foreach (Curve vLine in verticalLines)
            {
                foreach (Curve hLine in horizontalLines)
                {
                    IntersectionResultArray results;
                    SetComparisonResult result = vLine.Intersect(hLine, out results);
                    if (result == SetComparisonResult.Overlap)
                    {
                        foreach (IntersectionResult ir in results)
                        {
                            intersections.Add(ir.XYZPoint);
                        }
                    }
                }
            }

            List<XYZ> sortedPoints = intersections.OrderByDescending(p => p.Y).ThenBy(p => p.X).ToList();
            //MessageBox.Show(sortedPoints.Count().ToString());
            //MessageBox.Show(combinedMatrix.Count().ToString());

            FamilySymbol default_column = new FilteredElementCollector(doc)
            .OfClass(typeof(FamilySymbol))
            .OfCategory(BuiltInCategory.OST_Columns)
            .Cast<FamilySymbol>()
            .FirstOrDefault(q => q.Name == "default") as FamilySymbol;


            for(int i = 0; i < combinedMatrix.Count(); i++)
            {
                using (Transaction tx = new Transaction(doc))
                {
                    try
                    {
                        tx.Start("createColumn");
                        if (!default_column.IsActive)
                        {
                            default_column.Activate();
                        }

                        string levelstring_1F = "1F";
                        string levelstring_2F = "2F";
                        Level placeLevel_1F = null;
                        Level placeLevel_2F = null;
                        FilteredElementCollector collector_level = new FilteredElementCollector(doc);
                        collector_level.OfClass(typeof(Level));
                        foreach (Level level_1 in collector_level.Cast<Level>())
                        {
                            if (levelstring_1F == level_1.Name.ToString())
                            {
                                placeLevel_1F = level_1;
                            }
                        }
                        foreach (Level level_1 in collector_level.Cast<Level>())
                        {
                            if (levelstring_2F == level_1.Name.ToString())
                            {
                                placeLevel_2F = level_1;
                            }
                        }

                        if (i < 36)
                        {
                            if (combinedMatrix[i] == 1)
                            {
                                FamilyInstance familyInstance = doc.Create.NewFamilyInstance(sortedPoints[i % 36], default_column, placeLevel_1F, StructuralType.Column);
                            }
                        }
                        else
                        {
                            if (combinedMatrix[i] == 1)
                            {
                                FamilyInstance familyInstance = doc.Create.NewFamilyInstance(sortedPoints[i % 36], default_column, placeLevel_2F, StructuralType.Column);
                            }
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

            //foreach (XYZ inter in intersections)
            //{
            //    using (Transaction tx = new Transaction(doc))
            //    {
            //        try
            //        {
            //            tx.Start("createColumn");
            //            if (!default_column.IsActive)
            //            {
            //                default_column.Activate();
            //            }

            //            string levelstring = "1F";
            //            Level placeLevel = null;
            //            FilteredElementCollector collector_level = new FilteredElementCollector(doc);
            //            collector_level.OfClass(typeof(Level));
            //            foreach (Level level_1 in collector_level.Cast<Level>())
            //            {
            //                if (levelstring == level_1.Name.ToString())
            //                {
            //                    placeLevel = level_1;
            //                }
            //            }
                        
            //            FamilyInstance familyInstance = doc.Create.NewFamilyInstance(inter, default_column, placeLevel, StructuralType.Column);
            //            tx.Commit();
            //        }
            //        catch (Exception ex)
            //        {
            //            TaskDialog td = new TaskDialog("error")
            //            {
            //                Title = "error",
            //                AllowCancellation = true,
            //                MainInstruction = "error",
            //                MainContent = "Error" + ex.Message,
            //                CommonButtons = TaskDialogCommonButtons.Close
            //            };
            //            td.Show();

            //            Debug.Print(ex.Message);
            //            tx.RollBack();
            //        }
            //    }
            //}


            //XYZ[,] matrix = new XYZ[count_verti_grid, count_horiz_grid];

            //for (int i = 0; i < count_verti_grid; i++)
            //{
            //    for(int j = 0; j < count_horiz_grid; j++)
            //    {
            //        matrix[i, 0] = intersections[i];
            //    }
            //}




            

            return Result.Succeeded;
        }
        // The end of the main code.

        private static List<int> FlattenMatrix(JToken matrix)
        {
            List<int> flatList = new List<int>();
            foreach (var row in matrix)
            {
                foreach (int num in row)
                {
                    flatList.Add(num);
                }
            }
            return flatList;
        }

        private string[,] ReadCsvToMatrix(string filePath)
        {
            var lines = File.ReadAllLines(filePath);
            int numRows = lines.Length;
            int numCols = lines[0].Split(',').Length;

            string[,] matrix = new string[numRows, numCols];

            for (int i = 0; i < numRows; i++)
            {
                string[] cols = lines[i].Split(',');
                for (int j = 0; j < numCols; j++)
                {
                    matrix[i, j] = cols[j];
                }
            }

            return matrix;
        }
        public void ChangeColumnType(Document doc, String columnSize)
        {
            // 50x30
            char separator = 'x';
            string[] parts = columnSize.Split(separator);
            // Int32.Parse:字串轉整數
            // double.Parse：字串轉浮點數
            double width = int.Parse(parts[0]);
            double depth = int.Parse(parts[1]);
            // Size: 50x30

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
            XYZ MiddlePoint = (startPoint + endPoint) / 2;
            return MiddlePoint;
        }

        public string GetColumnSize(XYZ point_a, XYZ point_b)
        {
            String columnSize = Math.Abs(Math.Round(Algorithm.UnitsToCentimeters(point_a.X - point_b.X))).ToString() + "x"
                                + Math.Abs(Math.Round(Algorithm.UnitsToCentimeters(point_a.Y - point_b.Y))).ToString();
            return columnSize;
        }

        public bool IsVertical(XYZ vector)
        {
            // 检查X和Y分量是否接近于0，这里使用1e-6作为容差
            return Math.Abs(vector.X) < 1e-6 && Math.Abs(vector.Y) < 1e-6;
        }
    }
}
