using System;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json.Linq;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Grounding : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // 取得 Revit 文件
            Document doc = commandData.Application.ActiveUIDocument.Document;

            // 建立最外層 JSON 物件
            JObject resultJson = new JObject();

            // 定義各類別對應的 BuiltInCategory
            var categoryMapping = new[]
            {
                new { Name = "column", Category = BuiltInCategory.OST_Columns },
                new { Name = "wall", Category = BuiltInCategory.OST_Walls },
                new { Name = "beam", Category = BuiltInCategory.OST_StructuralFraming },
                new { Name = "slab", Category = BuiltInCategory.OST_Floors }, // 板類可視專案需求調整
                new { Name = "opening", Category = BuiltInCategory.OST_Windows } // 以窗戶作為 opening
            };

            // 逐一處理各個類別
            foreach (var mapping in categoryMapping)
            {
                // 利用 FilteredElementCollector 抓取該類別的元素
                var collector = new FilteredElementCollector(doc)
                                .OfCategory(mapping.Category)
                                .WhereElementIsNotElementType()
                                .ToElements();

                // 建立存放該類別元素資料的 JSON 物件
                JObject categoryJson = new JObject();
                int count = 1;

                foreach (Element elem in collector)
                {
                    // 取得元素的 3D bounding box
                    BoundingBoxXYZ bb = elem.get_BoundingBox(null);
                    if (bb == null)
                        continue;

                    // 計算寬度、高度、深度 (注意：這裡假設 X 為寬、Y 為深、Z 為高)
                    double width = Algorithm.UnitsToCentimeters(bb.Max.X - bb.Min.X);
                    double depth = Algorithm.UnitsToCentimeters(bb.Max.Y - bb.Min.Y);
                    double height = Algorithm.UnitsToCentimeters(bb.Max.Z - bb.Min.Z);

                    // 計算中心點 (bb.min 與 bb.max 的平均值)
                    double centerX = Algorithm.UnitsToCentimeters((bb.Min.X + bb.Max.X) / 2);
                    double centerY = Algorithm.UnitsToCentimeters((bb.Min.Y + bb.Max.Y) / 2);
                    double centerZ = Algorithm.UnitsToCentimeters((bb.Min.Z + bb.Max.Z) / 2);


                    // 四捨五入到小數點後 2 位
                    width = Math.Round(width, 2);
                    depth = Math.Round(depth, 2);
                    height = Math.Round(height, 2);

                    centerX = Math.Round(centerX, 2);
                    centerY = Math.Round(centerY, 2);
                    centerZ = Math.Round(centerZ, 2);

                    // 建立單一元素的 JSON 物件
                    JObject elementJson = new JObject
                    {
                        { "center", new JArray(centerX, centerY, centerZ) },
                        { "bounding box", new JArray(width, depth, height) }
                    };

                    // 加入該類別的 JSON 中
                    categoryJson[count.ToString()] = elementJson;
                    count++;
                }

                // 將該類別資料加入最終 JSON 物件
                resultJson[mapping.Name] = categoryJson;
            }

            // 設定輸出檔案路徑
            string desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            // 組合成「scan-to-BIM」資料夾路徑
            string folderPath = Path.Combine(desktopFolder, "scan-to-BIM");
            string outputPath = Path.Combine(folderPath, "bounding_boxes.json");


            // 若資料夾不存在則建立
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }

            try
            {
                File.WriteAllText(outputPath, resultJson.ToString());
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }

            TaskDialog.Show("完成", "Bounding Box 資料已輸出至：" + outputPath);
            return Result.Succeeded;
        }
    }
}
