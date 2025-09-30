// GetElementInfo.cs (full amended version)
// ------------------------------------------------------------
// 依據使用者需求：
// 1. 自動掃描整個模型中的 OST_Walls / OST_Windows / OST_Doors / OST_Floors。
// 2. 讀取 ./ElementInfo/<scene>_bbox.json（由 pt2bbox_json.py 產生，單位 m）。
// 3. 以 AABB 重疊體積最大原則指派 instance_id；重疊比例 <1% 忽略。
// 4. 組裝為 elementInfo.json 格式輸出至 ./ElementInfo/<scene>_elementInfo.json
// ------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;

namespace Modeling
{
    [Transaction(TransactionMode.Manual)]
    public class GetElementInfo : IExternalCommand
    {
        // 1 ft = 0.3048 m
        private const double Ft2M = 0.3048;
        private const double Ft2Cm = 30.48;

        #region ──────────────── 入口 ────────────────
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                Run(commandData.Application);
                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("GetElementInfo", ex.ToString());
                return Result.Failed;
            }
        }
        #endregion

        #region ──────────────── 主流程 ────────────────
        private void Run(UIApplication uiapp)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            // 1. 決定 sceneId 與 bbox JSON 路徑
            string sceneId = GetSceneId(doc);                       // 如 "scene0046"
            string exeDir = Path.GetDirectoryName(doc.PathName) ?? string.Empty;
            // 若要向上兩層找 ElementInfo 資料夾：
            string twoUp = Path.GetFullPath(Path.Combine(exeDir, "..", ".."));
            string inputDir = Path.Combine(twoUp, "ElementInfo"); // 改於上兩層
            string bboxPath = Path.Combine(inputDir, $"{sceneId}_bbox.json");
            if (!File.Exists(bboxPath))
                throw new FileNotFoundException($"找不到 bbox 檔：{bboxPath}\n請先執行 pt2bbox_json.py 產生。\n");

            // 2. 讀 bounding box JSON (meter)
            Dictionary<string, List<InstanceBBox>> bboxData = JsonConvert.DeserializeObject<Dictionary<string, List<InstanceBBox>>>(File.ReadAllText(bboxPath));
            // bboxData["walls"] / ["floors"] / ["doors"] / ["windows"]
            MessageBox.Show("JSON file loaded.");

            // 3. 收集 Revit 元件
            IEnumerable<Wall> walls = Collect<Wall>(doc, BuiltInCategory.OST_Walls);
            IEnumerable<Floor> floors = Collect<Floor>(doc, BuiltInCategory.OST_Floors);
            IEnumerable<FamilyInstance> doors = Collect<FamilyInstance>(doc, BuiltInCategory.OST_Doors);
            IEnumerable<FamilyInstance> windows = Collect<FamilyInstance>(doc, BuiltInCategory.OST_Windows);
            MessageBox.Show("Collect all the elements!");

            // 4. 建立 element → info 物件
            SceneInfo sceneInfo = new SceneInfo();

            // 4‑1 牆
            foreach (Wall w in walls) sceneInfo.walls.Add(BuildWallInfo(w));
            // 4‑2 樓板
            foreach (Floor f in floors) sceneInfo.floors.Add(BuildFloorInfo(f));
            // 4‑3 窗 & 門
            foreach (FamilyInstance wi in windows) sceneInfo.window.Add(BuildWindowOrDoorInfo(wi, isDoor: false));
            foreach (FamilyInstance di in doors) sceneInfo.door.Add(BuildWindowOrDoorInfo(di, isDoor: true));
            MessageBox.Show("Info written.");

            // 5. 執行 instance_id 分派
            AssignInstanceIds(sceneInfo, bboxData, 0.01 /*1%*/);
            MessageBox.Show("InstanceIds assigned.");


            // 6. 轉換 host_ids 為對應牆的 instance_ids
            UpdateHostIds(sceneInfo);
            MessageBox.Show("Host instanceIds updated.");


            // 7. 輸出 JSON
            Directory.CreateDirectory(inputDir);
            string outPath = Path.Combine(inputDir, $"{sceneId}_elementInfo.json");
            Dictionary<string, SceneInfo> wrapper = new Dictionary<string, SceneInfo> { [$"{sceneId}"] = sceneInfo };
            File.WriteAllText(outPath, JsonConvert.SerializeObject(wrapper, Formatting.Indented));
            TaskDialog.Show("GetElementInfo", $"輸出完成 → {outPath}");
        }
        #endregion

        #region ──────────────── 資料結構 ────────────────
        private class InstanceBBox
        {
            public int instance_id { get; set; }
            public List<double> Bbox_m { get; set; } // [cx,cy,cz,dx,dy,dz]

            // JSON 存的是中心 + 尺寸 (m) → 轉 Revit feet min/max
            public BoundingBoxXYZ ToBboxFt()
            {
                double cx = Bbox_m[0], cy = Bbox_m[1], cz = Bbox_m[2];
                double dx = Bbox_m[3], dy = Bbox_m[4], dz = Bbox_m[5];
                double hx = dx / 2.0, hy = dy / 2.0, hz = dz / 2.0;
                XYZ min = new XYZ((cx - hx) / Ft2M, (cy - hy) / Ft2M, (cz - hz) / Ft2M);
                XYZ max = new XYZ((cx + hx) / Ft2M, (cy + hy) / Ft2M, (cz + hz) / Ft2M);
                return new BoundingBoxXYZ { Min = min, Max = max };
            }
        }

        private class SceneInfo
        {
            public List<WallJson> walls = new List<WallJson>();
            public List<FloorJson> floors = new List<FloorJson>();
            public List<WinDoorJson> window = new List<WinDoorJson>();
            public List<WinDoorJson> door = new List<WinDoorJson>();
        }

        private class WallJson
        {
            public List<int> instance_ids = new List<int>();
            public List<Segment> segments = new List<Segment>();
            public double height_cm;
            public double thickness_cm;
            public int revit_id; // 用來轉 host_ids
            [JsonIgnore] public BoundingBoxXYZ bbox; // Revit 內建
        }
        private class Segment { public List<double> start; public List<double> end; }

        private class FloorJson
        {
            public List<int> instance_ids = new List<int>();
            public List<List<double>> outline = new List<List<double>>();
            public double thickness_cm;
            [JsonIgnore] public BoundingBoxXYZ bbox;
        }

        private class WinDoorJson
        {
            public List<int> instance_ids = new List<int>();
            public List<double> location; // XYZ (cm)
            public double width_cm;
            public double height_cm;
            public List<int> host_ids = new List<int>(); // 先放 Revit Id，稍後轉換
            [JsonIgnore] public BoundingBoxXYZ bbox;
        }
        #endregion

        #region ──────────────── Collectors & Builders ────────────────
        private static IEnumerable<T> Collect<T>(Document doc, BuiltInCategory cat) where T : Element
        {
            return new FilteredElementCollector(doc)
                .OfCategory(cat)
                .WhereElementIsNotElementType()
                .Cast<T>();
        }

        private static WallJson BuildWallInfo(Wall w)
        {
            WallJson res = new WallJson();
            LocationCurve lc = w.Location as LocationCurve;
            if (lc != null)
            {
                Curve c = lc.Curve;
                XYZ sFt = c.GetEndPoint(0);
                XYZ eFt = c.GetEndPoint(1);
                res.segments.Add(new Segment { start = XYZ2cm(sFt), end = XYZ2cm(eFt) });
            }
            res.height_cm = Math.Round(Ft2Cm * w.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble(), 1);
            res.thickness_cm = Math.Round(Ft2Cm * w.Width, 1);
            res.revit_id = w.Id.IntegerValue;
            res.bbox = w.get_BoundingBox(null);
            return res;
        }

        private static FloorJson BuildFloorInfo(Floor f)
        {
            FloorJson res = new FloorJson();
            IList<Reference> topFaces = HostObjectUtils.GetTopFaces(f);
            if (topFaces?.Count > 0)
            {
                Face face = f.GetGeometryObjectFromReference(topFaces[0]) as Face;
                foreach (CurveLoop loop in face.GetEdgesAsCurveLoops())
                {
                    foreach (Curve edge in loop)
                    {
                        XYZ p = edge.GetEndPoint(0);
                        res.outline.Add(XYZ2cm(p));
                    }
                    break;
                }
            }
            res.thickness_cm = Ft2Cm * f.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM).AsDouble();
            res.bbox = f.get_BoundingBox(null);
            return res;
        }

        //private static WinDoorJson BuildWindowOrDoorInfo(FamilyInstance fi, bool isDoor)
        //{
        //    WinDoorJson res = new WinDoorJson();
        //    LocationPoint lp = fi.Location as LocationPoint;
        //    res.location = XYZ2cm(lp.Point);
        //    BuiltInParameter wParam = isDoor ? BuiltInParameter.DOOR_WIDTH : BuiltInParameter.WINDOW_WIDTH;
        //    BuiltInParameter hParam = isDoor ? BuiltInParameter.DOOR_HEIGHT : BuiltInParameter.WINDOW_HEIGHT;
        //    res.width_cm = Ft2Cm * fi.get_Parameter(wParam).AsDouble();
        //    res.height_cm = Ft2Cm * fi.get_Parameter(hParam).AsDouble();
        //    if (fi.Host != null) res.host_ids.Add(fi.Host.Id.IntegerValue);
        //    res.bbox = fi.get_BoundingBox(null);
        //    return res;
        //}
        private static WinDoorJson BuildWindowOrDoorInfo(FamilyInstance fi, bool isDoor)
        {
            WinDoorJson res = new WinDoorJson();
            LocationPoint lp = fi.Location as LocationPoint;
            res.location = XYZ2cm(lp.Point);

            // ► 改用「參數名稱」抓取尺寸
            string wName = isDoor ? "粗略寬度" : "寬度";
            string hName = isDoor ? "粗略高度" : "高度";

            Parameter wPara = fi.Symbol.LookupParameter(wName);
            Parameter hPara = fi.Symbol.LookupParameter(hName);

            double widthFt = wPara != null ? wPara.AsDouble() : 0;
            double heightFt = hPara != null ? hPara.AsDouble() : 0;

            res.width_cm = Ft2Cm * widthFt;
            res.height_cm = Ft2Cm * heightFt;

            // 其餘維持原樣
            if (fi.Host != null) res.host_ids.Add(fi.Host.Id.IntegerValue);
            res.bbox = fi.get_BoundingBox(null);
            return res;
        }

        private static List<double> XYZ2cm(XYZ p) => new List<double> { Math.Round(p.X * Ft2Cm), Math.Round(p.Y * Ft2Cm), Math.Round(p.Z * Ft2Cm) };
        #endregion

        #region ──────────────── Instance Matching ────────────────
        private static void AssignInstanceIds(SceneInfo info, Dictionary<string, List<InstanceBBox>> bboxData, double thresh)
        {
            Dictionary<object, BoundingBoxXYZ> elementBBoxes = new Dictionary<object, BoundingBoxXYZ>();
            void Add(object key, BoundingBoxXYZ bb) { if (bb != null) elementBBoxes[key] = bb; }

            foreach (WallJson w in info.walls) Add(w, w.bbox);
            foreach (FloorJson f in info.floors) Add(f, f.bbox);
            foreach (WinDoorJson wd in info.window) Add(wd, wd.bbox);
            foreach (WinDoorJson d in info.door) Add(d, d.bbox);

            foreach (KeyValuePair<string, List<InstanceBBox>> kv in bboxData)
            {
                List<InstanceBBox> list = kv.Value; // cat 未使用
                foreach (InstanceBBox inst in list)
                {
                    BoundingBoxXYZ instBBFt = inst.ToBboxFt();
                    double instVol = Volume(instBBFt);
                    double bestOverlap = 0;
                    object bestElem = null;
                    foreach (KeyValuePair<object, BoundingBoxXYZ> kvElem in elementBBoxes)
                    {
                        object elem = kvElem.Key;
                        BoundingBoxXYZ bb = kvElem.Value;
                        double ov = IntersectionVolume(bb, instBBFt);
                        if (ov > bestOverlap) { bestOverlap = ov; bestElem = elem; }
                    }
                    if (bestElem != null && (bestOverlap / instVol) >= thresh)
                    {
                        switch (bestElem)
                        {
                            case WallJson w: w.instance_ids.Add(inst.instance_id); break;
                            case FloorJson f: f.instance_ids.Add(inst.instance_id); break;
                            case WinDoorJson wd when info.window.Contains(wd): wd.instance_ids.Add(inst.instance_id); break;
                            case WinDoorJson d when info.door.Contains(d): d.instance_ids.Add(inst.instance_id); break;
                        }
                    }
                }
            }
        }
        #endregion

        #region ──────────────── AABB & Host 轉換 ────────────────
        private static void UpdateHostIds(SceneInfo info)
        {
            Dictionary<int, List<int>> wallMap = info.walls.ToDictionary(w => w.revit_id, w => w.instance_ids);
            IEnumerable<WinDoorJson> allWD = info.window.Concat(info.door);
            foreach (WinDoorJson wd in allWD)
            {
                List<int> newList = new List<int>();
                foreach (int hostRevitId in wd.host_ids)
                    if (wallMap.TryGetValue(hostRevitId, out List<int> instList)) newList.AddRange(instList);
                wd.host_ids.Clear();
                wd.host_ids.AddRange(newList.Distinct());
            }
        }

        private static double Volume(BoundingBoxXYZ bb) => (bb.Max.X - bb.Min.X) * (bb.Max.Y - bb.Min.Y) * (bb.Max.Z - bb.Min.Z);
        private static double IntersectionVolume(BoundingBoxXYZ a, BoundingBoxXYZ b)
        {
            double dx = Math.Max(0, Math.Min(a.Max.X, b.Max.X) - Math.Max(a.Min.X, b.Min.X));
            double dy = Math.Max(0, Math.Min(a.Max.Y, b.Max.Y) - Math.Max(a.Min.Y, b.Min.Y));
            double dz = Math.Max(0, Math.Min(a.Max.Z, b.Max.Z) - Math.Max(a.Min.Z, b.Min.Z));
            return dx * dy * dz;
        }
        #endregion

        #region ──────────────── 其他 ────────────────
        private static string GetSceneId(Document doc)
        {
            string name = Path.GetFileNameWithoutExtension(doc.PathName); // e.g. scene0046
            System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(name, @"scene\d{4}_00");
            if (!m.Success) throw new Exception("無法從檔名擷取 scene 編號，請確認 .rvt 檔名格式 (sceneXXXX.rvt)");
            return m.Value;
        }
        #endregion
    }
}
