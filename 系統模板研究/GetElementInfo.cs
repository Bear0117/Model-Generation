using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms.VisualStyles;

namespace FormworkPlanning.Accessory
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class GetElementInfo : IExternalCommand
    {
        public StringBuilder st = new StringBuilder();

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            this.Run(commandData.Application, "1F");
            return Result.Succeeded;
        }

        // 輔助函式：英尺轉公分並四捨五入
        private double FootToCentimeter(double feet)
        {
            // 1 ft = 30.48 cm
            double cm = feet * 30.48;

            // 除以 5、四捨、乘回 5
            return Math.Round(cm / 5.0, MidpointRounding.AwayFromZero) * 5.0;
        }

        public class SelectFilter : ISelectionFilter
        {
            public Document Document { get; set; }
            public bool AllowElement(Element element)
            {
                int id = element.Category.Id.IntegerValue;
                return id == (int)BuiltInCategory.OST_GenericModel
                    || id == (int)BuiltInCategory.OST_StructuralFraming
                    || id == (int)BuiltInCategory.OST_Walls
                    || id == (int)BuiltInCategory.OST_Windows
                    || id == (int)BuiltInCategory.OST_Doors
                    || id == (int)BuiltInCategory.OST_Floors
                    || id == (int)BuiltInCategory.OST_StructuralColumns;
            }

            public bool AllowReference(Reference refer, XYZ point)
            {
                return false;
            }
        }

        public void Run(UIApplication uiapp, String levelstring)
        {
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;
            SelectFilter selectFilter = new SelectFilter { Document = doc };
            IList<Reference> refs = uidoc.Selection.PickObjects(ObjectType.Element, selectFilter);

            List<Element> elems = new List<Element>();
            foreach (Reference reference in refs)
            {
                Element element = doc.GetElement(reference);
                elems.Add(element);
            }

            foreach (Element element in elems)
            {
                // Wall
                if (element is Wall wall)
                {
                    var locCurve = wall.Location as LocationCurve;
                    if (locCurve != null)
                    {
                        var curve = locCurve.Curve;
                        XYZ startFt = curve.GetEndPoint(0);
                        XYZ endFt = curve.GetEndPoint(1);
                        double startX = FootToCentimeter(startFt.X);
                        double startY = FootToCentimeter(startFt.Y);
                        double endX = FootToCentimeter(endFt.X);
                        double endY = FootToCentimeter(endFt.Y);
                        TaskDialog.Show("Wall Info",
                            $"Id: {wall.Id}\n" +
                            $"Start: ({startX}cm, {startY}cm)\n" +
                            $"End:   ({endX}cm, {endY}cm)"
                        );
                    }
                    double thicknessCm = FootToCentimeter(wall.Width);
                    double heightCm = FootToCentimeter(
                        wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble());
                    TaskDialog.Show("Wall Params",
                        $"Thickness: {thicknessCm} cm\n" +
                        $"Height:    {heightCm} cm"
                    );
                }
                // Window
                else if (element is FamilyInstance fi &&
                         fi.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Windows)
                {
                    var locPt = fi.Location as LocationPoint;
                    XYZ ptFt = locPt.Point;
                    double xCm = FootToCentimeter(ptFt.X);
                    double yCm = FootToCentimeter(ptFt.Y);
                    var widthCm = FootToCentimeter(fi.get_Parameter(BuiltInParameter.WINDOW_WIDTH).AsDouble());
                    var heightCm = FootToCentimeter(fi.get_Parameter(BuiltInParameter.WINDOW_HEIGHT).AsDouble());
                    TaskDialog.Show("Window Info",
                        $"Location: ({xCm}cm, {yCm}cm)\n" +
                        $"Width:    {widthCm} cm\n" +
                        $"Height:   {heightCm} cm"
                    );
                }
                // Door
                else if (element is FamilyInstance di &&
                         di.Category.Id.IntegerValue == (int)BuiltInCategory.OST_Doors)
                {
                    var locPt = di.Location as LocationPoint;
                    XYZ ptFt = locPt.Point;
                    double xCm = FootToCentimeter(ptFt.X);
                    double yCm = FootToCentimeter(ptFt.Y);
                    var widthCm = FootToCentimeter(di.get_Parameter(BuiltInParameter.DOOR_WIDTH).AsDouble());
                    var heightCm = FootToCentimeter(di.get_Parameter(BuiltInParameter.DOOR_HEIGHT).AsDouble());
                    TaskDialog.Show("Door Info",
                        $"Location: ({xCm}cm, {yCm}cm)\n" +
                        $"Width:    {widthCm} cm\n" +
                        $"Height:   {heightCm} cm"
                    );
                }
                // Floor
                else if (element is Floor floor)
                {
                    // 1. 取得樓板頂面參考
                    IList<Reference> topFaces = HostObjectUtils.GetTopFaces(floor);

                    // 2. 逐一處理每個面
                    foreach (Reference faceRef in topFaces)
                    {
                        Face topFace = floor.GetGeometryObjectFromReference(faceRef) as Face;
                        if (topFace == null) continue;

                        // 3. 取得所有閉合曲線環
                        IList<CurveLoop> loops = topFace.GetEdgesAsCurveLoops();
                        foreach (CurveLoop loop in loops)
                        {
                            foreach (Curve edge in loop)
                            {
                                // 4. 擷取端點
                                XYZ ptFeet = edge.GetEndPoint(0);

                                // (假設已定義 FootToCentimeter 方法)
                                double xCm = FootToCentimeter(ptFeet.X);
                                double yCm = FootToCentimeter(ptFeet.Y);
                                var thicknessCm = FootToCentimeter(
                        floor.get_Parameter(BuiltInParameter.FLOOR_ATTR_THICKNESS_PARAM).AsDouble());
                                TaskDialog.Show("Floor Info",
                                    //$"Loops:     {loop.Count()}\n" +
                                    //$"Points:    {ptsCm.Count}\n" +
                                    $"Thickness: {thicknessCm} cm"
                                );
                                // 輸出或儲存……
                            }
                        }
                    }
                    
                }
                // Column
                else if (element is FamilyInstance col &&
                         col.Category.Id.IntegerValue == (int)BuiltInCategory.OST_StructuralColumns)
                {
                    var locPt = col.Location as LocationPoint;
                    XYZ ptFt = locPt.Point;
                    double xCm = FootToCentimeter(ptFt.X);
                    double yCm = FootToCentimeter(ptFt.Y);
                    var widthCm = FootToCentimeter(col.get_Parameter(BuiltInParameter.FAMILY_WIDTH_PARAM).AsDouble());
                    var depthCm = FootToCentimeter(col.get_Parameter(BuiltInParameter.GENERIC_DEPTH).AsDouble());
                    TaskDialog.Show("Column Info",
                        $"Location: ({xCm}cm, {yCm}cm)\n" +
                        $"Width:    {widthCm} cm\n" +
                        $"Depth:    {depthCm} cm"
                    );
                }
            }
        }
    }
}
