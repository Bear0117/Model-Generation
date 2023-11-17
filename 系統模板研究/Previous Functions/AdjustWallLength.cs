using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Collections.Generic;


namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class AdjustWallLength : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            Document doc = uidoc.Document;
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiApp.Application;
            foreach (ElementId elemId in ids)
            {

                Element elem = doc.GetElement(elemId);
                if (elem is Wall)
                {
                    Wall wall = (Wall)elem;
                    Parameter wallLengthParam = wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM);
                    ElementId WalltypeId = wall.GetTypeId();
                    ElementId levelId = null;
                    string basconstraint = wall.get_Parameter(BuiltInParameter.WALL_BASE_CONSTRAINT).AsValueString();
                    FilteredElementCollector LevelFilter = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(Level));
                    foreach (Level level in LevelFilter)
                    {
                        if (level.Name.Contains(basconstraint))
                        {
                            levelId = level.Id;
                        }
                    }

                    LocationCurve Locationcurve = wall.Location as LocationCurve;
                    Curve curve = Locationcurve.Curve;
                    double extensionLength = 0.003;
                    Line line = curve as Line;
                    // 取得線段的方向向量
                    XYZ direction = line.Direction;

                    // 將方向向量標準化
                    direction = direction.Normalize();

                    // 取得線段的起點與終點
                    XYZ start = line.GetEndPoint(0);
                    XYZ end = line.GetEndPoint(1);

                    // 分別計算起點與終點的延長點
                    XYZ startExtension = start - direction * extensionLength;
                    XYZ endExtension = end + direction * extensionLength;

                    // 建立延長後的線段
                    Line extendedLine = Line.CreateBound(startExtension, endExtension);
                    curve = extendedLine as Curve;
                    using (Transaction t = new Transaction(doc, "My transaction"))
                    {
                        t.Start();


                        try
                        {

                            Wall w = Wall.Create(doc, curve, WalltypeId, levelId, wallLengthParam.AsDouble(), 0, false, false);
                            WallUtils.DisallowWallJoinAtEnd(w, 0);
                            WallUtils.DisallowWallJoinAtEnd(w, 1);
                            ICollection<ElementId> iids = doc.Delete(elemId);
                            FilteredElementCollector collector = ElementCollector(w, doc);
                            foreach (Element e in collector)
                            {
                                if (e.Category.BuiltInCategory.Equals(BuiltInCategory.OST_Columns))
                                {

                                    try
                                    {
                                        RunJoinGeometryAndSwitch(doc, e, w);

                                    }

                                    catch { }

                                }
                                else if (e.Category.BuiltInCategory.Equals(BuiltInCategory.OST_StructuralFraming))
                                {
                                    try
                                    {
                                        RunJoinGeometry(doc, e, w);
                                    }

                                    catch { }
                                }

                            }

                        }
                        catch { }

                        t.Commit();

                    }
                }
                else if (elem.Category.BuiltInCategory.Equals(BuiltInCategory.OST_StructuralFraming))
                {
                    FamilyInstance beam = elem as FamilyInstance;
                    Parameter BeamLevel = beam.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM);
                    Level levelId = null;
                    FilteredElementCollector LevelFilter = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfClass(typeof(Level));
                    FilteredElementCollector collector = ElementCollector(beam, doc);
                    foreach (Level level in LevelFilter)
                    {

                        if (level.Name.Contains(BeamLevel.AsValueString()))
                        {
                            levelId = level;
                        }
                    }

                    LocationCurve Locationcurve = beam.Location as LocationCurve;
                    Curve curve = Locationcurve.Curve;
                    double extensionLength = 0.001;
                    Line line = curve as Line;
                    // 取得線段的方向向量
                    XYZ direction = line.Direction;

                    // 將方向向量標準化
                    direction = direction.Normalize();

                    // 取得線段的起點與終點
                    XYZ start = line.GetEndPoint(0);
                    XYZ end = line.GetEndPoint(1);

                    // 分別計算起點與終點的延長點
                    XYZ startExtension = start - direction * extensionLength;
                    XYZ endExtension = end + direction * extensionLength;

                    // 建立延長後的線段
                    Line extendedLine = Line.CreateBound(startExtension, endExtension);
                    curve = extendedLine as Curve;

                    //using (Transaction t = new Transaction(doc, "My transaction"))
                    //{

                    //    t.Start();

                    //    try
                    //    {

                    //        FamilyInstance Beam = doc.Create.NewFamilyInstance(curve, beam.Symbol, levelId, Autodesk.Revit.DB.Structure.StructuralType.Beam);
                    //        ICollection<ElementId> iids = doc.Delete(elemId);

                    //    }
                    //    catch { }

                    //    t.Commit();

                    //}
                    using (Transaction tr = new Transaction(doc, "My Join"))
                    {
                        tr.Start();
                        foreach (Element e in collector)
                        {
                            if (e.Category.BuiltInCategory.Equals(BuiltInCategory.OST_Columns))
                            {

                                try
                                {
                                    JoinGeometryUtils.JoinGeometry(doc, e, beam);

                                }

                                catch { }

                            }
                            if (e.Category.BuiltInCategory.Equals(BuiltInCategory.OST_StructuralFraming))
                            {
                                try
                                {
                                    JoinGeometryUtils.JoinGeometry(doc, e, beam);

                                }

                                catch { }
                            }

                        }
                        tr.Commit();

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

        // 執行接合並改變接合順序
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


        static public FilteredElementCollector ElementCollector(Autodesk.Revit.DB.Element elem, Document doc)
        {
            BoundingBoxXYZ bb = elem.get_BoundingBox(doc.ActiveView);
            Outline outline = new Outline(bb.Min, bb.Max);
            BoundingBoxIntersectsFilter bbfilter = new BoundingBoxIntersectsFilter(outline);
            FilteredElementCollector collector = new FilteredElementCollector(doc, doc.ActiveView.Id);
            ICollection<ElementId> idsExclude = new List<ElementId>();
            idsExclude.Add(elem.Id);
            collector.Excluding(idsExclude).WherePasses(bbfilter);

            return collector;
        }
    }
}
   
