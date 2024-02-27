using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Forms;
using System.Collections.Generic;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.ApplicationServices;
using Application = Autodesk.Revit.ApplicationServices.Application;
using System;
using Autodesk.Revit.UI.Selection;
using System.IO;
using System.Text;
using System.Linq;
using Aspose.Cells.Charts;
using Aspose.Pdf.Facades;
using System.Windows.Media.Animation;
using Teigha.DatabaseServices;
using Solid = Autodesk.Revit.DB.Solid;
using Line = Autodesk.Revit.DB.Line;
using Face = Autodesk.Revit.DB.Face;
using Curve = Autodesk.Revit.DB.Curve;
using System.Windows.Media;
using Dimension = Autodesk.Revit.DB.Dimension;
using System.Drawing.Printing;
using Transform = Autodesk.Revit.DB.Transform;


namespace Modeling.AreaEstimation
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class Test2_Wei1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            Document doc = uidoc.Document;
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiApp.Application;
            Reference paramerterRef = uidoc.Selection.PickObject(ObjectType.Element);
            ElementId selectedElementId = paramerterRef.ElementId;
            Element selectedElement = doc.GetElement(selectedElementId);
            List<Element> elems = new List<Element>();
            String info = "";
            Stairs stair = (Stairs)selectedElement;

            Element panelElement1 = doc.GetElement(stair.Id);
            FamilyInstance familyInstance = panelElement1 as FamilyInstance;
            ElementId pickedtypeId = panelElement1.GetTypeId();
            ElementType family = doc.GetElement(pickedtypeId) as ElementType;

            IList<Element> stairLandingsList = new List<Element>();

            ICollection<ElementId> stairLandingsICollectionId = stair.GetStairsLandings();
            MessageBox.Show("平台:" + stairLandingsICollectionId.Count.ToString());
            ICollection<ElementId> stairRunsICollectionId = stair.GetStairsRuns();
            MessageBox.Show("梯段:" + stairRunsICollectionId.Count.ToString());

            //選到的樓梯裡的平台
            foreach (ElementId stairLandingsId in stairLandingsICollectionId)
            {
                Element elem = doc.GetElement(stairLandingsId);
                if (elem != null)
                {
                    stairLandingsList.Add(elem);
                }

            }
            stairLandingsList = stairLandingsList
        .OrderBy(elem => elem.LookupParameter("相對高度")?.AsDouble() ?? 0.0)
        .ToList();

            info += "平台數量:" + stairLandingsList.Count() + "\n";
            int x = 1;
            foreach (Element elem in stairLandingsList)
            {
                List<Solid> solids = GetElementSolidList(elem);
                info += "平台: " + x + "\n";
                foreach (Solid sol in solids)
                {
                    info += "面的数量: " + sol.Faces.Size + "\n";

                    foreach (Face face in sol.Faces)
                    {
                        XYZ faceNormal = face.ComputeNormal(new UV(0, 0));
                        //找出垂直於Z軸的面:頂面與底面，但是是水平
                        if (faceNormal.IsAlmostEqualTo(XYZ.BasisZ) || faceNormal.IsAlmostEqualTo(-XYZ.BasisZ))
                        {
                            double area = face.Area * 0.3048 * 0.3048;
                            info += "面積: " + area + " 平方米\n";
                        }
                    }
                }

                x++;
            }

            //foreach (ElementId elemId in ids)
            //{
            //    Element elem = doc.GetElement(elemId);
            //    elems.Add(elem);
            //}
            //List<Solid> solids = new List<Solid>();
            //FilteredElementCollector collector = FormworkClass.ElementCollector(elems[0], doc);
            //foreach (Solid sol in s)
            //{
            //    solids.Add(sol);
            //}
            //foreach (Element e in collector)
            //{
            //    List<Solid> s2 = FormworkClass.GetElementSolidList(e);
            //    foreach (Solid sol in s2)
            //    {
            //        solids.Add(sol);
            //    }
            //}

            //Solid solid = BooleanOperationsUtils.ExecuteBooleanOperation(solids[0], solids[1], BooleanOperationsType.Union);

            //Solid solid = s[0];
            //info += "面的数量: " + solid.Faces.Size + "\n";

            ////int faceId = 0; // 用于追踪面的 ID
            //foreach (Face face in solid.Faces)
            //{
            //    //info += "面 ID: " + faceId + "\n";

            //    //EdgeArrayArray edgeLoops = face.EdgeLoops;
            //    //foreach (EdgeArray edgeArray in edgeLoops)
            //    //{
            //    //    foreach (Edge edge in edgeArray)
            //    //    {
            //    //        info += "邊長: " + edge.ApproximateLength * 0.3048 + "\n";
            //    //    }
            //    //}

            //    info += "面積: " + face.Area * 0.3048 * 0.3048 + "\n";
            //    //faceId++;
            //}
            TaskDialog.Show("test", info);
            return Result.Succeeded;
        }
        static public List<Solid> GetElementSolidList(Element elem, Options opt = null, bool useOriginGeom4FamilyInstance = false)
        {
            if (null == elem)
            {
                return null;
            }
            if (null == opt)
                opt = new Options();
            opt.IncludeNonVisibleObjects = false;
            opt.DetailLevel = ViewDetailLevel.Medium;

            GeometryElement gElem;
            List<Solid> solidList = new List<Solid>();
            try
            {
                if (useOriginGeom4FamilyInstance && elem is FamilyInstance)
                {
                    // we transform the geometry to instance coordinate to reflect actual geometry 
                    FamilyInstance fInst = elem as FamilyInstance;
                    gElem = fInst.GetOriginalGeometry(opt);
                    Transform trf = fInst.GetTransform();
                    if (!trf.IsIdentity)
                        gElem = gElem.GetTransformed(trf);
                }
                else
                    gElem = elem.get_Geometry(opt);
                if (null == gElem)
                {
                    return null;
                }
                IEnumerator<GeometryObject> gIter = gElem.GetEnumerator();
                gIter.Reset();
                while (gIter.MoveNext())
                {
                    // solidList.AddRange(getSolids(gIter.Current));
                }
            }
            catch (Exception ex)
            {
                // In Revit, sometime get the geometry will failed.
                string error = ex.Message;
            }
            return solidList;
        }
    }
}
