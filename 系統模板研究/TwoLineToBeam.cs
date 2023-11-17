using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class TwoLineToBeam : IExternalCommand
    {
        Autodesk.Revit.ApplicationServices.Application app;
        Document doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            Application app = uiapp.Application;
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("1F")) as Level;

            double beamWidth = 3300 * 0.003281;
            double beamDepth = 500 * 0.003281;
            double top = 230 * 0.003281;
            double width = 150 * 0.003281;
            String beamSize = "30x50";

            

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
                width = Math.Abs(Line2.Distance(points[0]));
                if (Math.Abs(width - 150 / 304.8) < 1 * 0.003281)
                {
                    beamSize = "15x30";
                }
                else if (Math.Abs(width - 250 / 304.8) < 1 * 0.003281)
                {
                    beamSize = "25x40";
                }
                else if (Math.Abs(width - 275 / 304.8) < 1 * 0.003281)
                {
                    beamSize = "27.5x30";
                }
                else if (Math.Abs(width - 300 / 304.8) < 1 * 0.003281)
                {
                    beamSize = "30x50";
                }
                else if (Math.Abs(width - 350 / 304.8) < 1 * 0.003281)
                {
                    beamSize = "35x50";
                }
                else if (Math.Abs(width - 375 / 304.8) < 1 * 0.003281)
                {
                    beamSize = "37.5x50";
                }
                else if (Math.Abs(width - 430 / 304.8) < 1 * 0.003281)
                {
                    beamSize = "43x50";
                }
                else if (Math.Abs(width - 550 / 304.8) < 1 * 0.003281)
                {
                    beamSize = "55x75";
                }
                else if (Math.Abs(width - 600 / 304.8) < 1 * 0.003281)
                {
                    beamSize = "60x75";
                }
                else if (Math.Abs(width - 650 / 304.8) < 1 * 0.003281)
                {
                    beamSize = "65x75";
                }
                else if (Math.Abs(width - 725 / 304.8) < 1 * 0.003281)
                {
                    beamSize = "72.5x50";
                }
                else if (Math.Abs(width - 740 / 304.8) < 1 * 0.003281)
                {
                    beamSize = "74x75";
                }
                else if (Math.Abs(width - 750 / 304.8) < 1 * 0.003281)
                {
                    beamSize = "75x75";
                }
                else if (Math.Abs(width - 800 / 304.8) < 1 * 0.003281)
                {
                    beamSize = "80x75";
                }
                else if (Math.Abs(width - 850 / 304.8) < 1 * 0.003281)
                {
                    beamSize = "85x75";
                }
                else if (Math.Abs(width - 875 / 304.8) < 1 * 0.003281)
                {
                    beamSize = "87.5x75";
                }
                else
                {
                    break;
                }

                //TaskDialog.Show("線條資訊", width.ToString(), TaskDialogCommonButtons.Ok);
                TaskDialog.Show("線條資訊", beamSize, TaskDialogCommonButtons.Ok);
                XYZ midPoint1 = midPoint(points[0], points[3]);
                XYZ midPoint2 = midPoint(points[1], points[2]);


                Line BreakLine = Line.CreateBound(points[0], points[2]);
                Line STDLine1 = Line.CreateBound(points[0], points[1]);
                Line STDLine2 = Line.CreateBound(points[1], points[3]);
                //1mm = 0.03937 inch
                if (Math.Abs(Line1.Length - Line2.Length) < 2/304.8)
                {
                    //TaskDialog.Show("線條資訊","兩線等長", TaskDialogCommonButtons.Ok);
                    if (midPoint1.DistanceTo(midPoint2) < 5 / 304.78)
                    {
                        //TaskDialog.Show("線條資訊", "兩線中點重疊", TaskDialogCommonButtons.Ok);
                        midPoint1 = midPoint(points[0], points[2]);
                        midPoint2 = midPoint(points[1], points[3]);
                        
                    }
                }
                else if (Line1.Length > Line2.Length)
                {
                    //TaskDialog.Show("線條資訊", "第二條線較短", TaskDialogCommonButtons.Ok);
                    XYZ UnitZ = new XYZ(0, 0, 1);
                    XYZ Vector2 = points[3] - points[2];
                    XYZ offset = Cross(UnitZ, Vector2) / Line2.Length * width / 2;
                    //TaskDialog.Show("線條資訊", offset.ToString(), TaskDialogCommonButtons.Ok);
                    if (Dot(offset, points[1] - points[2]) > 0)
                    {
                        midPoint1 = points[2] + offset;
                        midPoint2 = points[3] + offset;
                    }
                    else
                    {
                        midPoint1 = points[2] - offset;
                        midPoint2 = points[3] - offset;
                    }

                }
                else
                {
                    //TaskDialog.Show("線條資訊", "第一條線較短", TaskDialogCommonButtons.Ok);
                    XYZ UnitZ = new XYZ(0, 0, 1);
                    XYZ Vector1 = points[1] - points[0];
                    XYZ offset = Cross(UnitZ, Vector1) / Line1.Length * width / 2;
                    //TaskDialog.Show("線條資訊", offset.ToString(), TaskDialogCommonButtons.Ok);
                    if (Dot(offset, points[2] - points[1]) > 0)
                    {
                        midPoint1 = points[1] + offset;
                        midPoint2 = points[0] + offset;
                    }
                    else
                    {
                        midPoint1 = points[1] - offset;
                        midPoint2 = points[0] - offset;
                    }

                }
                foreach (XYZ p in points)
                {
                    //TaskDialog.Show("線條資訊", p.ToString(), TaskDialogCommonButtons.Ok);
                }
                Line wallLine = Line.CreateBound(midPoint1, midPoint2);
                // get a family symbol
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                collector.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralFraming);

                FamilySymbol gotSymbol = collector.FirstElement() as FamilySymbol;
                Transaction trans = new Transaction(doc, "自動建梁");
                trans.Start();
                FamilyInstance instance = doc.Create.NewFamilyInstance(wallLine, gotSymbol, level, Autodesk.Revit.DB.Structure.StructuralType.Beam);
                
                trans.Commit();
                changeBeamType(doc, beamSize);

            }
            return Result.Succeeded;
        }
        public FamilyInstance CreateBeam(Document document, Line beamLine)
        {
            // get the given view's level for beam creation
            Level level = new FilteredElementCollector(document).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("1F")) as Level;
            // get a family symbol
            FilteredElementCollector collector = new FilteredElementCollector(document);
            collector.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralFraming);

            FamilySymbol gotSymbol = collector.FirstElement() as FamilySymbol;
            //FamilySymbol BeamTypeName = document.GetElement(new ElementId(342873)) as FamilySymbol;
            Transaction trans = new Transaction(document, "自動建梁");
            trans.Start();
            //CADModel cADModel = new CADModel();
            //cADModel.curveArray = curveArray;
            //cADModel.length = newLine.Length;
            //cADModel.shape = "矩形梁";
            //cADModel.width = 300 / 304.8;
            //cADModel.location = MiddlePoint;
            //cADModel.rotation = rotation;
            FamilyInstance instance = doc.Create.NewFamilyInstance(beamLine, gotSymbol, level, Autodesk.Revit.DB.Structure.StructuralType.Beam);// create a new beam
            //Parameter offset = instance.get_Parameter(BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION);
            //Parameter offset1 = instance.get_Parameter(BuiltInParameter.STRUCTURAL_BEAM_END1_ELEVATION);
            //offset.Set(0.003281 * (3300-50));
            //offset1.Set(0.003281 * (3300-50));

            trans.Commit();
            return instance;
        }
        public void changeBeamType(Document doc, String beamSize)
        {
            using (Transaction t = new Transaction(doc, "Change Beam Type"))
            {
                t.Start();
                // the familyinstance you want to change -> e.g. UB-Universal Beam 305x165x40UB
                List<FamilyInstance> beams = new FilteredElementCollector(doc, doc.ActiveView.Id)
               .OfClass(typeof(FamilyInstance))
               .OfCategory(BuiltInCategory.OST_StructuralFraming)
               .Cast<FamilyInstance>()
               .Where(q => q.Name == "W310X38.7").ToList();

                // the target familyinstance(familysymbol) what it should be -> e.g. UB-Universal Beam 533x210x92UB         
                FamilySymbol fs = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .Cast<FamilySymbol>()
                .FirstOrDefault(q => q.Name == beamSize) as FamilySymbol;

                foreach (FamilyInstance beam in beams)
                {
                    beam.Symbol = fs;
                }

                t.Commit();
            }
        }

        private XYZ midPoint(XYZ a, XYZ b)
        {
            XYZ c = (a + b) / 2;
            return c;
        }

        private Line TransformLine(Transform transform, Line line)
        {
            XYZ startPoint = transform.OfPoint(line.GetEndPoint(0));
            XYZ endPoint = transform.OfPoint(line.GetEndPoint(1));
            Line newLine = Line.CreateBound(startPoint, endPoint);
            return newLine;
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
    }
}