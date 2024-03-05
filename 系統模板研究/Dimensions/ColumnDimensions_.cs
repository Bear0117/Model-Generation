using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows.Forms;
using System.Windows.Media.Animation;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Aspose.Cells.Drawing;
using Aspose.Pdf.Operators;
using Line = Autodesk.Revit.DB.Line;
using Teigha.GraphicsSystem;
using Autodesk.Revit.DB.Structure;
using Aspose.Pdf.Forms;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class ColumnDimensions_ : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();
            Document doc = uidoc.Document;


            ids = ids.OrderByDescending(e => doc.GetElement(e).get_BoundingBox(null).Max.Y).ToList();
            double height = uidoc.ActiveGraphicalView.GenLevel.Elevation;

            foreach (ElementId elemId in ids)
            {


                Element elem = doc.GetElement(elemId);
                if (elem.Category.BuiltInCategory.Equals(BuiltInCategory.OST_Columns))
                {


                    BoundingBoxXYZ bb = elem.get_BoundingBox(null);
                    XYZ location = new XYZ((bb.Max.X + bb.Min.X) / 2, (bb.Max.Y + bb.Min.Y) / 2, height);
                    Options options = new Options();
                    options.ComputeReferences = true;
                    ReferenceArray arrRefs = new ReferenceArray();
                    XYZ p1 = new XYZ(bb.Min.X, bb.Max.Y + 1, bb.Min.Z);
                    XYZ p2 = new XYZ(bb.Max.X, bb.Max.Y + 1, bb.Min.Z);
                    XYZ WasllXL = p1.Subtract(p2);
                    GeometryElement geometryElement = elem.get_Geometry(options);
                    foreach (GeometryObject @object in geometryElement)
                    {
                        Solid solid = @object as Solid;
                        if (solid != null)
                        {
                            FaceArray faces = solid.Faces;
                            foreach (PlanarFace planarFace in faces)
                            {
                                Reference reference = planarFace.Reference;
                                if (reference == null) continue;
                                Element element = doc.GetElement(reference.ElementId);
                                if (element == null) continue;
                                XYZ xYZ = planarFace.FaceNormal;
                                XYZ c = xYZ.CrossProduct(WasllXL);
                                if (c.IsZeroLength())
                                {
                                    arrRefs.Append(planarFace.Reference);
                                }
                            }
                        }
                    }

;
                    Line line = Line.CreateBound(p1, p2);


                    ReferenceArray arrRefs2 = new ReferenceArray();
                    XYZ P1 = new XYZ(bb.Min.X - 1, bb.Min.Y, bb.Min.Z);
                    XYZ P2 = new XYZ(bb.Min.X - 1, bb.Max.Y, bb.Min.Z);
                    XYZ WasllXL2 = P1.Subtract(P2);
                    GeometryElement geometryElement2 = elem.get_Geometry(options);
                    foreach (GeometryObject @object in geometryElement2)
                    {
                        Solid solid = @object as Solid;
                        if (solid != null)
                        {
                            FaceArray faces = solid.Faces;
                            foreach (PlanarFace planarFace in faces)
                            {
                                Reference reference = planarFace.Reference;
                                if (reference == null) continue;
                                Element element = doc.GetElement(reference.ElementId);
                                if (element == null) continue;
                                XYZ xYZ = planarFace.FaceNormal;
                                XYZ c = xYZ.CrossProduct(WasllXL2);
                                if (c.IsZeroLength())
                                {
                                    arrRefs2.Append(planarFace.Reference);
                                }
                            }
                        }
                    }
;

                    Line line2 = Line.CreateBound(P1, P2);





                    FilteredElementCollector CollectorText = new FilteredElementCollector(doc).OfClass(typeof(TextElementType));
                    ElementId elemid = null;
                    foreach (Element e in CollectorText)
                    {
                        if (e.Name.Contains("C"))
                        {
                            elemid = e.Id;
                        }
                    }




                    Transaction trans = new Transaction(doc);
                    trans.Start("Auto Mark Start");
                    Options geomOptions = new Options();
                    geomOptions.ComputeReferences = true;
                    {
                        TextNoteOptions textNoteOptions = new TextNoteOptions
                        {

                            KeepRotatedTextReadable = true,
                            HorizontalAlignment = HorizontalTextAlignment.Center,
                            VerticalAlignment = VerticalTextAlignment.Middle,
                            TypeId = elemid
                        };
                        String Text = elem.LookupParameter("備註").AsValueString().Trim();
                        TextNote textNote = TextNote.Create(doc, uidoc.ActiveGraphicalView.Id, location, Text, textNoteOptions);
                        if (arrRefs.Size < 2 || arrRefs2.Size < 2)
                        {
                            trans.Commit();
                            continue;
                        }

                        Dimension dim = doc.Create.NewDimension(doc.ActiveView, line, arrRefs);
                        Dimension dim2 = doc.Create.NewDimension(doc.ActiveView, line2, arrRefs2);
                        trans.Commit();
                    }
                }

            }


            return Result.Succeeded;
        }
    }
}

