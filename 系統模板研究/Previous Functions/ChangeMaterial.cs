using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class ChangeMaterial : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Get the active document
            Document doc = commandData.Application.ActiveUIDocument.Document;

            // Create a filter for walls
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            ICollection<ElementId> ids = uidoc.Selection.GetElementIds();

            // Retrieve the concrete material
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            collector.OfClass(typeof(Material));
            IEnumerable<Material> materials = collector.Cast<Material>();
            Material concreteMaterial = materials.FirstOrDefault(m => m.Name == "Concrete, Precast"); // replace "Concrete" with the name of your material

            MessageBox.Show(ids.Count.ToString());

            foreach (ElementId id in ids)
            {
                // Set the wall's material to the concrete material
                using (Transaction tx = new Transaction(doc, "Paint Wall"))
                {
                    tx.Start();
                    Wall wall = doc.GetElement(id) as Wall;
                    wall.WallType.GetCompoundStructure().GetLayers()[0].MaterialId = concreteMaterial.Id;
                    tx.Commit();
                }
            }
            return Result.Succeeded;
        }
    }
}
