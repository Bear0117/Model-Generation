using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;

namespace Modeling
{
    public class ElemsSelect
    {
        public IList<Element> SelectElements(ExternalCommandData commandData, ISelectionFilter selectionFilter)
        {
            IList<Reference> references = commandData.Application.ActiveUIDocument.Selection.PickObjects(ObjectType.Element, selectionFilter);
            IList<Element> elems = new List<Element>();
            foreach (var item in references)
            {
                elems.Add(commandData.Application.ActiveUIDocument.Document.GetElement(item));
            }
            return elems;
        }
        public IList<Element> SelectElements(ExternalCommandData commandData)
        {
            IList<Reference> references = commandData.Application.ActiveUIDocument.Selection.PickObjects(ObjectType.Element);
            IList<Element> elems = new List<Element>();
            foreach (var item in references)
            {
                elems.Add(commandData.Application.ActiveUIDocument.Document.GetElement(item));
            }
            return elems;
        }
        public IList<Element> SelectElements(ExternalCommandData commandData, ISelectionFilter selectionFilter, string statusPrompt)
        {
            IList<Reference> references = commandData.Application.ActiveUIDocument.Selection.PickObjects(ObjectType.Element, selectionFilter, statusPrompt);
            IList<Element> elems = new List<Element>();
            foreach (var item in references)
            {
                elems.Add(commandData.Application.ActiveUIDocument.Document.GetElement(item));
            }
            return elems;
        }
        public IList<Element> SelectElements(ExternalCommandData commandData, string statusPrompt)
        {
            IList<Reference> references = commandData.Application.ActiveUIDocument.Selection.PickObjects(ObjectType.Element, statusPrompt);
            IList<Element> elems = new List<Element>();
            foreach (var item in references)
            {
                elems.Add(commandData.Application.ActiveUIDocument.Document.GetElement(item));
            }
            return elems;
        }
        public Element SelectElement(ExternalCommandData commandData, ISelectionFilter selectionFilter)
        {
            Reference reference = commandData.Application.ActiveUIDocument.Selection.PickObject(ObjectType.Element, selectionFilter);
            return commandData.Application.ActiveUIDocument.Document.GetElement(reference);
        }
        public Element SelectElement(ExternalCommandData commandData)
        {
            Reference reference = commandData.Application.ActiveUIDocument.Selection.PickObject(ObjectType.Element);
            return commandData.Application.ActiveUIDocument.Document.GetElement(reference);
        }
        public Element SelectElement(ExternalCommandData commandData, ISelectionFilter selectionFilter, string statusPrompt)
        {
            Reference reference = commandData.Application.ActiveUIDocument.Selection.PickObject(ObjectType.Element, selectionFilter, statusPrompt);
            return commandData.Application.ActiveUIDocument.Document.GetElement(reference);
        }
        public Element SelectElement(ExternalCommandData commandData, string statusPrompt)
        {
            Reference reference = commandData.Application.ActiveUIDocument.Selection.PickObject(ObjectType.Element, statusPrompt);
            return commandData.Application.ActiveUIDocument.Document.GetElement(reference);
        }
    }
}