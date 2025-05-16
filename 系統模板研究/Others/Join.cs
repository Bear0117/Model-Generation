using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.DB.Structure;
using System.Management.Instrumentation;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class Join : IExternalCommand
    {
        Autodesk.Revit.ApplicationServices.Application app;
        Document doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            UIApplication uiApp = commandData.Application;

            Selection sel = uiApp.ActiveUIDocument.Selection;
            Reference r1 = sel.PickObject(ObjectType.Element, "選擇第一個物件");
            Reference r2 = sel.PickObject(ObjectType.Element, "選擇第二個物件");

            using (Transaction trans = new Transaction(doc))
            {
                trans.Start("start");
                Element e1 = doc.GetElement(r1.ElementId);
                Element e2 = doc.GetElement(r2.ElementId);

                // 判斷選擇順序與優先順序是否相同
                // 若是，則直接執行接合
                // 若否，則先執行接合後再改變順序
                if (RunWithDefaultOrder(e1, e2))
                    RunJoinGeometry(doc, e1, e2);
                else
                    RunJoinGeometryAndSwitch(doc, e1, e2);

                trans.Commit();
            }


            return Result.Succeeded;
        }
        // 判斷優先順序與選擇順序是否一致
        bool RunWithDefaultOrder(Element e1, Element e2)
        {

            int order1 = GetOrderNumber(e1);
            int order2 = GetOrderNumber(e2);

            if (order1 >= order2)
                return true;
            else
                return false;
        }

        // 判斷選擇物件的種類 柱=2 梁=1 版=4 牆=3
        int GetOrderNumber(Element e)
        {
            const string columnString = "柱";
            const string slabString = "樓板";
            const string beamString = "結構構架";
            const string wallString = "牆";

            if (e.Category.Name == columnString)
            {
                TaskDialog.Show("Info", e.Category.Name);
                return 4;
            }
            else if (e.Category.Name == slabString)
            {
                TaskDialog.Show("Info", e.Category.Name);
                return 1;
            }
            else if (e.Category.Name == beamString)
            {
                TaskDialog.Show("Info", e.Category.Name);
                return 3;
            }
            else if (e.Category.Name == wallString)
            {
                TaskDialog.Show("Info", e.Category.Name);
                return 2;
            }
            else
            {
                TaskDialog.Show("Info", e.Category.Name);
                return 0;
            }
        }

        // 執行接合
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
    }
}