using System; // 要使用Exception需引用此參考
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class AutoJoin : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            ICollection<ElementId> eleids = uidoc.Selection.GetElementIds();
            List<ElementId> eleids_List = eleids.ToList();

            for (int i = 0; i < eleids_List.Count; i++)
            {
                Element element1 = doc.GetElement(eleids_List[i]);

                BoundingBoxXYZ bbxyz = element1.get_BoundingBox(null);
                Outline outline = new Outline(bbxyz.Min, bbxyz.Max);
                BoundingBoxIntersectsFilter filter = new BoundingBoxIntersectsFilter(outline);
                IList<Element> inter = new FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements();
                
                for (int j = 0; j < eleids_List.Count; j++)
                {
                    if (i <= j || inter.Count == 0)
                    {
                        continue;
                    }

                    Element element2 = doc.GetElement(eleids_List[j]);
                    String ele2_id = element2.Id.ToString();

                    foreach (Element ele_foreach in inter)
                    {
                        //if (ele_foreach.Id.ToString() == ele2_id)
                        {
                            if (AdjustOrder(ele_foreach, element2))
                            {
                                StartJoin(doc, element1, element2);
                                break; 
                            }
                            
                        }
                    }                    
                }
            }      
            return Result.Succeeded;
        }

        public void StartJoin(Document doc, Element elem1, Element elem2)
        {
            Transaction trans = new Transaction(doc);
            trans.Start("Join");

            // 判斷選擇順序與優先順序是否相同
            // 若是，則直接執行接合
            // 若否，則先執行接合後再改變順序
            if (RunWithDefaultOrder(elem1, elem2))
                RunJoinGeometry(doc, elem1, elem2);
            else
                RunJoinGeometryAndSwitch(doc, elem1, elem2);

            trans.Commit();

        }

        // 執行接合
        public void RunJoinGeometry(Document doc, Element e1, Element e2)
        {
            if (!JoinGeometryUtils.AreElementsJoined(doc, e1, e2))
            {
                //MessageBox.Show("直接接合");
                JoinGeometryUtils.JoinGeometry(doc, e1, e2);               
            }            
            else
            {
                //MessageBox.Show("先取消接合再執行接合");
                JoinGeometryUtils.UnjoinGeometry(doc, e1, e2);
                JoinGeometryUtils.JoinGeometry(doc, e1, e2);
                
            }
        }

        // 執行接合並改變接合順序
        public void RunJoinGeometryAndSwitch(Document doc, Element e1, Element e2)
        {
            if (!JoinGeometryUtils.AreElementsJoined(doc, e1, e2))
            {
                //MessageBox.Show("直接接合再交換順序");
                JoinGeometryUtils.JoinGeometry(doc, e1, e2);
                JoinGeometryUtils.SwitchJoinOrder(doc, e1, e2);
                
            }
            else
            {
                //MessageBox.Show("先取消接合再執行接合再交換順序");
                JoinGeometryUtils.UnjoinGeometry(doc, e1, e2);
                //MessageBox.Show(JoinGeometryUtils.AreElementsJoined(doc, e1, e2).ToString());
                JoinGeometryUtils.UnjoinGeometry(doc, e1, e2);
                JoinGeometryUtils.UnjoinGeometry(doc, e1, e2);
                JoinGeometryUtils.JoinGeometry(doc, e1, e2);
                JoinGeometryUtils.SwitchJoinOrder(doc, e1, e2);
            }
        }
        public bool RunWithDefaultOrder(Element e1, Element e2)
        {
            int sum = GetOrderOfElement(e1) + GetOrderOfElement(e2);

            if (sum == 3 || sum == 7)
            {
                //MessageBox.Show("Correct order");
                return true;
            }

            else
            {
                //MessageBox.Show("Wrong order");
                return false;

            }
                
        }

        public int GetOrderOfElement(Element e)
        {
            if (e.Category.Name == "柱" || e.Category.Name == "結構柱")
            {
                //TaskDialog.Show("Info", e.Category.Name);
                return 1;
            }
            else if (e.Category.Name == "結構構架")
            {
                //TaskDialog.Show("Info", e.Category.Name);
                return 2;
            }
            else if (e.Category.Name == "樓板")
            {
                //TaskDialog.Show("Info", e.Category.Name);
                return 3;
            }
            else if (e.Category.Name == "牆")
            {
                //TaskDialog.Show("Info", e.Category.Name);
                return 4;
            }
            else
            {
                TaskDialog.Show("Info", "我是" + e.Category.Name.ToString());
                return 0;
            }
        }
        public Boolean AdjustOrder(Element e1, Element e2)
        {
            //該兩種元件之間是否執行接合
            if(GetOrderOfElement(e1) == 4 || GetOrderOfElement(e2) == 4)
            {
                return true;
            }
            else if (GetOrderOfElement(e1) == 2 && GetOrderOfElement(e2) == 3)
            {
                return true;
            }
            else if (GetOrderOfElement(e1) == 3 && GetOrderOfElement(e2) == 2)
            {
                return true;
            }          
            else
                return false;
        }
    }
}
