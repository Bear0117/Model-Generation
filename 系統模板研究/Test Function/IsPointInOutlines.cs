using System.Collections.Generic;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Windows;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class IsPointInOutlines : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            XYZ point = new XYZ(5, 3, 0);
            List<Line> lines = new List<Line>
            {
                Line.CreateBound(new XYZ(0, 0, 0), new XYZ(5, 0, 0)),
                Line.CreateBound(new XYZ(5, 0, 0), new XYZ(5, 5, 0)),
                Line.CreateBound(new XYZ(5, 5, 0), new XYZ(0, 5, 0)),
                Line.CreateBound(new XYZ(0, 5, 0), new XYZ(0, 0, 0))
            };

            MessageBox.Show(IsInsideOutline_new(point,lines).ToString());

            return Result.Succeeded;
        }
        public bool IsInsideOutline_new(XYZ TargetPoint, List<Line> lines)
        {
            bool result = true;
            int insertCount = 0;
            //常有例外
            //+ new XYZ(500000, 0, 0)
            //Line rayLine = Line.CreateBound(TargetPoint, TargetPoint.Add(XYZ.BasisY * 100000000 ));
            Line rayLine = Line.CreateBound(TargetPoint, TargetPoint.Add(new XYZ(1, 0, 0) * 100000000));
            foreach (Line areaLine in lines)
            {
                XYZ mimdPoint = (areaLine.GetEndPoint(0) + areaLine.GetEndPoint(1)) / 2;
                //MessageBox.Show(mimdPoint.ToString());
                SetComparisonResult interResult = areaLine.Intersect(rayLine, out IntersectionResultArray resultArray);
                //MessageBox.Show(interResult.ToString());
                IntersectionResult insPoint = resultArray?.get_Item(0);
                if (insPoint != null)
                {
                    //MessageBox.Show(insPoint.ToString());
                    //MessageBox.Show("邊線與射線有交集");
                    insertCount++;
                }
                else
                {
                    //MessageBox.Show("邊線與射線無交集");
                } 
            }
            MessageBox.Show(insertCount.ToString());
            //如果次數為偶數就在外面，次數為奇數就在裡面
            //如果點在Outline上的話，也算是在外面
            if (insertCount % 2 == 0)//偶数
            {
                return result = false;
            }
            return result;
        }
    }
}