using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Modeling.AreaEstimation
{
    internal class Test1
    {
        //    // 幾何面的邊緣
        //    EdgeArrayArray egdearrayarray = topFace.EdgeLoops;
        //    List<XYZ> points_Top = new List<XYZ>(); //平台頂面所有的點

        //    foreach (EdgeArray edgearray in egdearrayarray)
        //    {
        //        // 將邊緣線的點存到points裡
        //        foreach (Edge edge in edgearray)
        //        {
        //            Curve curve = edge.AsCurve();
        //            points_Top.Add(Modeling.Algorithm.RoundPoint(curve.GetEndPoint(0), 5));
        //            points_Top.Add(Modeling.Algorithm.RoundPoint(curve.GetEndPoint(1), 5));
        //        }
        //    }

        //    // 將points裡相同的點移除
        //    points_Top = RemoveDuplicatePoints(points_Top, 0.01);

        //    List<XYZ> newPoints = new List<XYZ>();
        //    foreach (XYZ point in points_Top)
        //    {
        //        newPoints.Add(point);
        //    }

        //    //foreach (XYZ p in points_Top)
        //    //{
        //    //    MessageBox.Show("平台頂面所有點:" + p.X.ToString() + "," + p.Y.ToString());
        //    //}


        //    List<double> points_maxminX= new List<double>();
        //    List<double> points_maxminY = new List<double>();

        //    // 使用 LINQ 获取最大和最小的 X 和 Y
        //    double minX = points_Top.Min(p => p.X);
        //    double maxX = points_Top.Max(p => p.X);
        //    double minY = points_Top.Min(p => p.Y);
        //    double maxY = points_Top.Max(p => p.Y);

        //    points_maxminX.Add(minX);
        //    points_maxminX.Add(maxX);
        //    points_maxminY.Add(minY);
        //    points_maxminY.Add(maxY);

        //    foreach (XYZ p in points_Top)
        //    {
        //        MessageBox.Show("平台頂面最大最小的XY點:" + p.X.ToString() + "," + p.Y.ToString());
        //    }

        //    break;
        //    int points_maxmin = points_maxminX.Count + points_maxminY.Count;

        //    MessageBox.Show(points_maxmin.ToString());

        //    for (int i = 0; i < points_Top.Count; i++) 
        //    {
        //        for (int j = 0; j < points_maxmin; j++)
        //        {
        //            if (Math.Abs(points_Top[i].X- points_maxminX[j]) < 1e-9|| Math.Abs(points_Top[i].Y - points_maxminY[j]) < 1e-9)
        //            {
        //                newPoints.Remove(points_Top[i]);
        //            }
        //        }
        //    }

        //    bool rectangle = true;

        //    if(newPoints.Count == 4)
        //    {
        //        rectangle = true;

        //        double length = Math.Abs(points_maxminX[0] - points_maxminX[1]);
        //        double width = Math.Abs(points_maxminY[0] - points_maxminY[1]);

        //        foreach (double point in points_maxminX)
        //        {
        //            MessageBox.Show("矩形最大最小X點"+point.ToString());
        //        }
        //        foreach (double point in points_maxminY)
        //        {
        //            MessageBox.Show("矩形最大最小Y點" + point.ToString());
        //        }

        //    }
        //    else
        //    {
        //        rectangle= false;
        //    }

        //    if (rectangle == false)
        //    {
        //        List<double> smallpoints_maxminX = new List<double>();
        //        List<double> samllpoints_maxminY = new List<double>();

        //        // 使用 LINQ 获取最大和最小的 X 和 Y
        //        double samllmaxX = newPoints.Max(p => p.X);
        //        double samllminX = newPoints.Min(p => p.X);
        //        double samllmaxY = newPoints.Max(p => p.Y);
        //        double samllminY = newPoints.Min(p => p.Y);

        //        smallpoints_maxminX.Add(maxX);
        //        smallpoints_maxminX.Add(minX);
        //        samllpoints_maxminY.Add(maxY);
        //        samllpoints_maxminY.Add(minY);

        //        double length_big = Math.Abs(points_maxminX[0] - points_maxminX[1]);
        //        double width_big = Math.Abs(points_maxminY[0] - points_maxminY[1]);
        //        double length_small = Math.Abs(smallpoints_maxminX[0] - smallpoints_maxminX[1]);
        //        double width_samll = Math.Abs(samllpoints_maxminY[0] - samllpoints_maxminY[1]);

        //        foreach (double point in points_maxminX)
        //        {
        //            MessageBox.Show("矩形最大最小X點" + point.ToString());
        //        }
        //        foreach (double point in points_maxminY)
        //        {
        //            MessageBox.Show("矩形最大最小Y點" + point.ToString());
        //        }

        //        foreach (XYZ p in newPoints)
        //        {
        //            MessageBox.Show("非矩形的小矩形的點" + p.X.ToString() + "," + p.Y.ToString());
        //        }
        //    }
    }
}
