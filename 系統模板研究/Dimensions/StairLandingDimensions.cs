#region Namespaces
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Windows;

#endregion

namespace Modeling
{
    [Transaction(TransactionMode.Manual)]
    public class StairLandingDimensions : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            // Select stair to dimension
            Reference pickRef = uiapp.ActiveUIDocument.Selection.PickObject(ObjectType.Element, "Select stair to dimension");
            Element selectedElem = doc.GetElement(pickRef);

            if (selectedElem is Stairs)
            {
                Stairs selectedStair = selectedElem as Stairs;
                
                ICollection<ElementId> stairLandingsICollectionId = selectedStair.GetStairsLandings();
                //MessageBox.Show("平台:" + stairLandingsICollectionId.Count.ToString());
                ICollection<ElementId> stairRunsICollectionId = selectedStair.GetStairsRuns();
                //MessageBox.Show("梯段:" + stairRunsICollectionId.Count.ToString());

                //選到的樓梯裡的平台
                IList<Element> stairLandingsList = new List<Element>();
                foreach (ElementId stairLandingsId in stairLandingsICollectionId)
                {
                    Element elem = doc.GetElement(stairLandingsId);
                    if (elem != null)
                    {
                        stairLandingsList.Add(elem);
                    }
                }
                MessageBox.Show(stairLandingsList.Count.ToString());

                //平台的參數
                foreach (StairsLanding stairLanding in stairLandingsList)
                {
                    ReferenceArray referenceArray = new ReferenceArray();
                    Reference r1 = null, r2 = null;

                    Face stairLandingFace = GetFace(stairLanding/*, selectedStair.Direction*/);
                    EdgeArrayArray edgeArrays = stairLandingFace.EdgeLoops;
                    EdgeArray edges = edgeArrays.get_Item(0);

                    List<Edge> edgeList = new List<Edge>();
                    foreach (Edge edge in edges)
                    {
                        Line line = edge.AsCurve() as Line;

                        if (IsLineVertical(line) == true)
                        {
                            edgeList.Add(edge);
                        }
                    }
                    MessageBox.Show(edgeList.Count.ToString());

                    List<Edge> sortedEdges = edgeList.OrderByDescending(e => e.AsCurve().Length).ToList();
                    r1 = sortedEdges[0].Reference;

                    referenceArray.Append(r1);

                    // create dimension line
                    LocationCurve stairLandingLoc = stairLanding.Location as LocationCurve;
                    MessageBox.Show(stairLandingLoc.ToString());
                    Line stairLandingLine = stairLandingLoc.Curve as Line;

                    XYZ offset1 = GetOffsetByStairLandingOrientation(stairLandingLine.GetEndPoint(0), /*selectedStair.Direction,*/ 5);
                    XYZ offset2 = GetOffsetByStairLandingOrientation(stairLandingLine.GetEndPoint(1), /*selectedStair.Direction,*/ 5);
                    MessageBox.Show(offset1.ToString());
                    MessageBox.Show(offset2.ToString());

                    Line dimLine = Line.CreateBound(offset1, offset2);

                   

                    using (Transaction t = new Transaction(doc))
                    {
                        t.Start("Create new dimension");

                        Dimension newDim = doc.Create.NewDimension(doc.ActiveView, dimLine, referenceArray);

                        MessageBox.Show("建立標註完成");
                        t.Commit();
                    }
                }
            }
            else
            {
                return Result.Failed;
            }

            return Result.Succeeded;
        }

        private XYZ GetOffsetByStairLandingOrientation(XYZ point/*, XYZ orientation*/, int value)
        {
            XYZ newVector = point.Multiply(value);
            XYZ returnPoint = point.Add(newVector);

            return returnPoint;
        }

        public enum SpecialReferenceType
        {
            Left = 0,
            CenterLR = 1,
            Right = 2,
            Front = 3,
            CenterFB = 4,
            Back = 5,
            Bottom = 6,
            CenterElevation = 7,
            Top = 8
        }

        //private Reference GetSpecialFamilyReference(FamilyInstance inst, SpecialReferenceType refType)
        //{
        //    // source for this method: https://thebuildingcoder.typepad.com/blog/2016/04/stable-reference-string-magic-voodoo.html

        //    Reference indexRef = null;

        //    int idx = (int)refType;

        //    if (inst != null)
        //    {
        //        Document dbDoc = inst.Document;

        //        Options geomOptions = new Options();
        //        geomOptions.ComputeReferences = true;
        //        geomOptions.DetailLevel = ViewDetailLevel.Undefined;
        //        geomOptions.IncludeNonVisibleObjects = true;

        //        GeometryElement gElement = inst.get_Geometry(geomOptions);
        //        GeometryInstance gInst = gElement.First() as GeometryInstance;

        //        String sampleStableRef = null;

        //        if (gInst != null)
        //        {
        //            GeometryElement gSymbol = gInst.GetSymbolGeometry();

        //            if (gSymbol != null)
        //            {
        //                foreach (GeometryObject geomObj in gSymbol)
        //                {
        //                    if (geomObj is Solid)
        //                    {
        //                        Solid solid = geomObj as Solid;

        //                        if (solid.Faces.Size > 0)
        //                        {
        //                            Face face = solid.Faces.get_Item(0);
        //                            sampleStableRef = face.Reference.ConvertToStableRepresentation(dbDoc);
        //                            break;
        //                        }
        //                    }
        //                    else if (geomObj is Curve)
        //                    {
        //                        Curve curve = geomObj as Curve;
        //                        Reference curveRef = curve.Reference;
        //                        if (curveRef != null)
        //                        {
        //                            sampleStableRef = curve.Reference.ConvertToStableRepresentation(dbDoc);
        //                            break;
        //                        }

        //                    }
        //                    else if (geomObj is Point)
        //                    {
        //                        Point point = geomObj as Point;
        //                        sampleStableRef = point.Reference.ConvertToStableRepresentation(dbDoc);
        //                        break;
        //                    }
        //                }
        //            }

        //            if (sampleStableRef != null)
        //            {
        //                String[] refTokens = sampleStableRef.Split(new char[] { ':' });

        //                String customStableRef = refTokens[0] + ":"
        //                  + refTokens[1] + ":" + refTokens[2] + ":"
        //                  + refTokens[3] + ":" + idx.ToString();

        //                indexRef = Reference.ParseFromStableRepresentation(dbDoc, customStableRef);

        //                GeometryObject geoObj = inst.GetGeometryObjectFromReference(indexRef);

        //                if (geoObj != null)
        //                {
        //                    String finalToken = "";
        //                    if (geoObj is Edge)
        //                    {
        //                        finalToken = ":LINEAR";
        //                    }

        //                    if (geoObj is Face)
        //                    {
        //                        finalToken = ":SURFACE";
        //                    }

        //                    customStableRef += finalToken;
        //                    indexRef = Reference.ParseFromStableRepresentation(dbDoc, customStableRef);
        //                }
        //                else
        //                {
        //                    indexRef = null;
        //                }
        //            }
        //        }
        //        else
        //        {
        //            throw new Exception("No Symbol Geometry found...");
        //        }
        //    }
        //    return indexRef;
        //}

        private bool IsLineVertical(Line line)
        {
            if (line.Direction.IsAlmostEqualTo(XYZ.BasisZ) || line.Direction.IsAlmostEqualTo(-XYZ.BasisZ))
                return true;
            else
                return false;
        }

        private Face GetFace(Element element/*, XYZ orientation*/)
        {
            GeometryElement geometryElement = element.get_Geometry(new Options());
            PlanarFace returnFace = null;

            // 編歷Geometry對象以獲取尺寸信息
            foreach (GeometryObject geomObj in geometryElement)
            {
                GeometryInstance geomInstance = geomObj as GeometryInstance;

                // 抓取在實體上的幾何參數
                if (geomObj is Solid solid)
                {
                    // 抓取幾何面
                    FaceArray faces = solid.Faces;
                    // MessageBox.Show("平台面數量:" + faces.Size.ToString());

                    foreach (Face face in faces)
                    {
                        XYZ targetFaceNormal = face.ComputeNormal(UV.Zero);
                        XYZ zDirection = XYZ.BasisZ;
                        double dotProduct = zDirection.DotProduct(targetFaceNormal);
                        //Zdirection.DotProduct(targetFaceNormal))  == 1 或是 == -1

                        if (face is PlanarFace)
                        {
                            PlanarFace pf = face as PlanarFace;
                            XYZ normal = pf.FaceNormal;
                            TaskDialog.Show("法线向量", $"X: {normal.X}, Y: {normal.Y}, Z: {normal.Z}");
                            if(normal.Z==1|| normal.Z == -1)
                            {
                                returnFace = pf;
                            }
                            //if (pf.FaceNormal.IsAlmostEqualTo(orientation))
                            //returnFace = pf;

                            //if (Math.Abs(dotProduct - 1.0) < 1e-9 || Math.Abs(dotProduct + 1.0) < 1e-9)
                            //{
                            //    returnFace = pf;
                            //}
                        }
                    }
                }
            }

            return returnFace;
        }

        //internal static PushButtonData GetButtonData()
        //{
        //    // use this method to define the properties for this command in the Revit ribbon
        //    string buttonInternalName = "btnCommand1";
        //    string buttonTitle = "Button 1";

        //    ButtonDataClass myButtonData1 = new ButtonDataClass(
        //        buttonInternalName,
        //        buttonTitle,
        //        MethodBase.GetCurrentMethod().DeclaringType?.FullName,
        //        Properties.Resources.Blue_32,
        //        Properties.Resources.Blue_16,
        //        "This is a tooltip for Button 1");

        //    return myButtonData1.Data;
        //}
    }
}
