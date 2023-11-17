using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System;
using System.Linq;
using Teigha.Runtime;
using Teigha.DatabaseServices;
using System.IO;
using Teigha.Geometry;
using Autodesk.Revit.DB.Structure;
using System.Diagnostics;
using System.Windows;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateBeams_Test : IExternalCommand
    {
        Document doc;
        UIDocument uidoc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            uidoc = commandData.Application.ActiveUIDocument;
            doc = uidoc.Document;

            // Select a CAD layer and get its name.
            Reference reference = uidoc.Selection.PickObject(ObjectType.PointOnElement, "Select a CAD Layer");
            Element elem = doc.GetElement(reference);
            GeometryElement geoElem = elem.get_Geometry(new Options());
            GeometryObject geoObj = elem.GetGeometryObjectFromReference(reference);
            Category targetCategory = null;
            ElementId graphicsStyleId = null;

            string layer_name = "S-Beam"; //Default
            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                graphicsStyleId = geoObj.GraphicsStyleId;
                using (GraphicsStyle gs = doc.GetElement(geoObj.GraphicsStyleId) as GraphicsStyle)
                {
                    if (gs != null)
                    {
                        targetCategory = gs.GraphicsStyleCategory;
                        layer_name = gs.GraphicsStyleCategory.Name;
                    }
                }
            }

            if (geoElem == null || graphicsStyleId == null)
            {
                message = "GeometryElement or ID does not exist！";
                return Result.Failed;
            }

            // To get the "Level" of imported CAD 
            Level level = doc.GetElement(elem.LevelId) as Level;
            MessageBox.Show(level.Name.ToString());

            // Pair two lines for creating beams. 
            const double NormBeamWidth = 150; // The maximum width of beams.(Unit:cm)
            string path = GetCADPath(elem.GetTypeId(), doc);
            List<CADModel> curveArray_List = GetCurveArray(geoElem, graphicsStyleId);// A list of Original CADModels.
            List<CADModel> NotMatchCadModel = new List<CADModel>(); // A list to store unpaired CADModel.     
            List<List<CADModel>> CADModelList_List = new List<List<CADModel>>();// A list to store paired CADModels.

            List<CADModel> curveArray_List_copy = new List<CADModel>(); // A list of copied CADModels. (lines)
            foreach (CADModel OrginCADModle in curveArray_List)
            {
                curveArray_List_copy.Add(OrginCADModle);
            }

            // Start pairing.
            while (curveArray_List.Count > 0)
            {
                // A list to store distances.
                List<double> distanceList = new List<double>();

                // A list to store two paired CADModel.
                List<CADModel> cADModel_B_List = new List<CADModel>();

                CADModel CadModel_A = curveArray_List[0]; // CadModel_A is a CADModel to be paired.
                curveArray_List.Remove(CadModel_A); // Remove CadModel_A from original list

                // If there is only CADModel left, add it to the unpaired list directly.
                if (curveArray_List.Count > 0)
                {
                    foreach (CADModel CadModel_B in curveArray_List)
                    {
                        // To find all the CADModel which has the same length and rotation angle with CADModel_A.
                        // If the length of  two lines are different, the max error is 0.5 cm.
                        if ((float)Math.Abs(CadModel_A.Rotation) == (float)Math.Abs(CadModel_B.Rotation)
                            && Math.Abs(CadModel_A.Length - CadModel_B.Length) < CentimetersToUnits(0.5))
                        {
                            double distance = CadModel_A.Location.DistanceTo(CadModel_B.Location);
                            distanceList.Add(distance);
                            cADModel_B_List.Add(CadModel_B);
                        }
                    }

                    if (distanceList.Count != 0 && cADModel_B_List.Count != 0)
                    {
                        // Get the distance of closest CADModel.
                        double distanceTwoLine = distanceList.Min();

                        // The width of beams must between 10cm to 150cm, and the length larger than 100cm.
                        if (UnitsToCentimeters(distanceTwoLine) < NormBeamWidth 
                            && UnitsToCentimeters(distanceTwoLine) > 10
                            && CadModel_A.Length > CentimetersToUnits(100))
                        {
                            // Remove CADModel_B from original lsit
                            CADModel CadModel_shortDistance = cADModel_B_List[distanceList.IndexOf(distanceTwoLine)];
                            curveArray_List.Remove(CadModel_shortDistance);

                            // Store two paired CADModel into a list
                            List<CADModel> cADModels = new List<CADModel>{ CadModel_A, CadModel_shortDistance};
                            CADModelList_List.Add(cADModels);
                        }
                    }
                    else NotMatchCadModel.Add(CadModel_A);
                }
                else NotMatchCadModel.Add(CadModel_A);
            }

            // Pair the "NotMatch" CADModel with copied list.
            while (NotMatchCadModel.Count > 0)
            {
                List<double> distanceList_No = new List<double>();
                List<CADModel> cADModel_B_List_No = new List<CADModel>();

                CADModel CadModel_A_No = NotMatchCadModel[0];
                NotMatchCadModel.Remove(CadModel_A_No);

                if (curveArray_List_copy.Count > 0)
                {
                    foreach (CADModel CadModel_B_No in curveArray_List_copy)
                    {
                        if ((float)Math.Abs(CadModel_A_No.Rotation) == (float)Math.Abs(CadModel_B_No.Rotation)
                            && Math.Abs(CadModel_A_No.Length - CadModel_B_No.Length) < CentimetersToUnits(0.5)
                            && UnitsToCentimeters(CadModel_A_No.Location.DistanceTo(CadModel_B_No.Location)) > 10)
                        {
                            double distance = CadModel_A_No.Location.DistanceTo(CadModel_B_No.Location);
                            distanceList_No.Add(distance);
                            cADModel_B_List_No.Add(CadModel_B_No);
                        }
                    }

                    if (distanceList_No.Count != 0 && distanceList_No.Count != 0)
                    {
                        // Get the distance of closest CADModel.
                        double distanceTwoLine_No = distanceList_No.Min();

                        // The width of beams must between 10cm to 150cm, and the length larger than 100cm.
                        if (UnitsToCentimeters(distanceTwoLine_No) < NormBeamWidth
                            && UnitsToCentimeters(distanceTwoLine_No) > 10
                            && CadModel_A_No.Length > CentimetersToUnits(100))
                        {
                            // Remove CADModel_B from original lsit
                            CADModel CadModel_shortDistance_No = cADModel_B_List_No[distanceList_No.IndexOf(distanceTwoLine_No)];
                            curveArray_List.Remove(CadModel_shortDistance_No);

                            // Store two paired CADModel into a list
                            List<CADModel> cADModels = new List<CADModel> { CadModel_A_No, CadModel_shortDistance_No };
                            CADModelList_List.Add(cADModels);
                        }
                    }
                    else continue;
                }
                else continue;
            }

            // Get the beam depth by closest label and store it in CADModel.
            List < List < CADModel >> sss = GetPairedLabel(CADModelList_List, path, layer_name);

            // Create beams with corresponding sizes.
            foreach (List<CADModel> cadModelList in CADModelList_List)
            {
                // The depth is recorded in cADModel_A
                CADModel cADModel_A = cadModelList[0];
                CADModel cADModel_B = cadModelList[1];

                XYZ cADModel_A_StratPoint = RoundPoint(cADModel_A.CurveArray.get_Item(0).GetEndPoint(0));
                XYZ cADModel_A_EndPoint = RoundPoint(cADModel_A.CurveArray.get_Item(0).GetEndPoint(1));
                XYZ cADModel_B_StratPoint = RoundPoint(cADModel_B.CurveArray.get_Item(0).GetEndPoint(0));
                XYZ cADModel_B_EndPoint = RoundPoint(cADModel_B.CurveArray.get_Item(0).GetEndPoint(1));

                XYZ ChangeXYZ = new XYZ();

                double LineLength = (GetMiddlePoint(cADModel_A_StratPoint, cADModel_B_StratPoint)).DistanceTo(GetMiddlePoint(cADModel_A_EndPoint, cADModel_B_EndPoint));
                if (UnitsToCentimeters(LineLength) < 1)
                {
                    ChangeXYZ = cADModel_B_StratPoint;
                    cADModel_B_StratPoint = cADModel_B_EndPoint;
                    cADModel_B_EndPoint = ChangeXYZ;
                }

                XYZ beamLocation = RoundPoint((cADModel_A_StratPoint + cADModel_B_EndPoint) / 2);


                ///XYZ extension = (cADModel_A_StratPoint - cADModel_A_EndPoint) / (cADModel_A_StratPoint - cADModel_A_EndPoint).GetLength();
                Autodesk.Revit.DB.Curve curve = Autodesk.Revit.DB.Line.CreateBound(
                    RoundPoint(GetMiddlePoint(cADModel_A_StratPoint, cADModel_B_StratPoint)),
                    RoundPoint(GetMiddlePoint(cADModel_A_EndPoint, cADModel_B_EndPoint))
                    );

                Autodesk.Revit.DB.Line lineb = cADModel_A.CurveArray.get_Item(0) as Autodesk.Revit.DB.Line;
                double width = lineb.Distance(cADModel_B.Location);
                width = Math.Round(UnitsToCentimeters(width));
                String beamSize = width.ToString() + "x" + cadModelList[0].Depth;

                using (Autodesk.Revit.DB.Transaction transaction = new Autodesk.Revit.DB.Transaction(doc))
                {
                    transaction.Start("Beam Strart Build");
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    collector.OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_StructuralFraming);
                    FamilySymbol gotSymbol = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(q => q.Name == "default") as FamilySymbol;
                    
                    FamilyInstance instance = doc.Create.NewFamilyInstance(curve, gotSymbol, level, Autodesk.Revit.DB.Structure.StructuralType.Beam);
                    StructuralFramingUtils.DisallowJoinAtEnd(instance, 1);
                    StructuralFramingUtils.DisallowJoinAtEnd(instance, 0);
                    transaction.Commit();
                }
                ChangeBeamType(doc, beamSize);
            }

            FilteredElementCollector collector_b = new FilteredElementCollector(doc);
            List<Element> beams = collector_b
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .WhereElementIsNotElementType()
                .ToList();
            //foreach(Element beam in beams)
            //{
            //    BoundingBoxXYZ bbxyzElement = beam.get_BoundingBox(null);
            //    Outline outline = new Outline(bbxyzElement.Min, bbxyzElement.Max);

            //    //Create a filter to get all the intersection elements with wall.
            //    BoundingBoxIntersectsFilter filterW = new BoundingBoxIntersectsFilter(outline);

            //    //Create a filter to get StructuralFraming (which include beam and column) and Slabs.
            //    ElementCategoryFilter filterSt = new ElementCategoryFilter(BuiltInCategory.OST_StructuralFraming);
            //    ElementCategoryFilter filterSl = new ElementCategoryFilter(BuiltInCategory.OST_Floors);
            //    LogicalOrFilter filterS = new LogicalOrFilter(filterSt, filterSl);

            //    //Combine two filter.
            //    LogicalAndFilter filter = new LogicalAndFilter(filterS, filterW);

            //    //A list to store the intersected elements.
            //    IList<Element> inter = new FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements();

            //    for (int i = 0; i < inter.Count; i++)
            //    {
            //        if (inter[i] != null)
            //        {
            //            String elementName = inter[i].Category.Name;
            //            if (elementName == "結構構架")
            //            {
            //                Autodesk.Revit.DB.Transaction t1 = new Autodesk.Revit.DB.Transaction(doc, "Join");
            //                t1.Start();
            //                if(inter[i].Id != beam.Id)
            //                {
            //                    if (!JoinGeometryUtils.AreElementsJoined(doc, inter[i], beam))
            //                        JoinGeometryUtils.JoinGeometry(doc, inter[i], beam);
            //                    else
            //                    {
            //                        JoinGeometryUtils.UnjoinGeometry(doc, inter[i], beam);
            //                        JoinGeometryUtils.JoinGeometry(doc, inter[i], beam);
            //                    }
            //                }
            //                t1.Commit();
            //            }
            //            if (elementName == "柱")
            //            {
            //                Autodesk.Revit.DB.Transaction t1 = new Autodesk.Revit.DB.Transaction(doc, "Join");
            //                t1.Start();
            //                if (!JoinGeometryUtils.AreElementsJoined(doc, inter[i], beam))
            //                    JoinGeometryUtils.JoinGeometry(doc, inter[i], beam);
            //                else
            //                {
            //                    JoinGeometryUtils.UnjoinGeometry(doc, inter[i], beam);
            //                    JoinGeometryUtils.JoinGeometry(doc, inter[i], beam);
            //                }
            //                t1.Commit();
            //            }
            //            else continue;
            //        }
            //    }

            //}
            //for(int i = 0; i < beams.Count; i++)
            //{
            //    Element beam1 = beams[i];
            //    for(int j = 0; j < beams.Count; j++)
            //    {
            //        if(i >= j)
            //        {
            //            continue;
            //        }

            //        BoundingBoxXYZ bbxyzElement = beam1.get_BoundingBox(null);
            //        Outline outline = new Outline(bbxyzElement.Min, bbxyzElement.Max);

            //        //Create a filter to get all the intersection elements with wall.
            //        BoundingBoxIntersectsFilter filterW = new BoundingBoxIntersectsFilter(outline);

            //        //Create a filter to get column.
            //        ElementCategoryFilter filterSc = new ElementCategoryFilter(BuiltInCategory.OST_Columns);

            //        //Combine two filter.
            //        LogicalAndFilter filter = new LogicalAndFilter(filterSc, filterW);

            //        //A list to store the intersected elements.
            //        IList<Element> inter = new FilteredElementCollector(doc).WherePasses(filter).WhereElementIsNotElementType().ToElements();




            //        Element beam2 = beams[j];
            //        Autodesk.Revit.DB.Transaction t1 = new Autodesk.Revit.DB.Transaction(doc, "Join");
            //        t1.Start();
            //        if (!JoinGeometryUtils.AreElementsJoined(doc, beam2, beam1))
            //            JoinGeometryUtils.JoinGeometry(doc, beam2, beam1);
            //        else
            //        {
            //            JoinGeometryUtils.UnjoinGeometry(doc, beam2, beam1);
            //            JoinGeometryUtils.JoinGeometry(doc, beam2, beam1);
            //        }
            //        t1.Commit();
            //    }
            //}
            
            return Result.Succeeded;
        }

        /// <summary>
        /// Get all the lines in this layer. 
        /// </summary>
        /// <param name="doc">RevitDocument</param>
        /// <param name="geoElem">GeometryElement</param>
        /// <param name="graphicsStyleId">GeometryElementID</param>
        /// <returns></returns>
        private List<CADModel> GetCurveArray(GeometryElement geoElem, ElementId graphicsStyleId)
        {
            List<CADModel> curveArray_List = new List<CADModel>();
            //TransactionGroup transGroup = new TransactionGroup(doc, "繪製模型線");
            //transGroup.Start();

            // To verify the type of elements.
            foreach (GeometryObject gObj in geoElem)
            {
                GeometryInstance geomInstance = gObj as GeometryInstance;

                // Coordinates transformation
                Transform transform = geomInstance.Transform;

                if (null != geomInstance)
                {
                    foreach (GeometryObject insObj in geomInstance.SymbolGeometry)
                    {
                        if (insObj.GraphicsStyleId.IntegerValue != graphicsStyleId.IntegerValue)
                            continue;

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.Line")
                        {
                            Autodesk.Revit.DB.Line line = insObj as Autodesk.Revit.DB.Line;
                            Autodesk.Revit.DB.Line newLine = TransformLine(transform, line);

                            CurveArray curveArray = new CurveArray();
                            curveArray.Append(newLine);

                            XYZ startPoint = newLine.GetEndPoint(0);
                            XYZ endPoint = newLine.GetEndPoint(1);
                            XYZ MiddlePoint = GetMiddlePoint(startPoint, endPoint);
                            double angle = (startPoint.Y - endPoint.Y) / startPoint.DistanceTo(endPoint);
                            double rotation = Math.Asin(angle);

                            CADModel cADModel = new CADModel
                            {
                                CurveArray = curveArray,
                                Length = newLine.Length,
                                Shape = "矩形梁",
                                Location = MiddlePoint,
                                Rotation = rotation
                            };
                            curveArray_List.Add(cADModel);
                        }

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> points = polyLine.GetCoordinates();

                            for (int i = 0; i < points.Count - 1; i++)
                            {
                                Autodesk.Revit.DB.Line line = Autodesk.Revit.DB.Line.CreateBound(points[i], points[i + 1]);
                                line = TransformLine(transform, line);
                                Autodesk.Revit.DB.Line newLine = line;
                                CurveArray curveArray = new CurveArray();
                                curveArray.Append(newLine);

                                XYZ startPoint = newLine.GetEndPoint(0);
                                XYZ endPoint = newLine.GetEndPoint(1);
                                XYZ MiddlePoint = GetMiddlePoint(startPoint, endPoint);
                                double angle = (startPoint.Y - endPoint.Y) / startPoint.DistanceTo(endPoint);
                                double rotation = Math.Asin(angle);

                                CADModel cADModel = new CADModel
                                {
                                    CurveArray = curveArray,
                                    Length = newLine.Length,
                                    Shape = "矩形梁",
                                    Width = 300 / 304.8,
                                    Location = MiddlePoint,
                                    Rotation = rotation
                                };
                                curveArray_List.Add(cADModel);
                            }
                        }
                    }
                }
            }
            //transGroup.Assimilate();
            return curveArray_List;
        }

        /// <summary>
        /// Transform particular line
        /// </summary>
        /// <param name="transform">TransformMatrix</param>
        /// <param name="line">TransformedLine</param>
        /// <returns></returns>
        private Autodesk.Revit.DB.Line TransformLine(Transform transform, Autodesk.Revit.DB.Line line)
        {
            XYZ startPoint = transform.OfPoint(line.GetEndPoint(0));
            XYZ endPoint = transform.OfPoint(line.GetEndPoint(1));
            Autodesk.Revit.DB.Line newLine = Autodesk.Revit.DB.Line.CreateBound(startPoint, endPoint);
            return newLine;
        }

        public XYZ RoundPoint(XYZ point)
        {
            XYZ newPoint = point;
            //XYZ newPoint = new XYZ(
            //    CentimetersToUnits(Math.Round(UnitsToCentimeters(point.X) * 10) / 10),
            //    CentimetersToUnits(Math.Round(UnitsToCentimeters(point.Y) * 10) / 10),
            //    CentimetersToUnits(Math.Round(UnitsToCentimeters(point.Z) * 10) / 10)
            //    );
            return newPoint;
        }

        public void ChangeBeamType(Document doc, string beamSize)
        {
            char separator = 'x';
            string[] parts = beamSize.Split(separator);
            double width = int.Parse(parts[0]);
            double depth = int.Parse(parts[1]);

            FilteredElementCollector Collector = new FilteredElementCollector(doc);
            List<FamilySymbol> familySymbolList = Collector.OfClass(typeof(FamilySymbol))
            .OfCategory(BuiltInCategory.OST_StructuralFraming)
            .Cast<FamilySymbol>().ToList();

            Boolean IsColumnTypeExist = false;
            foreach (FamilySymbol fs in familySymbolList)
            {
                if (fs.Name != beamSize)
                {
                    continue;
                }
                else
                {
                    IsColumnTypeExist = true;
                    break;
                }
            }
            if (!IsColumnTypeExist)
            {
                using (Autodesk.Revit.DB.Transaction t_createNewColumnType = new Autodesk.Revit.DB.Transaction(doc, "Ｃreate New Beam Type"))
                {
                    try
                    {
                        t_createNewColumnType.Start();

                        FamilySymbol default_FamilySymbol = new FilteredElementCollector(doc)
                        .OfClass(typeof(FamilySymbol))
                        .OfCategory(BuiltInCategory.OST_StructuralFraming)
                        .Cast<FamilySymbol>()
                        .FirstOrDefault(q => q.Name == "default") as FamilySymbol;

                        // set the "h" and "b" to a new value:
                        FamilySymbol newFamSym = default_FamilySymbol.Duplicate(beamSize) as FamilySymbol;
                        Parameter parah = newFamSym.LookupParameter("h");
                        Parameter parab = newFamSym.LookupParameter("b");
                        parah.Set(CentimetersToUnits(depth));
                        parab.Set(CentimetersToUnits(width));

                        t_createNewColumnType.Commit();
                    }
                    catch (System.Exception ex)
                    {
                        TaskDialog td = new TaskDialog("error")
                        {
                            Title = "error",
                            AllowCancellation = true,
                            MainInstruction = "error",
                            MainContent = "Error" + ex.Message,
                            CommonButtons = TaskDialogCommonButtons.Close
                        };
                        td.Show();

                        Debug.Print(ex.Message);
                        t_createNewColumnType.RollBack();
                    }
                }
            }
            using (Autodesk.Revit.DB.Transaction t = new Autodesk.Revit.DB.Transaction(doc, "Change Beam Type"))
            {
                t.Start();

                // The familyinstance you want to change -> e.g. UB-Universal Beam 305x165x40UB
                List<FamilyInstance> beams = new FilteredElementCollector(doc, doc.ActiveView.Id)
               .OfClass(typeof(FamilyInstance))
               .OfCategory(BuiltInCategory.OST_StructuralFraming)
               .Cast<FamilyInstance>()
               .Where(q => q.Name == "default").ToList();

                // The target familyinstance(familysymbol) what it should be.
                foreach (FamilyInstance beam in beams)
                {
                    if (new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .Cast<FamilySymbol>()
                .FirstOrDefault(q => q.Name == beamSize) is FamilySymbol fs)
                    {
                        beam.Symbol = fs;
                    }
                    else
                        continue;
                }
                t.Commit();
            }
        }

        /// <summary>
        /// Get the midpoint.
        /// </summary>
        /// <param name="startPoint">StartPoint</param>
        /// <param name="endPoint">EndPoint</param>
        /// <returns></returns>
        private XYZ GetMiddlePoint(XYZ startPoint, XYZ endPoint)
        {
            XYZ MiddlePoint = new XYZ((startPoint.X + endPoint.X) / 2, (startPoint.Y + endPoint.Y) / 2, (startPoint.Z + endPoint.Z) / 2);
            return MiddlePoint;
        }

        public XYZ ConverCADPointToRevitPoint(Point3d point)
        {
            return new XYZ(point.X, point.Y, point.Z);
        }

        public string GetCADPath(ElementId cadLinkTypeID, Document revitDoc)
        {
            CADLinkType cadLinkType = revitDoc.GetElement(cadLinkTypeID) as CADLinkType;
            return ModelPathUtils.ConvertModelPathToUserVisiblePath(cadLinkType.GetExternalFileReference().GetAbsolutePath());
        }

        public double UnitsToCentimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
        }

        public double CentimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Centimeters);
        }

        public XYZ PointCentimeterToUnit(XYZ point)
        {
            XYZ newPoint = new XYZ(
                CentimetersToUnits(point.X),
                CentimetersToUnits(point.Y),
                CentimetersToUnits(point.Z)
                );
            return newPoint;
        }

        public List<List<CADModel>> GetPairedLabel(List<List<CADModel>> CADModelList_List, String path, String layer_name)
        {
            foreach (List<CADModel> cadModelList in CADModelList_List)
            {
                string beamDepth = "0";
                double distanceBetweenTB = 10000000;
                XYZ location_beam = cadModelList[0].Location;

                using (new Services())
                {
                    using (Database database = new Database(false, false))
                    {
                        database.ReadDwgFile(path, FileShare.Read, true, "");
                        using (Teigha.DatabaseServices.Transaction trans1 = database.TransactionManager.StartTransaction())
                        {
                            using (BlockTable table = (BlockTable)database.BlockTableId.GetObject(OpenMode.ForRead))
                            {
                                using (SymbolTableEnumerator enumerator = table.GetEnumerator())
                                {
                                    enumerator.MoveNext();
                                    {
                                        using (BlockTableRecord record = (BlockTableRecord)enumerator.Current.GetObject(OpenMode.ForRead))
                                        {
                                            foreach (ObjectId id in record)
                                            {
                                                Entity entity2 = (Entity)id.GetObject(OpenMode.ForRead, false, false);
                                                switch (entity2.GetRXClass().Name)
                                                {
                                                    case "AcDbText":
                                                        Teigha.DatabaseServices.DBText text = (Teigha.DatabaseServices.DBText)entity2;
                                                        if (IsSameRotation(cadModelList[0], text.Rotation)
                                                            && text.TextString.Contains("H")
                                                            && text.Layer == layer_name)
                                                        {
                                                            double distance_cad_text = (PointCentimeterToUnit(ConverCADPointToRevitPoint(text.Position))).DistanceTo(location_beam);
                                                            if (distance_cad_text < distanceBetweenTB)
                                                            {
                                                                {
                                                                    distanceBetweenTB = distance_cad_text;
                                                                    int c = text.TextString.LastIndexOf('H');
                                                                    beamDepth = text.TextString.Substring(c + 2, 2);
                                                                }
                                                            }
                                                        }
                                                        break;

                                                    case "AcDbMText":
                                                        Teigha.DatabaseServices.MText text_m = (Teigha.DatabaseServices.MText)entity2;
                                                        if (IsSameRotation(cadModelList[0], text_m.Rotation)
                                                            && text_m.Text.Contains("H")
                                                            && text_m.Layer == layer_name)
                                                        {
                                                            double distance_cad_text = (PointCentimeterToUnit(ConverCADPointToRevitPoint(text_m.Location))).DistanceTo(location_beam);
                                                            if (distance_cad_text < distanceBetweenTB)
                                                            {
                                                                {
                                                                    distanceBetweenTB = distance_cad_text;
                                                                    int c = text_m.Text.LastIndexOf('H');
                                                                    beamDepth = text_m.Text.Substring(c + 2, 2);
                                                                }
                                                            }
                                                        }
                                                        break;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                cadModelList[0].Depth = beamDepth;
            }
            return CADModelList_List;
        }

        public Boolean IsSameRotation(CADModel cadmodel, double textRotation)
        {
            if(Math.Abs(Math.Abs(cadmodel.Rotation) - Math.Abs(textRotation)) < 0.001)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }

    public class CADModel
    {
        public CADModel()
        {
            CurveArray = null;
            Shape = "";
            Length = 0;
            Width = 0;
            Depth = "";
            FamilySymbol = "";
            Location = new XYZ(0, 0, 0);
            Rotation = 0;
            Elevation = 0.0;
        }

        /// <summary>
        /// Elevation
        /// </summary>
        public double Elevation { get; set; }

        /// <summary>
        /// CurveArray
        /// </summary>
        public CurveArray CurveArray { get; set; }

        /// <summary>
        /// shape
        /// </summary>
        public string Shape { get; set; }

        /// <summary>
        /// length
        /// </summary>
        public double Length { get; set; }

        /// <summary>
        /// width
        /// </summary>
        public double Width { get; set; }

        /// <summary>
        /// depth
        /// </summary>
        public string Depth { get; set; }

        /// <summary>
        /// familySymbol
        /// </summary>
        public string FamilySymbol { get; set; }

        /// <summary>
        /// location
        /// </summary>
        public XYZ Location { get; set; }

        /// <summary>
        /// rotation
        /// </summary>
        public double Rotation { get; set; }

        public static explicit operator CADModel(CurveArray v)
        {
            throw new NotImplementedException();
        }

    }

    public class SelectCADLinkFilter : ISelectionFilter
    {
        public bool AllowElement(Element elem)
        {
            if (elem is ImportInstance)
            {
                return true;
            }
            return false;
        }

        public bool AllowReference(Reference reference, XYZ position)
        {
            throw new NotImplementedException();
        }
    }
}
