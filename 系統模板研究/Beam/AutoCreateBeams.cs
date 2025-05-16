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
using System.Text.RegularExpressions;
using Line = Autodesk.Revit.DB.Line;
using Curve = Autodesk.Revit.DB.Curve;
using Transaction = Autodesk.Revit.DB.Transaction;
using Exception = System.Exception;
using Transform = Autodesk.Revit.DB.Transform;
using System.Windows;
using System.Text;

namespace Modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class AutoCreateBeams : IExternalCommand
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

            // Initialize Parameters.
            ModelingParam.Initialize();
            double gridline_size = ModelingParam.parameters.General.GridSize * 10; // unit: mm
            double[] beamWidthRange = ModelingParam.parameters.BeamParam.BeamWidthRange;
            double beamWidthMin = beamWidthRange[0];
            double beamWidthMax = beamWidthRange[1];
            //double beamMaxLength = ModelingParam.parameters.BeamParam.BeamMinLength;

            // To get the "Level" of imported CAD 
            Level level = doc.GetElement(elem.LevelId) as Level;
            //MessageBox.Show(level.Name.ToString());

            // Pair two lines for creating beams. 
            //const double NormBeamWidth = 150; // The maximum width of beams.(Unit:cm)

            string path = GetCADPath(elem.GetTypeId(), doc);
            List<CADModel> curveArray_List = GetCurveArray(geoElem, graphicsStyleId, gridline_size);// A list of Original CADModels.
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

                        // The width of beams must between 10cm to 150cm, and the length larger than 80cm.
                        if (UnitsToCentimeters(distanceTwoLine) < beamWidthMax && UnitsToCentimeters(distanceTwoLine) > CentimetersToUnits(beamWidthMin))
                        {
                            // Remove CADModel_B from original lsit
                            CADModel CadModel_shortDistance = cADModel_B_List[distanceList.IndexOf(distanceTwoLine)];
                            curveArray_List.Remove(CadModel_shortDistance);

                            // Store two paired CADModel into a list
                            List<CADModel> cADModels = new List<CADModel> { CadModel_A, CadModel_shortDistance };
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
                            && Math.Abs(CadModel_A_No.Length - CadModel_B_No.Length) < CentimetersToUnits(0.5))
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
                        if (UnitsToCentimeters(distanceTwoLine_No) < beamWidthMax && UnitsToCentimeters(distanceTwoLine_No) > beamWidthMin)
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

            // Set the location of beams
            //List<List<CADModel>> se

            // Get the beam depth by closest label and store it in CADModel.
            //List<List<CADModel>> sss = GetPairedLabel(CADModelList_List, path, layer_name);
            CADModelList_List = GetPairedLabel(CADModelList_List, path, layer_name, geoElem);

            List<(ElementId, string)> errorData = new List<(ElementId, string)>();
            // Create beams with corresponding sizes.
            foreach (List<CADModel> cadModelList in CADModelList_List)
            {
                // The depth is recorded in cADModel_A
                CADModel cADModel_A = cadModelList[0];
                CADModel cADModel_B = cadModelList[1];
                if (cADModel_A.Depth == "") continue;

                XYZ cADModel_A_StratPoint = cADModel_A.CurveArray.get_Item(0).GetEndPoint(0);
                XYZ cADModel_A_EndPoint = cADModel_A.CurveArray.get_Item(0).GetEndPoint(1);
                XYZ cADModel_B_StratPoint = cADModel_B.CurveArray.get_Item(0).GetEndPoint(0);
                XYZ cADModel_B_EndPoint = cADModel_B.CurveArray.get_Item(0).GetEndPoint(1);

                XYZ ChangeXYZ = new XYZ();

                double LineLength = (GetMiddlePoint(cADModel_A_StratPoint, cADModel_B_StratPoint)).DistanceTo(GetMiddlePoint(cADModel_A_EndPoint, cADModel_B_EndPoint));

                if (UnitsToCentimeters(LineLength) < 1)
                {
                    ChangeXYZ = cADModel_B_StratPoint;
                    cADModel_B_StratPoint = cADModel_B_EndPoint;
                    cADModel_B_EndPoint = ChangeXYZ;
                }

                XYZ mp1 = GetMiddlePoint(cADModel_A_StratPoint, cADModel_B_StratPoint);
                XYZ mp2 = GetMiddlePoint(cADModel_A_EndPoint, cADModel_B_EndPoint);
                Curve curve = Line.CreateBound(mp1, mp2);

                Line lineb = cADModel_A.CurveArray.get_Item(0) as Line;
                double width = lineb.Distance(cADModel_B.Location);
                width = Math.Round(Algorithm.UnitsToMillimeters(width)) / 10;
                String beamSize = width.ToString() + "x" + cadModelList[0].Depth;

                using (Transaction transaction = new Transaction(doc))
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
                    Parameter startOffsetParam = instance.get_Parameter(BuiltInParameter.STRUCTURAL_BEAM_END0_ELEVATION);
                    Parameter endOffsetParam = instance.get_Parameter(BuiltInParameter.STRUCTURAL_BEAM_END1_ELEVATION);
                    //MessageBox.Show(cadModelList[0].Elevation.ToString());
                    startOffsetParam.Set(MillimetersToUnits(cadModelList[0].Elevation));
                    endOffsetParam.Set(MillimetersToUnits(cadModelList[0].Elevation));
                    StructuralFramingUtils.DisallowJoinAtEnd(instance, 1);
                    StructuralFramingUtils.DisallowJoinAtEnd(instance, 0);
                    transaction.Commit();
                    if(cADModel_A.Error != "") errorData.Add((instance.Id, cADModel_A.Error));
                }
                ChangeBeamType(doc, beamSize);
            }
            if(errorData != null) ExportChecklistToCsv(errorData);

            FilteredElementCollector collector_b = new FilteredElementCollector(doc);
            List<Element> beams = collector_b
                .OfCategory(BuiltInCategory.OST_StructuralFraming)
                .WhereElementIsNotElementType()
                .ToList();
            return Result.Succeeded;
        }

        public XYZ PointCentimeterToUnit(XYZ point)
        {
            XYZ newPoint = new XYZ(
                CentimetersToUnits(point.X),
                CentimetersToUnits(point.Y),
                CentimetersToUnits(point.Z)
                );
            //XYZ newPoint = new XYZ(
            //    MillimetersToUnits(point.X),
            //    MillimetersToUnits(point.Y),
            //    MillimetersToUnits(point.Z)
            //    );
            return newPoint;
        }

        /// <summary>
        /// Get all the lines in this layer. 
        /// </summary>
        /// <param name="doc">RevitDocument</param>
        /// <param name="geoElem">GeometryElement</param>
        /// <param name="graphicsStyleId">GeometryElementID</param>
        /// <returns></returns>
        private List<CADModel> GetCurveArray(GeometryElement geoElem, ElementId graphicsStyleId, double gridSize)
        {
            List<CADModel> curveArray_List = new List<CADModel>();

            CurveArray allCurveArray = new CurveArray();

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
                            Line line = insObj as Line;
                            Line newLine = TransformLine(transform, line);
                            Line roundedLine = Line.CreateBound(Algorithm.RoundPoint(newLine.GetEndPoint(0), gridSize), Algorithm.RoundPoint(newLine.GetEndPoint(1), gridSize));
                            allCurveArray.Append(newLine);
                        }

                        if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
                        {
                            PolyLine polyLine = insObj as PolyLine;
                            IList<XYZ> points = polyLine.GetCoordinates();

                            for (int i = 0; i < points.Count - 1; i++)
                            {
                                Line line = Line.CreateBound(points[i], points[i + 1]);
                                line = TransformLine(transform, line);
                                Line roundedLine = Line.CreateBound(Algorithm.RoundPoint(line.GetEndPoint(0), gridSize), Algorithm.RoundPoint(line.GetEndPoint(1), gridSize));
                                Line newLine = roundedLine;
                                allCurveArray.Append(newLine);
                            }
                        }
                    }
                }
            }

            // Sort the CurveArray.
            List<Curve> sortedCurves = allCurveArray.Cast<Curve>().OrderBy(c => c.GetEndPoint(0).X).ToList();
            CurveArray sortedCurveArray = new CurveArray();
            foreach (Curve curve in sortedCurves)
            {
                sortedCurveArray.Append(curve);
            }

            foreach (Curve curve in sortedCurves)
            {
                CurveArray curveArray = new CurveArray();
                curveArray.Append(curve);

                XYZ startPoint = curve.GetEndPoint(0);
                XYZ endPoint = curve.GetEndPoint(1);
                XYZ MiddlePoint = GetMiddlePoint(startPoint, endPoint);
                double angle = (startPoint.Y - endPoint.Y) / startPoint.DistanceTo(endPoint);
                double rotation = Math.Asin(angle);

                CADModel cADModel = new CADModel
                {
                    CurveArray = curveArray,
                    Length = curve.Length,
                    Shape = "矩形梁",
                    Width = 300 / 304.8,
                    Location = MiddlePoint,
                    Rotation = rotation
                };
                curveArray_List.Add(cADModel);
            }

            // To verify the type of elements.
            //foreach (GeometryObject gObj in geoElem)
            //{
            //    GeometryInstance geomInstance = gObj as GeometryInstance;

            //    // Coordinates transformation
            //    Transform transform = geomInstance.Transform;

            //    if (null != geomInstance)
            //    {
            //        foreach (GeometryObject insObj in geomInstance.SymbolGeometry)
            //        {
            //            if (insObj.GraphicsStyleId.IntegerValue != graphicsStyleId.IntegerValue)
            //                continue;

            //            if (insObj.GetType().ToString() == "Autodesk.Revit.DB.Line")
            //            {
            //                Line line = insObj as Autodesk.Revit.DB.Line;
            //                Line newLine = TransformLine(transform, line);

            //                CurveArray curveArray = new CurveArray();
            //                curveArray.Append(newLine);

            //                XYZ startPoint = newLine.GetEndPoint(0);
            //                XYZ endPoint = newLine.GetEndPoint(1);
            //                XYZ MiddlePoint = GetMiddlePoint(startPoint, endPoint);
            //                double angle = (startPoint.Y - endPoint.Y) / startPoint.DistanceTo(endPoint);
            //                double rotation = Math.Asin(angle);

            //                CADModel cADModel = new CADModel
            //                {
            //                    CurveArray = curveArray,
            //                    Length = newLine.Length,
            //                    Shape = "矩形梁",
            //                    Location = MiddlePoint,
            //                    Rotation = rotation
            //                };
            //                curveArray_List.Add(cADModel);
            //            }

            //            if (insObj.GetType().ToString() == "Autodesk.Revit.DB.PolyLine")
            //            {
            //                PolyLine polyLine = insObj as PolyLine;
            //                IList<XYZ> points = polyLine.GetCoordinates();

            //                for (int i = 0; i < points.Count - 1; i++)
            //                {
            //                    Line line = Line.CreateBound(points[i], points[i + 1]);
            //                    line = TransformLine(transform, line);
            //                    Line newLine = line;
            //                    CurveArray curveArray = new CurveArray();
            //                    curveArray.Append(newLine);

            //                    XYZ startPoint = newLine.GetEndPoint(0);
            //                    XYZ endPoint = newLine.GetEndPoint(1);
            //                    XYZ MiddlePoint = GetMiddlePoint(startPoint, endPoint);
            //                    double angle = (startPoint.Y - endPoint.Y) / startPoint.DistanceTo(endPoint);
            //                    double rotation = Math.Asin(angle);

            //                    CADModel cADModel = new CADModel
            //                    {
            //                        CurveArray = curveArray,
            //                        Length = newLine.Length,
            //                        Shape = "矩形梁",
            //                        Width = 300 / 304.8,
            //                        Location = MiddlePoint,
            //                        Rotation = rotation
            //                    };
            //                    curveArray_List.Add(cADModel);
            //                }
            //            }
            //        }
            //    }
            //}
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

        public void ChangeBeamType(Document doc, string beamSize)
        {
            char separator = 'x';
            string[] parts = beamSize.Split(separator);
            double width = double.Parse(parts[0]);
            double depth = double.Parse(parts[1]);

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

        public List<List<CADModel>> GetPairedLabel(List<List<CADModel>> CADModelList_List, String path, String layer_name, GeometryElement geoElem)
        {
            Transform transform = null;
            foreach (GeometryObject gObj in geoElem)
            {
                GeometryInstance geomInstance = gObj as GeometryInstance;
                transform = geomInstance.Transform;
            }

            if (transform == null)
            {
                MessageBox.Show("Empty!");
            }

            foreach (List<CADModel> cadModelList in CADModelList_List)
            {

                XYZ location_beam = (cadModelList[0].Location + cadModelList[1].Location) / 2;
                XYZ start1 = cadModelList[0].CurveArray.get_Item(0).GetEndPoint(0);
                XYZ end1 = cadModelList[0].CurveArray.get_Item(0).GetEndPoint(1);
                XYZ start2 = cadModelList[1].CurveArray.get_Item(0).GetEndPoint(0);
                XYZ end2 = cadModelList[1].CurveArray.get_Item(0).GetEndPoint(1);

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
                                                        XYZ position = ConverCADPointToRevitPoint(text.Position);
                                                        position = PointCentimeterToUnit(position);
                                                        position = new XYZ(position.X, position.Y, location_beam.Z);
                                                        position = transform.OfPoint(position);

                                                        // If label in Beam lines.
                                                        if (!IsPointInsideRectangleOnXYPlane(start1, end1, start2, end2, position)) continue;

                                                        // If same rotation.
                                                        if (IsSameRotation(cadModelList[0], text.Rotation) && text.Layer == layer_name)
                                                        {
                                                            if (text.TextString.Contains("±"))
                                                            {
                                                                string pattern = @"±(\d+)";
                                                                Match match = Regex.Match(text.TextString, pattern);
                                                                if (match.Success) cadModelList[0].Elevation = 0; //mm
                                                            }
                                                            else if (text.TextString.Contains("+"))
                                                            {
                                                                string pattern = @"\+(\d+)";
                                                                Match match = Regex.Match(text.TextString, pattern);
                                                                if (match.Success)
                                                                {
                                                                    string numbersString = match.Groups[1].Value;
                                                                    double elevation = double.Parse(numbersString);
                                                                    cadModelList[0].Elevation = elevation;
                                                                }

                                                            }
                                                            else if (text.TextString.Contains("-"))
                                                            {
                                                                string pattern = @"\-(\d+)";
                                                                Match match = Regex.Match(text.TextString, pattern);
                                                                if (match.Success)
                                                                {
                                                                    string numbersString = match.Groups[1].Value;
                                                                    double elevation = double.Parse(numbersString) * (-1);
                                                                    cadModelList[0].Elevation = elevation;
                                                                }

                                                            }
                                                            else cadModelList[0].Elevation = 0;

                                                            if (text.TextString.Contains("H"))
                                                            {
                                                                try
                                                                {
                                                                    // Find the substring between "H(" and the next ")"
                                                                    int startIndex1 = text.TextString.IndexOf("H(") + 2;
                                                                    int endIndex1 = text.TextString.IndexOf(")", startIndex1);

                                                                    // Extract beamDepth from the substring
                                                                    string beamDepth = text.TextString.Substring(startIndex1, endIndex1 - startIndex1);

                                                                    // Try to parse and convert beamDepth
                                                                    int intBeamDepth = Int32.Parse(beamDepth) / 10;
                                                                    cadModelList[0].Depth = intBeamDepth.ToString(); //cm
                                                                }
                                                                catch (FormatException)
                                                                {
                                                                    // If parsing fails, show an error message with the actual text
                                                                    MessageBox.Show($"\"{text.TextString}\" is not in the correct format.", "Error");
                                                                    cadModelList[0].Depth = "400"; //cm
                                                                    cadModelList[0].Error = $"\"{text.TextString}\" is not in the correct format.";
                                                                }
                                                                catch (ArgumentOutOfRangeException)
                                                                {
                                                                    // If substring extraction fails due to invalid indices
                                                                    MessageBox.Show($"\"{text.TextString}\" does not contain the expected format (missing parentheses).", "Error");
                                                                    cadModelList[0].Depth = "400"; //cm
                                                                    cadModelList[0].Error = $"\"{text.TextString}\" does not contain the expected format (missing parentheses).";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                cadModelList[0].Depth = "400"; //cm
                                                                cadModelList[0].Error = $"\"{text.TextString}\" does not contain H.";
                                                            }
                                                        }
                                                        break;

                                                    case "AcDbMText":
                                                        Teigha.DatabaseServices.MText text_m = (Teigha.DatabaseServices.MText)entity2;
                                                        XYZ location = ConverCADPointToRevitPoint(text_m.Location);
                                                        location = PointCentimeterToUnit(location);
                                                        location = new XYZ(location.X, location.Y, location_beam.Z);
                                                        location = transform.OfPoint(location);

                                                        // If label in Beam lines.
                                                        if (!IsPointInsideRectangleOnXYPlane(start1, end1, start2, end2, location)) continue;

                                                        // If same roatation.
                                                        if (IsSameRotation(cadModelList[0], text_m.Rotation) && text_m.Layer == layer_name)
                                                        {
                                                            //string beamDepth = "400";
                                                            if (text_m.Text.Contains("±"))
                                                            {
                                                                string pattern = @"±(\d+)";
                                                                Match match = Regex.Match(text_m.Text, pattern);
                                                                if (match.Success) cadModelList[0].Elevation = 0; // mm
                                                            }
                                                            else if (text_m.Text.Contains("+"))
                                                            {
                                                                string pattern = @"\+(\d+)";
                                                                Match match = Regex.Match(text_m.Text, pattern);
                                                                if (match.Success)
                                                                {
                                                                    string numbersString = match.Groups[1].Value;
                                                                    int elevation = int.Parse(numbersString);
                                                                    cadModelList[0].Elevation = elevation; // mm
                                                                }

                                                            }
                                                            else if (text_m.Text.Contains("-"))
                                                            {
                                                                string pattern = @"\-(\d+)";
                                                                Match match = Regex.Match(text_m.Text, pattern);
                                                                if (match.Success)
                                                                {
                                                                    string numbersString = match.Groups[1].Value;
                                                                    int elevation = int.Parse(numbersString) * (-1);
                                                                    cadModelList[0].Elevation = elevation;
                                                                }
                                                            }
                                                            else cadModelList[0].Elevation = 0;

                                                            if (text_m.Text.Contains("H"))
                                                            {
                                                                try
                                                                {
                                                                    // Find the substring between "H(" and the next ")"
                                                                    int startIndex1 = text_m.Text.IndexOf("H(") + 2;
                                                                    int endIndex1 = text_m.Text.IndexOf(")", startIndex1);

                                                                    // Extract beamDepth from the substring
                                                                    string beamDepth = text_m.Text.Substring(startIndex1, endIndex1 - startIndex1);

                                                                    // Try to parse and convert beamDepth
                                                                    int intBeamDepth = Int32.Parse(beamDepth) / 10;
                                                                    cadModelList[0].Depth = intBeamDepth.ToString(); //cm
                                                                }
                                                                catch (FormatException)
                                                                {
                                                                    // If parsing fails, show an error message with the actual text
                                                                    MessageBox.Show($"\"{text_m.Text}\" is not in the correct format.", "Error");
                                                                    cadModelList[0].Depth = "400"; //cm
                                                                    cadModelList[0].Error = $"\"{text_m.Text}\" is not in the correct format.";
                                                                }
                                                                catch (ArgumentOutOfRangeException)
                                                                {
                                                                    // If substring extraction fails due to invalid indices
                                                                    MessageBox.Show($"\"{text_m.Text}\" does not contain the expected format (missing parentheses).", "Error");
                                                                    cadModelList[0].Depth = "400"; //cm
                                                                    cadModelList[0].Error = $"\"{text_m.Text}\" does not contain the expected format (missing parentheses).";
                                                                }
                                                            }
                                                            else
                                                            {
                                                                cadModelList[0].Depth = "400"; //cm
                                                                cadModelList[0].Error = $"\"{text_m.Text}\" does not contain H.";
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
            }
            return CADModelList_List;
        }

        public static bool IsPointInsideRectangleOnXYPlane(XYZ p1, XYZ p2, XYZ p3, XYZ p4, XYZ pA)
        {
            // 將四個頂點與測試點投影到 XY 平面，只取 X、Y
            // 取得最小、最大 X 與 Y
            double minX = new double[] { p1.X, p2.X, p3.X, p4.X }.Min();
            double maxX = new double[] { p1.X, p2.X, p3.X, p4.X }.Max();
            double minY = new double[] { p1.Y, p2.Y, p3.Y, p4.Y }.Min();
            double maxY = new double[] { p1.Y, p2.Y, p3.Y, p4.Y }.Max();

            // 取得測試點投影後的 X、Y
            double testX = pA.X;
            double testY = pA.Y;

            // 若測試點 X、Y 都在 min ~ max 範圍內，則視為在矩形內
            return (testX >= minX && testX <= maxX &&
                    testY >= minY && testY <= maxY);
        }

        public Boolean IsSameRotation(CADModel cadmodel, double textRotation)
        {
            // Take the difference between the two angles.
            double diff = cadmodel.Rotation - textRotation;

            // Normalize the difference to the range -pi to pi.
            diff = diff % (2 * Math.PI);
            if (diff < -Math.PI)
            {
                diff += 2 * Math.PI;
            }
            else if (diff > Math.PI)
            {
                diff -= 2 * Math.PI;
            }

            // Check if the absolute difference is less than a certain tolerance or close to pi.
            // Here, the tolerance is set to 0.01 radians, but you can adjust this value as needed.
            double tolerance = 0.01;
            if (Math.Abs(diff) < tolerance || Math.Abs(Math.Abs(diff) - Math.PI) < tolerance)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public double UnitsToCentimeters(double value)
        {
            return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Centimeters);
        }

        public double CentimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Centimeters);
        }

        public double MillimetersToUnits(double value)
        {
            return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Millimeters);
        }

        /// <summary>
        /// Exports a list of (ElementId, string) pairs to a CSV file named "Checklist_Beam.csv"
        /// in the "Checklist_Modeling" folder on the user's Desktop. The CSV has two columns:
        /// - ElementId
        /// - Error Message
        /// </summary>
        /// <param name="data">A collection of pairs containing ElementId and an error message.</param>
        public static void ExportChecklistToCsv(IEnumerable<(ElementId ElementId, string ErrorMessage)> data)
        {
            // Determine folder path on user's desktop
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string folderPath = Path.Combine(desktopPath, "Checklist_Modeling");

            // Ensure folder exists (create if not)
            Directory.CreateDirectory(folderPath);

            // Create the CSV file path
            string filePath = Path.Combine(folderPath, "Checklist_Beam.csv");

            // Write to CSV
            // - Overwrite file if it already exists (set append to 'false')
            // - Use UTF-8 encoding
            using (StreamWriter writer = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                // Write CSV header
                writer.WriteLine("ElementId,Error Message");

                // Write each row
                foreach ((ElementId ElementId, string ErrorMessage) pair in data)
                {
                    // ElementId is typically an integer; use pair.ElementId.IntegerValue
                    writer.WriteLine($"{pair.ElementId.IntegerValue},{pair.ErrorMessage}");
                }
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
            Error = "";
        }

        /// <summary>
        /// Error
        /// </summary>
        public string Error { get; set; }

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
}
