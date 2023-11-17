using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;
using Teigha.Runtime;
using Teigha.DatabaseServices;
using System.IO;
using Teigha.Geometry;
using System;
using System.Text.RegularExpressions;
using System.Reflection;

namespace modeling
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]

    class PJ_catchspecificlayerandtext : IExternalCommand
    {
        Autodesk.Revit.ApplicationServices.Application app;
        Document doc;
        string layername = null;


        public class CADTextModel
        {
            private string Htext;

            private string Wtext;

            private XYZ location;

            private double distant;

            private double rotation;


            public string HText
            {
                get
                {
                    return Htext;
                }

                set
                {
                    Htext = value;
                }
            }

            public string WText
            {
                get
                {
                    return Wtext;
                }

                set
                {
                    Wtext = value;
                }
            }

            public double Distant
            {
                get
                {
                    return distant;
                }

                set
                {
                    distant = value;
                }
            }

            public XYZ Location
            {
                get
                {
                    return location;
                }

                set
                {
                    location = value;
                }
            }

            public double Rotation
            {
                get
                {
                    return rotation;
                }

                set
                {
                    rotation = value;
                }
            }


        }
        public string GetCADPath(ElementId cadLinkTypeID, Document revitDoc)
        {
            CADLinkType cadLinkType = revitDoc.GetElement(cadLinkTypeID) as CADLinkType;
            return ModelPathUtils.ConvertModelPathToUserVisiblePath(cadLinkType.GetExternalFileReference().GetAbsolutePath());
        }
        /// CAD內的點為轉換
        public static XYZ ConverCADPointToRevitPoint(Point3d point)
        {
            double MillimetersToUnits(double value)
            {
                return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Millimeters);
            }

            return new XYZ(MillimetersToUnits(point.X), MillimetersToUnits(point.Y), MillimetersToUnits(point.Z));


        }

        public CADTextModel GetCADTextInfoparing(string dwgFile, XYZ midpoint)
        {
            CADTextModel CADModels = new CADTextModel();
            List<ObjectId> allObjectId = new List<ObjectId>();

            double comparedistanceH = 0;
            double distanceBetweenH = 10000000;

            //double comparedistanceW = 0;
            //double distanceBetweenW = 10000000;

            using (new Services())
            {
                using (Database database = new Database(false, false))
                {
                    database.ReadDwgFile(dwgFile, FileShare.Read, true, "");
                    using (var trans = database.TransactionManager.StartTransaction())
                    {
                        using (BlockTable table = (BlockTable)database.BlockTableId.GetObject(OpenMode.ForRead))
                        {
                            using (SymbolTableEnumerator enumerator = table.GetEnumerator())
                            {
                                while (enumerator.MoveNext())
                                {
                                    using (BlockTableRecord record = (BlockTableRecord)enumerator.Current.GetObject(OpenMode.ForRead))
                                    {
                                        foreach (ObjectId id in record)
                                        {
                                            Entity entity = (Entity)id.GetObject(OpenMode.ForRead, false, false);

                                            if (entity.Layer == layername)
                                            {
                                                switch (entity.GetRXClass().Name)
                                                {
                                                    case "AcDbText":
                                                        Teigha.DatabaseServices.DBText text = (Teigha.DatabaseServices.DBText)entity;



                                                        comparedistanceH = midpoint.DistanceTo(ConverCADPointToRevitPoint(text.Position));
                                                        if (distanceBetweenH > comparedistanceH)
                                                        {
                                                            distanceBetweenH = comparedistanceH;
                                                            CADModels.HText = text.TextString;
                                                            CADModels.Location = ConverCADPointToRevitPoint(text.Position);
                                                            CADModels.Rotation = text.Rotation;
                                                        }



                                                        break;

                                                    case "AcDbMText":
                                                        Teigha.DatabaseServices.MText mText = (Teigha.DatabaseServices.MText)entity;



                                                        comparedistanceH = midpoint.DistanceTo(ConverCADPointToRevitPoint(mText.Location));
                                                        if (distanceBetweenH > comparedistanceH)
                                                        {

                                                            distanceBetweenH = comparedistanceH;
                                                            CADModels.HText = mText.Text;
                                                            CADModels.Location = ConverCADPointToRevitPoint(mText.Location);
                                                            CADModels.Rotation = mText.Rotation;
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



            return CADModels;
        }

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;
            Level level = new FilteredElementCollector(doc).OfClass(typeof(Level)).First<Element>(e => e.Name.Equals("2F")) as Level;

            //string dll = @"C:\Users\james\Desktop\manicotti-main\Manicotti\Resources\lib\TD_Mgd_3.03_9.dll";
            //Assembly a = Assembly.UnsafeLoadFrom(dll);

            //抓取正確位置的文字
            Reference reference = uidoc.Selection.PickObject(ObjectType.PointOnElement);
            Element elem2 = doc.GetElement(reference);
            GeometryObject geoObj = elem2.GetGeometryObjectFromReference(reference);

            //抓取正確文字的世界座標
            XYZ pointontext = (XYZ)reference.GlobalPoint;

            TaskDialog.Show("Revit", pointontext.ToString());


            //建立CAD連結路徑
            string path = GetCADPath(elem2.GetTypeId(), doc);


            //抓取圖層名稱
            Category targetCategory1 = null;
            ElementId graphicsStyleId1 = null;

            if (geoObj.GraphicsStyleId != ElementId.InvalidElementId)
            {
                graphicsStyleId1 = geoObj.GraphicsStyleId;
                GraphicsStyle gs1 = doc.GetElement(geoObj.GraphicsStyleId) as GraphicsStyle;
                if (gs1 != null)
                {
                    targetCategory1 = gs1.GraphicsStyleCategory;
                    layername = gs1.GraphicsStyleCategory.Name;

                    TaskDialog.Show("Revit", layername);
                }
            }

            CADTextModel opening = GetCADTextInfoparing(path, pointontext);

            TaskDialog.Show("Revit", opening.HText);
            //TaskDialog.Show("Revit", opening.WText);
            TaskDialog.Show("Revit", opening.Location.ToString());
            TaskDialog.Show("Revit", opening.Rotation.ToString());


            return Result.Succeeded;

        }
    }
}