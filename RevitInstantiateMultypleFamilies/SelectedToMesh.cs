using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Kinross.Share.Utils;
using Kinross.RevitShare.Utils;

namespace RevitInstantiateMultypleFamilies
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class SelectedToMesh : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            ICollection<ElementId> selectedElements = commandData.Application.ActiveUIDocument.Selection.GetElementIds();

            if (selectedElements.Count == 0)
            {
                message = "Select elements first!";
                return Result.Failed;
            }


            foreach (ElementId id in selectedElements)
            {
                try
                {
                    
                }
                catch(Exception)
                {

                    return Result.Failed;
                }

                Element elem = doc.GetElement(id);
                //var geomElem = GeometryUtil.GetGeometryElement(elem);
                Options opt = new Options();
                opt.DetailLevel = ViewDetailLevel.Fine;
                GeometryElement geomElem = elem.get_Geometry(opt);

                var meshSolidCapsule = GeometryUtil.GetMeshSolidCapsule(geomElem);
                if (meshSolidCapsule == null)
                {
                    TaskDialog.Show("Mesh Solid Capsule",
                        "Element has no mesh/solid information. Try selecting elements one by one.");
                    break;
                }

                var mesh = ConvertUtil.ConvertMesh(doc, meshSolidCapsule);
                if (mesh == null)
                {
                    TaskDialog.Show("Mesh Converter",
                        "Unable to convert element. Try selecting elements one by one.");
                    break;
                }

                string outPath = @"D:\SECTIONS\rectangular\"+id+".json";
                using (StreamWriter sw = new StreamWriter(outPath))
                {
                    sw.WriteLine(JsonSerializer.Serialize(mesh));
                }
            }

            TaskDialog.Show("Exporting meshes",
                "Completed. Check output folder");

            return Result.Succeeded;
        }
    }



}
