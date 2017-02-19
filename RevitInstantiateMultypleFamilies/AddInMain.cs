using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI.Events;

namespace RevitInstantiateMultypleFamilies
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    //[Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    //[Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    public class AddInMain : IExternalCommand
    {
        public Autodesk.Revit.UI.Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //hardcode folder with families path here
            const string folderPath = @"D:\FAMILIES TO LOAD\Lighting fixtures Not classified";
            string[] paths = Directory.GetFiles(folderPath);
            TaskDialog.Show("Loading families",
                "Path to folder, containing families (should be changed manually in code):\n" +
                folderPath +
                "\n\nNo more than first 100 families will be loaded in current project.");

            Document doc = null;
            try
            {
                doc = commandData.Application.ActiveUIDocument.Document;
            }
            catch (Exception)
            {
                doc = commandData.Application.Application.NewProjectDocument(UnitSystem.Metric);
                //Show document view
            }

#region familyLoading
            List<Family> families = new List<Family>();
            using (Transaction trans = new Transaction(doc, "Family loading"))
            {
                trans.Start();
#if DEBUG
                Debug.WriteLine("LOGGING STARTED");
#endif
                //loading no more than 100 families
                for (int i = 0; i < paths.Length & i < 76; ++i)
                {
                    try
                    {
                        //trans.Start();
                        Family family = null;
                        //use LoadFamilySymbol to improve performance
                        if (doc.LoadFamily(paths[i], out family))
                            families.Add(family);
                        //trans.Commit();
                        Debug.WriteLine("Loaded " + paths[i]);
                    }
                    catch (Exception)
                    {
                        //trans.RollBack();
                        Debug.WriteLine("Unable to load " + paths[i]);
                    }
                }
                trans.Commit();
            }
#endregion

            using (Transaction trans = new Transaction(doc, "Instantiation"))
            {
                try
                {
                    trans.Start();

                    instantiateFamilySymbols(families, doc);

                    trans.Commit();
                }
                catch (Exception)
                {
                    trans.RollBack();
                }
#if DEBUG
                Debug.WriteLine("LOGGING STOPED");
#endif
            }

            //Save and Close doc if it was programmatically created
            //const string savePath = @"D:\LOADED FAMILIES.rvt";
            //if (System.IO.File.Exists(savePath))
            //    File.Delete(savePath);
            //doc.SaveAs(savePath);
            //doc.Close(false);

            TaskDialog.Show("Loading families",
                "Completed. Save project, open new and don't forget to remove first 100 files from folder!");

            return Autodesk.Revit.UI.Result.Succeeded;
        }

        void instantiateFamilySymbols(List<Family> families, Document doc)
        {
            double x = 1.0, y = 1.0;
            foreach (Family fml in families)
            {
                ISet<ElementId> familySymbolIds = fml.GetFamilySymbolIds();
                foreach (ElementId id in familySymbolIds)
                {
                    //subtransaction to be added?
                    try
                    {
                        FamilySymbol symbol = fml.Document.GetElement(id) as FamilySymbol;
                        symbol.Activate();
                        XYZ location = new XYZ(x, y, 0.0);
                        doc.Create.NewFamilyInstance(location, symbol, StructuralType.NonStructural);
                        y += 2.0;
                        Debug.WriteLine("Created " + fml.Name);
                        break; //need to create only one type of the family
                    }
                    catch
                    {
                        Debug.WriteLine("Unable to create " + fml.Name);
                    }
                }
            }
        }

        void createScene(Document doc)
        {
            //get 0.000 level
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> collection = collector.OfClass(typeof(Level)).ToElements();
            Level lvl = null;
            foreach (Element elem in collection)
            {
                if (elem is Level)
                {
                    lvl = (Level) elem;
                    break; //need only one of existing levels
                }
            }
            if (lvl == null) lvl = Level.Create(doc, 0d);

            XYZ start = new XYZ(0, 0, 0);
            XYZ end = new XYZ(0, 100, 0);
            XYZ right = new XYZ(25, 100, 0);
            XYZ left = new XYZ(25, 0, 0);
            Line wallNorth = Line.CreateBound(start, end);
            Line wallEast = Line.CreateBound(end, right);
            Line wallSouth = Line.CreateBound(right, left);
            Line wallWest = Line.CreateBound(left, start);

            //subtransaction to be added
            Wall.Create(doc, wallNorth, lvl.Id, false);
            Wall.Create(doc, wallEast, lvl.Id, false);
            Wall.Create(doc, wallSouth, lvl.Id, false);
            Wall.Create(doc, wallWest, lvl.Id, false);

            //create ceiling

            //delete front walls


            CurveArray flrScetch = new CurveArray();
            flrScetch.Append(wallNorth);
            flrScetch.Append(wallEast);
            flrScetch.Append(wallSouth);
            flrScetch.Append(wallWest);

            //doc.Create.NewFloor(flrScetch, false);
        }
    }
}