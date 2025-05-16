using Autodesk.Revit.UI;
using System;
using System.Reflection;
using System.IO;
using System.Windows.Media.Imaging;

namespace Modeling
{
    public class UIPanel : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            // 1. Create new tab.
            string tabName = "Modeling";
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch (Exception)
            {

            }

            // 2. Add new panel to tab.
            RibbonPanel panel = application.CreateRibbonPanel(tabName, "Modeling Tools");

            string assemblyPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string assemblyDirectory = Path.GetDirectoryName(assemblyPath);
            string anotherDllPath = Path.Combine(assemblyDirectory,"Debug" ,"CreatBridgeForRevit2018.dll");


            // 3. Add new buttons.

            AddButton(panel, "Tool1", "Columns", "Modeling.AutoCreateColumns", "This function will automatically create columns", Path.Combine(assemblyDirectory, "icons", "column.png"));
            AddButton(panel, "Tool2", " Beams ", "Modeling.AutoCreateBeams", "This function will automatically create beams", Path.Combine(assemblyDirectory, "icons", "beam.png"));
            AddButton(panel, "Tool3", " Slabs ", "Modeling.AutoCreateSlabs", "This function will automatically create slabs", Path.Combine(assemblyDirectory, "icons", "slab.png"));
            AddButton(panel, "Tool4", " Walls ", "Modeling.AutoCreateWalls", "This function will automatically create walls", Path.Combine(assemblyDirectory, "icons", "wall.png"));
            AddButton(panel, "Tool5", "Additional Walls", "Modeling.AutoCreateAdditionalWalls", "This function will automatically create additional walls", Path.Combine(assemblyDirectory, "icons", "addWall.png"));
            AddButton(panel, "Tool6", "Openings", "Modeling.AutoCreateWindows", "This function will automatically create openings", Path.Combine(assemblyDirectory, "icons", "opening.png"));

            RibbonPanel panel_adj = application.CreateRibbonPanel(tabName, "Adjusting Tools");
            AddButtonToAnotherDll(panel_adj, "Tool0", "Read CAD", "CreatBridgeForRevit2018.ReadCAD.ReadCADCmd", "This function read the CAD text info.", Path.Combine(assemblyDirectory, "icons", "read.png"), anotherDllPath);
            AddButton(panel_adj, "Tool1", "Model Lines", "Modeling.CreateWallModelline", "This function will create model lines according to the selected layer.", Path.Combine(assemblyDirectory, "icons", "wallLines.png"));
            AddButton(panel_adj, "Tool2", "Customized Walls", "Modeling.TwoLineToWall_A", "Select two lines in sequence to create a wall.", Path.Combine(assemblyDirectory, "icons", "twoLines.png"));
            AddButton(panel_adj, "Tool3", "Adjust Wall Height", "Modeling.AdjustWallHeight", "This function will adjust the height of walls and join the contacted elements.", Path.Combine(assemblyDirectory, "icons", "adjustment.png"));
            AddButton(panel_adj, "Tool4", "Gap Detection", "Modeling.GapDetection", "This function will list the elements in group.", Path.Combine(assemblyDirectory, "icons", "gap.png"));

            //  Test//132594//55d6v2
            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }

        private void AddButton(RibbonPanel panel, string name, string text, string className, string toolTip, string imagePath)
        {
            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            PushButtonData buttonData = new PushButtonData(name, text, assemblyPath, className)
            {
                ToolTip = toolTip
            };
                
            if (!string.IsNullOrEmpty(imagePath))
            {
                buttonData.LargeImage = new System.Windows.Media.Imaging.BitmapImage(new Uri(imagePath, UriKind.RelativeOrAbsolute));
            }

            panel.AddItem(buttonData);
        }
        private void AddButtonToAnotherDll(RibbonPanel panel,string name,string text,string className,string toolTip,string imagePath,string dllPath)
        {
            PushButtonData buttonData = new PushButtonData(name, text, dllPath, className)
            {
                ToolTip = toolTip
            };

            if (!string.IsNullOrEmpty(imagePath))
            {
                buttonData.LargeImage = new BitmapImage(new Uri(imagePath, UriKind.RelativeOrAbsolute));
            }

            panel.AddItem(buttonData);
        }
    }
}
