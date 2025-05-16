using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Newtonsoft.Json;
using System;
using System.IO;
using Microsoft.Win32;

namespace Modeling
{
    public static class ModelingParam
    {
        public static Parameters parameters { get; set; }

        public static void Initialize()
        {
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string jsonFilePath = Path.Combine(desktopPath, "ModelingParam.json");
            if (!File.Exists(jsonFilePath))
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Title = "Select ModelingParam.json",
                    Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*",
                    InitialDirectory = desktopPath
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    jsonFilePath = openFileDialog.FileName; // 更新路徑
                }
                else
                {
                    System.Windows.MessageBox.Show("File not found. Initialization aborted.", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
                    return;
                }
            }

            try
            {
                string jsonContent = File.ReadAllText(jsonFilePath);
                parameters = JsonConvert.DeserializeObject<Parameters>(jsonContent);

                System.Windows.MessageBox.Show("Parameters successfully loaded!", "Success", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Failed to load parameters: {ex.Message}", "Error", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            }
        }
    }

    public class Parameters
    {
        public General General { get; set; }
        public ColumnParam ColumnParam { get; set; }
        public BeamParam BeamParam { get; set; }
        public SlabParam SlabParam { get; set; }
        public WallParam WallParam { get; set; }
        public OpeningParam OpeningParam { get; set; }
    }

    public class General
    {
        public double GridSize { get; set; }
        public double LevelHeight { get; set; }
        public double Gap {  get; set; }
    }

    public class ColumnParam { public double[] ColumnWidthsRange { get; set; } }
    public class BeamParam {
        public double[] BeamWidthRange { get; set; }
    }
    public class SlabParam { public double[] slabThicknessRange { get; set; } }
    public class WallParam { 
        public int[] WallWidths { get; set; }
        public int[] AdditionalWallWidths { get; set; }
    }
    public class OpeningParam { public double KerbHeight { get; set; } }
}
