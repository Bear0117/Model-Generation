//using Autodesk.Revit.Attributes;
//using Autodesk.Revit.UI;
//using System;
//using System.Configuration.Assemblies;
//using System.IO;
//using System.Windows.Media.Imaging;

//namespace Modeling
//{
//    [Autodesk.Revit.Attributes.Transaction(TransactionMode.Manual)]
//    [Autodesk.Revit.Attributes.Regeneration(RegenerationOption.Manual)]
//    public class Icon : IExternalApplication
//    {
//        public Result OnStartup(UIControlledApplication app)
//        {
//            // 要執行的dll與按鈕圖片檔案路徑
//            string firstBtnDll = @"C:\Users\Bear\OneDrive\桌面11\系統模板研究\系統模板研究\bin\Debug\Modeling.dll";
//            if (!File.Exists(firstBtnDll))
//            {
//                TaskDialog.Show("Error", "DLL not found: " + firstBtnDll);
//                return Result.Failed; // 或者其他错误处理
//            }

//            string picPath = @"C:\Users\Bear\OneDrive\桌面11\系統模板研究\系統模板研究\bin\Debug\Auto Creation Icon.png";

//            // 創建一個新的工具列
//            string tabName = "Revit API";
//            app.CreateRibbonTab(tabName);
//            // 添加面板
//            RibbonPanel firstBtnPanel = app.CreateRibbonPanel(tabName, "Model Generation");
//            // FirstBtn按鈕創建
//            PushButton firstBtn = firstBtnPanel.AddItem(new PushButtonData("Model Generation", "Modeling", firstBtnDll, "Modeling.Icon")) as PushButton;
//            // 給按鈕添加圖片
//            Uri firstBtnImage = new Uri(picPath);
//            BitmapImage firstBtnLargeImage = new BitmapImage(firstBtnImage);
//            firstBtn.LargeImage = firstBtnLargeImage;

//            return Result.Succeeded;
//        }
//        public Result OnShutdown(UIControlledApplication app)
//        {
//            return Result.Succeeded;
//        }
//    }
//}