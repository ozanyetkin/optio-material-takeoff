using System;
using System.IO;

using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using DesignAutomationFramework;

namespace OPTIO_Converter
{
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    public class Class1 : IExternalDBApplication
    {
        public ExternalDBApplicationResult OnStartup(Autodesk.Revit.ApplicationServices.ControlledApplication app)
        {
            DesignAutomationBridge.DesignAutomationReadyEvent += HandleDesignAutomationReadyEvent;
            return ExternalDBApplicationResult.Succeeded;
        }

        public ExternalDBApplicationResult OnShutdown(Autodesk.Revit.ApplicationServices.ControlledApplication app)
        {
            return ExternalDBApplicationResult.Succeeded;
        }

        public void HandleDesignAutomationReadyEvent(object sender, DesignAutomationReadyEventArgs e)
        {
            e.Succeeded = true;
            RevitExportMaterialTakeoff(e.DesignAutomationData);
        }

        public static void RevitExportMaterialTakeoff(DesignAutomationData data)
        {
            if (data == null) throw new ArgumentNullException(nameof(data));

            Application rvtApp = data.RevitApp;
            if (rvtApp == null) throw new InvalidDataException(nameof(rvtApp));

            string modelPath = data.FilePath;
            if (String.IsNullOrWhiteSpace(modelPath)) throw new InvalidDataException(nameof(modelPath));

            Document doc = data.RevitDoc;
            if (doc == null) throw new InvalidOperationException("Could not open document.");

            using (Transaction transaction = new Transaction(doc))
            {
                transaction.Start("Export Material Takeoff");

                Console.WriteLine("Transaction has started without a problem");

                FilteredElementCollector col = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule));

                ViewScheduleExportOptions opt = new ViewScheduleExportOptions();

                foreach (ViewSchedule vs in col)
                {
                    Directory.CreateDirectory("materialTakeoffs");

                    vs.Export("materialTakeoffs", vs.Name + ".csv", opt);
                }

                transaction.Commit();
            }
        }
    }
}
