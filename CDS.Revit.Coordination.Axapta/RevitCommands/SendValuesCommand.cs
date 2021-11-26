using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CDS.Revit.Coordination.Services.Axapta;
using CDS.Revit.Coordination.Services.Excel;
using CDS.Revit.Coordination.Services.Revit;


namespace CDS.Revit.Coordination.Axapta
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    public class SendValuesCommand : IExternalCommand
    {
        public static Document Doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            UIApplication uiapp = uidoc.Application;
            Document doc = commandData.Application.ActiveUIDocument.Document;
            Doc = doc;
            var mainWindow = new MainWindow();
            mainWindow.ShowDialog();
            return Result.Succeeded;
        }
    }
}
