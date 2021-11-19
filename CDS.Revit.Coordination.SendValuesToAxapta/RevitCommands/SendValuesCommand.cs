using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CDS.Revit.Coordination.Services;

namespace CDS.Revit.Coordination.SendValuesToAxapta
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    public class SendValuesCommand : IExternalCommand
    {
        private Document _doc;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            UIApplication uiapp = uidoc.Application;
            Document doc = commandData.Application.ActiveUIDocument.Document;
            _doc = doc;
            var mainWindow = new MainWindow();
            mainWindow.ShowDialog();
            return Result.Succeeded;
        }
    }
}
