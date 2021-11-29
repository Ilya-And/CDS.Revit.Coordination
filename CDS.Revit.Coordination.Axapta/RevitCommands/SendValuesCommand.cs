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
using CDS.Revit.Coordination.Axapta.Views;
using Autodesk.Revit.ApplicationServices;
using System.Reflection;
using System.Windows;

namespace CDS.Revit.Coordination.Axapta
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    public class SendValuesCommand : IExternalCommand
    {
        public static Autodesk.Revit.ApplicationServices.Application App;
        public static string AssemblyPath { get; private set; }
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;
            var arrayStr = thisAssemblyPath.Split('\\').ToList();
            arrayStr.Remove(arrayStr[arrayStr.Count - 1]);
            string folderPath = string.Join("\\", arrayStr.ToArray());
            AssemblyPath = folderPath;

            try
            {
                Assembly.LoadFrom(folderPath + "\\Newtonsoft.Json.dll");
                Assembly.LoadFrom(folderPath + "\\UnidecodeSharpFork.dll");
                Assembly.LoadFrom(folderPath + "\\ExcelDataReader.dll");
                Assembly.LoadFrom(folderPath + "\\ExcelDataReader.DataSet.dll");
                Assembly.LoadFrom(folderPath + "\\CDS.Revit.Coordination.Services.dll");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return 0;
            }

            Autodesk.Revit.ApplicationServices.Application app = commandData.Application.Application;
            App = app;

            var mainWindow = new MainWindow();
            mainWindow.ShowDialog();
            return Result.Succeeded;
        }
    }
}
