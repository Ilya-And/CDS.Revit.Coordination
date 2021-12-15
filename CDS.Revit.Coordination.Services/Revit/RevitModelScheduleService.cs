using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace CDS.Revit.Coordination.Services.Revit
{
    public class RevitModelScheduleService
    {
        private static string MakeValidFileName(string name)
        {
            string invalidChars = System.Text.RegularExpressions.Regex.Escape(new string(System.IO.Path.GetInvalidFileNameChars()));
            string invalidReStr = string.Format(@"([{0}]*\.+$)|([{0}]+)", invalidChars);
            return System.Text.RegularExpressions.Regex.Replace(name, invalidReStr, "_");
        }

        private List<ViewSchedule> GetViewSchedulesForExport(Document doc)
        {
            List<ViewSchedule> viewSchedulesResult = new List<ViewSchedule>();

            var viewScheduleIds = new FilteredElementCollector(doc).OfClass(typeof(ViewSchedule)).WhereElementIsNotElementType().ToElementIds();

            foreach(ElementId elementId in viewScheduleIds)
            {
                ViewSchedule viewSchedule = doc.GetElement(elementId) as ViewSchedule;
                string viewScheduleName = viewSchedule?.Name;

                if (viewScheduleName.Split('_')[0] == "Axapta") viewSchedulesResult.Add(viewSchedule);
            }

            return viewSchedulesResult;
        }

        /*Метод экспорта спецификации в формат .csv
         */
        public void ExportToCSV(Document doc, string filePath, string projName)
        {
            var viewSchedulesList = GetViewSchedulesForExport(doc);

            foreach(ViewSchedule viewSchedule in viewSchedulesList)
            {
                string viewName = viewSchedule.Name;
                ViewScheduleExportOptions exportOptions = new ViewScheduleExportOptions();
                exportOptions.FieldDelimiter = ";";
                exportOptions.TextQualifier = ExportTextQualifier.None;

                string docName = doc.Title.ToString();

                docName = docName.Split('.')[0];

                if (!viewSchedule.Name.Contains("<"))
                {
                    viewSchedule.Export(filePath, $"{projName}_{docName}_{viewName}.csv", exportOptions);
                }
            }
        }
    }
}
