using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDS.Revit.Coordination.Axapta.Services.Revit
{
    public class RevitModelScheduleService
    {
        private static string MakeValidFileName(string name)
        {
            string invalidChars = System.Text.RegularExpressions.Regex.Escape(new string(System.IO.Path.GetInvalidFileNameChars()));
            string invalidReStr = string.Format(@"([{0}]*\.+$)|([{0}]+)", invalidChars);
            return System.Text.RegularExpressions.Regex.Replace(name, invalidReStr, "_");
        }

        //Метод экспорта спецификации в формат CSV
        public void ExportToCSV(ViewSchedule viewSchedule, string filePath)
        {
            string viewName = MakeValidFileName(viewSchedule.Name);
            ViewScheduleExportOptions exportOptions = new ViewScheduleExportOptions();
            if (!viewSchedule.Name.Contains("<"))
            {
                viewSchedule.Export(filePath, viewName + ".csv", exportOptions);
            }
        }
        public Dictionary<string, List<string>> ToDictionary(ViewSchedule viewSchedule)
        {
            Dictionary<string, List<string>> dictionaryFromSchedule = new Dictionary<string, List<string>>();

            return dictionaryFromSchedule;
        }
        public Dictionary<string, List<string>> ToDictionary(string filePath)
        {
            Dictionary<string, List<string>> dictionaryFromSchedule = new Dictionary<string, List<string>>();

            return dictionaryFromSchedule;
        }
    }
}
