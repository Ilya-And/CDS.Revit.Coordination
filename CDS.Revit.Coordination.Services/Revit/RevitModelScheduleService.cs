using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

        /*Метод экспорта спецификации в формат .csv
         */
        public void ExportToCSV(ViewSchedule viewSchedule, string filePath, string projName, Document doc)
        {
            string viewName = MakeValidFileName(viewSchedule.Name);
            ViewScheduleExportOptions exportOptions = new ViewScheduleExportOptions();

            string docName = doc.PathName.Split('.')[0];
            int index = docName.LastIndexOf('_');
            docName = docName.Substring(0, index);


            if (!viewSchedule.Name.Contains("<"))
            {
                viewSchedule.Export(filePath, $"{projName}_{docName}_{viewName}.csv", exportOptions);
            }
        }
    }
}
