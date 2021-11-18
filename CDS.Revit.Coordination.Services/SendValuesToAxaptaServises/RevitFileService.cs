using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDS.Revit.Coordination.Services
{
    public class RevitFileService
    {
        private Application _applicationDoc;
        public RevitFileService(Application application)
        {
            _applicationDoc = application;
        }
        public Document OpenRVTFile(string filePath)
        {
            // Переводим путь к файлу из string в ModelPath.

            ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);

            // Создаем настройки открытия файла. Открываем хранилище с отсоединением.

            OpenOptions openOptions = new OpenOptions();
            openOptions.DetachFromCentralOption = DetachFromCentralOption.DetachAndDiscardWorksets;

            Document doc = _applicationDoc.OpenDocumentFile(modelPath, openOptions);

            return doc;
        }
        public void SaveRVTFile(string filePath, Document doc)
        {
            // Переводим путь к файлу из string в ModelPath.

            ModelPath modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(filePath);

            // Создаем настройки сохранения файла. Сохраняем файл без сохранения рабочих наборов.

            WorksharingSaveAsOptions worksharingSaveAsOptions = new WorksharingSaveAsOptions();
            worksharingSaveAsOptions.SaveAsCentral = false;

            SaveAsOptions saveAsOptions = new SaveAsOptions();
            saveAsOptions.SetWorksharingOptions(worksharingSaveAsOptions);

            doc.SaveAs(modelPath, saveAsOptions);
        }
        public void ExportToNWC(string filePath, Document doc)
        {
            // Получаем 3D вид, который будем экспортировать в NWC.
            var all3DViews = new FilteredElementCollector(doc).OfClass(typeof(View3D)).WhereElementIsNotElementType().ToElements();
            List<View3D> view3Ds = all3DViews.Select(view => view as View3D).ToList();
            View3D view3DForExport = view3Ds.Single(view => view.Name.Contains("Axapta_Navisworks"));

            // Получаем список элементов для экспорта.

            var allElementIdsBy3DView = new FilteredElementCollector(doc, view3DForExport.Id).WhereElementIsNotElementType().ToElementIds();
            var allElementIdsForExport = allElementIdsBy3DView.Where(element => element != null).ToList();

            // Создаем настройки экспорта в Navisworks. Экспортируем документ в формат NWC.

            NavisworksExportOptions navisworksExportOptions = new NavisworksExportOptions();
            navisworksExportOptions.SetSelectedElementIds(allElementIdsForExport);
            doc.Export(filePath, doc.PathName, navisworksExportOptions);
        }
    }
}
