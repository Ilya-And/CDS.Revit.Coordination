using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.Exceptions;
using CDS.Revit.Coordination.Axapta.Models;
using CDS.Revit.Coordination.Services.Axapta;
using CDS.Revit.Coordination.Services.Excel;
using CDS.Revit.Coordination.Services.Revit;
using Microsoft.Win32;

namespace CDS.Revit.Coordination.Axapta.ViewModels
{
    class MainWindowViewModel : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        public void OnPropertyChanged([CallerMemberName] string prop = "")
        {
            if (PropertyChanged != null)
                PropertyChanged(this, new PropertyChangedEventArgs(prop));
        }
        private bool _isAllFilesWork { get; set; } = true;

        public bool IsAllFilesWork
        {
            get => _isAllFilesWork;
            set
            {
                _isAllFilesWork = value;
                OnPropertyChanged("IsAllFilesWork");
            }
        }

        public string UnprocessedFiles { get; set; }

        #region ДАННЫЕ ИЗ ОСНОВНОГО ФАЙЛА EXCEL

        private ColumnValues _projectNameColumn { get; set; }

        public ColumnValues ProjectNameColumn
        {
            get => _projectNameColumn;
            set
            {
                _projectNameColumn = value;
                OnPropertyChanged("ProjectNameColumn");
            }
        }

        private ColumnValues _projectSectionColumn { get; set; }
        public ColumnValues ProjectSectionColumn
        {
            get => _projectSectionColumn;
            set
            {
                _projectSectionColumn = value;
                OnPropertyChanged("ProjectSectionColumn");
            }
        }

        private ColumnValues _pathToSaveColumn;
        public ColumnValues PathToSaveColumn
        {
            get => _pathToSaveColumn;
            set
            {
                _pathToSaveColumn = value;
                OnPropertyChanged("PathToSaveColumn");
            }
        }

        private ColumnValues _pathToOpenColumn;
        public ColumnValues PathToOpenColumn
        {
            get => _pathToOpenColumn;
            set
            {
                _pathToOpenColumn = value;
                OnPropertyChanged("PathToOpenColumn");
            }
        }

        private ColumnValues _pathTableColumn;
        public ColumnValues PathTableColumn
        {
            get => _pathTableColumn;
            set
            {
                _pathTableColumn = value;
                OnPropertyChanged("PathTableColumn");
            }
        }
        private void GetColumnsFromExcel()
        {
            ProjectNameColumn = (from column in GeneralTable
                                 where column.ColumnName == "Проект"
                                 select column).FirstOrDefault();

            ProjectSectionColumn = (from column in GeneralTable
                                    where column.ColumnName == "Раздел"
                                    select column).FirstOrDefault();

            PathToSaveColumn = (from column in GeneralTable
                                where column.ColumnName == "ПутьДляСохранения"
                                select column).FirstOrDefault();

            PathToOpenColumn = (from column in GeneralTable
                                 where column.ColumnName == "Путь"
                                 select column).FirstOrDefault();

            PathTableColumn = (from column in GeneralTable
                               where column.ColumnName == "ПутьТаблицыВыбора"
                               select column).FirstOrDefault();
        }
        

        #endregion

        #region ПУТИ ДО ФАЙЛОВ

        private string _pathToAllFiles;
        public string PathToAllFiles
        {
            get => _pathToAllFiles;
            set
            {
                _pathToAllFiles = value;
                OnPropertyChanged("PathToAllFiles");
            }
        }

        private string[] _pathToCSVFiles;
        public string[] PathToCSVFiles
        {
            get => _pathToCSVFiles;
            set
            {
                _pathToCSVFiles = value;
                OnPropertyChanged("PathToCSVFiles");
            }
        }

        private string _fileNamesCSVFiles;
        public string FileNamesCSVFiles
        {
            get => _fileNamesCSVFiles;
            set
            {
                _fileNamesCSVFiles = value;
                OnPropertyChanged("FileNamesCSVFiles");
            }
        }

        #endregion

        #region НАСТРОЙКИ ВЫГРУЗКИ ОБЪЕМОВ
        private bool _isExportToExcel;
        public bool IsExportToExcel
        {
            get => _isExportToExcel;
            set
            {
                _isExportToExcel = value;
                OnPropertyChanged("IsExportToExcel");
            }
        }

        private bool _isExportToJSON;
        public bool IsExportToJSON
        {
            get => _isExportToJSON;
            set => _isExportToJSON = value;
        }

        private bool _isExportToAxapta = true;
        public bool IsExportToAxapta
        {
            get => _isExportToAxapta;
            set
            {
                _isExportToAxapta = value;
                OnPropertyChanged("IsExportToAxapta");
            }
        }

        private bool _isTestExport;
        public bool IsTestExport
        {
            get => _isTestExport;
            set
            {
                _isTestExport = value;
                OnPropertyChanged("IsTestExport");
            }
        }
        #endregion

        #region НАСТРОЙКИ СОХРАНЕНИЯ МОДЕЛИ

        private bool _isSaveRVT = true;
        public bool IsSaveRVT
        {
            get => _isSaveRVT;
            set => _isSaveRVT = value;
        }

        private bool _isSaveNWC = true;
        public bool IsSaveNWC
        {
            get => _isSaveNWC;
            set => _isSaveNWC = value;
        }

        #endregion

        /*Поле для хранения данных из основной таблицы Excel
         с информацией по файлам, которые нужно обрабатывать*/
        public List<ColumnValues> GeneralTable { get; set; }

        /*Поле для хранения данных для отправки работ в Axapta
         */
        public List<WorkToSend> WorksListToSentValuesToAxapta { get; set; }

        /*Поле для хранения данных для отправки номенклатур в Axapta
         */
        public List<MaterialToSend> MaterialsListToSentValuesToAxapta { get; set; }

        #region КОМАНДЫ

        /*Основная команда, запускающая процесс обработки моделей, 
         подготовки и отправки данных в Axapta,
        а также сохранения данных в различных форматах*/
        private RelayCommand _startPreparationCommand;
        public RelayCommand StartPreparationCommand
        {
            get
            {
                return _startPreparationCommand ?? new RelayCommand(obj =>
                {
                    if(PathToAllFiles != null)
                    {
                        ExcelService excelService = new ExcelService();
                        RevitFileService revitFileService = new RevitFileService(SendValuesCommand.App);

                        bool isContinue = true;

                        try
                        {
                            List<ColumnValues> generalTable = excelService.GetValuesFromExcelTable(PathToAllFiles);
                            GeneralTable = generalTable;
                            GetColumnsFromExcel();

                        }
                        catch (Exception ex)
                        {
                            isContinue = false;
                            MessageBox.Show("Не удалось обработать Excel файл!\nИ вот почему:\n" + ex.Message + "\n" + ex.StackTrace);
                        }

                        if(isContinue == true)
                        {

                            for (int i = 0; i <= GeneralTable[1].RowValues.Count - 1; i++)
                            {
                                Document doc = revitFileService.OpenRVTFile(PathToOpenColumn.RowValues[i].Value);
                                RevitModelElementService revitModelElementService = new RevitModelElementService(doc);
                                var unCorrectLevelsNames = revitModelElementService.GetUnCorrectLevelsNames();
                                if(unCorrectLevelsNames == "")
                                {
                                    var allElementsInModel = new FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements();

                                    var views3DForSetSectionParameter = revitModelElementService.Get3DViewForSetSectionParameter();

                                    var floors = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements();
                                    var roofs = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().ToElements();
                                    var ceilings = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElements();
                                    var walls = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements();
                                    var foundations = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFoundation).WhereElementIsNotElementType().ToElements();

                                    var floorIds = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElementIds();
                                    var roofIds = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().ToElementIds();
                                    var ceilingIds = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Ceilings).WhereElementIsNotElementType().ToElementIds();
                                    var wallIds = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElementIds();
                                    var foundationIds = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StructuralFoundation).WhereElementIsNotElementType().ToElementIds();

                                    using (Transaction tr = new Transaction(doc))
                                    {
                                        tr.Start("Заполнение параметров номера секции и этажа");

                                        foreach (ElementId elementId in views3DForSetSectionParameter)
                                        {
                                            View3D view3D = doc.GetElement(elementId) as View3D;
                                            string sectionNumber = view3D.Name.Split('_')[2];

                                            var allElementsByView = new FilteredElementCollector(doc, elementId).WhereElementIsNotElementType().ToElements();

                                            if(allElementsByView != null)
                                            {
                                                foreach (Element element in allElementsByView)
                                                {
                                                    if (element != null)
                                                    {
                                                        try
                                                        {
                                                            if (element.Category != null)
                                                            {
                                                                revitModelElementService.SetParametersValuesToElement(element, sectionNumber);
                                                            }
                                                        }
                                                        catch (InvalidObjectException)
                                                        {
                                                            continue;
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        tr.Commit();
                                    }

                                    using (Transaction tr = new Transaction(doc))
                                    {
                                        tr.Start("Восстановление формы");

                                        if (floors != null)
                                            revitModelElementService.RestorateForm(floors);
                                        if (roofs != null)
                                            revitModelElementService.RestorateForm(roofs);

                                        tr.Commit();
                                    }

                                    using (Transaction tr = new Transaction(doc))
                                    {
                                        tr.Start("Создание частей");

                                        if(floorIds != null)
                                            revitModelElementService.CreatePartsFromElement(floorIds);

                                        if (roofIds != null)
                                            revitModelElementService.CreatePartsFromElement(roofIds);

                                        if (ceilingIds != null)
                                            revitModelElementService.CreatePartsFromElement(ceilingIds);

                                        if (wallIds != null)
                                            revitModelElementService.CreatePartsFromElement(wallIds);

                                        if (foundationIds != null)
                                            revitModelElementService.CreatePartsFromElement(foundationIds);

                                        doc.Regenerate();

                                        tr.Commit();
                                    }

                                    using (Transaction tr = new Transaction(doc))
                                    {
                                        tr.Start("Заполнение классификатора");

                                        var classTable = excelService.GetValuesFromExcelTable(PathTableColumn.RowValues[i].Value);

                                        ColumnValues category = (from column in classTable
                                                                    where column.ColumnName == "Категория"
                                                                    select column).FirstOrDefault();

                                        ColumnValues parameter1 = (from column in classTable
                                                                    where column.ColumnName == "Параметр_1"
                                                                    select column).FirstOrDefault();

                                        ColumnValues parameter1Condition = (from column in classTable
                                                                            where column.ColumnName == "Условие_1"
                                                                            select column).FirstOrDefault();

                                        ColumnValues parameterValue1 = (from column in classTable
                                                                        where column.ColumnName == "Значение_1"
                                                                        select column).FirstOrDefault();

                                        ColumnValues parameter2 = (from column in classTable
                                                                    where column.ColumnName == "Параметр_2"
                                                                    select column).FirstOrDefault();

                                        ColumnValues parameter2Condition = (from column in classTable
                                                                            where column.ColumnName == "Условие_2"
                                                                            select column).FirstOrDefault();

                                        ColumnValues parameterValue2 = (from column in classTable
                                                                        where column.ColumnName == "Значение_2"
                                                                        select column).FirstOrDefault();

                                        ColumnValues parameter3 = (from column in GeneralTable
                                                                    where column.ColumnName == "Параметр_3"
                                                                    select column).FirstOrDefault();

                                        ColumnValues parameter3Condition = (from column in classTable
                                                                            where column.ColumnName == "Условие_3"
                                                                            select column).FirstOrDefault();

                                        ColumnValues parameterValue3 = (from column in classTable
                                                                        where column.ColumnName == "Значение_3"
                                                                        select column).FirstOrDefault();


                                        ColumnValues parameter4 = (from column in classTable
                                                                    where column.ColumnName == "Параметр_4"
                                                                    select column).FirstOrDefault();

                                        ColumnValues parameter4Condition = (from column in classTable
                                                                            where column.ColumnName == "Условие_4"
                                                                            select column).FirstOrDefault();

                                        ColumnValues parameterValue4 = (from column in classTable
                                                                        where column.ColumnName == "Значение_4"
                                                                        select column).FirstOrDefault();

                                        ColumnValues parameter5 = (from column in classTable
                                                                    where column.ColumnName == "Параметр_5"
                                                                    select column).FirstOrDefault();

                                        ColumnValues parameter5Condition = (from column in classTable
                                                                            where column.ColumnName == "Условие_5"
                                                                            select column).FirstOrDefault();

                                        ColumnValues parameterValue5 = (from column in classTable
                                                                        where column.ColumnName == "Значение_5"
                                                                        select column).FirstOrDefault();

                                        ColumnValues classifierForLine = (from column in classTable
                                                                    where column.ColumnName == "ЦДС_Классификатор(Пр)"
                                                                    select column).FirstOrDefault();

                                        ColumnValues classifierForOther = (from column in classTable
                                                                            where column.ColumnName == "ЦДС_Классификатор(Кр)"
                                                                            select column).FirstOrDefault();

                                        ColumnValues materialClassifier = (from column in classTable
                                                                            where column.ColumnName == "ЦДС_Классификатор материалов"
                                                                            select column).FirstOrDefault();

                                        
                                        ElementId view3D = revitModelElementService.Get3DViewForExportToNWC().Id;

                                        var allElementsByView = new FilteredElementCollector(doc, view3D).WhereElementIsNotElementType().ToElements();
                                        foreach (Element element in allElementsByView)
                                        {
                                            try
                                            {
                                                var categoryNameFromElement = "";
                                                try
                                                {
                                                    categoryNameFromElement = element.Category.Name;
                                                }
                                                catch
                                                {
                                                    continue;
                                                }

                                                if (categoryNameFromElement != "")
                                                {
                                                    for (int n = 0; n <= classTable[0].RowValues.Count - 1; n++)
                                                    {
                                                        var categoryNameFromTable = "";

                                                        var parameter1FromTable = "";
                                                        var parameter1ConditionFromTable = "";
                                                        var parameterValue1FromTable = "";

                                                        var parameter2FromTable = "";
                                                        var parameter2ConditionFromTable = "";
                                                        var parameterValue2FromTable = "";

                                                        var parameter3FromTable = "";
                                                        var parameter3ConditionFromTable = "";
                                                        var parameterValue3FromTable = "";

                                                        var parameter4FromTable = "";
                                                        var parameter4ConditionFromTable = "";
                                                        var parameterValue4FromTable = "";

                                                        var parameter5FromTable = "";
                                                        var parameter5ConditionFromTable = "";
                                                        var parameterValue5FromTable = "";

                                                        var classifierLineFromTable = "";
                                                        var classifierOtherFromTable = "";

                                                        var materialClassifierFromTable = "";

                                                        categoryNameFromTable = category.RowValues[n].Value;

                                                        if (parameter1 != null)
                                                        {
                                                            parameter1FromTable = parameter1.RowValues[n].Value;
                                                            parameter1ConditionFromTable = parameter1Condition.RowValues[n].Value;
                                                            parameterValue1FromTable = parameterValue1.RowValues[n].Value;
                                                        }

                                                        if (parameter2 != null)
                                                        {
                                                            parameter2FromTable = parameter2.RowValues[n].Value;
                                                            parameter2ConditionFromTable = parameter2Condition.RowValues[n].Value;
                                                            parameterValue2FromTable = parameterValue2.RowValues[n].Value;
                                                        }

                                                        if (parameter3 != null)
                                                        {
                                                            parameter3FromTable = parameter3.RowValues[n].Value;
                                                            parameter3ConditionFromTable = parameter3Condition.RowValues[n].Value;
                                                            parameterValue3FromTable = parameterValue3.RowValues[n].Value;
                                                        }

                                                        if (parameter4 != null)
                                                        {
                                                            parameter4FromTable = parameter4.RowValues[n].Value;
                                                            parameter4ConditionFromTable = parameter4Condition.RowValues[n].Value;
                                                            parameterValue4FromTable = parameterValue4.RowValues[n].Value;
                                                        }

                                                        if (parameter5 != null)
                                                        {
                                                            parameter5FromTable = parameter5.RowValues[n].Value;
                                                            parameter5ConditionFromTable = parameter5Condition.RowValues[n].Value;
                                                            parameterValue5FromTable = parameterValue5.RowValues[n].Value;
                                                        }

                                                        classifierLineFromTable = classifierForLine.RowValues[n].Value;
                                                        classifierOtherFromTable = classifierForOther.RowValues[n].Value;

                                                        materialClassifierFromTable = materialClassifier.RowValues[n].Value;

                                                        if(categoryNameFromElement == categoryNameFromTable)
                                                        {
                                                            var parameterValue1FromElement = "";
                                                            var parameterValue2FromElement = "";
                                                            var parameterValue3FromElement = "";
                                                            var parameterValue4FromElement = "";
                                                            var parameterValue5FromElement = "";

                                                            bool isFirstParameterValueMatch = false;
                                                            bool isSecondParameterValueMatch = false;
                                                            bool isThirdParameterValueMatch = false;
                                                            bool isFouthParameterValueMatch = false;
                                                            bool isFifthParameterValueMatch = false;

                                                            string classifierForSet = "";

                                                            switch (categoryNameFromElement)
                                                            {
                                                                case "Части":
                                                                    Part asPart = element as Part;
                                                                    var categoryNameHost = asPart.get_Parameter(BuiltInParameter.DPART_ORIGINAL_CATEGORY).AsString();

                                                                    parameterValue1FromElement = asPart.LookupParameter(parameter1FromTable)?.AsString();
                                                                    parameterValue2FromElement = asPart.LookupParameter(parameter2FromTable)?.AsString();
                                                                    parameterValue3FromElement = asPart.LookupParameter(parameter3FromTable)?.AsString();
                                                                    parameterValue4FromElement = asPart.LookupParameter(parameter4FromTable)?.AsString();
                                                                    parameterValue5FromElement = asPart.LookupParameter(parameter5FromTable)?.AsString();


                                                                    isFirstParameterValueMatch = IsParameterValueMatch(parameter1ConditionFromTable, parameterValue1FromTable, parameterValue1FromElement);
                                                                    isSecondParameterValueMatch = IsParameterValueMatch(parameter2ConditionFromTable, parameterValue2FromTable, parameterValue2FromElement);
                                                                    isThirdParameterValueMatch = IsParameterValueMatch(parameter3ConditionFromTable, parameterValue3FromTable, parameterValue3FromElement);
                                                                    isFouthParameterValueMatch = IsParameterValueMatch(parameter4ConditionFromTable, parameterValue4FromTable, parameterValue4FromElement);
                                                                    isFifthParameterValueMatch = IsParameterValueMatch(parameter5ConditionFromTable, parameterValue5FromTable, parameterValue5FromElement);

                                                                    if (isFirstParameterValueMatch == true &&
                                                                        isSecondParameterValueMatch == true &&
                                                                        isThirdParameterValueMatch == true &&
                                                                        isFouthParameterValueMatch == true &&
                                                                        isFifthParameterValueMatch == true)
                                                                    {
                                                                        if (categoryNameHost == "Стены")
                                                                        {
                                                                            var hostElementId = asPart.GetSourceElementIds().ToList()[0].HostElementId;
                                                                            var hostElement = doc.GetElement(hostElementId) as Wall;
                                                                            var hostElementCurve = ((LocationCurve)hostElement.Location).Curve;
                                                                            var asLine = hostElementCurve as Line;
                                                                            if (asLine != null)
                                                                            {
                                                                                classifierForSet = classifierLineFromTable;
                                                                            }
                                                                            else
                                                                            {
                                                                                classifierForSet = classifierOtherFromTable;
                                                                            }
                                                                        }
                                                                        else if (categoryNameHost == "Выступающие профили")
                                                                        {
                                                                            var hostElementId = asPart.GetSourceElementIds().ToList()[0].HostElementId;
                                                                            var hostWallSweepElement = doc.GetElement(hostElementId) as WallSweep;

                                                                            if(hostWallSweepElement != null)
                                                                            {
                                                                                var hostElement = doc.GetElement(hostWallSweepElement.GetHostIds()[0]) as Wall;
                                                                                if(hostElement != null)
                                                                                {
                                                                                    var hostElementCurve = ((LocationCurve)hostElement.Location).Curve;
                                                                                    var asLine = hostElementCurve as Line;
                                                                                    if (asLine != null)
                                                                                    {
                                                                                        classifierForSet = classifierLineFromTable;
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        classifierForSet = classifierOtherFromTable;
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                        else
                                                                        {
                                                                            classifierForSet = classifierLineFromTable;
                                                                        }
                                                                        asPart.LookupParameter("ЦДС_Классификатор")?.Set(classifierForSet);

                                                                        if (materialClassifierFromTable != "" || materialClassifierFromTable != "по материалу")
                                                                        {
                                                                            asPart.LookupParameter("ЦДС_Классификатор материалов")?.Set(materialClassifierFromTable);
                                                                        }
                                                                    }

                                                                    break;

                                                                case "Воздуховоды":
                                                                    Duct asDuct = element as Duct;

                                                                    parameterValue1FromElement = asDuct.LookupParameter(parameter1FromTable)?.AsString();
                                                                    parameterValue2FromElement = asDuct.LookupParameter(parameter2FromTable)?.AsString();
                                                                    parameterValue3FromElement = asDuct.LookupParameter(parameter3FromTable)?.AsString();
                                                                    parameterValue4FromElement = asDuct.LookupParameter(parameter4FromTable)?.AsString();
                                                                    parameterValue5FromElement = asDuct.LookupParameter(parameter5FromTable)?.AsString();


                                                                    isFirstParameterValueMatch = IsParameterValueMatch(parameter1ConditionFromTable, parameterValue1FromTable, parameterValue1FromElement);
                                                                    isSecondParameterValueMatch = IsParameterValueMatch(parameter2ConditionFromTable, parameterValue2FromTable, parameterValue2FromElement);
                                                                    isThirdParameterValueMatch = IsParameterValueMatch(parameter3ConditionFromTable, parameterValue3FromTable, parameterValue3FromElement);
                                                                    isFouthParameterValueMatch = IsParameterValueMatch(parameter4ConditionFromTable, parameterValue4FromTable, parameterValue4FromElement);
                                                                    isFifthParameterValueMatch = IsParameterValueMatch(parameter5ConditionFromTable, parameterValue5FromTable, parameterValue5FromElement);

                                                                    if (isFirstParameterValueMatch == true &&
                                                                        isSecondParameterValueMatch == true &&
                                                                        isThirdParameterValueMatch == true &&
                                                                        isFouthParameterValueMatch == true &&
                                                                        isFifthParameterValueMatch == true)
                                                                    {
                                                                        classifierForSet = classifierLineFromTable;

                                                                        try
                                                                        {
                                                                            asDuct.LookupParameter("ЦДС_Классификатор")?.Set(classifierForSet);

                                                                            if (materialClassifierFromTable != "" || materialClassifierFromTable != "по материалу")
                                                                            {
                                                                                asDuct.LookupParameter("ЦДС_Классификатор материалов")?.Set(materialClassifierFromTable);
                                                                            }
                                                                        }
                                                                        catch
                                                                        {
                                                                            break;
                                                                        }
                                                                    }

                                                                    break;

                                                                case "Трубы":
                                                                    Pipe asPipe = element as Pipe;

                                                                    parameterValue1FromElement = asPipe.LookupParameter(parameter1FromTable)?.AsString();
                                                                    parameterValue2FromElement = asPipe.LookupParameter(parameter2FromTable)?.AsString();
                                                                    parameterValue3FromElement = asPipe.LookupParameter(parameter3FromTable)?.AsString();
                                                                    parameterValue4FromElement = asPipe.LookupParameter(parameter4FromTable)?.AsString();
                                                                    parameterValue5FromElement = asPipe.LookupParameter(parameter5FromTable)?.AsString();

                                                                    isFirstParameterValueMatch = IsParameterValueMatch(parameter1ConditionFromTable, parameterValue1FromTable, parameterValue1FromElement);
                                                                    isSecondParameterValueMatch = IsParameterValueMatch(parameter2ConditionFromTable, parameterValue2FromTable, parameterValue2FromElement);
                                                                    isThirdParameterValueMatch = IsParameterValueMatch(parameter3ConditionFromTable, parameterValue3FromTable, parameterValue3FromElement);
                                                                    isFouthParameterValueMatch = IsParameterValueMatch(parameter4ConditionFromTable, parameterValue4FromTable, parameterValue4FromElement);
                                                                    isFifthParameterValueMatch = IsParameterValueMatch(parameter5ConditionFromTable, parameterValue5FromTable, parameterValue5FromElement);

                                                                    if (isFirstParameterValueMatch == true &&
                                                                        isSecondParameterValueMatch == true &&
                                                                        isThirdParameterValueMatch == true &&
                                                                        isFouthParameterValueMatch == true &&
                                                                        isFifthParameterValueMatch == true)
                                                                    {
                                                                        classifierForSet = classifierLineFromTable;

                                                                        try
                                                                        {
                                                                            asPipe.LookupParameter("ЦДС_Классификатор")?.Set(classifierForSet);

                                                                            if (materialClassifierFromTable != "" || materialClassifierFromTable != "по материалу")
                                                                            {
                                                                                asPipe.LookupParameter("ЦДС_Классификатор материалов")?.Set(materialClassifierFromTable);
                                                                            }
                                                                        }
                                                                        catch
                                                                        {
                                                                            break;
                                                                        }
                                                                    }

                                                                    break;

                                                                case "Материалы изоляции воздуховодов":
                                                                    DuctInsulation asDuctInsulation = element as DuctInsulation;

                                                                    parameterValue1FromElement = asDuctInsulation.LookupParameter(parameter1FromTable)?.AsString();
                                                                    parameterValue2FromElement = asDuctInsulation.LookupParameter(parameter2FromTable)?.AsString();
                                                                    parameterValue3FromElement = asDuctInsulation.LookupParameter(parameter3FromTable)?.AsString();
                                                                    parameterValue4FromElement = asDuctInsulation.LookupParameter(parameter4FromTable)?.AsString();
                                                                    parameterValue5FromElement = asDuctInsulation.LookupParameter(parameter5FromTable)?.AsString();

                                                                    isFirstParameterValueMatch = IsParameterValueMatch(parameter1ConditionFromTable, parameterValue1FromTable, parameterValue1FromElement);
                                                                    isSecondParameterValueMatch = IsParameterValueMatch(parameter2ConditionFromTable, parameterValue2FromTable, parameterValue2FromElement);
                                                                    isThirdParameterValueMatch = IsParameterValueMatch(parameter3ConditionFromTable, parameterValue3FromTable, parameterValue3FromElement);
                                                                    isFouthParameterValueMatch = IsParameterValueMatch(parameter4ConditionFromTable, parameterValue4FromTable, parameterValue4FromElement);
                                                                    isFifthParameterValueMatch = IsParameterValueMatch(parameter5ConditionFromTable, parameterValue5FromTable, parameterValue5FromElement);

                                                                    if (isFirstParameterValueMatch == true &&
                                                                        isSecondParameterValueMatch == true &&
                                                                        isThirdParameterValueMatch == true &&
                                                                        isFouthParameterValueMatch == true &&
                                                                        isFifthParameterValueMatch == true)
                                                                    {
                                                                        classifierForSet = classifierLineFromTable;

                                                                        try
                                                                        {
                                                                            asDuctInsulation.LookupParameter("ЦДС_Классификатор")?.Set(classifierForSet);

                                                                            if (materialClassifierFromTable != "" || materialClassifierFromTable != "по материалу")
                                                                            {
                                                                                asDuctInsulation.LookupParameter("ЦДС_Классификатор материалов")?.Set(materialClassifierFromTable);
                                                                            }
                                                                        }
                                                                        catch
                                                                        {
                                                                            break;
                                                                        }
                                                                    }

                                                                    break;

                                                                case "Материалы изоляции труб":
                                                                    PipeInsulation asPipeInsulation = element as PipeInsulation;

                                                                    parameterValue1FromElement = asPipeInsulation.LookupParameter(parameter1FromTable)?.AsString();
                                                                    parameterValue2FromElement = asPipeInsulation.LookupParameter(parameter2FromTable)?.AsString();
                                                                    parameterValue3FromElement = asPipeInsulation.LookupParameter(parameter3FromTable)?.AsString();
                                                                    parameterValue4FromElement = asPipeInsulation.LookupParameter(parameter4FromTable)?.AsString();
                                                                    parameterValue5FromElement = asPipeInsulation.LookupParameter(parameter5FromTable)?.AsString();

                                                                    isFirstParameterValueMatch = IsParameterValueMatch(parameter1ConditionFromTable, parameterValue1FromTable, parameterValue1FromElement);
                                                                    isSecondParameterValueMatch = IsParameterValueMatch(parameter2ConditionFromTable, parameterValue2FromTable, parameterValue2FromElement);
                                                                    isThirdParameterValueMatch = IsParameterValueMatch(parameter3ConditionFromTable, parameterValue3FromTable, parameterValue3FromElement);
                                                                    isFouthParameterValueMatch = IsParameterValueMatch(parameter4ConditionFromTable, parameterValue4FromTable, parameterValue4FromElement);
                                                                    isFifthParameterValueMatch = IsParameterValueMatch(parameter5ConditionFromTable, parameterValue5FromTable, parameterValue5FromElement);

                                                                    if (isFirstParameterValueMatch == true &&
                                                                        isSecondParameterValueMatch == true &&
                                                                        isThirdParameterValueMatch == true &&
                                                                        isFouthParameterValueMatch == true &&
                                                                        isFifthParameterValueMatch == true)
                                                                    {
                                                                        classifierForSet = classifierLineFromTable;

                                                                        try
                                                                        {
                                                                            asPipeInsulation.LookupParameter("ЦДС_Классификатор")?.Set(classifierForSet);

                                                                            if (materialClassifierFromTable != "" || materialClassifierFromTable != "по материалу")
                                                                            {
                                                                                asPipeInsulation.LookupParameter("ЦДС_Классификатор материалов")?.Set(materialClassifierFromTable);
                                                                            }
                                                                        }
                                                                        catch
                                                                        {
                                                                            break;
                                                                        }
                                                                    }
                                                                    break;

                                                                default:
                                                                    try
                                                                    {
                                                                        FamilyInstance asFamilyInstance = element as FamilyInstance;

                                                                        if (asFamilyInstance != null)
                                                                        {
                                                                            if(!asFamilyInstance.Name.Contains("Свая") && !asFamilyInstance.Name.Contains("свая") && !asFamilyInstance.Name.Contains("ЛМ"))
                                                                            {
                                                                                parameterValue1FromElement = asFamilyInstance.LookupParameter(parameter1FromTable)?.AsString();
                                                                                parameterValue2FromElement = asFamilyInstance.LookupParameter(parameter2FromTable)?.AsString();
                                                                                parameterValue3FromElement = asFamilyInstance.LookupParameter(parameter3FromTable)?.AsString();
                                                                                parameterValue4FromElement = asFamilyInstance.LookupParameter(parameter4FromTable)?.AsString();
                                                                                parameterValue5FromElement = asFamilyInstance.LookupParameter(parameter5FromTable)?.AsString();

                                                                                isFirstParameterValueMatch = IsParameterValueMatch(parameter1ConditionFromTable, parameterValue1FromTable, parameterValue1FromElement);
                                                                                isSecondParameterValueMatch = IsParameterValueMatch(parameter2ConditionFromTable, parameterValue2FromTable, parameterValue2FromElement);
                                                                                isThirdParameterValueMatch = IsParameterValueMatch(parameter3ConditionFromTable, parameterValue3FromTable, parameterValue3FromElement);
                                                                                isFouthParameterValueMatch = IsParameterValueMatch(parameter4ConditionFromTable, parameterValue4FromTable, parameterValue4FromElement);
                                                                                isFifthParameterValueMatch = IsParameterValueMatch(parameter5ConditionFromTable, parameterValue5FromTable, parameterValue5FromElement);

                                                                                if (isFirstParameterValueMatch == true &&
                                                                                    isSecondParameterValueMatch == true &&
                                                                                    isThirdParameterValueMatch == true &&
                                                                                    isFouthParameterValueMatch == true &&
                                                                                    isFifthParameterValueMatch == true)
                                                                                {
                                                                                    classifierForSet = classifierLineFromTable;
                                                                                    asFamilyInstance.LookupParameter("ЦДС_Классификатор")?.Set(classifierForSet);

                                                                                    if (materialClassifierFromTable != "" || materialClassifierFromTable != "по материалу")
                                                                                    {
                                                                                        asFamilyInstance.LookupParameter("ЦДС_Классификатор материалов")?.Set(materialClassifierFromTable);
                                                                                    }

                                                                                }
                                                                            }
                                                                        }
                                                                        break;
                                                                    }
                                                                    catch
                                                                    {
                                                                        break;
                                                                    }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                continue;
                                                //MessageBox.Show(ex.Message + "\n" + ex.StackTrace + "\n" + element);
                                            }
                                        }

                                        tr.Commit();
                                    }

                                    //using (Transaction tr = new Transaction(doc))
                                    //{
                                    //    tr.Start("");

                                    //    tr.Commit();
                                    //}

                                    //using (Transaction tr = new Transaction(doc))
                                    //{
                                    //    tr.Start("");

                                    //    tr.Commit();
                                    //}

                                    if(IsSaveNWC == true)
                                    {
                                        revitFileService.ExportToNWC(PathToSaveColumn.RowValues[i].Value, doc);
                                    }

                                    if(IsSaveRVT == true)
                                    {
                                        revitFileService.SaveAndCloseRVTFile(PathToSaveColumn.RowValues[i].Value, doc);
                                    }

                                }
                                else
                                {
                                    IsAllFilesWork = false;
                                    continue;
                                }
                            }
                        }
                       
                    }
                    else
                    {
                        MessageBox.Show("Укажите путь до файла Excel!!!");
                    }

                    //int allElementsByAllViews = 0;
                    //foreach (ElementId viewId in selectedViews)
                    //{
                    //    allElementsByAllViews += new FilteredElementCollector(doc, viewId).WhereElementIsNotElementType().ToElements().Count;
                    //}
                    //using (Transaction tr = new Transaction(doc))
                    //{
                    //    tr.Start("SetLevelAndSection");
                    //    foreach (ElementId view3DId in selectedViews)
                    //    {
                    //        var view3D = _doc.GetElement(view3DId) as View3D;
                    //        string sectionNumber = view3D.Name;
                    //        progressBarView.NameOfView.Text = $"Обработка вида - \"{sectionNumber}\"";


                    //        var allElementsByView = new FilteredElementCollector(doc, view3DId).WhereElementIsNotElementType().ToElements();
                    //        progressBarView.ProgressBarStatusView.Maximum = allElementsByAllViews;
                    //        progressBarView.ProgressBarStatusView.Value = 1;

                    //        List<Element> allElements = new List<Element>();

                    //        foreach (Element element in allElementsByView)
                    //        {
                    //            if (element != null)
                    //            {
                    //                try
                    //                {
                    //                    if (element.Category != null)
                    //                    {
                    //                        SetParametersValuesToElement(element, levelsService, sectionNumber);
                    //                    }
                    //                }
                    //                catch (InvalidObjectException)
                    //                {
                    //                    continue;
                    //                }
                    //            }
                    //            progressBarView.CountOfElement.Text = $"Обработано {_countElements} из {allElementsByAllViews} элементов.";
                    //            progressBarView.ProgressBarStatusView.Dispatcher.Invoke(new ProgressBarDelegate(progressBarView.UpdateProgress), System.Windows.Threading.DispatcherPriority.Background);
                    //        }
                    //    }
                    //    tr.Commit();
                    //}

                    //progressBarView.Close();
                    //if (_countElements != 0)
                    //{
                    //    MessageBox.Show(resultMessage);
                    //}
                    //else
                    //{
                    //    MessageBox.Show($"Обработано элементов - {_countElements}");
                    //}
                }
                );
            }
        }

        private RelayCommand _startSendCommand;
        public RelayCommand StartSendCommand
        {
            get
            {
                return _startSendCommand ?? new RelayCommand(obj =>
                {
                    CSVService csvService = new CSVService();
                    ExcelService excelService = new ExcelService();
                    AxaptaService axaptaService = new AxaptaService();

                    bool isContinue = true;
                    Dictionary<string, List<AxaptaWorkset>> worksFromAxapta = new Dictionary<string, List<AxaptaWorkset>>();

                    try
                    {
                        worksFromAxapta = axaptaService.GetWorksFromAxapta();
                    }
                    catch
                    {
                        isContinue = false;
                        MessageBox.Show("Невозможно продолжить! Нет связи с Axapta.");
                    }

                    List<ColumnValues> tableWithValues = new List<ColumnValues>();

                    if(isContinue == true)
                    {
                        foreach (string filePath in PathToCSVFiles)
                        {
                            if (filePath.Contains(".csv"))
                            {
                                tableWithValues = csvService.GetValuesFromCSVTable(filePath);
                            }
                            if (filePath.Contains(".xlsx"))
                            {
                                tableWithValues = excelService.GetValuesFromExcelTable(filePath);
                            }

                            var classifierColumn = (from column in tableWithValues
                                                    where column.ColumnName == "ЦДС_Классификатор"
                                                    select column).FirstOrDefault();
                            var materialColumn = (from column in tableWithValues
                                                  where column.ColumnName == "ЦДС_Классификатор материалов"
                                                  select column).FirstOrDefault();
                            var sectionNumberColumn = (from column in tableWithValues
                                                       where column.ColumnName == "Номер секции"
                                                       select column).FirstOrDefault();
                            var levelNumberColumn = (from column in tableWithValues
                                                     where column.ColumnName == "Этаж"
                                                     select column).FirstOrDefault();


                            for (int i = 0; i <= tableWithValues[0].RowValues.Count - 1; i++)
                            {

                            }
                        }
                    }
                }
                );
            }
        }

        private static bool IsParameterValueMatch(string parameter1ConditionFromTable, string parameterValue1FromTable, string parameterValue1FromElement)
        {
            bool isParameterValueMatch = false;

            if (parameter1ConditionFromTable == "Равно")
            {
                isParameterValueMatch = parameterValue1FromElement == parameterValue1FromTable;
            }
            else if (parameter1ConditionFromTable == "НеРавно")
            {
                isParameterValueMatch = parameterValue1FromElement == parameterValue1FromTable;
            }
            else if (parameter1ConditionFromTable == "Содержит")
            {
                isParameterValueMatch = parameterValue1FromElement.Contains(parameterValue1FromTable);
            }
            else if (parameter1ConditionFromTable == "НеСодержит")
            {
                isParameterValueMatch = !parameterValue1FromElement.Contains(parameterValue1FromTable);
            }
            else
            {
                isParameterValueMatch = true;
            }

            return isParameterValueMatch;
        }

        private RelayCommand _getAllFilesCommand;
        public RelayCommand GetAllFilesCommand
        {
            get
            {
                return _getAllFilesCommand ?? new RelayCommand(obj =>
                {
                    GetGeneralExcelTableMethod();
                }
                );
            }
        }

        private RelayCommand _getCSVFilesCommand;
        public RelayCommand GetCSVFilesCommand
        {
            get
            {
                return _getCSVFilesCommand ?? new RelayCommand(obj =>
                {
                    GetCSVTableMethod();
                }
                );
            }
        }

        #endregion

        private void GetCSVTableMethod()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "CSV (*.csv)|*.csv|Excel (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
            openFileDialog.FilterIndex = 2;
            openFileDialog.ShowDialog();

            PathToCSVFiles = openFileDialog.FileNames;

            foreach(string filePath in PathToCSVFiles)
            {
                var splitedFileName = filePath.Split('\\');
                string fileName = splitedFileName[splitedFileName.Length - 1];

                if (fileName.Contains(".csv") || fileName.Contains(".xlsx"))
                {
                    if(FileNamesCSVFiles == "Выбран неверный формат!")
                    {
                        FileNamesCSVFiles = "";
                    }

                    FileNamesCSVFiles = FileNamesCSVFiles + fileName + "; ";
                }
                else
                {
                    FileNamesCSVFiles = "Выбран неверный формат!";
                    break;
                }
            }
        }

        private void GetGeneralExcelTableMethod()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Excel (*.xlsx)|*.xlsx";
            openFileDialog.FilterIndex = 2;

            openFileDialog.ShowDialog();
            PathToAllFiles = openFileDialog.FileName;
        }
    }
}
