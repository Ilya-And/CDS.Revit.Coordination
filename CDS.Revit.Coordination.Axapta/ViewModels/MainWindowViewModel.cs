﻿using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
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
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using Newtonsoft.Json;
using Application = Microsoft.Office.Interop.Excel.Application;

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

        /*Поле для хранения данных из основной таблицы Excel
         с информацией по файлам, которые нужно обрабатывать*/
        public List<ColumnValues> GeneralTable { get; set; }

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

        private string _pathToUnitsFile;
        public string PathToUnitsFile
        {
            get => _pathToUnitsFile;
            set
            {
                _pathToUnitsFile = value;
                OnPropertyChanged("PathToUnitsFile");
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

        #region ДАННЫЕ ДЛЯ ОТПРАВКИ

        /*Поле для хранения данных для отправки работ в Axapta
         */
        public ObservableCollection<WorkToSend> WorksListToSentValuesToAxapta { get; private set; } = new ObservableCollection<WorkToSend>();

        /*Поле для хранения данных для отправки номенклатур в Axapta
         */
        public ObservableCollection<MaterialToSend> MaterialsListToSentValuesToAxapta { get; private set; } = new ObservableCollection<MaterialToSend>();

        #endregion

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
                                                                            where column.ColumnName == "ЦДС_Классификатор_Материалов"
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
                                                                            var asLine = hostElementCurve as Autodesk.Revit.DB.Line;
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
                                                                                    var asLine = hostElementCurve as Autodesk.Revit.DB.Line;
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
                                                                            asPart.LookupParameter("ЦДС_Классификатор_Материалов")?.Set(materialClassifierFromTable);
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
                                                                                asDuct.LookupParameter("ЦДС_Классификатор_Материалов")?.Set(materialClassifierFromTable);
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
                                                                                asPipe.LookupParameter("ЦДС_Классификатор_Материалов")?.Set(materialClassifierFromTable);
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
                                                                                asDuctInsulation.LookupParameter("ЦДС_Классификатор_Материалов")?.Set(materialClassifierFromTable);
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
                                                                                asPipeInsulation.LookupParameter("ЦДС_Классификатор_Материалов")?.Set(materialClassifierFromTable);
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
                                                                                        asFamilyInstance.LookupParameter("ЦДС_Классификатор_Материалов")?.Set(materialClassifierFromTable);
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
                                    else
                                    {
                                        doc.Close(false);
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
                catch (Exception ex)
                {
                    isContinue = false;
                    MessageBox.Show(ex.StackTrace);
                    MessageBox.Show("Невозможно продолжить! Нет связи с Axapta.");
                }

                List<ColumnValues> tableWithValues = new List<ColumnValues>();

                if (isContinue == true)
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
                                              where column.ColumnName == "ЦДС_Классификатор_Материалов"
                                              select column).FirstOrDefault();
                        var sectionNumberColumn = (from column in tableWithValues
                                                   where column.ColumnName == "ADSK_Номер секции"
                                                   select column).FirstOrDefault();
                        var levelNumberColumn = (from column in tableWithValues
                                                 where column.ColumnName == "ADSK_Этаж"
                                                 select column).FirstOrDefault();

                        int lastElementNumberFromFilePath = filePath.Split('\\').ToList().Count - 1;
                        string nameFromFilePath = filePath.Split('\\')[lastElementNumberFromFilePath];
                        string projName = nameFromFilePath.Split('_')[0];

                        var unitsMaterialTable = excelService.GetValuesFromExcelTable(PathToUnitsFile);
                        var dictionaryUnits = new Dictionary<string, string>();

                        for (int n = 0; n <= unitsMaterialTable[0].RowValues.Count - 1; n++)
                        {
                            if (n == 0)
                            {
                                continue;
                            }
                            else
                            {
                                dictionaryUnits[unitsMaterialTable[0].RowValues[n].Value] = unitsMaterialTable[1].RowValues[n].Value;
                            }
                        }

                        for (int i = 0; i <= tableWithValues[0].RowValues.Count - 1; i++)
                        {
                            string classifier = classifierColumn.RowValues[i].Value;

                            List<AxaptaWorkset> works = new List<AxaptaWorkset>();

                            try
                            {
                                works = worksFromAxapta[classifier];
                            }
                            catch
                            {
                                continue;
                            }

                            if (works.Count > 0)
                            {
                                foreach (AxaptaWorkset axaptaWorkset in works)
                                {
                                    if (WorksListToSentValuesToAxapta.Count > 0)
                                    {
                                            int indexEqualElement = GetEqualWorkElement(WorksListToSentValuesToAxapta, projName, sectionNumberColumn.RowValues[i].Value, levelNumberColumn.RowValues[i].Value, axaptaWorkset.ProjWorkCodeId);

                                            if(indexEqualElement != -1)
                                            {
                                                var workToSend = WorksListToSentValuesToAxapta[indexEqualElement];
                                                workToSend.Volume += Double.Parse((from column in tableWithValues
                                                                                   where column.ColumnName == workToSend.Units
                                                                                   select column).FirstOrDefault().RowValues[i].Value);
                                            }

                                            else
                                            {
                                                WorkToSend newWorkToSend = new WorkToSend()
                                                {
                                                    ProjName = projName,
                                                    SectionName = sectionNumberColumn.RowValues[i].Value,
                                                    FloorName = levelNumberColumn.RowValues[i].Value,
                                                    ProjWorkName = axaptaWorkset.Name,
                                                    ProjWorkCodeId = axaptaWorkset.ProjWorkCodeId,
                                                    Volume = Double.Parse((from column in tableWithValues
                                                                           where column.ColumnName == axaptaWorkset.UnitId
                                                                           select column).FirstOrDefault().RowValues[i].Value),
                                                    Units = axaptaWorkset.UnitId
                                                };

                                                WorksListToSentValuesToAxapta.Add(newWorkToSend);
                                            }
                                        }
                                    else
                                    {
                                        WorkToSend newWorkToSend = new WorkToSend()
                                        {
                                            ProjName = projName,
                                            SectionName = sectionNumberColumn.RowValues[i].Value,
                                            FloorName = levelNumberColumn.RowValues[i].Value,
                                            ProjWorkName = axaptaWorkset.Name,
                                            ProjWorkCodeId = axaptaWorkset.ProjWorkCodeId,
                                            Volume = Double.Parse((from column in tableWithValues
                                                                   where column.ColumnName == axaptaWorkset.UnitId
                                                                   select column).FirstOrDefault().RowValues[i].Value),
                                            Units = axaptaWorkset.UnitId
                                        };

                                        WorksListToSentValuesToAxapta.Add(newWorkToSend);
                                    }

                                    string materialClassifierValue = "";

                                    if (materialColumn.RowValues[i].Value.Contains("//"))
                                    {
                                        materialClassifierValue = materialColumn.RowValues[i].Value.Split('/')[2];
                                    }
                                    if (materialColumn.RowValues[i].Value.Contains('<'))
                                    {
                                        materialClassifierValue = materialColumn.RowValues[i].Value.Split('<')[0];
                                    }

                                    if (MaterialsListToSentValuesToAxapta.Count > 0)
                                    {
                                        int indexEqualElement = GetEqualMaterialElement(MaterialsListToSentValuesToAxapta, projName, sectionNumberColumn.RowValues[i].Value, levelNumberColumn.RowValues[i].Value, axaptaWorkset.ProjWorkCodeId, materialClassifierValue);

                                        if (indexEqualElement != -1)
                                        {
                                            var materialToSend = WorksListToSentValuesToAxapta[indexEqualElement];
                                                materialToSend.Volume += Double.Parse((from column in tableWithValues
                                                                                where column.ColumnName == materialToSend.Units
                                                                                select column).FirstOrDefault().RowValues[i].Value);
                                        }
                                        else
                                        {
                                            string units = dictionaryUnits[materialClassifierValue];

                                            MaterialToSend newMaterialToSend = new MaterialToSend()
                                            {
                                                ProjName = projName,
                                                SectionName = sectionNumberColumn.RowValues[i].Value,
                                                FloorName = levelNumberColumn.RowValues[i].Value,
                                                ProjWorkName = axaptaWorkset.Name,
                                                ProjWorkCodeId = axaptaWorkset.ProjWorkCodeId,
                                                ProjMaterialCodeId = materialClassifierValue,
                                                Volume = Double.Parse((from column in tableWithValues
                                                                        where column.ColumnName == units
                                                                        select column).FirstOrDefault().RowValues[i].Value),
                                                Units = units
                                            };

                                            MaterialsListToSentValuesToAxapta.Add(newMaterialToSend);
                                        }
                                           
                                    }
                                    else
                                    {
                                        string units = dictionaryUnits[materialClassifierValue];

                                        MaterialToSend newMaterialToSend = new MaterialToSend()
                                        {
                                            ProjName = projName,
                                            SectionName = sectionNumberColumn.RowValues[i].Value,
                                            FloorName = levelNumberColumn.RowValues[i].Value,
                                            ProjWorkName = axaptaWorkset.Name,
                                            ProjWorkCodeId = axaptaWorkset.ProjWorkCodeId,
                                            ProjMaterialCodeId = materialClassifierValue,
                                            Volume = Double.Parse((from column in tableWithValues
                                                                    where column.ColumnName == units
                                                                    select column).FirstOrDefault().RowValues[i].Value),
                                            Units = units
                                        };

                                        MaterialsListToSentValuesToAxapta.Add(newMaterialToSend);
                                    }
                                }
                            }
                        }
                    }

                    if(IsExportToJSON == true)
                    {
                        string jsonWorks = JsonConvert.SerializeObject(WorksListToSentValuesToAxapta);
                        string newjsonWorks = UnidecodeSharpFork.Unidecoder.Unidecode(jsonWorks);

                        using (StreamWriter sw = new StreamWriter($"D:\\Выгрузка работ.txt", false, System.Text.Encoding.UTF8))
                        {
                            sw.Write(newjsonWorks);
                        }

                        string jsonMaterials = JsonConvert.SerializeObject(MaterialsListToSentValuesToAxapta);
                        string newjsonMaterials = UnidecodeSharpFork.Unidecoder.Unidecode(jsonMaterials);

                        using (StreamWriter sw = new StreamWriter($"D:\\Выгрузка материалов.txt", false, System.Text.Encoding.UTF8))
                        {
                            sw.Write(newjsonMaterials);
                        }
                    }

                    if (IsExportToExcel == true)
                    {

                    }
                    if (IsExportToAxapta == true)
                    {
                        try
                        {
                            axaptaService.SendToAxapta(WorksListToSentValuesToAxapta, SenderType.Work);
                            MessageBox.Show("Работы выгружены.");
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show($"Работы не выгружены.\n{ex.Message}\n{ex.StackTrace}");
                        }

                        try
                        {
                            axaptaService.SendToAxapta(MaterialsListToSentValuesToAxapta, SenderType.Material);
                            MessageBox.Show("Материалы выгружены.");
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Материалы не выгружены.\n{ex.Message}\n{ex.StackTrace}");
                        }
                    }
                    }
                }
                );
            }
        }


        private RelayCommand _getPathToUnitsFileCommand;

        public RelayCommand GetPathToUnitsFileCommand
        {
            get
            {
                return _getPathToUnitsFileCommand ?? new RelayCommand(obj =>
                {
                    GetPathToUnitsFileMethod();
                }
                );
            }
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

        #region ПРИВАТНЫЕ МЕТОДЫ
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

        private void GetPathToUnitsFileMethod()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;
            openFileDialog.Filter = "Excel (*.xlsx)|*.xlsx";
            openFileDialog.FilterIndex = 1;
            openFileDialog.ShowDialog();

            PathToUnitsFile = openFileDialog.FileName;
        }
        private void GetCSVTableMethod()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = true;
            openFileDialog.Filter = "CSV (*.csv)|*.csv|Excel (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*";
            openFileDialog.FilterIndex = 1;
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
            openFileDialog.FilterIndex = 1;

            openFileDialog.ShowDialog();
            PathToAllFiles = openFileDialog.FileName;
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

        private int GetEqualWorkElement(ObservableCollection<WorkToSend> inputList, string projName, string sectionName, string floorName, string projWorkCodeId)
        {
            int workToSendResult = -1;

            for (int i = 0; i <= inputList.Count - 1; i++)
            {
                var workToSend = inputList[i];

                if (workToSend.ProjName == projName &&
                    workToSend.SectionName == sectionName &&
                    workToSend.FloorName == floorName &&
                    workToSend.ProjWorkCodeId == projWorkCodeId)
                {
                    workToSendResult = i;
                    break;
                }

            }

            return workToSendResult;
        }

        private int GetEqualMaterialElement(ObservableCollection<MaterialToSend> inputList, string projName, string sectionName, string floorName, string projWorkCodeId, string projMaterialCodeId)
        {
            int materialToSendResult = -1;

            for (int i = 0; i <= inputList.Count - 1; i++)
            {
                var materialToSend = inputList[i];

                if (materialToSend.ProjName == projName &&
                    materialToSend.SectionName == sectionName &&
                    materialToSend.FloorName == floorName &&
                    materialToSend.ProjWorkCodeId == projWorkCodeId &&
                    materialToSend.ProjWorkCodeId == projMaterialCodeId)
                {
                    materialToSendResult = i;
                    break;
                }

            }

            return materialToSendResult;
        }

        /*Метод создания таблицы в формате .xlsx
         */
        private void SaveWorksAsExcel(List<WorkToSend> inputElementList, string path)
        {
            try
            {
                Application objWorkExcel = new Application();
                Workbook objWorkBook = objWorkExcel.Workbooks.Add();

                objWorkBook.Sheets.Add();
                var workSheet = (Worksheet)objWorkBook.Sheets[1];
                workSheet.Name = "Объемы";

                int rowCount = 1;

                workSheet.Cells[rowCount, 1] = "Номер секции";
                workSheet.Cells[rowCount, 2] = "Этаж";
                workSheet.Cells[rowCount, 3] = "Работа";
                workSheet.Cells[rowCount, 3] = "Код работы";
                workSheet.Cells[rowCount, 5] = "Ед. изм.";
                workSheet.Cells[rowCount, 6] = "Объем";
                Range header_range = workSheet.Range["A1", "F1"];

                foreach (var element in inputElementList)
                {
                    rowCount++;
                    workSheet.Cells[rowCount, 1] = element.SectionName;
                    workSheet.Cells[rowCount, 2] = element.FloorName;
                    workSheet.Cells[rowCount, 3] = element.ProjWorkName;
                    workSheet.Cells[rowCount, 4] = element.ProjWorkCodeId;
                    workSheet.Cells[rowCount, 5] = element.Units;
                    workSheet.Cells[rowCount, 6] = element.Volume;
                }

                objWorkBook.SaveAs(path);

                objWorkBook.Close(true, Type.Missing, Type.Missing);

                objWorkExcel.Quit();
                MessageBox.Show("Сделано!");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion
    }
}
