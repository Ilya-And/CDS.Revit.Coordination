using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Autodesk.Revit.DB;
using Autodesk.Revit.Exceptions;
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

        #region КОМАНДЫ

        /*Основная команда, запускающая процесс обработки моделей, 
         подготовки и отправки данных в Axapta,
        а также сохранения данных в различных форматах*/
        private RelayCommand _startCommand;
        public RelayCommand StartCommand
        {
            get
            {
                return _startCommand ?? new RelayCommand(obj =>
                {
                    if(PathToAllFiles != null)
                    {
                        ExcelService excelService = new ExcelService();
                        RevitFileService revitFileService = new RevitFileService(SendValuesCommand.App);
                         
                        try
                        {
                            List<ColumnValues> generalTable = excelService.GetValuesFromExcelTable(PathToAllFiles);
                            GeneralTable = generalTable;
                            GetColumnsFromExcel();

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
                                                            switch (categoryNameFromElement)
                                                            {
                                                                case "Части":
                                                                    Part asPart = element as Part;
                                                                    var categoryNameHost = asPart.get_Parameter(BuiltInParameter.DPART_ORIGINAL_CATEGORY).AsString();

                                                                    var parameterValue1FromElement = "";
                                                                    var parameterValue2FromElement = "";
                                                                    var parameterValue3FromElement = "";
                                                                    var parameterValue4FromElement = "";
                                                                    var parameterValue5FromElement = "";

                                                                    parameterValue1FromElement = asPart.LookupParameter(parameter1FromTable)?.AsString();
                                                                    parameterValue2FromElement = asPart.LookupParameter(parameter2FromTable)?.AsString();
                                                                    parameterValue3FromElement = asPart.LookupParameter(parameter3FromTable)?.AsString();
                                                                    parameterValue4FromElement = asPart.LookupParameter(parameter4FromTable)?.AsString();
                                                                    parameterValue5FromElement = asPart.LookupParameter(parameter5FromTable)?.AsString();

                                                                    bool isFirstParameterValueMatch = false;
                                                                    bool isSecondParameterValueMatch = false;
                                                                    bool isThirdParameterValueMatch = false;
                                                                    bool isFouthParameterValueMatch = false;
                                                                    bool isFifthParameterValueMatch = false;

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
                                                                        string classifierForSet = "";
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
                                                                        else
                                                                        {
                                                                            classifierForSet = classifierLineFromTable;
                                                                        }
                                                                        asPart.LookupParameter("ЦДС_Классификатор")?.Set(classifierLineFromTable);

                                                                        if (materialClassifierFromTable != "" || materialClassifierFromTable != "по материалу")
                                                                        {
                                                                            asPart.LookupParameter("ЦДС_Классификатор материалов")?.Set(materialClassifierFromTable);
                                                                        }
                                                                    }

                                                                    break;

                                                                case "Несущие колонны":

                                                                    break;

                                                                case "Каркас несущий":

                                                                    break;

                                                                case "Лестницы":

                                                                    break;

                                                                case "Фундамент несущей конструкции":

                                                                    break;

                                                                case "Ограждение":

                                                                    break;

                                                                case "Обобщенные модели":

                                                                    break;

                                                                case "Двери":

                                                                    break;

                                                                case "Окна":

                                                                    break;

                                                                case "Стены":

                                                                    break;

                                                                case "Панели витража":

                                                                    break;

                                                                case "Импосты витража":

                                                                    break;

                                                                case "Воздуховоды":

                                                                    break;

                                                                case "Трубы":

                                                                    break;

                                                                case "Материалы изоляции воздуховодов":

                                                                    break;

                                                                case "Материалы изоляции труб":

                                                                    break;

                                                                case "Арматура воздуховодов":

                                                                    break;

                                                                case "Арматура трубопроводов":

                                                                    break;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                MessageBox.Show(ex.Message + "\n" + ex.StackTrace + "\n" + element);
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


                                    revitFileService.ExportToNWC(PathToSaveColumn.RowValues[i].Value, doc);
                                    revitFileService.SaveAndCloseRVTFile(PathToSaveColumn.RowValues[i].Value, doc);

                                }
                                else
                                {
                                    IsAllFilesWork = false;
                                    continue;
                                }
                            }
                        }
                        catch(Exception ex)
                        {
                            MessageBox.Show(ex.StackTrace);
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

        #endregion

        private void GetGeneralExcelTableMethod()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            PathToAllFiles = openFileDialog.FileName;
        }

        private void StartCommandMethod()
        {
            MessageBox.Show("Work!");
        }

    }
}
