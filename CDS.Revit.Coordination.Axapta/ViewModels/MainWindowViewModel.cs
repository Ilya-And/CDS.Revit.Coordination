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
                _projectNameColumn = (from column in GeneralTable
                                where column.ColumnName == "Проект"
                                select column).FirstOrDefault();
                OnPropertyChanged("ProjectNameColumn");
            }
        }

        private ColumnValues _projectSectionColumn { get; set; }
        public ColumnValues ProjectSectionColumn
        {
            get => _projectSectionColumn;
            set
            {
                _projectSectionColumn = (from column in GeneralTable
                                   where column.ColumnName == "Раздел"
                                   select column).FirstOrDefault();
                OnPropertyChanged("ProjectSectionColumn");
            }
        }

        private ColumnValues _pathToSaveColumn;
        public ColumnValues PathToSaveColumn
        {
            get => _pathToSaveColumn;
            set
            {
                _pathToSaveColumn = (from column in GeneralTable
                               where column.ColumnName == "ПутьДляСохранения"
                               select column).FirstOrDefault();
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
        }
        private ColumnValues _pathTableColumn;
        public ColumnValues PathTableColumn
        {
            get => _pathTableColumn;
            set
            {
                _pathTableColumn = (from column in GeneralTable
                              where column.ColumnName == "ПутьТаблицыВыбора"
                              select column).FirstOrDefault();
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
                                    var views3DForSetSectionParameter = revitModelElementService.Get3DViewForSetSectionParameter();

                                    var floors = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType().ToElements();
                                    var roofs = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Roofs).WhereElementIsNotElementType().ToElements();

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

                                    //using (Transaction tr = new Transaction(doc))
                                    //{
                                    //    tr.Start("Заполнение параметров классификатора");


                                    //    doc.Regenerate();

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
                            MessageBox.Show(ex.Message);
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
