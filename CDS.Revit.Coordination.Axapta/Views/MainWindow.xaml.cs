using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Autodesk.Revit.DB;
using CDS.Revit.Coordination.Axapta.Services.Axapta;
using CDS.Revit.Coordination.Axapta.Services.Excel;
using CDS.Revit.Coordination.Axapta.Services.Revit;
using Microsoft.Win32;
using CDS.Revit.Coordination.Axapta.ViewModels;

namespace CDS.Revit.Coordination.Axapta.Views
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            DataContext = new MainWindowViewModel();
        }


        //private void StartButton_Click(object sender, RoutedEventArgs e)
        //{
        //    AxaptaService axaptaService = new AxaptaService("https://tstaxapi.cds.spb.ru/", "nevis", "HPJoP/Y/33NPdTeITGd0WQ==");

        //    ExcelService excelService = new ExcelService();

        //    CSVService CSVService = new CSVService();

        //    RevitModelElementService revitModelElementService = new RevitModelElementService(SendValuesCommand.Doc);

        //    //OpenFileDialog openFileDialog = new OpenFileDialog();
        //    //openFileDialog.ShowDialog();

        //    var rebars = new FilteredElementCollector(SendValuesCommand.Doc).OfCategory(BuiltInCategory.OST_Rebar).WhereElementIsNotElementType().ToElements();
        //    using (Transaction tr = new Transaction(SendValuesCommand.Doc))
        //    {
        //        tr.Start("Заполнение классификатора для арматуры");
        //        revitModelElementService.SetClassifierParameterValueToRebar(rebars);
        //        tr.Commit();
        //    }
        //    //MessageBox.Show(result);
        //    //var resultaxapta = axaptaService.GetWorksFromAxapta();
        //    //var classifiers = axaptaService.GetAllClassifiersSections();

        //    //var elementClassifier = axaptaService.GetAllElementClassifiersDict(classifiers);
        //    //try
        //    //{
        //    //    var worksFromAxapta = axaptaService.GetAllAxaptaWorksetsMethod();
        //    //}
        //    //catch (Exception ex)
        //    //{
        //    //    MessageBox.Show("Error");
        //    //}

        //}
    }
}