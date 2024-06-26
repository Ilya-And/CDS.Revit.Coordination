﻿using System;
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
using CDS.Revit.Coordination.Services.Axapta;
using CDS.Revit.Coordination.Services.Revit;
using CDS.Revit.Coordination.Services.Excel;
using Microsoft.Win32;
using Autodesk.Revit.UI;

namespace CDS.Revit.Coordination.Test
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void StartButton_Click(object sender, RoutedEventArgs e)
        {
            AxaptaService axaptaService = new AxaptaService("https://tstaxapi.cds.spb.ru/", "nevis", "HPJoP/Y/33NPdTeITGd0WQ==");

            ExcelService excelService = new ExcelService();

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.ShowDialog();
            var result = excelService.GetValuesFromExcelTable(openFileDialog.FileName);
            var resultaxapta = axaptaService.GetWorksFromAxapta();
            //var classifiers = axaptaService.GetAllClassifiersSections();

            //var elementClassifier = axaptaService.GetAllElementClassifiersDict(classifiers);
            //try
            //{
            //    var worksFromAxapta = axaptaService.GetAllAxaptaWorksetsMethod();
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show("Error");
            //}

        }
    }
}