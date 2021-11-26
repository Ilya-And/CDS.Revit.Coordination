using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Prism.Commands;
using Prism.Mvvm;

namespace CDS.Revit.Coordination.Axapta.ViewModels
{
    class MainWindowViewModel : BindableBase
    {
        private bool _isExportToExcel;
        public bool IsExportToExcel
        {
            get => _isExportToExcel;
            set => SetProperty(ref _isExportToExcel, value);
        }

        public MainWindowViewModel()
        {
            StartCommand = new DelegateCommand(StartCommandMethod).ObservesCanExecute(() => IsExportToExcel);
        }
        public DelegateCommand StartCommand { get; private set; }

        private void StartCommandMethod()
        {
            MessageBox.Show("Work!");
        }
    }
}
