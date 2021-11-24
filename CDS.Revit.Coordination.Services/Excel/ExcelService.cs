using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MExcel = Microsoft.Office.Interop.Excel;

namespace CDS.Revit.Coordination.Services.Excel
{
    public class ExcelService
    {
        public ExcelService()
        {

        }
        public string GetValuesFromExcelTable(string filePath)
        {
            var resultList = new List<ColumnValues>();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;

                reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                };

                var dataSet = reader.AsDataSet(conf);

                var dataTable = dataSet.Tables[0];
                return "OK";
            }

            //MExcel.Application objWorkExcel = new MExcel.Application();
            //MExcel.Workbook workbook = objWorkExcel.Workbooks.Open(filePath, Type.Missing, Type.Missing, Type.EmptyTypes, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            //MExcel.Worksheet worksheet = workbook.Sheets[1];
            //MExcel.Range xlRange = worksheet.UsedRange;
            //var countColumns = worksheet.Cells.Count;
            //var countRows = worksheet.Rows.Count;

            //for (int i = 1; i <= countColumns; i++)
            //{
            //    var newColumn = new ColumnValues();
            //    newColumn.Id = i;

            //    for(int n = 1; n <= countRows; n++)
            //    {
            //        if(n == 1)
            //        {
            //            newColumn.ColumnName = (worksheet.Cells[n, i] as MExcel.Range).Text;
            //        }
            //        else
            //        {
            //            var newRow = new RowValue();
            //            newRow.Id = n;
            //            newRow.Value = (worksheet.Cells[n, i] as MExcel.Range).Text;
            //            newColumn.RowValues.Add(newRow);
            //        }
            //    }
            //}

            //return resultList;
        }
    }
}
