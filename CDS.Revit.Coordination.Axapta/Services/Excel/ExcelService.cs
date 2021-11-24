using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MExcel = Microsoft.Office.Interop.Excel;

namespace CDS.Revit.Coordination.Axapta.Services.Excel
{
    public class ExcelService
    {
        /*Метод получения данных из таблицы .xlsx в формате: Столбец : Список строк столбца
         */
        public List<ColumnValues> GetValuesFromExcelTable(string filePath)
        {
            // Инициализируем возвращаемый список
            var resultList = new List<ColumnValues>();

            //Открываем файл по указанному пути.
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                IExcelDataReader reader;

                reader = ExcelReaderFactory.CreateReader(stream);

                var conf = new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = false
                    }
                };

                //Получаем таблицу
                var dataSet = reader.AsDataSet(conf);

                var dataTable = dataSet.Tables[0];

                //Получаем строки, столбцы и их количество
                var rows = dataTable.Rows;
                var columns = dataTable.Columns;

                var countColumns = columns.Count;
                var countRows = rows.Count;

                //Заполняем возвращаемый список
                for (int i = 0; i <= countColumns - 1; i++)
                {
                    var newColumn = new ColumnValues();
                    newColumn.Id = i;

                    for (int n = 0; n <= countRows - 1; n++)
                    {
                        if (n == 0)
                        {
                            newColumn.ColumnName = rows[n][i].ToString();
                        }
                        else
                        {
                            var newRow = new RowValue();
                            newRow.Id = n;
                            try
                            {
                                newRow.Value = rows[n][i].ToString();
                            }
                            catch
                            {
                                newRow.Value = "";
                            }
                            newColumn.RowValues.Add(newRow);
                        }
                    }
                    resultList.Add(newColumn);
                }   
            }
            return resultList;
        }
    }
}
