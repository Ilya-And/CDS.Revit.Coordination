using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CDS.Revit.Coordination.Services.Excel
{
    public class CSVService
    {
        /*Метод получения данных из таблицы .csv в формате: Столбец : Список строк столбца
         */
        public List<ColumnValues> GetValuesFromCSVTable (string filePath)
        {
            // Инициализируем возвращаемый список
            var result = new List<ColumnValues>();

            //Открываем файл по указанному пути.
            string[] lines = System.IO.File.ReadAllLines(filePath, Encoding.Default);

            //Получаем значения первой строки
            var firstLineAsString = lines[0];
            var firstLineAsList = firstLineAsString.Split(',').ToList();

            //Получаем количество столбцов и строк
            var columnCount = firstLineAsList.Count;
            var rowCount = lines.Count();

            //Заполняем возвращаемый список
            for(int i = 0; i <= columnCount - 1; i++)
            {
                var newColumn = new ColumnValues();
                newColumn.Id = i;
                newColumn.ColumnName = firstLineAsList[i];

                for(int n = 1; n <= rowCount; n++)
                {
                    var newRow = new RowValue();
                    newRow.Id = n;
                    try
                    {
                        newRow.Value = lines[n].Split(',')[i];
                    }
                    catch
                    {
                        newRow.Value = "";
                    }
                    newColumn.RowValues.Add(newRow);
                    

                }
                result.Add(newColumn);
            }

            return result;
        }
    }
}
