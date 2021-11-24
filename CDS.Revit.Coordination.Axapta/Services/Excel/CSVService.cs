using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDS.Revit.Coordination.Axapta.Services.Excel
{
    public class CSVService
    {
        public List<ColumnValues> GetValuesFromCSVTable (string filePath)
        {
            var result = new List<ColumnValues>();
            string[] lines = System.IO.File.ReadAllLines(filePath);
            var firstLineAsString = lines[0];
            var firstLineAsList = firstLineAsString.Split(',').ToList();
            var columnCount = firstLineAsList.Count;
            var rowCount = lines.Count();
            for(int i = 0; i <= columnCount; i++)
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
