using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDS.Revit.Coordination.Services.Excel
{
    public class ColumnValues
    {
        public int Id { get; set; }
        public string ColumnName { get; set; }
        public List<RowValue> RowValues { get; set; } = new List<RowValue>();
    }
}
