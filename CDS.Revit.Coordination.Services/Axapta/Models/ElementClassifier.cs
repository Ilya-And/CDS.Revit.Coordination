using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDS.Revit.Coordination.Services.Axapta
{
    public class ElementClassifier
    {
        public string id { get; set; }
        public string Name { get; set; }
        public List<string> WorkCodeID { get; set; }
        public string WorkCodeIDListAsString
        {
            get
            {
                return String.Join("; ", WorkCodeID);
            }
        }
    }
}
