using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDS.Revit.Coordination.Axapta.Models
{
    [Serializable]
    public class WorkToSend
    {
        public string ProjName { get; set; }
        public string SectionName { get; set; }
        public string FloorName { get; set; }
        public string ProjWorkName { get; set; }
        public string ProjWorkCodeId { get; set; }
        public double Volume { get; set; }
        public string Units { get; set; }
    }
}
