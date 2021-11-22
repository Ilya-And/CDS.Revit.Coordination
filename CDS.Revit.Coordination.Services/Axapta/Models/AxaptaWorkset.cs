using System;
using System.Collections.ObjectModel;
using System.ComponentModel.DataAnnotations.Schema;

namespace CDS.Revit.Coordination.Services
{
    [Serializable]
    public class AxaptaWorkset
    {
        public bool Archive { get; set; }
        public string Name { get; set; }
        public string ParentCodeId { get; set; }
        public string ProjWorkCodeId { get; set; }
        public string UnitId { get; set; }
        public AxaptaWorkset() { }
        [NotMapped]
        public ObservableCollection<AxaptaWorkset> AxaptaWorksets { get; set; }
    }
}
