using Newtonsoft.Json;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel.DataAnnotations.Schema;

namespace CDS.Revit.Coordination.Services.Axapta
{
    public class WorksetToSend
    {
        public string ProjName { get; set; }
        public string SectionName { get; set; }
        public string FloorName { get; set; }
        [JsonIgnore]
        private AxaptaWorkset _axaptaWorkset { get; set; }
        [JsonIgnore]
        public string ProjWorkName
        {
            get
            {
                return this._axaptaWorkset.Name;
            }
        }
        public string ProjWorkCodeId
        {
            get
            {
                return this._axaptaWorkset.ProjWorkCodeId;
            }
        }
        public double Volume { get; set; }
        public string Units
        {
            get
            {
                return this._axaptaWorkset.UnitId;
            }
        }
        public WorksetToSend(AxaptaWorkset axaptaWorkset)
        {
            this._axaptaWorkset = axaptaWorkset;
        }

    }
}
