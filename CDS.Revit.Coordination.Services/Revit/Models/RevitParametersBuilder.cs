using System;

namespace CDS.Revit.Coordination.Services.Revit.Models
{
    public class RevitParametersBuilder
    {
        public static RevitParameter FloorNumber => new RevitParameter("ADSK_Этаж", new Guid("9eabf56c-a6cd-4b5c-a9d0-e9223e19ea3f"));
        public static RevitParameter SectionNumber => new RevitParameter("ADSK_Номер секции", new Guid("b59a3474-a5f4-430a-b087-a20f1a4eb57e"));
        public static RevitParameter Classifier => new RevitParameter("ЦДС_Классификатор");
        public static RevitParameter ClassifierMaterial => new RevitParameter("ЦДС_Классификатор_Материалов");
    }

    public class RevitParameter
    {
        public string Name { get; private set; }
        public Guid Guid { get; private set; }

        public RevitParameter(string name, Guid guid)
        {
            Name = name;
            Guid = guid;
        }

        public RevitParameter(string name)
        {
            Name = name;
        }
    }
}
