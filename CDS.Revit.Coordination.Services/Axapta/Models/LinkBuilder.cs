using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDS.Revit.Coordination.Services.Axapta
{
    public class LinkBuilder
    {
        public static string Work => HOST + "api/Navis/AddNavisData";
        public static string Material => HOST + "api/Navis/AddNavisDataItem";
        public static string Token => HOST + "api/Account/token";
        private static string Classifier => HOST + "api/Navis/ClassifierCodeTable";
        public static string ClassifierType => HOST + "api/Navis/ClassifierCodeTableType";
        public static string ProjWorkTable => HOST + "api/Navis/ProjWorkTable";

        public static string HOST => "https://tstaxapi.cds.spb.ru/";

        public static string GetClassifierLink(string type)
        {
            return Classifier + "?type=" + type;
        }
    }


    public enum SenderType
    {
        Work,
        Material
    }
}
