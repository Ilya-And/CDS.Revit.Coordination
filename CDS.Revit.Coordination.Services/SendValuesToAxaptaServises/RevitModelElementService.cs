using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDS.Revit.Coordination.Services
{
    public class RevitModelElementService
    {
        RevitModelElementService()
        {

        }

        //Метод заполнения параметра элемента
        public void SetParameter(Element element, string parameterName, string parameterValue)
        {
            element.LookupParameter(parameterName).Set(parameterValue);
        }
        public void SetParameter(Element element, string parameterName, int parameterValue)
        {
            element.LookupParameter(parameterName).Set(parameterValue);
        }
        public void SetParameter(Element element, string parameterName, double parameterValue)
        {
            element.LookupParameter(parameterName).Set(parameterValue);
        }

        //Метод восстановления формы
        public void RestorateForm(Element element)
        {

        }
        public void RestorateForm(List<Element> elements)
        {
            
        }

        //Метод создания частей
        public void CreateParts(ElementId elementId)
        {

        }
        public void CreateParts(List<ElementId> elementIds)
        {

        }
    }
}