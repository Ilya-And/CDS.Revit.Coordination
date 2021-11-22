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
        private Document _doc;
        public RevitModelElementService(Document doc)
        {
            _doc = doc;
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
            Floor asFloor = element as Floor;
            FootPrintRoof asFootPrintRoof = element as FootPrintRoof;
            ExtrusionRoof asExtrusionRoof = element as ExtrusionRoof;

            if (asFloor != null)
            {
                try
                {
                    if (asFloor.SlabShapeEditor.IsEnabled == true)
                    {
                        asFloor.SlabShapeEditor.ResetSlabShape();
                    }
                }
                catch
                {
                    
                }

            }

            if (asFootPrintRoof != null)
            {
                try
                {
                    if (asFootPrintRoof.SlabShapeEditor.IsEnabled == true)
                    {
                        asFootPrintRoof.SlabShapeEditor.ResetSlabShape();
                    }
                }
                catch
                {
                    
                }
            }

            if (asExtrusionRoof != null)
            {
                try
                {
                    if (asExtrusionRoof.SlabShapeEditor.IsEnabled == true)
                    {
                        asExtrusionRoof.SlabShapeEditor.ResetSlabShape();
                    }
                }
                catch
                {
                    
                }
            }
        }
        public void RestorateForm(List<Element> elements)
        {
            foreach (Element element in elements)
            {
                Floor asFloor = element as Floor;
                FootPrintRoof asFootPrintRoof = element as FootPrintRoof;
                ExtrusionRoof asExtrusionRoof = element as ExtrusionRoof;

                if (asFloor != null)
                {
                    try
                    {
                        if (asFloor.SlabShapeEditor.IsEnabled == true)
                        {
                            asFloor.SlabShapeEditor.ResetSlabShape();
                        }
                    }
                    catch
                    {
                        continue;
                    }

                }

                if (asFootPrintRoof != null)
                {
                    try
                    {
                        if (asFootPrintRoof.SlabShapeEditor.IsEnabled == true)
                        {
                            asFootPrintRoof.SlabShapeEditor.ResetSlabShape();
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }

                if (asExtrusionRoof != null)
                {
                    try
                    {
                        if (asExtrusionRoof.SlabShapeEditor.IsEnabled == true)
                        {
                            asExtrusionRoof.SlabShapeEditor.ResetSlabShape();
                        }
                    }
                    catch
                    {
                        continue;
                    }
                }
            }
        }

        //Метод создания частей
        public void CreateParts(ElementId elementId)
        {
            var listWithElementId = new List<ElementId>() { elementId };
            PartUtils.CreateParts(_doc, listWithElementId);
        }
        public void CreateParts(List<ElementId> elementIds)
        {
            foreach (ElementId elementId in elementIds)
            {
                var listWithElementId = new List<ElementId>() { elementId };
                try
                {
                    PartUtils.CreateParts(_doc, listWithElementId);
                }
                catch
                {
                    continue;
                }
            }
        }

        public void SetPartsParametersByHost(ICollection<Part> parts)
        {
            foreach (Part part in parts)
            {
                var hostId = part.GetSourceElementIds().ToList()[0].HostElementId;
                var hostElement = _doc.GetElement(hostId);


            }
        }
    }
}