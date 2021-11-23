using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDS.Revit.Coordination.Services.Revit
{
    public class RevitModelElementService
    {
        private Document _doc;
        public RevitModelElementService(Document doc)
        {
            _doc = doc;
        }

        /*Метод заполнения параметра элемента
         */
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

        /*Метод восстановления формы у элементов:
        Перекрытия
        Крыша по контуру
        Крыша выдавливанием*/
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

        /*Метод создания частей из элементов:
         */
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
                if (part != null)
                {
                    var hostId = part.GetSourceElementIds().ToList()[0].HostElementId;
                    var hostElement = _doc.GetElement(hostId);
                    if(hostElement != null)
                    {

                    }
                }


            }
        }

        /*Метод получения неправильно названных уровней в проекте
         */
        public string GetUnCorrectLevelsNames()
        {
            // Получаем список всех уровней в документе

            var allLevelsInDocument = new FilteredElementCollector(_doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements();
            string unCorrectLevelsNames = "";

            // Итерируемся по списку

            foreach (Element levelAsElement in allLevelsInDocument)
            {
                Level level = levelAsElement as Level;
                
                // Проверяем корректность имени уровня

                string levelName = level.Name;
                var splitedLevelName = levelName.Split(' ');
                bool isCorrect = false;

                if (splitedLevelName.Length == 2)
                {
                    if (splitedLevelName[0].ToLower() == "этаж")
                    {
                        isCorrect = true;
                    }
                }
                else if (splitedLevelName.Length == 1)
                {
                    if (levelName.ToLower() == "подвал"
                        || levelName.ToLower() == "кровля"
                        || levelName.ToLower() == "паркинг"
                        || levelName.ToLower() == "тех.этаж")
                    {
                        isCorrect = true;
                    }


                }
                if (isCorrect == false)
                {

                    unCorrectLevelsNames = unCorrectLevelsNames + "\n" + levelName;

                }

            }
            return unCorrectLevelsNames;
        }

        /*Метод получения номера/имени этажа для заполнения параметра ADSK_Этаж
         */
        public string GetLevelNumber(Element element, ElementId levelId, Parameter levelParam)
        {
            //Получаем имя уровня

            Level level = _doc.GetElement(levelId) as Level;
            string levelName = level.Name;
            var levelNumberList = levelName.Split(' ');
            var levelNumber = levelName;

            // Получаем порядковый номер или наименование этажа

            if (levelNumberList.Length == 2 && levelNumberList[0].ToLower() == "этаж")
            {
                if (levelNumberList[1].StartsWith("0"))
                {
                    levelNumber = levelNumberList[1].Substring(1);
                }
                else if (levelNumberList[1].Contains(',') || levelNumberList[1].Contains('.'))
                {
                    levelNumber = levelNumberList[1].Substring(0, 1);
                }
                else
                {
                    levelNumber = levelNumberList[1];
                }
            }
            else if (levelNumberList.Length == 1)
            {
                switch (levelName.ToLower())
                {
                    case "подвал":
                        levelNumber = "Подвал";
                        break;
                    case "кровля":
                        levelNumber = "Кровля";
                        break;
                    case "паркинг":
                        levelNumber = "Паркинг";
                        break;
                    case "тех.этаж":
                        levelNumber = "Тех.этаж";
                        break;
                }
            }

            return levelNumber;
        }

        /*Метод получения 3D вида для экспорта модели в формат .nwc
         */
        public View3D Get3DViewForExportToNWC()
        {
            var allViews = new FilteredElementCollector(_doc).OfClass(typeof(View3D)).WhereElementIsNotElementType().ToElements();
            View3D view3DForExport = null;
            foreach(Element elementView in allViews)
            {
                View3D view3D = elementView as View3D;
                string groupParameterValue = view3D.LookupParameter("ADSK_Подгруппа").AsString();
                string viewName = view3D.Name;
                if(viewName.Contains("Axapta") && viewName.Contains("Navisworks"))
                {
                    if (groupParameterValue.Contains("Axapta"))
                    {
                        view3DForExport = view3D;
                        break;
                    }
                }
            }
            return view3DForExport;
        }
    }
}