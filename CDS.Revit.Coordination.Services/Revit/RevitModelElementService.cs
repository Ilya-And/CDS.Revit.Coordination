using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using CDS.Revit.Coordination.Services.Revit.Models;
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
        private int _countElements = 0;
        public string SectionParameter = RevitParametersBuilder.SectionNumber.Name;
        public string LevelParameter = RevitParametersBuilder.FloorNumber.Name;
        public string ClassifierParameter = RevitParametersBuilder.Classifier.Name;

        const string EMBEDDED_PARTS = "детали закладных";

        private const string BASEMENT = "Подвал";
        private const string BASEMENT_LOWER = "подвал";

        private const string ROOF = "Кровля";
        private const string ROOF_LOWER = "кровля";

        private const string PARKING = "Паркинг";
        private const string PARKING_LOWER = "паркинг";

        private const string TECH_FLOOR = "Тех.этаж";
        private const string TECH_FLOOR_LOWER = "тех.этаж";

        private const string FLOOR_LOWER = "этаж";

        public RevitModelElementService(Document doc)
        {
            _doc = doc;
        }

        /*Метод получения номера/имени этажа для заполнения параметра ADSK_Этаж
         */
        private string GetLevelNumber(Element element, ElementId levelId, Parameter levelParam)
        {
            //Получаем имя уровня

            Level level = _doc.GetElement(levelId) as Level;
            string levelName = level.Name;
            var levelNumberList = levelName.Split(' ');
            var levelNumber = levelName;

            // Получаем порядковый номер или наименование этажа

            if (levelNumberList.Length == 2 && levelNumberList[0].ToLower() == FLOOR_LOWER)
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
                    case BASEMENT_LOWER:
                        levelNumber = BASEMENT;
                        break;
                    case ROOF_LOWER:
                        levelNumber = ROOF;
                        break;
                    case PARKING_LOWER:
                        levelNumber = PARKING;
                        break;
                    case TECH_FLOOR_LOWER:
                        levelNumber = TECH_FLOOR;
                        break;
                }
            }

            return levelNumber;
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
        public void RestorateForm(ICollection<Element> elements)
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
        public void CreatePartsFromElement(ICollection<ElementId> elementIds)
        {
            foreach(ElementId elementId in elementIds)
            {
                try
                {
                    var listWithElementId = new List<ElementId>() { elementId };
                    PartUtils.CreateParts(_doc, listWithElementId);
                }
                catch
                {
                    continue;
                }
            }
        }
        public void CreatePartsFromElement(List<ElementId> elementIds)
        {
            PartUtils.CreateParts(_doc, elementIds);
        }

        /*Метод заполнения у частей параметров ADSK_Номер секции и ADSK_Этаж
         */
        public void SetPartsParametersLevelAndSectionByHost(ICollection<Part> parts)
        {
            foreach (Part part in parts)
            {
                if (part != null)
                {
                    var hostId = part.GetSourceElementIds().ToList()[0].HostElementId;
                    var hostElement = _doc.GetElement(hostId);
                    string sectionNumber = "";
                    string levelNumber = "";

                    if(hostElement != null)
                    {
                        var asFamilyInstance = hostElement as FamilyInstance;
                        var asWall = hostElement as Wall;
                        var asFloor = hostElement as Floor;
                        var asWallSweep = hostElement as WallSweep;
                        var asRoof = hostElement as FootPrintRoof;
                        var asExtrusionRoof = hostElement as ExtrusionRoof;
                        var asPanel = hostElement as Panel;
                        var asCeiling = hostElement as Ceiling;

                        if (asFamilyInstance != null)
                        {
                            sectionNumber = asFamilyInstance.LookupParameter(SectionParameter).AsString();
                            levelNumber = asFamilyInstance.LookupParameter(LevelParameter).AsString();
                        }

                        if (asWall != null)
                        {
                            sectionNumber = asWall.LookupParameter(SectionParameter).AsString();
                            levelNumber = asWall.LookupParameter(LevelParameter).AsString();
                        }

                        if (asFloor != null)
                        {
                            sectionNumber = asFloor.LookupParameter(SectionParameter).AsString();
                            levelNumber = asFloor.LookupParameter(LevelParameter).AsString();
                        }

                        if (asWallSweep != null)
                        {
                            sectionNumber = asWallSweep.LookupParameter(SectionParameter).AsString();
                            levelNumber = asWallSweep.LookupParameter(LevelParameter).AsString();
                        }

                        if (asRoof != null)
                        {
                            sectionNumber = asRoof.LookupParameter(SectionParameter).AsString();
                            levelNumber = asRoof.LookupParameter(LevelParameter).AsString();
                        }

                        if (asExtrusionRoof != null)
                        {
                            sectionNumber = asExtrusionRoof.LookupParameter(SectionParameter).AsString();
                            levelNumber = asExtrusionRoof.LookupParameter(LevelParameter).AsString();
                        }

                        if (asPanel != null)
                        {
                            sectionNumber = asPanel.LookupParameter(SectionParameter).AsString();
                            levelNumber = asPanel.LookupParameter(LevelParameter).AsString();
                        }

                        if (asCeiling != null)
                        {
                            sectionNumber = asCeiling.LookupParameter(SectionParameter).AsString();
                            levelNumber = asCeiling.LookupParameter(LevelParameter).AsString();
                        }

                        part.LookupParameter(SectionParameter).Set(sectionNumber);
                        part.LookupParameter(LevelParameter).Set(levelNumber);
                    }
                }
            }
        }

        /*Метод заполнения у элемента параметров ADSK_Номер секции и ADSK_Этаж
         */
        public void SetParametersValuesToElement(Element element, string sectionNumber)
        {
            var category = element.Category;
            var levelsList = new FilteredElementCollector(_doc).OfClass(typeof(Level)).WhereElementIsNotElementType().ToElements() as List<Level>;

            switch ((BuiltInCategory)category.Id.IntegerValue)
            {
                case BuiltInCategory.OST_Walls:
                    var asWall = element as Wall;
                    asWall.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                    var levelWallParam = asWall.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                    var levelIdWall = asWall.LevelId;
                    if (levelWallParam != null && levelWallParam.AsValueString() != "Cекция" || levelWallParam.AsValueString() != "секция")
                    {
                        levelWallParam.Set(GetLevelNumber(element, levelIdWall, levelWallParam));
                        _countElements++;
                    }
                    break;

                case BuiltInCategory.OST_Cornices:
                    var asWallSweep = element as WallSweep;
                    asWallSweep.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                    var levelWallSweepParam = asWallSweep.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                    var levelIdWallSweep = asWallSweep.LevelId;
                    if (levelWallSweepParam != null)
                    {
                        asWallSweep.LookupParameter(RevitParametersBuilder.FloorNumber.Name).Set(GetLevelNumber(element, levelIdWallSweep, levelWallSweepParam));
                        _countElements++;
                    }
                    break;

                case BuiltInCategory.OST_Floors:
                    var asFloor = element as Floor;

                    if(asFloor != null)
                    {
                        asFloor.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                        var levelFloorParam = asFloor.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                        var levelIdFloor = asFloor.LevelId;
                        if (levelFloorParam != null)
                        {
                            levelFloorParam.Set(GetLevelNumber(element, levelIdFloor, levelFloorParam));
                            _countElements++;
                        }
                    }
                    break;

                case BuiltInCategory.OST_StructuralFoundation:
                    var asFloorFoundation = element as Floor;
                    var asFamilyInstanceFoundation = element as FamilyInstance;
                    if (asFloorFoundation != null)
                    {
                        asFloorFoundation.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                        var levelFloorFoundationParam = asFloorFoundation.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                        var levelIdFloorFoundation = asFloorFoundation.LevelId;
                        if (levelFloorFoundationParam != null)
                        {
                            levelFloorFoundationParam.Set(GetLevelNumber(element, levelIdFloorFoundation, levelFloorFoundationParam));
                            _countElements++;
                        }
                    }
                    else if (asFamilyInstanceFoundation != null)
                    {
                        asFamilyInstanceFoundation.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                        var levelInstanceFoundationParam = asFamilyInstanceFoundation.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                        var levelIdInstanceFoundation = asFamilyInstanceFoundation.LevelId;
                        if (levelInstanceFoundationParam != null)
                        {
                            levelInstanceFoundationParam.Set(GetLevelNumber(element, levelIdInstanceFoundation, levelInstanceFoundationParam));
                            _countElements++;
                        }
                    }
                    break;

                case BuiltInCategory.OST_Roofs:
                    var asRoof = element as FootPrintRoof;
                    var asExtrusionRoof = element as ExtrusionRoof;
                    if (asRoof != null)
                    {
                        asRoof.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                        var levelRoofParam = asRoof.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                        var levelIdRoof = asRoof.LevelId;
                        if (levelRoofParam != null)
                        {
                            levelRoofParam.Set(GetLevelNumber(element, levelIdRoof, levelRoofParam));
                            _countElements++;
                        }
                    }
                    if (asExtrusionRoof != null)
                    {
                        asExtrusionRoof.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                        var levelExtrusionRoofParam = asExtrusionRoof.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                        var levelIdRoof = asExtrusionRoof.get_Parameter(BuiltInParameter.ROOF_CONSTRAINT_LEVEL_PARAM).AsElementId();
                        if (levelExtrusionRoofParam != null)
                        {
                            asExtrusionRoof.LookupParameter(RevitParametersBuilder.FloorNumber.Name).Set(GetLevelNumber(element, levelIdRoof, levelExtrusionRoofParam));
                            _countElements++;
                        }
                    }
                    break;

                case BuiltInCategory.OST_CurtainWallMullions:
                    var asMullion = element as Mullion;
                    asMullion.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                    var levelMullionParam = asMullion.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                    var hostMullion = asMullion.Host as Wall;
                    var levelIdMullion = hostMullion.LevelId;
                    if (levelMullionParam != null)
                    {
                        levelMullionParam.Set(GetLevelNumber(element, levelIdMullion, levelMullionParam));
                        _countElements++;
                    }
                    break;

                case BuiltInCategory.OST_CurtainWallPanels:
                    var asPanel = element as Autodesk.Revit.DB.Panel;
                    var asPanelWall = element as Wall;
                    var asPanelInstance = element as FamilyInstance;
                    var levelIdPanel = new ElementId(0);
                    if (asPanel != null)
                    {
                        asPanel.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                        var levelPanelParam = asPanel.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                        var hostPanel = asPanel.Host as Wall;
                        levelIdPanel = hostPanel.LevelId;
                        if (levelPanelParam != null)
                        {
                            levelPanelParam.Set(GetLevelNumber(element, levelIdPanel, levelPanelParam));
                            _countElements++;
                        }
                    }
                    else if (asPanelWall != null)
                    {
                        asPanelWall.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                        var levelPanelWallParam = asPanelWall.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                        levelIdPanel = asPanelWall.LevelId;
                        if (levelPanelWallParam != null)
                        {
                            levelPanelWallParam.Set(GetLevelNumber(element, levelIdPanel, levelPanelWallParam));
                            _countElements++;
                        }
                    }
                    else
                    {
                        asPanelInstance.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                        var levelPanelInstanceParam = asPanelInstance.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                        levelIdPanel = asPanelWall.LevelId;
                        if (levelPanelInstanceParam != null)
                        {
                            levelPanelInstanceParam.Set(GetLevelNumber(element, levelIdPanel, levelPanelInstanceParam));
                            _countElements++;
                        }
                    }
                    break;

                case BuiltInCategory.OST_Ceilings:
                    var asCeiling = element as Ceiling;
                    asCeiling.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                    var levelCeilingParam = asCeiling.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                    var levelIdCeiling = asCeiling.LevelId;
                    if (levelCeilingParam != null)
                    {
                        levelCeilingParam.Set(GetLevelNumber(element, levelIdCeiling, levelCeilingParam));
                        _countElements++;
                    }
                    break;

                case BuiltInCategory.OST_StairsRailing:
                    var asRailing = element as Railing;
                    asRailing.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                    var levelRailingParam = asRailing.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                    ElementId levelIdRailing = asRailing.LevelId;

                    if (levelIdRailing == new ElementId(-1))
                    {
                        var hostStairs = _doc.GetElement(asRailing.HostId) as Stairs;
                        levelIdRailing = hostStairs.get_Parameter(BuiltInParameter.STAIRS_BASE_LEVEL_PARAM).AsElementId();
                        _countElements++;
                    }

                    if (levelRailingParam != null)
                    {
                        levelRailingParam.Set(GetLevelNumber(element, levelIdRailing, levelRailingParam));
                        _countElements++;
                    }
                    break;

                case BuiltInCategory.OST_Stairs:
                    try
                    {
                        var asStairs = element as Stairs;
                        var asFamInstance = element as FamilyInstance;

                        if(asStairs != null)
                        {
                            asStairs.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                            var levelStairsParam = asStairs.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                            var levelIdStairs = asStairs.get_Parameter(BuiltInParameter.STAIRS_BASE_LEVEL_PARAM).AsElementId();
                            if (levelStairsParam != null)
                            {
                                levelStairsParam.Set(GetLevelNumber(element, levelIdStairs, levelStairsParam));
                                _countElements++;
                            }
                        }

                        if(asFamInstance != null)
                        {
                            asFamInstance.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                            var levelStairsParam = asFamInstance.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                            var levelIdStairs = asFamInstance.LevelId;

                            if(levelIdStairs == null)
                            {
                                var host = asFamInstance.Host;
                                if(host != null)
                                {
                                    var hostAsFloor = host as Floor;
                                    if (hostAsFloor != null) levelIdStairs = hostAsFloor.LevelId;
                                }
                                else
                                {
                                    var boundingBoxMinHeightMark = asFamInstance.get_BoundingBox(null).Min.Z;
                                    if(levelsList != null && boundingBoxMinHeightMark != null)
                                    {
                                        double levelFirstHeightMark;
                                        double levelSecondHeightMark;

                                        for (int i = 0; i<= levelsList.Count - 1; i++)
                                        {
                                            if(i > 0)
                                            {
                                                levelFirstHeightMark = levelsList[i - 1].Elevation;
                                                levelSecondHeightMark = levelsList[i].Elevation;
                                                double differenceLevelsHeightMark = (levelFirstHeightMark + levelSecondHeightMark) / 2;

                                                if(boundingBoxMinHeightMark < levelSecondHeightMark && boundingBoxMinHeightMark > levelFirstHeightMark)
                                                {
                                                    if(boundingBoxMinHeightMark < differenceLevelsHeightMark)
                                                    {
                                                        levelIdStairs = levelsList[i - 1].Id;
                                                        break;
                                                    }

                                                    else
                                                    {
                                                        levelIdStairs = levelsList[i].Id;
                                                        break;
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (levelStairsParam != null && levelIdStairs != null)
                            {
                                levelStairsParam.Set(GetLevelNumber(element, levelIdStairs, levelStairsParam));
                                _countElements++;
                            }
                        }

                        break;
                    }
                    catch
                    {
                        break;
                    }

                case BuiltInCategory.OST_Rebar:
                    var asRebar = element as Rebar;
                    var asRebarInSystem = element as RebarInSystem;
                    var asFamilyInstance = element as FamilyInstance;
                    var levelIdRebar = new ElementId(0);
                    if (asRebar != null)
                    {
                        asRebar.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                        var levelRebarParam = asRebar.LookupParameter(RevitParametersBuilder.FloorNumber.Name);

                        string levelRebarParamValue = "";

                        string levelRebarParamAsString = levelRebarParam.AsString();
                        string levelRebarParamAsValueString = levelRebarParam.AsValueString();

                        if (levelRebarParamAsString != null && levelRebarParamAsString != "") levelRebarParamValue = levelRebarParamAsString;
                        if (levelRebarParamAsValueString != null && levelRebarParamAsValueString != "") levelRebarParamValue = levelRebarParamAsValueString;

                        if(levelRebarParamValue != null)
                        {
                            if (!levelRebarParamValue.Contains('-'))
                            {
                                var host = _doc.GetElement(asRebar.GetHostId());

                                var hostAsFloor = host as Floor;
                                var hostAsWall = host as Wall;
                                var hostAsFamilyInstance = host as FamilyInstance;

                                if (hostAsFloor != null) levelIdRebar = hostAsFloor.LevelId;

                                if (hostAsWall != null) levelIdRebar = hostAsWall.LevelId;

                                if (hostAsFamilyInstance != null)
                                {
                                    levelIdRebar = hostAsFamilyInstance.LevelId;

                                    if(levelIdRebar == null || levelIdRebar == new ElementId(0))
                                    {
                                        var boundingBoxMinHeightMark = hostAsFamilyInstance.get_BoundingBox(null).Min.Z;
                                        if (levelsList != null && boundingBoxMinHeightMark != null)
                                        {
                                            double levelFirstHeightMark;
                                            double levelSecondHeightMark;

                                            for (int i = 0; i <= levelsList.Count - 1; i++)
                                            {
                                                if (i > 0)
                                                {
                                                    levelFirstHeightMark = levelsList[i - 1].Elevation;
                                                    levelSecondHeightMark = levelsList[i].Elevation;
                                                    double differenceLevelsHeightMark = (levelFirstHeightMark + levelSecondHeightMark) / 2;

                                                    if (boundingBoxMinHeightMark < levelSecondHeightMark && boundingBoxMinHeightMark > levelFirstHeightMark)
                                                    {
                                                        if (boundingBoxMinHeightMark < differenceLevelsHeightMark)
                                                        {
                                                            levelIdRebar = levelsList[i - 1].Id;
                                                            break;
                                                        }

                                                        else
                                                        {
                                                            levelIdRebar = levelsList[i].Id;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                if (levelRebarParam != null)
                                {
                                    levelRebarParam.Set(GetLevelNumber(element, levelIdRebar, levelRebarParam));
                                    _countElements++;
                                }
                            }
                        }
                    }
                    else if (asRebarInSystem != null)
                    {
                        asRebarInSystem.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                        var levelRebarInSystemParam = asRebarInSystem.LookupParameter(RevitParametersBuilder.FloorNumber.Name);

                        string levelRebarParamValue = "";

                        string levelRebarParamAsString = levelRebarInSystemParam.AsString();
                        string levelRebarParamAsValueString = levelRebarInSystemParam.AsValueString();

                        if (levelRebarParamAsString != null && levelRebarParamAsString != "") levelRebarParamValue = levelRebarParamAsString;
                        if (levelRebarParamAsValueString != null && levelRebarParamAsValueString != "") levelRebarParamValue = levelRebarParamAsValueString;

                        if(levelRebarParamValue != null)
                        {
                            if (!levelRebarParamValue.Contains('-'))
                            {
                                var host = _doc.GetElement(asRebarInSystem.GetHostId());

                                var hostAsFloor = host as Floor;
                                var hostAsWall = host as Wall;
                                var hostAsFamilyInstance = host as FamilyInstance;

                                if (hostAsFloor != null) levelIdRebar = hostAsFloor.LevelId;

                                if (hostAsWall != null) levelIdRebar = hostAsWall.LevelId;

                                if (hostAsFamilyInstance != null)
                                {
                                    levelIdRebar = hostAsFamilyInstance.LevelId;

                                    if (levelIdRebar == null || levelIdRebar == new ElementId(0))
                                    {
                                        var boundingBoxMinHeightMark = hostAsFamilyInstance.get_BoundingBox(null).Min.Z;
                                        if (levelsList != null && boundingBoxMinHeightMark != null)
                                        {
                                            double levelFirstHeightMark;
                                            double levelSecondHeightMark;

                                            for (int i = 0; i <= levelsList.Count - 1; i++)
                                            {
                                                if (i > 0)
                                                {
                                                    levelFirstHeightMark = levelsList[i - 1].Elevation;
                                                    levelSecondHeightMark = levelsList[i].Elevation;
                                                    double differenceLevelsHeightMark = (levelFirstHeightMark + levelSecondHeightMark) / 2;

                                                    if (boundingBoxMinHeightMark < levelSecondHeightMark && boundingBoxMinHeightMark > levelFirstHeightMark)
                                                    {
                                                        if (boundingBoxMinHeightMark < differenceLevelsHeightMark)
                                                        {
                                                            levelIdRebar = levelsList[i - 1].Id;
                                                            break;
                                                        }

                                                        else
                                                        {
                                                            levelIdRebar = levelsList[i].Id;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                if (levelRebarInSystemParam != null && levelIdRebar != null)
                                {
                                    levelRebarInSystemParam.Set(GetLevelNumber(element, levelIdRebar, levelRebarInSystemParam));
                                    _countElements++;
                                }
                            }
                        }
                    }

                    else if (asFamilyInstance != null)
                    {
                        asFamilyInstance.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                        var levelFamilyInstanceParam = asFamilyInstance.LookupParameter(RevitParametersBuilder.FloorNumber.Name);

                        string levelRebarParamValue = "";

                        string levelRebarParamAsString = levelFamilyInstanceParam.AsString();
                        string levelRebarParamAsValueString = levelFamilyInstanceParam.AsValueString();

                        if (levelRebarParamAsString != null && levelRebarParamAsString != "") levelRebarParamValue = levelRebarParamAsString;
                        if (levelRebarParamAsValueString != null && levelRebarParamAsValueString != "") levelRebarParamValue = levelRebarParamAsValueString;

                        if(levelRebarParamValue != null)
                        {
                            if (!levelRebarParamValue.Contains('-'))
                            {
                                var hostForInstance = asFamilyInstance.Host;

                                if (hostForInstance == null)
                                {
                                    var superComponent = asFamilyInstance.SuperComponent as FamilyInstance;
                                    hostForInstance = superComponent.Host;

                                    if (hostForInstance == null)
                                    {
                                        var boundingBoxMinHeightMark = asFamilyInstance.get_BoundingBox(null).Min.Z;
                                        if (levelsList != null && boundingBoxMinHeightMark != null)
                                        {
                                            double levelFirstHeightMark;
                                            double levelSecondHeightMark;

                                            for (int i = 0; i <= levelsList.Count - 1; i++)
                                            {
                                                if (i > 0)
                                                {
                                                    levelFirstHeightMark = levelsList[i - 1].Elevation;
                                                    levelSecondHeightMark = levelsList[i].Elevation;
                                                    double differenceLevelsHeightMark = (levelFirstHeightMark + levelSecondHeightMark) / 2;

                                                    if (boundingBoxMinHeightMark < levelSecondHeightMark && boundingBoxMinHeightMark > levelFirstHeightMark)
                                                    {
                                                        if (boundingBoxMinHeightMark < differenceLevelsHeightMark)
                                                        {
                                                            levelIdRebar = levelsList[i - 1].Id;
                                                            break;
                                                        }

                                                        else
                                                        {
                                                            levelIdRebar = levelsList[i].Id;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                else
                                {
                                    var hostAsFloor = hostForInstance as Floor;
                                    var hostAsWall = hostForInstance as Wall;
                                    var hostAsInstance = hostForInstance as FamilyInstance;
                                    var hostAsStairs = hostForInstance as Stairs;
                                    var hostAsLevel = hostForInstance as Level;

                                    if (hostAsFloor != null)
                                    {
                                        levelIdRebar = hostAsFloor.LevelId;
                                    }

                                    else if (hostAsWall != null)
                                    {
                                        levelIdRebar = hostAsWall.LevelId;
                                    }

                                    else if (hostAsInstance != null)
                                    {
                                        levelIdRebar = hostAsInstance.LevelId;

                                        if (levelIdRebar == null || levelIdRebar == new ElementId(0))
                                        {
                                            var boundingBoxMinHeightMark = hostAsInstance.get_BoundingBox(null).Min.Z;
                                            if (levelsList != null && boundingBoxMinHeightMark != null)
                                            {
                                                double levelFirstHeightMark;
                                                double levelSecondHeightMark;

                                                for (int i = 0; i <= levelsList.Count - 1; i++)
                                                {
                                                    if (i > 0)
                                                    {
                                                        levelFirstHeightMark = levelsList[i - 1].Elevation;
                                                        levelSecondHeightMark = levelsList[i].Elevation;
                                                        double differenceLevelsHeightMark = (levelFirstHeightMark + levelSecondHeightMark) / 2;

                                                        if (boundingBoxMinHeightMark < levelSecondHeightMark && boundingBoxMinHeightMark > levelFirstHeightMark)
                                                        {
                                                            if (boundingBoxMinHeightMark < differenceLevelsHeightMark)
                                                            {
                                                                levelIdRebar = levelsList[i - 1].Id;
                                                                break;
                                                            }

                                                            else
                                                            {
                                                                levelIdRebar = levelsList[i].Id;
                                                                break;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    else if (hostAsStairs != null)
                                    {
                                        levelIdRebar = hostAsStairs.get_Parameter(BuiltInParameter.STAIRS_BASE_LEVEL_PARAM).AsElementId();
                                    }

                                    else if (hostAsLevel != null)
                                    {
                                        levelIdRebar = hostAsLevel.Id;
                                    }
                                }

                                if (levelFamilyInstanceParam != null)
                                {
                                    levelFamilyInstanceParam.Set(GetLevelNumber(element, levelIdRebar, levelFamilyInstanceParam));
                                    _countElements++;
                                }
                            }
                        }
                    }
                    break;

                case BuiltInCategory.OST_PipeCurves:
                    var asPipes = element as Pipe;

                    asPipes.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);

                    var levelPipesParam = asPipes.LookupParameter(RevitParametersBuilder.FloorNumber.Name);

                    var levelIdPipes = asPipes.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();

                    if (levelPipesParam != null && levelPipesParam.AsString() != "Секция" || levelPipesParam.AsValueString() != "Секция")
                    {
                        levelPipesParam.Set(GetLevelNumber(element, levelIdPipes, levelPipesParam));
                        _countElements++;
                    }
                    break;

                case BuiltInCategory.OST_PipeInsulations:
                    var asPipeInsulation = element as PipeInsulation;

                    asPipeInsulation.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);

                    var levelPipeInsulationParam = asPipeInsulation.LookupParameter(RevitParametersBuilder.FloorNumber.Name);

                    var levelIdPipeInsulation = _doc.GetElement(asPipeInsulation.HostElementId).get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();

                    if (levelPipeInsulationParam != null && levelPipeInsulationParam.AsString() != "Секция" || levelPipeInsulationParam.AsString() != "Секция")
                    {
                        levelPipeInsulationParam.Set(GetLevelNumber(element, levelIdPipeInsulation, levelPipeInsulationParam));
                        _countElements++;
                    }
                    break;

                case BuiltInCategory.OST_DuctCurves:
                    var asDuct = element as Duct;
                    asDuct.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                    var levelDuctParam = asDuct.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                    var levelIdDuct = asDuct.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    if (levelDuctParam != null && levelDuctParam.AsString() != "Секция" || levelDuctParam.AsString() != "Секция" || levelDuctParam.AsString() == "" || levelDuctParam.AsString() == "")
                    {
                        levelDuctParam.Set(GetLevelNumber(element, levelIdDuct, levelDuctParam));
                        _countElements++;
                    }
                    break;

                case BuiltInCategory.OST_DuctInsulations:
                    var asDuctInsulation = element as DuctInsulation;
                    asDuctInsulation.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                    var levelDuctInsulationParam = asDuctInsulation.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                    var levelIdDuctInsulation = _doc.GetElement(asDuctInsulation.HostElementId).get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    if (levelDuctInsulationParam != null && levelDuctInsulationParam.AsString() != "Секция" || levelDuctInsulationParam.AsString() != "секция")
                    {
                        levelDuctInsulationParam.Set(GetLevelNumber(element, levelIdDuctInsulation, levelDuctInsulationParam));
                        _countElements++;
                    }
                    break;


                default:
                    try
                    {
                        var asInstance = element as FamilyInstance;
                        asInstance.LookupParameter(RevitParametersBuilder.SectionNumber.Name).Set(sectionNumber);
                        var levelInstanceParam = asInstance.LookupParameter(RevitParametersBuilder.FloorNumber.Name);
                        var levelIdInstance = new ElementId(0);

                        try
                        {
                            levelIdInstance = asInstance.LevelId;
                        }
                        catch
                        {
                            levelIdInstance = asInstance.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM).AsElementId();
                        }

                        if (levelIdInstance == new ElementId(-1) || levelIdInstance == new ElementId(0) || levelIdInstance == null)
                        {
                            var host = asInstance.Host;
                            var superComponent = asInstance.SuperComponent;

                            if(host == null)
                            {
                                if(superComponent != null)
                                {
                                    var superComponentAsInstance = superComponent as FamilyInstance;

                                    if (superComponentAsInstance != null)
                                    {
                                        host = superComponentAsInstance.Host;
                                    }
                                }
                            }

                            if(host != null)
                            {
                                var asHostWall = host as Wall;
                                var asHostWallSweep = host as WallSweep;
                                var asHostFloor = host as Floor;
                                var asLevel = host as Level;
                                var asPipe = host as Pipe;

                                if (asHostWall != null) levelIdInstance = asHostWall.LevelId;

                                if (asHostWallSweep != null) levelIdInstance = asHostWallSweep.LevelId;

                                if (asHostFloor != null) levelIdInstance = asHostFloor.LevelId;

                                if (asLevel != null) levelIdInstance = asLevel.Id;

                                else if (asPipe != null) levelIdInstance = asPipe.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                            }
                        }

                        if (levelInstanceParam != null)
                        {
                            levelInstanceParam.Set(GetLevelNumber(element, levelIdInstance, levelInstanceParam));
                            _countElements++;
                        }
                        break;
                    }
                    catch
                    {
                        _countElements++;
                        break;
                    }
            }
        }

        /*Метод заполнния у несущей арматуры параметра ЦДС_Классификатор
         */
        public void SetClassifierParameterValueToRebar(ICollection<Element> rebarElements)
        {
            var allPartsInDocument = new FilteredElementCollector(_doc).OfClass(typeof(Part)).WhereElementIsNotElementType().ToElements();

            Dictionary<ElementId, ElementId> partsClassifierDictionary = new Dictionary<ElementId, ElementId>();

            foreach(Element element in allPartsInDocument)
            {
                Part part = element as Part;
                if(part != null)
                {
                    ElementId elementIdPart = part.Id;
                    ElementId elementIdHost = part.GetSourceElementIds().ToList()[0].HostElementId;
                    if (!partsClassifierDictionary.Keys.Contains(elementIdHost))
                    {
                        partsClassifierDictionary[elementIdHost] = elementIdPart;
                    }
                }
            }

            foreach (Element elem in rebarElements)
            {
                Rebar asRebar = elem as Rebar;
                RebarInSystem asRebarInSystem = elem as RebarInSystem;
                FamilyInstance asFamInstance = elem as FamilyInstance;

                if (asRebar != null)
                {
                    var parameterClassifierRebarForSet = asRebar.LookupParameter(RevitParametersBuilder.Classifier.Name);

                    if (parameterClassifierRebarForSet != null)
                    {
                        var host = _doc.GetElement(asRebar.GetHostId());

                        var hostAsWall = host as Wall;
                        var hostAsFloor = host as Autodesk.Revit.DB.Floor;
                        var hostAsInstance = host as FamilyInstance;

                        if (hostAsWall != null)
                        {
                            ElementId idFromWall = hostAsWall.Id;

                            string classifier = ((Part)_doc.GetElement(partsClassifierDictionary[idFromWall])).LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsString();
                            if (classifier == null || classifier == "") classifier = ((Part)_doc.GetElement(partsClassifierDictionary[idFromWall])).LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsValueString();

                            parameterClassifierRebarForSet.Set(classifier);
                        }

                        if (hostAsFloor != null)
                        {
                            ElementId idFromFloor = hostAsFloor.Id;

                            string classifier = ((Part)_doc.GetElement(partsClassifierDictionary[idFromFloor])).LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsString();
                            if (classifier == null || classifier == "") classifier = ((Part)_doc.GetElement(partsClassifierDictionary[idFromFloor])).LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsValueString();

                            parameterClassifierRebarForSet.Set(classifier);
                        }

                        if (hostAsInstance != null)
                        {
                            var classifierAsString = hostAsInstance.LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsString();
                            var classifierAsValueString = hostAsInstance.LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsValueString();

                            string classifier = "";

                            if (classifierAsString != null && classifierAsString != "") classifier = classifierAsString;
                            if (classifierAsValueString != null && classifierAsValueString != "") classifier = classifierAsString;

                            if (classifier == null || classifier == "")
                            {
                                ElementId idFromInstance = hostAsInstance.Id;

                                classifier = ((Part)_doc.GetElement(partsClassifierDictionary[idFromInstance])).LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsString();
                                if (classifier == null || classifier == "") classifier = ((Part)_doc.GetElement(partsClassifierDictionary[idFromInstance])).LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsValueString();
                            }

                            parameterClassifierRebarForSet.Set(classifier);
                        }
                    }
                }
                else if (asRebarInSystem != null)
                {
                    var parameterClassifierRebarForSet = asRebarInSystem.LookupParameter(RevitParametersBuilder.Classifier.Name);

                    if (parameterClassifierRebarForSet != null)
                    {
                        var host = _doc.GetElement(asRebarInSystem.GetHostId());

                        var hostAsWall = host as Wall;
                        var hostAsFloor = host as Autodesk.Revit.DB.Floor;
                        var hostAsInstance = host as FamilyInstance;

                        if (hostAsWall != null)
                        {
                            ElementId idFromWall = hostAsWall.Id;

                            string classifier = ((Part)_doc.GetElement(partsClassifierDictionary[idFromWall])).LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsString();
                            if (classifier == null || classifier == "") classifier = ((Part)_doc.GetElement(partsClassifierDictionary[idFromWall])).LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsValueString();

                            parameterClassifierRebarForSet.Set(classifier);
                        }

                        if (hostAsFloor != null)
                        {
                            ElementId idFromFloor = hostAsFloor.Id;

                            string classifier = ((Part)_doc.GetElement(partsClassifierDictionary[idFromFloor])).LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsString();
                            if (classifier == null || classifier == "") classifier = ((Part)_doc.GetElement(partsClassifierDictionary[idFromFloor])).LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsValueString();

                            parameterClassifierRebarForSet.Set(classifier);
                        }

                        if (hostAsInstance != null)
                        {
                            var classifierAsString = hostAsInstance.LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsString();
                            var classifierAsValueString = hostAsInstance.LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsValueString();

                            if (classifierAsString != null && classifierAsString != "") parameterClassifierRebarForSet.Set(classifierAsString);
                            if (classifierAsValueString != null && classifierAsValueString != "") parameterClassifierRebarForSet.Set(classifierAsValueString);
                        }
                    }
                }

                else if (asFamInstance != null)
                {
                    var asSymbol = asFamInstance.Symbol;
                    string typeCommentParameterValue = asSymbol.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString();

                    
                    if (typeCommentParameterValue != EMBEDDED_PARTS)
                    {
                        var parameterClassifierRebarForSet = asFamInstance.LookupParameter(RevitParametersBuilder.Classifier.Name);

                        if (parameterClassifierRebarForSet != null)
                        {
                            var hostForInstance = asFamInstance.Host;

                            if (hostForInstance == null)
                            {
                                var superComponent = asFamInstance.SuperComponent as FamilyInstance;
                                hostForInstance = superComponent?.Host;
                            }

                            if(hostForInstance != null)
                            {
                                var hostAsFloor = hostForInstance as Autodesk.Revit.DB.Floor;
                                var hostAsWall = hostForInstance as Wall;
                                var hostAsInstance = hostForInstance as FamilyInstance;
                                var hostAsStairs = hostForInstance as Stairs;
                                var hostAsLevel = hostForInstance as Level;

                                if (hostAsFloor != null)
                                {
                                    ElementId idFromFloor = hostAsFloor.Id;

                                    string classifier = ((Part)_doc.GetElement(partsClassifierDictionary[idFromFloor])).LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsString();
                                    if (classifier == null || classifier == "") classifier = ((Part)_doc.GetElement(partsClassifierDictionary[idFromFloor])).LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsValueString();

                                    parameterClassifierRebarForSet.Set(classifier);
                                }

                                else if (hostAsWall != null)
                                {
                                    ElementId idFromWall = hostAsWall.Id;

                                    string classifier = ((Part)_doc.GetElement(partsClassifierDictionary[idFromWall])).LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsString();
                                    if (classifier == null || classifier == "") classifier = ((Part)_doc.GetElement(partsClassifierDictionary[idFromWall])).LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsValueString();

                                    parameterClassifierRebarForSet.Set(classifier);
                                }

                                else if (hostAsInstance != null)
                                {
                                    var classifierAsString = hostAsInstance.LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsString();
                                    var classifierAsValueString = hostAsInstance.LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsValueString();

                                    if (classifierAsString != null && classifierAsString != "") parameterClassifierRebarForSet.Set(classifierAsString);
                                    if (classifierAsValueString != null && classifierAsValueString != "") parameterClassifierRebarForSet.Set(classifierAsValueString);
                                }

                                else if (hostAsStairs != null)
                                {
                                    var classifierAsString = hostAsStairs.LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsString();
                                    var classifierAsValueString = hostAsStairs.LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsValueString();

                                    if (classifierAsString != null && classifierAsString != "") parameterClassifierRebarForSet.Set(classifierAsString);
                                    if (classifierAsValueString != null && classifierAsValueString != "") parameterClassifierRebarForSet.Set(classifierAsValueString);
                                }

                                else if (hostAsLevel != null)
                                {
                                    var intersectFilter = new ElementIntersectsElementFilter(asFamInstance);
                                    var intersectElements = new FilteredElementCollector(_doc).WherePasses(intersectFilter).WhereElementIsNotElementType().ToElements();

                                    var intersectFirstPartElement = (from e in intersectElements
                                                                     where e as Part != null
                                                                     select e).FirstOrDefault() as Part;
                                    if (intersectFirstPartElement != null)
                                    {
                                        var classifierAsString = intersectFirstPartElement.LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsString();
                                        var classifierAsValueString = intersectFirstPartElement.LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsValueString();

                                        if (classifierAsString != null && classifierAsString != "") parameterClassifierRebarForSet.Set(classifierAsString);
                                        if (classifierAsValueString != null && classifierAsValueString != "") parameterClassifierRebarForSet.Set(classifierAsValueString);
                                    }
                                    else
                                    {
                                        FamilyInstance intersectFirstFamilyInstance = (from e in intersectElements
                                                                                       where e as FamilyInstance != null
                                                                                       select e).FirstOrDefault() as FamilyInstance;
                                        if(intersectFirstFamilyInstance != null)
                                        {
                                            var classifierAsString = intersectFirstFamilyInstance.LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsString();
                                            var classifierAsValueString = intersectFirstFamilyInstance.LookupParameter(RevitParametersBuilder.Classifier.Name)?.AsValueString();

                                            if (classifierAsString != null && classifierAsString != "") parameterClassifierRebarForSet.Set(classifierAsString);
                                            if (classifierAsValueString != null && classifierAsValueString != "") parameterClassifierRebarForSet.Set(classifierAsValueString);
                                        }
                                    }
                                }
                            }
                        }
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

        /*Метод получения 3D вида для экспорта модели в формат .nwc
         */
        public View3D Get3DViewForExportToNWC()
        {
            var allViews = new FilteredElementCollector(_doc).OfClass(typeof(View3D)).WhereElementIsNotElementType().ToElements();
            View3D view3DForExport = null;
            foreach (Element elementView in allViews)
            {
                View3D view3D = elementView as View3D;

                string viewName = view3D.Name;
                if (viewName.Split('_')[0].Contains("xapta") && viewName.Split('_')[1].Contains("avisworks"))
                {
                    view3DForExport = view3D;
                    break;
                }
            }
            return view3DForExport;
        }

        public List<ElementId> Get3DViewForSetSectionParameter()
        {
            var allViews = new FilteredElementCollector(_doc).OfClass(typeof(View3D)).WhereElementIsNotElementType().ToElements();
            List<ElementId> view3DForExport = new List<ElementId>();
            foreach (Element elementView in allViews)
            {
                View3D view3D = elementView as View3D;

                string viewName = view3D.Name;
                if (viewName.Split('_')[0].Contains("xapta") && viewName.Split('_')[1].Contains("екция"))
                {
                    if (!view3DForExport.Contains(view3D.Id))
                    {
                        view3DForExport.Add(view3D.Id);
                    }
                }
            }
            return view3DForExport;
        }
    }
}