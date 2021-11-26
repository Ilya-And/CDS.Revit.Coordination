using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CDS.Revit.Coordination.Axapta.Services.Revit
{
    public class RevitModelElementService
    {
        private Document _doc;
        private int _countElements = 0;
        public string SectionParameter = "ADSK_Номер секции";
        public string LevelParameter = "ADSK_Этаж";
        public string ClassifierParameter = "ЦДС_Классификатор";
        public string ClassMaterialParameter = "ЦДС_Классификатор_Материалов";
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

        /*Метод заполнния у элемента параметров ADSK_Номер секции и ADSK_Этаж
         */
        public void SetParametersValuesToElement(Element element, string sectionNumber)
        {
            var category = element.Category;

            switch ((BuiltInCategory)category.Id.IntegerValue)
            {
                case BuiltInCategory.OST_Walls:
                    var asWall = element as Wall;
                    asWall.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                    var levelWallParam = asWall.LookupParameter("ADSK_Этаж");
                    var levelIdWall = asWall.LevelId;
                    if (levelWallParam != null && levelWallParam.AsValueString() != "Cекция" || levelWallParam.AsValueString() != "секция")
                    {
                        levelWallParam.Set(GetLevelNumber(element, levelIdWall, levelWallParam));
                        _countElements++;
                    }
                    break;

                case BuiltInCategory.OST_Cornices:
                    var asWallSweep = element as WallSweep;
                    asWallSweep.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                    var levelWallSweepParam = asWallSweep.LookupParameter("ADSK_Этаж");
                    var levelIdWallSweep = asWallSweep.LevelId;
                    if (levelWallSweepParam != null)
                    {
                        asWallSweep.LookupParameter("ADSK_Этаж").Set(GetLevelNumber(element, levelIdWallSweep, levelWallSweepParam));
                        _countElements++;
                    }
                    break;

                case BuiltInCategory.OST_Floors:
                    var asFloor = element as Floor;
                    asFloor.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                    var levelFloorParam = asFloor.LookupParameter("ADSK_Этаж");
                    var levelIdFloor = asFloor.LevelId;
                    if (levelFloorParam != null)
                    {
                        levelFloorParam.Set(GetLevelNumber(element, levelIdFloor, levelFloorParam));
                        _countElements++;
                    }
                    break;

                case BuiltInCategory.OST_StructuralFoundation:
                    var asFloorFoundation = element as Floor;
                    var asFamilyInstanceFoundation = element as FamilyInstance;
                    if (asFloorFoundation != null)
                    {
                        asFloorFoundation.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                        var levelFloorFoundationParam = asFloorFoundation.LookupParameter("ADSK_Этаж");
                        var levelIdFloorFoundation = asFloorFoundation.LevelId;
                        if (levelFloorFoundationParam != null)
                        {
                            levelFloorFoundationParam.Set(GetLevelNumber(element, levelIdFloorFoundation, levelFloorFoundationParam));
                            _countElements++;
                        }
                    }
                    else if (asFamilyInstanceFoundation != null)
                    {
                        asFamilyInstanceFoundation.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                        var levelInstanceFoundationParam = asFamilyInstanceFoundation.LookupParameter("ADSK_Этаж");
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
                        asRoof.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                        var levelRoofParam = asRoof.LookupParameter("ADSK_Этаж");
                        var levelIdRoof = asRoof.LevelId;
                        if (levelRoofParam != null)
                        {
                            levelRoofParam.Set(GetLevelNumber(element, levelIdRoof, levelRoofParam));
                            _countElements++;
                        }
                    }
                    else if (asExtrusionRoof != null)
                    {
                        asExtrusionRoof.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                        var levelExtrusionRoofParam = asExtrusionRoof.LookupParameter("ADSK_Этаж");
                        var levelIdRoof = asExtrusionRoof.get_Parameter(BuiltInParameter.ROOF_CONSTRAINT_LEVEL_PARAM).AsElementId();
                        if (levelExtrusionRoofParam != null)
                        {
                            asExtrusionRoof.LookupParameter("ADSK_Этаж").Set(GetLevelNumber(element, levelIdRoof, levelExtrusionRoofParam));
                            _countElements++;
                        }
                    }
                    break;

                case BuiltInCategory.OST_CurtainWallMullions:
                    var asMullion = element as Mullion;
                    asMullion.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                    var levelMullionParam = asMullion.LookupParameter("ADSK_Этаж");
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
                        asPanel.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                        var levelPanelParam = asPanel.LookupParameter("ADSK_Этаж");
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
                        asPanelWall.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                        var levelPanelWallParam = asPanelWall.LookupParameter("ADSK_Этаж");
                        levelIdPanel = asPanelWall.LevelId;
                        if (levelPanelWallParam != null)
                        {
                            levelPanelWallParam.Set(GetLevelNumber(element, levelIdPanel, levelPanelWallParam));
                            _countElements++;
                        }
                    }
                    else
                    {
                        asPanelInstance.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                        var levelPanelInstanceParam = asPanelInstance.LookupParameter("ADSK_Этаж");
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
                    asCeiling.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                    var levelCeilingParam = asCeiling.LookupParameter("ADSK_Этаж");
                    var levelIdCeiling = asCeiling.LevelId;
                    if (levelCeilingParam != null)
                    {
                        levelCeilingParam.Set(GetLevelNumber(element, levelIdCeiling, levelCeilingParam));
                        _countElements++;
                    }
                    break;

                case BuiltInCategory.OST_StairsRailing:
                    var asRailing = element as Railing;
                    asRailing.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                    var levelRailingParam = asRailing.LookupParameter("ADSK_Этаж");
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
                        asStairs.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                        var levelStairsParam = asStairs.LookupParameter("ADSK_Этаж");
                        var levelIdStairs = asStairs.get_Parameter(BuiltInParameter.STAIRS_BASE_LEVEL_PARAM).AsElementId();
                        if (levelStairsParam != null)
                        {
                            levelStairsParam.Set(GetLevelNumber(element, levelIdStairs, levelStairsParam));
                            _countElements++;
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
                        asRebar.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                        var levelRebarParam = asRebar.LookupParameter("ADSK_Этаж");
                        try
                        {
                            levelIdRebar = _doc.GetElement(asRebar.GetHostId()).LevelId;
                        }
                        catch
                        {
                            levelIdRebar = _doc.GetElement(asRebar.GetHostId()).get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM).AsElementId();
                        }
                        if (levelRebarParam != null)
                        {
                            levelRebarParam.Set(GetLevelNumber(element, levelIdRebar, levelRebarParam));
                            _countElements++;
                        }
                    }
                    else if (asRebarInSystem != null)
                    {
                        asRebarInSystem.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                        var levelRebarInSystemParam = asRebarInSystem.LookupParameter("ADSK_Этаж");
                        try
                        {
                            levelIdRebar = _doc.GetElement(asRebarInSystem.GetHostId()).LevelId;
                        }
                        catch
                        {
                            levelIdRebar = _doc.GetElement(asRebarInSystem.GetHostId()).get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM).AsElementId();
                        }
                        if (levelRebarInSystemParam != null)
                        {
                            levelRebarInSystemParam.Set(GetLevelNumber(element, levelIdRebar, levelRebarInSystemParam));
                            _countElements++;
                        }
                    }

                    else if (asFamilyInstance != null)
                    {
                        asFamilyInstance.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                        var levelFamilyInstanceParam = asFamilyInstance.LookupParameter("ADSK_Этаж");
                        var hostForInstance = asFamilyInstance.Host;
                        if (hostForInstance == null)
                        {
                            var superComponent = asFamilyInstance.SuperComponent as FamilyInstance;
                            hostForInstance = superComponent.Host;
                        }
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
                        }
                        else if (hostAsStairs != null)
                        {
                            levelIdRebar = hostAsStairs.get_Parameter(BuiltInParameter.STAIRS_BASE_LEVEL_PARAM).AsElementId();
                        }
                        else if (hostAsLevel != null)
                        {
                            levelIdRebar = hostAsLevel.Id;
                        }

                        if (levelFamilyInstanceParam != null)
                        {
                            levelFamilyInstanceParam.Set(GetLevelNumber(element, levelIdRebar, levelFamilyInstanceParam));
                            _countElements++;
                        }
                    }
                    break;

                case BuiltInCategory.OST_PipeCurves:
                    var asPipes = element as Pipe;
                    asPipes.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                    var levelPipesParam = asPipes.LookupParameter("ADSK_Этаж");
                    var levelIdPipes = asPipes.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();

                    if (levelPipesParam != null && levelPipesParam.AsString() != "Секция" || levelPipesParam.AsValueString() != "секция")
                    {
                        levelPipesParam.Set(GetLevelNumber(element, levelIdPipes, levelPipesParam));
                        _countElements++;
                    }
                    break;

                case BuiltInCategory.OST_PipeInsulations:
                    var asPipeInsulation = element as PipeInsulation;
                    asPipeInsulation.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                    var levelPipeInsulationParam = asPipeInsulation.LookupParameter("ADSK_Этаж");
                    var levelIdPipeInsulation = _doc.GetElement(asPipeInsulation.HostElementId).get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    if (levelPipeInsulationParam != null && levelPipeInsulationParam.AsString() != "Секция" || levelPipeInsulationParam.AsString() != "секция")
                    {
                        levelPipeInsulationParam.Set(GetLevelNumber(element, levelIdPipeInsulation, levelPipeInsulationParam));
                        _countElements++;
                    }
                    break;

                case BuiltInCategory.OST_DuctCurves:
                    var asDuct = element as Duct;
                    asDuct.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                    var levelDuctParam = asDuct.LookupParameter("ADSK_Этаж");
                    var levelIdDuct = asDuct.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
                    if (levelDuctParam != null && levelDuctParam.AsString() != "Секция" || levelDuctParam.AsString() != "секция")
                    {
                        levelDuctParam.Set(GetLevelNumber(element, levelIdDuct, levelDuctParam));
                        _countElements++;
                    }
                    break;

                case BuiltInCategory.OST_DuctInsulations:
                    var asDuctInsulation = element as DuctInsulation;
                    asDuctInsulation.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                    var levelDuctInsulationParam = asDuctInsulation.LookupParameter("ADSK_Этаж");
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
                        asInstance.LookupParameter("ADSK_Номер секции").Set(sectionNumber);
                        var levelInstanceParam = asInstance.LookupParameter("ADSK_Этаж");
                        var levelIdInstance = new ElementId(0);
                        try
                        {
                            levelIdInstance = asInstance.LevelId;
                        }
                        catch
                        {

                            levelIdInstance = asInstance.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM).AsElementId();

                        }
                        if (levelIdInstance == new ElementId(-1))
                        {
                            var host = asInstance.Host;
                            var asLevel = host as Level;
                            var asPipe = host as Pipe;
                            if (asLevel != null)
                            {
                                levelIdInstance = asLevel.Id;
                            }
                            else if (asPipe != null)
                            {
                                levelIdInstance = asPipe.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM).AsElementId();
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
            foreach(Element rebarElement in rebarElements)
            {
                var asRebar = rebarElement as Rebar;
                var asRebarInSystem = rebarElement as RebarInSystem;
                var asFamilyInstance = rebarElement as FamilyInstance;

                if(asRebar != null)
                {
                    var host = _doc.GetElement(asRebar.GetHostId());
                    var hostAsPart = host as Part;

                    if(hostAsPart != null)
                    {
                        string classifier = hostAsPart.LookupParameter(ClassifierParameter).AsString();
                        asRebar.LookupParameter(ClassifierParameter).Set(classifier);
                    }
                }

                if (asRebarInSystem != null)
                {
                    var host = _doc.GetElement(asRebarInSystem.GetHostId());
                    var hostAsPart = host as Part;

                    if (hostAsPart != null)
                    {
                        string classifier = hostAsPart.LookupParameter(ClassifierParameter).AsString();
                        asRebarInSystem.LookupParameter(ClassifierParameter).Set(classifier);
                    }
                }

                if (asFamilyInstance != null)
                {
                    var asSymbol = asFamilyInstance.Symbol;
                    string typeCommentParameterValue = asSymbol.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_COMMENTS).AsString();

                    if(typeCommentParameterValue != "детали закладных")
                    {
                        var hostAsPart = asFamilyInstance.Host as Part;

                        if (hostAsPart != null)
                        {
                            string classifier = hostAsPart.LookupParameter(ClassifierParameter).AsString();
                            asFamilyInstance.LookupParameter(ClassifierParameter).Set(classifier);
                        }

                        else
                        {
                            var intersectFilter = new ElementIntersectsElementFilter(asFamilyInstance);
                            var intersectElements = new FilteredElementCollector(_doc).WherePasses(intersectFilter).WhereElementIsNotElementType().ToElements();
                            var intersectFirstPartElement = (from e in intersectElements
                                                            where e as Part != null
                                                            select e).FirstOrDefault() as Part;

                            if (intersectFirstPartElement != null)
                            {
                                string classifier = intersectFirstPartElement.LookupParameter(ClassifierParameter).AsString();
                                asFamilyInstance.LookupParameter(ClassifierParameter).Set(classifier);
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
    }
}