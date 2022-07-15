using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreationHouseModelRevitPlugin
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CreationHouseModel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document document = commandData.Application.ActiveUIDocument.Document;

            List<Level> levelsList = Levels.getLevelsList(document);
            Level level1 = Levels.getLevelByName(levelsList, "Уровень 1");
            Level level2 = Levels.getLevelByName(levelsList, "Уровень 2");

            List<Wall> wallsList = Walls.constructWalls(document, level1, level2, 10000, 5000);

            Doors.constructDoor(document, level1, wallsList[0]);

            for (int i = 1; i < wallsList.Count; i++)
            {
                Windows.constructWindow(document, level1, wallsList[i], 1000);
            }

            Roofs.constructRoof(document, level2, wallsList, 10000, 5000, 3);
            //Roofs.constructRoof(document, level2, wallsList);

            return Result.Succeeded;
        }
    }

    public static class Levels
    {
        public static List<Level> getLevelsList(Document document)
        {
            List<Level> levelsList = new FilteredElementCollector(document)
                .OfClass(typeof(Level))
                .OfType<Level>()
                .ToList();
            return levelsList;
        }
        public static Level getLevelByName(List<Level> levelsList, string levelName)
        {
            Level level = levelsList
               .Where(x => x.Name.Equals(levelName))
               .FirstOrDefault();
            return level;
        }
    }

    public static class Walls
    {
        public static List<Wall> constructWalls(Document document, Level levelBase, Level levelUp, double wallWidth, double wallDepth)
        {
            List<XYZ> wallAnchorPointsList = defineAnchorPointsFromTheOriginForConstructingWalls(wallWidth, wallDepth);
            List<Wall> wallsList = new List<Wall>();
            for (int i = 0; i < wallAnchorPointsList.Count - 1; i++)
            {
                Line constructionLine = Line.CreateBound(wallAnchorPointsList[i], wallAnchorPointsList[i + 1]);
                Wall wall = constructWallTransaction(document, levelBase, levelUp, constructionLine);
                wallsList.Add(wall);
            }
            return wallsList;
        }
        public static List<XYZ> defineAnchorPointsFromTheOriginForConstructingWalls(double wallWidth, double wallDepth)
        {
            double revitWallWidth = UnitUtils.ConvertToInternalUnits(wallWidth, UnitTypeId.Millimeters);
            double revitWallDepth = UnitUtils.ConvertToInternalUnits(wallDepth, UnitTypeId.Millimeters);

            double xIncrement = Increment.getIncrement(revitWallWidth);
            double yIncrement = Increment.getIncrement(revitWallDepth);

            List<XYZ> anchorPointsList = AnchorPoints.getListOfAnchorPoints(xIncrement, yIncrement);
            return anchorPointsList;
        }
        public static Wall constructWallTransaction(Document document, Level levelBase, Level levelUp, Line constructionLine)
        {
            Transaction transaction = new Transaction(document, "Построение стен");
            transaction.Start();
            Wall wall = Wall.Create(document, constructionLine, levelBase.Id, false);
            wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(levelUp.Id);
            transaction.Commit();
            return wall;
        }
    }

    public static class Doors
    {
        private static FamilySymbol getDoorTypeFromProject(Document document)
        {
            FamilySymbol doorType = new FilteredElementCollector(document)
              .OfClass(typeof(FamilySymbol))
              .OfCategory(BuiltInCategory.OST_Doors)
              .OfType<FamilySymbol>()
              .Where(x => x.Name.Equals("0915 x 2134 мм"))
              .Where(x => x.FamilyName.Equals("Одиночные-Щитовые"))
              .FirstOrDefault();
            return doorType;
        }
        public static void constructDoor(Document document, Level level, Wall wall)
        {
            FamilySymbol doorType = getDoorTypeFromProject(document);
            XYZ insertionPoint = HostCurveForObject.getPointForLocateElement(wall);
            constructDoorTransaction(document, level, wall, doorType, insertionPoint);
        }
        private static void constructDoorTransaction(Document document, Level level, Wall wall, FamilySymbol doorType, XYZ insertionPoint)
        {
            Transaction transaction = new Transaction(document, "Построение двери");
            transaction.Start();
            if (!doorType.IsActive) doorType.Activate();
            document.Create.NewFamilyInstance(insertionPoint, doorType, wall, level, StructuralType.NonStructural);
            transaction.Commit();
        }
    }

    public static class Windows
    {
        private static FamilySymbol getWindowTypeFromProject(Document document)
        {
            FamilySymbol windowType = new FilteredElementCollector(document)
               .OfClass(typeof(FamilySymbol))
               .OfCategory(BuiltInCategory.OST_Windows)
               .OfType<FamilySymbol>()
               .Where(x => x.Name.Equals("0915 x 1830 мм"))
               .Where(x => x.FamilyName.Equals("Фиксированные"))
               .FirstOrDefault();
            return windowType;
        }
        public static void constructWindow(Document document, Level level, Wall wall, double floorToWindowHeight)
        {
            FamilySymbol windowType = getWindowTypeFromProject(document);
            XYZ insertionPoint = HostCurveForObject.getPointForLocateElement(wall);
            constructWindowTransaction(document, level, wall, floorToWindowHeight, windowType, insertionPoint);
        }
        private static void constructWindowTransaction(Document document, Level level, Wall wall, double floorToWindowHeight, FamilySymbol windowType, XYZ insertionPoint)
        {
            Transaction transaction = new Transaction(document, "Построение окна");
            transaction.Start();
            if (!windowType.IsActive) windowType.Activate();
            FamilyInstance window = document.Create.NewFamilyInstance(insertionPoint, windowType, wall, level, StructuralType.NonStructural);
            setFloorToWindowHeight(window, floorToWindowHeight);
            transaction.Commit();
        }
        private static void setFloorToWindowHeight(FamilyInstance window, double floorToWindowHeight)
        {
            double revitFloorToWindowHeight = UnitUtils.ConvertToInternalUnits(floorToWindowHeight, UnitTypeId.Millimeters);
            window.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).Set(revitFloorToWindowHeight);
        }
    }

    public static class Roofs
    {
        private static RoofType getRoofTypeFromProject(Document document)
        {
            RoofType roofType = new FilteredElementCollector(document)
               .OfClass(typeof(RoofType))
               .OfType<RoofType>()
               .Where(x => x.Name.Equals("Типовой - 400мм"))
               .Where(x => x.FamilyName.Equals("Базовая крыша"))
               .FirstOrDefault();
            return roofType;
        }
        public static void constructRoof(Document document, Level level, List<Wall> wallsList)
        {
            RoofType roofType = getRoofTypeFromProject(document);
            CurveArray constructionLinesArray = getConstructionLinesArray(document, wallsList);
            constructRoofTransaction(document, level, roofType, constructionLinesArray);
        }
        private static void constructRoofTransaction(Document document, Level level, RoofType roofType, CurveArray constructionLinesArray)
        {
            Transaction transaction = new Transaction(document, "Построение крыши");
            transaction.Start();
            ModelCurveArray footPrintToModelCurveMapping = new ModelCurveArray();
            FootPrintRoof footPrintRoof = document.Create.NewFootPrintRoof(constructionLinesArray, level, roofType, out footPrintToModelCurveMapping);
            foreach (ModelCurve modelCurve in footPrintToModelCurveMapping)
            {
                footPrintRoof.set_DefinesSlope(modelCurve, true);
                footPrintRoof.set_SlopeAngle(modelCurve, 0.5);
            }
            transaction.Commit();
        }
        private static CurveArray getConstructionLinesArray(Document document, List<Wall> wallsList)
        {
            Application application = document.Application;
            CurveArray ConstructionLinesArray = application.Create.NewCurveArray();

            List<XYZ> roofAnchorPointsList = getRoofAnchorPointsList(wallsList[0]);

            for (int i = 0; i < 4; i++)
            {
                LocationCurve hostCurveForRoofRib = HostCurveForObject.getHostCurveForElement(wallsList[i]);
                XYZ startPointOfHostCurveForRoofRib = HostCurveForObject.getStartPointOfHostCurve(hostCurveForRoofRib);
                XYZ endPointOfHostCurveForRoofRib = HostCurveForObject.getEndPointOfHostCurve(hostCurveForRoofRib);
                Line constructionLine = Line.CreateBound(startPointOfHostCurveForRoofRib + roofAnchorPointsList[i], endPointOfHostCurveForRoofRib + roofAnchorPointsList[i + 1]);
                ConstructionLinesArray.Append(constructionLine);
            }
            return ConstructionLinesArray;
        }
        private static List<XYZ> getRoofAnchorPointsList(Wall wall)
        {
            double increment = Increment.getIncrement(wall.Width);
            List<XYZ> roofAnchorPointsList = AnchorPoints.getListOfAnchorPoints(increment, increment);
            return roofAnchorPointsList;
        }

        public static void constructRoof(Document document, Level level, List<Wall> wallsList, double roofWidth, double roofDepth, double roofHeight)
        {
            RoofType roofType = getRoofTypeFromProject(document);

            double revitRoofDepth = UnitUtils.ConvertToInternalUnits(roofDepth, UnitTypeId.Millimeters);
            double revitRoofWidth = UnitUtils.ConvertToInternalUnits(roofWidth, UnitTypeId.Millimeters);
            double increment = Increment.getIncrement(wallsList[0].Width);

            CurveArray ConstructionLinesArray = getConstructionLinesArray(document, level, increment, revitRoofDepth, roofHeight);

            double extrusionStart = -Increment.getIncrementedHalfValue(increment, revitRoofWidth);
            double extrusionEnd = Increment.getIncrementedHalfValue(increment, revitRoofWidth);

            constructRoofTransaction(document, level, ConstructionLinesArray, roofType, extrusionStart, extrusionEnd);
        }
        private static void constructRoofTransaction(Document document, Level level, CurveArray ConstructionLinesArray, RoofType roofType, double extrusionStart, double extrusionEnd)
        {
            Transaction tr = new Transaction(document, "Построение крыши");
            tr.Start();
            ReferencePlane plane = document.Create.NewReferencePlane(new XYZ(0, 0, 0), new XYZ(0, 0, 20), new XYZ(0, 20, 0), document.ActiveView);
            ExtrusionRoof extrusionRoof = document.Create.NewExtrusionRoof(ConstructionLinesArray, plane, level, roofType, extrusionStart, extrusionEnd);
            extrusionRoof.EaveCuts = EaveCutterType.TwoCutSquare;
            tr.Commit();
        }
        private static CurveArray getConstructionLinesArray(Document document, Level level, double increment, double roofDepth, double roofHeight)
        {
            double curveStart = -Increment.getIncrementedHalfValue(increment, roofDepth);
            double curveEnd = Increment.getIncrementedHalfValue(increment, roofDepth);

            Application application = document.Application;
            CurveArray ConstructionLinesArray = application.Create.NewCurveArray();
            ConstructionLinesArray.Append(Line.CreateBound(new XYZ(0, curveStart, level.Elevation), new XYZ(0, 0, level.Elevation + roofHeight)));
            ConstructionLinesArray.Append(Line.CreateBound(new XYZ(0, 0, level.Elevation + roofHeight), new XYZ(0, curveEnd, level.Elevation)));
            return ConstructionLinesArray;
        }
    }   

    public static class HostCurveForObject
    {
        public static XYZ getPointForLocateElement(Wall wall)
        {
            LocationCurve hostCurve = getHostCurveForElement(wall);
            XYZ startPointOfHostCurve = getStartPointOfHostCurve(hostCurve);
            XYZ endPointOfHostCurve = getEndPointOfHostCurve(hostCurve);
            XYZ middlePointOfHostCurve = (startPointOfHostCurve + endPointOfHostCurve) / 2;
            return middlePointOfHostCurve;
        }
        public static LocationCurve getHostCurveForElement(Wall wall)
        {
            LocationCurve hostCurveForElement = wall.Location as LocationCurve;
            return hostCurveForElement;
        }
        public static XYZ getStartPointOfHostCurve(LocationCurve hostCurveForElement)
        {
            XYZ startPointOfHostCurveForElement = hostCurveForElement.Curve.GetEndPoint(0);
            return startPointOfHostCurveForElement;
        }
        public static XYZ getEndPointOfHostCurve(LocationCurve hostCurveForElement)
        {
            XYZ endPointOfHostCurveForElement = hostCurveForElement.Curve.GetEndPoint(1);
            return endPointOfHostCurveForElement;
        }
    }

    public static class AnchorPoints
    {
        public static List<XYZ> getListOfAnchorPoints(double xIncrement, double yIncrement)
        {
            List<XYZ> anchorPointsList = new List<XYZ>();
            anchorPointsList.Add(new XYZ(-xIncrement, -yIncrement, 0));
            anchorPointsList.Add(new XYZ(xIncrement, -yIncrement, 0));
            anchorPointsList.Add(new XYZ(xIncrement, yIncrement, 0));
            anchorPointsList.Add(new XYZ(-xIncrement, yIncrement, 0));
            anchorPointsList.Add(new XYZ(-xIncrement, -yIncrement, 0));
            return anchorPointsList;
        }
    }

    public static class Increment
    {
        public static double getIncrementedHalfValue(double increment, double value)
        {
            return (value / 2 + increment);
        }
        public static double getIncrement(double value)
        {
            return value / 2;
        }
    }
}
