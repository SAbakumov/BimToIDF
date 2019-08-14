#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
//using IDFFile;
#endregion

namespace BIMToIDF
{
    [Transaction(TransactionMode.Manual)]
  
    public class Command : IExternalCommand
    {
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            System.Windows.Forms.Application.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            string path = Environment.CurrentDirectory + "/Data/";
            InputData userData = new InputData();
            userData.ShowDialog();
            List<List<Utility.StructureFloor>> sampledData = GetWallsAndFloorsAndCeilingsFromMasses(doc,userData);
            foreach (List<Utility.StructureFloor>  sample in sampledData)
            {
                IDFFile.IDFFile WrittenIDFFile = new IDFFile.IDFFile();
                WrittenIDFFile = Utility.AssignIDFFileParameters(WrittenIDFFile, sample);
                string fullFileName = "C:/Documents/TestFile_"+sampledData.IndexOf(sample)+".idf" ;
                WrittenIDFFile.GenerateOutput(false, "Annual");
                File.WriteAllLines(fullFileName, WrittenIDFFile.WriteFile());
            }
            int j = 0;
            foreach (List<Utility.StructureFloor> Modelsample in sampledData)
            {
                
                IDFFile.IDFFile WrittenIDFFile = new IDFFile.IDFFile();
                WrittenIDFFile = Utility.AssignIDFFileParameters(WrittenIDFFile, Modelsample);
                IDFFile.Building Bui = WrittenIDFFile.building;
                Tuple<List<Dictionary<string, double>>, List<Dictionary<string, double>>> RandomAttributes = Utility.GenerateProbabilisticRandomValueLists(Bui, userData.numSamples);
                List<IDFFile.Building> AllRandomBuildings = new List<IDFFile.Building>();
                for (int i=0;i<userData.numSamples; i++)
                {
                    IDFFile.Building RandBuild = Utility.GetRandomCopyOfModelBuilding(Utility.DeepClone(Bui), RandomAttributes.Item1[i], RandomAttributes.Item2[i]);
                    AllRandomBuildings.Add(Utility.DeepClone(RandBuild));
                }
                for (int i = 0; i < userData.numSamples; i++)
                {
                    IDFFile.IDFFile IDFFileSample= new IDFFile.IDFFile();
                    IDFFileSample.building = AllRandomBuildings[i];
                    string fullFileName = "C:/Documents/FilesForModelBuildings/File_" + i +"_FromModel_"+ j + ".idf";
                    IDFFileSample.GenerateOutput(false, "Annual");
                    File.WriteAllLines(fullFileName, IDFFileSample.WriteFile());
                }
                j++;
            }
                return Result.Succeeded;
        }
        public static List<List<Utility.StructureFloor>> GetWallsAndFloorsAndCeilingsFromMasses(Document doc,InputData userData)
        {
      

 
            int numberofFloors = userData.numFloors;

            //Initialize building elements
            List<List<Utility.StructureFloor>> structures = new List<List<Utility.StructureFloor>>();
            //Get masses from object in Revit

            FilteredElementCollector masses = new FilteredElementCollector(doc).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_Mass);
            foreach (Element mass in masses)
            {
                Options op = new Options() { ComputeReferences = true };

                GeometryElement gElement = mass.get_Geometry(op);

                foreach (GeometryObject SolidStructure in gElement)
                {
                    Solid GeoObject = SolidStructure as Solid;
                    if (GeoObject != null)
                    {

                        FaceArray AllFacesFromModel = GeoObject.Faces;
                        if (AllFacesFromModel.Size != 0)
                        {

                            List<Face> AllWalls = new List<Face>();
                            List<IDFFile.XYZList> WallsBuildingSurface = new List<IDFFile.XYZList>();
                            List<IDFFile.XYZList> FloorsOrCeilingsBuildingSurface = new List<IDFFile.XYZList>();
                            List<Tuple<double, IDFFile.XYZList>> FacesWithAreasWall = new List<Tuple<double, IDFFile.XYZList>>();
                            List<Tuple<double, IDFFile.XYZList>> FacesWithAreasPerp = new List<Tuple<double, IDFFile.XYZList>>();
                            List<IDFFile.XYZList> EdgeLoopRoofOrFloor = new List<IDFFile.XYZList>();

                            double MaximalHeight = 0;
                            double MinimalHeight = 0;
                            List<Tuple<IDFFile.XYZ, IDFFile.XYZ>> EdgeArrayForPossibleWall = new List<Tuple<IDFFile.XYZ, IDFFile.XYZ>>();
                            foreach (Face PossibleWall in AllFacesFromModel)
                            {
                                XYZ fNormal = (PossibleWall as PlanarFace).FaceNormal;

                                //checks if it is indeed a wall by computing the normal with respect to 001
                                if (Math.Abs(fNormal.Z) <= 0.001)
                                {
                                    (IDFFile.XYZList WallToAddToTotalBuildingSurface, double MaxH, double MinH) = Utility.GetIfFloorOrCeilingOrWall(PossibleWall, IDFFile.SurfaceType.Wall, MaximalHeight, MinimalHeight);

                                    double area = PossibleWall.Area;
                                    Tuple<double, IDFFile.XYZList> FaceToAdd = new Tuple<double, IDFFile.XYZList>(area, WallToAddToTotalBuildingSurface);

                                    FacesWithAreasWall.Add(FaceToAdd);

                                    MaximalHeight = MaxH;
                                    MinimalHeight = MinH;
                                }
                                else
                                {
                                    (IDFFile.XYZList FloorOrCelingToAddToTotalBuildingSurface, double MaxH, double MinH) = Utility.GetIfFloorOrCeilingOrWall(PossibleWall, IDFFile.SurfaceType.Floor, MaximalHeight, MinimalHeight);
                                    double area = PossibleWall.Area;
                                    Tuple<double, IDFFile.XYZList> FaceToAdd = new Tuple<double, IDFFile.XYZList>(area, FloorOrCelingToAddToTotalBuildingSurface);
                                    FacesWithAreasPerp.Add(FaceToAdd);
                                    EdgeArrayForPossibleWall = Utility.GetEdgesOfPolygon(PossibleWall as PlanarFace);


                                    MaximalHeight = MaxH;
                                    MinimalHeight = MinH;

                                }

                            }

                            List<Utility.StructureFloor> AllItemsPerStructure = new List<Utility.StructureFloor>();
                            for (int i = 0; i < numberofFloors; i++)
                            {
                                Utility.StructureFloor FloorNumber = new Utility.StructureFloor();
                                AllItemsPerStructure.Add(FloorNumber);
                            }
                            List<List<Tuple<double, IDFFile.XYZList>>> Roof = new List<List<Tuple<double, IDFFile.XYZList>>>();
                            Roof = Utility.GetRoof(FacesWithAreasPerp, MaximalHeight);
                            IDFFile.XYZList RoofPoints = Roof[0][0].Item2;
                            List<IDFFile.XYZ> DayPoints  =Utility.GetDayLightPointsXYZList(RoofPoints, EdgeArrayForPossibleWall);
                            IDFFile.Building BuildingToInitialize = Utility.InitializeBuildingFromInput(userData,DayPoints, MinimalHeight, MaximalHeight);
                            AllItemsPerStructure = Utility.AssignZonePerFloorElementOfCertainType(AllItemsPerStructure, BuildingToInitialize, IDFFile.SurfaceType.Roof, Roof,numberofFloors-1);

                            List<List<Tuple<double, IDFFile.XYZList>>> WallPerFloor = new List<List<Tuple<double, IDFFile.XYZList>>>();
                            WallPerFloor = Utility.GetWallsPerFloor(FacesWithAreasWall, MaximalHeight, MinimalHeight, numberofFloors , doc);
                            AllItemsPerStructure = Utility.AssignZonePerFloorElementOfCertainType(AllItemsPerStructure, BuildingToInitialize, IDFFile.SurfaceType.Wall, WallPerFloor, numberofFloors-1);


                            List<List<Tuple<double, IDFFile.XYZList>>> FloorsPerFloor = new List<List<Tuple<double, IDFFile.XYZList>>>();
                            FloorsPerFloor = Utility.GetFloorsAndCeilingPerFloor(FacesWithAreasPerp, MaximalHeight, MinimalHeight, 3);
                            AllItemsPerStructure = Utility.AssignZonePerFloorElementOfCertainType(AllItemsPerStructure, BuildingToInitialize, IDFFile.SurfaceType.Floor, FloorsPerFloor, numberofFloors-1);

                            structures.Add(AllItemsPerStructure);
                        }
                    }
                }

            }
            return structures;
        }


    }

    public static class Utility
    {

        public static List<StructureFloor> AssignZonePerFloorElementOfCertainType(List<StructureFloor> AllItemsPerStructure, IDFFile.Building Bui, IDFFile.SurfaceType Type, List<List<Tuple<double, IDFFile.XYZList>>> AllElementsInFloor,int maxnumfloor)
        {

            int numFloor = 0;
            foreach (List<Tuple<double, IDFFile.XYZList>> Element in AllElementsInFloor)
            {
                List<IDFFile.BuildingSurface> SurfaceTypesPerFloor = new List<IDFFile.BuildingSurface>();
                IDFFile.ZoneList AllZonesPerFloor = Bui.zoneLists[0];
                foreach (Tuple<double, IDFFile.XYZList> ElementCoordinates in Element)
                {
                    if (Type == IDFFile.SurfaceType.Wall)
                    {
                        IDFFile.BuildingSurface ElementSurface = new IDFFile.BuildingSurface(AllZonesPerFloor.listZones[numFloor], ElementCoordinates.Item2, ElementCoordinates.Item1, Type);

                        ElementSurface.SunExposed = "NoSun";
                        ElementSurface.WindExposed = "NoWind";
                        ElementSurface.zone.CalcAreaVolumeHeatCapacity();
                        SurfaceTypesPerFloor.Add(ElementSurface);
                    }
                    if (Type == IDFFile.SurfaceType.Floor)
                    {
                        if (numFloor > 0)
                        {
                            IDFFile.BuildingSurface ElementSurface = new IDFFile.BuildingSurface(AllZonesPerFloor.listZones[numFloor], ElementCoordinates.Item2, ElementCoordinates.Item1, IDFFile.SurfaceType.Ceiling);
                            ElementSurface.OutsideObject = AllZonesPerFloor.listZones[numFloor - 1].name;
                            ElementSurface.OutsideCondition = "Zone";
                            ElementSurface.SunExposed = "NoSun";
                            ElementSurface.WindExposed = "NoWind";
                            ElementSurface.zone.CalcAreaVolumeHeatCapacity();
                            SurfaceTypesPerFloor.Add(ElementSurface);
                        }
                        else
                        {
                            IDFFile.BuildingSurface ElementSurface = new IDFFile.BuildingSurface(AllZonesPerFloor.listZones[numFloor], ElementCoordinates.Item2, ElementCoordinates.Item1, Type);
                            ElementSurface.SunExposed = "NoSun";
                            ElementSurface.WindExposed = "NoWind";
                            ElementSurface.zone.CalcAreaVolumeHeatCapacity();
                            SurfaceTypesPerFloor.Add(ElementSurface);
                        }

                    }
                    if (Type == IDFFile.SurfaceType.Roof)
                    {
                        IDFFile.BuildingSurface ElementSurface = new IDFFile.BuildingSurface(AllZonesPerFloor.listZones[maxnumfloor], ElementCoordinates.Item2, ElementCoordinates.Item1, Type);

                        ElementSurface.SunExposed = "NoSun";
                        ElementSurface.WindExposed = "NoWind";
                        ElementSurface.zone.CalcAreaVolumeHeatCapacity();
                        SurfaceTypesPerFloor.Add(ElementSurface);
                    }


                }

                if (Type == IDFFile.SurfaceType.Wall)
                {
                    AllItemsPerStructure[numFloor].FloorWalls = SurfaceTypesPerFloor;
                }
                if (Type == IDFFile.SurfaceType.Floor)
                {
                    AllItemsPerStructure[numFloor].FloorFloors = SurfaceTypesPerFloor;
                }
                if (Type == IDFFile.SurfaceType.Roof )
                {
                    AllItemsPerStructure[maxnumfloor].GlobalRoof = SurfaceTypesPerFloor;
                }

                numFloor++;
            }
            if (Type == IDFFile.SurfaceType.Wall)
            {
                foreach (StructureFloor floor in AllItemsPerStructure)
                {

                    floor.FloorWalls[0].zone.building.AddZone(floor.FloorWalls[0].zone);
                    floor.FloorWalls[0].zone.building.GeneratePeopleLightingElectricEquipment();
                    floor.FloorWalls[0].zone.building.GenerateInfiltraitionAndVentillation();

                }
                AllItemsPerStructure[0].FloorWalls[0].zone.building.GenerateHVAC(true, false, false);

            }

            return AllItemsPerStructure;
        }
        public static Tuple<List<Dictionary<string, double>>, List<Dictionary<string,double>>> GenerateProbabilisticRandomValueLists(IDFFile.Building ModelBuilding,int numsamples)

        {


            List<Dictionary<string, double>> AllWWRvals = new List<Dictionary<string, double>>();
            List<Dictionary<string, double>> AllRandValuesList = new List<Dictionary<string, double>>();

            IDFFile.ProbabilisticBuildingConstruction ProbValBuild = ModelBuilding.pBuildingConstruction;
            IDFFile.ProbabilisticWWR ProbValWWR = ModelBuilding.pWWR;


            for (int i = 0; i< numsamples; i++)
            {
                Random r = new Random();

                Dictionary<string, double> WWRvals = new Dictionary<string, double>();
                Dictionary<string, double> AllRandValues = new Dictionary<string, double>();

                AllRandValues["uWall"]=GetRandom(ProbValBuild.uWall ,r);
                AllRandValues["GFloor"] = GetRandom(ProbValBuild.uGFloor, r);
                AllRandValues["uRoof"] = GetRandom(ProbValBuild.uRoof, r);
                AllRandValues["uWindow"] = GetRandom(ProbValBuild.uWindow, r);
                AllRandValues["gWindow"] = GetRandom(ProbValBuild.gWindow, r);
                AllRandValues["uIWall"] = GetRandom(ProbValBuild.uIWall, r);
                AllRandValues["uIFloor"] = GetRandom(ProbValBuild.uIFloor, r);
                AllRandValues["HCFloor"] = GetRandom(ProbValBuild.HCFloor, r);

                WWRvals["East"] = GetRandom(ProbValWWR.east, r);
                WWRvals["South"] = GetRandom(ProbValWWR.south, r);
                WWRvals["North"] = GetRandom(ProbValWWR.north, r);
                WWRvals["West"] = GetRandom(ProbValWWR.west, r);

                AllWWRvals.Add(Utility.DeepClone(WWRvals));
                AllRandValuesList.Add(Utility.DeepClone(AllRandValues));

            }
            Tuple<List<Dictionary<string, double>>, List<Dictionary<string, double>>> RandVals = Tuple.Create(AllRandValuesList,AllWWRvals);

            return RandVals;
        }
        public static IDFFile.Building InitializeBuildingFromInput(InputData userData, List<IDFFile.XYZ> DayLightPoints, double minHeightOfTheSolid, double maxHeightOfTheSolid)
        {
            IDFFile.ZoneList ListOfZonesPerFloor = new IDFFile.ZoneList("Office");



            Dictionary<string, double[]> WindowConstructData = userData.windowConstruction;
            Dictionary<string, double[]> BuildingConstructionData = userData.buildingConstruction;
            double[] uWall = BuildingConstructionData["uWall"];
            double[] uGFloor = BuildingConstructionData["uGFloor"];
            double[] uRoof = BuildingConstructionData["uRoof"];
            double[] uWindow = BuildingConstructionData["uWindow"];
            double[] gWindow = BuildingConstructionData["gWindow"];
            double[] cCOP = BuildingConstructionData["CCOP"];
            double[] BEFF = BuildingConstructionData["BEff"];
            double[] uIFloor = BuildingConstructionData["uIFloor"];
            double[] HCFloor = BuildingConstructionData["HCFloor"];
            double Hours = userData.operatingHours[0];
            double infiltration = userData.infiltration[0];
            double Heatgain = userData.iHG[0];

            double[] heatingSetPoints = new double[] { 10, 20 };
            double[] coolingSetPoints = new double[] { 28, 24 };
            double equipOffsetFraction = 0.1;
            //THIS NEEDS TO BE CHECKED - INSTEAD OF SECOND uWall there should be uIWall
            IDFFile.ProbabilisticWWR ProbabilisticWindowConstruction = new IDFFile.ProbabilisticWWR(WindowConstructData["wWR1"], WindowConstructData["wWR2"], WindowConstructData["wWR3"], WindowConstructData["wWR4"]);
            IDFFile.ProbabilisticBuildingConstruction ProbabilisticAttributes = new IDFFile.ProbabilisticBuildingConstruction(uWall, uGFloor, uRoof, uIFloor, uWall, uWindow, gWindow, HCFloor);
            IDFFile.Building bui = new IDFFile.Building
            {
                buildingConstruction = new IDFFile.BuildingConstruction(uWall[0], uGFloor[0], uRoof[0], uWindow[0], gWindow[0], 0.25, 0.25, 1050),
                WWR = new IDFFile.WWR(WindowConstructData["wWR1"][0], WindowConstructData["wWR2"][0], WindowConstructData["wWR3"][0], WindowConstructData["wWR4"][0]),

                chillerCOP = cCOP[0], boilerEfficiency = BEFF[0],
                LightHeatGain = Heatgain / 2,
                ElectricHeatGain = Heatgain / 2,
                operatingHours = Hours,
                infiltration = infiltration,
            };
            bui.pWWR = ProbabilisticWindowConstruction;
            bui.pBuildingConstruction = ProbabilisticAttributes;
            bui.CreateSchedules(heatingSetPoints, coolingSetPoints, equipOffsetFraction);
            bui.GenerateConstructionWithIComponentsU();
            IDFFile.ScheduleCompact Schedule = new IDFFile.ScheduleCompact();


            for (int i = 0; i <= userData.numFloors - 1; i++)
            {
                IDFFile.Zone zone = new IDFFile.Zone(bui, "Zone_" + i, i);
                zone.name = "Zone_" + i;
                IDFFile.People newpeople = new IDFFile.People(10);
                zone.people = newpeople;
                ListOfZonesPerFloor.listZones.Add(zone);

                List<IDFFile.XYZ> NewDayPoints = new List<IDFFile.XYZ>();
                NewDayPoints = Utility.DeepClone(DayLightPoints);

                for (int j = 0; j < DayLightPoints.Count; j++)
                {
                    NewDayPoints[j].Z = i * (maxHeightOfTheSolid - minHeightOfTheSolid) / (userData.numFloors) + 0.9;

                }

                IDFFile.DayLighting DayPoints = new IDFFile.DayLighting(ListOfZonesPerFloor.listZones[i], Schedule, Utility.DeepClone(NewDayPoints), 500);
                ListOfZonesPerFloor.listZones[i].DayLightControl = DayPoints;

            }

            bui.AddZoneList(ListOfZonesPerFloor);

            return bui;
        }


        public static List<List<Tuple<double, IDFFile.XYZList>>> GetWallsPerFloor(List<Tuple<double, IDFFile.XYZList>> TotalWall, double maxHeightOfTheSolid, double minHeightOfTheSolid, int nFloor, Document doc)
        {
            List<List<Tuple<double, IDFFile.XYZList>>> FloorWalls = new List<List<Tuple<double, IDFFile.XYZList>>>();
            for (int i = 0; i < nFloor; i++)
            {

                List<Tuple<double, IDFFile.XYZList>> FloorNumber = new List<Tuple<double, IDFFile.XYZList>>();
                FloorWalls.Add(FloorNumber);
            }
            foreach (Tuple<double, IDFFile.XYZList> wall in TotalWall)
            {
                //double TotalArea = wall.area;
                double minimalHeightOfWall = 0;
                double maximalHeightOfWall = 0;

                foreach (IDFFile.XYZ pointsOfWall in wall.Item2.xyzs)
                {
                    if (pointsOfWall.Z > maximalHeightOfWall)
                    {
                        maximalHeightOfWall = pointsOfWall.Z;
                    }
                    if (pointsOfWall.Z < minimalHeightOfWall)
                    {
                        minimalHeightOfWall = pointsOfWall.Z;
                    }
                }


                for (int i = 0; i < nFloor; i++)
                {
                    double WhichFloor = i * (maximalHeightOfWall - minimalHeightOfWall) / nFloor;
                    double NextFloor = (i + 1) * (maximalHeightOfWall - minimalHeightOfWall) / nFloor;

                    if (minimalHeightOfWall <= WhichFloor & maximalHeightOfWall >= NextFloor)
                    {
                        List<IDFFile.XYZ> AllPointsOfWallPerFloor = new List<IDFFile.XYZ>();

                        foreach (IDFFile.XYZ wallpart in wall.Item2.xyzs)
                        {
                            if (Math.Round(wallpart.Z, 4) <= Math.Round(WhichFloor, 4))
                            {
                                IDFFile.XYZ wallpartToAdd = new IDFFile.XYZ();
                                wallpartToAdd.Z = WhichFloor;
                                wallpartToAdd.X = wallpart.X;
                                wallpartToAdd.Y = wallpart.Y;
                                AllPointsOfWallPerFloor.Add(wallpartToAdd);
                            }
                            if (Math.Round(wallpart.Z, 4) >= Math.Round(NextFloor, 4))
                            {
                                IDFFile.XYZ wallpartToAdd = new IDFFile.XYZ();
                                wallpartToAdd.Z = NextFloor;
                                wallpartToAdd.X = wallpart.X;
                                wallpartToAdd.Y = wallpart.Y;
                                AllPointsOfWallPerFloor.Add(wallpartToAdd);

                            }
                        }
                        //double AreaN = ((NextFloor - WhichFloor) / (maxHeightOfTheSolid - minHeightOfTheSolid)) * TotalArea;
                        IDFFile.XYZList XYZCoordsOfFloorWallXYZList = new IDFFile.XYZList(AllPointsOfWallPerFloor);

                        Tuple<double, IDFFile.XYZList> WallToAdd = new Tuple<double, IDFFile.XYZList>((maxHeightOfTheSolid - minHeightOfTheSolid) / (maxHeightOfTheSolid * nFloor) * wall.Item1, XYZCoordsOfFloorWallXYZList);
                        FloorWalls[i].Add(WallToAdd);
                    }

                }
            }

            return FloorWalls;
        }

        public static List<XYZ> GetPoints(Face face)
        {
            List<XYZ> vectorList = new List<XYZ>();
            EdgeArray edgeArray = face.EdgeLoops.get_Item(0);
            foreach (Edge e in edgeArray)
            {
                vectorList.Add(e.AsCurveFollowingFace(face).GetEndPoint(0));
            }
            return vectorList;
        }

        public static List<List<Tuple<double, IDFFile.XYZList>>> GetFloorsAndCeilingPerFloor(List<Tuple<double, IDFFile.XYZList>> TotalFloor, double maxHeightOfTheSolid, double minHeightOfTheSolid, int nFloor)
        {
            List<List<Tuple<double, IDFFile.XYZList>>> FloorFloors = new List<List<Tuple<double, IDFFile.XYZList>>>();
            for (int i = 0; i < nFloor; i++)
            {
                List<Tuple<double, IDFFile.XYZList>> FloorNumber = new List<Tuple<double, IDFFile.XYZList>>();
                FloorFloors.Add(FloorNumber);
            }
            List<IDFFile.XYZ> FloorToCopyThroughBuilding = new List<IDFFile.XYZ>();
            double Ground = 0;
            double TotalArea = 0;

            foreach (Tuple<double, IDFFile.XYZList> Floor in TotalFloor)
            {
                TotalArea = Floor.Item1;
                List<IDFFile.XYZ> ZlocationOfFloor = Floor.Item2.xyzs;
                IDFFile.XYZ FloorToIntersect = ZlocationOfFloor[0];
                Ground = FloorToIntersect.Z;
                if (Ground == minHeightOfTheSolid)
                {
                    FloorToCopyThroughBuilding = ZlocationOfFloor;
                }
            }

            for (int i = 0; i < nFloor; i++)
            {
                double WhichFloor = i * (maxHeightOfTheSolid - minHeightOfTheSolid) / nFloor;
                double NextFloor = (i + 1) * (maxHeightOfTheSolid - minHeightOfTheSolid) / nFloor;
                List<IDFFile.XYZ> AllPointsOfFloorPerFloor = new List<IDFFile.XYZ>();

                foreach (IDFFile.XYZ floorpoint in FloorToCopyThroughBuilding)
                {
                    if (i == 0)
                    {
                        IDFFile.XYZ FloorpartToAdd = new IDFFile.XYZ();
                        FloorpartToAdd.Z = floorpoint.Z;
                        FloorpartToAdd.X = floorpoint.X;
                        FloorpartToAdd.Y = floorpoint.Y;
                        AllPointsOfFloorPerFloor.Add(FloorpartToAdd);
                    }
                    else
                    {
                        IDFFile.XYZ FloorpartToAdd = new IDFFile.XYZ();
                        FloorpartToAdd.Z = WhichFloor + Ground;
                        FloorpartToAdd.X = floorpoint.X;
                        FloorpartToAdd.Y = floorpoint.Y;
                        AllPointsOfFloorPerFloor.Add(FloorpartToAdd);
                    }
                }
                IDFFile.XYZList XYZCoordsOfFloorWallXYZList = new IDFFile.XYZList(AllPointsOfFloorPerFloor);
                Tuple<double, IDFFile.XYZList> FloorToAdd = new Tuple<double, IDFFile.XYZList>(TotalArea, XYZCoordsOfFloorWallXYZList);
                FloorFloors[i].Add(FloorToAdd);

            }
            return FloorFloors;
        }
        public static List<IDFFile.XYZ> GetDayLightPointsXYZList(IDFFile.XYZList FloorFace, List<Tuple<IDFFile.XYZ, IDFFile.XYZ>> EdgeArrayForPossibleWall)
        {

            List<IDFFile.XYZ> DayLightPoints = new List<IDFFile.XYZ>();
            List<IDFFile.XYZ> FloorFacePoints = FloorFace.xyzs;
            List<EnergyObjects.Vector> VectorList = new List<EnergyObjects.Vector>();
            foreach (IDFFile.XYZ Point in FloorFacePoints)
            {
                double xcoord = Math.Round(Point.X, 4);
                double ycoord = Math.Round(Point.Y, 4);
                double zcoord = Math.Round(Point.Z, 4);
                EnergyObjects.Vector NewVector = new EnergyObjects.Vector(xcoord, ycoord, zcoord);
                VectorList.Add(NewVector);
            }
            EnergyObjects.Vector[] AllPoints = new EnergyObjects.Vector[VectorList.Count];
            for (int i = 0; i < VectorList.Count; i++)
            {
                AllPoints[i] = VectorList[i];

            }
            EnergyObjects.Vector[] CentersOfMass = TriangulateAndGetCenterOfMass(AllPoints);
            foreach (EnergyObjects.Vector CM in CentersOfMass)
            {
                if (RayCastToCheckIfIsInside(EdgeArrayForPossibleWall, CM))
                {
                    IDFFile.XYZ CMToAdd = new IDFFile.XYZ();
                    CMToAdd.X = CM.x;
                    CMToAdd.Y = CM.y;
                    CMToAdd.Z = CM.z;
                    DayLightPoints.Add(CMToAdd);
                }
            }
            //IDFFile.XYZList AllDayLightPoints = new IDFFile.XYZList(DayLightPoints);
            return DayLightPoints;
        }
        public static EnergyObjects.Vector[] TriangulateAndGetCenterOfMass(EnergyObjects.Vector[] AllPoints)
        {
            int[] PointNumbers = Enumerable.Range(-1, AllPoints.Length + 2).ToArray();

            PointNumbers[0] = AllPoints.Length - 1;
            PointNumbers[PointNumbers.Length - 1] = 0;
            EnergyObjects.Vector[] CentersOfMass = new EnergyObjects.Vector[AllPoints.Length];

            for (int i = 1; i < PointNumbers.Length - 1; i++)
            {

                EnergyObjects.Vector[] Triangle = new EnergyObjects.Vector[3];
                Triangle[0] = AllPoints[PointNumbers[i - 1]];
                Triangle[1] = AllPoints[PointNumbers[i]];
                Triangle[2] = AllPoints[PointNumbers[i + 1]];

                double xCm = Math.Round((Triangle[0].x + Triangle[1].x + Triangle[2].x) / 3, 4);
                double yCm = Math.Round((Triangle[0].y + Triangle[1].y + Triangle[2].y) / 3, 4);
                double zCm = Math.Round((Triangle[0].z + Triangle[1].z + Triangle[2].z) / 3, 4);

                EnergyObjects.Vector CenterOfMass = new EnergyObjects.Vector(xCm, yCm, zCm);
                CentersOfMass[i - 1] = CenterOfMass;
            }
            return CentersOfMass;
        }
        public static (IDFFile.XYZList, double, double) GetIfFloorOrCeilingOrWall(Face PossibleWall, IDFFile.SurfaceType CurrentSurfaceType, double MaximalHeight, double MinimalHeight)
        {
            XYZ fNormal = (PossibleWall as PlanarFace).FaceNormal;


            double Arean;
            Arean = PossibleWall.Area;
            List<XYZ> NewVertex = Utility.GetPoints(PossibleWall);

            List<IDFFile.XYZ> AllSingleWallPointList = new List<IDFFile.XYZ>();

            foreach (XYZ Item in NewVertex)
            {
                IDFFile.XYZ verteces = new IDFFile.XYZ();
                verteces.X = Item[0];
                verteces.Y = Item[1];
                verteces.Z = Math.Round(Item[2], 4);

                AllSingleWallPointList.Add(verteces);
                if (verteces.Z > MaximalHeight)
                {
                    MaximalHeight = verteces.Z;
                }
                if (verteces.Z < MinimalHeight)
                {
                    MinimalHeight = verteces.Z;
                }

            }


            IDFFile.XYZList pointList1 = new IDFFile.XYZList(AllSingleWallPointList);


            return (pointList1, MaximalHeight, MinimalHeight);
        }
        public static List<List<Tuple<double, IDFFile.XYZList>>> GetRoof(List<Tuple<double, IDFFile.XYZList>> PossibleRoof, double MaxH)
        {
            List<List<Tuple<double, IDFFile.XYZList>>> AllRoofs = new List<List<Tuple<double, IDFFile.XYZList>>>();
            List<Tuple<double, IDFFile.XYZList>> AllRoofsPossible = new List<Tuple<double, IDFFile.XYZList>>();
            foreach (Tuple<double, IDFFile.XYZList> roofP in PossibleRoof)
            {
                IDFFile.XYZList RoofCoordinates = roofP.Item2;
                double GetHeight = RoofCoordinates.xyzs[0].Z;
                if (GetHeight == MaxH)
                {
                    roofP.Item2.reverse();
                    AllRoofsPossible.Add(roofP);
                    AllRoofs.Add(AllRoofsPossible);
                }

            }
            return AllRoofs;
        }
        public static bool RayCastToCheckIfIsInside(List<Tuple<IDFFile.XYZ, IDFFile.XYZ>> EdgeArrayForPossibleWall, EnergyObjects.Vector CM)
        {
            bool isInside = false;
            int count = 0;
            foreach (Tuple<IDFFile.XYZ, IDFFile.XYZ> EdgeOfWall in EdgeArrayForPossibleWall)
            {

                double r = (CM.y - EdgeOfWall.Item2.Y) / (EdgeOfWall.Item1.Y - EdgeOfWall.Item2.Y);
                if (r > 0 && r < 1)
                {
                    double Xvalue = r * (EdgeOfWall.Item1.X - EdgeOfWall.Item2.X) + EdgeOfWall.Item2.X;
                    if (CM.x < Xvalue)
                    {
                        count++;
                    }
                }
            }
            if (count % 2 == 0)
            {
                isInside = false;
            }
            else
            {
                isInside = true;
            }
            return isInside;
        }
        public static List<Tuple<IDFFile.XYZ, IDFFile.XYZ>> GetEdgesOfPolygon(PlanarFace PossibleWall)
        {

            List<Tuple<IDFFile.XYZ, IDFFile.XYZ>> EdgeLoop = new List<Tuple<IDFFile.XYZ, IDFFile.XYZ>>();

            EdgeArrayArray edgeArrays = PossibleWall.EdgeLoops;
            foreach (EdgeArray edges in edgeArrays)
            {
                foreach (Edge edge in edges)
                {
                    // Get one test point
                    XYZ Point1 = edge.Evaluate(1);
                    XYZ Point2 = edge.Evaluate(0);
                    IDFFile.XYZ Vertex1 = new IDFFile.XYZ(Point1.X, Point1.Y, Point1.Z);
                    IDFFile.XYZ Vertex2 = new IDFFile.XYZ(Point2.X, Point2.Y, Point2.Z);

                    Tuple<IDFFile.XYZ, IDFFile.XYZ> EdgeToAdd = new Tuple<IDFFile.XYZ, IDFFile.XYZ>(Vertex1, Vertex2);
                    EdgeLoop.Add(EdgeToAdd);
                }
            }
            return EdgeLoop;
        }
        public class BuildingDesignParameters
        {
            public double Length, Width, Height, rLenA, rWidA, BasementDepth, Orientation,
                uWall, uGFloor, uRoof, hcSlab, uWindow, gWindow, uIWall, uIFloor,
                infiltration, hours, lhg, ehg, LEHG,
                beff, cCOP;
            public IDFFile.WWR wwr;
            public BuildingDesignParameters() { }
        }

        public class Structure
        {
            public List<StructureFloor> FloorOfStructure;


        }
        public class StructureFloor
        {
            public List<IDFFile.BuildingSurface> FloorWalls;
            public List<IDFFile.BuildingSurface> FloorFloors;
            public List<IDFFile.BuildingSurface> GlobalRoof;
        }
        public static IDFFile.IDFFile AssignIDFFileParameters(IDFFile.IDFFile IDFFileToReadAndWrite, List<Utility.StructureFloor> ExportedStructure)
        {



            IDFFileToReadAndWrite.building = ExportedStructure[0].FloorWalls[0].zone.building;


            return IDFFileToReadAndWrite;
        }
        public static T DeepClone<T>(T obj)
        {
            using (var ms = new MemoryStream())
            {
                var formatter = new BinaryFormatter();
                formatter.Serialize(ms, obj);
                ms.Position = 0;
                return (T)formatter.Deserialize(ms);
            }
        }
        public static double GetRandom(double[] range, Random r)
        {
            return (range[0] * (1 + range[1] * (2 * r.NextDouble() - 1)));
        }
        public static IDFFile.Building GetRandomCopyOfModelBuilding(IDFFile.Building ModelBuilding, Dictionary<string, double> AllRandValues, Dictionary<string, double> WWRPr)
        {

            ModelBuilding.buildingConstruction.uWall = AllRandValues["uWall"];
            ModelBuilding.buildingConstruction.uGFloor = AllRandValues["GFloor"];
            ModelBuilding.buildingConstruction.uRoof  =AllRandValues["uRoof"];
            ModelBuilding.buildingConstruction.uWindow = AllRandValues["uWindow"];
            ModelBuilding.buildingConstruction.gWindow = AllRandValues["gWindow"];
            ModelBuilding.buildingConstruction.uIWall = AllRandValues["uIWall"];
            ModelBuilding.buildingConstruction.uIFloor = AllRandValues["uIFloor"];
            ModelBuilding.buildingConstruction.hcIFloor = AllRandValues["HCFloor"];
            IDFFile.WWR WWR = new IDFFile.WWR(WWRPr["North"], WWRPr["East"], WWRPr["South"], WWRPr["West"]);

            ModelBuilding.WWR = WWR;
            ModelBuilding = UpdateFenestrations(ModelBuilding);
            ModelBuilding.GeneratePeopleLightingElectricEquipment();
            ModelBuilding.GenerateInfiltraitionAndVentillation();
            ModelBuilding.shadingControls = new List<IDFFile.ShadingControl>();
            ModelBuilding.windowMaterialShades = new List<IDFFile.WindowMaterialShade>();

            //ModelBuilding.CreateSchedules(heatingSetPoints, coolingSetPoints, equipOffsetFraction);
            ModelBuilding.GenerateConstructionWithIComponentsU();
            return ModelBuilding;

        }
        public static IDFFile.Building UpdateFenestrations(IDFFile.Building BuiRand)
        {
            foreach (IDFFile.BuildingSurface toupdate in BuiRand.bSurfaces)
            {
                if (toupdate.surfaceType == IDFFile.SurfaceType.Wall)
                {
                    toupdate.AssociateWWRandShadingLength();
                    toupdate.fenestrations = toupdate.CreateFenestration(1);
                }

            }
            return BuiRand;
        }

    }

}
