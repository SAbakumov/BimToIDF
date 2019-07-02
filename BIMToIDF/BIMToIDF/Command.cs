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
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Autodesk.Revit.ApplicationServices.Application app = uiapp.Application;
            Document doc = uidoc.Document;
            string path = Environment.CurrentDirectory + "/Data/";

            List<List<Utility.StructureFloor>> sampledData = GetWallsAndFloorsAndCeilingsFromMasses(doc);

            return Result.Succeeded;
        }
        public static List<List<Utility.StructureFloor>> GetWallsAndFloorsAndCeilingsFromMasses(Document doc)
        {
      

            InputData userData = new InputData();
            userData.ShowDialog();
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

                            double MaximalHeight = 0;
                            double MinimalHeight = 0;
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
                            IDFFile.ZoneList AllZonesPerFloor = Utility.InitializeBuildingFromInput(userData);

                            List<List<Tuple<double, IDFFile.XYZList>>> WallPerFloor = new List<List<Tuple<double, IDFFile.XYZList>>>();
                            WallPerFloor = Utility.GetWallsPerFloor(FacesWithAreasWall, MaximalHeight, MinimalHeight, numberofFloors , doc);
                            AllItemsPerStructure = Utility.AssignZonePerFloorElementOfCertainType(AllItemsPerStructure, AllZonesPerFloor, IDFFile.SurfaceType.Wall, WallPerFloor);



                            List<List<Tuple<double, IDFFile.XYZList>>> Roof = new List<List<Tuple<double, IDFFile.XYZList>>>();
                            Roof = Utility.GetRoof(FacesWithAreasPerp, MaximalHeight);
                            AllItemsPerStructure = Utility.AssignZonePerFloorElementOfCertainType(AllItemsPerStructure, AllZonesPerFloor, IDFFile.SurfaceType.Roof, Roof);



                            List<List<Tuple<double, IDFFile.XYZList>>> FloorsPerFloor = new List<List<Tuple<double, IDFFile.XYZList>>>();
                            FloorsPerFloor = Utility.GetFloorsAndCeilingPerFloor(FacesWithAreasPerp, MaximalHeight, MinimalHeight, 3);
                            AllItemsPerStructure = Utility.AssignZonePerFloorElementOfCertainType(AllItemsPerStructure, AllZonesPerFloor, IDFFile.SurfaceType.Floor, FloorsPerFloor);

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
   
        public static List<StructureFloor> AssignZonePerFloorElementOfCertainType(List<StructureFloor> AllItemsPerStructure , IDFFile.ZoneList AllZonesPerFloor, IDFFile.SurfaceType Type, List<List<Tuple<double, IDFFile.XYZList>>> AllElementsInFloor )
        {

            int numFloor = 0;
            foreach (List<Tuple<double, IDFFile.XYZList>> Element in AllElementsInFloor)
            {
                List<IDFFile.BuildingSurface> SurfaceTypesPerFloor = new List<IDFFile.BuildingSurface>();

                foreach (Tuple<double, IDFFile.XYZList> ElementCoordinates in Element)
                {
                    IDFFile.Zone checkZone = AllZonesPerFloor.listZones[numFloor];
                    IDFFile.BuildingSurface ElementSurface = new IDFFile.BuildingSurface(AllZonesPerFloor.listZones[1], ElementCoordinates.Item2, ElementCoordinates.Item1, Type);
                    ElementSurface.ConstructionName = "General_Floor_Ceiling";
                    ElementSurface.OutsideCondition = "Zone";
                    ElementSurface.OutsideObject = "Zone_numFloor_" + numFloor;
                    ElementSurface.SunExposed = "NoSun";
                    ElementSurface.WindExposed = "NoWind";
                    ElementSurface.zone.CalcAreaVolumeHeatCapacity();
                    SurfaceTypesPerFloor.Add(ElementSurface);
                }

                if( Type == IDFFile.SurfaceType.Wall)
                {
                    AllItemsPerStructure[numFloor].FloorWalls = SurfaceTypesPerFloor;
                }
                if (Type == IDFFile.SurfaceType.Floor)
                {
                    AllItemsPerStructure[numFloor].FloorFloors = SurfaceTypesPerFloor;
                }
                if (Type == IDFFile.SurfaceType.Roof)
                {
                    AllItemsPerStructure[numFloor].GlobalRoof = SurfaceTypesPerFloor;
                }
                numFloor++;
            }
            return AllItemsPerStructure;
        }
        public static IDFFile.ZoneList InitializeBuildingFromInput( InputData userData)
        {
            IDFFile.ZoneList ListOfZonesPerFloor = new IDFFile.ZoneList("Shape");



            Dictionary<string, double[]> WindowConstructData = userData.windowConstruction;
            Dictionary<string, double[]> BuildingConstructionData = userData.buildingConstruction;
            double[] uWall = BuildingConstructionData["uWall"];
            double[] uGFloor = BuildingConstructionData["uGFloor"];
            double[] uRoof = BuildingConstructionData["uRoof"];
            double[] uWindow = BuildingConstructionData["uWindow"];
            double[] gWindow = BuildingConstructionData["gWindow"];
            double[] cCOP = BuildingConstructionData["CCOP"];
            double[] BEFF = BuildingConstructionData["BEff"];
            double[] heatingSetPoints = new double[] { 10, 20 };
            double[] coolingSetPoints = new double[] { 28, 24 };
            double equipOffsetFraction = 0.1;



            //BuildingConstructionData['uWall'];
            IDFFile.Building bui = new IDFFile.Building
            {
                buildingConstruction = new IDFFile.BuildingConstruction(uWall[0], uGFloor[0], uRoof[0], uWindow[0], gWindow[0], 0.25, 0.25, 1050),
                WWR = new IDFFile.WWR(WindowConstructData["wWR1"][0], WindowConstructData["wWR2"][0], WindowConstructData["wWR3"][0], WindowConstructData["wWR4"][0]),

                chillerCOP = cCOP[0], boilerEfficiency = BEFF[0],
            };
            bui.CreateSchedules(heatingSetPoints, coolingSetPoints, equipOffsetFraction);
            bui.GenerateConstructionWithIComponentsU();
            bui.GeneratePeopleLightingElectricEquipment();
            bui.GenerateInfiltraitionAndVentillation();
            //bui.GenerateHVAC(true, false, false);

            for (int i=0; i <= userData.numFloors; i++)
            {
                IDFFile.Zone zone = new IDFFile.Zone(bui, "Zone_numFloor_" + i, i);
                ListOfZonesPerFloor.listZones.Add(zone);

            }

            return ListOfZonesPerFloor;
        }


        public static List<List<Tuple<double, IDFFile.XYZList>>> GetWallsPerFloor(List<Tuple<double, IDFFile.XYZList>> TotalWall, double maxHeightOfTheSolid,double minHeightOfTheSolid, int nFloor, Document doc)
        {
            List<List< Tuple<double, IDFFile.XYZList>> > FloorWalls = new List<List<Tuple<double, IDFFile.XYZList>>>();
            for(int i=0;i<nFloor; i++)
            {
             
                List<Tuple<double, IDFFile.XYZList>> FloorNumber = new List<Tuple<double, IDFFile.XYZList>>();
                FloorWalls.Add(FloorNumber);
            }
            foreach(Tuple<double, IDFFile.XYZList> wall in TotalWall)
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
                    double WhichFloor = i * (maximalHeightOfWall - minimalHeightOfWall) /nFloor;
                    double NextFloor = (i+1) * (maximalHeightOfWall - minimalHeightOfWall) / nFloor;

                    if (minimalHeightOfWall <= WhichFloor & maximalHeightOfWall >= NextFloor)
                    {
                        List<IDFFile.XYZ> AllPointsOfWallPerFloor = new List<IDFFile.XYZ>();

                        foreach (IDFFile.XYZ wallpart in wall.Item2.xyzs)
                        {
                            if (Math.Round(wallpart.Z, 4)  <= Math.Round(WhichFloor , 4) )
                            {
                                IDFFile.XYZ wallpartToAdd = new IDFFile.XYZ();
                                wallpartToAdd.Z = WhichFloor;
                                wallpartToAdd.X = wallpart.X;
                                wallpartToAdd.Y = wallpart.Y;
                                AllPointsOfWallPerFloor.Add(wallpartToAdd);
                            }
                            if (Math.Round(wallpart.Z, 4) >= Math.Round(NextFloor,4))
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
                       
                        Tuple<double, IDFFile.XYZList> WallToAdd = new Tuple<double, IDFFile.XYZList>((maxHeightOfTheSolid - minHeightOfTheSolid) /(maxHeightOfTheSolid  *nFloor )*wall.Item1, XYZCoordsOfFloorWallXYZList);
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
            double Ground =0;
            double TotalArea = 0;

            foreach (Tuple<double, IDFFile.XYZList> Floor in TotalFloor)
            {
                TotalArea =  Floor.Item1; 
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
                //IDFFile.Building NewBuilding = new IDFFile.Building();
                //IDFFile.Zone NewZone = new IDFFile.Zone();
                //IDFFile.SurfaceType CurrentSurfType = IDFFile.SurfaceType.Floor;
                //IDFFile.BuildingSurface WallToAddPerFloor = new IDFFile.BuildingSurface(NewZone, XYZCoordsOfFloorWallXYZList, TotalArea, CurrentSurfType);
                Tuple<double, IDFFile.XYZList> FloorToAdd = new Tuple<double, IDFFile.XYZList>(TotalArea, XYZCoordsOfFloorWallXYZList);
                FloorFloors[i].Add(FloorToAdd);

            }
            return FloorFloors;
        }
 

        public static (IDFFile.XYZList ,double,double ) GetIfFloorOrCeilingOrWall(Face PossibleWall, IDFFile.SurfaceType CurrentSurfaceType,double MaximalHeight, double MinimalHeight)
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
                    verteces.Z = Math.Round(Item[2],4);
                
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
            //IDFFile.Building NewBuilding = new IDFFile.Building();
            //IDFFile.Zone NewZone = new IDFFile.Zone(NewBuilding, "FirstFloorWall", 1);

            //IDFFile.BuildingSurface WallAddedToBuildingSurface = new IDFFile.BuildingSurface(NewZone, pointList1, Arean, CurrentSurfaceType);

            //return (WallAddedToBuildingSurface, MaximalHeight,MinimalHeight);

            return (pointList1, MaximalHeight, MinimalHeight);
        }
        public static List<List<Tuple<double, IDFFile.XYZList>>> GetRoof(List<Tuple<double, IDFFile.XYZList>> PossibleRoof ,double MaxH)
        {
            List<List<Tuple<double, IDFFile.XYZList>>> AllRoofs = new List<List<Tuple<double, IDFFile.XYZList>>>();
            List<Tuple<double, IDFFile.XYZList>> AllRoofsPossible = new List<Tuple<double, IDFFile.XYZList>>();
            //IDFFile.SurfaceType CurrentSurfaceType = IDFFile.SurfaceType.Roof;
            foreach (Tuple<double, IDFFile.XYZList> roofP in PossibleRoof)
            {
                IDFFile.XYZList RoofCoordinates = roofP.Item2;
                double GetHeight = RoofCoordinates.xyzs[0].Z;
                if (GetHeight == MaxH)
                {   
                    AllRoofsPossible.Add(roofP);
                    //roofP.surfaceType = CurrentSurfaceType;
                    AllRoofs.Add(AllRoofsPossible);
                }

            }
            return AllRoofs;
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
    }
    
}
