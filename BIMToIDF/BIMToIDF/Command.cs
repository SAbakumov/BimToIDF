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

            List<List<int>> sampledData = GetWallsAndFloorsAndCeilingsFromMasses(doc);

            return Result.Succeeded;
        }
        public static List<List<int>> GetWallsAndFloorsAndCeilingsFromMasses(Document doc)
        {
            List<List<IDFFile.BuildingSurface>> AllWallsFromBuildingPerFloor = new List<List<IDFFile.BuildingSurface>>();
            List<List<IDFFile.BuildingSurface>> AllFloorsFromBuildingPerFloor = new List<List<IDFFile.BuildingSurface>>();
            List<List<IDFFile.BuildingSurface>> AllCeilingsFromBuildingPerFloor = new List<List<IDFFile.BuildingSurface>>();


            InputData userData = new InputData();
            userData.ShowDialog();
            //Initialize building elements
            List<List<int>> structures = new List<List<int>>();
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
                        List<Face> AllWalls = new List<Face>();
                        List<IDFFile.BuildingSurface> WallsBuildingSurface = new List<IDFFile.BuildingSurface>();
                        List<IDFFile.BuildingSurface> FloorsOrCeilingsBuildingSurface = new List<IDFFile.BuildingSurface>();

                        double MaximalHeight = 0;
                        double MinimalHeight = 0;
                        foreach (Face PossibleWall in AllFacesFromModel)
                        {
                            XYZ fNormal = (PossibleWall as PlanarFace).FaceNormal;

                            //checks if it is indeed a wall by computing the normal with respect to 001
                            if (fNormal.Z == 0)
                            {
                                (IDFFile.BuildingSurface WallToAddToTotalBuildingSurface, double MaxH,double MinH) = Utility.GetIfFloorOrCeilingOrWall(PossibleWall, IDFFile.SurfaceType.Wall, MaximalHeight,MinimalHeight);
                                WallsBuildingSurface.Add(WallToAddToTotalBuildingSurface);
                                MaximalHeight = MaxH;
                                MinimalHeight = MinH;
                            }
                            else
                            {
                                (IDFFile.BuildingSurface FloorOrCelingToAddToTotalBuildingSurface, double MaxH, double MinH) = Utility.GetIfFloorOrCeilingOrWall(PossibleWall, IDFFile.SurfaceType.Floor, MaximalHeight,MinimalHeight);
                                FloorsOrCeilingsBuildingSurface.Add(FloorOrCelingToAddToTotalBuildingSurface);
                                MaximalHeight = MaxH;
                                MinimalHeight = MinH;

                            }

                        }
                        List < List < IDFFile.BuildingSurface >>  WallPerFloor= new List<List<IDFFile.BuildingSurface>>();
                        WallPerFloor = Utility.GetWallsPerFloor(WallsBuildingSurface,MaximalHeight, MinimalHeight,3);
                        List <IDFFile.BuildingSurface> Roof = new List<IDFFile.BuildingSurface>();

                        Roof = Utility.GetRoof(FloorsOrCeilingsBuildingSurface, MaximalHeight);

                        List<List<IDFFile.BuildingSurface>> FloorsPerFloor = new List<List<IDFFile.BuildingSurface>>();
                        FloorsPerFloor = Utility.GetFloorsAndCeilingPerFloor(FloorsOrCeilingsBuildingSurface, MaximalHeight, MinimalHeight, 3);


                    }
                }

            }
            return structures;
        }
        public List<double> InitializeZoneFromInput()
        {
            List<double> ZoneParameters = new List<double>();
            return ZoneParameters;
        }

    }

    public static class Utility
    {

        public static List<List<IDFFile.BuildingSurface>> GetWallsPerFloor(List<IDFFile.BuildingSurface> TotalWall, double maxHeightOfTheSolid,double minHeightOfTheSolid, int nFloor)
        {
            List<List<IDFFile.BuildingSurface>> FloorWalls = new List<List<IDFFile.BuildingSurface>>();
            for(int i=0;i<nFloor; i++)
            {
                List<IDFFile.BuildingSurface> FloorNumber = new List<IDFFile.BuildingSurface>();
                FloorWalls.Add(FloorNumber);
            }
            foreach(IDFFile.BuildingSurface wall in TotalWall)
            {
                double TotalArea = wall.area;
                double minimalHeightOfWall = wall.verticesList.xyzs[0].Z;
                double maximalHeightOfWall = wall.verticesList.xyzs[0].Z;
                foreach (IDFFile.XYZ pointsOfWall in wall.verticesList.xyzs)
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
                    double WhichFloor = i * (maxHeightOfTheSolid - minHeightOfTheSolid)/nFloor;
                    double NextFloor = (i+1) * (maxHeightOfTheSolid - minHeightOfTheSolid)/ nFloor;

                    if (minimalHeightOfWall <= WhichFloor & maximalHeightOfWall >= NextFloor)
                    {
                        List<IDFFile.XYZ> AllPointsOfWallPerFloor = new List<IDFFile.XYZ>();

                        foreach (IDFFile.XYZ wallpart in wall.verticesList.xyzs)
                        {
                            if( wallpart.Z <= WhichFloor )
                            {
                                IDFFile.XYZ wallpartToAdd = new IDFFile.XYZ();
                                wallpartToAdd.Z = WhichFloor;
                                wallpartToAdd.X = wallpart.X;
                                wallpartToAdd.Y = wallpart.Y;
                                AllPointsOfWallPerFloor.Add(wallpartToAdd);
                            }
                            if (wallpart.Z >= NextFloor)
                            {
                                IDFFile.XYZ wallpartToAdd = new IDFFile.XYZ();
                                wallpartToAdd.Z = NextFloor;
                                wallpartToAdd.X = wallpart.X;
                                wallpartToAdd.Y = wallpart.Y;
                                AllPointsOfWallPerFloor.Add(wallpartToAdd);
                            }
                        }
                        double AreaN = ((NextFloor - WhichFloor) / (maxHeightOfTheSolid - minHeightOfTheSolid)) * TotalArea;
                        IDFFile.XYZList XYZCoordsOfFloorWallXYZList = new IDFFile.XYZList(AllPointsOfWallPerFloor);
                        IDFFile.Building NewBuilding = new IDFFile.Building();
                        IDFFile.Zone NewZone = new IDFFile.Zone(NewBuilding, "FirstFloorWall", 1);
                        IDFFile.SurfaceType CurrentSurfType = IDFFile.SurfaceType.Wall;
                        IDFFile.BuildingSurface WallToAddPerFloor = new IDFFile.BuildingSurface(NewZone, XYZCoordsOfFloorWallXYZList, AreaN, CurrentSurfType);
                        FloorWalls[i].Add(WallToAddPerFloor);
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

        public static List<List<IDFFile.BuildingSurface>> GetFloorsAndCeilingPerFloor(List<IDFFile.BuildingSurface> TotalFloor, double maxHeightOfTheSolid, double minHeightOfTheSolid, int nFloor)
        {
            List<List<IDFFile.BuildingSurface>> FloorFloors = new List<List<IDFFile.BuildingSurface>>();
            for (int i = 0; i < nFloor; i++)
            {
                List<IDFFile.BuildingSurface> FloorNumber = new List<IDFFile.BuildingSurface>();
                FloorFloors.Add(FloorNumber);
            }
            List<IDFFile.XYZ> FloorToCopyThroughBuilding = new List<IDFFile.XYZ>();
            double Ground =0;
            double TotalArea=0;
            foreach (IDFFile.BuildingSurface Floor in TotalFloor)
            {
                TotalArea = Floor.area;
                List<IDFFile.XYZ> ZlocationOfFloor = Floor.verticesList.xyzs;
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
                IDFFile.Building NewBuilding = new IDFFile.Building();
                IDFFile.Zone NewZone = new IDFFile.Zone(NewBuilding, "FirstFloorWall", 1);
                IDFFile.SurfaceType CurrentSurfType = IDFFile.SurfaceType.Floor;
                IDFFile.BuildingSurface WallToAddPerFloor = new IDFFile.BuildingSurface(NewZone, XYZCoordsOfFloorWallXYZList, TotalArea, CurrentSurfType);
                FloorFloors[i].Add(WallToAddPerFloor);

            }
            return FloorFloors;
        }
 

        public static (IDFFile.BuildingSurface ,double,double ) GetIfFloorOrCeilingOrWall(Face PossibleWall, IDFFile.SurfaceType CurrentSurfaceType,double MaximalHeight, double MinimalHeight)
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
                    verteces.Z = Item[2];
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
                IDFFile.Building NewBuilding = new IDFFile.Building();
                IDFFile.Zone NewZone = new IDFFile.Zone(NewBuilding, "FirstFloorWall", 1);
                
                IDFFile.BuildingSurface WallAddedToBuildingSurface = new IDFFile.BuildingSurface(NewZone, pointList1, Arean, CurrentSurfaceType);

                return (WallAddedToBuildingSurface, MaximalHeight,MinimalHeight);

            
        }
        public static List<IDFFile.BuildingSurface> GetRoof(List<IDFFile.BuildingSurface> PossibleRoof ,double MaxH)
        {
            List<IDFFile.BuildingSurface> AllRoofs = new List<IDFFile.BuildingSurface>();
            IDFFile.SurfaceType CurrentSurfaceType = IDFFile.SurfaceType.Roof;
            foreach (IDFFile.BuildingSurface roofP in PossibleRoof)
            {
                IDFFile.XYZList RoofCoordinates = roofP.verticesList;
                double GetHeight = RoofCoordinates.xyzs[0].Z;
                if (GetHeight == MaxH)
                {
                    roofP.surfaceType = CurrentSurfaceType;
                    AllRoofs.Add(roofP);
                }

            }
            return AllRoofs;
        }
    }
    
}
