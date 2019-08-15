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

            string fullPath = doc.PathName;
            string folderPath = fullPath.Remove(fullPath.IndexOf('.'));

            try { Directory.CreateDirectory(folderPath); } catch { }
            Random r = new Random();

            InputData userData = new InputData();
            userData.ShowDialog();
            if (userData.cancel) { return Result.Failed; }

            List<IDFFile.BuildingConstruction> constructions = Utility.GenerateRandomConstruction(userData.pBuildingConstruction, userData.numSamples, r);
            List<IDFFile.WWR> WWRs = Utility.GenerateRandomWWR(userData.pWindowConstruction, userData.numSamples, r);
            List<IDFFile.BuildingOperation> buildingOperations = Utility.GenerateRandomBuildingParameters(userData.pBuildingOperation, userData.numSamples, r);

            File.WriteAllLines(string.Format("{0}/{1}.csv", folderPath, "Avg"), userData.pBuildingConstruction.ToCSVString());
            File.WriteAllLines(string.Format("{0}/{1}.csv", folderPath, "All"), constructions.Select(c=>c.ToCSVString()));

            Dictionary<string, IDFFile.Building> modelBuildings = Utility.GetAverageBuildingFromMasses(doc, userData.numFloors, userData.pBuildingConstruction.GetAverage(),
               userData.pWindowConstruction.GetAverage(), userData.pBuildingOperation.GetAverage() );

            foreach (KeyValuePair<string, IDFFile.Building> modelBuilding in modelBuildings)
            {
                IDFFile.IDFFile file = new IDFFile.IDFFile() { name = modelBuilding.Key, building = modelBuilding.Value };
                string fullFileName = string.Format("{0}/{1}.idf", folderPath, modelBuilding.Key);
                file.GenerateOutput(true, "Annual");
                File.WriteAllLines(fullFileName, file.WriteFile());

                for (int i=0; i<userData.numSamples; i++)
                {
                    file.building.UpdateBuildingConstructionWWROperations(constructions[i], WWRs[i], buildingOperations[i]);
                    fullFileName = string.Format("{0}/{1}_{2}.idf", folderPath, modelBuilding.Key, i.ToString());
                    File.WriteAllLines(fullFileName, file.WriteFile());
                }
            }
            return Result.Succeeded;
        }
    }

    public static class Utility
    {
        public static double FtToM(double value)
        {
            return (Math.Round(value * 0.3048, 5));
        }
        public static double SqFtToSqM(double value)
        {
            return (Math.Round(value * 0.092903, 5));
        }      
        public static XYZ ftToM(this XYZ point)
        {
            return new XYZ(FtToM(point.X), FtToM(point.Y), FtToM(point.Z));
        }
        public static Dictionary<string, IDFFile.Building> GetAverageBuildingFromMasses(Document doc, int nFloors, IDFFile.BuildingConstruction constructions,
        IDFFile.WWR WWR, IDFFile.BuildingOperation buildingOperations )
        {
            Dictionary<string, IDFFile.Building> buildings = new Dictionary<string, IDFFile.Building>();

            List<View3D> views = (new FilteredElementCollector(doc).OfClass(typeof(View3D)).Cast<View3D>()).Where(v => v.Name.Contains("Op - ")).ToList();
            foreach (View3D v1 in views)
            {
                //Initialize building elements
                FilteredElementCollector masses = new FilteredElementCollector(doc, v1.Id).WhereElementIsNotElementType().OfCategory(BuiltInCategory.OST_Mass);
                List<IDFFile.XYZ> groundPoints = new List<IDFFile.XYZ>();
                List<IDFFile.XYZ> roofPoints = new List<IDFFile.XYZ>();
                double floorArea = 0;
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
                                foreach (Face face1 in AllFacesFromModel)
                                {
                                    XYZ fNormal = (face1 as PlanarFace).FaceNormal;

                                    //checks if it is indeed a wall by computing the normal with respect to 001                                   
                                    if (Math.Round(fNormal.Z, 3) == -1)
                                    {
                                        groundPoints.AddRange(GetPoints(face1));
                                        floorArea = SqFtToSqM(face1.Area);
                                    }
                                    if (Math.Round(fNormal.Z, 3) == 1)
                                    {
                                        roofPoints.AddRange(GetPoints(face1));
                                    }
                                }                                                            
                            }
                        }
                    }
                }
                IDFFile.Building building = InitialiseModelBuilding(groundPoints, roofPoints, floorArea, nFloors, constructions, WWR, buildingOperations );

                buildings.Add(v1.Name.Remove(0, 4), building);
            }
            return buildings;
        }
        public static List<Tuple<IDFFile.XYZ, IDFFile.XYZ>> GetAllWallEdges(List<IDFFile.XYZ> groundPoints)
        {
            List<Tuple<IDFFile.XYZ, IDFFile.XYZ>> wallEdges = new List<Tuple<IDFFile.XYZ, IDFFile.XYZ>>();
            for (int i = 0; i<groundPoints.Count; i++)
            {
                try
                {
                    wallEdges.Add(new Tuple<IDFFile.XYZ, IDFFile.XYZ>( groundPoints[i], groundPoints[i + 1]));
                }
                catch
                {
                    wallEdges.Add(new Tuple<IDFFile.XYZ, IDFFile.XYZ>(groundPoints[i], groundPoints[0]));
                }
            }
            return wallEdges;
        }     
        public static List<IDFFile.BuildingConstruction> GenerateRandomConstruction(IDFFile.ProbabilisticBuildingConstruction pConstruction, int numsamples, Random r)
        {
            List<IDFFile.BuildingConstruction> constructions = new List<IDFFile.BuildingConstruction>();
            for (int i = 0; i < numsamples; i++)
            {
                constructions.Add(new IDFFile.BuildingConstruction()
                {
                    uWall = GetRandom(pConstruction.uWall, r),
                    uGFloor = GetRandom(pConstruction.uGFloor, r),
                    uRoof = GetRandom(pConstruction.uRoof, r),
                    uWindow = GetRandom(pConstruction.uWindow, r),
                    gWindow = GetRandom(pConstruction.gWindow, r),
                    uIWall = GetRandom(pConstruction.uIWall, r),
                    uIFloor = GetRandom(pConstruction.uIFloor, r),
                    hcSlab = GetRandom(pConstruction.hcSlab, r),
                    infiltration = GetRandom(pConstruction.infiltration, r),
                });              
            }            
            return constructions;
        }
        public static List<IDFFile.WWR> GenerateRandomWWR(IDFFile.ProbabilisticWWR pWWR, int numsamples, Random r)
        {
            List<IDFFile.WWR> WWRs = new List<IDFFile.WWR>();
            for (int i = 0; i < numsamples; i++)
            {
                WWRs.Add(new IDFFile.WWR()
                {
                    east = GetRandom(pWWR.east, r),
                    south = GetRandom(pWWR.south, r),
                    north = GetRandom(pWWR.north, r),
                    west = GetRandom(pWWR.west, r)                
                });
                
            }
            return WWRs;
        }
        public static List<IDFFile.BuildingOperation> GenerateRandomBuildingParameters(IDFFile.ProbabilisticBuildingOperation pOP, int numsamples, Random r)
        {
            List<IDFFile.BuildingOperation> bParameters = new List<IDFFile.BuildingOperation>();
            for (int i = 0; i < numsamples; i++)
            {
                bParameters.Add(new IDFFile.BuildingOperation()
                {
                    internalHeatGain = GetRandom(pOP.internalHeatGain, r),
                    operatingHours = GetRandom(pOP.operatingHours, r),
                    boilerEfficiency = GetRandom(pOP.boilerEfficiency, r),
                    chillerCOP = GetRandom(pOP.chillerCOP, r)
                });

            }
            return bParameters;
        }
        public static IDFFile.Building InitialiseModelBuilding(List<IDFFile.XYZ> groundPoints, List<IDFFile.XYZ> roofPoints, double area, int nFloors,
            IDFFile.BuildingConstruction constuction, IDFFile.WWR wwr, IDFFile.BuildingOperation operation)
        {
            IDFFile.ZoneList zoneList = new IDFFile.ZoneList("Office");

            double[] heatingSetPoints = new double[] { 10, 20 };
            double[] coolingSetPoints = new double[] { 28, 24 };
            double equipOffsetFraction = 0.1;
            IDFFile.Building bui = new IDFFile.Building
            {
                buildingConstruction = constuction,
                WWR = wwr,
                buildingOperation = operation
            };
            bui.UpdataBuildingOperations();
            bui.CreateSchedules(heatingSetPoints, coolingSetPoints, equipOffsetFraction);
            bui.GenerateConstructionWithIComponentsU();
            IDFFile.ScheduleCompact Schedule = bui.schedulescomp.First(s=>s.name.Contains("Occupancy"));

            double baseZ = groundPoints.First().Z;
            double roofZ = roofPoints.First().Z;
            double heightFl = (roofZ - baseZ) / nFloors;

            IDFFile.XYZList dlPoints = GetDayLightPointsXYZList(groundPoints, GetAllWallEdges(groundPoints));
            for (int i = 0; i < nFloors; i++)
            {
                double floorZ = baseZ + i * heightFl;
                IDFFile.XYZList floorPoints = new IDFFile.XYZList(groundPoints).ChangeZValue(floorZ);
                IDFFile.Zone zone = new IDFFile.Zone(bui, "Zone_" + i, i);
                zone.name = "Zone_" + i;
                IDFFile.People newpeople = new IDFFile.People(10);
                zone.people = newpeople;
                if (i == 0)
                {
                    IDFFile.BuildingSurface floor = new IDFFile.BuildingSurface(zone, floorPoints, area, IDFFile.SurfaceType.Floor);
                }
                else
                {
                    IDFFile.BuildingSurface floor = new IDFFile.BuildingSurface(zone, floorPoints, area, IDFFile.SurfaceType.Floor)
                    {
                        ConstructionName = "General_Floor_Ceiling",
                        OutsideCondition = "Zone", OutsideObject = "Zone_" + (i-1).ToString()
                    };
                }
                floorPoints.createWalls(zone, heightFl);
                if (i == nFloors - 1)
                {
                    roofPoints.Reverse();
                    IDFFile.BuildingSurface roof = new IDFFile.BuildingSurface(zone, new IDFFile.XYZList(roofPoints), area, IDFFile.SurfaceType.Roof);
                }     

                IDFFile.DayLighting DayPoints = new IDFFile.DayLighting(zone, Schedule, dlPoints.ChangeZValue(floorZ+0.9).xyzs, 500);

                bui.AddZone(zone);
                zoneList.listZones.Add(zone);
            }
            bui.AddZoneList(zoneList);

            bui.GeneratePeopleLightingElectricEquipment();
            bui.GenerateInfiltraitionAndVentillation();
            bui.GenerateHVAC(true, false, false);
            return bui;
        }
        public static List<IDFFile.XYZ> GetPoints(Face face)
        {
            List<IDFFile.XYZ> vectorList = new List<IDFFile.XYZ>();
            EdgeArray edgeArray = face.EdgeLoops.get_Item(0);
            foreach (Edge e in edgeArray)
            {
                vectorList.Add(e.AsCurveFollowingFace(face).GetEndPoint(0).ftToM().ConvertToIDF());
            }
            return vectorList;
        }
        public static IDFFile.XYZList GetDayLightPointsXYZList(List<IDFFile.XYZ> FloorFacePoints, List<Tuple<IDFFile.XYZ, IDFFile.XYZ>> WallEdges)
        {
            List<IDFFile.XYZ> CentersOfMass = TriangulateAndGetCenterOfMass(FloorFacePoints);
            return new IDFFile.XYZList(CentersOfMass.Where(p => RayCastToCheckIfIsInside(WallEdges, p)).ToList());
        }
        public static IDFFile.XYZ CenterOfMass(List<IDFFile.XYZ> points)
        {
            return new IDFFile.XYZ()
            {
                X = points.Select(p=>p.X).Average(),
                Y = points.Select(p => p.Y).Average(),
                Z = points.Select(p => p.Z).Average(),
            };
        }
        public static List<IDFFile.XYZ> TriangulateAndGetCenterOfMass(List<IDFFile.XYZ> AllPoints)
        {
            int[] PointNumbers = Enumerable.Range(-1, AllPoints.Count() + 2).ToArray();

            PointNumbers[0] = AllPoints.Count() - 1;
            PointNumbers[PointNumbers.Length - 1] = 0;
            List<IDFFile.XYZ> CentersOfMass = new List<IDFFile.XYZ>();

            for (int i= 0; i<AllPoints.Count(); i++)
            {
                IDFFile.XYZ pointCM = new IDFFile.XYZ();
                try
                {
                    pointCM = CenterOfMass(new List<IDFFile.XYZ>() { AllPoints[i - 1], AllPoints[i], AllPoints[i + 1] });
                }
                catch
                {
                    try
                    {
                        pointCM = CenterOfMass(new List<IDFFile.XYZ>() { AllPoints.Last(), AllPoints[i], AllPoints[i + 1] });
                    }
                    catch
                    {
                        pointCM = CenterOfMass(new List<IDFFile.XYZ>() { AllPoints[i - 1], AllPoints[i], AllPoints.First() });
                    }
                }
                CentersOfMass.Add(pointCM);
            }
            CentersOfMass.Add(CenterOfMass(AllPoints));
            return CentersOfMass;
        }
        public static bool RayCastToCheckIfIsInside(List<Tuple<IDFFile.XYZ, IDFFile.XYZ>> EdgeArrayForPossibleWall, IDFFile.XYZ CM)
        {
            int count = 0;
            foreach (Tuple<IDFFile.XYZ, IDFFile.XYZ> EdgeOfWall in EdgeArrayForPossibleWall)
            {

                double r = (CM.Y - EdgeOfWall.Item2.Y) / (EdgeOfWall.Item1.Y - EdgeOfWall.Item2.Y);
                if (r > 0 && r < 1)
                {
                    double Xvalue = r * (EdgeOfWall.Item1.X - EdgeOfWall.Item2.X) + EdgeOfWall.Item2.X;
                    if (CM.X < Xvalue)
                    {
                        count++;
                    }
                }
            }
            if (count % 2 == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        public static IDFFile.XYZ ConvertToIDF(this XYZ point)
        {
            return new IDFFile.XYZ(point.X, point.Y, point.Z);
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
            return range[0] + ((range[1] - range[0]) * r.NextDouble());
        }
    }
}
