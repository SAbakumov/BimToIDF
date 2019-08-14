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

            InputData userData = new InputData();
            userData.ShowDialog();
            Dictionary<string, IDFFile.Building> sampledData = Utility.GetAverageBuildingFromMasses(doc,userData);
            foreach (KeyValuePair<string, IDFFile.Building> sample in sampledData)
            {
                IDFFile.IDFFile file = new IDFFile.IDFFile() { name = sample.Key, building = sample.Value };
                string fullFileName = string.Format("{0}/{1}.idf", folderPath, sample.Key);
                file.GenerateOutput(true, "Annual");
                File.WriteAllLines(fullFileName, file.WriteFile());
            }

            foreach (KeyValuePair<string, IDFFile.Building> sample in sampledData)
            {
                IDFFile.IDFFile file = new IDFFile.IDFFile() { name = sample.Key, building = sample.Value };
                IDFFile.Building Bui = sample.Value;

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
                    string fullFileName = string.Format("{0}/{1}_{2}.idf", folderPath, sample.Key, i.ToString());
                    IDFFileSample.GenerateOutput(false, "Annual");
                    File.WriteAllLines(fullFileName, IDFFileSample.WriteFile());
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
        public static Dictionary<string, IDFFile.Building> GetAverageBuildingFromMasses(Document doc, InputData userData)
        {
            int numberofFloors = userData.numFloors;
            Dictionary<string, IDFFile.Building> structures = new Dictionary<string, IDFFile.Building>();

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
                IDFFile.Building building = InitialiseAverageBuilding(userData, groundPoints, roofPoints, floorArea);

                structures.Add(v1.Name.Remove(0, 4), building);
            }
            return structures;
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
        public static IDFFile.Building InitialiseAverageBuilding(InputData userData, List<IDFFile.XYZ> groundPoints, List<IDFFile.XYZ> roofPoints, double area)
        {
            IDFFile.ZoneList zoneList = new IDFFile.ZoneList("Office");

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
            IDFFile.ScheduleCompact Schedule = bui.schedulescomp.First(s=>s.name.Contains("Occupancy"));

            double baseZ = groundPoints.First().Z;
            double roofZ = roofPoints.First().Z;

            double heightFl = (roofZ - baseZ) / userData.numFloors;

            IDFFile.XYZList dlPoints = GetDayLightPointsXYZList(groundPoints, GetAllWallEdges(groundPoints));

            for (int i = 0; i < userData.numFloors; i++)
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
                    IDFFile.BuildingSurface floor = new IDFFile.BuildingSurface(zone, floorPoints, area, IDFFile.SurfaceType.InternalFloor) { OutsideObject = "Zone_" + (i-1).ToString()};
                }

                if (i == userData.numFloors - 1)
                {
                    IDFFile.BuildingSurface roof = new IDFFile.BuildingSurface(zone, new IDFFile.XYZList(roofPoints), area, IDFFile.SurfaceType.Floor) { OutsideObject = "Zone_" + (i - 1).ToString() };
                }

                floorPoints.createWalls(zone, heightFl);

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

            List<IDFFile.XYZ> DayLightPoints = new List<IDFFile.XYZ>();
            List<IDFFile.XYZ> VectorList = new List<IDFFile.XYZ>();
            foreach (IDFFile.XYZ Point in FloorFacePoints)
            {
                double xcoord = Math.Round(Point.X, 4);
                double ycoord = Math.Round(Point.Y, 4);
                double zcoord = Math.Round(Point.Z, 4);
                IDFFile.XYZ NewVector = new IDFFile.XYZ (xcoord, ycoord, zcoord);
                VectorList.Add(NewVector);
            }
            IDFFile.XYZ[] AllPoints = new IDFFile.XYZ[VectorList.Count];
            for (int i = 0; i < VectorList.Count; i++)
            {
                AllPoints[i] = VectorList[i];

            }
            IDFFile.XYZ[] CentersOfMass = TriangulateAndGetCenterOfMass(AllPoints);
            
            foreach (IDFFile.XYZ CM in CentersOfMass)
            {
                if (RayCastToCheckIfIsInside(WallEdges, CM))
                {
                    DayLightPoints.Add(DeepClone(CM));
                }
            }

            //IDFFile.XYZList AllDayLightPoints = new IDFFile.XYZList(DayLightPoints);
            return new IDFFile.XYZList(DayLightPoints);
        }
        public static IDFFile.XYZ CenterOfMass(IDFFile.XYZ[] points)
        {
            return new IDFFile.XYZ()
            {
                X = points.Select(p=>p.X).Average(),
                Y = points.Select(p => p.Y).Average(),
                Z = points.Select(p => p.Z).Average(),
            };
        }
        public static IDFFile.XYZ[] TriangulateAndGetCenterOfMass(IDFFile.XYZ[] AllPoints)
        {
            int[] PointNumbers = Enumerable.Range(-1, AllPoints.Length + 2).ToArray();

            PointNumbers[0] = AllPoints.Length - 1;
            PointNumbers[PointNumbers.Length - 1] = 0;
            IDFFile.XYZ[] CentersOfMass = new IDFFile.XYZ[AllPoints.Length];

            for (int i = 1; i < PointNumbers.Length - 1; i++)
            {
                IDFFile.XYZ[] Triangle = new IDFFile.XYZ[3];
                Triangle[0] = AllPoints[PointNumbers[i - 1]];
                Triangle[1] = AllPoints[PointNumbers[i]];
                Triangle[2] = AllPoints[PointNumbers[i + 1]];

                IDFFile.XYZ cM = CenterOfMass(Triangle);
                CentersOfMass[i - 1] = cM;
            }
            CentersOfMass.Append(CenterOfMass(AllPoints));
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
        public static List<Tuple<IDFFile.XYZ, IDFFile.XYZ>> GetEdgesOfPolygon(PlanarFace PossibleWall)
        {

            List<Tuple<IDFFile.XYZ, IDFFile.XYZ>> EdgeLoop = new List<Tuple<IDFFile.XYZ, IDFFile.XYZ>>();

            EdgeArrayArray edgeArrays = PossibleWall.EdgeLoops;
            foreach (EdgeArray edges in edgeArrays)
            {
                foreach (Edge edge in edges)
                {
                    // Get one test point
                    XYZ Point1 = edge.Evaluate(1).ftToM();
                    XYZ Point2 = edge.Evaluate(0).ftToM();
                    IDFFile.XYZ Vertex1 = Point1.ConvertToIDF();
                    IDFFile.XYZ Vertex2 = Point2.ConvertToIDF();

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
