using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml;

namespace HarvestArcGISComponentCategories
{
    public class Program
    {
        private const string _zipFileName = "Config.xml";

        internal static string GetAssemblyShortName(Assembly assembly)
        {
            string location = assembly.Location;
            string assemblyShortName = Path.GetFileNameWithoutExtension(location);
            return assemblyShortName;
        }

        internal static string GetFolderName(string assemblyName, Guid assemblyGuid)
        {
            return string.Format("{0}_{1}.ecfg",
                                 assemblyGuid.ToString("B"),
                                 assemblyName);
        }

        private static void CreateCsv(HarvestResults componentCategories, string assemblyPath, string txtPath)
        {
            Console.WriteLine("********************************");
            Console.WriteLine($"Assembly = {assemblyPath}");

            // Get assembly
            Assembly assembly = Assembly.LoadFrom(assemblyPath);
            Console.WriteLine($"{assembly.FullName}");
            Type[] types = assembly.GetExportedTypes();

            // map clsid to type
            Dictionary<string, Type> typeMap = new Dictionary<string, Type>();
            foreach (Type type in types)
            {
                object[] customAttributes = type.GetCustomAttributes(typeof(GuidAttribute), false);
                if (customAttributes.Length <= 0)
                {
                    //Console.WriteLine($"{type.Name} = ???");
                    continue;
                }

                var guidAttribute = (GuidAttribute)customAttributes[0];
                string classGuidString = guidAttribute.Value.ToUpper();
                typeMap[classGuidString] = type;
                Console.WriteLine($"{type.Name} = {classGuidString}");
            }

            // Category names
            Dictionary<string, string> catMap = new Dictionary<string, string>();
            catMap.Add("B56A7C42-83D4-11D2-A2E9-080009B6F22B", "MxCommands");
            catMap.Add("B56A7C4A-83D4-11D2-A2E9-080009B6F22B", "MxCommandBars");
            catMap.Add("B56A7C45-83D4-11D2-A2E9-080009B6F22B", "MxExtensions");
            catMap.Add("117623B5-F9D1-11D3-A67F-0008C7DF97B9", "MxDockableWindows");
            catMap.Add("D4E2A322-5D59-11D2-89FD-006097AFF44E", "GeoObjectClassExtensions");
            catMap.Add("0813623A-A72E-11D2-8924-0000F877762D", "GeoObjectClassDescriptions");
            catMap.Add("5F08CBCA-E91F-11D1-AEE8-080009EC734B", "GxCommands ");

            List <ComponentInfo> list = new List<ComponentInfo>();
            foreach (
                KeyValuePair<Guid, IList<Guid>> componentCategory in
                    componentCategories.HarvestedRegistryValues)
            {
                // Category
                string categoryGuid = componentCategory.Key.ToString().ToUpper();
                string categoryName = "Unknown";
                if (catMap.ContainsKey(categoryGuid))
                {
                    categoryName = catMap[categoryGuid];
                    Console.WriteLine($"Category {categoryName} = {categoryGuid}");
                }
                else
                {
                    Console.WriteLine($"Missing CatID {categoryGuid}");
                }

                // Class ID in Category
                foreach (Guid classId in componentCategory.Value)
                {
                    string clsid = classId.ToString().ToUpper();
                    if (!typeMap.ContainsKey(clsid))
                    {
                        Console.WriteLine($"Missing CLSID {clsid}");
                        continue;
                    }

                    // Get Class from CLSID
                    Type type = typeMap[clsid];


                    // New Info
                    ComponentInfo info = new ComponentInfo();
                    info.CategoryGuid = categoryGuid;
                    info.CategoryName = categoryName;
                    info.ClassGuid = clsid;
                    info.ClassName = type.Name;
                    info.BaseClassName = type.BaseType.Name;

                    // Try get ProgID
                    object[] customAttributes = type.GetCustomAttributes(typeof(ProgIdAttribute), false);
                    if (customAttributes.Length <= 0) continue;
                    var progIdAttribute = (ProgIdAttribute)customAttributes[0];
                    string progId = progIdAttribute.Value;
                    info.ProgId = progId;

                    Console.WriteLine($"{info.ClassName} = {info.CategoryName}");
                    list.Add(info);
                }
            }

            // Generate the CSV file...
            StringBuilder sb = new StringBuilder();
            string header = $"Class,Base Class,ProgID,CLSID,Category,CatID";
            sb.AppendLine(header);
            var sorted = list.OrderBy(i => i.ProgId).OrderBy(i => i.CategoryName);
            foreach (var info in sorted)
            {
                string line = $"{info.ClassName},{info.BaseClassName},{info.ProgId},{info.ClassGuid},{info.CategoryName},{info.CategoryGuid}";
                sb.AppendLine(line);
            }
            File.WriteAllText(txtPath, sb.ToString());
            Console.WriteLine($"Created file {txtPath}");
        }

        private static void CreateXml(HarvestResults componentCategories, string zipPath)
        {
            var xmlWriter = new XmlTextWriter(zipPath, null);
            xmlWriter.Formatting = Formatting.Indented;

            xmlWriter.WriteStartDocument();

            xmlWriter.WriteStartElement("ESRI.Configuration");
            xmlWriter.WriteAttributeString("ver", "1");

            xmlWriter.WriteStartElement("Categories");

            foreach (
                KeyValuePair<Guid, IList<Guid>> componentCategory in
                    componentCategories.HarvestedRegistryValues)
            {
                xmlWriter.WriteStartElement("Category");
                xmlWriter.WriteAttributeString("CATID",
                                               "{" + componentCategory.Key.ToString().ToUpper() + "}");

                Console.WriteLine("Component Category: {0}", componentCategory.Key);

                foreach (Guid classId in componentCategory.Value)
                {
                    xmlWriter.WriteStartElement("Class");
                    xmlWriter.WriteAttributeString("CLSID", classId.ToString("B").ToUpper());
                    xmlWriter.WriteEndElement();

                    Console.WriteLine("   <Class CLSID=\"" + classId.ToString("B").ToUpper());
                }
                xmlWriter.WriteEndElement();
            }
            xmlWriter.WriteEndElement();
            xmlWriter.WriteEndElement();
            xmlWriter.WriteEndDocument();

            xmlWriter.Close();
        }

        private static HarvestResults HarvestRegistryValues(string path)
        {
            var regSvcs = new RegistrationServices();
            Assembly assembly = Assembly.LoadFrom(path);

            object[] customAttributes = assembly.GetCustomAttributes(typeof(GuidAttribute), false);

            if (customAttributes.Length <= 0)
            {
                Console.WriteLine(string.Format("Assembly {0} does not have a GUID", assembly.FullName));

                return null;
            }

            var guidAttribute = (GuidAttribute)customAttributes[0];
            string assemblyGuidString = guidAttribute.Value;

            var assemblyGuid = new Guid(assemblyGuidString);

            Console.WriteLine("Assembly {0}_{1}", assemblyGuid.ToString("B"),
                              GetAssemblyShortName(assembly));

            try
            {
                // must call this before overriding registry hives to prevent
                // binding failures on exported types during RegisterAssembly
                assembly.GetExportedTypes();
            }
            catch (Exception)
            {
                Console.WriteLine(
                    "Error getting types from assembly. Make sure the referenced assemblies exist in the output folder.");
                throw;
            }

            const bool remapRegistration = true;

            using (var componentCategoryHarvester = new ComponentCategoryHarvester(remapRegistration))
            {
                regSvcs.RegisterAssembly(assembly, AssemblyRegistrationFlags.SetCodeBase);

                return new HarvestResults(assemblyGuid,
                                          GetAssemblyShortName(assembly),
                                          componentCategoryHarvester.HarvestRegistry());
            }
        }

        private static void Main(string[] args)
        {
            string inputAssemblyFileName = null;

            try
            {
                string outputFolder;

                if (args.Length == 0)
                {
                    Console.WriteLine("ERROR: Incorrect number of arguments.");
                    Console.WriteLine();
                    Console.WriteLine(
                        "Usage: HarvestArcGISCategories.exe <input assembly> {output folder}");

                    return;
                }
                if (args.Length == 1)
                {
                    inputAssemblyFileName = args[0];
                    outputFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                }
                else
                {
                    inputAssemblyFileName = args[0];
                    outputFolder = args[1];
                }

                Console.WriteLine(string.Format("Harvesting Categories for {0}. Output folder: {1}",
                                                inputAssemblyFileName, outputFolder));

                if (!PathsValid(inputAssemblyFileName, outputFolder))
                {
                    return;
                }

                Console.WriteLine("Execution Path: " + outputFolder);

                HarvestResults componentCategories = HarvestRegistryValues(inputAssemblyFileName);

                if (componentCategories == null)
                {
                    return;
                }

                string folderName = GetFolderName(componentCategories.AssemblyName,
                                                  componentCategories.AssemblyGuid);

                string dirPath = string.Format("{0}/{1}", outputFolder, folderName);

                if (File.Exists(dirPath))
                {
                    File.Delete(dirPath);
                }

                if (!Directory.Exists(dirPath))
                {
                    Directory.CreateDirectory(dirPath);
                }

                string zipPath = string.Format("{0}/{1}/{2}", outputFolder, folderName, _zipFileName);

                string txtPath = $"{outputFolder}/{componentCategories.AssemblyName}_categories.csv";

                // Hacked in a CSV output
                CreateCsv(componentCategories, inputAssemblyFileName, txtPath);

                // Make the ecfg file
                if (false)
                {
                    CreateXml(componentCategories, zipPath);
                    string zipFolder = string.Format("{0}/{1}", outputFolder, folderName);
                    string zipSaveName = string.Format("{0}.zip", zipFolder);

                    using (var zip = new ZipFile())
                    {
                        zip.AddDirectory(zipFolder);
                        zip.Save(zipSaveName);
                    }

                    File.Delete(zipPath);

                    Directory.Delete(zipFolder);

                    File.Move(zipSaveName, zipFolder);
                }
            }
            catch (Exception ex)
            {
                // TODO: find how to communicate with msbuild that it failed and to print the error in red
                Console.WriteLine(string.Format("Error Harvesting Categories in {0}: {1}",
                                                inputAssemblyFileName, ex.Message));
                Console.WriteLine(ex.ToString());

                // NOTE: throw results in the crash-dialog to come up... but at
                //       least msbuild realises that there was an error
                throw;
            }
        }

        private static bool PathsValid(string inputAssemblyFileName, string outputFolder)
        {
            if (!File.Exists(inputAssemblyFileName))
            {
                //TODO Error handling

                Console.WriteLine(string.Format("File {0} does not exist", inputAssemblyFileName));

                return false;
            }

            if (outputFolder != null && !Directory.Exists(outputFolder))
            {
                //TODO Error handling

                Console.WriteLine(string.Format("Folder {0} does not exist", outputFolder));

                return false;
            }
            return true;
        }
    }
}