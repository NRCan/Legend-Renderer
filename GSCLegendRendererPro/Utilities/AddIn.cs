using ArcGIS.Desktop.Editing.Attributes;
using ArcGIS.Desktop.Internal.Editing.COGO;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using System.Xml;

namespace GSCLegendRendererPro.Utilities
{
    /// <summary>
    /// Encapsulated AddIn metadata
    /// https://github.com/Esri/arcgis-pro-sdk-community-samples/blob/master/Content/AddInInfoManager/AddIn.cs
    /// </summary>
    public class AddIn
    {
        public string Name { get; set; }
        public string Version { get; set; }
        public string DesktopVersion { get; set; }
        public string AddInPath { get; set; }
        public string Id { get; set; }
        public string AddInDate { get; set; }

        /// <summary>
        /// Get the current add-in module's daml / AddInInfo Id tag (which is the same as the Assembly GUID)
        /// </summary>
        /// <returns></returns>
        public static string GetAddInId()
        {
            string fileName = string.Empty;

            // Module.Id is internal, but we can still get the ID from the assembly
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            object[] guidAttribute = assembly.GetCustomAttributes(typeof(GuidAttribute), true);
            if (guidAttribute != null && guidAttribute.Count() > 0)
            {
                var attribute = (GuidAttribute)assembly.GetCustomAttributes(typeof(GuidAttribute), true)[0];
                fileName = Path.Combine($@"{{{attribute.Value.ToString()}}}", $@"{assembly.FullName.Split(',')[0]}.esriAddInX");
            }
            else
            {
                //Get addin guid with regex
                var result = Regex.Match(
                      assembly.Location,
                      @"[({]?[a-fA-F0-9]{8}[-]?([a-fA-F0-9]{4}[-]?){3}[a-fA-F0-9]{12}[})]?",
                      RegexOptions.IgnoreCase
                );
                fileName = Path.Combine($@"{result}", $@"{assembly.FullName.Split(',')[0]}.esriAddInX");
            }

            return fileName;
        }

        /// <summary>
        /// returns a tuple with version and desktopVersion using the given addin file path
        /// </summary>
        /// <param name="fileName">file path (partial) of esriAddinX package</param>
        /// <returns>tuple: version, desktopVersion</returns>
        public static AddIn GetConfigDamlAddInInfo(string fileName)
        {
            var esriAddInX = new AddIn();
            XmlDocument xDoc = new XmlDocument();
            var esriAddInXPath = FindEsriAddInXPath(fileName);

            try
            {
                esriAddInX.AddInPath = esriAddInXPath;
                using (ZipArchive zip = ZipFile.OpenRead(esriAddInXPath))
                {
                    ZipArchiveEntry zipEntry = zip.GetEntry("Config.daml");
                    MemoryStream ms = new MemoryStream();
                    string daml = string.Empty;
                    using (Stream stmZip = zipEntry.Open())
                    {
                        StreamReader streamReader = new StreamReader(stmZip);
                        daml = streamReader.ReadToEnd();
                        xDoc.LoadXml(daml); // @"<?xml version=""1.0"" encoding=""utf - 8""?>" + 
                    }
                }
                XmlNodeList items = xDoc.GetElementsByTagName("AddInInfo");
                foreach (XmlNode xItem in items)
                {
                    esriAddInX.Version = xItem.Attributes["version"].Value;
                    esriAddInX.DesktopVersion = xItem.Attributes["desktopVersion"].Value;
                    esriAddInX.Id = xItem.Attributes["id"].Value;
                    esriAddInX.Name = "N/A";
                    esriAddInX.AddInDate = "N/A";
                    foreach (XmlNode xChild in xItem.ChildNodes)
                    {
                        switch (xChild.Name)
                        {
                            case "Name":
                                esriAddInX.Name = xChild.InnerText;
                                break;
                            case "Image":
                                break;
                            case "Date":
                                esriAddInX.AddInDate = xChild.InnerText;
                                break;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                throw new Exception($@"Unable to parse config.daml {esriAddInXPath}: {ex.Message}");
            }
            return esriAddInX;
        }

        private static readonly string AddInSubFolderPath = @"ArcGIS\AddIns\ArcGISPro";
        private static string FindEsriAddInXPath(string fileName)
        {
            string defaultAddInPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), AddInSubFolderPath);
            string thePath = Path.Combine(defaultAddInPath, fileName);
            if (File.Exists(thePath)) return thePath;
            foreach (var addinPath in GetAddInFolders())
            {
                thePath = Path.Combine(addinPath, fileName);
                if (File.Exists(thePath)) return thePath;
            }
            throw new FileNotFoundException($@"esriAddInX file for {fileName} was not found");
        }

        /// <summary>
        /// Returns the list of all Addins
        /// </summary>
        /// <returns></returns>
        public static List<AddIn> GetAddIns()
        {
            List<AddIn> addIns = new List<AddIn>();
            List<string> lstPaths = GetAddInFolders();
            string defaultAddInPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), AddInSubFolderPath);
            lstPaths.Insert(0, defaultAddInPath);
            foreach (var addinPath in lstPaths)
            {
                foreach (var addinDirs in Directory.GetDirectories(addinPath))
                {
                    foreach (var addinFile in Directory.GetFiles(addinDirs, "*.esriAddinX"))
                    {
                        addIns.Add(GetConfigDamlAddInInfo(addinFile));
                    }
                }
            }
            return addIns;
        }

        /// <summary>
        /// Gets the well-known Add-in folders on the machine
        /// </summary>
        /// <returns>List of all well-known add-in folders</returns>
        public static List<string> GetAddInFolders()
        {
            List<string> myAddInPathKeys = new List<string>();
            string regPath = string.Format(@"Software\ESRI\ArcGISPro\Settings\Add-In Folders");
            //string path = "";
            string err1 = "This is an error";
            try
            {
                Microsoft.Win32.RegistryKey localKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.LocalMachine, Microsoft.Win32.RegistryView.Registry64);
                Microsoft.Win32.RegistryKey esriKey = localKey.OpenSubKey(regPath);
                if (esriKey == null)
                {
                    localKey = Microsoft.Win32.RegistryKey.OpenBaseKey(Microsoft.Win32.RegistryHive.CurrentUser, Microsoft.Win32.RegistryView.Registry64);
                    esriKey = localKey.OpenSubKey(regPath);
                }
                if (esriKey != null)
                    myAddInPathKeys.AddRange(esriKey.GetValueNames().Select(key => key.ToString()));
            }
            catch (InvalidOperationException ie)
            {
                //this is ours
                throw ie;
            }
            catch (Exception ex)
            {
                throw new System.Exception(err1, ex);
            }
            return myAddInPathKeys;
        }

    }
    
}
