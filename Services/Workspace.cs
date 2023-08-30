using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.GeoDatabaseDistributed;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Catalog;
using ESRI.ArcGIS.DataSourcesOleDB;

namespace GSC_Legend_Renderer.Services
{
    public class Workspace
    {
        #region GET METHODS

        /// <summary>
        /// Will return a list of dataset that resides inside a workspace, like a file geodatabase
        /// </summary>
        /// <param name="inputWorkspace">The workspace to get a list of dataset from</param>
        /// <returns></returns>
        public static List<IDataset> GetDatasetListFromWorkspace(IWorkspace inputWorkspace)
        {
            //Variables
            List<IDataset> dataList = new List<IDataset>();

            //Get collection of datasets
            IEnumDataset enumData = inputWorkspace.get_Datasets(esriDatasetType.esriDTAny);

            //Fill the list
            IDataset currentData = enumData.Next();
            while (currentData != null)
            {
                dataList.Add(currentData);


                //Check for different inside feature dataset
                if (currentData.Category.Contains("Dataset"))
                {
                    //Cast as feature dataset
                    IFeatureDataset featDataset = currentData as IFeatureDataset;
                    IEnumDataset enumFeatDataset = featDataset.Subsets;

                    IDataset currentFData = enumFeatDataset.Next();
                    while (currentFData != null)
                    {
                        dataList.Add(currentFData);
                        currentFData = enumFeatDataset.Next();
                    }

                }

                currentData = enumData.Next();
            }

            return dataList;
        }

        /// <summary>
        /// Will return a list of dataset browse name that resides inside a workspace, like a file geodatabase
        /// </summary>
        /// <param name="inputWorkspace">The workspace to get a list of dataset from</param>
        /// <returns></returns>
        public static List<string> GetDatasetNameListFromWorkspace(IWorkspace inputWorkspace)
        {
            //Variables
            List<string> dataList = new List<string>();

            //Get collection of datasets
            IEnumDataset enumData = inputWorkspace.get_Datasets(esriDatasetType.esriDTAny);

            //Fill the list
            IDataset currentData = enumData.Next();
            while (currentData != null)
            {
                dataList.Add(currentData.BrowseName);

                //Check for different inside feature dataset
                if (currentData.Category.Contains("Dataset"))
                {
                    //Cast as feature dataset
                    IFeatureDataset featDataset = currentData as IFeatureDataset;
                    IEnumDataset enumFeatDataset = featDataset.Subsets;

                    IDataset currentFData = enumFeatDataset.Next();
                    while (currentFData != null)
                    {
                        dataList.Add(currentFData.BrowseName);
                        currentFData = enumFeatDataset.Next();
                    }

                }


                currentData = enumData.Next();
            }

            return dataList;
        }

        /// <summary>
        /// Will return true if a given data type name already exist in project database
        /// </summary>
        /// <param name="dataType"></param>
        /// <param name="inName"></param>
        /// <returns></returns>
        public static bool GetNameExistsFromWorkspace(IWorkspace inputWorkspace, esriDatasetType dataType, string inName)
        {
            //Access workspace
            IWorkspace2 work2 = inputWorkspace as IWorkspace2;

            return work2.get_NameExists(dataType, inName);
        }

        /// <summary>
        /// Will validate a dataset name from a given workspace and remove any problemactic characters or keyword.
        /// </summary>
        /// <param name="inputWorkspace">The workspace in which the new dataset will be added.</param>
        /// <param name="inputName">The wanted name for the output dataset name</param>
        /// <returns></returns>
        public static string GetValidDatasetName(IWorkspace inputWorkspace, string inputName)
        {
            //Variable
            string validName = inputName;

            //Validate
            IFieldChecker validChecker = new FieldChecker() as IFieldChecker;
            validChecker.InputWorkspace = inputWorkspace;
            validChecker.ValidateTableName(inputName, out validName);

            return validName;
        }

        #endregion

        #region ACCESS METHODS

        /// <summary>
        /// Create a workspace factory to access a file database, personal database or shapefile
        /// </summary>
        /// <param name="inputFeaturePath">Reference full path to wanted workspace, usually project main geodatabase</param>
        /// <returns></returns>
        public static IWorkspace AccessWorkspace(string inputPath)
        {
            //Variables
            IWorkspaceFactory workFactory = null;
            IWorkspace workspc = null;

            try
            {
                if (inputPath.Substring(inputPath.Length - 4, 4).ToLower() == ".gdb")
                {
                    //Activate the singleton, or else System.__ComObject errors could pop.
                    Type t = Type.GetTypeFromProgID("esriDataSourcesGDB.FileGDBWorkspaceFactory");
                    System.Object obj = Activator.CreateInstance(t);
                    workFactory = obj as FileGDBWorkspaceFactory;

                }
                else if (inputPath.Substring(inputPath.Length - 4, 4).ToLower() == ".mdb")
                {
                    //Activate the singleton, or else System.__ComObject errors could pop.
                    Type t = Type.GetTypeFromProgID("esriDataSourcesGDB.AccessWorkspaceFactory");
                    System.Object obj = Activator.CreateInstance(t);
                    workFactory = obj as AccessWorkspaceFactory;
                }
                else if (inputPath.Substring(inputPath.Length - 4, 4).ToLower() == ".shp" || inputPath.Substring(inputPath.Length - 4, 4).ToLower() == ".dbf")
                {
                    //Get workspace path and file name from input
                    inputPath = System.IO.Path.GetDirectoryName(inputPath); //Rest path to directory, for shapefiles only.
                    workFactory = new ShapefileWorkspaceFactory();

                }
                workspc = workFactory.OpenFromFile(inputPath, 0);

                //Release workspace factory
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workFactory);

            }
            catch (Exception accessWorkspaceExcept)
            {
                int lineNumber = Exceptions.LineNumber(accessWorkspaceExcept);
                MessageBox.Show("accessWorkspaceExcept (" + lineNumber.ToString() + "):" + accessWorkspaceExcept.Message);
            }

            return workspc;
        }

        /// <summary>
        /// Will return a workspace factory for excel given an input path without the file sheet name 
        /// </summary>
        /// <param name="inputFeaturePath">Reference full path excel sheet</param>
        /// <returns></returns>
        public static IWorkspace AccessExcelWorkspace(string inputPath)
        {
            //Variables
            IWorkspaceFactory workFactory = null;
            IWorkspace workspace = null;

            try
            {
                if (inputPath.Contains(".xls") || inputPath.Contains(".xlsx"))
                {
                    //Activate the singleton, or else System.__ComObject errors could pop.
                    Type t = Type.GetTypeFromProgID("esriDataSourcesOleDB.ExcelWorkspaceFactory");
                    System.Object obj = Activator.CreateInstance(t);
                    workFactory = obj as ExcelWorkspaceFactory;

                }
                workspace = workFactory.OpenFromFile(inputPath, 0);

                //Release workspace factory
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workFactory);

            }
            catch (Exception accessWorkspaceExcept)
            {
                int lineNumber = Exceptions.LineNumber(accessWorkspaceExcept);
                MessageBox.Show("AccessExcelWorkspace (" + lineNumber.ToString() + "):" + accessWorkspaceExcept.Message);
            }

            return workspace;
        }

        /// <summary>
        /// Will return a workspace factory for a given textfile path. 
        /// </summary>
        /// <param name="inputFeaturePath">Reference full path to wanted text file</param>
        /// <returns></returns>
        public static IWorkspace AccessTextfileWorkspace(string inputPath)
        {
            //Variables
            IWorkspaceFactory workFactory = null;
            IWorkspace workspace = null;

            //Get parent folder
            string inputPathParentFolder = System.IO.Path.GetDirectoryName(inputPath);

            try
            {
                if (inputPath.Contains(".txt") || inputPath.Contains(".csv"))
                {
                    //Activate the singleton, or else System.__ComObject errors could pop.
                    Type t = Type.GetTypeFromProgID("esriDataSourcesFile.TextFileWorkspaceFactory");
                    workFactory = (IWorkspaceFactory)Activator.CreateInstance(t);

                }
                workspace = workFactory.OpenFromFile(inputPathParentFolder, 0);

                //Release workspace factory
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workFactory);

            }
            catch (Exception accessWorkspaceExcept)
            {
                int lineNumber = Exceptions.LineNumber(accessWorkspaceExcept);
                MessageBox.Show("AccessTextfileWorkspace (" + lineNumber.ToString() + "):" + accessWorkspaceExcept.Message);
            }

            return workspace;
        }

        /// <summary>
        /// Create a workspace factory to access a file database, personal database or shapefile, used mainly to find if a given feature exists inside database.
        /// </summary>
        /// <param name="inputFeaturePath">Reference full path to wanted workspace, usually project main geodatabase</param>
        /// <returns></returns>
        public static IWorkspace2 AccessWorkspace2(string inputPath)
        {
            //Variables
            IWorkspaceFactory workFactory = null;
            IWorkspace2 workspc = null;

            try
            {
                if (inputPath.Substring(inputPath.Length - 4, 4).ToLower() == ".gdb")
                {
                    //Activate the singleton, or else System.__ComObject errors could pop.
                    Type t = Type.GetTypeFromProgID("esriDataSourcesGDB.FileGDBWorkspaceFactory");
                    System.Object obj = Activator.CreateInstance(t);
                    workFactory = obj as FileGDBWorkspaceFactory;

                }
                else if (inputPath.Substring(inputPath.Length - 4, 4).ToLower() == ".mdb")
                {
                    //Activate the singleton, or else System.__ComObject errors could pop.
                    Type t = Type.GetTypeFromProgID("esriDataSourcesGDB.AccessWorkspaceFactory");
                    System.Object obj = Activator.CreateInstance(t);
                    workFactory = obj as AccessWorkspaceFactory;
                }
                else if (inputPath.Substring(inputPath.Length - 4, 4).ToLower() == ".shp" || inputPath.Substring(inputPath.Length - 4, 4).ToLower() == ".dbf")
                {
                    //Get workspace path and file name from input
                    inputPath = System.IO.Path.GetDirectoryName(inputPath); //Rest path to directory, for shapefiles only.
                    workFactory = new ShapefileWorkspaceFactory();

                }

                workspc = workFactory.OpenFromFile(inputPath, 0) as IWorkspace2;

                //Release workspace factory
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workFactory);

            }
            catch (Exception accessWorkspaceExcept)
            {
                int lineNumber = Exceptions.LineNumber(accessWorkspaceExcept);
                MessageBox.Show("accessWorkspaceExcept (" + lineNumber.ToString() + "):" + accessWorkspaceExcept.Message);
            }

            return workspc;
        }

        /// <summary>
        /// Not Working... was meant to give access to a in_memory workspace.Instead use
        /// </summary>
        /// <param name="getWorkspaceFactory"></param>
        /// <returns></returns>
        public static IWorkspace AccessInMemoryWorkspace()
        {

            //Create an in-memory workspace factory
            IWorkspaceFactory workFact = new InMemoryWorkspaceFactory();

            //Create a new workspace name to init a new workspace
            IWorkspaceName workName = workFact.Create(null, "inMemory", null, 0);
            IName wName = (IName)workName;

            //Open
            IWorkspace workspace = (IWorkspace)wName.Open();

            return workspace;
        }

        #endregion

        #region CREATE METHODS

        /// <summary>
        /// Will create a file geodatabase and return a workspace object.
        /// </summary>
        /// <param name="newPath"> Input new database full path (C:\Folder\Folder\...)</param>
        /// <param name="newName"> Input new database name</param>
        /// <returns></returns>
        public static IWorkspace CreateWorkspace(string parentDirectory, string newName)
        {
            IWorkspace getWorkspace = null;

            //Create a factor
            Type factoryType = Type.GetTypeFromProgID("esriDataSourcesGDB.FileGDBWorkspaceFactory");
            IWorkspaceFactory workFactory = Activator.CreateInstance(factoryType) as IWorkspaceFactory;
            IWorkspaceName workName = workFactory.Create(parentDirectory, newName + ".gdb", null, 0);

            //Cast and return
            IName name = workName as IName;
            getWorkspace = name.Open() as IWorkspace;

            //Release workspace factory
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workFactory);

            return getWorkspace;
        }

        /// <summary>
        /// Will create and return an in_memory workspace to be used in scratchWorkspaces
        /// </summary>
        /// <returns></returns>
        public static IWorkspace CreateInMemoryWorkspace()
        {
            //Create a factory 
            IWorkspaceFactory imWork = new InMemoryWorkspaceFactory();

            //Create a work name
            IWorkspaceName workName = imWork.Create(null, "IMeMineWorkspace", null, 0);
            IName name = workName as IName;

            //Build the workspace
            IWorkspace inMemoryWorkspace = name.Open() as IWorkspace;

            //Release workspace factory
            System.Runtime.InteropServices.Marshal.ReleaseComObject(imWork);

            return inMemoryWorkspace;
        }

        /// <summary>
        /// Will create a new workspace for shapefiles, based on the parent folder name.
        /// </summary>
        /// <param name="folderForShapefile">The parent folder path. New shapefiles can be created inside this folder</param>
        /// <returns></returns>
        public static IWorkspace CreateWorkspaceForShapefiles(string folderForShapefile)
        {
            //string environmentDBPath = Properties.MainRepoSettings.Default.Tools4ProjectFolder + Constants.Environment.envRelPath;
            bool dbExists = System.IO.Directory.Exists(folderForShapefile);

            if (!dbExists)
            {
                //Create the folder
                Directory.CreateDirectory(folderForShapefile);
            }

            //Create a factory 
            IWorkspaceFactory shapeWorkspaceFactory = new ShapefileWorkspaceFactory();

            //Build the workspace
            IWorkspace shapeWorkspace = shapeWorkspaceFactory.OpenFromFile(folderForShapefile, 0);

            //Release workspace factory
            System.Runtime.InteropServices.Marshal.ReleaseComObject(shapeWorkspaceFactory);

            return shapeWorkspace;

        }

        #endregion

        #region SCHEMA LOCKS

        public static void ListSchemaLocksForObjectClass(IObjectClass objectClass)
        {
            //Get an exclusive schema lock on the dataset.
            ISchemaLock schemaLock = (ISchemaLock)objectClass;

            // Get an enumerator over the current schema locks.
            IEnumSchemaLockInfo enumSchemaLockInfo = null;
            schemaLock.GetCurrentSchemaLocks(out enumSchemaLockInfo);

            // Iterate through the locks.
            ISchemaLockInfo schemaLockInfo = null;
            while ((schemaLockInfo = enumSchemaLockInfo.Next()) != null)
            {
                MessageBox.Show(schemaLockInfo.TableName + "; " + schemaLockInfo.UserName + "; " + schemaLockInfo.SchemaLockType);

            }
        }
        #endregion

    }
}
