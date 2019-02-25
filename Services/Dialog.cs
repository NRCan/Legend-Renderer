using System.Windows.Forms;
using ESRI.ArcGIS.CatalogUI;
using ESRI.ArcGIS.Catalog;
using ESRI.ArcGIS.Geometry;

namespace GSC_Legend_Renderer.Services
{
    public class Dialog
    {

        ///A special XML File prompt exist, within the Arc Catalog addin only.
        ///Procedure to create custom file filter to user in file prompts:
        //1) Right-click the project and add new item
        //2) Select a ArcGIS --> Extending Arc GIS objects --> Class
        //3) Select as com type Arc Catalog
        //4) Select as type ObjectFilter
        //5) Check the option to create the gxObjectfilter and click finish.
        //WARNING: This operation might trigger a "Cannot register Assembly Error", you need to uncheck the "Register COM component" option in the project build option.

        ///There is a working example withiin Addin_ArcCatalog for pdf, doc and docx files, see CustomGxFilter.cs

        /// <summary>
        /// Will prompt a simple folder dialog
        /// </summary>
        /// <param name="folderDescription">Enter a string as for description that will be showned within the dialog.</param>
        /// <returns></returns>
        public static string GetFolderPrompt(string folderDescription)
        {

            //Force user to browser for a new folder
            FolderBrowserDialog folderDialog = new FolderBrowserDialog();
            folderDialog.Description = folderDescription;
            
            DialogResult result = folderDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                return folderDialog.SelectedPath;
            }
            else
            {
                //Return old value
                return "";
            }


        }

        /// <summary>
        /// Will prompt a simple file dialog
        /// </summary>
        /// <param name="folderDescription">Enter a string as for description that will be showned within the dialog.</param>
        /// <returns></returns>
        public static string GetFilePrompt(string fileDescription)
        {

            //Force user to browser for a new folder
            OpenFileDialog fileDialog = new OpenFileDialog();
            
            DialogResult result = fileDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                return fileDialog.FileName;
            }
            else
            {
                //Return old value
                return "";
            }


        }

        /// <summary>
        /// Will prompt an ESRI folder prompt
        /// </summary>
        /// <param name="hwnd"></param>
        /// <param name="promptTitle"></param>
        /// <returns></returns>
        public static string GetESRIFolderPrompt(int hwnd, string promptTitle)
        {
            //Create a new dialog instance
            IGxDialog gxDialog = new GxDialog();

            //Create a new object filter for all kind of features
            gxDialog.ObjectFilter = new GxFilterFileFolder();
            gxDialog.AllowMultiSelect = false;
            gxDialog.Title = promptTitle;
            IEnumGxObject folders;

            //Variable
            string userSelectedFolderPath = "";

            //Open dialog and retrieve user's answer
            if (gxDialog.DoModalOpen(hwnd, out folders) && folders != null)
            {
                IGxObject currentObj = folders.Next();
                userSelectedFolderPath = currentObj.FullName;
            }

            //Close
            gxDialog.InternalCatalog.Close();

            //Release dialog
            System.Runtime.InteropServices.Marshal.ReleaseComObject(gxDialog);


            return userSelectedFolderPath;
        }

        ///// <summary>
        ///// Will prompt an ESRI file prompt
        ///// </summary>
        ///// <param name="hwnd"></param>
        ///// <param name="promptTitle"></param>
        ///// <returns></returns>
        //public static string GetESRIFilePrompt(int hwnd, string promptTitle)
        //{
        //    //Create a new dialog instance
        //    IGxDialog gxDialog = new GxDialog();
        //    //Create a new object filter for all kind of features
        //    gxDialog.ObjectFilter = new GxFilterFilesClass();
        //    gxDialog.AllowMultiSelect = false;
        //    gxDialog.Title = promptTitle;
        //    IEnumGxObject files;

        //    //Variable
        //    string userSelectedFilePath = "";

        //    //Open dialog and retrieve user's answer
        //    if (gxDialog.DoModalOpen(hwnd, out files) && files != null)
        //    {
        //        IGxObject currentObj = files.Next();
        //        userSelectedFilePath = currentObj.FullName;
        //    }

        //    //Close
        //    gxDialog.InternalCatalog.Close();

        //    //Release dialog
        //    System.Runtime.InteropServices.Marshal.ReleaseComObject(gxDialog);


        //    return userSelectedFilePath;
        //}


        /// <summary>
        /// Will prompt a data dialog
        /// </summary>
        /// <param name="hwnd">Enter current handle window number (this.Handle.ToInt32 usually makes it)</param>
        /// <returns></returns>
        public static string GetDataPrompt(int hwnd, string promptTitle)
        {
            //Create a new dialog instance
            IGxDialog gxDialog = new GxDialog();
            //Create a new object filter for all kind of features
            gxDialog.ObjectFilter = new GxFilterDatasets();
            gxDialog.AllowMultiSelect = false;
            gxDialog.Title = promptTitle;
            IEnumGxObject datas;

            //Variable
            string userSelectedFeaturePath = "";

            //Open dialog and retrieve user's answer
            if (gxDialog.DoModalOpen(hwnd, out datas) && datas != null)
            {
                IGxObject currentObj = datas.Next();
                userSelectedFeaturePath = currentObj.FullName;
            }

            //Close
            gxDialog.InternalCatalog.Close();

            //Release dialog
            System.Runtime.InteropServices.Marshal.ReleaseComObject(datas);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(gxDialog);


            return userSelectedFeaturePath;
        }

        /// <summary>
        /// Will prompt a dialog for user to select and enter a new folder path to create
        /// </summary>
        /// <param name="hwnd">Enter current handle window number (this.Handle.ToInt32 usually makes it)</param>
        /// <param name="folderDescription">Enter a description for prompt interface</param>
        /// <returns>The returned string is the full path containing the new folder name</returns>
        public static string GetFolderSavePrompt(int hwnd, string folderDescription)
        {
            //Create a new dialog instance
            IGxDialog gxDialog = new GxDialog();

            //Create a new object filter for all kind of features
            gxDialog.ObjectFilter = new GxFilterBasicTypes();
            gxDialog.AllowMultiSelect = false;
            gxDialog.Title = folderDescription;

            //Variable
            string userSelectedFeaturePath = "";

            //Open dialog and retrieve user's answer
            if (gxDialog.DoModalSave(hwnd))
            {
                userSelectedFeaturePath = System.IO.Path.Combine(gxDialog.FinalLocation.FullName, gxDialog.Name);
            }

            //Close
            gxDialog.InternalCatalog.Close();

            //Release dialog
            System.Runtime.InteropServices.Marshal.ReleaseComObject(gxDialog);


            return userSelectedFeaturePath;
        }

        /// <summary>
        /// Will prompt a dialog for user to select an existing file geodatabase
        /// </summary>
        /// <param name="hwnd">Enter current handle window number (this.Handle.ToInt32 usually makes it)</param>
        /// <param name="promptTitle">Enter a prompt title</param>
        /// <returns></returns>
        public static string GetFGDBPrompt(int hwnd, string promptTitle)
        {
            //Create a new dialog instance
            IGxDialog gxDialog = new GxDialog();

            //Create a new object filter for all kind of features
            gxDialog.ObjectFilter = new GxFilterFileGeodatabases();
            gxDialog.AllowMultiSelect = false;
            gxDialog.Title = promptTitle;
            IEnumGxObject features;

            //Variable
            string userSelectedFeaturePath = "";

            //Open dialog and retrieve user's answer
            if (gxDialog.DoModalOpen(hwnd, out features) && features != null)
            {
                IGxObject currentObj = features.Next();
                userSelectedFeaturePath = currentObj.FullName;
            }

            //Close
            gxDialog.InternalCatalog.Close();

            //Release dialog
            System.Runtime.InteropServices.Marshal.ReleaseComObject(gxDialog);


            return userSelectedFeaturePath;

        }


        /// <summary>
        /// Will prompt a workspace (.gdb, .mdb, folders) dialog, for user to select an existing workspace file.
        /// </summary>
        /// <param name="hwnd">Enter current handle window number (this.Handle.ToInt32, usually makes it)</param>
        /// <param name="promptTitle"></param>
        public static string GetWorkspacePrompt(int hwnd, string promptTitle)
        {
            //Create a new dialog instance
            IGxDialog gxDialog = new GxDialog();

            //Create a new object filter for all kind of features
            gxDialog.ObjectFilter = new GxFilterWorkspaces();
            gxDialog.AllowMultiSelect = false;
            gxDialog.Title = promptTitle;
            IEnumGxObject workspaces;

            //Variable
            string userSelectedWorkPath = "";

            //Open dialog and retrieve user's answer
            if (gxDialog.DoModalOpen(hwnd, out workspaces) && workspaces != null)
            {
                IGxObject currentObj = workspaces.Next();
                userSelectedWorkPath = currentObj.FullName;
            }

            //Close
            gxDialog.InternalCatalog.Close();

            //Release dialog
            System.Runtime.InteropServices.Marshal.ReleaseComObject(gxDialog);


            return userSelectedWorkPath;
        }

    }
}
