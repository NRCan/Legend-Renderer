using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Geoprocessing;
using ESRI.ArcGIS.GeoDatabaseUI;
using ESRI.ArcGIS.esriSystem;

namespace GSC_Legend_Renderer.Services
{
    public class FolderAndFiles
    {
        //Global variables
        public static int ColumnIterationsExcelSheet = 1;
        public static string parallelLoopSymbol = string.Empty;

        /// <summary>
        /// Will return a list of file path, from a given folder.
        /// </summary>
        /// <param name="folderPath">Folder path to list files from</param>
        /// <param name="fileExtension">A wildCard to filter the seach (ex: "*.shp")</param>
        /// <returns></returns>
        public static List<string> GetListOfFilesFromFolder(string folderPath, string wildCard)
        {
            //https://docs.microsoft.com/en-us/dotnet/csharp/programming-guide/file-system/how-to-iterate-through-a-directory-tree

            //Variables
            List<string> outputListOfFiles = new List<string>();

            Stack<string> dirs = new Stack<string>(200);

            if (!System.IO.Directory.Exists(folderPath))
            {
                throw new ArgumentException();
            }
            dirs.Push(folderPath);

            while (dirs.Count > 0)
            {
                string currentDir = dirs.Pop();
                string[] subDirs;
                try
                {
                    subDirs = System.IO.Directory.GetDirectories(currentDir);
                }
                // An UnauthorizedAccessException exception will be thrown if we do not have
                // discovery permission on a folder or file. It may or may not be acceptable 
                // to ignore the exception and continue enumerating the remaining files and 
                // folders. It is also possible (but unlikely) that a DirectoryNotFound exception 
                // will be raised. This will happen if currentDir has been deleted by
                // another application or thread after our call to Directory.Exists. The 
                // choice of which exceptions to catch depends entirely on the specific task 
                // you are intending to perform and also on how much you know with certainty 
                // about the systems on which this code will run.
                catch (UnauthorizedAccessException e)
                {
                    MessageBox.Show(e.Message, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    continue;
                }
                catch (System.IO.DirectoryNotFoundException e)
                {
                    continue;
                }
                catch (Exception getlistOffilesFromFolderException)
                {
                    int lineNumber = Services.Exceptions.LineNumber(getlistOffilesFromFolderException);
                    MessageBox.Show("getListofFilesFromFolder (" + lineNumber.ToString() + "):" + getlistOffilesFromFolderException.Message);
                    continue;
                }


                string[] files = null;
                try
                {
                    files = System.IO.Directory.GetFiles(currentDir,wildCard);
                }

                catch (UnauthorizedAccessException e)
                {
                    continue;
                }

                catch (System.IO.DirectoryNotFoundException e)
                {
                    continue;
                }

                // Perform the required action on each file here.
                // Modify this block to perform your required task.
                foreach (string file in files)
                {
                    try
                    {
                        outputListOfFiles.Add(file);
                    }
                    catch (System.IO.FileNotFoundException e)
                    {
                        continue;
                    }
                }

                // Push the subdirectories onto the stack for traversal.
                // This could also be done before handing the files.
                foreach (string str in subDirs)
                    dirs.Push(str);
            }

            return outputListOfFiles;

        }


        /// <summary>
        /// Will append some text into an existing textfile
        /// </summary>
        public static void WriteToTextFile(string filePath, string textToAppend)
        {
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(filePath, true))
            {
                file.WriteLine(textToAppend);
            }
            
        }

        /// <summary>
        /// Will return a string array coming from a textfile for all the lines inside it.
        /// </summary>
        /// <returns></returns>
        public static string[] ReadTextFile(string filePath)
        {
            return System.IO.File.ReadAllLines(filePath);
        }

        /// <summary>
        /// Will write a resource from a folder inside a given namepsace into an output folder.
        /// </summary>
        /// <param name="inFile">The file to extract and copy to, needs to reside inside project</param>
        /// <param name="inFolder">The folder inside the project containing the file to copy</param>
        /// <param name="inNameSpace">The namespace of the project that has the file and folder</param>
        /// <param name="outFolder">The output folder that will contain a copy of the file</param>
        public static void WriteResourceToFile(string inFile, string inFolder, string inNameSpace, string outFolder)
        {
            //Get the assembly object to access files
            Assembly ass = Assembly.GetCallingAssembly();  

            //Make sure output folder exists, else create it
            System.IO.Directory.CreateDirectory(outFolder);

            //Access the file with a stream
            using (Stream stream = ass.GetManifestResourceStream(inNameSpace + "." + (inFolder == "" ? "" : inFolder + ".") + inFile))
            {
                //Read the file with a reader
                using(BinaryReader binReader = new BinaryReader(stream))
                {
                    //Init the stream to write the file at the right place
                    using (System.IO.FileStream fileStream = new System.IO.FileStream(outFolder + "\\" + inFile, FileMode.OpenOrCreate))
                    {
                        //Write the file
                        using(BinaryWriter binWriter = new BinaryWriter(fileStream))
                        {
                            binWriter.Write(binReader.ReadBytes((int)stream.Length));
                            binWriter.Close();
                        }
                    }
                }
            }

        }

        /// <summary>
        /// Will return a list of folder path, from a given folder.
        /// </summary>
        /// <param name="folderPath">The folder to search in</param>
        /// <param name="wildCard">A wildcard to refine the search</param>
        /// <returns></returns>
        public static List<string> GetListOfFoldersFromFolder(string folderPath, string wildCard)
        {
            //Variables
            List<string> outputListOfFolders = new List<string>();

            try
            {
                foreach (string files in Directory.GetDirectories(folderPath, wildCard, SearchOption.AllDirectories))
                {
                    outputListOfFolders.Add(files);
                }


            }
            catch (Exception getlistOffoldersFromFolderException)
            {
                int lineNumber = Services.Exceptions.LineNumber(getlistOffoldersFromFolderException);
                MessageBox.Show("getlistOffoldersFromFolderException (" + lineNumber.ToString() + "):" + getlistOffoldersFromFolderException.Message);
            }

            return outputListOfFolders;
        }

        /// <summary>
        /// Will rename a given folder path to an output folder path
        /// </summary>
        /// <param name="fromFolder">From folder path to rename</param>
        /// <param name="toFolder">To folder path to replace old</param>
        public static void RenameFolder(string fromFolder, string toFolder)
        {
            Directory.Move(fromFolder, toFolder);
        }

    }
}
