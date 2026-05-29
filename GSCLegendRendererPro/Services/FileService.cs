using GSCLegendRendererPro.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GSCLegendRendererPro.Services
{
    public  class FileService
    {
        /// <summary>
        /// Will save an embedded resource in the working environment folder
        /// </summary>
        /// <param name="resourceBytes"></param>
        /// <param name="outputPath"></param>
        public static void WriteStreamResource(byte[] resourceBytes, string outputPath)
        {
            try
            {
                Stream outputStream = new FileStream(outputPath, FileMode.Create, FileAccess.Write, FileShare.None);
                using (BinaryWriter fileWriter = new BinaryWriter(outputStream))
                {
                    fileWriter.Write(resourceBytes);
                    fileWriter.Close();
                }
            }
            catch (Exception e)
            {
                new ErrorService(e).WriteToFile();
            }

        }
    }
}
