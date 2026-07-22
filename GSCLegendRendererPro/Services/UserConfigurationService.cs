using ArcGIS.Desktop.Mapping;
using GSCLegendRendererPro.Models;
using GSCLegendRendererPro.Utilities;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading.Tasks;

namespace GSCLegendRendererPro.Services
{
    public class UserConfigurationService: UserConfiguration
    {
        public UserConfigurationService() { }

        /// <summary>
        /// Will validate json configuration files and style file, if it doesn't exist a copy will be made inside My Document\Arc GIS folder
        /// </summary>
        /// <returns>Will return output folder path</returns>
        public static async Task ValidateAssetsExistance()
        {
            UserConfigurationSetup _userConfigurationSetup = new UserConfigurationSetup();

            //Make sure the directory exists
            if (!Directory.Exists(_userConfigurationSetup.DefaultPath))
            {
                Directory.CreateDirectory(_userConfigurationSetup.DefaultPath);
            }

            string jsonYSpacingFilePath = System.IO.Path.Combine(_userConfigurationSetup.DefaultPath, Constants.Assets.jsonYSpacingEmbeddedFile);
            byte[] jsonYBytes = Properties.Resources.Configuration_Y_Spacings;
            if (!System.IO.File.Exists(jsonYSpacingFilePath))
            {
                Services.FileService.WriteStreamResource(jsonYBytes, jsonYSpacingFilePath);
            }

            string jsonXSpacingFilePath = System.IO.Path.Combine(_userConfigurationSetup.DefaultPath, Constants.Assets.jsonXSpacingEmbeddedFile);
            byte[] jsonXBytes = Properties.Resources.Configuration_X_Spacings;
            if (!System.IO.File.Exists(jsonXSpacingFilePath))
            {
                Services.FileService.WriteStreamResource(jsonXBytes, jsonXSpacingFilePath);
            }

            string jsonOtherFilePath = System.IO.Path.Combine(_userConfigurationSetup.DefaultPath, Constants.Assets.jsonStyleFontsOtherEmbeddedFile);
            byte[] jsonOtherBytes = Properties.Resources.Configuration_Other;
            if (!System.IO.File.Exists(jsonOtherFilePath))
            {
                Services.FileService.WriteStreamResource(jsonOtherBytes, jsonOtherFilePath);
            }

            string styleFilePath = System.IO.Path.Combine(_userConfigurationSetup.DefaultPath, Constants.Assets.gscSymbolStandardStyle);
            byte[] styleBytes = Properties.Resources.GSC_SymbolStandard;
            if (!System.IO.File.Exists(styleFilePath))
            {
                Services.FileService.WriteStreamResource(styleBytes, styleFilePath);
            }

            string templateFilePath = System.IO.Path.Combine(_userConfigurationSetup.DefaultPath, Constants.Assets.pageLayoutTemplateFile);
            byte[] templateBytes = Properties.Resources.LegendRendererTemplate;
            if (!System.IO.File.Exists(templateFilePath))
            {
                Services.FileService.WriteStreamResource(templateBytes, templateFilePath);
            }

            string demPictureFilePath = System.IO.Path.Combine(_userConfigurationSetup.DefaultPath, Constants.Assets.demPicture);
            byte[] demBytes = Properties.Resources.UNIT_BOX_DEM;
            if (!System.IO.File.Exists(demPictureFilePath))
            {
                Services.FileService.WriteStreamResource(demBytes, demPictureFilePath);
            }

        }

        /// <summary>
        /// Will deserialize a json file and return the object
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        /// <exception cref="FileNotFoundException"></exception>
        public static async Task<T> DeserializeJsonFile<T>(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException(string.Format(Properties.Resources.ErrorFileNotFound, filePath));
            }

            using (FileStream openStream = File.OpenRead(filePath))
            {
                // Enable support 
                JsonSerializerOptions options = new JsonSerializerOptions { IncludeFields = true, NumberHandling = JsonNumberHandling.WriteAsString | JsonNumberHandling.AllowReadingFromString};
                T deserializedObject = await JsonSerializer.DeserializeAsync<T>(openStream, options);
                openStream.Close();
                openStream.Dispose();
                return deserializedObject;
            }

        }
    }
}
