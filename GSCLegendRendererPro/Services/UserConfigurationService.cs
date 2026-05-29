using ArcGIS.Desktop.Mapping;
using GSCLegendRendererPro.Models;
using GSCLegendRendererPro.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace BedrockEditorPro.Services
{
    public class UserConfigurationService: UserConfiguration
    {
        public UserConfigurationService() { }

        public static async Task<UserConfiguration> GetUserConfigurationAsync()
        {

            UserConfigurationSetup _userConfigurationSetup = new UserConfigurationSetup();

            //Make sure the directory exists
            if (!Directory.Exists(Path.GetDirectoryName(_userConfigurationSetup.DefaultPath)))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_userConfigurationSetup.DefaultPath));
            }

            //If the file exists, read it
            if (File.Exists(_userConfigurationSetup.DefaultPath))
            {
                
                using (FileStream openStream = File.OpenRead(_userConfigurationSetup.DefaultPath))
                {
                    // Enable support
                    var options = new JsonSerializerOptions { IncludeFields = true };

                    _userConfigurationSetup.UserConfiguration = await JsonSerializer.DeserializeAsync<UserConfiguration>(openStream, options);

                    openStream.Close();
                    openStream.Dispose();
                }

            }
            else
            {
                //else write it

                //Manage default symbol style file
                string stylePath = Symbols.ManageStyleFile();
                _userConfigurationSetup.UserConfiguration.StyleFilePath = stylePath;

                await using FileStream fStream = File.Create(_userConfigurationSetup.DefaultPath);
                await JsonSerializer.SerializeAsync(fStream, _userConfigurationSetup.UserConfiguration);
                fStream.Close();

            }

            return _userConfigurationSetup.UserConfiguration;
        }
    }
}
