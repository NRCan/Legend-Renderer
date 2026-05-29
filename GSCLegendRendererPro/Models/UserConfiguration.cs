using GSCLegendRendererPro.Utilities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GSCLegendRendererPro.Models
{
    /// <summary>
    /// A class that will be used to store user configuration settings
    /// It will be saved in the local app data folder as a json file
    /// C:\Users\<username>\Documents\ArcGIS\GSC Bedrock Project <version #>\userconfig.json
    /// </summary>  
    public class UserConfiguration
    {

        public string StyleFilePath { get; set; } = string.Empty;

    }

    public class UserConfigurationSetup
    {

        public string DefaultPath
        {
            get
            {

                var folder = WorkingEnvironment.WorkingEnvironmentPath;

                return Path.Combine(folder, Constants.Configuration.userConfigFileName);
            }

        }

        private WorkingEnvironment WorkingEnvironment
        {
            get
            {
                return new WorkingEnvironment();
            }
        }

        public UserConfiguration UserConfiguration { get; set; }

        public UserConfigurationSetup()
        {
            UserConfiguration = new UserConfiguration();
        }
    }
}
