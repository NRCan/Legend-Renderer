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
    /// C:\Users\<username>\Documents\ArcGIS\GSC Legend Renderer <version #>\configuration_other.json
    /// </summary>  
    public class UserConfiguration
    {

        public string GEOLOGY_FONT_NAME { get; set; } 
        public string GEOLOGY_STYLE_NAME { get; set; }
        public int DEM_OPACITY_PERCENT { get; set; } 
    }

    public class UserConfigurationSetup
    {

        public string DefaultPath
        {
            get
            {
                return WorkingEnvironment.WorkingEnvironmentPath;
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
