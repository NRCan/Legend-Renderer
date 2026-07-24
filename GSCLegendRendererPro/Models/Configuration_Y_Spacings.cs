using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GSCLegendRendererPro.Models
{
    /// <summary>
    /// Base class that will hold all Y spacing values as string, because the position could be added to the float value 
    /// like "5.0 from LL".
    /// In millimeters
    /// </summary>
    public class ConfigurationBase
    {
        public string HEADING1 { get; set; }
        public string HEADING2 { get; set; }
        public string HEADING3 { get; set; }
        public string HEADING4 { get; set; }
        public string HEADING5 { get; set; }
        public string DESCRIPTION { get; set; }
        public string NOTE { get; set; }
        public string UNIT_BOX { get; set; }
        public string UNIT_SPLIT { get; set; }
        public string UNIT_LINE { get; set; }
        public string UNIT_PARENT { get; set; }
        public string UNIT_CHILD { get; set; }
        public string UNIT_CHILD_LINE { get; set; }
        public string UNIT_INDENT { get; set; }
        public string UNIT_INDENT2 { get; set; }
        public string BREAK { get; set; }
        public string OVERLAY { get; set; }
        public string WAVE { get; set; }
        public string LINE { get; set; }
        public string TWOSIDE { get; set; }
        public string TWOSIDE_FLIP { get; set; }
        public string TWOSIDE_FLOW { get; set; }
        public string BLOB { get; set; }
        public string BEACH { get; set; }
        public string DUNES { get; set; }
        public string LANDSLIDE { get; set; }
        public string MORAINES { get; set; }
        public string POINT_CC { get; set; }
        public string POINT_CC_45 { get; set; }
        public string POINT_LC_45 { get; set; }
    }

    /// <summary>
    /// A class that will track Y spacing between graphics. 
    /// Example: Heading 1 below another Heading 1, Y spacing is = to 8.0,
    /// Heading 1 below a description, Y spacing is 5.0
    /// In millimeters
    /// </summary>
    public class Configuration_Y_Spacings
    {
        public ConfigurationBase HEADING1 { get; set; }
        public ConfigurationBase HEADING2 { get; set; }
        public ConfigurationBase HEADING3 { get; set; }
        public ConfigurationBase HEADING4 { get; set; }
        public ConfigurationBase HEADING5 { get; set; }
        public ConfigurationBase DESCRIPTION { get; set; }
        public ConfigurationBase NOTE { get; set; }
        public ConfigurationBase UNIT_BOX { get; set; }
        public ConfigurationBase UNIT_SPLIT { get; set; }
        public ConfigurationBase UNIT_LINE { get; set; }
        public ConfigurationBase UNIT_PARENT { get; set; }
        public ConfigurationBase UNIT_CHILD { get; set; }
        public ConfigurationBase UNIT_CHILD_LINE { get; set; }
        public ConfigurationBase UNIT_INDENT { get; set; }
        public ConfigurationBase UNIT_INDENT2 { get; set; }
        public ConfigurationBase BREAK { get; set; }
        public ConfigurationBase OVERLAY { get; set; }
        public ConfigurationBase WAVE { get; set; }
        public ConfigurationBase LINE { get; set; }
        public ConfigurationBase TWOSIDE { get; set; }
        public ConfigurationBase TWOSIDE_FLIP { get; set; }
        public ConfigurationBase TWOSIDE_FLOW { get; set; }
        public ConfigurationBase BLOB { get; set; }
        public ConfigurationBase BEACH { get; set; }
        public ConfigurationBase DUNES { get; set; }
        public ConfigurationBase LANDSLIDE { get; set; }
        public ConfigurationBase MORAINES { get; set; }
        public ConfigurationBase POINT_CC { get; set; }
        public ConfigurationBase POINT_CC_45 { get; set; }
        public ConfigurationBase POINT_LC_45 { get; set; }

        public string GetSpacing(string currentType, string nextType)
        {
            object? row = GetType().GetProperty(currentType)?.GetValue(this);

            return row?
                .GetType()
                .GetProperty(nextType)?
                .GetValue(row)?
                .ToString();
        }
    }

}
