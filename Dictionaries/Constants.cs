
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GSC_Legend_Renderer.Dictionaries
{
    public class Constants
    {
        public static class ESRI
        {
            //http://resources.arcgis.com/en/help/arcobjects-net/conceptualhelp/index.html#//00010000029s000000

            public const string UIDLayoutViewCommand = "esriArcMapUI.LayoutViewCommand";
            public const string UIDLayoutAlignVerticalCenterCommand = "esriArcMapUI.AlignMiddleCommand";
            public const string defaultArcGISFolderName = "ArcGIS";
        }

        public static class Extensions
        {
            //Extensions
            public const string shpExt = ".shp";
            public const string gdbExt = ".gdb";
            public const string mdbExt = ".mdb";
            public const string xlExt = ".xls";
            public const string txtExt = ".txt";
            public const string csvExt = ".csv";
            public const string dbfExt = ".dbf";
        }

        /// <summary>
        /// Mandatory field names found in the legend table
        /// </summary>
        public static class LegendTable
        {
            public const string legendOrderField = "ELEMENT_ORDER";
            public const string legendColumnField = "ELEMENT_COLUMN";
            public const string legendElementField = "ELEMENT";
            public const string legendStyle1Field = "STYLE1";
            public const string legendStyle2Field = "STYLE2";
            public const string legendLabel1Field = "LABEL1";
            public const string legendLabel1StyleField = "LABEL1STYLE";
            public const string legendLabel2Field = "LABEL2";
            public const string legendLabel2StyleField = "LABEL2STYLE";
            public const string legendHeadingField = "HEADING";
            public const string legendDescriptionField = "DESCRIPTION";
        }

        /// <summary>
        /// Graphic names found in template.mxd and values for field ELEMENT
        /// </summary>
        public static class Graphics
        {
            public const string unitBox = "UNIT_BOX";

            public const string unitSplit = "UNIT_SPLIT";
            public const string subUnitSplit1 = "UNIT_SPLIT1";
            public const string subUnitSplit2 = "UNIT_SPLIT2";

            public const string unitLine = "UNIT_LINE";
            public const string subUnitLine = "LINEUNIT_LINE";

            public const string unitParent = "UNIT_PARENT";
            public const string subUnitParentChild = "UNIT_CHILD";
            public const string subUnitParentChildLine = "UNIT_CHILD_LINE";

            public const string unitindent1 = "UNIT_INDENT";
            public const string unitindent2 = "UNIT_INDENT2";

            public const string lineDouble = "TWOSIDE";
            public const string lineDoubleFlip = "TWOSIDE_FLIP";
            public const string subLineDoubleTop = "TWOSIDE_TOP";
            public const string subLineDoubleBottom = "TWOSIDE_BOTTOM";

            public const string lineDoubleFLow = "TWOSIDE_FLOW";
            public const string subLineDoubleFLowTop = "TWOSIDE_TOP";
            public const string subLineDoubleFLowMiddle = "TWOSIDE_FLOW_MIDDLE";
            public const string subLineDoubleFLowBottom= "TWOSIDE_BOTTOM";

            public const string bracketLeftLower = "L_BRACKET_L";
            public const string bracketLeftCenter = "L_BRACKET_C";
            public const string bracketLeftUpper = "L_BRACKET_U";
            public const string bracketRightLower = "R_BRACKET_L";
            public const string bracketRightCenter = "R_BRACKET_C";
            public const string bracketRightUpper = "R_BRACKET_U";
            public const string bracketSpine = "BRACKET_SPINE";
            //public const string bracketLeftStart = "L_BRACKET_START";
            //public const string bracketLeftEnd = "L_BRACKET_END";
            //public const string bracketRightStart = "R_BRACKET_START";
            //public const string bracketRightEnd = "R_BRACKET_END";

            public const string overlay = "OVERLAY";
            public const string breakLine = "BREAK";
            public const string point = "POINT_CC";
            public const string pointAngle = "POINT_CC_45";
            public const string pointAngleLine = "POINT_LC_45";
            public const string wave = "WAVE";
            public const string line = "LINE";
            public const string blob = "BLOB";
            public const string landslide = "LANDSLIDE";
            public const string beach = "BEACH";
            public const string dunes = "DUNES";
            public const string moraines = "MORAINES";
            public const string annotationBreak = "ANNO_BREAK";
            public const string legendBoxDEM = "LegendBoxDEM";

            public const string topNote = "TOP_NOTE";
            public const string note = "NOTE";

            public const string unitLabel = "UNIT_LABEL"; //Added by GHV
            public const string heading1 = "HEADING1"; //Added by GHV
            public const string heading2 = "HEADING2"; //Added by GHV
            public const string heading3 = "HEADING3"; //Added by GHV
            public const string heading4 = "HEADING4"; //Added by GHV
            public const string heading5 = "HEADING5"; //Added by GHV
            public const string heading5_end = "HEADING5_END"; //Added by GHV
            public const string annoBracket = "ANNO_BRACKET"; //Added by GHV
            public const string description = "DESCRIPTION"; //Added by GHV
            public const string description_indent = "DESCRIPTION_INDENT";//Added by GHV
            public const string description_indent2 = "DESCRIPTION_INDENT2";//Added by GHV
            public const string heading5Description = "GROUP_DESCRIPTION";

            //Other element names
            public const string columnWidth = "COLUMN_WIDTH";
            public const string elementWidth = "ELEMENT_WIDTH";
            public const string elementDescriptionGapWidth = "ELEMENT_DESCRIPTION_GAP_WIDTH"; //Added by GHV
            public const string descriptionWidth = "DESCRIPTION_WIDTH"; 
            public const string columnColumnGapWidth = "COLUMN_COLUMN_GAP_WIDTH"; //Added by GHV
            public const string bracketLeftGap = "L_BRACKET_GAP"; //Added by GHV
            public const string bracketRightGap = "R_BRACKET_GAP"; //Added by GHV
            public const string unitBoxBracket = "UNIT_BOX_R_BRACKET_GAP"; //Added by GHV
            public const string measurementLabel = "MEASUREMENT_LABEL"; //Added by GHV
            public const string generationLabel = "GENERATION_LABEL"; //Added by GHV
            public const string defaultpointLabel = "DEFAULT_POINT_LABEL"; //Added by GHV
            public const string bracketLeftColumnGap = "L_BRACKET_COLUMN_GAP"; //Added by GHV
            public const string annotationBlob = "ANNO_BLOB"; //Added by GHV
            public const string groupDescriptionWidth = "GROUP_DESCRIPTION_WIDTH"; //Added by GHV

            public const string anchorCenterCenter = "CC";
            public const string anchorUupperLeft = "UL";
            public const string anchorCenterLeft = "CL";
            public const string anchorLowerCenter = "LC";
            public const string anchorUpperCenter = "UC";
            public const string anchorLowerLeft = "LL";
            public const string anchorCenterRight = "CR";
            public const string anchorLowerRight = "LR";
            public const string anchorUpperRight = "UR";

            public const string defaultStartEmplacement = "";

            public const string keywordBracket = "BRACKET";
            public const string keywordEnd = "END"; //Will be used for column auto-calculate

            public enum UnitBoxType { normal, split1, split2, parent, child, line, child_line}; //Will be used to correctly place label inside unit boxes

            //CGM template elements
            public const string cgmLegendElement = "Legend";
            public const string cgmCitation = "Citation";
            public const string cgmDetectorKeyword = "CGM";

        }

        /// <summary>
        /// Asset file names.
        /// </summary>
        public static class Assets
        {
            public const string AssetFolder = "Assets";
            public const string mxdEmbeddedFile = "LegendRendererTemplate.mxd";
            public const string jsonYSpacingEmbeddedFile = "Configuration_Y_Spacings.json";
            public const string jsonXSpacingEmbeddedFile = "Configuration_X_Spacings.json";
            public const string jsonStyleFontsOtherEmbeddedFile = "Configuration_Other.json";
            public const string demPicture = "LegendBoxDEM.png";
        }

        public static class Namespaces
        {
            public const string mainNamespace = "GSC_Legend_Renderer";
        }

        public static class Styles
        {
            //public const string styleName = "GSC_SymbolStandard"; //Moved to JSON file
            public const string styleNameJSON = "GEOLOGY_STYLE_NAME";
            public const string styleExtension = ".style";
            public const string styleFillClass = "Fill Symbols";
            public const string styleMarkerClass = "Marker Symbols";
            public const string styleLineClass = "Line Symbols";
            //public const string styleRepresentationMarkerClass = "Representation Markers";
            public const string styleTextClass = "Text Symbols";

            /// <summary>
            /// A list of possible position of label surrounding marker points.
            /// </summary>
            public enum MarkerLabelPositioning { FromCenterToUpperLeft, FromCenterToUpperRight, RightAboveCenter, FromCenterToUpperRightTight}
        }

        /// <summary>
        /// Some hardcoded values for Y spacings
        /// </summary>
        public static class YSpacings
        {
            public const double smallDescriptionHeightLimit = 10.0; //Beyond this, it is a long description
            public const double smallDescriptionHeightLimitLines = 6.56; //Beyond this, it is a long description for a line symbol.
            public const double legendEnd_Citation = 15.0; //This include a bufer of 10 for a better fit. Used to calculate number of columns
            public const double lineDescriptionHeightAdjustement = 0.5; //Aligning lines with long description with their top introduces a small gap because text as some slight margin at the top
            public const double lineHeight0DescriptionHeightAdjustement = 1.5; //Flat lines symbols have a height of 0, take this instead when a height is needed.
            public const double markerMeanHeight = 2.5; //Used to position subsequent element after a long marker description
        }

        public static class TextConfiguration
        {
            public const int charactersPerLine = 76;
            public const double lineHeight = 3.28; //mm, used to calculate text box height approx.
            public const double header3LineHeight = 3.66; //mm, used to calculate header 3 text box height approx.
            //HTML related tags
            public const string tagBold = "<bol>";
            public const string endTagBold = "</bol>";
            public const string tagAllCaps = "<ACP>";
            public const string endTagAllCaps = "</ACP>";
            public const string tagItalic = "<ITA>";
            public const string endTagItalic = "</ITA>";
            public const string tagFont = "<FNT name = ";
            public const string endTagFont = "</FNT>";
            //Missing terms
            public const string missingText = "Missing";
            public const string NullLiteral = "<Null>";
            //Defaults
            public const double defaultUnitBoxLabelFontSize = 8.0; //Mainly used for labels that see their font size change when using a new style.
            public const double tooLongLabelUnitBoxLabelFontSize = 7.5; //Mainly used for labels that see their font size change when using a new style.
        }

        public static class ImageConfiguration
        {
            //public const int demTransparency = 178; // 178/255 --> 70% opaque which is 30% transparent.
            public const string demTransparencyNameJSON = "DEM_OPACITY_PERCENT";
            public const string monoColoredImageNamePrefix = "Mono_";
        }

        public static class GraphicConfiguration
        {
            public const double outlineWidth = 0.43;
        }

        public static class Fonts
        {
            public const string geologytFontNameJSON = "GEOLOGY_FONT_NAME";
            public const double geologyFontHeightAjustement = 0.4; //Add 0.4 mm so it widens the box.
        }

        public static class ObjectNames
        {
            public const string fillTypeSimple = "ISimpleFillSymbol";
            public const string fillTypeGradient = "IGradientFillSymbol";
            public const string fillTypeLine = "ILineFillSymbol";
            public const string fillTypeMarker = "IMarkerFillSymbol";
            public const string fillTypePicture = "IPictureFillSymbol";
            public const string fillTypeMultilayer = "IMultiLayerFillSymbol";

        }
    }
}
