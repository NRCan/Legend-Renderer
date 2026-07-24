using ArcGIS.Core.CIM;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Internal.Mapping;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using GSCLegendRendererPro.Models;
using GSCLegendRendererPro.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xaml;

namespace GSCLegendRendererPro.Utilities
{
    public class Symbols
    {
        /// <summary>
        /// Creates a point symbol renderer
        /// </summary>
        /// <param name="pointRGB">A list containing red green blue numerical codes for point color</param>
        /// <param name="pointSize">A point size</param>
        /// <returns></returns>
        public static CIMPointSymbol GetLabelDefaultRenderer(CIMColor inColor)
        {
            CIMPointSymbol pntSym = SymbolFactory.Instance.ConstructPointSymbol(ColorFactory.Instance.RedRGB, 4, SimpleMarkerStyle.Circle);

            //Create color object to apply to symbol
            CIMColor pntColor = CIMColor.CreateRGBColor(inColor.GetColorComponent(0), inColor.GetColorComponent(1), inColor.GetColorComponent(2));

            ///Note: We used to remove the outline of the circle in version 3.X
            ///it's actually easier to validate when keeping it but still coloring it
            ///just like the map unit

            //Add color and width
            pntSym.SetColor(pntColor);

            return pntSym;
        }

        /// <summary>
        /// Get a grey point symbol for default values or null values
        /// </summary>
        /// <returns></returns>
        public static CIMPointSymbol GetDefaultPointSymbol()
        {
            return SymbolFactory.Instance.ConstructPointSymbol(ColorFactory.Instance.RedRGB, 4, SimpleMarkerStyle.Circle);
        }

        /// <summary>
        /// Get an empty fill polygon symbol for default values or null values
        /// Black outline with an empty filling
        /// </summary>
        /// <returns></returns>
        public static CIMPolygonSymbol GetDefaultPolygonSymbol()
        {
            CIMStroke outline = SymbolFactory.Instance.ConstructStroke(
                ColorFactory.Instance.BlackRGB, 1.0, SimpleLineStyle.Solid);

            CIMPolygonSymbol nullFillWithOutline = SymbolFactory.Instance.ConstructPolygonSymbol(
                ColorFactory.Instance.BlueRGB, SimpleFillStyle.Null , outline);

            return nullFillWithOutline;
        }

        /// <summary>
        /// Will validate the existance of the default style file from
        /// the embedded resource and will return it's path
        /// </summary>
        /// <returns></returns>
        public static string ManageStyleFile()
        {
            WorkingEnvironment workingEnvironment = new WorkingEnvironment();
            string StyleFilePath = System.IO.Path.Combine(workingEnvironment.WorkingEnvironmentPath, nameof(Properties.Resources.GSC_SymbolStandard) + ".stylx");

            if (!File.Exists(StyleFilePath))
            {
                try
                {
                    FileService.WriteStreamResource(Properties.Resources.GSC_SymbolStandard, StyleFilePath);
                }
                catch (Exception e)
                {
                    new ErrorService(e).WriteToFile();
                }
                
            }
            
            return StyleFilePath;

        }

        /// <summary>
        /// Will return a default arcGIS color ramp for layer symbolization fall back
        /// </summary>
        /// <returns></returns>
        public static CIMColorRamp GetDefaultColorRamp(Project inProject)
        {
            StyleProjectItem style = inProject.GetItems<StyleProjectItem>().FirstOrDefault(x => x.Name == "ArcGIS Colors");
            if (style != null)
            {
                List<ColorRampStyleItem> colorRampSI = style.SearchColorRamps("Viridis").ToList();

                if (colorRampSI != null && colorRampSI.Count() != 0)
                {
                    return colorRampSI[0].ColorRamp;
                }
                else
                {
                    return null;
                }
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Will return the style object and even add it to the current project if it's not already there.
        /// </summary>
        /// <param name="stylePath"></param>
        /// <returns></returns>
        public static StyleProjectItem GetStyleItemProject(string stylePath)
        {
            List<StyleProjectItem> styleItems = Project.Current.GetItems<StyleProjectItem>().Where(x => x.Path == stylePath).ToList();
            if (styleItems == null || styleItems.Count() == 0)
            {
                Project.Current.AddStyle(stylePath);
            }
            StyleProjectItem workingStyle = Project.Current.GetItems<StyleProjectItem>().FirstOrDefault(x => x.Path == stylePath);

            return workingStyle;
        }

        /// <summary>
        /// Will return a text symbol of red color. If a parent symbol is passed, font config will be taken from it.
        /// </summary>
        /// <param name="parentSymbol">Can be null, font config will be taken from it, else arial 10 is the default.</param>
        /// <returns></returns>
        public static TextElement SetMissingTextSymbol(TextElement parentSymbol, string inText = "")
        {
            string missingText = Properties.Resources.ErrorHeadingMissingText;
            if (inText != string.Empty)
            {
                missingText = Properties.Resources.ErrorHeadingMissingText;
            }

            CIMGraphic cimGraphic = parentSymbol.GetGraphic();

            if (cimGraphic != null)
            {
                CIMTextSymbol cIMTextSymbol = cimGraphic.Symbol.Symbol as CIMTextSymbol;
                cIMTextSymbol.SetColor(ColorFactory.Instance.RedRGB);
                cIMTextSymbol.FontFamilyName = "Arial";
                cIMTextSymbol.SetSize(8);
                parentSymbol.SetGraphic(cimGraphic);
            }

            parentSymbol.SetTextProperties(new TextProperties(missingText, parentSymbol.TextProperties.Font, parentSymbol.TextProperties.FontSize, parentSymbol.TextProperties.FontStyle));

            return parentSymbol;

        }

        /// <summary>
        /// Will return a text symbol of red color. If a parent symbol is passed, font config will be taken from it.
        /// </summary>
        /// <param name="parentSymbol">Can be null, font config will be taken from it, else arial 10 is the default.</param>
        /// <returns></returns>
        public static GraphicElement SetMissingPolygonSymbol(GraphicElement parentSymbol)
        {

            CIMGraphic cimGraphic = parentSymbol.GetGraphic();
            if (cimGraphic != null)
            {
                CIMPolygonSymbol cimPolySymbol = cimGraphic.Symbol.Symbol as CIMPolygonSymbol;
                cimPolySymbol.SetColor(ColorFactory.Instance.RedRGB);
                parentSymbol.SetGraphic(cimGraphic);
            }

            return parentSymbol;

        }

        /// <summary>
        /// Default style for line symbol that are missing in a style
        /// </summary>
        /// <param name="parentSymbol">Can be null, font config will be taken from it, else arial 10 is the default.</param>
        /// <returns></returns>
        public static GraphicElement SetMissingLineSymbol(GraphicElement parentSymbol)
        {
            if (parentSymbol != null)
            {
                CIMGraphic cimGraphic = parentSymbol.GetGraphic();
                if (cimGraphic != null)
                {
                    CIMLineGraphic cimLineSymbol = cimGraphic as CIMLineGraphic;
                    cimLineSymbol.Symbol.Symbol.SetColor(ColorFactory.Instance.RedRGB);
                    parentSymbol.SetGraphic(cimGraphic);
                }
            }

            return parentSymbol;

        }

        /// <summary>
        /// Get a grey point symbol for default values or null values
        /// </summary>
        /// <returns></returns>
        public static CIMPointSymbol GetMissingPointSymbol()
        {

            CIMMarker missingMarker = SymbolFactory.Instance.ConstructMarker(103, "Arial", "Regular", 10,ColorFactory.Instance.RedRGB);
            missingMarker.Rotation = 180;

            CIMPointSymbol missingSymbol = SymbolFactory.Instance.ConstructPointSymbol(missingMarker);

            return missingSymbol;
        }

    }
}
