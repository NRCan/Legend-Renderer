using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geodatabase;
using GSC_Legend_Renderer.Dictionaries;

namespace GSC_Legend_Renderer.Services
{
    public class Symbols
    {

        #region GET methods

        /// <summary>
        /// Get a grey line symbol for default values or null values
        /// </summary>
        /// <returns></returns>
        public static ISimpleLineSymbol GetDefaultLineSymbol()
        {
            //Create line symbol for default symbol
            ISimpleLineSymbol lineSym = new SimpleLineSymbol();
            lineSym.Style = esriSimpleLineStyle.esriSLSSolid; //Set it to null

            //Create an RGB object for the fill part
            RgbColor colorLine = new RgbColor();
            colorLine.Red = 0;
            colorLine.Green = 0;
            colorLine.Blue = 0;

            lineSym.Color = colorLine;

            return lineSym;
        }

        /// <summary>
        /// Create Simple fill renderer (for polygons), with outline too.
        /// </summary>
        /// <param name="fillRGB">A List containing red blue green numeric codes to fill surface</param>
        /// <param name="outlineRGB">A list containing red blue green numerci code for surface outline</param>
        /// <param name="outlineWidth">A double value to set ouline width</param>
        /// <returns></returns>
        public static ISimpleRenderer GetSimpleFillRenderer(List<int> fillRGB, List<int> outlineRGB, double outlineWidth)
        {

            //Create a renderer for style
            ISimpleRenderer renderer = new SimpleRenderer();

            //Create poly symbol 
            ISimpleFillSymbol polySym = new SimpleFillSymbol();
            polySym.Style = esriSimpleFillStyle.esriSFSSolid; //Default to solid

            //Create an RGB object for the fill part
            RgbColor inputFill = new RgbColor();
            inputFill.Red = fillRGB[0];
            inputFill.Green = fillRGB[1];
            inputFill.Blue = fillRGB[2];

            //Assign rgb to color
            polySym.Color = inputFill;

            //Create a line symbol for outline.
            ISimpleLineSymbol outlineSym = new SimpleLineSymbol();
            outlineSym.Style = esriSimpleLineStyle.esriSLSSolid; //Default to solid.

            //Create an RGB object for outline part
            RgbColor inputOutline = new RgbColor();
            inputOutline.Red = outlineRGB[0];
            inputOutline.Green = outlineRGB[1];
            inputOutline.Blue = outlineRGB[2];

            //Set color and width of outline
            outlineSym.Color = inputOutline;
            outlineSym.Width = outlineWidth;

            //Set outline symbol to fill symbol
            polySym.Outline = outlineSym;

            //Set symbol propertie of renderer.
            renderer.Symbol = polySym as ISymbol; //Give a symbol object to default symbol

            //Return object
            return renderer;

        }

        /// <summary>
        /// Create a line symbol renderer
        /// </summary>
        /// <param name="lineRGB">A list containing red green blue numerical codes for line color</param>
        /// <param name="lineWidth">A line width</param>
        /// <returns></returns>
        public static ISimpleRenderer GetSimpleLineRenderer(List<int> lineRGB, double lineWidth)
        {

            //Create a renderer for style
            ISimpleRenderer renderer = new SimpleRenderer();

            //Create line symbol for default symbol
            ISimpleLineSymbol lineSym = new SimpleLineSymbol();
            lineSym.Style = esriSimpleLineStyle.esriSLSSolid; //Set it to null

            //Create an RGB object for the fill part
            RgbColor colorLine = new RgbColor();
            colorLine.Red = lineRGB[0];
            colorLine.Green = lineRGB[1];
            colorLine.Blue = lineRGB[2];

            //Add color and width
            lineSym.Color = colorLine;
            lineSym.Width = lineWidth;

            //Set renderer to line symbol
            renderer.Symbol = lineSym as ISymbol;

            return renderer;
        }

        /// <summary>
        /// Creates a point symbol renderer
        /// </summary>
        /// <param name="pointRGB">A list containing red green blue numerical codes for point color</param>
        /// <param name="pointSize">A point size</param>
        /// <returns></returns>
        public static ISimpleRenderer GetPointRenderer(List<int> pointRGB, double pointSize)
        {
            //Create a renderer for style
            ISimpleRenderer pointRenderer = new SimpleRenderer();

            //Create line symbol for default symbol
            ISimpleMarkerSymbol pointSym = new SimpleMarkerSymbol();
            pointSym.Style = esriSimpleMarkerStyle.esriSMSCircle; //Set it to null

            //Create an RGB object for the fill part
            RgbColor colorLine = new RgbColor();
            colorLine.Red = pointRGB[0];
            colorLine.Green = pointRGB[1];
            colorLine.Blue = pointRGB[2];

            //Add color and width
            pointSym.Color = colorLine;
            pointSym.Size = pointSize;

            //Set renderer to line symbol
            pointRenderer.Symbol = pointSym as ISymbol;

            return pointRenderer;
        }

        /// <summary>
        /// Will return as a string, a line symbol type (MarkerLineSymbol, SimpleLineSymbol, etc.)
        /// </summary>
        /// <param name="inSymbol">The input symbol object to valide from</param>
        /// <param name="symbolTypeName"> Output of symbol type name</param>
        /// <returns></returns>
        public static object GetLineSymbolType(ISymbol inSymbol, out string symbolTypeName)
        {
            //Cast input symbol into all kinds of line symbol types
            IMarkerLineSymbol markerLine = inSymbol as MarkerLineSymbol;
            ICartographicLineSymbol cartoLine = inSymbol as CartographicLineSymbol;
            ISimpleLineSymbol simpleLine = inSymbol as SimpleLineSymbol;
            IMultiLayerLineSymbol multilayerLine = inSymbol as MultiLayerLineSymbol;
            IHashLineSymbol hashLine = inSymbol as HashLineSymbol;

            //Init a new object that will contain the correct symbol type
            object correctSymbol = new object();

            //Init symbol type name 
            symbolTypeName = "";

            if (markerLine != null)
            {
                correctSymbol = markerLine;
                symbolTypeName = "IMarkerLineSymbol";
            }
            else if (cartoLine != null)
            {
                correctSymbol = cartoLine;
                symbolTypeName = "ICartographicLineSymbol";
            }
            else if (simpleLine != null)
            {
                correctSymbol = simpleLine;
                symbolTypeName = "ISimpleLineSymbol";
            }
            else if (multilayerLine != null)
            {
                correctSymbol = multilayerLine;
                symbolTypeName = "IMultiLayerLineSymbol";
            }
            else if (hashLine != null)
            {
                correctSymbol = hashLine;
                symbolTypeName = "IHashLineSymbol";
            }

            return correctSymbol;
        }

        /// <summary>
        /// Will return as a string, a point symbol type (ArrowMarker, SimpleMarker, etc.)
        /// </summary>
        /// <param name="inSymbol">The input symbol object to validate from</param>
        /// <param name="symbolTypeName"> Output of symbol type name</param>
        /// <returns></returns>
        public static object GetPointSymbolType(ISymbol inSymbol, out string symbolTypeName)
        {
            //Cast input symbol into all kinds of line symbol types
            IArrowMarkerSymbol arrowPoint = inSymbol as ArrowMarkerSymbol;
            ICharacterMarkerSymbol charPoint = inSymbol as CharacterMarkerSymbol;
            IPictureMarkerSymbol picturePoint = inSymbol as PictureMarkerSymbol;
            ISimpleMarkerSymbol simplePoint = inSymbol as SimpleMarkerSymbol;
            IMultiLayerMarkerSymbol multiPoint = inSymbol as IMultiLayerMarkerSymbol;

            //Init a new object that will contain the correct symbol type
            object correctSymbol = new object();

            //Init symbol type name 
            symbolTypeName = "";

            if (arrowPoint != null)
            {
                correctSymbol = arrowPoint;
                symbolTypeName = "IArrowMarkerSymbol";
            }
            else if (charPoint != null)
            {
                correctSymbol = charPoint;
                symbolTypeName = "ICharacterMarkerSymbol";
            }
            else if (picturePoint != null)
            {
                correctSymbol = picturePoint;
                symbolTypeName = "IPictureMarkerSymbol";
            }
            else if (simplePoint != null)
            {
                correctSymbol = simplePoint;
                symbolTypeName = "ISimpleMarkerSymbol";
            }
            else if (multiPoint != null)
            {
                correctSymbol = multiPoint;
                symbolTypeName = "IMultiLayerMarkerSymbol";
            }

            return correctSymbol;
        }

        /// <summary>
        /// Will return as a string, a polygon symbol type (ArrowMarker, SimpleMarker, etc.)
        /// </summary>
        /// <param name="inSymbol">The input symbol object to validate from</param>
        /// <param name="symbolTypeName"> Output of symbol type name</param>
        /// <returns></returns>
        public static object GetPolygonSymbolType(ISymbol inSymbol, out string symbolTypeName)
        {
            //Cast input symbol into all kinds of line symbol types
            IGradientFillSymbol gradientFill = inSymbol as GradientFillSymbol;
            ILineFillSymbol lineFill = inSymbol as LineFillSymbol;
            IMarkerFillSymbol markerFil = inSymbol as MarkerFillSymbol;
            IPictureFillSymbol pictureFill = inSymbol as PictureFillSymbol;
            ISimpleFillSymbol simpleFill = inSymbol as SimpleFillSymbol;
            IMultiLayerFillSymbol multiFill = inSymbol as IMultiLayerFillSymbol;
            IDotDensityFillSymbol dotFill = inSymbol as IDotDensityFillSymbol;
            IReferenceFillSymbol refFill = inSymbol as IReferenceFillSymbol;

            //Init a new object that will contain the correct symbol type
            object correctSymbol = new object();

            //Init symbol type name 
            symbolTypeName = "";

            if (gradientFill != null)
            {
                correctSymbol = gradientFill;
                symbolTypeName = "IGradientFillSymbol";
            }
            else if (lineFill != null)
            {
                correctSymbol = lineFill;
                symbolTypeName = "ILineFillSymbol";
            }
            else if (markerFil != null)
            {
                correctSymbol = markerFil;
                symbolTypeName = "IMarkerFillSymbol";
            }
            else if (pictureFill != null)
            {
                correctSymbol = pictureFill;
                symbolTypeName = "IPictureFillSymbol";
            }
            else if (simpleFill != null)
            {
                correctSymbol = simpleFill;
                symbolTypeName = "ISimpleFillSymbol";
            }
            else if (multiFill != null)
            {
                correctSymbol = multiFill;
                symbolTypeName = "IMultiLayerFillSymbol";
            }
            else if(dotFill != null)
            {
                correctSymbol = dotFill;
                symbolTypeName = "IDotDensityFillSymbol";
            }
            else if (refFill != null)
            {
                correctSymbol = refFill;
                symbolTypeName = "IReferenceFillSymbol";
            }

            return correctSymbol;
        }

        /// <summary>
        /// Will return as a color object from given symbol, no matter symbol type
        /// </summary>
        /// <param name="inSymbol">The input symbol object to validate from</param>
        /// <param name="symbolTypeName"> Output of symbol type name</param>
        /// <returns></returns>
        public static IColor GetPolygonSymbolColor(ISymbol inSymbol, out string symbolTypeName)
        {
            //Cast input symbol into all kinds of line symbol types
            IGradientFillSymbol gradientFill = inSymbol as GradientFillSymbol;
            ILineFillSymbol lineFill = inSymbol as LineFillSymbol;
            IMarkerFillSymbol markerFil = inSymbol as MarkerFillSymbol;
            IPictureFillSymbol pictureFill = inSymbol as PictureFillSymbol;
            ISimpleFillSymbol simpleFill = inSymbol as SimpleFillSymbol;
            IMultiLayerFillSymbol multiFill = inSymbol as IMultiLayerFillSymbol;

            //Init a new object that will contain the correct symbol type
            IColor symbolColor = GetDefaultPolygonSymbol().Color;

            //Init symbol type name 
            symbolTypeName = "";

            if (gradientFill != null)
            {
                symbolColor = gradientFill.Color;
                symbolTypeName = Constants.ObjectNames.fillTypeGradient;
            }
            else if (lineFill != null)
            {
                symbolColor = lineFill.Color;
                symbolTypeName = Constants.ObjectNames.fillTypeLine;
            }
            else if (markerFil != null)
            {
                symbolColor = markerFil.Color;
                symbolTypeName = Constants.ObjectNames.fillTypeMarker;
            }
            else if (pictureFill != null)
            {
                symbolColor = pictureFill.Color;
                symbolTypeName = Constants.ObjectNames.fillTypePicture;
            }
            else if (simpleFill != null)
            {
                symbolColor = simpleFill.Color;
                symbolTypeName = Constants.ObjectNames.fillTypeSimple;
            }
            else if (multiFill != null)
            {
                symbolColor = multiFill.Color;
                symbolTypeName = Constants.ObjectNames.fillTypeMultilayer;
            }

            return symbolColor;
        }

        /// <summary>
        /// Get a grey point symbol for default values or null values
        /// </summary>
        /// <returns></returns>
        public static ISimpleMarkerSymbol GetDefaultPointSymbol()
        {
            //Create line symbol for default symbol
            ISimpleMarkerSymbol pointSym = new SimpleMarkerSymbol();
            pointSym.Style = esriSimpleMarkerStyle.esriSMSCircle;
            pointSym.Size = 2;

            //Create an RGB object for the fill part
            RgbColor colorLine = new RgbColor();
            colorLine.Red = 0;
            colorLine.Green = 0;
            colorLine.Blue = 0;

            pointSym.Color = colorLine;

            return pointSym;
        }

        /// <summary>
        /// Get an empty fill polygon symbol for default values or null values
        /// </summary>
        /// <returns></returns>
        public static ISimpleFillSymbol GetDefaultPolygonSymbol()
        {
            //Create empty poly symbol for default symbol
            ISimpleFillSymbol polySym = new SimpleFillSymbol();
            polySym.Style = esriSimpleFillStyle.esriSFSNull;
           
            //Create an RGB object for the outline
            RgbColor colorLine = new RgbColor();
            colorLine.Red = 0;
            colorLine.Green = 0;
            colorLine.Blue = 0;

            polySym.Outline.Color = colorLine;

            return polySym;
        }

        /// <summary>
        /// Default style for fill symbol that are missing in a style
        /// </summary>
        /// <returns></returns>
        public static ISimpleFillSymbol GetMissingPolygonSymbol()
        {
            //Create empty poly symbol for default symbol
            ISimpleFillSymbol polySym = new SimpleFillSymbol();
            polySym.Style = esriSimpleFillStyle.esriSFSSolid;

            //Create an RGB object for the outline
            RgbColor colorLine = new RgbColor();
            colorLine.Red = 255;
            colorLine.Green = 0;
            colorLine.Blue = 0;

            polySym.Color = colorLine;

            return polySym;
        }

        /// <summary>
        /// Default style for marker fill symbol that are missing in a style
        /// </summary>
        /// <returns></returns>
        public static IMarkerFillSymbol GetMissingOverlaySymbol()
        {
            //Create empty poly symbol for default symbol
            IMarkerFillSymbol markerFillSym = new MarkerFillSymbol();
            markerFillSym.Style = esriMarkerFillStyle.esriMFSRandom;

            //Create an RGB object for the outline
            RgbColor colorMarker = new RgbColor();
            colorMarker.Red = 255;
            colorMarker.Green = 0;
            colorMarker.Blue = 0;

            RgbColor colorOutline = new RgbColor();
            colorOutline.Red = 0;
            colorOutline.Green = 0;
            colorOutline.Blue = 0;

            markerFillSym.Color = colorMarker;
            markerFillSym.Outline.Color = colorOutline;

            return markerFillSym;
        }

        /// <summary>
        /// Default style for line symbol that are missing in a style
        /// </summary>
        /// <returns></returns>
        public static ISimpleLineSymbol GetMissingLineSymbol()
        {
            //Create empty poly symbol for default symbol
            ISimpleLineSymbol lineSym = new SimpleLineSymbol();
            lineSym.Style = esriSimpleLineStyle.esriSLSSolid;

            //Create an RGB object for the outline
            RgbColor colorLine = new RgbColor();
            colorLine.Red = 255;
            colorLine.Green = 0;
            colorLine.Blue = 0;
            lineSym.Width = 3.0;

            lineSym.Color = colorLine;

            return lineSym;
        }

        /// <summary>
        /// Default style for line symbol that are missing in a style
        /// </summary>
        /// <returns></returns>
        public static ISimpleFillSymbol GetWhiteBackgrounFillSymbol(ILineSymbol existingOutline)
        {
            //Create empty poly symbol for default symbol
            ISimpleFillSymbol polySym = new SimpleFillSymbol();
            polySym.Style = esriSimpleFillStyle.esriSFSSolid;

            //Set outline
            if (existingOutline !=null)
            {
                polySym.Outline = existingOutline;
            }
            else
            {
                //Create empty poly symbol for default symbol
                ISimpleLineSymbol outlineSym = new SimpleLineSymbol();
                outlineSym.Style = esriSimpleLineStyle.esriSLSSolid;

                //Create an RGB object for the outline
                RgbColor colorLine = new RgbColor();
                colorLine.Red = 0;
                colorLine.Green = 0;
                colorLine.Blue = 0;
                outlineSym.Width = 0.72;

                outlineSym.Color = colorLine;
            }
            

            //Create an RGB object for the outline
            RgbColor colorPoly = new RgbColor();
            colorPoly.Red = 255;
            colorPoly.Green = 255;
            colorPoly.Blue = 255;

            polySym.Color = colorPoly;

            return polySym;
        }

        /// <summary>
        /// Will return a text symbol of red color. If a parent symbol is passed, font config will be taken from it.
        /// </summary>
        /// <param name="parentSymbol">Can be null, font config will be taken from it, else arial 10 is the default.</param>
        /// <returns></returns>
        public static ITextSymbol GetMissingTextSymbol(ITextSymbol parentSymbol)
        {
            //Create empty symbol
            ITextSymbol textSymbol = new TextSymbol();

            if (parentSymbol != null)
            {
                textSymbol = Services.ObjectManagement.CopyInputObject(parentSymbol) as ISimpleTextSymbol;
            }
            else
            {
                textSymbol.Font.Name = "Arial";
                textSymbol.Font.Size = 8;
            }

            //Create an RGB object for the outline
            RgbColor fontColor = new RgbColor();
            fontColor.Red = 255;
            fontColor.Green = 0;
            fontColor.Blue = 0;

            textSymbol.Color = fontColor;



            return textSymbol;
            
        }

        /// <summary>
        /// Get a grey point symbol for default values or null values
        /// </summary>
        /// <returns></returns>
        public static ICharacterMarkerSymbol GetMissingPointSymbol()
        {
            //Create line symbol for default symbol
            ICharacterMarkerSymbol charSym = new CharacterMarkerSymbol();
            charSym.CharacterIndex = 103;
            charSym.Size = 10;
            charSym.Angle = 180;

            //Create an RGB object for the fill part
            RgbColor colorPoint = new RgbColor();
            colorPoint.Red = 255;
            colorPoint.Green = 0;
            colorPoint.Blue = 0;

            charSym.Color = colorPoint;

            return charSym;
        }

        #endregion

    }
}
