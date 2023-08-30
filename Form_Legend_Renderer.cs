using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geodatabase;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GSC_Legend_Renderer.Dictionaries;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Display;
using stdole;
using Newtonsoft.Json;
using System.IO;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Framework;
using System.Drawing.Imaging;
using System.Globalization;

namespace GSC_Legend_Renderer
{
    public partial class Form_Legend_Renderer : Form
    {
        #region Main Variables

        //Public variables
        public string dataPath { get; set; }
        public string dataExtension { get; set; }
        public List<string> dataFieldList { get; set; } //Alias will be used inside this list
        public ITable inputDataTable { get; set; }
        public ITable inMemoryDataTable { get; set; }

        List<IElement> legendElementList = new List<IElement>(); //Will hold all legend items to group them at the end of the process.

        public IElement originalCGMLegend { get; set; } //Will be used to move legend  right if a left bracket is found for CGM maps only.

        public IMxDocument currentDoc { get; set; } //Wil be used to keep track of current document throughout methods

        public Dictionary<int, double> arialCharactersWidth { get; set; } //Will be used to calculate text box height based on total lenght of characters

        //Extensions
        public const string gdbExt = Dictionaries.Constants.Extensions.gdbExt;
        public const string mdbExt = Dictionaries.Constants.Extensions.mdbExt;
        public const string xlExt = Dictionaries.Constants.Extensions.xlExt;
        public const string txtExt = Dictionaries.Constants.Extensions.txtExt;
        public const string csvExt = Dictionaries.Constants.Extensions.csvExt;
        public const string dbfExt = Dictionaries.Constants.Extensions.dbfExt;

        //Legend table
        public const string fieldOrder = Dictionaries.Constants.LegendTable.legendOrderField;
        public const string fieldColumn = Dictionaries.Constants.LegendTable.legendColumnField;
        public const string fieldElement = Dictionaries.Constants.LegendTable.legendElementField;
        public const string fieldStyle1 = Dictionaries.Constants.LegendTable.legendStyle1Field;
        public const string fieldStyle2 = Dictionaries.Constants.LegendTable.legendStyle2Field;
        public const string fieldLabel1 = Dictionaries.Constants.LegendTable.legendLabel1Field;
        public const string fieldLabel1Style = Dictionaries.Constants.LegendTable.legendLabel1StyleField;
        public const string fieldLabel2 = Dictionaries.Constants.LegendTable.legendLabel2Field;
        public const string fieldLabel2Style = Dictionaries.Constants.LegendTable.legendLabel2StyleField;
        public const string fieldHeading = Dictionaries.Constants.LegendTable.legendHeadingField;
        public const string fieldDescription = Dictionaries.Constants.LegendTable.legendDescriptionField;

        //UI
        public List<CboxTables> _cboxlayers = new List<CboxTables>();
        #endregion

        #region PROPERTIES

        //Graphics
        public Dictionary<string, IElement> templateGraphicDico { get; set; }

        //Fields
        public string orderFieldUser { get; set; }
        public string columnFieldUser { get; set; }
        public string elementFieldUser { get; set; }
        public string style1FieldUser { get; set; }
        public string style2FieldUser { get; set; }
        public string label1FieldUser { get; set; }
        public string label1StyleFieldUser { get; set; }
        public string label2FieldUser { get; set; }
        public string label2StyleFieldUser { get; set; }
        public string headingFieldUser { get; set; }
        public string descrîptionFieldUser { get; set; }

        //Symbols
        public Dictionary<string, object> fillSymbolDico { get; set; } //Will hold symbol name (01.01.0l) and it's associate style object
        public Dictionary<string, object> lineSymbolDico { get; set; } //Will hold symbol name (01.01.0l) and it's associate style object
        public Dictionary<string, object> markerSymbolDico { get; set; } //Will hold symbol name (01.01.0l) and it's associate style object
        public Dictionary<string, object> textSymbolDico { get; set; } //Will hold symbol name and style object

        //Style
        public IStyleGallery gscStyle { get; set; } //Whole arc map style gallery that holds all style files
        public string gscStylePath { get; set; } //For search purposes inside style gallery

        //JSON
        public string jsonYSpacingFilePath { get; set; }
        public string jsonXSpacingFilePath { get; set; }
        public string jsonOtherFilePath { get; set; }
        public Dictionary<string, Dictionary<string, string>> ySpacings { get; set; }
        public Dictionary<string, string> xSpacings { get; set; }
        public Dictionary<string, string> otherComponents { get; set; }

        //OTHER
        public double columnWidth { get; set; }
        public double elementWidth { get; set; }
        public double elementDescriptGapWidth { get; set; }
        public double descriptionWidth { get; set; }
        public double columnColumnGapWidth { get; set; }
        public double smallDescriptionHeight { get; set; }
        public double smallDescriptionHeightLine { get; set; }
        public double groupDescriptionWidth { get; set; }
        public List<string> heading5Text { get; set; } //Will be used to detect heading 5 elements, which will see their description made italic and indented of 10 points

        public bool isCGMTemplateMXD { get; set; } //Will be used to prevent legend grouping in a CGM template to prevent weird behavior.

        //UI
        public class CboxTables
        {
            public string cboxDataName { get; set; }
            public IStandaloneTable cboxSTLTable { get; set; } //This field can hold a layer object or will be set to null if a string path is available inside layer name
        }
        #endregion

        #region INIT
        public Form_Legend_Renderer()
        {
            InitializeComponent();
            FillTableViewCombobox();
            this.comboBox_SelectTable.SelectedIndexChanged += comboBox_SelectTable_SelectedIndexChanged;

            //Init of dictionaries
            templateGraphicDico = new Dictionary<string, IElement>();
            ySpacings = new Dictionary<string, Dictionary<string, string>>();
            xSpacings = new Dictionary<string, string>();
            otherComponents = new Dictionary<string, string>();

        }

        #endregion

        #region EVENTS

        /// <summary>
        /// Will be triggered when user selects a new table path or layer
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox_SelectTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox senderBox = sender as ComboBox;
            if (senderBox.SelectedIndex != -1)
            {
                //Init field list
                InitFieldList();

                //Fill the fields comboboxes
                FillFieldsComboboxes();
            }
        }

        /// <summary>
        /// Will close the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// Will initiate a graphic legend creation
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_Start_Click(object sender, EventArgs e)
        {

            //Process
            CreateLegend();

        }

        /// <summary>
        /// Will open a generic browse dialog in which the user can select the data that he wants to load inside the form
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_selectTable_Click(object sender, EventArgs e)
        {
            //Open dialog
            dataPath = Services.Dialog.GetDataPrompt(this.Handle.ToInt32(), Properties.GSC_LegendRenderer_Resources.Dialog_SelectTableTitle);

            //Unset list before change
            this.comboBox_SelectTable.DataSource = null;

            //Update
            _cboxlayers.Add(new CboxTables { cboxDataName = dataPath, cboxSTLTable = null });
            this.comboBox_SelectTable.DataSource = _cboxlayers;
            this.comboBox_SelectTable.DisplayMember = "cboxDataName";
            this.comboBox_SelectTable.ValueMember = "cboxSTLTable";
            this.comboBox_SelectTable.SelectedIndex = _cboxlayers.Count - 1;

        }

        private void comboBox_orderField_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.comboBox_orderField.SelectedIndex != -1)
            {
                orderFieldUser = this.comboBox_orderField.SelectedItem.ToString();
            }
        }

        private void comboBox_ColumnField_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (this.comboBox_ColumnField.SelectedIndex != -1)
            {
                columnFieldUser = this.comboBox_ColumnField.SelectedItem.ToString();
            }


        }

        private void comboBox_ElementField_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.comboBox_ElementField.SelectedIndex != -1)
            {
                elementFieldUser = this.comboBox_ElementField.SelectedItem.ToString();
            }


        }

        private void comboBox_Style1Field_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.comboBox_Style1Field.SelectedIndex != -1)
            {
                style1FieldUser = this.comboBox_Style1Field.SelectedItem.ToString();
            }


        }

        private void comboBox_Style2Field_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.comboBox_Style2Field.SelectedIndex != -1)
            {
                style2FieldUser = this.comboBox_Style2Field.SelectedItem.ToString();
            }


        }

        private void comboBox_Label1Field_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.comboBox_Label1Field.SelectedIndex != -1)
            {
                label1FieldUser = this.comboBox_Label1Field.SelectedItem.ToString();
            }


        }

        private void comboBox_Label1StyleField_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.comboBox_Label1StyleField.SelectedIndex != -1)
            {
                label1StyleFieldUser = this.comboBox_Label1StyleField.SelectedItem.ToString();
            }


        }

        private void comboBox_Label2Field_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.comboBox_Label2Field.SelectedIndex != -1)
            {
                label2FieldUser = this.comboBox_Label2Field.SelectedItem.ToString();
            }


        }

        private void comboBox_Label2StyleField_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.comboBox_Label2StyleField.SelectedIndex != -1)
            {
                label2StyleFieldUser = this.comboBox_Label2StyleField.SelectedItem.ToString();
            }


        }

        private void comboBox_HeadingField_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.comboBox_HeadingField.SelectedIndex != -1)
            {
                headingFieldUser = this.comboBox_HeadingField.SelectedItem.ToString();
            }

        }

        private void comboBox_DescriptionField_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (this.comboBox_DescriptionField.SelectedIndex != -1)
            {
                descrîptionFieldUser = this.comboBox_DescriptionField.SelectedItem.ToString();
            }
        }

        #endregion

        #region METHODS

        /// <summary>
        /// Will retrieve a list of table fields depending on the data extension type
        /// Will later be used to fill in the UI
        /// </summary>
        private void InitFieldList()
        {

            if (this.comboBox_SelectTable.SelectedIndex != -1)
            {
                CboxTables selectedlTables = this.comboBox_SelectTable.SelectedItem as CboxTables;

                if (selectedlTables.cboxSTLTable != null)
                {
                    #region FOR TABLE VIEWS IN TABLE OF CONTENT

                    IStandaloneTable selectedSTL = selectedlTables.cboxSTLTable;

                    //Fill the field list with field names
                    dataFieldList = Services.Tables.GetFieldList(selectedSTL.Table, true, selectedSTL);


                    #endregion
                }
                else if (selectedlTables.cboxDataName != null)
                {
                    #region FOR EXTERNAL DATA
                    if (dataPath.Contains(gdbExt) || dataPath.Contains(mdbExt))
                    {
                        //Open as table
                        inputDataTable = Services.Tables.OpenTableFromStringFaster(dataPath);


                        //Fill the field list with field names
                        dataFieldList = Services.Tables.GetFieldList(inputDataTable, true);

                    }
                    else if (dataPath.Contains(dbfExt))
                    {
                        //Get the table name and extension only
                        string fileNameOnly = System.IO.Path.GetFileName(dataPath);

                        //Get the excel workspace factory
                        IWorkspace dbfWorkspace = Services.Workspace.AccessWorkspace(dataPath);
                        inputDataTable = Services.Tables.OpenTableFromWorkspace(dbfWorkspace, fileNameOnly);

                        //Fill the field list with field names
                        dataFieldList = Services.Tables.GetFieldList(inputDataTable, true);

                        //Release workspace so user can keep on editing
                        Services.ObjectManagement.ReleaseObject(dbfWorkspace);

                    }
                    else if (dataPath.Contains(txtExt) || dataPath.Contains(csvExt))
                    {
                        //Get the sheet name only
                        string fileNameOnly = System.IO.Path.GetFileName(dataPath);

                        //Get the excel workspace factory
                        IWorkspace txtFileWorkspace = Services.Workspace.AccessTextfileWorkspace(dataPath);
                        inputDataTable = Services.Tables.OpenTableFromWorkspace(txtFileWorkspace, fileNameOnly);

                        //Fill the field list with field names
                        dataFieldList = Services.Tables.GetFieldList(inputDataTable, true);

                        //Release workspace so user can keep on editing
                        Services.ObjectManagement.ReleaseObject(txtFileWorkspace);

                    }
                    else if (dataPath.Contains(xlExt))
                    {
                        //Get the sheet name
                        string[] splitedPath = dataPath.Split('\\');

                        //Build path to the file itself without the sheet
                        string dataPathFileOnly = string.Empty;

                        foreach (string parts in splitedPath)
                        {
                            if (parts != splitedPath[splitedPath.Length - 1])
                            {
                                if (dataPathFileOnly != string.Empty)
                                {
                                    dataPathFileOnly = dataPathFileOnly + "\\" + parts;
                                }
                                else
                                {
                                    dataPathFileOnly = parts;
                                }

                            }

                        }

                        //Get the sheet name only
                        string fileSheetName = splitedPath[splitedPath.Length - 1];

                        //Get the excel workspace factory
                        IWorkspace excelWorkspace = Services.Workspace.AccessExcelWorkspace(dataPathFileOnly);
                        inputDataTable = Services.Tables.OpenTableFromWorkspace(excelWorkspace, fileSheetName);

                        //Fill the field list with field names
                        dataFieldList = Services.Tables.GetFieldList(inputDataTable, true);

                        dataFieldList.Add(string.Empty);

                        //Release workspace so user can keep on editing
                        Services.ObjectManagement.ReleaseObject(inputDataTable);
                        Services.ObjectManagement.ReleaseObject(excelWorkspace);
                    }

                    #endregion
                }
            }

        }

        /// <summary>
        /// Will create the legend.
        /// </summary>
        public void CreateLegend()
        {
            
            //Continu if table is valid
            if (this.comboBox_SelectTable.Text != null && this.comboBox_SelectTable.Text != string.Empty)
            {
                //Get json config files - check properties for path
                ValidateJsonFilesExistance();
                BuildOtherComponentsDictionary();

                //Validate if needed style has been loaded
                if (ValidateStyleFile())
                {

                    #region GET and BUILD information
                    //Get arial character widths
                    arialCharactersWidth = GetArialCharacterWidth();

                    //Set document units, else if it's not in mm the legend will be looking bad...
                    currentDoc = (IMxDocument)ArcMap.Application.Document;
                    esriUnits originalUnits = SetDocumentUnits(currentDoc, esriUnits.esriMillimeters);

                    //Force delay update
                    currentDoc.DelayUpdateContents = true;

                    //Variables
                    IElement parentElement = null; //Will be used to keep parent element that has embedded children
                    double originalYSpacing = 0;
                    double ySpacing = originalYSpacing;//Keep track of Y spacing
                    double xSpacing = 0; //Keep track of X spacing
                    bool firstIterationBreaker = true;
                    IElement lastElement = null;
                    string lastElementType = string.Empty;
                    int lastColumn = 1;
                    IElement waitingLeftBracket = null; //Will be used to move in Y axis an added left bracket that needs to know ySpacing for it's horizontal brother element.
                    IElement upLeftBracket = null; //Will be used to complete left bracket when end-point is reached
                    IElement waitingCenterLeftBracket = null; //Will be used to move bracket annotation when full bracket has been completed.
                    IElement annotationBracket = null; //Will be used to set text first and then move it.
                    IElement waitingRightBracket = null; //Will be used to move in XY axis an added right bracket
                    IElement upRightBracket = null; //Will be used to complete right bracket when end-point is reached.
                    IElement waitingCenterRightBracket = null; //Wil be used to move bracket associated map unit when full bracket has been completed
                    int howManyRightBrackets = 0; //Will be used to recalculate x spacing in case more columns are asked by user and that some right brackets are also found
                    Tuple<IElement, IElement, IElement, IElement> bracketMapUnit = new Tuple<IElement, IElement, IElement, IElement>(null, null, null, null); //Will be used to keep unit box for bracket and replace it at the right anchor when bracket is done drawing.
                    Tuple<double, double> anchorPoint = GetAnchorPointStart(); //TODO Find if mxd is a CGM one or not.
                    originalYSpacing = anchorPoint.Item2; //Synchronise with initial calculate anchor.
                    Tuple<double, double> anchorPointParent = new Tuple<double, double>(0, 0);
                    heading5Text = new List<string>(); //Init
                    double currentIteration = 0.0; //Will be used if user has forgot to enter an order.
                    bool nullOrderBreaker = false; //Will be used to show error message to user if null values are found, but only once.

                    //Get selection of graphic (will help doing a clear)
                    IPageLayout currentLayout = currentDoc.ActiveView as IPageLayout;
                    IGraphicsContainerSelect currentGrapSelection = currentLayout as IGraphicsContainerSelect;

                    //Get template graphics
                    GetTemplateGraphicList();

                    //Get legend table
                    ITable legendTable;
                    CboxTables selectedTable = this.comboBox_SelectTable.SelectedItem as CboxTables;
                    if (selectedTable.cboxSTLTable == null)
                    {
                        legendTable = Services.Tables.OpenTableFromString(selectedTable.cboxDataName);
                    }
                    else
                    {
                        legendTable = selectedTable.cboxSTLTable as ITable;
                    }


                    //Iterate through table and add elements
                    IQueryFilter ascendingOrderQuery = new QueryFilter();
                    IQueryFilterDefinition ascendingOrderQueryPostfix = ascendingOrderQuery as IQueryFilterDefinition;
                    ascendingOrderQueryPostfix.PostfixClause = "ORDER BY " + orderFieldUser;
                    ICursor legendCursor = legendTable.Search(ascendingOrderQuery, true);

                    //Get symbols
                    if (lineSymbolDico == null)
                    {
                        GetSymbols(Dictionaries.Constants.Styles.styleLineClass);
                    }
                    if (fillSymbolDico == null)
                    {
                        GetSymbols(Dictionaries.Constants.Styles.styleFillClass);
                    }
                    if (markerSymbolDico == null)
                    {
                        GetSymbols(Dictionaries.Constants.Styles.styleMarkerClass);
                    }
                    if (textSymbolDico == null)
                    {
                        GetSymbols(Dictionaries.Constants.Styles.styleTextClass);
                    }

                    //Get lower bound
                    double legendYLowerBound = GetCGMLegendLowerBound(Constants.YSpacings.legendEnd_Citation, Constants.Graphics.cgmCitation);

                    #endregion

                    #region CURSOR inside Legend Table
                    //Get fields indexes
                    int elementFieldIndex = legendCursor.FindField(elementFieldUser);
                    int orderFieldIndex = legendCursor.FindField(orderFieldUser);
                    int style1FieldIndex = legendCursor.FindField(style1FieldUser);
                    int style2FieldIndex = legendCursor.FindField(style2FieldUser);
                    int labelFieldIndex = legendCursor.FindField(label1FieldUser);
                    int descriptionFieldIndex = legendCursor.FindField(descrîptionFieldUser);
                    int headingFieldIndex = legendCursor.FindField(headingFieldUser);
                    int columnFieldIndex = legendCursor.FindField(columnFieldUser);
                    int label2FieldIndex = legendCursor.FindField(label2FieldUser);
                    int label1StyleFieldIndex = legendCursor.FindField(label1StyleFieldUser);
                    int label2StyleFieldIndex = legendCursor.FindField(label2StyleFieldUser);

                    int currentColumn = 1; //Default

                    IRow legendRow;
                    while ((legendRow = legendCursor.NextRow()) != null)
                    {
                        //Variables
                        currentIteration = currentIteration + 1.0;
                        double currentOrder = currentIteration;
                        Double.TryParse(legendRow.Value[orderFieldIndex].ToString(), out currentOrder);
                        string currentStyle1 = legendRow.Value[style1FieldIndex].ToString();
                        string currentStyle2 = legendRow.Value[style2FieldIndex].ToString();
                        string currentLabel1 = legendRow.Value[labelFieldIndex].ToString();
                        string currentLabel2 = legendRow.Value[label2FieldIndex].ToString();
                        string currentDescription = legendRow.Value[descriptionFieldIndex].ToString();
                        string currentHeading = legendRow.Value[headingFieldIndex].ToString();
                        string currentElement = legendRow.Value[elementFieldIndex].ToString();
                        string currentLabel1Style = legendRow.Value[label1StyleFieldIndex].ToString();
                        string currentLabel2Style = legendRow.Value[label2StyleFieldIndex].ToString();

                        //Clean and replace < characters from description
                        //Having <bol></bol> within description along an extra < symbol, breaks the bolding of the heading within the description
                        if (currentDescription != string.Empty && currentHeading != string.Empty && heading5Text.Count == 0)
                        {
                            currentDescription = currentDescription.Replace("<", "&lt;");
                        }


                        //Get related graphic, if exists
                        IElement currentElementObject = null;
                        if (templateGraphicDico.ContainsKey(currentElement))
                        {
                            currentElementObject = Services.ObjectManagement.CopyInputObject(templateGraphicDico[currentElement]) as IElement;
                        }
                        

                        //Manage null order
                        if (legendRow.Value[orderFieldIndex].ToString() == string.Empty || legendRow.Value[orderFieldIndex].ToString() == "<Null>" || legendRow.Value[orderFieldIndex] == null)
                        {
                            if (!nullOrderBreaker)
                            {
                                MessageBox.Show("Missing value found in " + Constants.LegendTable.legendOrderField + " field. This might cause some problems on item rendering. \n\nLook for:\n" +
                                  currentElement + " " + currentHeading + " " + currentDescription); //TODO change this message for localized one and better text.
                                nullOrderBreaker = true;
                            }
                            
                        }

                        //Set heading5 trigger for special style symbols
                        ///Two case, either user repeats heading5 text in wanted embedded symbols, or
                        ///uses the latest element named HEADING5_END without duplicating heading5 text in all symbols
                        if (heading5Text.Count > 0)
                        {
                            //Add any duplicate
                            if (heading5Text[0] == currentHeading)
                            {
                                heading5Text.Add(currentHeading);
                            }

                            //Detect suddent misrupt of heading 5 text in heading column
                            if (heading5Text[0] != currentHeading && heading5Text.Count > 1)
                            {
                                heading5Text = new List<string>(); //reinitialize
                            }
                            //Detect explicit use of a heading 5 end element
                            if (currentElement == Constants.Graphics.heading5_end)
                            {
                                heading5Text = new List<string>(); //reinitialize
                            }
                        }     
                        
                        //Get spacings dictionnary loaded up
                        if (!firstIterationBreaker && !currentElement.Contains(Constants.Graphics.keywordBracket))
                        {
                            ySpacing = GetYSpacing(lastElement, lastElementType, currentElement, anchorPoint.Item2);
                            xSpacing = GetXSpacing(currentElement);

                        }
                        else
                        {
                            BuildYSpacingsDictionary();
                            BuildXSpacingsDictionary();

                            //Widths
                            if (xSpacings != null)
                            {
                                columnWidth = GetXSpacing(Constants.Graphics.columnWidth);
                                elementWidth = GetXSpacing(Constants.Graphics.elementWidth);
                                elementDescriptGapWidth = GetXSpacing(Constants.Graphics.elementDescriptionGapWidth);
                                descriptionWidth = GetXSpacing(Constants.Graphics.descriptionWidth);
                                columnColumnGapWidth = GetXSpacing(Constants.Graphics.columnColumnGapWidth);
                                smallDescriptionHeight = Constants.YSpacings.smallDescriptionHeightLimit;
                                smallDescriptionHeightLine = Constants.YSpacings.smallDescriptionHeightLimitLines;
                                groupDescriptionWidth = GetXSpacing(Constants.Graphics.groupDescriptionWidth);

                                xSpacing = GetXSpacing(currentElement);
                            }


                            firstIterationBreaker = false;
                        }

                        //Manage columns
                        if (!this.checkBox_autoCalculateColumns.Checked)
                        {
                            //Track column number change in table
                            if (int.TryParse(legendRow.Value[columnFieldIndex].ToString(), out currentColumn))
                            {
                                currentColumn = Convert.ToInt32(legendRow.Value[columnFieldIndex]);
                            }
                        }
                        else
                        {
                            //Track column change with auto-calculate
                            if (legendYLowerBound != 0.0)
                            {
                                if ((anchorPoint.Item2 - ySpacing - currentElementObject.Geometry.Envelope.Height) < legendYLowerBound)
                                {
                                    currentColumn++;
                                }
                            }
                        }

                        if (currentColumn > 1 && lastColumn != currentColumn)
                        {
                            //Get x spacing based on how many brackets were found in previous column
                            double rightBracketSpacing = 0;
                            if (howManyRightBrackets > 0)
                            {
                                rightBracketSpacing = (descriptionWidth + elementDescriptGapWidth + elementWidth + GetXSpacing(Constants.Graphics.bracketRightCenter) + GetXSpacing(Constants.Graphics.unitBoxBracket));
                            }

                            //Move to right and reset Y.
                            ySpacing = 0; //Reset y spacing so it appears at the top of the page 
                            anchorPoint = new Tuple<double, double>(anchorPoint.Item1 + columnWidth + columnColumnGapWidth + rightBracketSpacing, originalYSpacing);

                            //Adjust  anchorpoint in case current element as an inner centered y anchor (CC, CL and CR)
                            if (templateGraphicDico.ContainsKey(currentElement))
                            {
                                IElement newColumnFirstElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[currentElement]) as IElement;
                                //Get anchor type
                                IElementProperties3 newColumnProp = newColumnFirstElement as IElementProperties3;
                                esriAnchorPointEnum currentAnchorPointType = newColumnProp.AnchorPoint;

                                if (currentAnchorPointType == esriAnchorPointEnum.esriCenterPoint || currentAnchorPointType == esriAnchorPointEnum.esriLeftMidPoint || currentAnchorPointType == esriAnchorPointEnum.esriRightMidPoint)
                                {
                                    ySpacing = (newColumnFirstElement.Geometry.Envelope.Height / 2.0);
                                }
                            }
                            lastColumn = currentColumn;

                            //Reset right bracket number
                            howManyRightBrackets = 0;

                        }

                        #region HEADINGS
                        if (currentElement.Contains(Constants.Graphics.heading1.Substring(0, 6)))
                        {

                            //Get appropriate element
                            if (templateGraphicDico.ContainsKey(currentElement))
                            {
                                IElement headElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[currentElement]) as IElement;
                                IElementProperties headElProp = headElement as IElementProperties;

                                string originalElementName = headElProp.Name;

                                //TODO special cases heading 3 like description, heading 4 UL

                                #region Move to right anchor

                                //Set new anchor
                                anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing);


                                //Set height for heading3 
                                if (currentElement.Contains(Constants.Graphics.heading3))
                                {
                                    //Recalculate height
                                    string tempGroupHeadingDescription = currentHeading;
                                    if (currentDescription != null)
                                    {
                                        tempGroupHeadingDescription = currentHeading + currentDescription;
                                    }
                                    double heading3Height = GetTextHeight(tempGroupHeadingDescription, descriptionWidth, Constants.TextConfiguration.lineHeight);

                                    //Set new envelope
                                    SetRectnagularPolygonFromAnchorTypeAndHeight(headElement, anchorPoint, heading3Height);
                                }
                                else
                                {
                                    //Set new envelope
                                    SetRectangularPolygonFromAnchorType(headElement, anchorPoint);
                                }

                                //Move
                                ITransform2D transElement = headElement as ITransform2D;
                                transElement.Move(xSpacing, 0); //Move accordingly to x spacing if any

                                #endregion

                                //Rename
                                headElProp.Name = headElProp.Name + currentOrder.ToString();

                                //Add Heading text
                                ITextElement tElement = headElement as ITextElement;
                                if (currentHeading == null || currentHeading == string.Empty || currentHeading == " ")
                                {
                                    currentHeading = Constants.TextConfiguration.missingText;
                                    tElement.Symbol = Services.Symbols.GetMissingTextSymbol(tElement.Symbol);
                                }

                                //Special case for heading 3 since we can't have bolded all caps setting inside a graphic along
                                //no cap and not bolded description.
                                if (currentElement.Contains(Constants.Graphics.heading3))
                                {
                                    currentHeading = Constants.TextConfiguration.tagAllCaps + Constants.TextConfiguration.tagBold + currentHeading + Constants.TextConfiguration.endTagBold + Constants.TextConfiguration.endTagAllCaps + " ";

                                    //Add Description to text - Only for heading 3 in theory
                                    if (!IsTextEmpty(currentDescription))
                                    {
                                        //Add header if needed
                                        if (!IsTextEmpty(currentHeading))
                                        {
                                            currentHeading = currentHeading + currentDescription;
                                        }
                                    }

                                }
                                if (currentElement.Contains(Constants.Graphics.heading5))
                                {
                                    //Keep heading text so it can be used for a trigger to modify description style for heading 5 only.
                                    heading5Text.Add(currentHeading);
                                }

                                tElement.Text = currentHeading;

                                //Manage style if needed
                                if (currentStyle1 != "")
                                {
                                    if (textSymbolDico.ContainsKey(currentStyle1))
                                    {
                                        ISimpleTextSymbol inStyleSymbol = textSymbolDico[currentStyle1] as ISimpleTextSymbol;
                                        ISimpleTextSymbol currentStyleSymbol = tElement.Symbol as ISimpleTextSymbol;
                                        currentStyleSymbol.Font = inStyleSymbol.Font;
                                        currentStyleSymbol.Color = inStyleSymbol.Color;
                                        currentStyleSymbol.Size = currentStyleSymbol.Size; //Force size else incoming style might be too big.
                                        currentStyleSymbol.VerticalAlignment = currentStyleSymbol.VerticalAlignment; //Force vertical center for text else incoming style might be set to else where.
                                        tElement.Symbol = Services.ObjectManagement.CopyInputObject(currentStyleSymbol) as ISimpleTextSymbol;
                                    }
                                    else
                                    {
                                        //Missing or wrong style 
                                        tElement.Symbol = Services.Symbols.GetMissingTextSymbol(tElement.Symbol);
                                    }


                                }


                                //Add base element
                                currentDoc.ActiveView.GraphicsContainer.AddElement(headElement, 0);
                                currentDoc.ActiveView.GraphicsContainer.BringToFront(currentGrapSelection.SelectedElements);

                                //Unselect
                                currentGrapSelection.UnselectElement(headElement);

                                //Add to legend list
                                legendElementList.Add(headElement);

                                //Keep name
                                lastElement = headElement;
                                lastElementType = originalElementName;


                            }

                        }

                        #endregion

                        #region MAP UNITS 
                        if (currentElement == Constants.Graphics.unitBox || currentElement == Constants.Graphics.unitSplit || 
                            currentElement == Constants.Graphics.unitindent1 || currentElement == Constants.Graphics.unitindent2)
                        {

                            //Get appropriate element
                            IElement unitBoxElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[currentElement]) as IElement;
                            IElementProperties unitBoxElProp = unitBoxElement as IElementProperties;
                            currentDoc.ActiveView.GraphicsContainer.AddElement(unitBoxElement as IElement, 0);
                            string originalElementName = unitBoxElProp.Name;

                            //Init empty dem element if ever needed
                            IElement demUnitBoxElement = null;

                            #region Move to right anchor

                            //Set new anchor
                            anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing); //New anchor point with proper move inside it
                            SetRectangularPolygonFromAnchorType(unitBoxElement, anchorPoint);

                            //Move
                            if (currentElement == Constants.Graphics.unitindent1 || currentElement == Constants.Graphics.unitindent2)
                            {
                                ITransform2D transElement = unitBoxElement as ITransform2D;
                                transElement.Move(xSpacing, 0); //Move accordingly to x spacing if any
                            }

                            #endregion

                            //Rename
                            unitBoxElProp.Name = unitBoxElProp.Name + currentOrder.ToString();

                            //Symbolize
                            IElement unitBoxLabelElement = new MarkerElement();
                            IGroupElement inGroupElement = unitBoxElement as IGroupElement;

                            //Unselect
                            currentGrapSelection.UnselectElement(unitBoxElement);

                            if (inGroupElement != null)
                            {
                                //Check geometry of inner elements, if it's all lines

                                for (int el = 0; el < inGroupElement.ElementCount; el++)
                                {
                                    IElement innerElement = inGroupElement.Element[el];
                                    if (el == 0)
                                    {
                                        SetPolygonFill(innerElement, currentStyle1, true);

                                        //Add label
                                        if (currentLabel1 == null || currentLabel1 == string.Empty || currentLabel1 == " ")
                                        {
                                            currentLabel1 = Constants.TextConfiguration.missingText;
                                        }

                                        unitBoxLabelElement = AddLabelInUnitBox(currentLabel1, innerElement, currentDoc, anchorPoint, Constants.Graphics.UnitBoxType.split1, currentLabel1Style);

                                    }
                                    else if (el > 0)
                                    {
                                        SetPolygonFill(innerElement, currentStyle2, true);

                                        //Add label
                                        if (currentLabel2 == null || currentLabel2 == string.Empty || currentLabel2 == " ")
                                        {
                                            currentLabel2 = Constants.TextConfiguration.missingText;
                                        }
                                        unitBoxLabelElement = AddLabelInUnitBox(currentLabel2, innerElement, currentDoc, anchorPoint, Constants.Graphics.UnitBoxType.split2, currentLabel2Style);
                                    }

                                }
                            }
                            else
                            {
                                //Symbolize
                                demUnitBoxElement = SetPolygonFill(unitBoxElement, currentStyle1, true, true, anchorPoint, currentStyle2);
                                
                                //Add
                                currentDoc.ActiveView.GraphicsContainer.AddElement(demUnitBoxElement as IElement, 0);

                                //Add label
                                if (currentLabel1 == null || currentLabel1 == string.Empty || currentLabel1 == " ")
                                {
                                    currentLabel1 = Constants.TextConfiguration.missingText;
                                }

                                unitBoxLabelElement = AddLabelInUnitBox(currentLabel1, unitBoxElement, currentDoc, anchorPoint, Constants.Graphics.UnitBoxType.normal, currentLabel1Style);

                            }

                            //Move label and/or dem
                            if (currentElement == Constants.Graphics.unitindent1 || currentElement == Constants.Graphics.unitindent2)
                            {
                                //DEM
                                if (this.checkBox_DEMBoxes.Checked)
                                {
                                    ITransform2D transDEMElement = demUnitBoxElement as ITransform2D;
                                    transDEMElement.Move(xSpacing, 0); //Move accordingly to x spacing if any
                                }

                                //LABEL
                                ITransform2D transLabelElement = unitBoxLabelElement as ITransform2D;
                                transLabelElement.Move(xSpacing, 0); //Move accordingly to x spacing if any
                            }

                            //Keep name
                            lastElement = unitBoxElement;
                            lastElementType = originalElementName;

                            //Add header if needed
                            if (currentHeading != null && currentHeading != string.Empty && currentHeading != " ")
                            {
                                currentDescription = Constants.TextConfiguration.tagBold + currentHeading + Constants.TextConfiguration.endTagBold + " " + currentDescription;
                            }

                            //Add Description
                            IElement newDescriptionElement = AddDescription(currentDescription, unitBoxElement, currentDoc, anchorPoint, originalElementName);
                            double descriptionHeight = newDescriptionElement.Geometry.Envelope.Height;
                            if (descriptionHeight > smallDescriptionHeight)
                            {
                                //Reset anchor point for next element
                                if (currentColumn != 0)
                                {
                                    anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - descriptionHeight); //New anchor point with proper move inside it
                                                                                                                                       
                                }

                                //Keep name
                                lastElement = newDescriptionElement;
                                lastElementType = Constants.Graphics.description;

                            }

                            //Move description
                            if (currentElement == Constants.Graphics.unitindent1 || currentElement == Constants.Graphics.unitindent2)
                            {
                                ITransform2D transDescElement = newDescriptionElement as ITransform2D;
                                transDescElement.Move(xSpacing, 0); //Move accordingly to x spacing if any
                            }

                            //Keep element if for bracket
                            if (currentColumn == 0)
                            {
                                bracketMapUnit = new Tuple<IElement, IElement, IElement, IElement>(unitBoxElement, unitBoxLabelElement, newDescriptionElement, demUnitBoxElement);

                                //Reset anchor point
                                anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 + ySpacing);
                            }

                            //Add to legend list
                            if (demUnitBoxElement != null)
                            {
                                legendElementList.Add(demUnitBoxElement as IElement);
                            }
                            legendElementList.Add(unitBoxLabelElement as IElement);
                            legendElementList.Add(unitBoxElement as IElement);



                        }
                        #endregion

                        #region THIN UNITS (MAP UNIT LINE)
                        if (currentElement == Constants.Graphics.unitLine)
                        {
                            //Get appropriate element
                            IElement thinUnitElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[currentElement]) as IElement;
                            IElementProperties thinUnitElProp = thinUnitElement as IElementProperties;

                            string originalElementName = thinUnitElProp.Name;

                            #region Move to right anchor

                            //Set new anchor
                            anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing); //New anchor point with proper move inside it
                            SetRectangularPolygonFromAnchorType(thinUnitElement, anchorPoint);

                            #endregion

                            #region Set symbol

                            thinUnitElement = SetThinUnitSymbol(thinUnitElement, currentStyle1, currentStyle2);

                            #endregion

                            //Rename
                            thinUnitElProp.Name = thinUnitElProp.Name + currentOrder.ToString();

                            //Build group
                            IGroupElement3 thinUnitGroup = GetGroupLegendElement(thinUnitElProp.Name);
                            currentDoc.ActiveView.GraphicsContainer.MoveElementToGroup(thinUnitElement, thinUnitGroup as IGroupElement);

                            //Add label if needed
                            if (currentLabel1 != null && currentLabel1 != string.Empty && currentLabel1 != " ")
                            {
                                IElement thinUnitLabel = AddLabelInUnitBox(currentLabel1, thinUnitElement, currentDoc, anchorPoint, Constants.Graphics.UnitBoxType.line, currentLabel1Style);
                                currentDoc.ActiveView.GraphicsContainer.MoveElementToGroup(thinUnitLabel, thinUnitGroup as IGroupElement);
                            }

                            //Keep name
                            lastElement = thinUnitElement;
                            lastElementType = originalElementName;

                            //Add header if needed
                            if (currentHeading != null && currentHeading != string.Empty && currentHeading != " ")
                            {
                                currentDescription = Constants.TextConfiguration.tagBold + currentHeading + Constants.TextConfiguration.endTagBold + " " + currentDescription;
                            }

                            //Add Description
                            IElement newDescriptionElement = AddDescription(currentDescription, thinUnitElement, currentDoc, anchorPoint, originalElementName);
                            double descriptionHeight = newDescriptionElement.Geometry.Envelope.Height;
                            if (descriptionHeight > smallDescriptionHeight)
                            {
                                //Reset anchor point for next element
                                anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - descriptionHeight); //New anchor point with proper move inside it
                                                                                                                                   //Keep name
                                lastElement = newDescriptionElement;
                                lastElementType = Constants.Graphics.description;

                            }

                            //Add base element
                            currentDoc.ActiveView.GraphicsContainer.AddElement(thinUnitGroup as IElement, 0);

                            //Unselect
                            currentGrapSelection.UnselectElement(thinUnitGroup as IElement);

                            //Add to legend list
                            legendElementList.Add(thinUnitGroup as IElement);
                        }
                        #endregion

                        #region EMBEDDED MAP UNIT (PARENT/CHILD)
                        if (currentElement == Constants.Graphics.unitParent || currentElement == Constants.Graphics.subUnitParentChild || currentElement == Constants.Graphics.subUnitParentChildLine)
                        {
                            //Get appropriate element
                            IElement parentChildElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[currentElement]) as IElement;

                            //Reset element and grow parent if needed
                            if (currentElement == Constants.Graphics.unitParent)
                            {
                                //Reset parent element
                                parentElement = null;

                                //Keep normal unit box instead
                                parentChildElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[Constants.Graphics.unitBox]) as IElement;
                            }

                            IElementProperties parentChildElProp = parentChildElement as IElementProperties;

                            string originalElementName = parentChildElProp.Name;

                            //Apply conversion factor
                            double parentHeight = 0;
                            double parentChildHeight = parentChildElement.Geometry.Envelope.Height;
                            #region Move to right anchor

                            //Set new anchor
                            if (parentElement != null && lastElement == parentElement)
                            {
                                parentHeight = parentElement.Geometry.Envelope.Height;
                                double newYSpacing = ySpacing + (parentHeight - parentChildHeight);
                                anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - newYSpacing); //New anchor point with proper move inside it
                            }
                            else
                            {
                                anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing); //New anchor point with proper move inside it
                            }

                            if ((currentElement == Constants.Graphics.subUnitParentChild || currentElement == Constants.Graphics.subUnitParentChildLine) && parentElement != null)
                            {
                                parentHeight = parentElement.Geometry.Envelope.Height;
                                double newHeightFromChild = parentChildHeight + parentHeight;
                                SetRectnagularPolygonFromAnchorTypeAndHeight(parentElement, anchorPointParent, newHeightFromChild);

                                //Reset anchor point since height of the element has changed.
                                anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - parentChildHeight);
                            }

                            SetRectangularPolygonFromAnchorType(parentChildElement, anchorPoint);

                            //Move
                            ITransform2D transElement = parentChildElement as ITransform2D;
                            transElement.Move(xSpacing, 0); //Move accordingly to x spacing if any is found (child unit case)

                            #endregion

                            //Rename
                            parentChildElProp.Name = parentChildElProp.Name + currentOrder.ToString();

                            //Label
                            IElement labelParentChild = null;

                            //Symbolize
                            if (currentElement == Constants.Graphics.subUnitParentChildLine)
                            {
                                SetThinUnitSymbol(parentChildElement, currentStyle1, currentStyle2);
                                if (currentLabel1 != null && currentLabel1 != string.Empty && currentLabel1 != " ")
                                {
                                    labelParentChild = AddLabelInUnitBox(currentLabel1, parentChildElement, currentDoc, anchorPoint, Constants.Graphics.UnitBoxType.child_line, currentLabel1Style);
                                    //Move label
                                    ITransform2D transLabelElement = labelParentChild as ITransform2D;
                                    transLabelElement.Move(xSpacing, 0); //Move accordingly to x spacing if any is found (child unit case)
                                }

                            }
                            else
                            {
                                SetPolygonFill(parentChildElement, currentStyle1, true);

                                //Add label
                                if (currentLabel1 == null || currentLabel1 == string.Empty || currentLabel1 == " ")
                                {
                                    currentLabel1 = Constants.TextConfiguration.missingText;
                                }
                                labelParentChild = AddLabelInUnitBox(currentLabel1, parentChildElement, currentDoc, anchorPoint, Constants.Graphics.UnitBoxType.normal, currentLabel1Style);

                                //Move label
                                ITransform2D transLabelElement = labelParentChild as ITransform2D;
                                transLabelElement.Move(xSpacing, 0); //Move accordingly to x spacing if any is found (child unit case)
                            }

                            //Add header if needed
                            if (currentHeading != null && currentHeading != string.Empty && currentHeading != " ")
                            {
                                currentDescription = Constants.TextConfiguration.tagBold + currentHeading + Constants.TextConfiguration.endTagBold + " " + currentDescription;
                            }

                            //Add Description
                            IElement newDescriptionElement = AddDescription(currentDescription, parentChildElement, currentDoc, anchorPoint, originalElementName);
                            double descriptionHeight = newDescriptionElement.Geometry.Envelope.Height;
                            if (descriptionHeight > smallDescriptionHeight)
                            {
                                //Keep name
                                lastElement = newDescriptionElement;
                                lastElementType = Constants.Graphics.description;

                                //Reset height of unit parent
                                if (currentElement == Constants.Graphics.unitParent)
                                {
                                    //double newHeight = parentChildElement.Geometry.Envelope.Height + descriptionHeight;
                                    SetRectnagularPolygonFromAnchorTypeAndHeight(parentChildElement, anchorPoint, descriptionHeight);
                                }

                                //Reset anchor for next unit to be added
                                if (currentElement == Constants.Graphics.subUnitParentChild || currentElement == Constants.Graphics.subUnitParentChildLine)
                                {
                                    double newDescriptionHeight = descriptionHeight - smallDescriptionHeight;
                                    double newParentHeight = parentElement.Geometry.Envelope.Height + newDescriptionHeight;
                                    SetRectnagularPolygonFromAnchorTypeAndHeight(parentElement, anchorPointParent, newParentHeight); //Reset parent box height
                                    anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - newDescriptionHeight); //New anchor point with proper move inside it
                                }

                            }
                            //Keep name
                            lastElement = parentChildElement;
                            lastElementType = originalElementName;

                            //Keep parent information
                            if (currentElement == Constants.Graphics.unitParent)
                            {
                                parentElement = parentChildElement;
                                anchorPointParent = anchorPoint;
                            }

                            //Build group
                            IGroupElement3 unitParentChildGroup = GetGroupLegendElement(parentChildElProp.Name);
                            currentDoc.ActiveView.GraphicsContainer.MoveElementToGroup(parentChildElement, unitParentChildGroup as IGroupElement);
                            if (labelParentChild != null)
                            {
                                currentDoc.ActiveView.GraphicsContainer.MoveElementToGroup(labelParentChild, unitParentChildGroup as IGroupElement);
                            }


                            //Add base element
                            currentDoc.ActiveView.GraphicsContainer.AddElement(unitParentChildGroup as IElement, 0);

                            //Unselect
                            currentGrapSelection.UnselectElement(unitParentChildGroup as IElement);

                            //Add to legend list
                            legendElementList.Add(unitParentChildGroup as IElement);
                        }

                        #endregion

                        #region MARKERS/POINTS
                        if (currentElement == Constants.Graphics.point || currentElement == Constants.Graphics.pointAngle || currentElement == Constants.Graphics.pointAngleLine)
                        {
                            //Build marker element
                            Tuple<double, double> offset = null;
                            IElement pointElement = BuildMarker(currentElement, currentOrder, currentStyle1, out offset);

                            #region Move to right anchor

                            //Set new anchor
                            anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing); //New anchor point with proper move inside it
                            SetPointFromAnchorType(pointElement, anchorPoint, offset);

                            #endregion

                            //Add measurement value label
                            if (currentLabel1 != null && currentLabel1 != string.Empty && currentLabel1 != " ")
                            {

                                //Find proper placement for label
                                Constants.Styles.MarkerLabelPositioning placement = Constants.Styles.MarkerLabelPositioning.FromCenterToUpperLeft;
                                if (currentElement == Constants.Graphics.pointAngleLine)
                                {
                                    placement = Constants.Styles.MarkerLabelPositioning.FromCenterToUpperRight;
                                }
                                else if (currentElement == Constants.Graphics.point)
                                {
                                    placement = Constants.Styles.MarkerLabelPositioning.FromCenterToUpperRightTight;
                                }

                                if (currentLabel1Style == null ||  currentLabel1Style == " ")
                                {
                                    currentLabel1Style = string.Empty;
                                }
                                IElement markerLabel1 = AddLabelToMarker(currentLabel1, pointElement, currentDoc, anchorPoint, placement, currentLabel1Style);

                                //Add second label if any
                                if (currentLabel2 != null && currentLabel2 != string.Empty && currentLabel2 != " ")
                                {
                                    if (currentLabel2Style == null || currentLabel2Style == " ")
                                    {
                                        currentLabel2Style = string.Empty;
                                    }

                                    AddLabelToMarker(currentLabel2, markerLabel1, currentDoc, anchorPoint, Constants.Styles.MarkerLabelPositioning.RightAboveCenter, currentLabel2Style, placement);
                                }
                            }

                            //Add base element
                            currentDoc.ActiveView.GraphicsContainer.AddElement(pointElement, 0);
                            currentDoc.ActiveView.GraphicsContainer.SendToBack(currentGrapSelection.SelectedElements);

                            //Unselect
                            currentGrapSelection.UnselectElement(pointElement);

                            //Keep
                            lastElement = pointElement;
                            lastElementType = currentElement;

                            //Add to legend list
                            legendElementList.Add(lastElement);

                            //Add Description
                            IElement newDescriptionElement = AddDescription(currentDescription, pointElement, currentDoc, anchorPoint, lastElementType, true);
                            double descriptionHeight = newDescriptionElement.Geometry.Envelope.Height;
                            if (descriptionHeight > smallDescriptionHeightLine)
                            {
                                //Reset anchor point for next element
                                IElementProperties3 pointPropElement = pointElement as IElementProperties3;
                                esriAnchorPointEnum currentAnchorPointType = pointPropElement.AnchorPoint;
                                //double descriptionAdjustement = descriptionHeight;
                                double descriptionAdjustement = descriptionHeight - Constants.YSpacings.markerMeanHeight;

                                anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - descriptionAdjustement); //New anchor point with proper move inside it

                                //anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - descriptionHeight);

                                //Keep name
                                lastElement = newDescriptionElement;
                                lastElementType = Constants.Graphics.description;

                            }
                        }
                        #endregion

                        #region LINES
                        if (currentElement == Constants.Graphics.beach || currentElement == Constants.Graphics.moraines || currentElement == Constants.Graphics.dunes
                            || currentElement == Constants.Graphics.landslide || currentElement == Constants.Graphics.line || currentElement == Constants.Graphics.wave
                            || currentElement == Constants.Graphics.lineDouble || currentElement == Constants.Graphics.lineDoubleFLow || currentElement == Constants.Graphics.lineDoubleFlip)
                        {
                            //Get appropriate element
                            IElement lineElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[currentElement]) as IElement;
                            IElementProperties lineElProp = lineElement as IElementProperties;

                            string originalElementName = lineElProp.Name;

                            #region Move to right anchor

                            //Set new anchor
                            anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing); //New anchor point with proper move inside it
                            if (lineElement.Geometry.GeometryType == esriGeometryType.esriGeometryPolygon)
                            {
                                SetRectangularPolygonFromAnchorType(lineElement, anchorPoint);
                            }
                            else if (lineElement.Geometry.GeometryType == esriGeometryType.esriGeometryPolyline)
                            {
                                SetLineFromAnchorType(lineElement, lastElement, anchorPoint, elementWidth);
                            }


                            #endregion

                            #region Set symbol

                            IGroupElement inGroupElement = lineElement as IGroupElement;
                            if (inGroupElement != null)
                            {
                                //Check geometry of inner elements, if it's all lines
                                for (int el = 0; el < inGroupElement.ElementCount; el++)
                                {
                                    ILineElement currentShapeElement = inGroupElement.Element[el] as ILineElement;
                                    IElementProperties currentShapeProp = inGroupElement.Element[el] as IElementProperties;
                                    double currentLineWidth = currentShapeElement.Symbol.Width;

                                    if (currentShapeProp.Name != Constants.Graphics.subLineDoubleFLowBottom && currentShapeProp.Name != Constants.Graphics.subLineDoubleFLowMiddle)
                                    {
                                        if (lineSymbolDico.ContainsKey(currentStyle1))
                                        {
                                            currentShapeElement.Symbol = Services.ObjectManagement.CopyInputObject(lineSymbolDico[currentStyle1]) as ILineSymbol;

                                        }
                                        else
                                        {
                                            //Apply missing style
                                            ISimpleLineSymbol missingFillSymbol = Services.Symbols.GetMissingLineSymbol();
                                            currentShapeElement.Symbol = missingFillSymbol;
                                        }
                                        
                                    }
                                    if (currentShapeProp.Name == Constants.Graphics.subLineDoubleFLowBottom)
                                    {
                                        //If something isn't found in style2 revert to first one
                                        if (currentStyle2 != string.Empty && currentStyle2 != " " && currentStyle2 != null)
                                        {
                                            // For double line, two style might be used if it's not inside a double line flow symbol
                                            if (lineSymbolDico.ContainsKey(currentStyle2))
                                            {
                                                currentShapeElement.Symbol = Services.ObjectManagement.CopyInputObject(lineSymbolDico[currentStyle2]) as ILineSymbol;

                                                if (currentElement == Constants.Graphics.lineDoubleFLow)
                                                {
                                                    //For double line flow symbol keep bottom line just like the top one and take style2 field for the flow symbol
                                                    currentShapeElement.Symbol = Services.ObjectManagement.CopyInputObject(lineSymbolDico[currentStyle1]) as ILineSymbol;
                                                }
                                            }

                                        }
                                        else
                                        {
                                            currentShapeElement.Symbol = Services.ObjectManagement.CopyInputObject(lineSymbolDico[currentStyle1]) as ILineSymbol;
                                        }

                                    }

                                    if (currentShapeProp.Name == Constants.Graphics.subLineDoubleFLowMiddle)
                                    {
                                        if (currentStyle2 != null && lineSymbolDico.ContainsKey(currentStyle2))
                                        {
                                            currentShapeElement.Symbol = Services.ObjectManagement.CopyInputObject(lineSymbolDico[currentStyle2]) as ILineSymbol;
                                        }
                                        else
                                        {
                                            //Apply missing style
                                            ISimpleLineSymbol missingFillSymbol = Services.Symbols.GetMissingLineSymbol();
                                            currentShapeElement.Symbol = missingFillSymbol;
                                        }
                                    }

                                }
                            }
                            else
                            {
                                ILineElement areaLineElement = lineElement as ILineElement;
                                if (lineSymbolDico.ContainsKey(currentStyle1))
                                {
                                    areaLineElement.Symbol = Services.ObjectManagement.CopyInputObject(lineSymbolDico[currentStyle1]) as ILineSymbol;

                                }
                                else
                                {
                                    //Apply missing style
                                    ISimpleLineSymbol missingFillSymbol = Services.Symbols.GetMissingLineSymbol();
                                    areaLineElement.Symbol = missingFillSymbol;
                                }
                            }
                            #endregion

                            //Rename
                            lineElProp.Name = lineElProp.Name + currentOrder.ToString();

                            //Add base element
                            currentDoc.ActiveView.GraphicsContainer.AddElement(lineElement, 0);
                            currentDoc.ActiveView.GraphicsContainer.SendToBack(currentGrapSelection.SelectedElements);

                            //Unselect
                            currentGrapSelection.UnselectElement(lineElement);

                            //Keep name
                            lastElement = lineElement;
                            lastElementType = originalElementName;

                            //Add to legend list
                            legendElementList.Add(lastElement);

                            //Add Description
                            //NOTES: Usually lines have anchors in the center, that needs special attention
                            IElement newDescriptionElement = AddDescription(currentDescription, lineElement, currentDoc, anchorPoint, originalElementName, true);
                            double descriptionHeight = newDescriptionElement.Geometry.Envelope.Height;
                            if (descriptionHeight >= smallDescriptionHeightLine)
                            {
                                double descriptionAdjustement = descriptionHeight;
                                if (lineElement.Geometry.Envelope.Height == 0)
                                {
                                    descriptionAdjustement = descriptionAdjustement - Constants.YSpacings.lineHeight0DescriptionHeightAdjustement;
                                }
                                else
                                {
                                    descriptionAdjustement = descriptionAdjustement - (lineElement.Geometry.Envelope.Height / 2.0);
                                }

                                //Reset anchor point for next element
                                anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - descriptionAdjustement); //New anchor point with proper move inside it

                                //Keep name
                                lastElement = newDescriptionElement;
                                lastElementType = Constants.Graphics.description;

                            }
                        }

                        #endregion

                        #region SYMBOLIZED AREAS / OVERLAYS
                        if (currentElement == Constants.Graphics.overlay || currentElement == Constants.Graphics.blob)
                        {

                            //Get appropriate element
                            IElement symAreasElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[currentElement]) as IElement;
                            IElementProperties symAreaElProp = symAreasElement as IElementProperties;

                            string originalElementName = symAreaElProp.Name;

                            #region Move to right anchor

                            //Set new anchor
                            anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing); //New anchor point with proper move inside it
                            SetPolygonFromAnchorType(symAreasElement, anchorPoint);

                            #endregion

                            //Symbolize
                            SetPolygonFill(symAreasElement, currentStyle1, false);

                            //Rename
                            symAreaElProp.Name = symAreaElProp.Name + currentOrder.ToString();

                            //Add base element
                            currentDoc.ActiveView.GraphicsContainer.AddElement(symAreasElement, 0);
                            currentDoc.ActiveView.GraphicsContainer.SendToBack(currentGrapSelection.SelectedElements);

                            //Unselect
                            currentGrapSelection.UnselectElement(symAreasElement);

                            //Keep name
                            lastElement = symAreasElement;
                            lastElementType = originalElementName;

                            //Add to legend list
                            legendElementList.Add(lastElement);

                            //Add annotation or a marker in the middle if needed
                            if ((currentLabel1 != null && currentLabel1Style != string.Empty) || (currentStyle2 != null && currentStyle2 != string.Empty))
                            {
                                //Init
                                Tuple<double, double> offset = new Tuple<double, double>(0, 0);
                                IElement symAreaPointElement = null;


                                //For a label in the middle
                                if (textSymbolDico != null && textSymbolDico.Count > 0 && textSymbolDico.ContainsKey(currentLabel1Style))
                                {

                                    //Get appropriate element
                                    symAreaPointElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[Constants.Graphics.annotationBlob]) as IElement;

                                    //Create new text graphic with default style
                                    IMarkerElement saMarkerElement = symAreaPointElement as IMarkerElement;
                                    IMultiLayerMarkerSymbol mlCharacterElement = saMarkerElement.Symbol as IMultiLayerMarkerSymbol;
                                    ICharacterMarkerSymbol saCharacterElement = Services.ObjectManagement.CopyInputObject(mlCharacterElement.Layer[0]) as ICharacterMarkerSymbol;

                                    int labelCharset = 0;
                                    if (int.TryParse(currentLabel1, out labelCharset))
                                    {
                                        saCharacterElement.CharacterIndex = Convert.ToInt16(currentLabel1);
                                    }
                                    else
                                    {
                                        saCharacterElement.CharacterIndex = labelCharset;
                                    }

                                    mlCharacterElement.DeleteLayer(mlCharacterElement.Layer[0]);
                                    mlCharacterElement.AddLayer(saCharacterElement);
                                    mlCharacterElement.MoveLayer(saCharacterElement, 0);

                                    saMarkerElement.Symbol = Services.ObjectManagement.CopyInputObject(mlCharacterElement) as IMultiLayerMarkerSymbol;

                                    //Set new anchor
                                    SetPointFromAnchorType(symAreaPointElement, anchorPoint, offset);

                                }
                                //For a marker in the middle
                                if (markerSymbolDico != null && markerSymbolDico.Count > 0 && markerSymbolDico.ContainsKey(currentStyle2))
                                {
                                    //Build marker element
                                    symAreaPointElement = BuildMarker(Constants.Graphics.point, currentOrder, currentStyle2, out offset);

                                    //Set new anchor
                                    Tuple<double, double> overlayPointAnchor = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - symAreasElement.Geometry.Envelope.Height / 2.0);
                                    SetPointFromAnchorType(symAreaPointElement, overlayPointAnchor, offset);
                                }

                                //For a new color fill on the symbol
                                if (fillSymbolDico != null && fillSymbolDico.Count > 0 && fillSymbolDico.ContainsKey(currentStyle2))
                                {
                                    //Symbolize
                                    symAreaPointElement = SetOverlayFillColor(symAreasElement, currentStyle1, currentStyle2);
                                }

                                if (symAreaPointElement != null)
                                {

                                    //Add base element
                                    currentDoc.ActiveView.GraphicsContainer.AddElement(symAreaPointElement, 0);
                                    currentDoc.ActiveView.GraphicsContainer.BringToFront(currentGrapSelection.SelectedElements);

                                    //Unselect
                                    currentGrapSelection.UnselectElement(symAreaPointElement);

                                    //Add to legend list
                                    legendElementList.Add(symAreaPointElement);
                                }

                            }


                            //Add Description
                            IElement newDescriptionElement = AddDescription(currentDescription, symAreasElement, currentDoc, anchorPoint, originalElementName);
                            double descriptionHeight = newDescriptionElement.Geometry.Envelope.Height;
                            if (descriptionHeight > smallDescriptionHeight)
                            {
                                //Reset anchor point for next element
                                anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - descriptionHeight); //New anchor point with proper move inside it
                                                                                                                                   //Keep name
                                lastElement = newDescriptionElement;
                                lastElementType = Constants.Graphics.description;

                            }
                        }

                        #endregion

                        #region LINE HEADERS/BREAKERS

                        if (currentElement == Constants.Graphics.breakLine)
                        {

                            //Get appropriate element
                            IElement lineElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[currentElement]) as IElement;
                            IElementProperties lineElProp = lineElement as IElementProperties;

                            string originalElementName = lineElProp.Name;

                            #region Move to right anchor

                            //Set new anchor
                            anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing); //New anchor point with proper move inside it
                            IPolyline polylineElement = lineElement.Geometry as IPolyline;
                            SetLineFromAnchorType(lineElement, lastElement, anchorPoint, polylineElement.Length);

                            #endregion

                            #region Set symbol
                            ILineElement breakLineElement = lineElement as ILineElement;
                            if (currentStyle1 != string.Empty && currentStyle1 != " " && currentStyle1 != null)
                            {
                                if (lineSymbolDico.ContainsKey(currentStyle1))
                                {
                                    breakLineElement.Symbol = Services.ObjectManagement.CopyInputObject(lineSymbolDico[currentStyle1]) as ILineSymbol;
                                }
                                else
                                {
                                    //Apply missing style
                                    ISimpleLineSymbol missingFillSymbol = Services.Symbols.GetMissingLineSymbol();
                                    breakLineElement.Symbol = missingFillSymbol;
                                }
                                

                            }

                            
                            #endregion

                            //Rename
                            lineElProp.Name = lineElProp.Name + currentOrder.ToString();

                            //Add base element
                            currentDoc.ActiveView.GraphicsContainer.AddElement(lineElement, 0);
                            currentDoc.ActiveView.GraphicsContainer.SendToBack(currentGrapSelection.SelectedElements);

                            //Unselect
                            currentGrapSelection.UnselectElement(lineElement);

                            //Keep
                            lastElement = lineElement;
                            lastElementType = originalElementName;

                            //Add to legend list
                            legendElementList.Add(lastElement);

                        }

                        #endregion

                        #region ANNOTATIONS
                        if (currentElement == Constants.Graphics.annotationBreak)
                        {
                            //Get appropriate element
                            if (templateGraphicDico.ContainsKey(currentElement))
                            {
                                IElement annoElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[currentElement]) as IElement;
                                IElementProperties annoElProp = annoElement as IElementProperties;

                                string originalElementName = annoElProp.Name;

                                #region Move to right anchor

                                //Set new anchor
                                anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing);
                                SetRectangularPolygonFromAnchorType(annoElement, anchorPoint);

                                //Move
                                ITransform2D transElement = annoElement as ITransform2D;
                                transElement.Move(xSpacing, 0); //Move accordingly to x spacing if any

                                #endregion

                                //Rename
                                annoElProp.Name = annoElProp.Name + currentOrder.ToString();

                                //Add Heading text
                                ITextElement tElement = annoElement as ITextElement;
                                if (currentLabel1 == null || currentLabel1 == string.Empty || currentLabel1 == " ")
                                {
                                    currentLabel1 = Constants.TextConfiguration.missingText;
                                    tElement.Symbol = Services.Symbols.GetMissingTextSymbol(tElement.Symbol);
                                }

                                tElement.Text = currentLabel1;

                                //Add base element
                                currentDoc.ActiveView.GraphicsContainer.AddElement(annoElement, 0);
                                currentDoc.ActiveView.GraphicsContainer.BringToFront(currentGrapSelection.SelectedElements);

                                //Unselect
                                currentGrapSelection.UnselectElement(annoElement);

                                //Add to legend list
                                legendElementList.Add(annoElement);

                            }
                        }

                        #endregion

                        #region LEFT BRACKETS

                        //Process left bracket that was waiting to have Y spacing
                        if (waitingLeftBracket != null && !currentElement.Contains(Constants.Graphics.keywordBracket))
                        {
                            ITransform2D transElement = waitingLeftBracket as ITransform2D;
                            transElement.Move(0, -ySpacing); //Move accordingly to x spacing if any

                            waitingLeftBracket = null;
                        }

                        #region UPPER ELEMENT
                        if (currentElement == Constants.Graphics.bracketLeftUpper)
                        {
                            //Get appropriate element
                            if (templateGraphicDico.ContainsKey(currentElement))
                            {
                                IElement leftBracketElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[currentElement]) as IElement;
                                IElementProperties leftBracketdElProp = leftBracketElement as IElementProperties;

                                string originalElementName = leftBracketdElProp.Name;

                                #region Move to right anchor

                                //Set new anchor
                                Tuple<double, double> leftBracketAnchor = new Tuple<double, double>(anchorPoint.Item1, lastElement.Geometry.Envelope.YMax);
                                MoveItemToAnchorPoint(leftBracketElement, leftBracketAnchor);

                                //Move
                                ITransform2D transElement = leftBracketElement as ITransform2D;
                                transElement.Move(GetXSpacing(currentElement), 0); //Move accordingly to x spacing if any

                                #endregion

                                //Rename
                                leftBracketdElProp.Name = leftBracketdElProp.Name + currentOrder.ToString();

                                //Add base element
                                currentDoc.ActiveView.GraphicsContainer.AddElement(leftBracketElement, 0);
                                currentDoc.ActiveView.GraphicsContainer.BringToFront(currentGrapSelection.SelectedElements);

                                //Unselect
                                currentGrapSelection.UnselectElement(leftBracketElement);

                                // Don't Keep name

                                //Add to legend list
                                legendElementList.Add(leftBracketElement);

                                //Keep in waiting line, to be move in Y axis
                                waitingLeftBracket = leftBracketElement;
                                upLeftBracket = leftBracketElement;

                            }
                        }

                        #endregion

                        #region LOWER ELEMENT
                        if (currentElement == Constants.Graphics.bracketLeftLower)
                        {
                            //Get appropriate element
                            if (templateGraphicDico.ContainsKey(currentElement))
                            {
                                #region END BRACKET NOTCH
                                IElement leftEndBracketElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[currentElement]) as IElement;
                                IElementProperties leftEndBracketdElProp = leftEndBracketElement as IElementProperties;

                                string originalElementName = leftEndBracketdElProp.Name;

                                #region Move to right anchor

                                //Set new anchor
                                Tuple<double, double> bottomBracketAnchor = new Tuple<double, double>(anchorPoint.Item1, lastElement.Geometry.Envelope.YMin + leftEndBracketElement.Geometry.Envelope.Height);
                                MoveItemToAnchorPoint(leftEndBracketElement, bottomBracketAnchor);

                                //Move
                                ITransform2D transElement = leftEndBracketElement as ITransform2D;
                                transElement.Move(GetXSpacing(currentElement), 0); //Move accordingly to x spacing if any

                                #endregion

                                //Rename
                                leftEndBracketdElProp.Name = leftEndBracketdElProp.Name + currentOrder.ToString();

                                //Add base element
                                currentDoc.ActiveView.GraphicsContainer.AddElement(leftEndBracketElement, 0);
                                currentDoc.ActiveView.GraphicsContainer.BringToFront(currentGrapSelection.SelectedElements);

                                //Unselect
                                currentGrapSelection.UnselectElement(leftEndBracketElement);

                                //Add to legend list
                                legendElementList.Add(leftEndBracketElement);

                                #endregion

                                #region MIDDLE BRACKET NOTCH
                                IElement middleBracketElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[Constants.Graphics.bracketLeftCenter]) as IElement;
                                IElementProperties middleBracketdElProp = middleBracketElement as IElementProperties;

                                string middleElementName = middleBracketdElProp.Name;

                                #region Move to right anchor

                                //Set new anchor
                                double middleBracketY = upLeftBracket.Geometry.Envelope.YMin - Math.Abs(((leftEndBracketElement.Geometry.Envelope.YMax - upLeftBracket.Geometry.Envelope.YMin) / 2.0));
                                Tuple<double, double> middleBracketAnchor = new Tuple<double, double>(anchorPoint.Item1, middleBracketY);
                                MoveItemToAnchorPoint(middleBracketElement, middleBracketAnchor);

                                //Move
                                ITransform2D transMiddleElement = middleBracketElement as ITransform2D;
                                transMiddleElement.Move(GetXSpacing(Constants.Graphics.bracketLeftCenter), 0); //Move accordingly to x spacing if any

                                #endregion

                                //Rename
                                middleBracketdElProp.Name = middleBracketdElProp.Name + currentOrder.ToString();

                                //Add base element
                                currentDoc.ActiveView.GraphicsContainer.AddElement(middleBracketElement, 0);
                                currentDoc.ActiveView.GraphicsContainer.BringToFront(currentGrapSelection.SelectedElements);

                                //Unselect
                                currentGrapSelection.UnselectElement(middleBracketElement);

                                //Keep element to use with bracket annotation
                                waitingCenterLeftBracket = middleBracketElement;

                                //Add to legend list
                                legendElementList.Add(middleBracketElement);


                                #endregion

                                #region BRACKET SPINE 1
                                IElement spine1BracketElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[Constants.Graphics.bracketSpine]) as IElement;
                                IElementProperties spine1BracketdElProp = spine1BracketElement as IElementProperties;

                                string spine1ElementName = spine1BracketdElProp.Name;

                                #region Move to right anchor

                                //Set new anchor
                                IPoint startPointLine = new ESRI.ArcGIS.Geometry.Point();
                                startPointLine.X = upLeftBracket.Geometry.Envelope.XMin;
                                startPointLine.Y = upLeftBracket.Geometry.Envelope.YMin;

                                IPoint endPointLine = new ESRI.ArcGIS.Geometry.Point();
                                endPointLine.X = middleBracketElement.Geometry.Envelope.XMax;
                                endPointLine.Y = middleBracketElement.Geometry.Envelope.YMax;

                                IPolyline spineLine = new ESRI.ArcGIS.Geometry.Polyline() as IPolyline;
                                spineLine.ToPoint = endPointLine;
                                spineLine.FromPoint = startPointLine;

                                spine1BracketElement.Geometry = spineLine;

                                #endregion

                                //Rename
                                spine1BracketdElProp.Name = spine1BracketdElProp.Name + currentOrder.ToString();

                                //Add base element
                                currentDoc.ActiveView.GraphicsContainer.AddElement(spine1BracketElement, 0);
                                currentDoc.ActiveView.GraphicsContainer.BringToFront(currentGrapSelection.SelectedElements);

                                //Unselect
                                currentGrapSelection.UnselectElement(spine1BracketElement);

                                //Add to legend list
                                legendElementList.Add(spine1BracketElement);

                                #endregion

                                #region BRACKET SPINE 2
                                IElement spine2BracketElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[Constants.Graphics.bracketSpine]) as IElement;
                                IElementProperties spine2BracketdElProp = spine2BracketElement as IElementProperties;

                                string spine2ElementName = spine2BracketdElProp.Name;

                                #region Move to right anchor

                                //Set new anchor
                                IPoint startPointLine2 = new ESRI.ArcGIS.Geometry.Point();
                                startPointLine2.X = middleBracketElement.Geometry.Envelope.XMax;
                                startPointLine2.Y = middleBracketElement.Geometry.Envelope.YMin;

                                IPoint endPointLine2 = new ESRI.ArcGIS.Geometry.Point();
                                endPointLine2.X = leftEndBracketElement.Geometry.Envelope.XMin;
                                endPointLine2.Y = leftEndBracketElement.Geometry.Envelope.YMax;

                                IPolyline spineLine2 = new Polyline() as IPolyline;
                                spineLine2.ToPoint = endPointLine2;
                                spineLine2.FromPoint = startPointLine2;

                                spine2BracketElement.Geometry = spineLine2;

                                #endregion

                                //Rename
                                spine2BracketdElProp.Name = spine2BracketdElProp.Name + currentOrder.ToString();

                                //Add base element
                                currentDoc.ActiveView.GraphicsContainer.AddElement(spine2BracketElement, 0);
                                currentDoc.ActiveView.GraphicsContainer.BringToFront(currentGrapSelection.SelectedElements);

                                //Unselect
                                currentGrapSelection.UnselectElement(spine2BracketElement);

                                //Add to legend list
                                legendElementList.Add(spine2BracketElement);

                                #endregion
                            }
                        }

                        #endregion

                        #region ANNOTATION ELEMENT
                        if (currentElement == Constants.Graphics.annoBracket)
                        {
                            //Get appropriate element
                            if (templateGraphicDico.ContainsKey(currentElement))
                            {
                                IElement annoBracketElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[currentElement]) as IElement;
                                IElementProperties annoBracketdElProp = annoBracketElement as IElementProperties;

                                string originalElementName = annoBracketdElProp.Name;

                                //Rename
                                annoBracketdElProp.Name = annoBracketdElProp.Name + currentOrder.ToString();

                                //Add Heading text
                                ITextElement aElement = annoBracketElement as ITextElement;
                                if (currentLabel1 == null || currentLabel1 == string.Empty || currentLabel1 == " ")
                                {
                                    currentLabel1 = Constants.TextConfiguration.missingText;
                                    aElement.Symbol = Services.Symbols.GetMissingTextSymbol(aElement.Symbol);
                                }

                                aElement.Text = Constants.TextConfiguration.tagAllCaps + currentLabel1 + Constants.TextConfiguration.endTagAllCaps;

                                //Move to current anchor for a start
                                MoveItemToAnchorPoint(annoBracketElement, anchorPoint);


                                //Add base element
                                currentDoc.ActiveView.GraphicsContainer.AddElement(annoBracketElement, 0);
                                currentDoc.ActiveView.GraphicsContainer.BringToFront(currentGrapSelection.SelectedElements);

                                //Unselect
                                currentGrapSelection.UnselectElement(annoBracketElement);

                                //Keep
                                annotationBracket = annoBracketElement;

                                //Add to legend list
                                legendElementList.Add(annoBracketElement);

                            }
                        }
                        #endregion

                        //Process left bracket annotation waiting to have the right anchor point
                        if (waitingCenterLeftBracket != null && annotationBracket != null)
                        {
                            //Set new anchor
                            Tuple<double, double> annoBracketAnchor = new Tuple<double, double>(waitingCenterLeftBracket.Geometry.Envelope.XMin, waitingCenterLeftBracket.Geometry.Envelope.YMax - waitingCenterLeftBracket.Geometry.Envelope.Height / 2.0);
                            MoveItemToAnchorPoint(annotationBracket, annoBracketAnchor);

                            //Move accordingly to x spacing from json file
                            ITransform2D transAnnoBracketElement = annotationBracket as ITransform2D;
                            //transAnnoBracketElement.Move(GetXSpacing(currentElement), 0); 
                            
                            //Move at the right anchor
                            double xMove = Math.Abs(waitingCenterLeftBracket.Geometry.Envelope.XMin - annotationBracket.Geometry.Envelope.XMax);
                            transAnnoBracketElement.Move(xMove, 0);

                            //Move because flipping 90 degree text is done around mid-center, but doesn't impact coordinates, only visually
                            transAnnoBracketElement.Move(annotationBracket.Geometry.Envelope.Width / 2.0, 0);
                            transAnnoBracketElement.Move(-annotationBracket.Geometry.Envelope.Height / 2.0, 0);
                            transAnnoBracketElement.Move(GetXSpacing(currentElement), 0);

                            waitingCenterLeftBracket = null;
                            annotationBracket = null;
                        }

                        #endregion

                        #region RIGHT BRACKETS

                        //Process left bracket that was waiting to have Y spacing
                        if (waitingRightBracket != null && !currentElement.Contains(Constants.Graphics.keywordBracket))
                        {
                            ITransform2D transElement = waitingRightBracket as ITransform2D;
                            transElement.Move(GetXSpacing(Constants.Graphics.bracketRightUpper) + (currentColumn * GetXSpacing(Constants.Graphics.columnWidth)), -ySpacing); //Move accordingly to x spacing if any

                            waitingRightBracket = null;

                            howManyRightBrackets = howManyRightBrackets + 1;
                        }

                        #region UPPER ELEMENT
                        if (currentElement == Constants.Graphics.bracketRightUpper)
                        {
                            //Get appropriate element
                            if (templateGraphicDico.ContainsKey(currentElement) && lastElement != null)
                            {
                                IElement rightBracketElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[currentElement]) as IElement;
                                IElementProperties rightBracketdElProp = rightBracketElement as IElementProperties;

                                string originalElementName = rightBracketdElProp.Name;

                                #region Move to right anchor

                                //Set new anchor
                                Tuple<double, double> rightBracketAnchor = new Tuple<double, double>(anchorPoint.Item1, lastElement.Geometry.Envelope.YMax);
                                MoveItemToAnchorPoint(rightBracketElement, rightBracketAnchor);

                                #endregion

                                //Rename
                                rightBracketdElProp.Name = rightBracketdElProp.Name + currentOrder.ToString();

                                //Add base element
                                currentDoc.ActiveView.GraphicsContainer.AddElement(rightBracketElement, 0);
                                currentDoc.ActiveView.GraphicsContainer.BringToFront(currentGrapSelection.SelectedElements);

                                //Unselect
                                currentGrapSelection.UnselectElement(rightBracketElement);

                                // Don't Keep name

                                //Add to legend list
                                legendElementList.Add(rightBracketElement);

                                //Keep in waiting line, to be move in Y axis
                                waitingRightBracket = rightBracketElement;
                                upRightBracket = rightBracketElement;

                            }
                        }

                        #endregion

                        #region LOWER ELEMENT
                        if (currentElement == Constants.Graphics.bracketRightLower)
                        {
                            //Get appropriate element
                            if (templateGraphicDico.ContainsKey(currentElement) && upRightBracket != null)
                            {
                                #region END BRACKET NOTCH
                                IElement rightEndBracketElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[currentElement]) as IElement;
                                IElementProperties rightEndBracketdElProp = rightEndBracketElement as IElementProperties;

                                string originalElementName = rightEndBracketdElProp.Name;

                                #region Move to right anchor

                                //Set new anchor
                                Tuple<double, double> bottomBracketAnchor = new Tuple<double, double>(anchorPoint.Item1, lastElement.Geometry.Envelope.YMin + rightEndBracketElement.Geometry.Envelope.Height);
                                MoveItemToAnchorPoint(rightEndBracketElement, bottomBracketAnchor);

                                //Move
                                ITransform2D transElement = rightEndBracketElement as ITransform2D;
                                transElement.Move(xSpacing + (currentColumn * GetXSpacing(Constants.Graphics.columnWidth)), 0); //Move accordingly to x spacing if any

                                #endregion

                                //Rename
                                rightEndBracketdElProp.Name = rightEndBracketdElProp.Name + currentOrder.ToString();

                                //Add base element
                                currentDoc.ActiveView.GraphicsContainer.AddElement(rightEndBracketElement, 0);
                                currentDoc.ActiveView.GraphicsContainer.BringToFront(currentGrapSelection.SelectedElements);

                                //Unselect
                                currentGrapSelection.UnselectElement(rightEndBracketElement);

                                //Add to legend list
                                legendElementList.Add(rightEndBracketElement);

                                #endregion

                                #region MIDDLE BRACKET NOTCH
                                IElement rightMiddleBracketElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[Constants.Graphics.bracketRightCenter]) as IElement;
                                IElementProperties rightMiddleBracketdElProp = rightMiddleBracketElement as IElementProperties;

                                string rightMiddleElementName = rightMiddleBracketdElProp.Name;

                                #region Move to right anchor

                                //Set new anchor
                                double rightMiddleBracketY = upRightBracket.Geometry.Envelope.YMin - Math.Abs(((rightEndBracketElement.Geometry.Envelope.YMax - upRightBracket.Geometry.Envelope.YMin) / 2.0));
                                Tuple<double, double> rightMiddleBracketAnchor = new Tuple<double, double>(anchorPoint.Item1, rightMiddleBracketY);
                                MoveItemToAnchorPoint(rightMiddleBracketElement, rightMiddleBracketAnchor);

                                //Move
                                ITransform2D transMiddleElement = rightMiddleBracketElement as ITransform2D;
                                transMiddleElement.Move(GetXSpacing(Constants.Graphics.bracketRightCenter) + (currentColumn * GetXSpacing(Constants.Graphics.columnWidth)), 0); //Move accordingly to x spacing if any

                                #endregion

                                //Rename
                                rightMiddleBracketdElProp.Name = rightMiddleBracketdElProp.Name + currentOrder.ToString();

                                //Add base element
                                currentDoc.ActiveView.GraphicsContainer.AddElement(rightMiddleBracketElement, 0);
                                currentDoc.ActiveView.GraphicsContainer.BringToFront(currentGrapSelection.SelectedElements);

                                //Unselect
                                currentGrapSelection.UnselectElement(rightMiddleBracketElement);

                                //Add to legend list
                                legendElementList.Add(rightMiddleBracketElement);

                                //Keep element to use with bracket annotation
                                waitingCenterRightBracket = rightMiddleBracketElement;

                                #endregion

                                #region BRACKET SPINE 1
                                IElement rightSpine1BracketElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[Constants.Graphics.bracketSpine]) as IElement;
                                IElementProperties rightSpine1BracketdElProp = rightSpine1BracketElement as IElementProperties;

                                string rightSpine1ElementName = rightSpine1BracketdElProp.Name;

                                #region Move to right anchor

                                //Set new anchor
                                IPoint rightStartPointLine = new ESRI.ArcGIS.Geometry.Point();
                                rightStartPointLine.X = upRightBracket.Geometry.Envelope.XMax;
                                rightStartPointLine.Y = upRightBracket.Geometry.Envelope.YMin;

                                IPoint rightEndPointLine = new ESRI.ArcGIS.Geometry.Point();
                                rightEndPointLine.X = rightMiddleBracketElement.Geometry.Envelope.XMin;
                                rightEndPointLine.Y = rightMiddleBracketElement.Geometry.Envelope.YMax;

                                IPolyline rightSpineLine = new ESRI.ArcGIS.Geometry.Polyline() as IPolyline;
                                rightSpineLine.ToPoint = rightEndPointLine;
                                rightSpineLine.FromPoint = rightStartPointLine;

                                rightSpine1BracketElement.Geometry = rightSpineLine;

                                #endregion

                                //Rename
                                rightSpine1BracketdElProp.Name = rightSpine1BracketdElProp.Name + currentOrder.ToString();

                                //Add base element
                                currentDoc.ActiveView.GraphicsContainer.AddElement(rightSpine1BracketElement, 0);
                                currentDoc.ActiveView.GraphicsContainer.BringToFront(currentGrapSelection.SelectedElements);

                                //Unselect
                                currentGrapSelection.UnselectElement(rightSpine1BracketElement);

                                //Add to legend list
                                legendElementList.Add(rightSpine1BracketElement);
                                #endregion

                                #region BRACKET SPINE 2
                                IElement rightSpine2BracketElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[Constants.Graphics.bracketSpine]) as IElement;
                                IElementProperties rightSpine2BracketdElProp = rightSpine2BracketElement as IElementProperties;

                                string rightSpine2ElementName = rightSpine2BracketdElProp.Name;

                                #region Move to right anchor

                                //Set new anchor
                                IPoint rightStartPointLine2 = new ESRI.ArcGIS.Geometry.Point();
                                rightStartPointLine2.X = rightMiddleBracketElement.Geometry.Envelope.XMin;
                                rightStartPointLine2.Y = rightMiddleBracketElement.Geometry.Envelope.YMin;

                                IPoint rightEndPointLine2 = new ESRI.ArcGIS.Geometry.Point();
                                rightEndPointLine2.X = rightEndBracketElement.Geometry.Envelope.XMax;
                                rightEndPointLine2.Y = rightEndBracketElement.Geometry.Envelope.YMax;

                                IPolyline rightspineLine2 = new ESRI.ArcGIS.Geometry.Polyline() as IPolyline;
                                rightspineLine2.ToPoint = rightEndPointLine2;
                                rightspineLine2.FromPoint = rightStartPointLine2;

                                rightSpine2BracketElement.Geometry = rightspineLine2;

                                #endregion

                                //Rename
                                rightSpine2BracketdElProp.Name = rightSpine2BracketdElProp.Name + currentOrder.ToString();

                                //Add base element
                                currentDoc.ActiveView.GraphicsContainer.AddElement(rightSpine2BracketElement, 0);
                                currentDoc.ActiveView.GraphicsContainer.BringToFront(currentGrapSelection.SelectedElements);

                                //Unselect
                                currentGrapSelection.UnselectElement(rightSpine2BracketElement);

                                //Add to legend list
                                legendElementList.Add(rightSpine2BracketElement);
                                #endregion
                            }
                        }

                        #endregion

                        #region UNIT BOX ELEMENT
                        //Process left bracket annotation waiting to have the right anchor point
                        if (waitingCenterRightBracket != null && bracketMapUnit.Item3 != null)
                        {
                            //Set new anchor for unit box
                            //double xUnitBoxBracketAnchor = anchorPoint.Item1 + GetXSpacing(Constants.Graphics.bracketRightGap) + waitingCenterRightBracket.Geometry.Envelope.Width + GetXSpacing(Constants.Graphics.unitBoxBracket) -10;
                            double xUnitBoxBracketAnchor = waitingCenterRightBracket.Geometry.Envelope.XMax + GetXSpacing(Constants.Graphics.unitBoxBracket);
                            double yUnitBoxBracketAnchor = waitingCenterRightBracket.Geometry.Envelope.YMin + (waitingCenterRightBracket.Geometry.Envelope.YMax - waitingCenterRightBracket.Geometry.Envelope.YMin) / 2.0 + bracketMapUnit.Item1.Geometry.Envelope.Height / 2.0;
                            Tuple<double, double> mapUnitBracketAnchor = new Tuple<double, double>(xUnitBoxBracketAnchor, yUnitBoxBracketAnchor);
                            //MoveItemToAnchorPoint(bracketMapUnit.Item1, mapUnitBracketAnchor);
                            SetRectangularPolygonFromAnchorType(bracketMapUnit.Item1, mapUnitBracketAnchor);
                            //Set new anchor for unit box dem if any
                            if (bracketMapUnit.Item4 != null)
                            {
                                MoveItemToAnchorPoint(bracketMapUnit.Item4, mapUnitBracketAnchor);
                            }

                            //Set new anchor for label
                            double xLabelBracketAnchor = mapUnitBracketAnchor.Item1 + bracketMapUnit.Item1.Geometry.Envelope.Width / 2.0;
                            double yLabelBracketAnchor = mapUnitBracketAnchor.Item2 - bracketMapUnit.Item1.Geometry.Envelope.Height / 2.0;
                            Tuple<double, double> labelBracketAnchor = new Tuple<double, double>(xLabelBracketAnchor, yLabelBracketAnchor);
                            MoveItemToAnchorPoint(bracketMapUnit.Item2, labelBracketAnchor);

                            //Set new anchor for description
                            double xDescBracketAnchor = bracketMapUnit.Item1.Geometry.Envelope.XMax + GetXSpacing(Constants.Graphics.elementDescriptionGapWidth);
                            Tuple<double, double> descBracketAnchor = new Tuple<double, double>(xDescBracketAnchor, yUnitBoxBracketAnchor);
                            AddDescriptionFromElement(bracketMapUnit.Item3, bracketMapUnit.Item1, currentDoc, mapUnitBracketAnchor, Constants.Graphics.unitBox);

                            waitingCenterRightBracket = null;
                            bracketMapUnit = new Tuple<IElement, IElement, IElement, IElement>(null, null, null, null);
                        }
                        #endregion

                        #endregion

                        #region NOTES
                        if (currentElement == Constants.Graphics.topNote || currentElement == Constants.Graphics.note)
                        {

                            //Get appropriate element
                            if (templateGraphicDico.ContainsKey(currentElement))
                            {
                                IElement noteElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[currentElement]) as IElement;
                                IElementProperties noteElProp = noteElement as IElementProperties;

                                string originalElementName = noteElProp.Name;

                                //Add Description
                                ITextElement tElement = noteElement as ITextElement;
                                if (currentDescription == null || currentDescription == string.Empty || currentDescription == " ")
                                {
                                    currentDescription = Constants.TextConfiguration.missingText;
                                    tElement.Symbol = Services.Symbols.GetMissingTextSymbol(tElement.Symbol);
                                }
                                else
                                {

                                    //Add header if needed
                                    if (currentHeading != null && currentHeading != string.Empty && currentHeading != " ")
                                    {
                                        currentDescription = Constants.TextConfiguration.tagBold + currentHeading + Constants.TextConfiguration.endTagBold + " " + currentDescription;
                                    }

                                    tElement.Text = currentDescription;
                                }



                                //Set width and height
                                double wantedTextHeight = GetTextHeight(currentDescription, descriptionWidth);
                                IEnvelope env = noteElement.Geometry.Envelope;
                                env.Width = noteElement.Geometry.Envelope.Width; //Set width
                                env.Height = wantedTextHeight;

                                IPolygon pol = new Polygon() as IPolygon; //Create new polygon from wanted envelope
                                ISegmentCollection polSegment = pol as ISegmentCollection;
                                polSegment.SetRectangle(env);

                                noteElement.Geometry = pol as IGeometry;

                                #region Move to right anchor

                                //Set new anchor
                                anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing);
                                SetRectangularPolygonFromAnchorType(noteElement, anchorPoint);

                                //Move
                                ITransform2D transNoteElement = noteElement as ITransform2D;
                                transNoteElement.Move(xSpacing, 0); //Move accordingly to x spacing if any

                                #endregion

                                //Rename
                                noteElProp.Name = noteElProp.Name + currentOrder.ToString();

                                //Add base element
                                currentDoc.ActiveView.GraphicsContainer.AddElement(noteElement, 0);
                                currentDoc.ActiveView.GraphicsContainer.SendToBack(currentGrapSelection.SelectedElements);

                                //Unselect
                                currentGrapSelection.UnselectElement(noteElement);

                                //Keep name
                                double noteHeight = noteElement.Geometry.Envelope.Height;
                                double noteYMin = noteElement.Geometry.Envelope.YMin;
                                double noteYMax = noteElement.Geometry.Envelope.YMax;
                                lastElement = noteElement;
                                lastElementType = originalElementName;

                                //Add to legend list
                                legendElementList.Add(lastElement);
                            }

                        }


                        #endregion

                        #region AUTOMATIC COLUMNS

                        //if(this.checkBox_autoCalculateColumns.Checked)
                        //{
                        //    //Track column change with auto-calculate
                        //    if (legendYLowerBound != 0.0)
                        //    {
                        //        if ((anchorPoint.Item2 - ySpacing - currentElementObject.Geometry.Envelope.Height) < legendYLowerBound)
                        //        {
                        //            currentColumn++;
                        //        }
                        //    }

                        //}




                        #endregion

                        //Release each row, else a lock can occur.
                        Services.ObjectManagement.ReleaseObject(legendRow);
                    }

                    #endregion

                    #region Group all items 

                    ////TODO: Finish Issue #2 bc69669a latest commit about it.
                    ////if (!isCGMTemplateMXD)
                    ////{
                    //IGroupElement3 groupedLegend = GetGroupLegendElement("Test");
                    //legendElementList.Reverse();

                    //foreach (IElement tElements in legendElementList)
                    //{
                    //    currentDoc.ActiveView.GraphicsContainer.MoveElementToGroup(tElements, groupedLegend as IGroupElement);
                    //}


                    ////Add group legend
                    //currentDoc.ActiveView.GraphicsContainer.AddElement(groupedLegend as IElement, 0);


                    ////    #region Move whole legend if left bracket was found
                    ////    //TODO fix why when moving group element inner parts don't move at the right place
                    ////    if (upLeftBracket != null)
                    ////    {
                    ////        //ITransform2D moveGroupedLegend = groupedLegend as ITransform2D;
                    ////        //moveGroupedLegend.Move(9.9206, 0);
                    ////    }
                    ////    #endregion

                    ////}

                    #endregion

                    //Delete originalCGMLegend if exist
                    if (originalCGMLegend != null)
                    {
                        currentDoc.ActiveView.GraphicsContainer.DeleteElement(originalCGMLegend);
                    }

                    //TODO commit this part for issue #2 commit bc69669a
                    //Reset units to be like it was
                    if (originalUnits != esriUnits.esriMillimeters)
                    {
                        SetDocumentUnits(currentDoc, originalUnits);
                    }

                    currentDoc.ActiveView.Refresh();

                    //Release all objects so user can keep on editing the original legend file
                    Services.ObjectManagement.ReleaseObject(legendCursor);
                    Services.ObjectManagement.ReleaseObject(legendTable);



                    CloseForm();


                }
                else
                {
                    MessageBox.Show("Missing style file."); //TODO change this message for localized one and better text.
                }

            }
            else
            {
                MessageBox.Show("Fill in all inputs"); //TODO change text and make it localized.
            }

        }

        /// <summary>
        /// Will close current form.
        /// </summary>
        private void CloseForm()
        {
            this.Close();
        }

        /// <summary>
        /// Will take from template mxd a set of graphics to use in the legend
        /// </summary>
        public void GetTemplateGraphicList()
        {
            //Validate if template mxd is available, else send copy resource to local folder
            string templateMXDPath = ValidateTemplateMXDExistance();

            //Open template mxd to retrieve graphics out of it
            IMapDocument mapDoc = new MapDocument();
            mapDoc.Open(templateMXDPath);

            //Get to layout view
            Services.MXD mxdService = new Services.MXD();
            mxdService.ActivateLayoutView(mapDoc.ActiveView);

            //Get list of graphics
            IGraphicsContainer graphiclist = mapDoc.ActiveView.GraphicsContainer;

            //Filter graphics
            IElement nextGraphic = graphiclist.Next();
            while (nextGraphic != null)
            {
                IElementProperties nextGProperties = nextGraphic as IElementProperties;
                if (nextGProperties.Name != "New Data Frame")
                {
                    ///For testing purposes
                    //currentDoc.ActiveView.GraphicsContainer.AddElement(nextGraphic, 0);
                    templateGraphicDico[nextGProperties.Name] = nextGraphic;
                    //templateGraphics.Add(nextGraphic);
                }

                nextGraphic = graphiclist.Next();
            }

            //Close template mxd
            mapDoc.Close();
        }

        /// <summary>
        /// From table property, will fill all comboboxes with field values.
        /// </summary>
        private void FillFieldsComboboxes()
        {
            //Validate list existance
            if (dataFieldList == null)
            {
                InitFieldList();
            }

            if (dataFieldList != null)
            {
                //Clean up 
                this.comboBox_ColumnField.Items.Clear();
                this.comboBox_DescriptionField.Items.Clear();
                this.comboBox_ElementField.Items.Clear();
                this.comboBox_HeadingField.Items.Clear();
                this.comboBox_Label1Field.Items.Clear();
                this.comboBox_Label1StyleField.Items.Clear();
                this.comboBox_Label2Field.Items.Clear();
                this.comboBox_Label2StyleField.Items.Clear();
                this.comboBox_orderField.Items.Clear();
                this.comboBox_Style1Field.Items.Clear();
                this.comboBox_Style2Field.Items.Clear();

                //Fill boxes
                foreach (string fieldNames in dataFieldList)
                {
                    this.comboBox_ColumnField.Items.Add(fieldNames);
                    this.comboBox_DescriptionField.Items.Add(fieldNames);
                    this.comboBox_ElementField.Items.Add(fieldNames);
                    this.comboBox_HeadingField.Items.Add(fieldNames);
                    this.comboBox_Label1Field.Items.Add(fieldNames);
                    this.comboBox_Label1StyleField.Items.Add(fieldNames);
                    this.comboBox_Label2Field.Items.Add(fieldNames);
                    this.comboBox_Label2StyleField.Items.Add(fieldNames);
                    this.comboBox_orderField.Items.Add(fieldNames);
                    this.comboBox_Style1Field.Items.Add(fieldNames);
                    this.comboBox_Style2Field.Items.Add(fieldNames);
                }

                //Select proper field if in list
                SelectFields();
            }

        }

        /// <summary>
        /// Will automatically select field if they match default schema for legend table
        /// </summary>
        private void SelectFields()
        {
            if (this.comboBox_ColumnField.Items.Contains(fieldColumn))
            {
                int columnIndex = this.comboBox_ColumnField.Items.IndexOf(fieldColumn);
                this.comboBox_ColumnField.SelectedIndex = columnIndex;
            }
            if (this.comboBox_DescriptionField.Items.Contains(fieldDescription))
            {
                int descriptionIndex = this.comboBox_DescriptionField.Items.IndexOf(fieldDescription);
                this.comboBox_DescriptionField.SelectedIndex = descriptionIndex;
            }
            if (this.comboBox_ElementField.Items.Contains(fieldElement))
            {
                int elementIndex = this.comboBox_ElementField.Items.IndexOf(fieldElement);
                this.comboBox_ElementField.SelectedIndex = elementIndex;
            }
            if (this.comboBox_HeadingField.Items.Contains(fieldHeading))
            {
                int headingIndex = this.comboBox_HeadingField.Items.IndexOf(fieldHeading);
                this.comboBox_HeadingField.SelectedIndex = headingIndex;
            }
            if (this.comboBox_Label1Field.Items.Contains(fieldLabel1))
            {
                int label1Index = this.comboBox_Label1Field.Items.IndexOf(fieldLabel1);
                this.comboBox_Label1Field.SelectedIndex = label1Index;
            }
            if (this.comboBox_Label1StyleField.Items.Contains(fieldLabel1Style))
            {
                int label1StyleIndex = this.comboBox_Label1StyleField.Items.IndexOf(fieldLabel1Style);
                this.comboBox_Label1StyleField.SelectedIndex = label1StyleIndex;
            }
            if (this.comboBox_Label2Field.Items.Contains(fieldLabel2))
            {
                int label2Index = this.comboBox_Label2Field.Items.IndexOf(fieldLabel2);
                this.comboBox_Label2Field.SelectedIndex = label2Index;
            }
            if (this.comboBox_Label2StyleField.Items.Contains(fieldLabel2Style))
            {
                int label2StyleIndex = this.comboBox_Label2StyleField.Items.IndexOf(fieldLabel2Style);
                this.comboBox_Label2StyleField.SelectedIndex = label2StyleIndex;
            }
            if (this.comboBox_orderField.Items.Contains(fieldOrder))
            {
                int orderIndex = this.comboBox_orderField.Items.IndexOf(fieldOrder);
                this.comboBox_orderField.SelectedIndex = orderIndex;
            }
            if (this.comboBox_Style1Field.Items.Contains(fieldStyle1))
            {
                int style1Index = this.comboBox_Style1Field.Items.IndexOf(fieldStyle1);
                this.comboBox_Style1Field.SelectedIndex = style1Index;
            }
            if (this.comboBox_Style2Field.Items.Contains(fieldStyle2))
            {
                int style2Index = this.comboBox_Style2Field.Items.IndexOf(fieldStyle2);
                this.comboBox_Style2Field.SelectedIndex = style2Index;
            }
        }

        /// <summary>
        /// Will fill in the appropriate symbol dictionary from needed style
        /// </summary>
        /// <param name="symbolClassName">Fill Symbols, Marker Symbols, Line Symbols</param>
        private void GetSymbols(string symbolClassName)
        {

            //Iterate through all fill symbols
            IStyleGalleryItem styleItem = null;
            IEnumStyleGalleryItem enumStyle = gscStyle.Items[symbolClassName, gscStylePath, string.Empty];

            while ((styleItem = enumStyle.Next()) != null)
            {
                if (symbolClassName == Dictionaries.Constants.Styles.styleFillClass)
                {
                    if (fillSymbolDico == null)
                    {
                        fillSymbolDico = new Dictionary<string, object>();
                    }

                    fillSymbolDico[styleItem.Name] = styleItem.Item;
                }
                else if (symbolClassName == Dictionaries.Constants.Styles.styleLineClass)
                {
                    if (lineSymbolDico == null)
                    {
                        lineSymbolDico = new Dictionary<string, object>();
                    }

                    lineSymbolDico[styleItem.Name] = styleItem.Item;
                }
                else if (symbolClassName == Dictionaries.Constants.Styles.styleMarkerClass)
                {

                    if (markerSymbolDico == null)
                    {
                        markerSymbolDico = new Dictionary<string, object>();
                    }

                    markerSymbolDico[styleItem.Name] = styleItem.Item;
                }
                else if (symbolClassName == Dictionaries.Constants.Styles.styleTextClass)
                {
                    if (textSymbolDico == null)
                    {
                        textSymbolDico = new Dictionary<string, object>();
                    }

                    textSymbolDico[styleItem.Name] = styleItem.Item;
                }
                //else if (symbolClassName == Dictionaries.Constants.Styles.styleRepresentationMarkerClass)
                //{
                //    if (repMarkerSymbolDico == null)
                //    {
                //        repMarkerSymbolDico = new Dictionary<string, object>();
                //    }

                //    repMarkerSymbolDico[styleItem.Name] = styleItem.Item;
                //}
            }
        }

        /// <summary>
        /// Will deserialize a json file to get all key component for y spacings
        /// </summary>
        private void BuildYSpacingsDictionary()
        {
            if (ySpacings.Count() == 0)
            {
                ySpacings = JsonConvert.DeserializeObject<Dictionary<string, Dictionary<string, string>>>(File.ReadAllText(jsonYSpacingFilePath));
            }

        }

        /// <summary>
        /// Will deserialize a json file to get all key component for x spacings
        /// </summary>
        private void BuildXSpacingsDictionary()
        {
            if (xSpacings.Count() == 0)
            {
                xSpacings = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(jsonXSpacingFilePath));
            }

        }

        /// <summary>
        /// Will deserialize a json file to get all other key components required by the tool, especially for styles and fonts
        /// </summary>
        private void BuildOtherComponentsDictionary()
        {
            if (otherComponents.Count() == 0)
            {
                otherComponents = JsonConvert.DeserializeObject<Dictionary<string, string>>(File.ReadAllText(jsonOtherFilePath));
            }

        }

        /// <summary>
        /// Will return a y spacing based on from and to element names
        /// </summary>
        /// <returns></returns>
        private double GetYSpacing(IElement fromElement, string fromElementType, string toElementName, double lastYSpacing)
        {
            //Variable
            double y = 0.0;
            IElementProperties fromProperties = fromElement as IElementProperties;

            //Get x-y spacing
            if (ySpacings.ContainsKey(fromElementType))
            {
                if (ySpacings[fromElementType].ContainsKey(toElementName))
                {
                    if (ySpacings[fromElementType][toElementName] != null && ySpacings[fromElementType][toElementName] != string.Empty)
                    {
                        try
                        {
                            y = Convert.ToDouble(ySpacings[fromElementType][toElementName], CultureInfo.InvariantCulture);
                        }
                        catch (Exception)
                        {
                            if (ySpacings[fromElementType][toElementName].Contains(Constants.Graphics.anchorLowerLeft))
                            {
                                y = Math.Abs((fromElement.Geometry.Envelope.YMin - lastYSpacing)) + Convert.ToDouble(ySpacings[fromElementType][toElementName].Split(' ')[0], CultureInfo.InvariantCulture);
                            }
                        }

                        //Special case for heading3
                        if (fromElementType == Constants.Graphics.heading3 || fromElementType == Constants.Graphics.topNote || fromElementType == Constants.Graphics.note)
                        {
                            y = fromElement.Geometry.Envelope.Height + Convert.ToDouble(ySpacings[fromElementType][toElementName].Split(' ')[0], CultureInfo.InvariantCulture);
                        }

                    }
                }
            }

            return y;
        }

        /// <summary>
        /// Will return a x spacing based on element names
        /// </summary>
        /// <returns></returns>
        private double GetXSpacing(string toElementName)
        {
            //Variable
            double x = 0.0;

            if (xSpacings.ContainsKey(toElementName))
            {

                if (xSpacings[toElementName] != null && xSpacings[toElementName].ToString() != string.Empty)
                {
                    try
                    {
                        x = Convert.ToDouble(xSpacings[toElementName], CultureInfo.InvariantCulture);
                    }
                    catch (Exception)
                    {

                    }

                }

            }

            return x;
        }

        /// <summary>
        /// Will create a default group element that will act as the final legend element as a whole.
        /// </summary>
        /// <param name="groupAnchor"></param>
        /// <returns></returns>
        private IGroupElement3 GetGroupLegendElement(string elementName = Constants.Graphics.cgmLegendElement)
        {
            //Variables
            IGroupElement3 groupedLegend = new GroupElement() as IGroupElement3;

            //Set name
            IElementProperties elementGL = groupedLegend as IElementProperties;
            elementGL.Name = elementName;
            elementGL.AutoTransform = true;

            //Set anchor point
            IElementProperties3 elemProp3 = groupedLegend as IElementProperties3;
            elemProp3.AnchorPoint = esriAnchorPointEnum.esriTopLeftCorner;
            elemProp3.AutoTransform = true;

            //Set fixed aspect ratio, else moving a grouped element can see it's inner parts going elsewhere in the map...
            IBoundsProperties elementBoundsProp = groupedLegend as IBoundsProperties;
            elementBoundsProp.FixedAspectRatio = false;


            return groupedLegend;
        }


        /// <summary>
        /// Will fill the combobox control with all layers from table of content
        /// </summary>
        private void FillTableViewCombobox()
        {
            //Unset list before change
            this.comboBox_SelectTable.DataSource = null;

            //Get a list of all point layers inside table of content
            Services.Layers layerService = new Services.Layers();
            UID layerUID = new UID();
            layerUID.Value = "{34C20002-4D3C-11D0-92D8-00805F7C28B0}";
            List<IStandaloneTable> tocTables = layerService.GetListOfStandaloneTables((IMxDocument)ArcMap.Application.Document);

            //Init datasource of combobox
            _cboxlayers = new List<CboxTables>();

            //Fill in the control
            foreach (IStandaloneTable l in tocTables)
            {
                _cboxlayers.Add(new CboxTables { cboxDataName = l.Name, cboxSTLTable = l});
            }

            //Init defaults
            this.comboBox_SelectTable.SelectedIndex = -1;
            this.comboBox_SelectTable.DataSource = _cboxlayers;
            this.comboBox_SelectTable.DisplayMember = "cboxDataName";
            this.comboBox_SelectTable.ValueMember = "cboxSTLTable";

        }

        #endregion

        #region ADD GRAPHIC

        /// <summary>
        /// Will add a new text element at the center a of map unit box
        /// </summary>
        /// <param name="inText"></param>
        /// <param name="parentElement"></param>
        /// <param name="inDocument"></param>
        private IElement AddLabelInUnitBox(string inText, IElement parentElement, IMxDocument inDocument, Tuple<double, double> inAnchor, Constants.Graphics.UnitBoxType unitBoxType, string inStyle = "")
        {
            //TODO get symbol from style file?? Ask Dave

            //Get appropriate element
            IElement unitBoxLabelElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[Constants.Graphics.unitLabel]) as IElement;

            //Create new text graphic with default style
            ITextElement tElement = unitBoxLabelElement as ITextElement;

            //Manage incoming style if needed
            if (inStyle!="" && inStyle != null && textSymbolDico.ContainsKey(inStyle))
            {
                ISimpleTextSymbol inStyleSymbol = textSymbolDico[inStyle] as ISimpleTextSymbol;
                ISimpleTextSymbol currentStyleSymbol = tElement.Symbol as ISimpleTextSymbol;
                currentStyleSymbol.Font = inStyleSymbol.Font;
                currentStyleSymbol.Color = inStyleSymbol.Color;
                currentStyleSymbol.Size = Constants.TextConfiguration.defaultUnitBoxLabelFontSize; //Force size else incoming style might be too big.
                currentStyleSymbol.VerticalAlignment = esriTextVerticalAlignment.esriTVACenter; //Force vertical center for text else incoming style might be set to else where.
                tElement.Symbol = Services.ObjectManagement.CopyInputObject(currentStyleSymbol) as ISimpleTextSymbol;
                
            }

            //Mange too long text (mainly to fix when used in UNIT_SPLIT boxes).
            //Conditions on style to prevent trigger on special fonts
            if (inText.Length >= 6 && inStyle == "")
            {
                ISimpleTextSymbol currentStyleSymbol = tElement.Symbol as ISimpleTextSymbol;
                currentStyleSymbol.Size = Constants.TextConfiguration.tooLongLabelUnitBoxLabelFontSize; //Force size else incoming style might be too big.

                //Manage placement
                if (unitBoxType == Constants.Graphics.UnitBoxType.split2)
                {
                    //Shift down a bit for right part.
                    currentStyleSymbol.VerticalAlignment = esriTextVerticalAlignment.esriTVABaseline; 
                }
                else if (unitBoxType == Constants.Graphics.UnitBoxType.split1)
                {
                    //Shift up a bit for left part
                    currentStyleSymbol.VerticalAlignment = esriTextVerticalAlignment.esriTVATop;
                }
                
                tElement.Symbol = Services.ObjectManagement.CopyInputObject(currentStyleSymbol) as ISimpleTextSymbol;
            }

            //Manage missing
            if (inText == null || inText == string.Empty || inText == " ")
            {
                inText = Constants.TextConfiguration.missingText;
                tElement.Symbol = Services.Symbols.GetMissingTextSymbol(tElement.Symbol);
            }
            tElement.Text = inText;

            //Rename
            IElementProperties labelElementProperties = unitBoxLabelElement as IElementProperties;
            labelElementProperties.Name = inText;
            labelElementProperties.AutoTransform = false;

            //Get width and height of parent
            IGeometry parentGeom = parentElement.Geometry;
            IEnvelope parentEnvelope = parentGeom.Envelope;
            double parentHeight = parentEnvelope.Height;
            double parentWidth = parentEnvelope.Width;

            //Set new anchor point
            SetRectangularPolygonFromAnchorType(unitBoxLabelElement, inAnchor);

            //Move
            ITransform2D transElement = unitBoxLabelElement as ITransform2D;
            if (unitBoxType == Constants.Graphics.UnitBoxType.normal)
            {
                transElement.Move(parentWidth / 2.0, (-parentHeight / 2.0)); //Move accordingly to anchor point which is center center
            }
            else if (unitBoxType == Constants.Graphics.UnitBoxType.split1)
            {
                //https://www.mathopenref.com/coordcentroid.html
                double centerX = (3 * inAnchor.Item1 + parentWidth) / 3.0;
                double centerY = (3 * inAnchor.Item2 - parentHeight) / 3.0;
                transElement.Move(centerX - inAnchor.Item1, -(Math.Abs(centerY - inAnchor.Item2))); //Move accordingly to anchor point which is center center
            }
            else if (unitBoxType == Constants.Graphics.UnitBoxType.split2 || unitBoxType == Constants.Graphics.UnitBoxType.line || unitBoxType == Constants.Graphics.UnitBoxType.child_line)
            {
                //https://www.mathopenref.com/coordcentroid.html
                double centerX = (3 * inAnchor.Item1 + 2 * parentWidth) / 3.0;
                double centerY = (3 * inAnchor.Item2 - 2 * parentHeight) / 3.0;
                transElement.Move(centerX - inAnchor.Item1, -(Math.Abs(centerY - inAnchor.Item2))); //Move accordingly to anchor point which is center center
            }

            //Add base element
            IPageLayout l = inDocument.ActiveView as IPageLayout;
            IGraphicsContainerSelect gcs = l as IGraphicsContainerSelect;
            inDocument.ActiveView.GraphicsContainer.AddElement(unitBoxLabelElement as IElement, 0);
            gcs.UnselectAllElements();
            gcs.SelectElement(unitBoxLabelElement);
            
            inDocument.ActiveView.GraphicsContainer.BringToFront(gcs.SelectedElements);

            //Unselect
            gcs.UnselectElement(unitBoxLabelElement);

            //Add to legend list
            //legendElementList.Add(unitBoxLabelElement);

            return unitBoxLabelElement;
        }

        /// <summary>
        /// Will add anew text element for description in column description part
        /// </summary>
        /// <param name="inDescription"></param>
        /// <param name="parentElem"></param>
        /// <param name="inDocument"></param>
        /// <param name="inAnchor"></param>
        /// <returns> Description height for validation purposes</returns>
        private IElement AddDescription(string inDescription, IElement parentElem, IMxDocument inDocument, Tuple<double, double> inAnchor, string parentElemType, bool isLineOrPoint = false)
        {
            //Get appropriate element
            IElement descriptionElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[Constants.Graphics.description]) as IElement;

            //Get different size description
            if (parentElemType == Constants.Graphics.unitindent1)
            {
                descriptionElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[Constants.Graphics.description_indent]) as IElement;
            }
            else if (parentElemType == Constants.Graphics.unitindent2)
            {
                descriptionElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[Constants.Graphics.description_indent2]) as IElement;
            }

            //If description is meant for a group 5 heading, then modify style
            if (heading5Text.Count >= 1)
            {
                descriptionElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[Constants.Graphics.heading5Description]) as IElement;
                double indentation = GetXSpacing(Constants.Graphics.heading5Description) - GetXSpacing(Constants.Graphics.description);
                inAnchor = new Tuple<double, double>(inAnchor.Item1 + indentation, inAnchor.Item2);
            }

            //Create new text graphic with default style
            ITextElement dtElement = descriptionElement as ITextElement;

            //Manage missing
            if (inDescription == null || inDescription == string.Empty || inDescription == " ")
            {
                inDescription = Constants.TextConfiguration.missingText;
                dtElement.Symbol = Services.Symbols.GetMissingTextSymbol(dtElement.Symbol);
            }

            dtElement.Text = inDescription;

            legendElementList.Add(descriptionElement as IElement);
            currentDoc.ActiveView.GraphicsContainer.AddElement(descriptionElement as IElement, 0);

            return AddDescriptionFromElement(descriptionElement, parentElem, inDocument, inAnchor, parentElemType, isLineOrPoint);
        }

        /// <summary>
        /// Will add an existing text element for description at the right place beside it's parent.
        /// </summary>
        /// <param name="descriptionElement"></param>
        /// <param name="parentElem"></param>
        /// <param name="inDocument"></param>
        /// <param name="inAnchor"></param>
        /// <param name="parentElemType"></param>
        /// <returns></returns>
        private IElement AddDescriptionFromElement(IElement descriptionElement, IElement parentElem, IMxDocument inDocument, Tuple<double, double> inAnchor, string parentElemType, bool isLineOrPoint = false)
        {

            //Create new text graphic with default style
            ITextElement dtElement = descriptionElement as ITextElement;

            //Variables
            double wantedTextHeight = GetTextHeight(dtElement.Text, descriptionWidth);
            IElementProperties3 parentProperties = parentElem as IElementProperties3;

            IElementProperties3 dtElementPro = dtElement as IElementProperties3;
            dtElementPro.Name = parentProperties.Name + "_description";

            //Min height setting
            double smallDescHeight = smallDescriptionHeight;
            if (isLineOrPoint)
            {
                smallDescHeight = smallDescriptionHeightLine;
            }

            //Get width and height of parent
            IGeometry parentGeom = parentElem.Geometry;
            double parentHeight = 1.0;
            double parentWidth = elementWidth;

            try
            {
                //In case element is a unit box, try to get height
                IEnvelope parentEnv = parentGeom.Envelope;
                parentHeight = parentEnv.Height;
            }
            catch (Exception)
            {

            }

            //Parent height setting based on anchor position
            //double parentHeightCalculated = parentHeight;
            //if (parentProperties.AnchorPoint == esriAnchorPointEnum.esriCenterPoint || parentProperties.AnchorPoint == esriAnchorPointEnum.esriLeftMidPoint || parentProperties.AnchorPoint == esriAnchorPointEnum.esriRightMidPoint)
            //{
            //    parentHeightCalculated = parentHeightCalculated / 2.0;
            //}
            ////else if (parentProperties.AnchorPoint == esriAnchorPointEnum.esriBottomLeftCorner || parentProperties.AnchorPoint == esriAnchorPointEnum.esriBottomMidPoint || parentProperties.AnchorPoint == esriAnchorPointEnum.esriBottomRightCorner)
            ////{
            ////    parentHeightCalculated
            ////}

            //Set width and height and manage group description for heading 5
            IEnvelope env = descriptionElement.Geometry.Envelope;
            if (heading5Text.Count >= 1)
            {
                env.Width = groupDescriptionWidth;//Set width
            }
            //else
            //{
            //    env.Width = descriptionWidth; //Set width
            //}

            env.Height = wantedTextHeight;
            IPolygon pol = new Polygon() as IPolygon; //Create new polygon from wanted envelope
            ISegmentCollection polSegment = pol as ISegmentCollection;
            polSegment.SetRectangle(env);

            descriptionElement.Geometry = pol as IGeometry;

            //Rename
            IElementProperties parentElementProperties = parentElem as IElementProperties;
            IElementProperties descriptionElementProperties = descriptionElement as IElementProperties;
            descriptionElementProperties.Name = parentElementProperties.Name + "_" + Constants.Graphics.description;

            //Set anchor
            SetRectangularPolygonFromAnchorType(descriptionElement, inAnchor);

            //Move based on different length of description
            ITransform2D transElement = descriptionElement as ITransform2D;

            if (wantedTextHeight <= smallDescHeight)
            {
                //When description height is less then align its center on parent center
                if (wantedTextHeight <= parentHeight || parentHeight <= 1.0)
                {
                    if (!IsElementAllNonFlatLines(parentElem) && !parentProperties.Name.Contains(Constants.Graphics.blob))
                    {
                        transElement.Move(elementDescriptGapWidth + parentWidth, -(parentHeight / 2.0 - wantedTextHeight / 2.0)); //Anchor is upper left but needs to be centered on unit box.
                    }
                    else
                    {
                        //Special case for wave and blob since it's a line with anchor in center/center but behaves like bottom center...?
                        transElement.Move(elementDescriptGapWidth + parentWidth, (parentHeight / 2.0) - (parentHeight - wantedTextHeight) / 2.0); //Anchor is upper left but needs to be centered on unit box
                    }

                }
                else
                {
                    transElement.Move(elementDescriptGapWidth + parentWidth, wantedTextHeight / 2.0);
                }

            }
            else
            {
                if (isLineOrPoint)
                {
                    if (parentHeight <= 1.0)
                    {
                        transElement.Move(elementDescriptGapWidth + parentWidth, (parentHeight /2.0) + Constants.YSpacings.lineHeight0DescriptionHeightAdjustement);
                    }
                    else
                    {
                        transElement.Move(elementDescriptGapWidth + parentWidth, (parentHeight / 2.0) + Constants.YSpacings.lineDescriptionHeightAdjustement);
                    }

                }
                else
                {
                    transElement.Move(elementDescriptGapWidth + parentWidth, 0); // Anchor is upper left and needs to be horizontally aligned with it
                }
                

            }

            //Add to legend list
            if (!legendElementList.Contains(descriptionElement))
            {
                legendElementList.Add(descriptionElement);
            }

            //Add base element
            IPageLayout l = inDocument.ActiveView as IPageLayout;
            IGraphicsContainerSelect gcs = l as IGraphicsContainerSelect;
            gcs.UnselectElement(descriptionElement);


            return descriptionElement;
        }

        /// <summary>
        /// Will add a text label around a marker.
        /// </summary>
        /// <param name="inLabelText">The text that will be added to the label</param>
        /// <param name="parentElement">The element onto which a label will be added around it</param>
        /// <param name="inDocument">The document in which the label will be added</param>
        /// <param name="inAnchor">The anchor of the parent</param>
        /// <param name="parentElemType">The parent original name (type) to parse where to put the label (POINT_CC_45 vs POINT_LC_45)</param>
        /// <returns></returns>
        private IElement AddLabelToMarker(string inLabelText, IElement parentElement, IMxDocument inDocument, Tuple<double, double> inAnchor, Constants.Styles.MarkerLabelPositioning wantedPosition, string inLabelStyle = "", Constants.Styles.MarkerLabelPositioning parentPosition = Constants.Styles.MarkerLabelPositioning.FromCenterToUpperLeft)
        {
            //Variables
            string inElementType = string.Empty;

            //Get appropriate element (measurement or generation)
            IElement markerLabelElement = null;
            int measurementValue = -1;
            if (Int32.TryParse(inLabelText, out measurementValue))
            {
                inElementType = Constants.Graphics.measurementLabel;
                markerLabelElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[inElementType]) as IElement;
            }
            else
            {
                inElementType = Constants.Graphics.generationLabel;
                markerLabelElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[inElementType]) as IElement;
            }

            if (inLabelStyle != string.Empty && textSymbolDico.ContainsKey(inLabelStyle))
            {
                inElementType = Constants.Graphics.defaultpointLabel;
                markerLabelElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[inElementType]) as IElement;

                //Get wanted symbol
                ISimpleTextSymbol inStyleSymbol = textSymbolDico[inLabelStyle] as ISimpleTextSymbol;

                //Get inside marker label symbol to change it with found label
                ITextElement markerLabelTextElement = markerLabelElement as ITextElement;
                ISimpleTextSymbol currentStyleSymbol = markerLabelTextElement.Symbol as ISimpleTextSymbol;
                currentStyleSymbol.Font = inStyleSymbol.Font;
                currentStyleSymbol.Color = inStyleSymbol.Color;
                currentStyleSymbol.Size = currentStyleSymbol.Size; //Force size else incoming style might be too big.
                currentStyleSymbol.VerticalAlignment = currentStyleSymbol.VerticalAlignment; //Force vertical center for text else incoming style might be set to else where.
                markerLabelTextElement.Symbol = Services.ObjectManagement.CopyInputObject(currentStyleSymbol) as ISimpleTextSymbol;

            }

            //Create new text graphic with default style
            ITextElement lElement = markerLabelElement as ITextElement;

            //Manage missing
            if (inLabelText == null || inLabelText == string.Empty || inLabelText == " ")
            {
                inLabelText = Constants.TextConfiguration.missingText;
                lElement.Symbol = Services.Symbols.GetMissingTextSymbol(lElement.Symbol);
            }
            lElement.Text = inLabelText;

            //Move to right anchor
            double xLabelAnchor = inAnchor.Item1;
            double yLabelAnchor = inAnchor.Item2;

            //Apply conversion factor
            double markerWidth = markerLabelElement.Geometry.Envelope.Width;
            double markerHeight = markerLabelElement.Geometry.Envelope.Height;
            double parentWidth = parentElement.Geometry.Envelope.Width;
            double parentHeight = parentElement.Geometry.Envelope.Height;

            //For label right on top of another label
            switch (wantedPosition)
            {
                case Constants.Styles.MarkerLabelPositioning.FromCenterToUpperLeft:

                    //Validate x since esri sends an integer instead of a double, for some reasons ...
                    double doubleMin = parentElement.Geometry.Envelope.XMin;
                    decimal _xmin = Convert.ToDecimal(parentElement.Geometry.Envelope.XMin);
                    if (Decimal.Floor(_xmin) == _xmin)
                    {
                        doubleMin = doubleMin + 0.411; //This value was found by checking in Arc Map the true double xmin value in the properties of the parent element.
                    }

                    //Value were found from manually placing the label at wanted place and calculating the ratio for the best move. 
                    xLabelAnchor = doubleMin - (markerWidth * 0.5541573);  //TODO move hardcoded value somewhere else
                    yLabelAnchor = parentElement.Geometry.Envelope.YMin + (markerHeight * 0.5315338); //TODO move hardcoded value somewhere else

                    break;

                case Constants.Styles.MarkerLabelPositioning.FromCenterToUpperRight:
                    //Value were found from manually placing the label at wanted place and calculating the ratio for the best move. 
                    xLabelAnchor = parentElement.Geometry.Envelope.XMin + (markerWidth / 2.0) * 2.18849;  //TODO move hardcoded value somewhere else
                    yLabelAnchor = parentElement.Geometry.Envelope.YMin + (markerHeight) * 0.59923; //TODO move hardcoded value somewhere else
                    break;

                case Constants.Styles.MarkerLabelPositioning.FromCenterToUpperRightTight:
                    //Value were found from manually placing the label at wanted place and calculating the ratio for the best move. 
                    xLabelAnchor = parentElement.Geometry.Envelope.XMin + (markerWidth / 2.0) * 0.94849;  //TODO move hardcoded value somewhere else
                    yLabelAnchor = parentElement.Geometry.Envelope.YMin + (markerHeight) * 0.59923; //TODO move hardcoded value somewhere else
                    break;

                //This case is meant for when two labels must be added around a marker point
                case Constants.Styles.MarkerLabelPositioning.RightAboveCenter:

                    //Force y move on parent for a better fit of the two labels
                    ITransform2D transformParentElement = parentElement as ITransform2D;
                    if (parentPosition == Constants.Styles.MarkerLabelPositioning.FromCenterToUpperLeft)
                    {
                        transformParentElement.Move(-0.47, -parentHeight * 0.5);
                    }
                    else
                    {
                        transformParentElement.Move(0, -parentHeight * 0.5);
                    }
                    

                    //Value were found from manually placing the label at wanted place and calculating the ratio for the best move. 
                    xLabelAnchor = parentElement.Geometry.Envelope.XMin + parentWidth / 2.0;
                    yLabelAnchor = parentElement.Geometry.Envelope.YMax + markerHeight / 4.0;

                    break;
                default:
                    break;
            }

            Tuple<double, double> labelAnchor = new Tuple<double, double>(xLabelAnchor, yLabelAnchor);
            MoveItemToAnchorPoint(markerLabelElement, labelAnchor);

            //Add base element
            IPageLayout l = inDocument.ActiveView as IPageLayout;
            IGraphicsContainerSelect gcs = l as IGraphicsContainerSelect;
            inDocument.ActiveView.GraphicsContainer.AddElement(markerLabelElement as IElement, 0);
            inDocument.ActiveView.GraphicsContainer.BringToFront(gcs.SelectedElements);

            //Unselect
            gcs.UnselectElement(markerLabelElement);

            //Add to legend list
            legendElementList.Add(markerLabelElement);

            return markerLabelElement;
        }
        #endregion

        #region BUILD GRAPHIC

        /// <summary>
        /// Will create a marker symbol from given type, order and style. Will also return an offset parameter for linear markers
        /// </summary>
        /// <param name="markerType">element marker type (as stated in user table column element)</param>
        /// <param name="markerOrder">element order (for naming purposes)</param>
        /// <param name="markerStyle">element style</param>
        /// <param name="offset">out offset for linear markers</param>
        /// <returns></returns>
        private IElement BuildMarker(string markerType, double markerOrder, string markerStyle, out Tuple<double, double> offset)
        {
            //Get appropriate element
            IElement pointElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[markerType]) as IElement;
            IElementProperties pointElProp = pointElement as IElementProperties;

            //Rename
            pointElProp.Name = pointElProp.Name + markerOrder.ToString();

            //Symbolize if symbol can be found in style file
            IMarkerElement currentShapeElement = pointElement as IMarkerElement;
            double originalAngle = currentShapeElement.Symbol.Angle;
            offset = null;

            if (markerSymbolDico.ContainsKey(markerStyle))
            {
                IMarkerSymbol markSymbol = Services.ObjectManagement.CopyInputObject(markerSymbolDico[markerStyle]) as IMarkerSymbol;
                markSymbol.Angle = originalAngle; //Keep original angle (could be comming from POINT_CC_45), because style symbol doesn't have an angle by default.
                offset = new Tuple<double, double>(markSymbol.XOffset, markSymbol.YOffset);

                //Get rid of offset for linear markers
                if (markerType == Constants.Graphics.pointAngleLine)
                {
                    markSymbol.XOffset = 0;
                    markSymbol.YOffset = 0;
                }

                currentShapeElement.Symbol = markSymbol;
            }
            else
            {
                //Apply missing style
                ICharacterMarkerSymbol missingFillSymbol = Services.Symbols.GetMissingPointSymbol();
                currentShapeElement.Symbol = missingFillSymbol;
            }

            return pointElement;

        }

        #endregion

        #region VALIDATION
        /// <summary>
        /// Will save asset into default arcgis folder in user my document if it doesn't already exists.
        /// </summary>
        /// <returns>Output path for the new mxd.</returns>
        private string ValidateTemplateMXDExistance()
        {
            string outputFolderName = System.IO.Path.Combine(Dictionaries.Constants.ESRI.defaultArcGISFolderName, Dictionaries.Constants.Namespaces.mainNamespace + " " + ThisAddIn.Version.ToString());
            string outputFolderPath = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments), outputFolderName);
            string outputFilePath = System.IO.Path.Combine(outputFolderPath, Dictionaries.Constants.Assets.mxdEmbeddedFile);
            if (!System.IO.File.Exists(outputFilePath))
            {
                Services.FolderAndFiles.WriteResourceToFile(Dictionaries.Constants.Assets.mxdEmbeddedFile, Dictionaries.Constants.Assets.AssetFolder, Dictionaries.Constants.Namespaces.mainNamespace, outputFolderPath);
            }

            return outputFilePath;
        }

        /// <summary>
        /// Will return true if wanted style file is loaded inside map document
        /// </summary>
        /// <returns>true if style is loaded</returns>
        public bool ValidateStyleFile()
        {
            //Variable
            bool isLoaded = false;

            //Create a style gallery object from current arc map document
            IStyleGalleryStorage styleStore = ArcMap.Document.StyleGallery as IStyleGalleryStorage;

            //Iterate through storage and detect existing files
            int[] fileIndexes = Enumerable.Range(0, styleStore.FileCount).ToArray();
            foreach (int indexes in fileIndexes)
            {
                //Detect already loaded styles
                if (otherComponents.ContainsKey(Dictionaries.Constants.Styles.styleNameJSON))
                {
                    if (styleStore.File[indexes].Contains(otherComponents[Dictionaries.Constants.Styles.styleNameJSON]))
                    {
                        isLoaded = true;

                        //Keep style path
                        gscStyle = ArcMap.Document.StyleGallery;
                        gscStylePath = styleStore.File[indexes];
                        break;
                    }
                }

            }

            return isLoaded;
        }

        /// <summary>
        /// Will return wanted json file path, if it doesn't exist a copy will be made inside My Document\Arc GIS folder
        /// </summary>
        /// <returns></returns>
        public void ValidateJsonFilesExistance()
        {
            string outputFolderName = System.IO.Path.Combine(Dictionaries.Constants.ESRI.defaultArcGISFolderName, Dictionaries.Constants.Namespaces.mainNamespace + " " + ThisAddIn.Version.ToString());
            string outputFolderPath = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments), outputFolderName);
            jsonYSpacingFilePath = System.IO.Path.Combine(outputFolderPath, Dictionaries.Constants.Assets.jsonYSpacingEmbeddedFile);
            if (!System.IO.File.Exists(jsonYSpacingFilePath))
            {
                Services.FolderAndFiles.WriteResourceToFile(Dictionaries.Constants.Assets.jsonYSpacingEmbeddedFile, Dictionaries.Constants.Assets.AssetFolder, Dictionaries.Constants.Namespaces.mainNamespace, outputFolderPath);
            }

            jsonXSpacingFilePath = System.IO.Path.Combine(outputFolderPath, Dictionaries.Constants.Assets.jsonXSpacingEmbeddedFile);
            if (!System.IO.File.Exists(jsonXSpacingFilePath))
            {
                Services.FolderAndFiles.WriteResourceToFile(Dictionaries.Constants.Assets.jsonXSpacingEmbeddedFile, Dictionaries.Constants.Assets.AssetFolder, Dictionaries.Constants.Namespaces.mainNamespace, outputFolderPath);
            }

            jsonOtherFilePath = System.IO.Path.Combine(outputFolderPath, Dictionaries.Constants.Assets.jsonStyleFontsOtherEmbeddedFile);
            if (!System.IO.File.Exists(jsonOtherFilePath))
            {
                Services.FolderAndFiles.WriteResourceToFile(Dictionaries.Constants.Assets.jsonStyleFontsOtherEmbeddedFile, Dictionaries.Constants.Assets.AssetFolder, Dictionaries.Constants.Namespaces.mainNamespace, outputFolderPath);
            }

        }

        /// <summary>
        /// Will save asset into default arcgis folder in user my document if it doesn't alreayd exists.
        /// </summary>
        public string ValidateDEMPictureExistance()
        {
            string outputFolderName = System.IO.Path.Combine(Dictionaries.Constants.ESRI.defaultArcGISFolderName, Dictionaries.Constants.Namespaces.mainNamespace + " " + ThisAddIn.Version.ToString());
            string outputFolderPath = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments), outputFolderName);
            string outputFilePath = System.IO.Path.Combine(outputFolderPath, Dictionaries.Constants.Assets.demPicture);
            if (!System.IO.File.Exists(outputFilePath))
            {
                Services.FolderAndFiles.WriteResourceToFile(Dictionaries.Constants.Assets.demPicture, Dictionaries.Constants.Assets.AssetFolder, Dictionaries.Constants.Namespaces.mainNamespace, outputFolderPath);
            }

            return outputFilePath;
        }

        /// <summary>
        /// Will validate if input element is a  non flat (height 0) line element, or a group element composed of a bunch of non-flat lines.
        /// Knowing this can results in different methods of anchoring since they behave like left anchor compared to whatever they are set to.
        /// Flat line graphics always behave like anchor point is at the left side. Area of lines and heighted lines behaves the same but can be seen as polygons.
        /// </summary>
        /// <param name="inElement"></param>
        /// <returns></returns>
        public bool IsElementAllNonFlatLines(IElement inElement)
        {
            bool allLines = true;

            //For single line element
            if (inElement.Geometry.GeometryType != esriGeometryType.esriGeometryPolyline)
            {
                allLines = false;
            }
            else
            {
                //Check height
                if (inElement.Geometry.Envelope.Height == 0)
                {
                    allLines = false;
                }
                else
                {
                    allLines = true;
                }
            }

            //for group of elements
            IGroupElement inGroupElement = inElement as IGroupElement;
            if (inGroupElement != null)
            {
                //Check geometry of inner elements, if it's all lines
                for (int el = 0; el < inGroupElement.ElementCount; el++)
                {
                    IElement innerElement = inGroupElement.Element[el];
                    if (innerElement.Geometry.GeometryType != esriGeometryType.esriGeometryPolyline)
                    {
                        allLines = false;
                        break;
                    }
                    else
                    {
                        allLines = true;
                    }
                }
            }

            return allLines;
        }

        /// <summary>
        /// Will validate if input text is different then string.empty, a space, null or string literal "<null>"
        /// </summary>
        /// <param name="inputText"></param>
        /// <returns></returns>
        public bool IsTextEmpty(string inputText)
        {
            bool isEmpty = false;
            if (inputText == string.Empty || inputText == null || inputText == " " || inputText == Constants.TextConfiguration.NullLiteral)
            {
                isEmpty = true;
            }

            return isEmpty;
        }
        #endregion

        #region CALCULATIONS

        /// <summary>
        /// Will calculate a text element height based on wanted text inside it, if width and font size is fixed.
        /// </summary>
        /// <param name="inText"></param>
        /// <param name="minHeight">A minimal height in case text is a bit bolder or bigger, used for heading for example</param>
        /// <returns></returns>
        public double GetTextHeight(string inText, double maxWidth, double minHeight = 0.0, double fontSize = 8.0)
        {
            //Count total width of text
            double textWidth = 0.0;
            double tHeight = 0.0;
            int j;

            //Adjust with possible font GSCGeology2015. Need to have bigger box
            if (otherComponents.ContainsKey(Constants.Fonts.geologytFontNameJSON) && inText.Contains(otherComponents[Constants.Fonts.geologytFontNameJSON]))
            {
                tHeight = tHeight + Constants.Fonts.geologyFontHeightAjustement;

                //Strip text of tags that could make it look longer then it is
                inText = inText.Replace(Constants.TextConfiguration.tagFont + '"' + otherComponents[Constants.Fonts.geologytFontNameJSON] + '"' + ">", "");
            }

            if (arialCharactersWidth == null)
            {
                arialCharactersWidth = GetArialCharacterWidth();
            }

            //Strip text of tags that could make it look longer then it is
            inText = inText.Replace(Constants.TextConfiguration.tagAllCaps, "");
            inText = inText.Replace(Constants.TextConfiguration.tagBold, "");
            inText = inText.Replace(Constants.TextConfiguration.tagItalic, "");
            inText = inText.Replace(Constants.TextConfiguration.endTagAllCaps, "");
            inText = inText.Replace(Constants.TextConfiguration.endTagBold, "");
            inText = inText.Replace(Constants.TextConfiguration.endTagItalic, "");
            inText = inText.Replace(Constants.TextConfiguration.endTagFont, "");

            for (int i = 0; i < inText.Length; i++)
            {
                j = Encoding.Default.GetBytes(inText.Substring(i, 1))[0];
                if (j >= 32)
                {
                    if (arialCharactersWidth.ContainsKey(j))
                    {
                        textWidth = textWidth + (fontSize * arialCharactersWidth[j]);
                    }
                    else
                    {
                        textWidth = textWidth + (fontSize * 1);
                    }
                    
                }
            }

            //Calculate approx. number of lines
            double numberLines = (textWidth * 0.352778) / maxWidth;
            numberLines = Math.Ceiling(numberLines); //Round to upper boundary

            //Extra validation
            if (numberLines >= 6)
            {
                double extraWidth = (textWidth * 0.02) + textWidth; //Extra percent of width, in case
                numberLines = (extraWidth * 0.352778) / maxWidth;
                numberLines = Math.Ceiling(numberLines);
            }

            //Height
            if (Constants.TextConfiguration.lineHeight < minHeight)
            {
                tHeight = tHeight + (numberLines * (minHeight));
            }
            else
            {
                tHeight = tHeight + (numberLines * (Constants.TextConfiguration.lineHeight));
            }

            return tHeight;

        }

        /// <summary>
        /// Will output a dictionnary containing width in points for all arial character
        /// </summary>
        /// <returns></returns>
        public Dictionary<int, double> GetArialCharacterWidth()
        {
            Dictionary<int, double> arialCharacterWidth = new Dictionary<int, double>();

            for (int i = 32; i <= 127; i++)
            {
                switch (i)
                {
                    case 39:
                    case 106:
                    case 108:
                        arialCharacterWidth[i] = 0.1902;
                        break;
                    case 105:
                    case 116:
                        arialCharacterWidth[i] = 0.2526;
                        break;
                    case 32:
                    case 33:
                    case 44:
                    case 46:
                    case 47:
                    case 58:
                    case 59:
                    case 73:
                    case 91:
                    case 92:
                    case 93:
                    case 102:
                    case 124:
                        arialCharacterWidth[i] = 0.3144;
                        break;
                    case 34:
                    case 40:
                    case 41:
                    case 45:
                    case 96:
                    case 114:
                    case 123:
                    case 125:
                        arialCharacterWidth[i] = 0.3768;
                        break;
                    case 42:
                    case 94:
                    case 118:
                    case 120:
                        arialCharacterWidth[i] = 0.4392;
                        break;
                    case 107:
                    case 115:
                    case 122:
                        arialCharacterWidth[i] = 0.501;
                        break;
                    case 35:
                    case 36:
                    case 48:
                    case 49:
                    case 50:
                    case 51:
                    case 52:
                    case 53:
                    case 54:
                    case 55:
                    case 56:
                    case 57:
                    case 63:
                    case 74:
                    case 76:
                    case 84:
                    case 90:
                    case 95:
                    case 97:
                    case 98:
                    case 99:
                    case 100:
                    case 101:
                    case 103:
                    case 104:
                    case 110:
                    case 111:
                    case 112:
                    case 113:
                    case 117:
                    case 121:
                        arialCharacterWidth[i] = 0.5634;
                        break;
                    case 43:
                    case 60:
                    case 61:
                    case 62:
                    case 70:
                    case 126:
                        arialCharacterWidth[i] = 0.6252;
                        break;
                    case 38:
                    case 65:
                    case 66:
                    case 69:
                    case 72:
                    case 75:
                    case 78:
                    case 80:
                    case 82:
                    case 83:
                    case 85:
                    case 86:
                    case 88:
                    case 89:
                    case 119:
                        arialCharacterWidth[i] = 0.6876;
                        break;
                    case 67:
                    case 68:
                    case 71:
                    case 79:
                    case 81:
                        arialCharacterWidth[i] = 0.7494;
                        break;
                    case 77:
                    case 109:
                    case 127:
                        arialCharacterWidth[i] = 0.8118;
                        break;
                    case 37:
                        arialCharacterWidth[i] = 0.936;
                        break;
                    case 64:
                    case 87:
                        arialCharacterWidth[i] = 1.0602;
                        break;
                    default:
                        break;
                }
            }

            return arialCharacterWidth;
        }

        /// <summary>
        /// From a given element and anchor point, will calculate a new polygon enveloppe to fit anchor point so the element can
        /// be set at the right place on the layout before being moved.
        /// </summary>
        /// <param name="inElement"></param>
        /// <param name="inAnchor"></param>
        /// <returns></returns>
        public void SetRectangularPolygonFromAnchorType(IElement inElement, Tuple<double, double> inAnchor)
        {
            SetRectnagularPolygonFromAnchorTypeAndHeight(inElement, inAnchor, inElement.Geometry.Envelope.Height);
        }

        /// <summary>
        /// From a given element, anchor point and height, will calculate a new polygon envelope
        /// </summary>
        /// <param name="inElement"></param>
        /// <param name="inAnchor"></param>
        /// <param name="inHeight"></param>
        public void SetRectnagularPolygonFromAnchorTypeAndHeight(IElement inElement, Tuple<double, double> inAnchor, double inHeight)
        {
            //Get anchor type
            IElementProperties3 elemProp3 = inElement as IElementProperties3;
            esriAnchorPointEnum currentAnchorPointType = elemProp3.AnchorPoint;
            esriGeometryType inGeometryType = inElement.Geometry.GeometryType;

            //Apply conversion factor
            //inHeight = inHeight ;
            double inElementWidth = inElement.Geometry.Envelope.Width;
            double inElementHeight = inElement.Geometry.Envelope.Height;

            if (currentAnchorPointType == esriAnchorPointEnum.esriTopLeftCorner)
            {
                //Set new envelope
                IEnvelope envUnitBox = new Envelope() as IEnvelope;
                double minX = inAnchor.Item1;
                double minY = inAnchor.Item2 - inHeight;
                double maxX = inAnchor.Item1 + inElementWidth;
                double maxY = inAnchor.Item2;
                envUnitBox.PutCoords(minX, minY, maxX, maxY);
                inElement.Geometry = envUnitBox;
            }

            if (currentAnchorPointType == esriAnchorPointEnum.esriCenterPoint)
            {

                //Set new envelope
                IEnvelope envUnitLabelBox = new Envelope() as IEnvelope;
                double minX = inAnchor.Item1 - inElementWidth / 2.0;
                double minY = inAnchor.Item2 - inHeight / 2.0;
                double maxX = inAnchor.Item1 + inElementWidth / 2.0;
                double maxY = inAnchor.Item2 + inHeight / 2.0;

                //Validate if polygon graphic is a group of lines elements, which behaves differently 
                //and for some reasons even if the anchor point is center center, behaves likes it's left
                //see same case for line graphices.
                if (IsElementAllNonFlatLines(inElement))
                {
                    minX = inAnchor.Item1 + elementWidth / 2.0 - inElementWidth / 2.0;
                    minY = inAnchor.Item2 - inHeight / 2.0;
                    maxX = minX;
                    maxY = minY;
                }


                envUnitLabelBox.PutCoords(minX, minY, maxX, maxY);
                inElement.Geometry = envUnitLabelBox;
            }

            if (currentAnchorPointType == esriAnchorPointEnum.esriTopMidPoint)
            {

                //Set new envelope
                IEnvelope envUnitLabelBox = new Envelope() as IEnvelope;
                double minX = inAnchor.Item1 - inElementWidth / 2.0;
                double minY = inAnchor.Item2 - inHeight;
                double maxX = inAnchor.Item1 + inElementWidth / 2.0;
                double maxY = inAnchor.Item2;
                envUnitLabelBox.PutCoords(minX, minY, maxX, maxY);
                inElement.Geometry = envUnitLabelBox;
            }

            if (currentAnchorPointType == esriAnchorPointEnum.esriLeftMidPoint)
            {

                //Set new envelope
                IEnvelope envUnitLabelBox = new Envelope() as IEnvelope;
                double minX = inAnchor.Item1;
                double minY = inAnchor.Item2 - inHeight / 2.0;
                double maxX = inAnchor.Item1 + inElementWidth;
                double maxY = inAnchor.Item2 + inHeight / 2.0;
                envUnitLabelBox.PutCoords(minX, minY, maxX, maxY);
                inElement.Geometry = envUnitLabelBox;
            }

            if (currentAnchorPointType == esriAnchorPointEnum.esriBottomRightCorner)
            {
                //Set new envelope
                IEnvelope envUnitLabelBox = new Envelope() as IEnvelope;
                double minX = inAnchor.Item1 - inElementWidth;
                double minY = inAnchor.Item2;
                double maxX = inAnchor.Item1;
                double maxY = inAnchor.Item2 + inHeight;
                envUnitLabelBox.PutCoords(minX, minY, maxX, maxY);
                inElement.Geometry = envUnitLabelBox;
            }

            if (currentAnchorPointType == esriAnchorPointEnum.esriBottomLeftCorner)
            {
                //Set new envelope
                IEnvelope envUnitLabelBox = new Envelope() as IEnvelope;
                double minX = inAnchor.Item1;
                double minY = inAnchor.Item2;
                double maxX = inAnchor.Item1 + inElementWidth;
                double maxY = inAnchor.Item2 - inHeight;
                envUnitLabelBox.PutCoords(minX, minY, maxX, maxY);
                inElement.Geometry = envUnitLabelBox;
            }


        }

        /// <summary>
        /// From a given element and anchor point, will calculate a new polygon enveloppe to fit anchor point so the element can
        /// be set at the right place on the layout before being moved.
        /// </summary>
        /// <param name="inElement"></param>
        /// <param name="inAnchor"></param>
        /// <returns></returns>
        public void SetPolygonFromAnchorType(IElement inElement, Tuple<double, double> inAnchor)
        {
            //Get anchor type
            IElementProperties3 elemProp3 = inElement as IElementProperties3;
            esriAnchorPointEnum currentAnchorPointType = elemProp3.AnchorPoint;

            //Cast for transformation
            ITransform2D transElement = inElement as ITransform2D;

            //Cast original geometry
            IPolygon inPolygon = inElement.Geometry as IPolygon;


            //Create a point from anchor point coordaintes
            IPoint aPoint = new ESRI.ArcGIS.Geometry.Point();
            aPoint.SpatialReference = inPolygon.SpatialReference;

            //Apply conversion factor
            double inElementWidth = inElement.Geometry.Envelope.Width;
            double inElementHeight = inElement.Geometry.Envelope.Height;

            //Get max width
            double maxWidth = GetXSpacing(Constants.Graphics.elementWidth);

            if (currentAnchorPointType == esriAnchorPointEnum.esriTopLeftCorner)
            {
                if (inElementWidth >= maxWidth)
                {
                    aPoint.PutCoords(inAnchor.Item1 + inElementWidth / 2.0, inAnchor.Item2); //Transform.moveVector is runned with centroid.
                }
                else
                {
                    aPoint.PutCoords(inAnchor.Item1 + maxWidth / 2.0, inAnchor.Item2); //Transform.moveVector is runned with centroid.
                }


                //Create a vector that will be used to move original polygon
                ILine movingLine = new Line();
                movingLine.PutCoords(inPolygon.FromPoint, aPoint);

                //Move
                transElement.MoveVector(movingLine);


            }

            if (currentAnchorPointType == esriAnchorPointEnum.esriCenterPoint)
            {
                if (inElementWidth >= maxWidth)
                {
                    aPoint.PutCoords(inAnchor.Item1 + inElementWidth / 2.0, inAnchor.Item2); //Transform.moveVector is runned with centroid.
                }
                else
                {
                    aPoint.PutCoords(inAnchor.Item1 + maxWidth / 2.0, inAnchor.Item2); //Transform.moveVector is runned with centroid.
                }


                //Create a vector that will be used to move original polygon
                ILine movingLine = new Line();
                movingLine.PutCoords(inPolygon.FromPoint, aPoint);

                //Move
                transElement.MoveVector(movingLine);

                //Move again
                IArea inArea = inElement.Geometry as IArea;
                double xMove = Math.Abs(inArea.Centroid.X - aPoint.X);
                double yMove = Math.Abs(inArea.Centroid.Y - aPoint.Y);
                transElement.Move(xMove, -yMove);


            }

        }

        /// <summary>
        /// From a given element will calculate a new line geometry to fit anchor point so the element can
        /// be set at the right place on the layout before being moved.
        /// NOTE Anchor Point type doesnt' change a thing on the placement of the element.
        /// </summary>
        /// <param name="inElement"></param>
        /// <param name="inAnchor">Without embedded y spacing</param>
        /// <returns></returns>
        public void SetLineFromAnchorType(IElement inElement, IElement inParentElement, Tuple<double, double> inAnchor, double length)
        {

            //Get line
            IPolyline inPolyline = inElement.Geometry as IPolyline;

            //Get height
            double height = inPolyline.Envelope.Height;

            //Apply conversion factor
            IPolycurve polycurveElement = inElement.Geometry as IPolycurve;
            double polycurveWidth = polycurveElement.Envelope.Width;
            double polycurveHeight = polycurveElement.Envelope.Height;

            if (height == 0)
            {
                //Set new polyline
                IPolyline polylineElement = new ESRI.ArcGIS.Geometry.Polyline() as IPolyline;

                //Set new starting point to be center
                IPoint fromPoint = new ESRI.ArcGIS.Geometry.Point();
                fromPoint.X = inAnchor.Item1;
                fromPoint.Y = inAnchor.Item2;

                IPoint toPoint = new ESRI.ArcGIS.Geometry.Point();
                toPoint.X = inAnchor.Item1 + length;
                toPoint.Y = inAnchor.Item2;

                polylineElement.FromPoint = fromPoint;
                polylineElement.ToPoint = toPoint;
                inElement.Geometry = polylineElement;
            }
            else
            {
                //NOTE: This situation is for lines that has heights. In that case, it's easier to
                //Move the whole element then trying to reset anchor point, else we need to move
                // all curves start/end points inside the polycurve object and that is a pain.

                //Get anchor type
                IElementProperties3 inElementProperties = inElement as IElementProperties3;

                //Calculate distance between current location and new anchor point
                double oriCenterX = polycurveElement.Envelope.XMin + polycurveWidth / 2.0;
                double oriCenterY = polycurveElement.Envelope.YMin + polycurveHeight; //It behaves like bottom left anchor...weird
                double dX = 0;
                double dY = 0;

                if (inElementProperties.AnchorPoint == esriAnchorPointEnum.esriCenterPoint)
                {
                    #region CENTER POINT

                    //Calculate distance between current location and new anchor point
                    if (oriCenterX < inAnchor.Item1)
                    {
                        dX = -(oriCenterX) + inAnchor.Item1 + polycurveWidth / 2.0;
                        dY = (oriCenterY) - inAnchor.Item2 - polycurveHeight / 2.0;

                        //In cases where the element is smaller in width then the element column itself
                        if (polycurveWidth < elementWidth)
                        {
                            dX = -(oriCenterX) + inAnchor.Item1 + elementWidth / 2.0;
                            dY = (oriCenterY) - inAnchor.Item2 - polycurveHeight / 2.0;
                        }

                    }
                    else if (oriCenterX > inAnchor.Item1)
                    {
                        dX = (oriCenterX) - inAnchor.Item1 - polycurveWidth / 2.0;
                        dY = (oriCenterY) - inAnchor.Item2 - polycurveHeight / 2.0;

                        //In cases where the element is smaller in width then the element column itself
                        if (polycurveWidth < elementWidth)
                        {
                            dX = (oriCenterX) - inAnchor.Item1 - elementWidth / 2.0;
                            dY = (oriCenterY) - inAnchor.Item2 - polycurveHeight / 2.0;
                        }
                    }
                    #endregion
                }
                else if (inElementProperties.AnchorPoint == esriAnchorPointEnum.esriBottomRightCorner)
                {
                    #region BOTTOM RIGHT


                    if (oriCenterX < inAnchor.Item1)
                    {
                        dX = -(oriCenterX) + inAnchor.Item1 + polycurveWidth;
                        dY = (oriCenterY) - inAnchor.Item2 - polycurveHeight;

                        //In cases where the element is smaller in width then the element column itself
                        if (polycurveWidth < elementWidth)
                        {
                            dX = -(oriCenterX) + inAnchor.Item1 + elementWidth;
                            dY = (oriCenterY) - inAnchor.Item2 - polycurveHeight;
                        }
                    }
                    else if (oriCenterX > inAnchor.Item1)
                    {
                        dX = (oriCenterX) - inAnchor.Item1 - polycurveWidth;
                        dY = (oriCenterY) - inAnchor.Item2 - polycurveHeight;

                        //In cases where the element is smaller in width then the element column itself
                        if (polycurveWidth < elementWidth)
                        {
                            dX = (oriCenterX) - inAnchor.Item1 - elementWidth;
                            dY = (oriCenterY) - inAnchor.Item2 - polycurveHeight;
                        }
                    }
                    #endregion
                }

                //Move
                ITransform2D transElement = inElement as ITransform2D;


                if (oriCenterX < inAnchor.Item1 && oriCenterY < inAnchor.Item2)
                {
                    transElement.Move(dX, Math.Abs(dY)); //Move accordingly to anchor point which is center center
                }
                else if (oriCenterX > inAnchor.Item1 && oriCenterY < inAnchor.Item2)
                {
                    transElement.Move(-dX, Math.Abs(dY));
                }
                else if (oriCenterX > inAnchor.Item1 && oriCenterY > inAnchor.Item2)
                {
                    transElement.Move(-dX, -dY);
                }
                else if (oriCenterX < inAnchor.Item1 && oriCenterY > inAnchor.Item2)
                {
                    transElement.Move(dX, -dY);
                }

            }



        }

        /// <summary>
        /// From a given element will calculate a new point geometry to fit anchor point so the element can
        /// be set at the right place on the layout before being moved.
        /// NOTE Anchor Point type doesnt' change a thing on the placement of the element.
        /// </summary>
        /// <param name="inElement"></param>
        /// <param name="inAnchor">Without embedded y spacing</param>
        /// <returns></returns>
        public void SetPointFromAnchorType(IElement inElement, Tuple<double, double> inAnchor, Tuple<double, double> offset)
        {
            //Get anchor type
            IElementProperties3 elemProp3 = inElement as IElementProperties3;
            esriAnchorPointEnum currentAnchorPointType = elemProp3.AnchorPoint;
            IPoint newPoint = new ESRI.ArcGIS.Geometry.Point();

            //Apply conversion factor
            double inElementHeight = inElement.Geometry.Envelope.Height;

            //Get offset
            double xOff = 0;
            double yOff = 0;
            if (offset != null)
            {
                xOff = offset.Item1;
                yOff = offset.Item2;
            }
            else
            {
                offset = new Tuple<double, double>(xOff, yOff);
            }

            switch (currentAnchorPointType)
            {
                case esriAnchorPointEnum.esriTopLeftCorner:
                    break;
                case esriAnchorPointEnum.esriTopMidPoint:
                    break;
                case esriAnchorPointEnum.esriTopRightCorner:
                    break;
                case esriAnchorPointEnum.esriLeftMidPoint:
                    break;
                case esriAnchorPointEnum.esriCenterPoint:
                    newPoint.X = inAnchor.Item1 + elementWidth / 2.0 + xOff;
                    newPoint.Y = inAnchor.Item2 + offset.Item2;
                    inElement.Geometry = newPoint;
                    break;
                case esriAnchorPointEnum.esriRightMidPoint:
                    break;
                case esriAnchorPointEnum.esriBottomLeftCorner:
                    break;
                case esriAnchorPointEnum.esriBottomMidPoint:
                    newPoint.X = inAnchor.Item1 + elementWidth / 2.0 + yOff;
                    newPoint.Y = inAnchor.Item2 - inElementHeight / 2.0 + offset.Item2;
                    inElement.Geometry = newPoint;
                    break;
                case esriAnchorPointEnum.esriBottomRightCorner:
                    break;
                default:
                    break;
            }




        }

        /// <summary>
        /// Will output the x,y coordinate for the first anchor of the legend
        /// Default is center of layout, else if inside CGM template, will be placed inside the LEGEND element.
        /// </summary>
        /// <returns></returns>
        public Tuple<double, double> GetAnchorPointStart()
        {
            Tuple<double, double> anchorCoord = new Tuple<double, double>(0.0, 0.0);

            IMxDocument currentDoc = (IMxDocument)ArcMap.Application.Document;

            //Get if an element is named Legend, if yes take starting point from it.
            IGraphicsContainer graphics = currentDoc.ActiveView.GraphicsContainer;
           
            //Filter graphics
            IElement existingLegendElement = null;
            graphics.Reset();
            IElement nextElement = graphics.Next();
            while (nextElement != null)
            {
                IElementProperties nextGProperties = nextElement as IElementProperties;
                if (nextGProperties.Name == Constants.Graphics.cgmLegendElement)
                {
                    existingLegendElement = nextElement;
                }
                if (nextGProperties.Name.Contains(Constants.Graphics.cgmDetectorKeyword))
                {
                    isCGMTemplateMXD = true;
                }
                nextElement = graphics.Next();
            }

            //Get previously saved anchor
            anchorCoord = new Tuple<double, double>(GSC_Legend_Renderer.Properties.Settings.Default.LegendAnchorX, GSC_Legend_Renderer.Properties.Settings.Default.LegendAnchorY);

            //Case when CGM blue legend box is detected, get new anchor
            if (existingLegendElement != null)
            {

                anchorCoord = new Tuple<double, double>(existingLegendElement.Geometry.Envelope.XMin, existingLegendElement.Geometry.Envelope.YMax);

                //Keep element
                originalCGMLegend = existingLegendElement;
                isCGMTemplateMXD = true;


                //Keep in memory anchor, in case user needs to launch renderer again within same session
                GSC_Legend_Renderer.Properties.Settings.Default.LegendAnchorX = anchorCoord.Item1;
                GSC_Legend_Renderer.Properties.Settings.Default.LegendAnchorY = anchorCoord.Item2;
                GSC_Legend_Renderer.Properties.Settings.Default.Save();

            }
            //Case when nothing has been save relative to cgm blue legend box, default to upper left of paper layout
            else if (anchorCoord.Item2 == 0.0 || isCGMTemplateMXD == false)
            {
                anchorCoord = new Tuple<double, double>(0.0, currentDoc.PageLayout.Page.PrintableBounds.UpperLeft.Y);
                isCGMTemplateMXD = false;
            }


            return anchorCoord;

        }

        /// <summary>
        /// From a given anchor, will calculate the legend maximum height, 
        /// based on upper cgm citation graphic anchor. This will be used to
        /// automatically calculate the number of columns
        /// </summary>
        /// <param name="in_ySpacingWithCitation">The wanted y spacing between end of legend and the reference graphic</param>
        /// <returns></returns>
        public double GetCGMLegendLowerBound(double in_ySpacingWithCitation, string in_referenceCGMGraphicName)
        {
            //Variables
            IMxDocument currentDoc = (IMxDocument)ArcMap.Application.Document;
            double outYBound = 0.0;

            //Get if an element is named citation, if yes take starting point from it.
            IGraphicsContainer graphics = currentDoc.ActiveView.GraphicsContainer;
            IElement existingCitationElement = null;
            graphics.Reset();
            IElement nextElement = graphics.Next();
            while (nextElement != null)
            {
                IElementProperties nextGProperties = nextElement as IElementProperties;
                if (nextGProperties.Name == in_referenceCGMGraphicName)
                {
                    existingCitationElement = nextElement;

                    break;
                }

                nextElement = graphics.Next();
            }

            if (existingCitationElement != null)
            {

                outYBound = existingCitationElement.Geometry.Envelope.YMax + in_ySpacingWithCitation;


            }

            return outYBound;
        }

        /// <summary>
        /// Will move a given element to an anchor point by creating a vector line
        /// </summary>
        /// <param name="inElement"></param>
        /// <param name="inAnchor"></param>
        public void MoveItemToAnchorPoint(IElement inElement, Tuple<double, double> inAnchor)
        {
            //Create new vector line
            ILine moveVector = new Line();

            //Apply conversion factor
            double inElementWidth = inElement.Geometry.Envelope.Width;
            double inElementHeight = inElement.Geometry.Envelope.Height;

            //Get anchor type
            IElementProperties3 inElementProperties = inElement as IElementProperties3;

            if (inElementProperties.AnchorPoint == esriAnchorPointEnum.esriBottomRightCorner || inElementProperties.AnchorPoint == esriAnchorPointEnum.esriTopRightCorner)
            {
                //Create new start point
                IPoint startPoint = new ESRI.ArcGIS.Geometry.Point();
                startPoint.X = inElement.Geometry.Envelope.XMax;
                startPoint.Y = inElement.Geometry.Envelope.YMax;
                moveVector.FromPoint = startPoint;

            }
            if (inElementProperties.AnchorPoint == esriAnchorPointEnum.esriRightMidPoint)
            {
                //Create new start point
                IPoint startPoint = new ESRI.ArcGIS.Geometry.Point();
                startPoint.X = inElement.Geometry.Envelope.XMax;
                startPoint.Y = inElement.Geometry.Envelope.YMax - inElementHeight / 2.0;
                moveVector.FromPoint = startPoint;
            }
            if (inElementProperties.AnchorPoint == esriAnchorPointEnum.esriBottomLeftCorner)
            {
                //Create new start point
                IPoint startPoint = new ESRI.ArcGIS.Geometry.Point();
                startPoint.X = inElement.Geometry.Envelope.XMin;
                startPoint.Y = inElement.Geometry.Envelope.YMin;
                moveVector.FromPoint = startPoint;
            }
            if (inElementProperties.AnchorPoint == esriAnchorPointEnum.esriTopLeftCorner)
            {
                //Create new start point
                IPoint startPoint = new ESRI.ArcGIS.Geometry.Point();
                startPoint.X = inElement.Geometry.Envelope.XMin;
                startPoint.Y = inElement.Geometry.Envelope.YMax;
                moveVector.FromPoint = startPoint;
            }
            if (inElementProperties.AnchorPoint == esriAnchorPointEnum.esriLeftMidPoint)
            {
                //Create new start point
                IPoint startPoint = new ESRI.ArcGIS.Geometry.Point();
                startPoint.X = inElement.Geometry.Envelope.XMin;
                startPoint.Y = inElement.Geometry.Envelope.YMax - inElementHeight / 2.0;
                moveVector.FromPoint = startPoint;
            }
            if (inElementProperties.AnchorPoint == esriAnchorPointEnum.esriCenterPoint)
            {
                //Create new start point
                IPoint startPoint = new ESRI.ArcGIS.Geometry.Point();
                startPoint.X = inElement.Geometry.Envelope.XMin + inElementWidth / 2.0;
                startPoint.Y = inElement.Geometry.Envelope.YMax - inElementHeight / 2.0;
                moveVector.FromPoint = startPoint;
            }

            //Create new end point
            IPoint endPoint = new ESRI.ArcGIS.Geometry.Point();
            endPoint.X = inAnchor.Item1;
            endPoint.Y = inAnchor.Item2;
            moveVector.ToPoint = endPoint;

            //Move
            ITransform2D transElement = inElement as ITransform2D;
            transElement.MoveVector(moveVector); //Move accordingly to anchor point which is center center
        }

        /// <summary>
        /// Will set current document to be in given input unit.
        /// </summary>
        /// <param name="inDocument"></param>
        public esriUnits SetDocumentUnits(IMxDocument inDocument, esriUnits inUnitType)
        {
            esriUnits oriUnits = inDocument.PageLayout.Page.Units;
            inDocument.PageLayout.Page.Units = inUnitType;

            return oriUnits;
        }

        #endregion

        #region SYMBOLIZE GRAPHICS

        /// <summary>
        /// Will set the background color for given element with given style
        /// </summary>
        /// <param name="inElement"></param>
        /// <param name="style"></param>
        public IElement SetPolygonFill(IElement inElement, string style, bool isSimpleFill, bool isUnitBoxOnly = false, Tuple<double, double> inAnchor = null, string style2 = "")
        {

            //Symbolize if symbol can be found in style file
            //IGroupElement3 groupShapeElement = GetGroupLegendElement(Constants.Graphics.legendBoxDEM);
            IFillShapeElement intShapeElement = inElement as IFillShapeElement;

            //Detect cartographic line and force rounded joins
            ///Special case found for UNIT_SPLIT having the wrong join and showing a bad rendering.
            IMultiLayerLineSymbol multiLineSymbol = intShapeElement.Symbol.Outline as IMultiLayerLineSymbol;
            if (multiLineSymbol != null && style2 != "")
            {
                for (int i = 0; i < multiLineSymbol.LayerCount; i++)
                {
                    ICartographicLineSymbol cartoLineSymbol = multiLineSymbol.Layer[i] as ICartographicLineSymbol;
                    if (cartoLineSymbol != null)
                    {
                        cartoLineSymbol.Join = esriLineJoinStyle.esriLJSRound;
                    }

                }
            }
            ILineSymbol inOutline = intShapeElement.Symbol.Outline; //Cast to keep actual outline

            if (fillSymbolDico.ContainsKey(style) && isSimpleFill)
            {

                //Get symbol type and color
                string symbolTypeName = string.Empty;
                ISymbol fillSymbol = fillSymbolDico[style] as ISymbol;
                IColor symbolColor = Services.Symbols.GetPolygonSymbolColor(fillSymbol, out symbolTypeName);
                bool isFillNullColor = symbolColor.NullColor;
                IRgbColor rgbCol = symbolColor as IRgbColor;
                IFillSymbol iFillSymbol = fillSymbol as IFillSymbol;
                bool isOutlineNullColor = iFillSymbol.Outline.Color.NullColor;

                //Fill polygon or replace with related DEM image
                if (this.checkBox_DEMBoxes.Checked && isUnitBoxOnly && rgbCol != null)
                {
                    //Detect tranparent color and force it white
                    Color fillColors = Color.FromArgb(255, rgbCol.Red, rgbCol.Green, rgbCol.Blue);
                    if (isFillNullColor)
                    {
                        fillColors = Color.FromArgb(255, 255, 255, 255);
                    }
                    IElement demElement = SetPolygonDEM(inElement, fillColors, inAnchor);

                    //Create new symbol and apply, else it won't update...
                    ISimpleFillSymbol newSimpleFill = new SimpleFillSymbol();
                    newSimpleFill.Style = esriSimpleFillStyle.esriSFSHollow;

                    //Manage outline
                    if (isOutlineNullColor)
                    {
                        //Apply black outline 
                        newSimpleFill.Outline = inOutline;
                    }
                    else
                    {
                        //Keep wanted outline
                        newSimpleFill.Outline = iFillSymbol.Outline;
                    }


                    intShapeElement.Symbol = newSimpleFill;

                    return demElement;
                }
                //Overlay fill type 
                else if (symbolTypeName == Constants.ObjectNames.fillTypeMultilayer)
                {
                    //Will act as a non simple fill
                    IFillSymbol fillMulti = iFillSymbol;
                    
                    //Set color if needed
                    if (style2!= string.Empty && style2 != null && fillSymbolDico.ContainsKey(style2))
                    {
                        string symbolTypeName2 = string.Empty;
                        ISymbol fillSymbol2 = fillSymbolDico[style2] as ISymbol;
                        IColor symbolColor2 = Services.Symbols.GetPolygonSymbolColor(fillSymbol2, out symbolTypeName2);

                        fillMulti.Color = symbolColor2;
                    }

                    //Manage outline
                    if (isOutlineNullColor)
                    {
                        //Apply black outline 
                        fillMulti.Outline = inOutline;
                    }
                    else
                    {
                        //Keep wanted outline
                        fillMulti.Outline = multiLineSymbol;

                    }
                    intShapeElement.Symbol = fillMulti;

                    return inElement;
                }
                else
                {
                    //Create new symbol and apply, else it won't update...
                    ISimpleFillSymbol newSimpleFill = new SimpleFillSymbol();
                    newSimpleFill.Color = symbolColor;
                    newSimpleFill.Style = esriSimpleFillStyle.esriSFSSolid;
                    //Manage outline
                    if (isOutlineNullColor)
                    {
                        //Apply black outline 
                        newSimpleFill.Outline = inOutline;
                    }
                    else
                    {
                        //Keep wanted outline
                        newSimpleFill.Outline = iFillSymbol.Outline;
                    }

                    intShapeElement.Symbol = newSimpleFill;

                    return inElement;
                }

            }
            else if (fillSymbolDico.ContainsKey(style) && !isSimpleFill)
            {
                intShapeElement.Symbol = fillSymbolDico[style] as IFillSymbol;
                return inElement;
            }
            else
            {
                //Apply missing style
                ISimpleFillSymbol missingFillSymbol = Services.Symbols.GetMissingPolygonSymbol();
                missingFillSymbol.Style = esriSimpleFillStyle.esriSFSSolid;
                missingFillSymbol.Outline = inOutline;
                missingFillSymbol.Outline.Color = inOutline.Color;

                intShapeElement.Symbol = missingFillSymbol;
                return inElement;
            }



        }

        /// <summary>
        /// Will add a picture element with given colored added above it as transparent.
        /// </summary>
        /// <param name="inElement"></param>
        /// <param name="inColor"></param>
        public IElement SetPolygonDEM(IElement inElement, Color inColor, Tuple<double, double> inAnchor = null)
        {

            //Variables
            Services.ImageProcessing imProcessing = new Services.ImageProcessing();

            //Calculate DEM transparency
            int demtransparency = 178; //178/255 is 70% opacity
            if (otherComponents.ContainsKey(Constants.ImageConfiguration.demTransparencyNameJSON))
            {
                int.TryParse(otherComponents[Constants.ImageConfiguration.demTransparencyNameJSON], out demtransparency);
                double opacityConversion = Math.Round(((double)demtransparency / 100.0) * 255.0);
                demtransparency = Convert.ToInt16(opacityConversion);
            }
            

            //Validate DEM picture existance and get path
            string demImagePath = ValidateDEMPictureExistance();

            //Init new image object
            Image demImage = Image.FromFile(demImagePath);

            //Build path to new mono colored image
            string outputFolderName = System.IO.Path.Combine(Dictionaries.Constants.ESRI.defaultArcGISFolderName, Dictionaries.Constants.Namespaces.mainNamespace + " " + ThisAddIn.Version.ToString());
            string outputFolderPath = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments), outputFolderName);
            string monoColoredName = Constants.ImageConfiguration.monoColoredImageNamePrefix + demtransparency.ToString() + "_" + inColor.A.ToString() + "_" + inColor.R.ToString() + "_" + inColor.G.ToString() + "_" + inColor.B.ToString() + ".png";
            string demColoredName = Constants.Graphics.legendBoxDEM + demtransparency.ToString() + "_" + inColor.A.ToString() + "_" + inColor.R.ToString() + "_" + inColor.G.ToString() + "_" + inColor.B.ToString() + ".png";
            string monoColoredPath = System.IO.Path.Combine(outputFolderPath, monoColoredName);
            string demColoredPath = System.IO.Path.Combine(outputFolderPath, demColoredName);

            //Process and a get a copy of new mono colored image
            if (!System.IO.File.Exists(monoColoredPath))
            {
                imProcessing.CreateMonoColorFromImageCopy(demImage, inColor, monoColoredPath, demtransparency);
            }
            
            //Create bitmaps from original dem image and mono colored one
            Image monoColoredImage = Image.FromFile(monoColoredPath);
            Bitmap monoColoredBitmap = new Bitmap(monoColoredImage);
            Bitmap originalBitmap = new Bitmap(demImage);

            //Overlap both bitmaps
            Bitmap overlapImage = new Bitmap(monoColoredImage.Width, monoColoredImage.Height, PixelFormat.Format32bppPArgb);
            Graphics overlapGraphic = Graphics.FromImage(overlapImage);
            overlapGraphic.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceOver;
            overlapGraphic.DrawImage(originalBitmap, 0, 0);
            overlapGraphic.DrawImage(monoColoredBitmap, 0, 0);

            //Save result
            overlapImage.Save(demColoredPath, ImageFormat.Png);

            //Get graphic element for DEM
            IElement demElement = Services.ObjectManagement.CopyInputObject(templateGraphicDico[Constants.Graphics.legendBoxDEM]) as IElement;

            //Set path
            IPictureElement demPictureElement = demElement as IPictureElement;
            demPictureElement.ImportPictureFromFile(demColoredPath);
            demPictureElement.SavePictureInDocument = true;

            //Add to document
            currentDoc.ActiveView.GraphicsContainer.AddElement(demElement, 0);
            IPageLayout currentLayout = currentDoc.ActiveView as IPageLayout;
            IGraphicsContainerSelect currentGrapSelection = currentLayout as IGraphicsContainerSelect;

            currentDoc.ActivatedView.GraphicsContainer.SendToBack(currentGrapSelection.SelectedElements);
            legendElementList.Add(demElement);

            //Move if needed
            if (inAnchor != null)
            {
                //TODO find why we must substract 10 else the DEM picture is offset to the right by 10 if used inside a grouped element
                Tuple<double, double> newDEMAnchor = new Tuple<double, double>(inAnchor.Item1, inAnchor.Item2);
                SetRectangularPolygonFromAnchorType(demElement, newDEMAnchor);

            }

            return demElement;

        }

        /// <summary>
        /// Will symbolize thin unit by changing inner line symbol and match color to wanted map unit color.
        /// </summary>
        /// <param name="inThinUnitElement"></param>
        /// <param name="styleLineColorCode"></param>
        /// <param name="styleLineSymbolCode"></param>
        /// <returns></returns>
        public IElement SetThinUnitSymbol(IElement inThinUnitElement, string styleLineColorCode, string styleLineSymbolCode)
        {
            IGroupElement inGroupElement = inThinUnitElement as IGroupElement;
            if (inGroupElement != null)
            {
                //Check geometry of inner elements, if it's all lines
                for (int el = 0; el < inGroupElement.ElementCount; el++)
                {
                    
                    IElementProperties currentShapeProp = inGroupElement.Element[el] as IElementProperties;

                    if (currentShapeProp.Name == Constants.Graphics.subUnitLine)
                    {
                        ILineElement currentShapeElement = inGroupElement.Element[el] as ILineElement;

                        double currentLineWidth = currentShapeElement.Symbol.Width ;

                        //Set line style
                        if (lineSymbolDico.ContainsKey(styleLineSymbolCode))
                        {
                            currentShapeElement.Symbol = Services.ObjectManagement.CopyInputObject(lineSymbolDico[styleLineSymbolCode]) as ILineSymbol;

                        }
                        else
                        {
                            //Apply missing style
                            ISimpleLineSymbol missingFillSymbol = Services.Symbols.GetMissingLineSymbol();
                            currentShapeElement.Symbol = missingFillSymbol;
                        }

                        //Set line color 
                        if (styleLineColorCode != null && fillSymbolDico.ContainsKey(styleLineColorCode))
                        {
                            string symbolType = string.Empty;
                            ILineSymbol symbolCopy = Services.ObjectManagement.CopyInputObject(currentShapeElement.Symbol) as ILineSymbol;
                            symbolCopy.Color = Services.Symbols.GetPolygonSymbolColor(fillSymbolDico[styleLineColorCode] as ISymbol, out symbolType);
                            currentShapeElement.Symbol = symbolCopy;
                        }
                        else
                        {
                            //Apply missing style
                            ISimpleLineSymbol missingFillSymbol = Services.Symbols.GetMissingLineSymbol();
                            currentShapeElement.Symbol = missingFillSymbol;
                        }
                    }
                    else
                    {
                        //Set background color to white
                        IFillShapeElement currentPolyElement = inGroupElement.Element[el] as IFillShapeElement;
                        ISimpleFillSymbol whiteBackground = Services.Symbols.GetWhiteBackgrounFillSymbol(currentPolyElement.Symbol.Outline);

                        currentPolyElement.Symbol = whiteBackground;
                    }

                }
            }

            return inThinUnitElement;
        }

        /// <summary>
        /// Will set the marker fill of overlay symbol with a predefined color
        /// </summary>
        /// <returns></returns>
        public IElement SetOverlayFillColor(IElement inElement, string style, string style2)
        {
            //Symbolize if symbol can be found in style file
            IFillShapeElement inShapeElement = inElement as IFillShapeElement;

            if (fillSymbolDico.ContainsKey(style) && fillSymbolDico.ContainsKey(style2))
            {
                //Get symbol type and color
                string symbolTypeName = string.Empty;
                IColor symbolColor = Services.Symbols.GetPolygonSymbolColor(fillSymbolDico[style2] as ISymbol, out symbolTypeName);

                //Get symbol itself
                IMultiLayerFillSymbol inMarkerFill = fillSymbolDico[style] as IMultiLayerFillSymbol;

                //Create new symbol and apply, else it won't update...
                IMultiLayerFillSymbol newMarkerFill = new MultiLayerFillSymbol();

                //Add as many layer there is needed to the symbol.
                for (int i = 0; i < inMarkerFill.LayerCount; i++)
                {
                    newMarkerFill.AddLayer(inMarkerFill.Layer[i]);
                }

                //Set color at the end 
                newMarkerFill.Color = symbolColor;

                inShapeElement.Symbol = newMarkerFill;
                return inElement;
            }
            else
            {
                //Apply missing style
                IMarkerFillSymbol missingMarkerFillSymbol = Services.Symbols.GetMissingOverlaySymbol();

                inShapeElement.Symbol = missingMarkerFillSymbol;
                return inElement;
            }
        }
        #endregion

    }
}
