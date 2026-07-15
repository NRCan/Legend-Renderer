using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Mapping.Symbology;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using GSCLegendRendererPro.Models;
using GSCLegendRendererPro.Services;
using GSCLegendRendererPro.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;
using static ArcGIS.Desktop.Internal.GeoProcessing.Controls.rtbEditor;
using static GSCLegendRendererPro.Utilities.Layers;
using static System.Net.Mime.MediaTypeNames;
using Envelope = ArcGIS.Core.Geometry.Envelope;
using Field = ArcGIS.Core.Data.Field;
using LinearUnit = ArcGIS.Core.Geometry.LinearUnit;
using Table = ArcGIS.Core.Data.Table;
using TextElement = ArcGIS.Desktop.Layouts.TextElement;

namespace GSCLegendRendererPro.ProWindows
{
    public class Form_Legend_RendererViewModel: PropertyChangedBase
    {
        #region INIT

        //UI
        private Form_Legend_Renderer _view = null;
        private object _lock = new(); //For locking the threads to update obs. collection
        private Uri _legendTableWorkspaceUri = null;

        //JSON
        public Configuration_Y_Spacings ySpacings { get; set; }
        public Configuration_X_Spacings xSpacings { get; set; }
        public Configuration_Other otherComponents { get; set; }

        //STYLING
        public StyleProjectItem gscStyle { get; set; }
        public Dictionary<string, SymbolStyleItem> fillSymbolDico { get; set; } //Will hold symbol name (01.01.0l) and it's associate style object
        public Dictionary<string, SymbolStyleItem> lineSymbolDico { get; set; } //Will hold symbol name (01.01.0l) and it's associate style object
        public Dictionary<string, SymbolStyleItem> markerSymbolDico { get; set; } //Will hold symbol name (01.01.0l) and it's associate style object
        public Dictionary<string, SymbolStyleItem> textSymbolDico { get; set; } //Will hold symbol name and style object

        //FONT
        public Dictionary<int, double> arialCharactersWidth { get; set; } //Will be used to calculate text box height based on total lenght of characters

        //LAYOUT
        Layout pPage = null;
        LayoutView pLayoutView = null;

        //GRAPHICS PROCESSING
        public Dictionary<string, IElement> templateGraphicDico { get; set; }
        public Element parentElement = null; //Will be used to keep parent element that has embedded children
        public double originalYSpacing = 0;
        public double ySpacing = 0;//Keep track of Y spacing
        public double xSpacing = 0; //Keep track of X spacing
        public bool firstIterationBreaker = true;
        public Element lastElement = null;
        public string lastElementType = string.Empty;
        public int lastColumn = 1;
        public Element waitingLeftBracket = null; //Will be used to move in Y axis an added left bracket that needs to know ySpacing for it's horizontal brother element.
        public Element upLeftBracket = null; //Will be used to complete left bracket when end-point is reached
        public Element waitingCenterLeftBracket = null; //Will be used to move bracket annotation when full bracket has been completed.
        public Element annotationBracket = null; //Will be used to set text first and then move it.
        public Element waitingRightBracket = null; //Will be used to move in XY axis an added right bracket
        public Element upRightBracket = null; //Will be used to complete right bracket when end-point is reached.
        public Element waitingCenterRightBracket = null; //Wil be used to move bracket associated map unit when full bracket has been completed
        public int howManyRightBrackets = 0; //Will be used to recalculate x spacing in case more columns are asked by user and that some right brackets are also found
        public Tuple<Element, Element, Element, Element> bracketMapUnit = new Tuple<Element, Element, Element, Element>(null, null, null, null); //Will be used to keep unit box for bracket and replace it at the right anchor when bracket is done drawing.       
        public Tuple<double, double> anchorPoint = new Tuple<double, double>(0, 0);
        public Tuple<double, double> anchorPointParent = new Tuple<double, double>(0, 0);
        //public List<string> heading5Text = new List<string>(); //Init
        public double currentIteration = 0.0; //Will be used if user has forgot to enter an order.
        public bool nullOrderBreaker = false; //Will be used to show error message to user if null values are found, but only once.
        public double legendYLowerBound = 0.0; //Will be used to keep track of the lower bound of the legend element in case it's a CGM map.
        public int currentColumn = 1;
        double currentOrder = 0;
        string currentStyle1 = string.Empty;
        string currentStyle2 = string.Empty;
        string currentLabel1 = string.Empty;
        string currentLabel2 = string.Empty;
        string currentDescription = string.Empty;
        string currentHeading = string.Empty;
        string currentElementName = string.Empty;
        string currentLabel1Style = string.Empty;
        string currentLabel2Style = string.Empty;
        public Element currentElementObject { get; set; }
        List<Element> legendElementList = new List<Element>(); //Will hold all legend items to group them at the end of the process.

        //TABLE
        int elementFieldIndex = -1;
        int orderFieldIndex = -1;
        int style1FieldIndex = -1;
        int style2FieldIndex = -1;
        int labelFieldIndex = -1;
        int label2FieldIndex = -1;
        int descriptionFieldIndex = -1;
        int headingFieldIndex = -1;
        int columnFieldIndex = -1;
        int label1StyleFieldIndex = -1;
        int label2StyleFieldIndex = -1;

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

        #endregion

        #region UI PROPERTIES

        //Layer controls
        private ObservableCollection<LayerDisplay> _legendLayers = new();
        public ObservableCollection<LayerDisplay> LegendLayers
        {
            get { return _legendLayers; }
        }

        private int _legendSelectedLayerIndex = -1;
        public int LegendSelectedLayerIndex
        {
            get { return _legendSelectedLayerIndex; }
            set
            {
                SetProperty(ref _legendSelectedLayerIndex, value, () => _legendSelectedLayerIndex);

                //On index change, update the legend table field comboboxes
                FillFieldComboboxes();
            }
        }


        private ObservableCollection<CustomCombobox> _legendOrder = new();
        public ObservableCollection<CustomCombobox> LegendOrder
        {
            get { return _legendOrder; }
        }
        private int _legendSelectedOrderIndex = -1;
        public int LegendSelectedOrderIndex
        {
            get { return _legendSelectedOrderIndex; }
            set
            {
                SetProperty(ref _legendSelectedOrderIndex, value, () => _legendSelectedOrderIndex);
            }
        }


        private ObservableCollection<CustomCombobox> _legendColumn = new();
        public ObservableCollection<CustomCombobox> LegendColumn
        {
            get { return _legendColumn; }
        }
        private int _legendSelectedColumnIndex = -1;
        public int LegendSelectedColumnIndex
        {
            get { return _legendSelectedColumnIndex; }
            set
            {
                SetProperty(ref _legendSelectedColumnIndex, value, () => _legendSelectedColumnIndex);
            }
        }

        private ObservableCollection<CustomCombobox> _legendElement = new();
        public ObservableCollection<CustomCombobox> LegendElement
        {
            get { return _legendElement; }
        }
        private int _legendSelectedElementIndex = -1;
        public int LegendSelectedElementIndex
        {
            get { return _legendSelectedElementIndex; }
            set
            {
                SetProperty(ref _legendSelectedElementIndex, value, () => _legendSelectedElementIndex);
            }
        }

        private ObservableCollection<CustomCombobox> _legendStyle1 = new();
        public ObservableCollection<CustomCombobox> LegendStyle1
        {
            get { return _legendStyle1; }
        }
        private int _legendSelectedStyle1Index = -1;
        public int LegendSelectedStyle1Index
        {
            get { return _legendSelectedStyle1Index; }
            set
            {
                SetProperty(ref _legendSelectedStyle1Index, value, () => _legendSelectedStyle1Index);
            }
        }

        private ObservableCollection<CustomCombobox> _legendStyle2 = new();
        public ObservableCollection<CustomCombobox> LegendStyle2
        {
            get { return _legendStyle2; }
        }
        private int _legendSelectedStyle2Index = -1;
        public int LegendSelectedStyle2Index
        {
            get { return _legendSelectedStyle2Index; }
            set
            {
                SetProperty(ref _legendSelectedStyle2Index, value, () => _legendSelectedStyle2Index);
            }
        }

        private ObservableCollection<CustomCombobox> _legendLabel1 = new();
        public ObservableCollection<CustomCombobox> LegendLabel1
        {
            get { return _legendLabel1; }
        }
        private int _legendSelectedLabel1Index = -1;
        public int LegendSelectedLabel1Index
        {
            get { return _legendSelectedLabel1Index; }
            set
            {
                SetProperty(ref _legendSelectedLabel1Index, value, () => _legendSelectedLabel1Index);
            }
        }

        private ObservableCollection<CustomCombobox> _legendLabel1Style = new();
        public ObservableCollection<CustomCombobox> LegendLabel1Style
        {
            get { return _legendLabel1Style; }
        }
        private int _legendSelectedLabel1StyleIndex = -1;
        public int LegendSelectedLabel1StyleIndex
        {
            get { return _legendSelectedLabel1StyleIndex; }
            set
            {
                SetProperty(ref _legendSelectedLabel1StyleIndex, value, () => _legendSelectedLabel1StyleIndex);
            }
        }

        private ObservableCollection<CustomCombobox> _legendLabel2 = new();
        public ObservableCollection<CustomCombobox> LegendLabel2
        {
            get { return _legendLabel2; }
        }
        private int _legendSelectedLabel2Index = -1;
        public int LegendSelectedLabel2Index
        {
            get { return _legendSelectedLabel2Index; }
            set
            {
                SetProperty(ref _legendSelectedLabel2Index, value, () => _legendSelectedLabel2Index);
            }
        }

        private ObservableCollection<CustomCombobox> _legendLabel2Style = new();
        public ObservableCollection<CustomCombobox> LegendLabel2Style
        {
            get { return _legendLabel2Style; }
        }
        private int _legendSelectedLabel2StyleIndex = -1;
        public int LegendSelectedLabel2StyleIndex
        {
            get { return _legendSelectedLabel2StyleIndex; }
            set
            {
                SetProperty(ref _legendSelectedLabel2StyleIndex, value, () => _legendSelectedLabel2StyleIndex);
            }
        }

        private ObservableCollection<CustomCombobox> _legendHeading = new();
        public ObservableCollection<CustomCombobox> LegendHeading
        {
            get { return _legendHeading; }
        }
        private int _legendSelectedHeadingIndex = -1;
        public int LegendSelectedHeadingIndex
        {
            get { return _legendSelectedHeadingIndex; }
            set
            {
                SetProperty(ref _legendSelectedHeadingIndex, value, () => _legendSelectedHeadingIndex);
            }
        }

        private ObservableCollection<CustomCombobox> _legendDescription = new();
        public ObservableCollection<CustomCombobox> LegendDescription
        {
            get { return _legendDescription; }
        }
        private int _legendSelectedDescriptionIndex = -1;
        public int LegendSelectedDescriptionIndex
        {
            get { return _legendSelectedDescriptionIndex; }
            set
            {
                SetProperty(ref _legendSelectedDescriptionIndex, value, () => _legendSelectedDescriptionIndex);
            }
        }

        private bool _legendDEM = false;
        public bool LegendDEM
        {
            get { return _legendDEM; }
            set
            {
                SetProperty(ref _legendDEM, value, () => _legendDEM);
            }
        }

        private bool _legendAutoCalculateColumn = false;
        public bool LegendAutoCalculateColumn
        {
            get { return _legendAutoCalculateColumn; }
            set
            {
                SetProperty(ref _legendAutoCalculateColumn, value, () => _legendAutoCalculateColumn);
            }
        }

        private string _warningMessage = string.Empty;
        public string WarningMessage
        {
            get { return _warningMessage; }
            set
            {
                SetProperty(ref _warningMessage, value, () => _warningMessage);
            }
        }

        #endregion

        #region RELAYS

        private ICommand _runTool = null;
        public ICommand RunTool
        {
            get
            {
                if (_runTool == null)
                {
                    _runTool = new RelayCommand(() => CreateLegend(), () => true);
                }
                return _runTool;
            }
        }

        #endregion

        public Form_Legend_RendererViewModel(Form_Legend_Renderer view)
        {
            //Init as obs. collection the comboboxes
            BindingOperations.EnableCollectionSynchronization(_legendLayers, _lock);

            _view = view;

            //Init some components
            UpdateTableCombobox();
        }

        #region METHODS

        /// <summary>
        /// Will fill the layer combobox with all standalone tables in the map
        /// </summary>
        public async void UpdateTableCombobox()
        {
            //Init some components
            _legendLayers.Clear();
            Layers layerService = new Layers();
            await layerService.UpdateTableViewCombobox(_legendLayers, nameof(LegendLayers), _legendSelectedLayerIndex, nameof(LegendSelectedLayerIndex));

        }

        /// <summary>
        /// Will load the legend field items from the selected legend table
        /// </summary>
        public async void FillFieldComboboxes()
        {
            //Clear warnings
            _warningMessage = string.Empty;
            NotifyPropertyChanged(nameof(WarningMessage));

            ObservableCollection<CustomCombobox> fieldList = new ObservableCollection<CustomCombobox>();

            try
            {
                //Reset
                _legendColumn.Clear();
                _legendOrder.Clear();
                _legendDescription.Clear();
                _legendElement.Clear();
                _legendHeading.Clear();
                _legendLabel1.Clear();
                _legendLabel1Style.Clear();
                _legendLabel2Style.Clear();
                _legendStyle1.Clear();
                _legendStyle2.Clear();
                _legendLabel2.Clear();

                NotifyPropertyChanged(nameof(LegendColumn));
                NotifyPropertyChanged(nameof(LegendOrder));
                NotifyPropertyChanged(nameof(LegendDescription));
                NotifyPropertyChanged(nameof(LegendElement));
                NotifyPropertyChanged(nameof(LegendHeading));
                NotifyPropertyChanged(nameof(LegendLabel1));
                NotifyPropertyChanged(nameof(LegendLabel2));
                NotifyPropertyChanged(nameof(LegendStyle1));
                NotifyPropertyChanged(nameof(LegendStyle2));
                NotifyPropertyChanged(nameof(LegendLabel1Style));
                NotifyPropertyChanged(nameof(LegendLabel2Style));

                if (_legendSelectedLayerIndex != -1)
                {
                    await QueuedTask.Run(() =>
                    {

                        StandaloneTable legendSTable = LegendLayers[LegendSelectedLayerIndex].STable;
                        if (legendSTable != null)
                        {
                            _legendTableWorkspaceUri = Utilities.Workspace.GetWorkspacePath(legendSTable);

                            if (_legendTableWorkspaceUri != null)
                            {

                                //Legend table records 
                                using (Table legendTable = legendSTable.GetTable())
                                {
                                    if (legendTable != null)
                                    {
                                        List<Field> legendFields = legendTable.GetDefinition().GetFields().ToList();
                                        if (legendFields != null && legendFields.Count > 0)
                                        {
                                            foreach (Field field in legendFields)
                                            {
                                                CustomCombobox fieldItem = new CustomCombobox
                                                {
                                                    Name = field.Name,
                                                    Value = field.Name
                                                };

                                                fieldList.Add(fieldItem);
                                            }

                                        }
                                        else
                                        {
                                            _warningMessage = Properties.Resources.ErrorNoFields;
                                            NotifyPropertyChanged(nameof(WarningMessage));
                                        }
                                    }
                                }

                            }
                            else
                            {
                                _warningMessage = Properties.Resources.ErrorWrongData;
                                NotifyPropertyChanged(nameof(WarningMessage));
                            }

                        }
                    });

                    if (fieldList.Count() > 0)
                    {
                        _legendColumn = fieldList;
                        _legendOrder = fieldList;
                        _legendDescription = fieldList;
                        _legendElement = fieldList;
                        _legendHeading = fieldList;
                        _legendLabel1 = fieldList;
                        _legendLabel1Style = fieldList;
                        _legendLabel2 = fieldList;
                        _legendLabel2Style = fieldList;
                        _legendStyle1 = fieldList;
                        _legendStyle2 = fieldList;

                        NotifyPropertyChanged(nameof(LegendColumn));
                        NotifyPropertyChanged(nameof(LegendOrder));
                        NotifyPropertyChanged(nameof(LegendDescription));
                        NotifyPropertyChanged(nameof(LegendElement));
                        NotifyPropertyChanged(nameof(LegendHeading));
                        NotifyPropertyChanged(nameof(LegendLabel1));
                        NotifyPropertyChanged(nameof(LegendLabel2));
                        NotifyPropertyChanged(nameof(LegendStyle1));
                        NotifyPropertyChanged(nameof(LegendStyle2));
                        NotifyPropertyChanged(nameof(LegendLabel1Style));
                        NotifyPropertyChanged(nameof(LegendLabel2Style));

                        //Preselect values if matching fields are found
                        SelectFields();
                    }
                }
            }
            catch (Exception FillFieldComboboxesException)
            {
                new ErrorService(FillFieldComboboxesException).WriteToFile();
                _warningMessage = FillFieldComboboxesException.Message;
                NotifyPropertyChanged(nameof(WarningMessage));
            }


        }

        /// <summary>
        /// Will automatically select field if they match default schema for legend table
        /// </summary>
        private void SelectFields()
        {
            CustomCombobox fieldColumn = _legendColumn.Where(x => x.Name.Contains(Constants.LegendTable.legendColumnField, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            CustomCombobox fieldDescription = _legendDescription.Where(x => x.Name.Contains(Constants.LegendTable.legendDescriptionField, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            CustomCombobox fieldElement = _legendElement.Where(x => x.Name == Constants.LegendTable.legendElementField).FirstOrDefault();
            CustomCombobox fieldHeading = _legendHeading.Where(x => x.Name.Contains(Constants.LegendTable.legendHeadingField, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            CustomCombobox fieldLabel1 = _legendLabel1.Where(x => x.Name.Contains(Constants.LegendTable.legendLabel1Field, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            CustomCombobox fieldLabel1Style = _legendLabel1Style.Where(x => x.Name.Contains(Constants.LegendTable.legendLabel1StyleField, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            CustomCombobox fieldLabel2 = _legendLabel2.Where(x => x.Name.Contains(Constants.LegendTable.legendLabel2Field, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            CustomCombobox fieldLabel2Style = _legendLabel2Style.Where(x => x.Name.Contains(Constants.LegendTable.legendLabel2StyleField, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            CustomCombobox fieldOrder = _legendOrder.Where(x => x.Name.Contains(Constants.LegendTable.legendOrderField, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            CustomCombobox fieldStyle1 = _legendStyle1.Where(x => x.Name.Contains(Constants.LegendTable.legendStyle1Field, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
            CustomCombobox fieldStyle2 = _legendStyle2.Where(x => x.Name.Contains(Constants.LegendTable.legendStyle2Field, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();

            if (fieldColumn != null)
            {
                int columnIndex = _legendColumn.IndexOf(fieldColumn);
                _legendSelectedColumnIndex = columnIndex;
                NotifyPropertyChanged(nameof(LegendSelectedColumnIndex));
            }

            if (fieldDescription != null)
            {
                int descIndex = _legendDescription.IndexOf(fieldDescription);
                _legendSelectedDescriptionIndex = descIndex;
                NotifyPropertyChanged(nameof(LegendSelectedDescriptionIndex));
            }

            if (fieldElement != null)
            {
                int elementIndex = _legendElement.IndexOf(fieldElement);
                _legendSelectedElementIndex = elementIndex;
                NotifyPropertyChanged(nameof(LegendSelectedElementIndex));
            }

            if (fieldHeading != null)
            {
                int headingIndex = _legendHeading.IndexOf(fieldHeading);
                _legendSelectedHeadingIndex = headingIndex;
                NotifyPropertyChanged(nameof(LegendSelectedHeadingIndex));
            }

            if (fieldLabel1 != null)
            {
                int label1Index = _legendLabel1.IndexOf(fieldLabel1);
                _legendSelectedLabel1Index = label1Index;
                NotifyPropertyChanged(nameof(LegendSelectedLabel1Index));
            }

            if (fieldLabel2 != null)
            {
                int label2Index = _legendLabel2.IndexOf(fieldLabel2);
                _legendSelectedLabel2Index = label2Index;
                NotifyPropertyChanged(nameof(LegendSelectedLabel2Index));
            }

            if (fieldLabel1Style != null)
            {
                int label1StyleIndex = _legendLabel1Style.IndexOf(fieldLabel1Style);
                _legendSelectedLabel1StyleIndex = label1StyleIndex;
                NotifyPropertyChanged(nameof(LegendSelectedLabel1StyleIndex));
            }

            if (fieldLabel2Style != null)
            {
                int label2StyleIndex = _legendLabel2Style.IndexOf(fieldLabel2Style);
                _legendSelectedLabel2StyleIndex = label2StyleIndex;
                NotifyPropertyChanged(nameof(LegendSelectedLabel2StyleIndex));
            }

            if (fieldOrder != null)
            {
                int orderIndex = _legendOrder.IndexOf(fieldOrder);
                _legendSelectedOrderIndex = orderIndex;
                NotifyPropertyChanged(nameof(LegendSelectedOrderIndex));
            }

            if (fieldStyle1 != null)
            {
                int style1Index = _legendStyle1.IndexOf(fieldStyle1);
                _legendSelectedStyle1Index = style1Index;
                NotifyPropertyChanged(nameof(LegendSelectedStyle1Index));
            }

            if (fieldStyle2 != null)
            {
                int style2Index = _legendStyle2.IndexOf(fieldStyle2);
                _legendSelectedStyle2Index = style2Index;
                NotifyPropertyChanged(nameof(LegendSelectedStyle2Index));
            }
        }

        /// <summary>
        /// Will create the legend
        /// </summary>
        private async void CreateLegend()
        {
            _warningMessage = string.Empty;
            NotifyPropertyChanged(nameof(WarningMessage));

            try
            {
                if (_legendSelectedLayerIndex != -1)
                {
                    //Setup procedures
                    bool validateUIControls = ValidateFieldControls();
                    bool setupAddinCleared = await SetupAddinEnvironment();
                    bool setupLayoutCleared = await SetupLayoutAndGraphics();

                    if (validateUIControls && setupAddinCleared && setupLayoutCleared && _legendLayers[_legendSelectedLayerIndex].STable != null)
                    {
                        //Run on UI thread
                        await QueuedTask.Run(async () =>
                        {
                            //Prepare legend table and query filter to sort it on user predefined order
                            using (Table legendTable = _legendLayers[_legendSelectedLayerIndex].STable.GetTable())
                            {

                                ArcGIS.Core.Data.QueryFilter itemFilter = new ArcGIS.Core.Data.QueryFilter();
                                if (_legendSelectedOrderIndex != -1)
                                {
                                    itemFilter = new ArcGIS.Core.Data.QueryFilter
                                    {
                                        PostfixClause = $"ORDER BY {_legendOrder[_legendSelectedOrderIndex].Value} ASC"
                                    };
                                }

                                using (RowCursor legendCursor = legendTable.Search(itemFilter, false))
                                {
                                    //Get fields indexes
                                    if (GetFieldIndexesFromTable(legendCursor))
                                    {
                                        while (legendCursor.MoveNext())
                                        {
                                            using (Row legendRow = legendCursor.Current)
                                            {
                                                //Current row information collecting
                                                currentIteration = currentIteration + 1.0;
                                                //currentOrder = currentIteration;
                                                await GatherCurrentRowInformation(legendRow);

                                                #region GRAPHIC PREPARATION

                                                await CleanupDescription();

                                                //Get related graphic from template layout dictionary
                                                if (templateGraphicDico.ContainsKey(currentElementName))
                                                {
                                                    currentElementObject = CopyElementObject(templateGraphicDico[currentElementName] as Element, currentOrder.ToString());
                                                }
                                                else
                                                {
                                                    throw new Exception(string.Format(Properties.Resources.ErrorMissingLegendElement, currentElementName));
                                                }

                                                //Manage null order
                                                await ManageNullOrder(legendRow);

                                                //Set heading5 trigger for special style symbols
                                                await Heading5Preparation(legendRow);

                                                //Get spacings in x and Y for current row
                                                ySpacing = GetYSpacing(lastElement, lastElementType, currentElementName, anchorPoint.Item2);
                                                xSpacing = GetXSpacing(currentElementName);

                                                //Manage column change and spacing
                                                await ManageColumn(legendRow);

                                                #endregion

                                                #region GRAPHIC HANDLING

                                                await AddHeading(legendRow);

                                                #endregion

                                                #region FINALIZE

                                                //ElementFactory.Instance.CreateGroupElement(pPage, legendElementList);
                                                

                                                #endregion
                                            }
                                        }
                                    }  
                                }

                            }
                        });

                        if (_warningMessage == string.Empty)
                        {
                            //Show notication success
                            FrameworkApplication.AddNotification(new Notification()
                            {
                                Title = Properties.Resources.FormRendererTitle,
                                Message = Properties.Resources.GenericMessageCompleted,
                                ImageSource = System.Windows.Application.Current.Resources["Success_Toast48"] as ImageSource
                            });

                            //Close window
                            _view.Close();
                        }
                    }
                    else
                    {
                        _warningMessage = Properties.Resources.ErrorInvalidControls;
                        NotifyPropertyChanged(nameof(WarningMessage));
                    }
                }
                else
                {
                    _warningMessage = Properties.Resources.ErrorMissingTable;
                    NotifyPropertyChanged(nameof(WarningMessage));
                }
            }
            catch (Exception CreateLegendException)
            {
                new ErrorService(CreateLegendException).WriteToFile();
            }
        }

        /// <summary>
        /// Will initiate a setup procedure regarding the addin and its environment of work, asset files existing and copied locally,
        /// user or default style file loaded and json configurations deserialized and loaded in their proper models.
        /// </summary>
        /// <returns></returns>
        public async Task<bool> SetupAddinEnvironment()
        {
            try
            {
                //Validate configuration files and default style file
                await UserConfigurationService.ValidateAssetsExistance();

                //Deserialize the configuration files
                ySpacings = await UserConfigurationService.DeserializeJsonFile<Configuration_Y_Spacings>(System.IO.Path.Combine(Properties.Settings.Default.WorkingEnvironmentPath, Constants.Assets.jsonYSpacingEmbeddedFile));
                xSpacings = await UserConfigurationService.DeserializeJsonFile<Configuration_X_Spacings>(System.IO.Path.Combine(Properties.Settings.Default.WorkingEnvironmentPath, Constants.Assets.jsonXSpacingEmbeddedFile));
                otherComponents = await UserConfigurationService.DeserializeJsonFile<Configuration_Other>(System.IO.Path.Combine(Properties.Settings.Default.WorkingEnvironmentPath, Constants.Assets.jsonStyleFontsOtherEmbeddedFile));

                //Validate the style file loaded in map
                if (otherComponents != null)
                {
                    await QueuedTask.Run(async () =>
                    {
                        List<StyleProjectItem> styleItems = Project.Current.GetItems<StyleProjectItem>().Where(x => x.Name == otherComponents.GEOLOGY_STYLE_NAME).ToList();

                        //Add embeded style if user hasn't got it already loaded in the project
                        if ((styleItems == null || styleItems.Count() == 0 ) && otherComponents.GEOLOGY_STYLE_NAME == Constants.Assets.gscSymbolStandardStyle.Split(".")[0])
                        {
                            //Load up the style coming from the default setup
                            string defaultStylePath = System.IO.Path.Combine(Properties.Settings.Default.WorkingEnvironmentPath, otherComponents.GEOLOGY_STYLE_NAME + ".stylx");
                            Project.Current.AddStyle(defaultStylePath);
                        }
                        gscStyle = Project.Current.GetItems<StyleProjectItem>().Where(x => x.Name == otherComponents.GEOLOGY_STYLE_NAME).FirstOrDefault();

                        if (gscStyle == null)
                        {
                            throw new Exception(Properties.Resources.ErrorStyleNotFound);
                        }
                    });


                    return true;
                }
                else
                {
                    return false;
                }

            }

            catch (Exception ValidationRoundException)
            {
                new ErrorService(ValidationRoundException).WriteToFile();
                return false;
            }

        
        }

        /// <summary>
        /// Will initialized the layout page and set a couple of things like units and character widths
        /// </summary>
        /// <returns></returns>
        public async Task<bool> SetupLayoutAndGraphics()
        {
            try
            {
                //Navigate to a layout view
                await NavigateToLayoutViewAsync();

                //Get arial character widths
                arialCharactersWidth = GetArialCharacterWidth();

                //Set document units, else if it's not in mm the legend will be looking bad...
                SetPageUnits();

                //Force delay update
                //TODO: didn't find equivalent under ArcPro
                //currentDoc.DelayUpdateContents = true;

                //Get template graphics
                await GetTemplateGraphicList();

                //Get all symbols from style and keep them in a dictionary for later use
                await GetAllSymbols();

                //If within a CGM map, get the lower bound of the legend element
                legendYLowerBound = await GetCGMLegendLowerBound(Constants.YSpacings.legendEnd_Citation, Constants.Graphics.cgmCitation);

                //Get start anchor point (in case from CGM template)
                await GetAnchorPointStart();

                //Set default Y spacing at start (incase from CGM template)
                originalYSpacing = anchorPoint.Item2; //Synchronise with initial calculate anchor.

                //Initiate width and height of some graphics
                await InitWidthHeight();

                //Heading 5 initialization
                heading5Text = new List<string>();

                return true;
            }
            catch (Exception SetupLayoutAndGraphicsException)
            {
                new ErrorService(SetupLayoutAndGraphicsException).WriteToFile();
                return false;
            }

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
        /// If user is not already in a layout view, the tool will navigate to the first one in the available list
        /// As opposed to ArcMap, Pro allows multiple layouts per project, this method will just open the first one found
        /// </summary>
        public async Task<bool> NavigateToLayoutViewAsync()
        {
            try
            {
                bool paneFoundAndActivated = false;
                Pane currentActivePane = ProApp.Panes.ActivePane;
                if (currentActivePane != null)
                {
                    ILayoutPane layoutPane = currentActivePane as ILayoutPane;

                    if (layoutPane == null)
                    {
                        MessageBoxResult msgBoxResult = MessageBox.Show(Properties.Resources.FormRendererMovingToLayout, Properties.Resources.GenericWarningTitle, System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Exclamation);
                        if (msgBoxResult == MessageBoxResult.OK)
                        {

                            //First case - Layout already activated
                            PaneCollection panes = ProApp.Panes;
                            foreach (Pane pane in panes)
                            {
                                ILayoutPane lPane = pane as ILayoutPane;
                                if (lPane != null)
                                {

                                    //Keep for later use
                                    pPage = lPane.LayoutView.Layout;
                                    pLayoutView = lPane.LayoutView;

                                    //Activate
                                    pane.Activate();

                                    paneFoundAndActivated = true;

                                    break;
                                }
                            }

                            //Second case - check within project layouts and open first one
                            if (!paneFoundAndActivated)
                            {
                                List<LayoutProjectItem> layouts = Project.Current.GetItems<LayoutProjectItem>().ToList();
                                if (layouts != null && layouts.Count() > 0)
                                {
                                    Layout firstFoundLayout = await QueuedTask.Run(() => {

                                        LayoutProjectItem firstLayoutItem = layouts.First();
                                        return firstLayoutItem.GetLayout();

                                    });

                                    if (firstFoundLayout != null)
                                    {
                                        //Keep for later use
                                        pPage = firstFoundLayout;

                                        //Load it up
                                        ILayoutPane pPane = await ProApp.Panes.CreateLayoutPaneAsync(firstFoundLayout);
                                        pLayoutView = pPane.LayoutView;
                                        paneFoundAndActivated = true;
                                    }

                                }
                            }


                            //Third case - Create a new default one
                            if (!paneFoundAndActivated)
                            {
                                //Create new layout on UI thread.
                                //Can't use QueuedTask on the whole method since the layout pane build needs to run on another
                                Layout lyt = await QueuedTask.Run(() =>
                                {
                                    //Default size to A2
                                    Layout newLayout = LayoutFactory.Instance.CreateLayout(420, 594, ArcGIS.Core.Geometry.LinearUnit.Millimeters);
                                    newLayout.SetName(Properties.Resources.FormRendererNewLayoutName);
                                    return newLayout;
                                });

                                //Keep for later use
                                pPage = lyt;

                                //Build the layout pane
                                ILayoutPane pPane = await ProApp.Panes.CreateLayoutPaneAsync(lyt);
                                pLayoutView = pPane.LayoutView;
                                paneFoundAndActivated = true;
                            }

                        }
                    }
                    else
                    {
                        paneFoundAndActivated = true;

                        //Keep for later use
                        pPage = layoutPane.LayoutView.Layout;
                        pLayoutView = layoutPane.LayoutView;
                    }
                }
                else
                {
                    paneFoundAndActivated = false;
                }

                return paneFoundAndActivated;
            }
            catch (Exception panelLayoutException)
            {
                new ErrorService(panelLayoutException).WriteToFile();
                _view.Close();
                return false;
            }


        }

        /// <summary>
        /// Checks that page units are only in mm. All spacings between graphics within the configuration files are in mm.
        /// in the tool internal settings
        /// </summary>
        public void SetPageUnits()
        {
            try
            {
                QueuedTask.Run(() =>
                {
                    if (pPage != null)
                    {

                        CIMPage cIMPage = pPage.GetPage();
                        if (cIMPage != null)
                        {
                            if (cIMPage.Units.Name != LinearUnit.Millimeters.Name)
                            {
                                MessageBoxResult msgBoxResult = MessageBox.Show(Properties.Resources.FormRendererPageUnitsWarning, Properties.Resources.GenericWarningTitle, System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Exclamation);
                                if (msgBoxResult == MessageBoxResult.OK)
                                {
                                    //Enforce centimeters
                                    cIMPage.Units = LinearUnit.Millimeters;
                                    pPage.SetPage(cIMPage);

                                }
                            }
                        }
                    }

                }).Wait();
            }
            catch (Exception setPageUnitsException)
            {
                new ErrorService(setPageUnitsException).WriteToFile();
                _view.Close();
            }
        }

        /// <summary>
        /// Will take from template mxd a set of graphics to use in the legend
        /// </summary>
        public async Task GetTemplateGraphicList()
        {
            try
            {
                //Validate if template mxd is available, else send copy resource to local folder
                string templateLayoutPath = System.IO.Path.Combine(Properties.Settings.Default.WorkingEnvironmentPath, Constants.Assets.pageLayoutTemplateFile);
                templateGraphicDico = new Dictionary<string, IElement>();

                await QueuedTask.Run(async () =>
                {
                    //Check if template layout already exist somewhere in the project
                    LayoutProjectItem templateLayoutItem = Project.Current.GetItems<LayoutProjectItem>().Where(x => x.Name.Contains(Constants.Assets.pageLayoutTemplateFile.Split(".")[0])).FirstOrDefault();

                    if (templateLayoutItem == null)
                    {
                        //Add it
                        IProjectItem templateProjectItem = ItemFactory.Instance.Create(templateLayoutPath) as IProjectItem;
                        bool addedTemplateLayout = Project.Current.AddItem(templateProjectItem);
                        templateLayoutItem = Project.Current.GetItems<LayoutProjectItem>().Where(x => x.Name.Contains(Constants.Assets.pageLayoutTemplateFile.Split(".")[0])).FirstOrDefault();
                    }


                    if (templateLayoutItem != null)
                    {
                        Layout templateLayout = templateLayoutItem.GetLayout();

                        if (templateLayout != null && templateLayout.Elements != null && templateLayout.Elements.Count() > 0)
                        {
                            foreach (Element templateElements in templateLayout.Elements)
                            {
                                templateGraphicDico.Add(templateElements.Name, templateElements);
                            }
                        }
                        else
                        {
                            throw new Exception(Properties.Resources.ErrorLayoutLoading);
                        }
                    }

                });

            }
            catch (Exception GetTemplateGraphicListException)
            {
                new ErrorService(GetTemplateGraphicListException).WriteToFile();
            }

        }

        /// <summary>
        /// Will fill in the appropriate symbol dictionary from a needed style class from user loaded style
        /// </summary>
        public async Task GetAllSymbols()
        {
            try
            {
                await QueuedTask.Run(() =>
                {
                    if (gscStyle != null)
                    {
                        fillSymbolDico = GetSymbols(gscStyle, StyleItemType.PolygonSymbol);
                        lineSymbolDico = GetSymbols(gscStyle, StyleItemType.LineSymbol);
                        textSymbolDico = GetSymbols(gscStyle, StyleItemType.TextSymbol);
                        markerSymbolDico = GetSymbols(gscStyle, StyleItemType.PointSymbol);
                    }
                });
            }
            catch (Exception GetAllSymbolsException)
            {
                new ErrorService(GetAllSymbolsException).WriteToFile();
            }
        }

        /// <summary>
        /// Will fill out a dictionary based on a symbol type and a style project item
        /// </summary>
        /// <param name="symbolDictionary"></param>
        /// <param name="styleProjectItem"></param>
        public Dictionary<string, SymbolStyleItem> GetSymbols(StyleProjectItem styleProjectItem, StyleItemType styleItemType)
        {
            Dictionary<string, SymbolStyleItem> symbolDictionary = new Dictionary<string, SymbolStyleItem>();
            List<SymbolStyleItem> symbolList = styleProjectItem.SearchSymbols(styleItemType, null).ToList();
            if (symbolList != null && symbolList.Count() > 0)
            {
                foreach (SymbolStyleItem ssi in symbolList)
                {
                    symbolDictionary.Add(ssi.Name, ssi);
                }
            }

            return symbolDictionary;
        }

        /// <summary>
        /// From a given anchor, will calculate the legend maximum height, 
        /// based on upper cgm citation graphic anchor. This will be used to
        /// automatically calculate the number of columns
        /// </summary>
        /// <param name="in_ySpacingWithCitation">The wanted y spacing between end of legend and the reference graphic</param>
        /// <returns></returns>
        public async Task<double> GetCGMLegendLowerBound(double in_ySpacingWithCitation, string in_referenceCGMGraphicName)
        {
            double outYBound = 0.0;

            try
            {
                if (pPage != null)
                {
                    await QueuedTask.Run(async () =>
                    {
                        Element referenceCGMElement = pPage.Elements.Where(e => e.Name == in_referenceCGMGraphicName).FirstOrDefault();

                        if (referenceCGMElement != null && referenceCGMElement.GetBounds() != null)
                        {
                            outYBound = referenceCGMElement.GetBounds().YMax + in_ySpacingWithCitation;
                        }
                    });



                }
            }
            catch (Exception GetCGMLegendLowerBoundException)
            {
                new ErrorService(GetCGMLegendLowerBoundException).WriteToFile();
            }


            return outYBound;

        }

        /// <summary>
        /// From a given table row cursor, will keep in variables the index of the field names, 
        /// in order to prevent fetching the index each time a value is needed.
        /// </summary>
        /// <param name="legendCursor"></param>
        /// <returns></returns>
        public bool GetFieldIndexesFromTable(RowCursor legendCursor)
        {
            bool allValidates = true;
             
            if (_legendSelectedElementIndex != -1 || _legendSelectedOrderIndex != -1 || _legendSelectedLabel1Index != -1 ||
                _legendSelectedLabel2Index != -1 || _legendSelectedDescriptionIndex != -1 || _legendSelectedHeadingIndex != -1 ||
                _legendSelectedColumnIndex != -1 || _legendSelectedLabel1StyleIndex != -1 || _legendSelectedLabel2StyleIndex != -1 ||
                _legendSelectedStyle1Index != -1 || _legendSelectedStyle2Index != -1)
            {
                elementFieldIndex = legendCursor.FindField(_legendElement[_legendSelectedElementIndex].Value);
                orderFieldIndex = legendCursor.FindField(_legendOrder[_legendSelectedOrderIndex].Value);
                labelFieldIndex = legendCursor.FindField(_legendLabel1[_legendSelectedLabel1Index].Value);
                label2FieldIndex = legendCursor.FindField(_legendLabel2[_legendSelectedLabel2Index].Value);
                descriptionFieldIndex = legendCursor.FindField(_legendDescription[_legendSelectedDescriptionIndex].Value);
                headingFieldIndex = legendCursor.FindField(_legendHeading[_legendSelectedHeadingIndex].Value);
                columnFieldIndex = legendCursor.FindField(_legendColumn[_legendSelectedColumnIndex].Value);
                label1StyleFieldIndex = legendCursor.FindField(_legendLabel1Style[_legendSelectedLabel1StyleIndex].Value);
                label2StyleFieldIndex = legendCursor.FindField(_legendLabel2Style[_legendSelectedLabel2StyleIndex].Value);
                style1FieldIndex = legendCursor.FindField(_legendStyle1[_legendSelectedStyle1Index].Value);
                style2FieldIndex = legendCursor.FindField(_legendStyle2[_legendSelectedStyle2Index].Value);
            }
            else
            {
                allValidates = false;
            }

            return allValidates;
        }

        /// <summary>
        /// Will validate UI controls to make sure eveyrthing is filled out properly
        /// </summary>
        /// <returns></returns>
        public bool ValidateFieldControls()
        {
            bool allValidates = true;

            if (_legendSelectedElementIndex == -1 || _legendSelectedOrderIndex == -1 || _legendSelectedLabel1Index == -1 ||
                _legendSelectedLabel2Index == -1 || _legendSelectedDescriptionIndex == -1 || _legendSelectedHeadingIndex == -1 ||
                _legendSelectedColumnIndex == -1 || _legendSelectedLabel1StyleIndex == -1 || _legendSelectedLabel2StyleIndex == -1 ||
                _legendSelectedStyle1Index == -1 || _legendSelectedStyle2Index == -1)
            {
                allValidates = false;
            }

            return allValidates;
        }

        /// <summary>
        /// Will make a copy of a given element object via cloning with CIM
        /// </summary>
        /// <param name="inputOb">The object to get a copy rom</param>
        /// <returns></returns>
        public Element CopyElementObject(Element element, string newElementName)
        {
            Element copiedElement = null;
            
            //Get CIM definition
            CIMElement cimElement = element.GetDefinition();

            //Clone - DO NOT CLONE, IT ALREADY ADDS IT TO ORIGINAL LAYOUT
            //CIMElement copiedCIMElement = CIMElement.Clone(cimElement) as CIMElement;

            //Case when the element is a group, we need to iterate through its elements and copy them one by one
            CIMGroupElement cimGroupElement = cimElement as CIMGroupElement;
            if (cimGroupElement != null)
            {
                List<Element> cimElements = new List<Element>();
                foreach (CIMElement elem in cimGroupElement.Elements)
                {
                   cimElements.Add(ElementFactory.Instance.CreateElement(pPage, elem));
                }

                //Create new object and add it to current layout
                copiedElement = ElementFactory.Instance.CreateGroupElement(pPage, cimElements, newElementName + " " + element.Name, false);

            }
            else
            {
                //Create new object and add it to current layout
                cimElement.Name = newElementName + " " + cimElement.Name;
                copiedElement = ElementFactory.Instance.CreateElement(pPage, cimElement, false);
            }

     
            return copiedElement;
        }

        /// <summary>
        /// Will output the x,y coordinate for the first anchor of the legend
        /// Default is center of layout, else if inside CGM template, will be placed inside the LEGEND element.
        /// </summary>
        /// <returns></returns>
        public async Task<Tuple<double, double>> GetAnchorPointStart()
        {

            anchorPoint = new Tuple<double, double>(0.0, 0.0);

            try
            {
                if (pPage != null)
                {
                    await QueuedTask.Run(async () =>
                    {
                        Element referenceLegendElement = pPage.Elements.Where(e => e.Name == Constants.Graphics.cgmLegendElement ||
                        e.Name.Contains(Constants.Graphics.cgmDetectorKeyword)).FirstOrDefault();

                        //Case when CGM blue legend box is detected, get new anchor
                        if (referenceLegendElement != null && referenceLegendElement.GetBounds() != null)
                        {
                            anchorPoint = new Tuple<double, double>(referenceLegendElement.GetBounds().XMin, referenceLegendElement.GetBounds().YMax);
                            if (referenceLegendElement.Name.Contains(Constants.Graphics.cgmDetectorKeyword))
                            {
                                isCGMTemplateMXD = true;
                            }
                            
                        }
                        //Case when nothing has been save relative to cgm blue legend box, default to upper left of paper layout with a margin of 10
                        else if (anchorPoint.Item2 == 0.0 || isCGMTemplateMXD == false)
                        {
                            anchorPoint = new Tuple<double, double>(10.0, pPage.GetPage().Height - 10);
                            isCGMTemplateMXD = false;
                        }
                    });
                }
            }
            catch (Exception GetAnchorPointStartException)
            {
                new ErrorService(GetAnchorPointStartException).WriteToFile();
            }

            return anchorPoint;

        }

        /// <summary>
        /// Will initialize some width and height of some graphics
        /// </summary>
        /// <returns></returns>
        public async Task InitWidthHeight()
        {
            columnWidth = xSpacings.COLUMN_WIDTH;
            elementWidth = xSpacings.ELEMENT_WIDTH;
            elementDescriptGapWidth = xSpacings.ELEMENT_DESCRIPTION_GAP_WIDTH;
            descriptionWidth = xSpacings.DESCRIPTION_WIDTH;
            columnColumnGapWidth = xSpacings.COLUMN_COLUMN_GAP_WIDTH;
            smallDescriptionHeight = Constants.YSpacings.smallDescriptionHeightLimit;
            smallDescriptionHeightLine = Constants.YSpacings.smallDescriptionHeightLimitLines;
            groupDescriptionWidth = xSpacings.GROUP_DESCRIPTION_WIDTH;

        }

        /// <summary>
        /// Will retrieve all legend field values from a given row
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        public async Task GatherCurrentRowInformation(Row row)
        {
            if (row[orderFieldIndex] != null)
            {
                Double.TryParse(row[orderFieldIndex].ToString(), out currentOrder);
            }
            if (row[style1FieldIndex] != null)
            {
                currentStyle1 = row[style1FieldIndex].ToString();
            }
            if (row[style2FieldIndex] != null)
            {
                currentStyle2 = row[style2FieldIndex].ToString();
            }
            if (row[labelFieldIndex] != null)
            {
                currentLabel1 = row[labelFieldIndex].ToString();
            }
            if (row[label2FieldIndex] != null)
            {
                currentLabel2 = row[label2FieldIndex].ToString();
            }
            if (row[descriptionFieldIndex] != null)
            {
                currentDescription = row[descriptionFieldIndex].ToString();
            }
            if (row[headingFieldIndex] != null)
            {
                currentHeading = row[headingFieldIndex].ToString();
            }
            if (row[elementFieldIndex] != null)
            {
                currentElementName = row[elementFieldIndex].ToString();
            }
            if (row[label1StyleFieldIndex] != null)
            {
                currentLabel1Style = row[label1StyleFieldIndex].ToString();
            }
            if (row[label2StyleFieldIndex] != null)
            {
                currentLabel2Style = row[label2StyleFieldIndex].ToString();
            }
            if (row[columnFieldIndex] != null)
            {
                int.TryParse(row[columnFieldIndex].ToString(), out currentColumn);
            }
        }

        /// <summary>
        /// Clean and replace < characters from description
        /// Having <bol></bol> within description along an extra < symbol, breaks the bolding of the heading within the description
        /// </summary>
        /// <returns></returns>
        public async Task CleanupDescription()
        {
            if (currentDescription != string.Empty && currentHeading != string.Empty && heading5Text.Count == 0)
            {
                currentDescription = currentDescription.Replace("<", "&lt;");
            }
        }

        /// <summary>
        /// Will show a warning message to user regarding a missing, empty or null order value. The warning will be shown only once per legend rendering session.
        /// </summary>
        /// <returns></returns>
        public async Task ManageNullOrder(Row orderRow) 
        {
            //Manage null order
            if (orderRow[orderFieldIndex].ToString() == string.Empty || orderRow[orderFieldIndex].ToString() == "<Null>" || orderRow[orderFieldIndex] == null)
            {
                if (!nullOrderBreaker)
                {
                    MessageBoxResult msgBoxResult = MessageBox.Show(string.Format(Properties.Resources.WarningNullOrderFound, currentElementName + " " + currentHeading + " " + currentDescription), 
                        Properties.Resources.GenericWarningTitle, System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Exclamation);
                    if (msgBoxResult == MessageBoxResult.OK)
                    {
                        nullOrderBreaker = true;
                    }
                }
            }
        }

        /// <summary>
        /// Set heading5 for special style symbols
        /// Two cases, either user repeats heading5 text in wanted embedded symbols, or
        /// uses the latest element named HEADING5_END without duplicating heading5 text in all symbols
        /// </summary>
        /// <param name="heading5Row"></param>
        /// <returns></returns>
        public async Task Heading5Preparation(Row heading5Row)
        {
            //Check if anything in current list
            if (heading5Text.Count > 0)
            {
                //Add any duplicate with current heading
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
                if (currentElementName == Constants.Graphics.heading5_end)
                {
                    heading5Text = new List<string>(); //reinitialize
                }
            }
        }

        /// <summary>
        /// Will return a x spacing based on element names
        /// </summary>
        /// <returns></returns>
        private double GetXSpacing(string toElementName)
        {
            //Variable
            double x = 0.0;

            string xSpacingElement = xSpacings.GetType().GetProperty(toElementName)?.GetValue(xSpacings, null)?.ToString();

            if (xSpacingElement != null && xSpacingElement != string.Empty)
            {
                try
                {
                    x = Convert.ToDouble(xSpacingElement, CultureInfo.InvariantCulture);
                }
                catch (Exception)
                {

                }
            }

            return x;
        }

        /// <summary>
        /// Will return a y spacing based on from and to element names
        /// </summary>
        /// <returns></returns>
        private double GetYSpacing(Element fromElement, string fromElementType, string toElementName, double lastYSpacing)
        {
            //Variable
            double y = 0.0;

            if (fromElement != null)
            {
                object ySpacingFromElement = ySpacings.GetType().GetProperty(fromElementType);

                //Get x-y spacing
                if (ySpacingFromElement != null)
                {
                    string ySpacingToElement = ySpacings.GetSpacing(fromElementType, toElementName);

                    if (ySpacingToElement != null && ySpacingToElement != string.Empty)
                    {
                        try
                        {
                            y = Convert.ToDouble(ySpacingToElement, CultureInfo.InvariantCulture);
                        }
                        catch (Exception)
                        {
                            if (ySpacingToElement.Contains(Constants.Graphics.anchorLowerLeft))
                            {
                                y = Math.Abs((fromElement.GetBounds().YMin - lastYSpacing)) + Convert.ToDouble(ySpacingToElement.Split(' ')[0], CultureInfo.InvariantCulture);
                            }
                        }

                        //Special case for heading3
                        if (fromElementType == Constants.Graphics.heading3 || fromElementType == Constants.Graphics.topNote || fromElementType == Constants.Graphics.note)
                        {
                            y = fromElement.GetBounds().Height + Convert.ToDouble(ySpacingToElement.Split(' ')[0], CultureInfo.InvariantCulture);
                        }
                    }
                }
            }

            return y;
        }

        /// <summary>
        /// Will manage the column number, in case user wants it to be autocalculated based on the widght of a legend element template
        /// </summary>
        /// <returns></returns>
        private async Task ManageColumn(Row columnRow)
        {
            try
            {
                if (_legendAutoCalculateColumn)
                {
                    //Track column change with auto-calculate
                    if (legendYLowerBound != 0.0 && currentElementObject != null)
                    {
                        if ((anchorPoint.Item2 - ySpacing - currentElementObject.GetBounds().Height) < legendYLowerBound)
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
                    if (templateGraphicDico.ContainsKey(currentElementName))
                    {
                        //Element newColumnFirstElement = CopyElementObject(templateGraphicDico[currentElementName] as Element, currentOrder.ToString());

                        //Get anchor type and calculate y spacing based on it.
                        Anchor elementAnchor = currentElementObject.GetAnchor();
                        if (elementAnchor == Anchor.CenterPoint || elementAnchor == Anchor.LeftMidPoint || elementAnchor == Anchor.RightMidPoint)
                        {
                            ySpacing = (currentElementObject.GetBounds().Height / 2.0);
                        }
                    }
                    lastColumn = currentColumn;

                    //Reset right bracket number
                    howManyRightBrackets = 0;

                }
            }
            catch (Exception ManageColumnException)
            {
                new ErrorService(ManageColumnException).WriteToFile();
            }

        }

        /// <summary>
        /// Will add a heading element to the legend based on the current row information
        /// </summary>
        /// <returns></returns>
        private async Task AddHeading(Row headingRow)
        {
            try
            {
                if (currentElementObject != null && currentElementName.Contains(Constants.Graphics.heading1.Substring(0, 6)))
                {

                    string originalElementName = currentElementObject.Name.Split(" ")[1];

                    //TODO special cases heading 3 like description, heading 4 UL

                    #region Move to right anchor

                    //Set new anchor
                    anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing);
                    PositionElement(currentElementObject, anchorPoint.Item1, anchorPoint.Item2);

                    //Set height for heading3 
                    if (currentElementName.Contains(Constants.Graphics.heading3))
                    {
                        //Recalculate height
                        string tempGroupHeadingDescription = currentHeading;
                        if (currentDescription != null)
                        {
                            tempGroupHeadingDescription = currentHeading + currentDescription;
                        }
                        double heading3Height = GetTextHeight(tempGroupHeadingDescription, descriptionWidth, Constants.TextConfiguration.lineHeight);

                        //Set new envelope
                        SetRectangularPolygonFromAnchorTypeAndHeight(currentElementObject, anchorPoint, heading3Height);
                    }
                    else
                    {
                        //Set new envelope
                        SetRectangularPolygonFromAnchorType(currentElementObject, anchorPoint);
                    }

                    //Move in X
                    MoveElement(currentElementObject, xSpacing, 0);

                    #endregion

                    //Special case for heading 3 since we can't have bolded all caps setting inside a graphic along
                    //no cap and not bolded description.
                    if (currentElementName.Contains(Constants.Graphics.heading3))
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
                    if (currentElementName.Contains(Constants.Graphics.heading5))
                    {
                        //Keep heading text so it can be used for a trigger to modify description style for heading 5 only.
                        heading5Text.Add(currentHeading);
                    }

                    //Set heading text and manage empty
                    TextElement tElement = currentElementObject as TextElement;
                    tElement.SetTextProperties(new TextProperties(currentHeading, tElement.TextProperties.Font, tElement.TextProperties.FontSize, tElement.TextProperties.FontStyle));
                    if (currentHeading == null || currentHeading == string.Empty || currentHeading == " ")
                    {
                        tElement = Symbols.GetMissingTextSymbol(tElement);
                    }

                    //Manage style if needed
                    if (currentStyle1 != "")
                    {
                        if (textSymbolDico.ContainsKey(currentStyle1))
                        {
                            SymbolStyleItem inStyleSymbol = textSymbolDico[currentStyle1];

                            CIMGraphic cimGraphic = tElement.GetGraphic();
                            if (cimGraphic != null)
                            {
                                CIMTextSymbol cIMTextSymbol = cimGraphic.Symbol.Symbol as CIMTextSymbol;

                                if (cIMTextSymbol != null)
                                {
                                    cIMTextSymbol.SetColor(cIMTextSymbol.GetColor());
                                    cIMTextSymbol.FontFamilyName = cIMTextSymbol.FontFamilyName;
                                    cIMTextSymbol.SetSize(cIMTextSymbol.GetSize());
                                    cIMTextSymbol.VerticalAlignment = cIMTextSymbol.VerticalAlignment;
                                    tElement.SetGraphic(cimGraphic);
                                }

                            }
                        }
                        else
                        {
                            //Missing or wrong style 
                            tElement = Symbols.GetMissingTextSymbol(tElement);
                        }


                    }
                    //pLayoutView.Refresh();
                    ////Add base element
                    //currentDoc.ActiveView.GraphicsContainer.AddElement(headElement, 0);
                    //currentDoc.ActiveView.GraphicsContainer.BringToFront(currentGrapSelection.SelectedElements);

                    ////Unselect
                    //currentGrapSelection.UnselectElement(headElement);

                    //Add to legend list
                    legendElementList.Add(currentElementObject);

                    //Keep name
                    lastElement = currentElementObject;
                    lastElementType = originalElementName;

                }
            }
            catch (Exception AddHeadingException)
            {
                new ErrorService(AddHeadingException).WriteToFile();
            }

        }

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

            try
            {
                //Adjust with possible font GSCGeology2015. Need to have bigger box
                if (otherComponents.GEOLOGY_FONT_NAME != null && inText.Contains(otherComponents.GEOLOGY_FONT_NAME))
                {
                    tHeight = tHeight + Constants.Fonts.geologyFontHeightAjustement;

                    //Strip text of tags that could make it look longer then it is
                    inText = inText.Replace(Constants.TextConfiguration.tagFont + '"' + otherComponents.GEOLOGY_FONT_NAME + '"' + ">", "");
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
            }
            catch (Exception GetTextHeightException)
            {
                new ErrorService(GetTextHeightException).WriteToFile();
            }


            return tHeight;

        }

        /// <summary>
        /// From a given element and anchor point, will calculate a new polygon enveloppe to fit anchor point so the element can
        /// be set at the right place on the layout before being moved.
        /// </summary>
        /// <param name="inElement"></param>
        /// <param name="inAnchor"></param>
        /// <returns></returns>
        public void SetRectangularPolygonFromAnchorType(Element inElement, Tuple<double, double> inAnchor)
        {
            SetRectangularPolygonFromAnchorTypeAndHeight(inElement, inAnchor, inElement.GetBounds().Height);
        }

        /// <summary>
        /// From a given element, anchor point and height, will calculate a new polygon envelope
        /// </summary>
        /// <param name="inElement"></param>
        /// <param name="inAnchor"></param>
        /// <param name="inHeight"></param>
        public void SetRectangularPolygonFromAnchorTypeAndHeight(Element inElement, Tuple<double, double> inAnchor, double inHeight)
        {
            try
            {
                if (inElement != null)
                {
                    //Get anchor type
                    Anchor currentAnchorPointType = inElement.GetAnchor();
                    GeometryType inGeometryType = inElement.GetGeometry().GeometryType;

                    //Apply conversion factor
                    //inHeight = inHeight ;
                    double inElementWidth = inElement.GetWidth();
                    double inElementHeight = inElement.GetHeight();

                    //Default envelop based on top left corner anchor point
                    Coordinate2D lowerLeftPoint = new Coordinate2D(inAnchor.Item1, inAnchor.Item2 - inHeight);
                    Coordinate2D upperRightPoint = new Coordinate2D(inAnchor.Item1 + inElementWidth, inAnchor.Item2);

                    ArcGIS.Core.Geometry.Envelope elementEnvelope = inElement.GetBounds();
                    if (elementEnvelope != null)
                    {
                        switch(currentAnchorPointType)
                        {
                            case Anchor.TopLeftCorner:

                                lowerLeftPoint = new Coordinate2D(inAnchor.Item1, inAnchor.Item2 - inHeight);
                                upperRightPoint = new Coordinate2D(inAnchor.Item1 + inElementWidth, inAnchor.Item2);

                                Envelope llEnvelope = EnvelopeBuilderEx.CreateEnvelope(lowerLeftPoint, upperRightPoint);

                                ArcGIS.Core.Geometry.Polygon topLeftPoly = PolygonBuilderEx.CreatePolygon(llEnvelope);
                                inElement.SetGeometry(topLeftPoly);

                                break;

                            case Anchor.CenterPoint:

                                lowerLeftPoint = new Coordinate2D(inAnchor.Item1 - inElementWidth / 2.0, inAnchor.Item2 - inHeight / 2.0);
                                upperRightPoint = new Coordinate2D(inAnchor.Item1 + inElementWidth / 2.0, inAnchor.Item2 + inHeight / 2.0);

                                //Validate if polygon graphic is a group of lines elements, which behaves differently 
                                //and for some reasons even if the anchor point is center center, behaves likes it's left
                                //see same case for line graphices.
                                if (IsElementAllNonFlatLines(inElement))
                                {
                                    lowerLeftPoint = new Coordinate2D(inAnchor.Item1 + inElementWidth / 2.0 - inElementWidth / 2.0, inAnchor.Item2 - inHeight / 2.0);
                                    upperRightPoint = new Coordinate2D(inAnchor.Item1 + inElementWidth / 2.0 - inElementWidth / 2.0, inAnchor.Item2 - inHeight / 2.0);
                                }

                                Envelope cEnvelope = EnvelopeBuilderEx.CreateEnvelope(lowerLeftPoint, upperRightPoint);
                                ArcGIS.Core.Geometry.Polygon centerPoly = PolygonBuilderEx.CreatePolygon(cEnvelope);
                                inElement.SetGeometry(centerPoly);


                                break;
                            case Anchor.TopMidPoint:

                                lowerLeftPoint = new Coordinate2D(inAnchor.Item1 - inElementWidth / 2.0, inAnchor.Item2 - inHeight);
                                upperRightPoint = new Coordinate2D(inAnchor.Item1 + inElementWidth / 2.0, inAnchor.Item2);

                                Envelope tmEnvelope = EnvelopeBuilderEx.CreateEnvelope(lowerLeftPoint, upperRightPoint);

                                ArcGIS.Core.Geometry.Polygon tmtPoly = PolygonBuilderEx.CreatePolygon(tmEnvelope);
                                inElement.SetGeometry(tmtPoly);

                                break;
                            case Anchor.LeftMidPoint:

                                lowerLeftPoint = new Coordinate2D(inAnchor.Item1, inAnchor.Item2 - inHeight / 2.0);
                                upperRightPoint = new Coordinate2D(inAnchor.Item1 + inElementWidth, inAnchor.Item2 + inHeight / 2.0);

                                Envelope lmEnvelope = EnvelopeBuilderEx.CreateEnvelope(lowerLeftPoint, upperRightPoint);

                                ArcGIS.Core.Geometry.Polygon lmPoly = PolygonBuilderEx.CreatePolygon(lmEnvelope);
                                inElement.SetGeometry(lmPoly);

                                break;
                            case Anchor.BottomRightCorner:

                                lowerLeftPoint = new Coordinate2D(inAnchor.Item1 - inElementWidth, inAnchor.Item2);
                                upperRightPoint = new Coordinate2D(inAnchor.Item1, inAnchor.Item2 + inHeight);

                                Envelope brEnvelope = EnvelopeBuilderEx.CreateEnvelope(lowerLeftPoint, upperRightPoint);

                                ArcGIS.Core.Geometry.Polygon brPoly = PolygonBuilderEx.CreatePolygon(brEnvelope);
                                inElement.SetGeometry(brPoly);

                                break;
                            case Anchor.BottomLeftCorner:

                                lowerLeftPoint = new Coordinate2D(inAnchor.Item1, inAnchor.Item2);
                                upperRightPoint = new Coordinate2D(inAnchor.Item1 + inElementWidth, inAnchor.Item2 - inHeight);

                                Envelope blEnvelope = EnvelopeBuilderEx.CreateEnvelope(lowerLeftPoint, upperRightPoint);

                                ArcGIS.Core.Geometry.Polygon blPoly = PolygonBuilderEx.CreatePolygon(blEnvelope);
                                inElement.SetGeometry(blPoly);

                                break;
                        }

                    }
                }
            }
            catch (Exception SetRectangularPolygonFromAnchorTypeAndHeightException)
            {
                new ErrorService(SetRectangularPolygonFromAnchorTypeAndHeightException).WriteToFile();
            }

        }

        /// <summary>
        /// Will validate if input element is a  non flat (height 0) line element, or a group element composed of a bunch of non-flat lines.
        /// Knowing this can results in different methods of anchoring since they behave like left anchor compared to whatever they are set to.
        /// Flat line graphics always behave like anchor point is at the left side. Area of lines and heighted lines behaves the same but can be seen as polygons.
        /// </summary>
        /// <param name="inElement"></param>
        /// <returns></returns>
        public bool IsElementAllNonFlatLines(Element inElement)
        {
            bool allLines = true;
            ArcGIS.Core.Geometry.Geometry elementGeometry = inElement.GetGeometry();

            //For single line element
            if (elementGeometry != null && elementGeometry.GeometryType != GeometryType.Polyline)
            {
                allLines = false;
            }
            else
            {
                //Check height
                if (inElement.GetBounds().Height == 0)
                {
                    allLines = false;
                }
                else
                {
                    allLines = true;
                }
            }

            //for group of elements
            GroupElement inGroupElement = inElement as GroupElement;
            if (inGroupElement != null)
            {
                //Check geometry of inner elements, if it's all lines
                for (int el = 0; el < inGroupElement.GetElementsAsFlattenedList().Count(); el++)
                {
                    Element innerElement = inGroupElement.GetElementsAsFlattenedList()[el];
                    if (innerElement.GetGeometry().GeometryType != GeometryType.Polyline)
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
        /// Will position a given element to a new location coordinate. The new coordinates can fit a desire anchor point.
        /// The element anchor will change momentarily to set the new position, then it'll revert to it's original value
        /// </summary>
        /// <param name="pElement"></param>
        /// <param name="newX"></param>
        /// <param name="newY"></param>
        /// <param name="withAnchor"></param>
        public void PositionElement(Element pElement, double newX, double newY, Anchor withAnchor = Anchor.BottomLeftCorner)
        {

            try
            {
                //Get actual anchor point coordinates
                Coordinate2D newAnchorCoordinates = new Coordinate2D(newX, newY);

                //Keep actual anchor in case it's different
                //Anchor currentAnchor = pElement.GetAnchor();

                //Set anchor so it fits the desire coordinate
                //pElement.SetAnchor(withAnchor);

                //Set the new geometry
                pElement.SetAnchorPoint(newAnchorCoordinates);

                //Reset the old anchor
                //pElement.SetAnchor(currentAnchor);



            }
            catch (Exception positionElementException)
            {
                new ErrorService(positionElementException).WriteToFile();
            }

        }

        /// <summary>
        /// Will position a given element to a new location coordinate. The new coordinates can fit a desire anchor point.
        /// The element anchor will change momentarily to set the new position, then it'll revert to it's original value
        /// </summary>
        /// <param name="pElement"></param>
        /// <param name="newX"></param>
        /// <param name="newY"></param>
        /// <param name="withAnchor"></param>
        public void MoveElement(Element pElement, double deltaX, double deltaY, Anchor withAnchor = Anchor.BottomLeftCorner)
        {

            try
            {
                //Get actual anchor point coordinates
                Coordinate2D anchorCoordinates = pElement.GetAnchorPoint();
                anchorCoordinates = new Coordinate2D(anchorCoordinates.X + deltaX, anchorCoordinates.Y + deltaY);

                //Set the new geometry
                pElement.SetAnchorPoint(anchorCoordinates);

            }
            catch (Exception positionElementException)
            {
                new ErrorService(positionElementException).WriteToFile();
            }

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
    }
}
