using ActiproSoftware.Windows.Media;
using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Data.Analyst3D;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Core.Internal.Geometry;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Catalog.PropertyPages.NetworkDataset;
using ArcGIS.Desktop.Internal.KnowledgeGraph;
using ArcGIS.Desktop.Internal.Mapping.Symbology;
using ArcGIS.Desktop.Internal.Reports;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using GSCLegendRendererPro.Models;
using GSCLegendRendererPro.Services;
using GSCLegendRendererPro.Utilities;
using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;
using System.Windows.Navigation;
using System.Windows.Shapes;
using static ArcGIS.Desktop.Editing.Templates.EditingGroupTemplate;
using static ArcGIS.Desktop.Internal.GeoProcessing.Controls.rtbEditor;
using static GSCLegendRendererPro.Utilities.Layers;
using static System.Net.Mime.MediaTypeNames;
using Color = ArcGIS.Core.Internal.CIM.Color;
using Envelope = ArcGIS.Core.Geometry.Envelope;
using Field = ArcGIS.Core.Data.Field;
using Geometry = ArcGIS.Core.Geometry.Geometry;
using Line = ArcGIS.Core.Internal.CIM.Line;
using LinearUnit = ArcGIS.Core.Geometry.LinearUnit;
using Polyline = ArcGIS.Core.Geometry.Polyline;
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
        LinearUnit originalPageUnits = null;

        //GRAPHICS PROCESSING
        public Dictionary<string, Element> templateGraphicDico { get; set; }
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
        public Element demPictureElementObject { get; set; }
        List<Element> legendElementList = new List<Element>(); //Will hold all legend items to group them at the end of the process.
        List<String> legendOrderPrefixList = new List<string>();

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

                                                #region GRAPHIC PREPARATION

                                                await GatherCurrentRowInformation(legendRow);

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

                                                //EDGE CASE: Set heading5 trigger for special style symbols
                                                await Heading5Preparation(legendRow);

                                                //Get spacings in x and Y for current row
                                                ySpacing = GetYSpacing(lastElement, lastElementType, currentElementName, anchorPoint.Item2);
                                                xSpacing = GetXSpacing(currentElementName);

                                                //Manage column change and spacing
                                                await ManageColumn(legendRow);

                                                #endregion

                                                #region GRAPHIC HANDLING

                                                await AddHeading();

                                                await AddMapUnit();

                                                await AddThinUnit();

                                                await AddEmbeddedMapUnit();

                                                await AddMarker();

                                                await AddLine();

                                                await AddOverlay();

                                                await AddBreakLine();

                                                await AddBreakLineAnnotation();

                                                await AddLeftBracket();

                                                await AddRightBracket();

                                                await AddNote();

                                                #endregion

                                                #region FINALIZE

                                                //Add to legend list
                                                legendElementList.Add(currentElementObject);

                                                //Keep name
                                                //EDGE CASE - No bracket elements, because they wrap around other elements and do not interact
                                                //With the ordering and rendering of other legend items
                                                if (!currentElementName.Contains(Constants.Graphics.keywordBracket))
                                                {
                                                    lastElement = currentElementObject;
                                                    lastElementType = currentElementName;
                                                }

                                                await AutomaticColumnsCalculation();

                                                #endregion
                                            }
                                        }
                                    }  
                                }

                            }

                            //Finalize whole process
                            await GroupByOrder();
                            await OrderElementsInTOC();
                            await GroupLegendElements();
                            SetPageUnits(originalPageUnits);

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
                SetPageUnits(LinearUnit.Millimeters);

                //Force delay update
                //Was in ArcMap, no equivalent for ArcPro
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

                //Delete previous results if any
                await CleanUpOldLegend();

                return true;
            }
            catch (Exception SetupLayoutAndGraphicsException)
            {
                new ErrorService(SetupLayoutAndGraphicsException).WriteToFile();
                return false;
            }

        }

        /// <summary>
        /// Will search for the default legend group name and remove it, so user doesn't have
        /// to delete it each time they launches the tool.
        /// </summary>
        /// <returns></returns>
        public async Task CleanUpOldLegend()
        {
            await QueuedTask.Run(async () =>
            {
                //Cleanup - Delete previous legend if any
                Element legendElement = pPage.GetElements().Where(e => e.Name == Properties.Resources.LegendGroupName).FirstOrDefault();
                if (legendElement != null)
                {
                    pPage.DeleteElement(legendElement);
                }
            });

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
        public void SetPageUnits(LinearUnit units)
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
                            if (units == null)
                            {
                                //Set default to mm
                                units = LinearUnit.Millimeters;
                            }

                            if (originalPageUnits == null)
                            {
                                originalPageUnits = cIMPage.Units;
                            }

                            if (cIMPage.Units != units)
                            {
                                //Enforce centimeters
                                cIMPage.Units = units;
                                cIMPage.StretchElements = true;
                                pPage.SetPage(cIMPage);
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
                templateGraphicDico = new Dictionary<string, Element>();

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
        public Element CopyElementObject(Element element, string elementNamePrefix, string elementNameSuffix = "")
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
                copiedElement = ElementFactory.Instance.CreateGroupElement(pPage, cimElements, elementNamePrefix + " " + element.Name + elementNameSuffix, false);

            }
            else
            {
                //Create new object and add it to current layout
                cimElement.Name = elementNamePrefix + " " + cimElement.Name + elementNameSuffix;
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
                        Element referenceLegendElement = pPage.Elements.Where(e => e.Name == Constants.Graphics.cgmLegendElement).FirstOrDefault();

                        //Case when CGM blue legend box is detected, get new anchor
                        if (referenceLegendElement != null && referenceLegendElement.GetBounds() != null)
                        {
                            anchorPoint = new Tuple<double, double>(referenceLegendElement.GetBounds().XMin, referenceLegendElement.GetBounds().YMax);
                            isCGMTemplateMXD = true;

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
            else
            {
                currentStyle1 = null;
            }

            if (row[style2FieldIndex] != null)
            {
                currentStyle2 = row[style2FieldIndex].ToString();
            }
            else
            {
                currentStyle2 = null;
            }

            if (row[labelFieldIndex] != null)
            {
                currentLabel1 = row[labelFieldIndex].ToString();
            }
            else
            {
                currentLabel1 = null;
            }

            if (row[label2FieldIndex] != null)
            {
                currentLabel2 = row[label2FieldIndex].ToString();
            }
            else
            {
                currentLabel2 = null;
            }

            if (row[descriptionFieldIndex] != null)
            {
                currentDescription = row[descriptionFieldIndex].ToString();
            }
            else
            {
                currentDescription = null;
            }

            if (row[headingFieldIndex] != null)
            {
                currentHeading = row[headingFieldIndex].ToString();
            }
            else
            {
                currentHeading = null;
            }

            if (row[elementFieldIndex] != null)
            {
                currentElementName = row[elementFieldIndex].ToString();
            }
            else
            {
                currentElementName = null;
            }

            if (row[label1StyleFieldIndex] != null)
            {
                currentLabel1Style = row[label1StyleFieldIndex].ToString();
            }
            else
            {
                currentLabel1Style = null;
            }

            if (row[label2StyleFieldIndex] != null)
            {
                currentLabel2Style = row[label2StyleFieldIndex].ToString();
            }
            else
            {
                currentLabel2Style = null;
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
            if (currentDescription != null && currentDescription != string.Empty && currentHeading != string.Empty && heading5Text.Count == 0)
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
            if (orderRow[orderFieldIndex] == null || orderRow[orderFieldIndex].ToString() == string.Empty || orderRow[orderFieldIndex].ToString() == "<Null>")
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
                else
                {
                    //EDGE CASE - If element is not present and still need some space between the current and next element (usually meant for TOP_NOTE element)
                    y = fromElement.GetBounds().Height + Constants.YSpacings.minimumHeightBetweenTextElement;
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
                if (otherComponents.GEOLOGY_FONT_NAME != null && inText != null && inText.Contains(otherComponents.GEOLOGY_FONT_NAME))
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
                if (inText != null)
                {
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

            //Final validation
            if (tHeight == 0)
            {
                tHeight = Constants.TextConfiguration.lineHeight;
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
                    double inElementWidth = inElement.GetBounds().Width;
                    //double inElementHeight = inElement.GetBounds().Height;

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
        public void PositionElement(Element pElement, double newX, double newY, Anchor withAnchor = Anchor.Unspecified)
        {

            try
            {
                //Get actual anchor point coordinates
                Coordinate2D newAnchorCoordinates = new Coordinate2D(newX, newY);

                //Keep actual anchor in case it's different
                Anchor currentAnchor = pElement.GetAnchor();

                //Set anchor so it fits the desire coordinate
                if (withAnchor != Anchor.Unspecified)
                {
                    pElement.SetAnchor(withAnchor);
                }

                //Set the new geometry
                pElement.SetAnchorPoint(newAnchorCoordinates);

                //Reset the old anchor
                if (withAnchor != Anchor.Unspecified)
                {
                    pElement.SetAnchor(currentAnchor);
                }
                
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

        /// <summary>
        /// The tool processing each element from their order, each added item are being added on top of
        /// each other in the table of content. At the end the one at the top will be the last legend item.
        /// Resort them just like it was set in the table order field.
        /// </summary>
        /// <returns></returns>
        private async Task OrderElementsInTOC()
        {
            //Sort by the order value
            legendOrderPrefixList = legendOrderPrefixList.OrderByDescending(x => {double.TryParse(x.Split(' ')[0], out double value);return value;}).ToList();

            foreach (string legendOrder in legendOrderPrefixList)
            {
                Element orderGroupElement = pPage.GetElements().Where(x => x is GroupElement && x.Name == legendOrder).FirstOrDefault();

                if (orderGroupElement != null)
                {
                    pPage.SelectElement(orderGroupElement);
                    if (pPage.CanBringForward(orderGroupElement))
                    {
                        pPage.BringToFront(orderGroupElement);
                    }
                }
            }

            //EDGE CASE - UNIT_PARENT needs to be send lower than their child
            List<Element> parentElements = pPage.GetElementsAsFlattenedList().Where(e => e.Name.Contains(Constants.Graphics.unitParent)).ToList();
            if (parentElements != null && parentElements.Count() > 0)
            {
                foreach (Element pe in parentElements)
                {
                    //Select the whole group
                    string groupPrefix = pe.Name.Split(" ")[0];
                    Element parentGroupElement = pPage.GetElements().Where(e => (e is GroupElement) && e.Name == groupPrefix).FirstOrDefault();
                    if (parentGroupElement != null) 
                    {
                        pPage.SelectElement(parentGroupElement);
                        if (pPage.CanSendBackward(parentGroupElement))
                        {
                            pPage.SendToBack(parentGroupElement);
                        }
                    }
                }
            }

        }

        /// <summary>
        /// Will select all grouped order graphics and add them to a new group
        /// for ease of work after tool is done (ex. move, delete)
        /// </summary>
        /// <returns></returns>
        private async Task GroupLegendElements()
        {
            List<Element> groupedElements = pPage.GetElements().Where(x => x is GroupElement).ToList();
            List<Element> desiredGroupedElements = new List<Element>();

            //Validate that groups are associated to legend and not something else by the user
            foreach (Element element in groupedElements)
            {
                double elementOrder = 0.0;
                if (double.TryParse(element.Name, out elementOrder))
                {
                    desiredGroupedElements.Add(element);
                }
            }

            //Need to select elements first
            pPage.SelectElements(desiredGroupedElements);

            //Group
            ElementFactory.Instance.CreateGroupElement(pPage, pPage.GetSelectedElements(), Properties.Resources.LegendGroupName, false);
        }

        /// <summary>
        /// Will group each added items by their name prefix which is set as the user order for each table record from the legend table
        /// </summary>
        /// <returns></returns>
        private async Task GroupByOrder()
        {
            //Group by order
            List<List<Element>> groupedOrderList = legendElementList.GroupBy(x => x.Name.Split(" ")[0]).Select(g => g.ToList()).ToList();

            //Iterate through groups and select them and create a new group
            foreach (List<Element> group in groupedOrderList)
            {
                string groupName = group.First().Name.Split(" ")[0];

                //Check as to not incorporate anything else that could exist on user layout
                //If the prefix is only numbers, make a group
                double orderGroup = 0.0;
                if (double.TryParse(groupName, out orderGroup))
                {
                    legendOrderPrefixList.Add(groupName); //Keep for later
                    pPage.SelectElements(group);
                    ElementFactory.Instance.CreateGroupElement(pPage, pPage.GetSelectedElements(), groupName, false);
                }


            }
        
        }

        /// <summary>
        /// Will set the background color for given element with given style
        /// </summary>
        /// <param name="inElement"></param>
        /// <param name="style"></param>
        public Element SetPolygonFill(Element inElement, string style, bool isSimpleFill, bool isUnitBoxOnly = false, Tuple<double, double> inAnchor = null, string style2 = "")
        {
            try
            {
                GraphicElement graphicElement = inElement as GraphicElement;
                if (graphicElement != null)
                {

                    if (style != null && fillSymbolDico.ContainsKey(style) && isSimpleFill)
                    {

                        //Get symbol type and color
                        string symbolTypeName = string.Empty;
                        SymbolStyleItem fillSymbol = fillSymbolDico[style];

                        //Fill polygon or replace with related DEM image
                        if (_legendDEM && isUnitBoxOnly)
                        {
                            //Detect tranparent color and force it white
                            Element demElement = SetPolygonDEM(fillSymbol.Symbol.GetColor(), inAnchor);

                            return inElement;
                        }
                        //Overlay fill type 
                        else if (symbolTypeName == Constants.ObjectNames.fillTypeMultilayer)
                        {
                            CIMGraphic graphic = graphicElement.GetGraphic();
                            CIMPolygonSymbol cimPolySymbol = graphic.Symbol.Symbol as CIMPolygonSymbol;

                            if (graphic != null)
                            {
                                graphic.Symbol.Symbol = fillSymbol.Symbol;
                            }

                            //Set color if needed
                            if (style2 != null && style2 != string.Empty && fillSymbolDico.ContainsKey(style2))
                            {
                                SymbolStyleItem fillSymbol2 = fillSymbolDico[style2];
                                cimPolySymbol.SetColor(fillSymbol2.Symbol.GetColor());
                            }

                            //Manage outline
                            if (cimPolySymbol.GetOutlineColor() == CIMColor.NoColor())
                            {
                                //Apply black outline 
                                cimPolySymbol.SetOutlineColor(CIMColor.CreateRGBColor(0, 0, 0));
                            }

                            graphicElement.SetGraphic(graphic);

                            return inElement;
                        }
                        else
                        {
                            //Set background color
                            CIMGraphic graphic = graphicElement.GetGraphic();
                            if (graphic != null)
                            {
                                CIMPolygonSymbol cimPolySymbol = graphic.Symbol.Symbol as CIMPolygonSymbol;
                                cimPolySymbol.SetColor(fillSymbol.Symbol.GetColor());
                                graphicElement.SetGraphic(graphic);
                            }

                            return inElement;
                        }

                    }
                    else if (style != null && fillSymbolDico.ContainsKey(style) && !isSimpleFill)
                    {
                        SymbolStyleItem fillSymbol = fillSymbolDico[style];
                        CIMGraphic graphic = graphicElement.GetGraphic();
                        if (graphic != null)
                        {
                            graphic.Symbol.Symbol = fillSymbol.Symbol;
                            graphicElement.SetGraphic(graphic);
                        }

                        return inElement;
                    }
                    else
                    {
                        //Apply missing style
                        GraphicElement missingFillSymbol = Symbols.SetMissingPolygonSymbol(graphicElement);

                        return missingFillSymbol as Element;
                    }
                }

                return null;

            }
            catch (Exception SetPolygonFillException)
            {
                new ErrorService(SetPolygonFillException).WriteToFile();
                return null;
            }

        }

        /// <summary>
        /// Will add a picture element with given colored added above it as transparent.
        /// </summary>
        /// <param name="inElement"></param>
        /// <param name="inColor"></param>
        public Element SetPolygonDEM(CIMColor inColor, Tuple<double, double> inAnchor = null)
        {
            try
            {   
                //Variables
                Services.ImageProcessing imProcessing = new Services.ImageProcessing();

                //Calculate DEM transparency
                int demtransparency = 178; //178/255 is 70% opacity
                if (otherComponents.DEM_OPACITY_PERCENT != 70)
                {
                    double opacityConversion = Math.Round(((double)demtransparency / 100.0) * 255.0);
                    demtransparency = Convert.ToInt16(opacityConversion);
                }

                //Validate DEM picture existance and get path
                string demImagePath = System.IO.Path.Combine(Properties.Settings.Default.WorkingEnvironmentPath, Constants.Assets.demPicture);

                //Init new image object
                System.Drawing.Image demImage = System.Drawing.Image.FromFile(demImagePath);

                //Build path to new mono colored image
                //string outputFolderName = System.IO.Path.Combine(Dictionaries.Constants.ESRI.defaultArcGISFolderName, Dictionaries.Constants.Namespaces.mainNamespace + " " + ThisAddIn.Version.ToString());
                //string outputFolderPath = System.IO.Path.Combine(System.Environment.GetFolderPath(System.Environment.SpecialFolder.MyDocuments), outputFolderName);
                string monoColoredName = Constants.ImageConfiguration.monoColoredImageNamePrefix + demtransparency.ToString() + "_"
                    + inColor.GetAlphaValue().ToString() + "_" + inColor.GetColorComponent(0).ToString() + "_" + inColor.GetColorComponent(1).ToString() + "_" + inColor.GetColorComponent(2).ToString() + ".png";
                string demColoredName = Constants.Graphics.legendBoxDEM + demtransparency.ToString() + "_" + inColor.GetColorComponent(0).ToString() + "_" + inColor.GetColorComponent(1).ToString() + "_" + inColor.GetColorComponent(2).ToString() + ".png";
                string monoColoredPath = System.IO.Path.Combine(Properties.Settings.Default.WorkingEnvironmentPath, monoColoredName);
                string demColoredPath = System.IO.Path.Combine(Properties.Settings.Default.WorkingEnvironmentPath, demColoredName);

                //Process and a get a copy of new mono colored image
                if (!System.IO.File.Exists(monoColoredPath))
                {
                    imProcessing.CreateMonoColorFromImageCopy(demImage, inColor, monoColoredPath, demtransparency);
                }

                //Create bitmaps from original dem image and mono colored one
                System.Drawing.Image monoColoredImage = System.Drawing.Image.FromFile(monoColoredPath);
                Bitmap monoColoredBitmap = new Bitmap(monoColoredImage);
                Bitmap originalBitmap = new Bitmap(demImage);

                //Overlap both bitmaps
                Bitmap overlapImage = new Bitmap(monoColoredImage);
                Graphics overlapGraphic = Graphics.FromImage(overlapImage);
                overlapGraphic.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceOver;
                overlapGraphic.DrawImage(originalBitmap, 0, 0);
                overlapGraphic.DrawImage(monoColoredBitmap, 0, 0);

                //Save result
                overlapImage.Save(demColoredPath, ImageFormat.Png);

                //Get graphic element for DEM
                Element demElement = CopyElementObject(templateGraphicDico[Constants.Graphics.legendBoxDEM], currentOrder.ToString());

                //Set path
                PictureElement demPictureElement = demElement as PictureElement;
                demPictureElement.SetSourcePath(demColoredPath);

                legendElementList.Add(demElement);

                demPictureElementObject = demElement;

                //Move if needed
                if (inAnchor != null)
                {
                    //TODO find why we must substract 10 else the DEM picture is offset to the right by 10 if used inside a grouped element
                    Tuple<double, double> newDEMAnchor = new Tuple<double, double>(inAnchor.Item1, inAnchor.Item2);
                    SetRectangularPolygonFromAnchorType(demElement, newDEMAnchor);

                }

                OrderElement(demElement, false);

                return demElement;

            }
            catch (Exception SetPolygonDEMException)
            {
                new ErrorService(SetPolygonDEMException).WriteToFile();

                return null;
            }

        }

        /// <summary>
        /// Will add a new text element at the center a of map unit box
        /// </summary>
        /// <param name="inText"></param>
        /// <param name="parentElement"></param>
        /// <param name="inDocument"></param>
        private Element AddLabelInUnitBox(string inText, Element parentElement, Tuple<double, double> inAnchor, Constants.Graphics.UnitBoxType unitBoxType, string inStyle = "")
        {
            try
            {
                //Get appropriate element and set appropriate name
                string newLabelSuffix = string.Empty;
                if (unitBoxType == Constants.Graphics.UnitBoxType.split2)
                {
                    newLabelSuffix = "_2";
                }
                Element unitBoxLabelElement = CopyElementObject(templateGraphicDico[Constants.Graphics.unitLabel], currentOrder.ToString(), newLabelSuffix);

                //Create new text graphic with default style
                TextElement tElement = unitBoxLabelElement as TextElement;

                if (tElement != null)
                {
                    //Get graphic in order to change text and symbol
                    CIMGraphic currentElementGraphic = tElement.GetGraphic();

                    //Manage incoming style if needed
                    if (inStyle != null && inStyle != "" && textSymbolDico.ContainsKey(inStyle))
                    {

                        SymbolStyleItem fillSymbol = fillSymbolDico[inStyle];

                        if (currentElementGraphic != null)
                        {
                            CIMTextSymbol cIMTextSymbol = currentElementGraphic.Symbol.Symbol as CIMTextSymbol;
                            cIMTextSymbol.SetColor(ColorFactory.Instance.RedRGB);
                            cIMTextSymbol.FontFamilyName = "Arial";
                            cIMTextSymbol.SetSize(Constants.TextConfiguration.defaultUnitBoxLabelFontSize);
                            cIMTextSymbol.VerticalAlignment = ArcGIS.Core.CIM.VerticalAlignment.Center;
                            tElement.SetGraphic(currentElementGraphic);
                        }

                    }
                    
                    //Mange too long text (mainly to fix when used in UNIT_SPLIT boxes).
                    //Conditions on style to prevent trigger on special fonts
                    if (inText.Length >= 6 && inStyle == "")
                    { 

                        if (currentElementGraphic != null)
                        {
                            CIMTextSymbol cIMTextSymbol = currentElementGraphic.Symbol.Symbol as CIMTextSymbol;
                            cIMTextSymbol.SetSize(Constants.TextConfiguration.tooLongLabelUnitBoxLabelFontSize);

                            //Manage placement
                            if (unitBoxType == Constants.Graphics.UnitBoxType.split2)
                            {
                                //Shift down a bit for right part.
                                cIMTextSymbol.VerticalAlignment = ArcGIS.Core.CIM.VerticalAlignment.Baseline;
                            }
                            else if (unitBoxType == Constants.Graphics.UnitBoxType.split1)
                            {
                                //Shift up a bit for left part
                                cIMTextSymbol.VerticalAlignment = ArcGIS.Core.CIM.VerticalAlignment.Top;
                            }

                            tElement.SetGraphic(currentElementGraphic);
                        }

                    }

                    //Manage missing
                    if (inText == null || inText == string.Empty || inText == " " && currentElementGraphic != null)
                    {
                        inText = Constants.TextConfiguration.missingText;
                        tElement = Symbols.SetMissingTextSymbol(tElement, Properties.Resources.ErrorMissingLabel);
                    }
                    else
                    {
                        tElement.SetTextProperties(new TextProperties(inText, tElement.TextProperties.Font, tElement.TextProperties.FontSize, tElement.TextProperties.FontStyle));
                    }

                    //Get width and height of parent
                    ArcGIS.Core.Geometry.Geometry parentGeom = parentElement.GetGeometry();
                    Envelope parentEnvelope = parentGeom.Extent;
                    double parentHeight = parentEnvelope.Height;
                    double parentWidth = parentEnvelope.Width;

                    //Set new anchor point
                    SetRectangularPolygonFromAnchorType(unitBoxLabelElement, inAnchor);

                    //Move       
                    if (unitBoxType == Constants.Graphics.UnitBoxType.normal)
                    {
                        MoveElement(unitBoxLabelElement, parentWidth / 2.0, (-parentHeight / 2.0));//Move accordingly to anchor point which is center center
                    }
                    else if (unitBoxType == Constants.Graphics.UnitBoxType.split1)
                    {
                        //https://www.mathopenref.com/coordcentroid.html
                        double centerX = (3 * inAnchor.Item1 + parentWidth) / 3.0;
                        double centerY = (3 * inAnchor.Item2 - parentHeight) / 3.0;
                        MoveElement(unitBoxLabelElement, centerX - inAnchor.Item1, -(Math.Abs(centerY - inAnchor.Item2)));//Move accordingly to anchor point which is center center
                    }
                    else if (unitBoxType == Constants.Graphics.UnitBoxType.split2 || unitBoxType == Constants.Graphics.UnitBoxType.line || unitBoxType == Constants.Graphics.UnitBoxType.child_line)
                    {
                        //https://www.mathopenref.com/coordcentroid.html
                        double centerX = (3 * inAnchor.Item1 + 2 * parentWidth) / 3.0;
                        double centerY = (3 * inAnchor.Item2 - 2 * parentHeight) / 3.0;
                        MoveElement(unitBoxLabelElement, centerX - inAnchor.Item1, -(Math.Abs(centerY - inAnchor.Item2)));//Move accordingly to anchor point which is center center
                    }
                    else if (unitBoxType == Constants.Graphics.UnitBoxType.parent)
                    {
                        MoveElement(unitBoxLabelElement, parentWidth / 2.0, (-parentHeight / 4.0));
                    }
                    //Order
                    OrderElement(unitBoxLabelElement);

                    //Keep track of new item
                    legendElementList.Add(unitBoxLabelElement);

                }

                return unitBoxLabelElement;
            }
            catch (Exception AddLabelInUnitBoxException)
            {
                new ErrorService(AddLabelInUnitBoxException).WriteToFile();

                return null;
            }


        }

        /// <summary>
        /// Will merge a parent and child element and bring forward or backward the child, currentobjectelement will also be replaced
        /// </summary>
        /// <param name="parentElement"></param>
        /// <param name="childElement"></param>
        /// <param name="newName"></param>
        /// <param name="bringForward"></param>
        /// <returns></returns>
        public Element GroupElement(Element parentElement, Element childElement, string newName)
        {

            pPage.SelectElements(new List<Element>() { currentElementObject, childElement });

            GroupElement newGroup = ElementFactory.Instance.CreateGroupElement(pPage, pPage.GetSelectedElements(), newName, false);

            currentElementObject = newGroup;

            return newGroup;

        }

        /// <summary>
        /// Will reorder an element to set proper drawing order
        /// </summary>
        /// <param name="orderElement"></param>
        /// <param name="bringForward"></param>
        public void OrderElement(Element orderElement, bool bringForward = true)
        {
            //Send beneath unit box then group it and reset current object as a new grouped graphic
            pPage.SelectElement(orderElement);
            if (bringForward && pPage.CanBringForward(orderElement))
            {
                pPage.BringToFront(orderElement);
            }
            else if (!bringForward && pPage.CanSendBackward(orderElement))
            {
                pPage.SendBackward(orderElement);
            }
        }

        /// <summary>
        /// Will add to an existing group elemenmt instead of creating a new one each time
        /// </summary>
        /// <returns></returns>
        public Element AddToGroupElement(Element parentGroupElement, Element childElement)
        {
            GroupElement parentGroup = parentGroupElement as GroupElement;

            try
            {
                if (parentGroup != null)
                {
                    //Add
                    List<Element> parentElements = parentGroup.GetElementsAsFlattenedList().ToList();
                    parentElements.Add(childElement);

                    //Keep original nam
                    string originalGroupName = parentGroupElement.Name;

                    //Select before recreating group
                    pPage.SelectElements(parentElements);
      
                    //Recreate
                    parentGroup = ElementFactory.Instance.CreateGroupElement(pPage, pPage.GetSelectedElements(), originalGroupName + "NEW", false);

                    //Delete original group
                    pPage.DeleteElement(parentGroupElement);

                    //Rename
                    parentGroup.SetName(parentGroup.Name.Replace("NEW", ""));

                    currentElementObject = parentGroup;
                }

                return parentGroup;

            }
            catch (Exception AddToGroupElementException)
            {
                new ErrorService(AddToGroupElementException).WriteToFile();

                return null;
            }

        }

        /// <summary>
        /// Will add a new text element for description in column description part
        /// </summary>
        /// <param name="inDescription"></param>
        /// <param name="parentElem"></param>
        /// <param name="inDocument"></param>
        /// <param name="inAnchor"></param>
        /// <returns> Description height for validation purposes</returns>
        private Element AddDescription(string inDescription, Element parentElem, Tuple<double, double> inAnchor, string parentElemType, bool isLineOrPoint = false)
        {
            //Get appropriate element
            Element descriptionElement = null;

            //Get different size description
            if (parentElemType == Constants.Graphics.unitindent1)
            {
                descriptionElement = CopyElementObject(templateGraphicDico[Constants.Graphics.description_indent] as Element, currentOrder.ToString());
            }
            else if (parentElemType == Constants.Graphics.unitindent2)
            {
                descriptionElement = CopyElementObject(templateGraphicDico[Constants.Graphics.description_indent2] as Element, currentOrder.ToString());
            }
            else if (heading5Text.Count >= 1)
            {
                descriptionElement = CopyElementObject(templateGraphicDico[Constants.Graphics.heading5Description] as Element, currentOrder.ToString());
                double indentation = GetXSpacing(Constants.Graphics.heading5Description) - GetXSpacing(Constants.Graphics.description);
                inAnchor = new Tuple<double, double>(inAnchor.Item1 + indentation, inAnchor.Item2);
            }
            else
            {
                descriptionElement = CopyElementObject(templateGraphicDico[Constants.Graphics.description] as Element, currentOrder.ToString());
            }


            //Create new text graphic with default style
            TextElement dtElement = descriptionElement as TextElement;

            //Manage missing
            if (inDescription == null || inDescription == string.Empty || inDescription == " ")
            {
                dtElement = Symbols.SetMissingTextSymbol(dtElement, Properties.Resources.ErrorMissingLabel); 
            }
            else
            {
                dtElement.SetTextProperties(new TextProperties(inDescription, dtElement.TextProperties.Font, dtElement.TextProperties.FontSize, dtElement.TextProperties.FontStyle));
            }

            #region AddDescriptFromElement code
            double wantedTextHeight = GetTextHeight(dtElement.TextProperties.Text, descriptionWidth);

            //Min height setting
            double smallDescHeight = smallDescriptionHeight;
            if (isLineOrPoint)
            {
                smallDescHeight = smallDescriptionHeightLine;
            }

            //Get width and height of parent
            double parentHeight = 1.0;
            double parentWidth = elementWidth;

            //In case element is a unit box, try to get height
            if (parentElem is GroupElement)
            {
                GroupElement ge = parentElem as GroupElement;
                parentHeight = ge.GetElementsAsFlattenedList().First().GetHeight();
            }
            else
            {
                parentHeight = parentElem.GetHeight();
            }

            //Set width and height and manage group description for heading 5        
            ArcGIS.Core.Geometry.Geometry descriptionGeometry = descriptionElement.GetGeometry();
            double descriptionHeight = wantedTextHeight;

            if (heading5Text.Count >= 1)
            {
                descriptionWidth = groupDescriptionWidth;
            }

            Coordinate2D lowerLeftDescription = new Coordinate2D(descriptionGeometry.Extent.XMin, descriptionGeometry.Extent.YMin);
            Coordinate2D upperRightDescription = new Coordinate2D(descriptionGeometry.Extent.XMin + descriptionWidth, descriptionGeometry.Extent.YMin + descriptionHeight);

            Envelope llEnvelope = EnvelopeBuilderEx.CreateEnvelope(lowerLeftDescription, upperRightDescription);
            ArcGIS.Core.Geometry.Polygon descPoly = PolygonBuilderEx.CreatePolygon(llEnvelope);
            descriptionElement.SetGeometry(descPoly);

            //Set anchor
            SetRectangularPolygonFromAnchorType(descriptionElement, inAnchor);

            //Move based on different length of description
            if (wantedTextHeight <= smallDescHeight)
            {
                //When description height is less then align its center on parent center
                if (wantedTextHeight <= parentHeight || parentHeight <= 1.0)
                {
                    if (!IsElementAllNonFlatLines(parentElem) && !parentElem.Name.Contains(Constants.Graphics.blob) && !parentElem.Name.Contains(Constants.Graphics.unitParent)
                        && !parentElem.Name.Contains(Constants.Graphics.pointAngle) && !parentElem.Name.Contains(Constants.Graphics.pointAngleLine))
                    {
                        MoveElement(descriptionElement, elementDescriptGapWidth + parentWidth, -(parentHeight / 2.0 - wantedTextHeight / 2.0)); //Anchor is upper left but needs to be centered on unit box.
                    }
                    else if (parentElem.Name.Contains(Constants.Graphics.unitParent))
                    {
                        MoveElement(descriptionElement, elementDescriptGapWidth + parentWidth, -(parentHeight / 4.0 - wantedTextHeight / 2.0));
                    }
                    else
                    {
                        //Special case for wave and blob since it's a line with anchor in center/center but behaves like bottom center...?
                        MoveElement(descriptionElement, elementDescriptGapWidth + parentWidth, (parentHeight / 2.0) - (parentHeight - wantedTextHeight) / 2.0); //Anchor is upper left but needs to be centered on unit box
                    }

                }
                else
                {
                    MoveElement(descriptionElement, elementDescriptGapWidth + parentWidth, wantedTextHeight / 2.0);
                }

            }
            else
            {
                if (isLineOrPoint)
                {
                    if (parentHeight <= 1.0)
                    {
                        MoveElement(descriptionElement, elementDescriptGapWidth + parentWidth, (parentHeight / 2.0) + Constants.YSpacings.lineHeight0DescriptionHeightAdjustement);
                    }
                    else
                    {
                        MoveElement(descriptionElement, elementDescriptGapWidth + parentWidth, (parentHeight / 2.0) + Constants.YSpacings.lineDescriptionHeightAdjustement);
                    }

                }
                else
                {
                    MoveElement(descriptionElement, elementDescriptGapWidth + parentWidth, 0); // Anchor is upper left and needs to be horizontally aligned with it
                }


            }

            #endregion

            //Order
            OrderElement(descriptionElement);

            //Add to tracking list
            legendElementList.Add(descriptionElement);

            return descriptionElement;
        }

        /// <summary>
        /// Will detect the type of graphic and will extract the geometry out of it, for later calculation.
        /// </summary>
        /// <param name="inElement"></param>
        /// <returns></returns>
        public Geometry GetElementGeometry(Element inElement)
        {
            Geometry outGeometry = null;

            switch(inElement)
            {
                case GraphicElement graphicElement:
                    CIMGraphic graphic = graphicElement.GetGraphic();

                    switch (graphic)
                    {
                        case CIMPointGraphic pointGraphic:
                            outGeometry = pointGraphic.Location; break;
                        case CIMLineGraphic lineGraphic:
                            outGeometry = lineGraphic.Line; break;
                        case CIMPolygonGraphic polygonGraphic:
                            outGeometry = polygonGraphic.Polygon; break;
                        case CIMTextGraphic textGraphic:
                            outGeometry = textGraphic.Shape; break;
                        default:
                            break;
                    }

                    break;

                default:
                    outGeometry = inElement.GetBounds();
                    break;

            }

            return outGeometry;
        }

        /// <summary>
        /// Will symbolize thin unit by changing inner line symbol and match color to wanted map unit color.
        /// </summary>
        /// <param name="inThinUnitElement"></param>
        /// <param name="styleLineColorCode"></param>
        /// <param name="styleLineSymbolCode"></param>
        /// <returns></returns>
        public Element SetThinUnitSymbol(Element inThinUnitElement, string styleLineColorCode, string styleLineSymbolCode)
        {
            
            try
            {
                GroupElement inGroupElement = inThinUnitElement as GroupElement;
                if (inGroupElement != null)
                {
                    List<Element> groupElements = inGroupElement.GetElementsAsFlattenedList().ToList();

                    //Check geometry of inner elements, if it's all lines
                    for (int el = 0; el < groupElements.Count(); el++)
                    {
                        Element innerElement = groupElements[el];
                        GraphicElement graphicElement = innerElement as GraphicElement;

                        if (innerElement.Name.StartsWith(Constants.Graphics.subUnitLine))
                        {
                            CIMGraphic graphic = graphicElement.GetGraphic();
                            if (graphic != null && graphic is CIMLineGraphic)
                            {
                                CIMLineGraphic cimLineSymbol = graphic as CIMLineGraphic;

                                if (cimLineSymbol != null)
                                {
                                    double currentLineWidth = cimLineSymbol.Symbol.Symbol.GetSize();

                                    //Set line style
                                    if (styleLineSymbolCode != null && lineSymbolDico.ContainsKey(styleLineSymbolCode))
                                    {
                                        cimLineSymbol.Symbol.Symbol = lineSymbolDico[styleLineSymbolCode].Symbol;
                                        graphicElement.SetGraphic(graphic);
                                    }
                                    else
                                    {
                                        //Apply missing style
                                        Symbols.SetMissingLineSymbol(graphicElement);
                                    }

                                    //Set line color 
                                    if (styleLineColorCode != null && fillSymbolDico.ContainsKey(styleLineColorCode))
                                    {
                                        SymbolStyleItem fillSymbol = fillSymbolDico[styleLineColorCode];
                                        cimLineSymbol.Symbol.Symbol.SetColor(fillSymbolDico[styleLineColorCode].Symbol.GetColor());
                                        graphicElement.SetGraphic(graphic);
                                    }
                                    else
                                    {
                                        //Apply missing style
                                        Symbols.SetMissingLineSymbol(graphicElement);
                                    }
                                }
                            }
                        }
                        else
                        {
                            //Force white background in unit box (in case)
                            SetPolygonFill(innerElement, "1.01.01.001", true);
                        }

                    }
                    
                }
            }
            catch (Exception SetThinUnitSymbolException)
            {
                new ErrorService(SetThinUnitSymbolException).WriteToFile();
            }

            return inThinUnitElement;
        }

        /// <summary>
        /// Will create a marker symbol from given type, order and style. Will also return an offset parameter for linear markers
        /// </summary>
        /// <param name="markerType">element marker type (as stated in user table column element)</param>
        /// <param name="markerOrder">element order (for naming purposes)</param>
        /// <param name="markerStyle">element style</param>
        /// <param name="offset">out offset for linear markers</param>
        /// <returns></returns>
        private Element BuildMarker(Element element, string markerType, double markerOrder, string markerStyle, out Tuple<double, double> offset)
        {

            offset = new Tuple<double, double>(0,0);

            try
            {
                //Symbolize if symbol can be found in style file
                GraphicElement graphicElement = element as GraphicElement;
                if (graphicElement != null)
                {
                    CIMPointGraphic cimPoint = graphicElement.GetGraphic() as CIMPointGraphic;
                    if (cimPoint != null)
                    {
                        CIMPointSymbol pointSymbol = cimPoint.Symbol.Symbol as CIMPointSymbol;
                        
                        if (pointSymbol != null)
                        {
                            //Keep original angle (could be comming from POINT_CC_45), because style symbol doesn't have an angle by default.
                            CIMMarker originalMarker = pointSymbol.SymbolLayers[0] as CIMMarker;
                            double originalAngle = originalMarker.Rotation;

                            if (markerStyle != null && markerSymbolDico.ContainsKey(markerStyle))
                            {
                                SymbolStyleItem markerStyleItem = markerSymbolDico[markerStyle];
                                if (markerStyleItem != null && markerStyleItem.Symbol is CIMPointSymbol) 
                                {
                                    CIMPointSymbol newPointSymbol = markerStyleItem.Symbol as CIMPointSymbol;

                                    //Get new style offset
                                    CIMMarker newPointMarker = newPointSymbol.SymbolLayers[0] as CIMMarker;
                                    if (newPointMarker != null)
                                    {
                                        offset = new Tuple<double, double>(newPointMarker.OffsetX, newPointMarker.OffsetY);

                                        //Get rid of offset for linear markers
                                        if (markerType == Constants.Graphics.pointAngleLine)
                                        {
                                            newPointMarker.OffsetX = 0;
                                            newPointMarker.OffsetY = 0;
                                        }

                                        //Reset original angle
                                        newPointMarker.Rotation = originalAngle;
                                    }

                                    //Apply
                                    cimPoint.Symbol.Symbol = newPointSymbol;
                                }

                            }
                            else
                            {
                                //Apply missing style
                                cimPoint.Symbol.Symbol = Symbols.GetMissingPointSymbol();
                            }
                        }

                        graphicElement.SetGraphic(cimPoint);
                    }
                }
            }
            catch (Exception BuildMarkerException)
            {
                new ErrorService(BuildMarkerException).WriteToFile();
            }

            return element;

        }

        /// <summary>
        /// From a given element will calculate a new point geometry to fit anchor point so the element can
        /// be set at the right place on the layout before being moved.
        /// NOTE Anchor Point type doesnt' change a thing on the placement of the element.
        /// </summary>
        /// <param name="inElement"></param>
        /// <param name="inAnchor">Without embedded y spacing</param>
        /// <returns></returns>
        public void SetPointFromAnchorType(Element inElement, Tuple<double, double> inAnchor, Tuple<double, double> offset)
        {
            try
            {
                //Get info
                Coordinate2D anchorPoint = inElement.GetAnchorPoint();
                Anchor anchorPointType = inElement.GetAnchor();
                double inElementHeight = inElement.GetHeight();

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

                switch (anchorPointType)
                {
                    case Anchor.TopLeftCorner:
                        break;
                    case Anchor.TopMidPoint:
                        break;
                    case Anchor.TopRightCorner:
                        break;
                    case Anchor.LeftMidPoint:
                        break;
                    case Anchor.CenterPoint:
                        anchorPoint.X = inAnchor.Item1 + elementWidth / 2.0 + xOff;
                        anchorPoint.Y = inAnchor.Item2 + offset.Item2;
                        inElement.SetAnchorPoint(anchorPoint);
                        break;
                    case Anchor.RightMidPoint:
                        break;
                    case Anchor.BottomLeftCorner:
                        break;
                    case Anchor.BottomMidPoint:
                        anchorPoint.X = inAnchor.Item1 + elementWidth / 2.0 + yOff;
                        anchorPoint.Y = inAnchor.Item2 - inElementHeight / 2.0 + offset.Item2;
                        inElement.SetAnchorPoint(anchorPoint);
                        break;
                    case Anchor.BottomRightCorner:
                        break;
                    default:
                        break;
                }
            }
            catch (Exception SetPointFromAnchorTypeException)
            {
                new ErrorService(SetPointFromAnchorTypeException).WriteToFile();
            }

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
        private Element AddLabelToMarker(string inLabelText, Element pointElement, Tuple<double, double> inAnchor, 
            Constants.Styles.MarkerLabelPositioning wantedPosition, string inLabelStyle = "", Constants.Styles.MarkerLabelPositioning parentPosition = Constants.Styles.MarkerLabelPositioning.FromCenterToUpperLeft)
        {
            Element markerLabelElement = null;

            try
            {
                //Variables
                string inElementType = Constants.Graphics.measurementLabel;

                //Get appropriate element (measurement or generation)
                int measurementValue = -1;
                if (!Int32.TryParse(inLabelText, out measurementValue))
                {
                    inElementType = Constants.Graphics.generationLabel;
                }
                markerLabelElement = CopyElementObject(templateGraphicDico[inElementType], currentOrder.ToString());
                legendElementList.Add(markerLabelElement);

                //Set style
                TextElement labelElement = markerLabelElement as TextElement;
                if (labelElement != null && inLabelStyle != null && inLabelStyle != "")
                {
                    if (textSymbolDico.ContainsKey(inLabelStyle))
                    {
                        SymbolStyleItem inStyleSymbol = textSymbolDico[inLabelStyle];

                        if (labelElement != null)
                        {
                            CIMGraphic cimGraphic = labelElement.GetGraphic();
                            if (cimGraphic != null)
                            {
                                CIMTextSymbol cIMTextSymbol = cimGraphic.Symbol.Symbol as CIMTextSymbol;

                                if (cIMTextSymbol != null)
                                {
                                    cIMTextSymbol.SetColor(cIMTextSymbol.GetColor());
                                    cIMTextSymbol.FontFamilyName = cIMTextSymbol.FontFamilyName;
                                    cIMTextSymbol.SetSize(cIMTextSymbol.GetSize());
                                    cIMTextSymbol.VerticalAlignment = cIMTextSymbol.VerticalAlignment;
                                    labelElement.SetGraphic(cimGraphic);
                                }
                            }
                        }

                    }
                    else
                    {
                        //Missing or wrong style 
                        labelElement = Symbols.SetMissingTextSymbol(labelElement);
                    }
                }

                //Manage missing
                if (inLabelText == null || inLabelText == string.Empty || inLabelText == " ")
                {
                    labelElement = Symbols.SetMissingTextSymbol(labelElement);
                }
                labelElement.SetTextProperties(new TextProperties(inLabelText, labelElement.TextProperties.Font, labelElement.TextProperties.FontSize, labelElement.TextProperties.FontStyle));

                //Move to right anchor
                double xLabelAnchor = inAnchor.Item1;
                double yLabelAnchor = inAnchor.Item2;

                //Apply conversion factor
                double markerWidth = markerLabelElement.GetBounds().Width;
                double markerHeight = markerLabelElement.GetBounds().Height;
                double parentWidth = pointElement.GetBounds().Width;
                double parentHeight = pointElement.GetBounds().Height;

                //Original values
                Anchor originPointAnchor = pointElement.GetAnchor();

                //For label right on top of another label
                switch (wantedPosition)
                {
                    case Constants.Styles.MarkerLabelPositioning.FromCenterToUpperLeft:

                        //Set label center on upper left anchor of marker
                        pointElement.SetAnchor(Anchor.TopLeftCorner);
                        Coordinate2D topLeftCorner = pointElement.GetAnchorPoint();
                        xLabelAnchor = topLeftCorner.X;
                        yLabelAnchor = topLeftCorner.Y;
                        
                        break;

                    case Constants.Styles.MarkerLabelPositioning.FromCenterToUpperRight:

                        pointElement.SetAnchor(Anchor.TopRightCorner);
                        Coordinate2D topRighCorner = pointElement.GetAnchorPoint();
                        xLabelAnchor = (topRighCorner.X + markerWidth / 2.0) - 0.5;
                        yLabelAnchor = (topRighCorner.Y + markerHeight / 2.0) - 0.5;

                        break;

                    case Constants.Styles.MarkerLabelPositioning.FromCenterToUpperRightTight:

                        pointElement.SetAnchor(Anchor.TopRightCorner);
                        Coordinate2D topRightCorner = pointElement.GetAnchorPoint();
                        xLabelAnchor = (topRightCorner.X + markerWidth / 2.0) - 0.75;
                        yLabelAnchor = (topRightCorner.Y + markerHeight / 2.0) - 0.75;

                        break;

                    //This case is meant for when two labels must be added around a marker point
                    case Constants.Styles.MarkerLabelPositioning.RightAboveCenter:

                        //Force y move on parent for a better fit of the two labels
                        if (parentPosition == Constants.Styles.MarkerLabelPositioning.FromCenterToUpperLeft)
                        {
                            MoveElement(pointElement, -0.47, -parentHeight * 0.5);
                        }
                        else
                        {
                            MoveElement(pointElement, 0, -parentHeight * 0.5);
                        }

                        //Value were found from manually placing the label at wanted place and calculating the ratio for the best move. 
                        xLabelAnchor = pointElement.GetBounds().XMin + parentWidth / 2.0;
                        yLabelAnchor = pointElement.GetBounds().YMax + markerHeight / 4.0;

                        break;

                    default:
                        break;
                }

                //Reset anchor to original
                pointElement.SetAnchor(originPointAnchor);

                Tuple<double, double> labelAnchor = new Tuple<double, double>(xLabelAnchor, yLabelAnchor);
                PositionElement(markerLabelElement, labelAnchor.Item1, labelAnchor.Item2);

            }
            catch (Exception AddLabelToMarkerException)
            {
                new ErrorService(AddLabelToMarkerException).WriteToFile();
            }

            return markerLabelElement;
        }

        /// <summary>
        /// Will automatically calculate the column number needed. Usually, meant for CGM mapping 
        /// publication where the legend is confined within a given box.
        /// </summary>
        /// <returns></returns>
        private async Task AutomaticColumnsCalculation()
        {
            try
            {
                if (_legendAutoCalculateColumn)
                {
                    //Track column change with auto-calculate
                    if (legendYLowerBound != 0.0)
                    {
                        if ((anchorPoint.Item2 - ySpacing - currentElementObject.GetHeight()) < legendYLowerBound)
                        {
                            currentColumn++;
                        }
                    }

                }
            }
            catch (Exception AutomaticColumnsCalculationException)
            {
                new ErrorService(AutomaticColumnsCalculationException).WriteToFile();
            }
        }
        #endregion

        #region ADD GRAPHIC METHODS

        /// <summary>
        /// Will add a heading element to the legend based on the current row information
        /// </summary>
        /// <returns></returns>
        private async Task AddHeading()
        {
            try
            {
                if (currentElementObject != null && currentElementName.Contains(Constants.Graphics.heading1.Substring(0, 6)))
                {

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
                        tElement = Symbols.SetMissingTextSymbol(tElement);
                    }

                    //Manage style if needed
                    if (currentStyle1 != null && currentStyle1 != "")
                    {
                        if (textSymbolDico.ContainsKey(currentStyle1))
                        {
                            SymbolStyleItem inStyleSymbol = textSymbolDico[currentStyle1];

                            CIMGraphic cimGraphic = tElement.GetGraphic();
                            if (cimGraphic != null)
                            {
                                CIMTextSymbol cIMTextSymbol = cimGraphic.Symbol.Symbol as CIMTextSymbol;
                                CIMTextSymbol styleText = inStyleSymbol.Symbol as CIMTextSymbol;
                                if (cIMTextSymbol != null && styleText != null)
                                {
                                    cIMTextSymbol.SetColor(inStyleSymbol.Symbol.GetColor());
                                    cIMTextSymbol.FontFamilyName = styleText.FontFamilyName;
                                    cIMTextSymbol.SetSize(styleText.GetSize());
                                    cIMTextSymbol.VerticalAlignment = styleText.VerticalAlignment;
                                    tElement.SetGraphic(cimGraphic);
                                }

                            }
                        }
                        else
                        {
                            //Missing or wrong style 
                            tElement = Symbols.SetMissingTextSymbol(tElement);
                        }
                    }
                }
            }
            catch (Exception AddHeadingException)
            {
                new ErrorService(AddHeadingException).WriteToFile();
            }

        }

        /// <summary>
        /// Will add a map unit graphic
        /// </summary>
        /// <param name="mapUnitRow"></param>
        /// <returns></returns>
        private async Task AddMapUnit()
        {
            try
            {
                if (currentElementObject != null && (currentElementName == Constants.Graphics.unitBox || currentElementName == Constants.Graphics.unitSplit ||
                            currentElementName == Constants.Graphics.unitindent1 || currentElementName == Constants.Graphics.unitindent2))
                {

                    //Keep some information
                    Element originalParent = currentElementObject;

                    //Init empty element if ever needed for edge cases
                    Element demUnitBoxElement = null;
                    Element labelUnitBoxElement = null;
                    Element labelUnitBoxElement2 = null;

                    #region Move to right anchor

                    //Set new anchor and position
                    anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing); //New anchor point with proper move inside it

                    if (currentElementName != Constants.Graphics.unitSplit)
                    {
                        PositionElement(currentElementObject, anchorPoint.Item1, anchorPoint.Item2);
                    }
                    else
                    {
                        PositionElement(currentElementObject, anchorPoint.Item1, anchorPoint.Item2, Anchor.TopLeftCorner);
                    }

                    #endregion

                    //Manage label
                    if (currentLabel1 == null || currentLabel1 == string.Empty || currentLabel1 == " ")
                    {
                        currentLabel1 = Constants.TextConfiguration.missingText;
                    }

                    if (currentLabel2 == null || currentLabel2 == string.Empty || currentLabel2 == " ")
                    {
                        currentLabel2 = Constants.TextConfiguration.missingText;
                    }

                    //Symbolize
                    GroupElement inGroupElement = currentElementObject as GroupElement;
                    if (inGroupElement != null)
                    {
                        //Check geometry of inner elements, if it's all lines
                        List<Element> groupElements = inGroupElement.GetElementsAsFlattenedList().ToList();

                        for (int el = 0; el < groupElements.Count(); el++)
                        {
                            Element innerElement = groupElements[el];

                            if (el == 0)
                            {
                                SetPolygonFill(innerElement, currentStyle1, true);
                                labelUnitBoxElement = AddLabelInUnitBox(currentLabel1, innerElement, anchorPoint, Constants.Graphics.UnitBoxType.split1, currentLabel1Style);

                            }
                            else if (el > 0)
                            {
                                SetPolygonFill(innerElement, currentStyle2, true);
                                labelUnitBoxElement2 = AddLabelInUnitBox(currentLabel2, innerElement, anchorPoint, Constants.Graphics.UnitBoxType.split2, currentLabel2Style);
                            }

                        }
                    }
                    else
                    {
                        //Symbolize
                        labelUnitBoxElement = AddLabelInUnitBox(currentLabel1, currentElementObject, anchorPoint, Constants.Graphics.UnitBoxType.normal, currentLabel1Style);
                        demUnitBoxElement = SetPolygonFill(currentElementObject, currentStyle1, true, true, anchorPoint, currentStyle2);
                    }

                    //Add header if needed
                    if (currentHeading != null && currentHeading != string.Empty && currentHeading != " ")
                    {
                        currentDescription = Constants.TextConfiguration.tagBold + currentHeading + Constants.TextConfiguration.endTagBold + " " + currentDescription;
                    }

                    //Add Description
                    Element newDescriptionElement = AddDescription(currentDescription, originalParent, anchorPoint, currentElementName);

                    //Rest anchor point for next element
                    double descriptionHeight = Constants.TextConfiguration.lineHeight;
                    if (newDescriptionElement is GroupElement)
                    {
                        GroupElement newDescriptGroup = newDescriptionElement as GroupElement;
                        descriptionHeight = newDescriptGroup.GetElementsAsFlattenedList().First().GetHeight();
                    }
                    else
                    {
                        descriptionHeight = newDescriptionElement.GetHeight();
                    }

                    if (descriptionHeight > smallDescriptionHeight)
                    {
                        if (currentColumn != 0)
                        {
                            anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - descriptionHeight); //New anchor point with proper move inside it

                        }

                        //Keep name
                        lastElement = newDescriptionElement;
                        lastElementType = Constants.Graphics.description;

                    }

                    //CASE - UNIT INDENT: Move items (label, description and unit box) for unit indent items
                    if (currentElementName == Constants.Graphics.unitindent1 || currentElementName == Constants.Graphics.unitindent2)
                    {
                        MoveElement(newDescriptionElement, xSpacing, 0);
                        MoveElement(labelUnitBoxElement, xSpacing, 0);
                        MoveElement(demUnitBoxElement, xSpacing, 0);

                        if (_legendDEM && demPictureElementObject != null)
                        {
                            MoveElement(demPictureElementObject, xSpacing, 0);
                        }
                    }

                    //CASE - UNIT DEM needs to be send all the way back
                    if (_legendDEM && demPictureElementObject != null && demUnitBoxElement != null)
                    {
                        demPictureElementObject.SetTOCPositionRelative(demUnitBoxElement, false);
                    }

                    //Keep element if for bracket
                    if (currentColumn == 0)
                    {
                        bracketMapUnit = new Tuple<Element, Element, Element, Element>(demUnitBoxElement, labelUnitBoxElement, newDescriptionElement, demUnitBoxElement);

                        //Reset anchor point
                        anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 + ySpacing);
                    }

                }
            }
            catch (Exception AddMapUnitException)
            {
                new ErrorService(AddMapUnitException).WriteToFile();
            }
        }

        /// <summary>
        /// Will add a thin unit graphic
        /// </summary>
        /// <returns></returns>
        public async Task AddThinUnit()
        {
            try
            {
                if (currentElementObject != null && currentElementName == Constants.Graphics.unitLine)
                {
                    Element originalParent = currentElementObject;

                    //Set new anchor
                    anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing); //New anchor point with proper move inside it
                    PositionElement(currentElementObject, anchorPoint.Item1, anchorPoint.Item2, Anchor.TopLeftCorner); 

                    Element thinUnitElement = SetThinUnitSymbol(currentElementObject, currentStyle1, currentStyle2);

                    //Add label if needed
                    if (currentLabel1 != null && currentLabel1 != string.Empty && currentLabel1 != " ")
                    {
                        Element thinUnitLabel = AddLabelInUnitBox(currentLabel1, thinUnitElement, anchorPoint, Constants.Graphics.UnitBoxType.line, currentLabel1Style);
                    }

                    //Add header if needed
                    if (currentHeading != null && currentHeading != string.Empty && currentHeading != " ")
                    {
                        currentDescription = Constants.TextConfiguration.tagBold + currentHeading + Constants.TextConfiguration.endTagBold + " " + currentDescription;
                    }

                    //Add Description
                    Element newDescriptionElement = AddDescription(currentDescription, originalParent, anchorPoint, currentElementName);
                    double descriptionHeight = currentElementObject.GetHeight();
                    if (descriptionHeight > smallDescriptionHeight)
                    {
                        //Reset anchor point for next element
                        anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - descriptionHeight); //New anchor point with proper move inside it

                    }

                }
            }
            catch (Exception AddThinUnitException)
            {
                new ErrorService(AddThinUnitException).WriteToFile();
            }
        }

        /// <summary>
        /// Will add embedded units (parents, child and child line)
        /// </summary>
        /// <returns></returns>
        public async Task AddEmbeddedMapUnit()
        {
            try
            {
                if (currentElementObject != null && (currentElementName == Constants.Graphics.unitParent ||
                    currentElementName == Constants.Graphics.subUnitParentChild ||
                    currentElementName == Constants.Graphics.subUnitParentChildLine))
                {

                    //Reset element and grow parent if needed
                    if (currentElementName == Constants.Graphics.unitParent)
                    {
                        //Reset parent element
                        parentElement = null;
                    }

                    string originalElementName = currentElementName;

                    //Apply conversion factor
                    double parentHeight = 0;
                    double parentChildHeight = currentElementObject.GetHeight();

                    //Set new anchor
                    if (parentElement != null && lastElement == parentElement)
                    {
                        parentHeight = parentElement.GetHeight();
                        double newYSpacing = ySpacing + (parentHeight - parentChildHeight - templateGraphicDico[Constants.Graphics.unitBox].GetHeight());
                        anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - newYSpacing); //New anchor point with proper move inside it
                    }
                    else
                    {
                        anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing); //New anchor point with proper move inside it
                    }

                    if ((currentElementName == Constants.Graphics.subUnitParentChild || currentElementName == Constants.Graphics.subUnitParentChildLine) && parentElement != null)
                    {
                        if (lastElement != parentElement)
                        {
                            //Resize parent to match addition of child
                            parentHeight = parentElement.GetHeight();

                            double newHeightFromChild = parentChildHeight + parentHeight;

                            SetRectangularPolygonFromAnchorTypeAndHeight(parentElement, anchorPointParent, newHeightFromChild);

                            //Reset anchor point since height of the element has changed.
                            anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - parentChildHeight);

                            //Enforce position, especially for unit child line which has an anchor bottom left instead of top left
                            PositionElement(currentElementObject, anchorPoint.Item1, anchorPoint.Item2, Anchor.TopLeftCorner);
                        }

                    }

                    //Resize
                    SetRectangularPolygonFromAnchorType(currentElementObject, anchorPoint);

                    //Move
                    MoveElement(currentElementObject, xSpacing, 0);

                    //Symbolize
                    Element labelParentChild = null;
                    Constants.Graphics.UnitBoxType labelType = Constants.Graphics.UnitBoxType.normal;
                    if (currentElementName == Constants.Graphics.subUnitParentChildLine)
                    {
                        SetThinUnitSymbol(currentElementObject, currentStyle1, currentStyle2);
                        labelType = Constants.Graphics.UnitBoxType.child_line;
                    }
                    else if (currentElementName == Constants.Graphics.unitParent)
                    {
                        SetPolygonFill(currentElementObject, currentStyle1, true);
                        labelType = Constants.Graphics.UnitBoxType.parent;
                    }
                    else
                    {
                        SetPolygonFill(currentElementObject, currentStyle1, true);
                        labelType = Constants.Graphics.UnitBoxType.normal;
                    }

                    //Labelize
                    if (currentLabel1 != null && currentLabel1 != string.Empty && currentLabel1 != " ")
                    {
                        labelParentChild = AddLabelInUnitBox(currentLabel1, currentElementObject, anchorPoint, labelType, currentLabel1Style);
                    }
                    else
                    {
                        labelParentChild = AddLabelInUnitBox(Constants.TextConfiguration.missingText, currentElementObject, anchorPoint, labelType, currentLabel1Style);
                    }

                    //Move label
                    if (labelParentChild != null)
                    {
                        MoveElement(labelParentChild, xSpacing, 0);
                    }

                    //Add header if needed
                    if (currentHeading != null && currentHeading != string.Empty && currentHeading != " ")
                    {
                        currentDescription = Constants.TextConfiguration.tagBold + currentHeading + Constants.TextConfiguration.endTagBold + " " + currentDescription;
                    }

                    //Add Description
                    Element newDescriptionElement = AddDescription(currentDescription, currentElementObject, anchorPoint, originalElementName);
                    double descriptionHeight = newDescriptionElement.GetHeight();
                    if (descriptionHeight > smallDescriptionHeight)
                    {
                        //Keep name
                        lastElement = newDescriptionElement;
                        lastElementType = Constants.Graphics.description;

                        //Reset height of unit parent
                        if (currentElementName == Constants.Graphics.unitParent)
                        {
                            SetRectangularPolygonFromAnchorTypeAndHeight(currentElementObject, anchorPoint, descriptionHeight);
                        }

                        //Reset anchor for next unit to be added
                        if (currentElementName == Constants.Graphics.subUnitParentChild || currentElementName == Constants.Graphics.subUnitParentChildLine)
                        {
                            double newDescriptionHeight = descriptionHeight - smallDescriptionHeight;
                            double newParentHeight = parentElement.GetHeight() + newDescriptionHeight;
                            SetRectangularPolygonFromAnchorTypeAndHeight(parentElement, anchorPointParent, newParentHeight); //Reset parent box height
                            anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - newDescriptionHeight); //New anchor point with proper move inside it
                        }

                    }

                    //Keep parent information
                    if (currentElementName == Constants.Graphics.unitParent)
                    {
                        parentElement = currentElementObject;
                        anchorPointParent = anchorPoint;
                    }

                }
            }
            catch (Exception AddEmbeddedMapUnitException)
            {
                new ErrorService(AddEmbeddedMapUnitException).WriteToFile();
            }
        }

        /// <summary>
        /// Will add points/markers symbols
        /// </summary>
        /// <returns></returns>
        public async Task AddMarker()
        {
            try
            {
                if (currentElementObject != null && (currentElementName == Constants.Graphics.point || currentElementName == Constants.Graphics.pointAngle || currentElementName == Constants.Graphics.pointAngleLine))
                {
                    //Build and symbolize marker element
                    Tuple<double, double> offset = null;
                    Element pointElement = BuildMarker(currentElementObject, currentElementName, currentOrder, currentStyle1, out offset);

                    //Set new anchor
                    anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing); //New anchor point with proper move inside it
                    SetPointFromAnchorType(pointElement, anchorPoint, offset);

                    //Add measurement value label
                    if (currentLabel1 != null && currentLabel1 != string.Empty && currentLabel1 != " ")
                    {

                        //Find proper placement for label
                        Constants.Styles.MarkerLabelPositioning placement = Constants.Styles.MarkerLabelPositioning.FromCenterToUpperLeft;
                        if (currentElementName == Constants.Graphics.pointAngleLine)
                        {
                            placement = Constants.Styles.MarkerLabelPositioning.FromCenterToUpperRight;
                        }
                        else if (currentElementName == Constants.Graphics.point)
                        {
                            placement = Constants.Styles.MarkerLabelPositioning.FromCenterToUpperRightTight;
                        }

                        if (currentLabel1Style == null || currentLabel1Style == " ")
                        {
                            currentLabel1Style = string.Empty;
                        }
                        Element markerLabel1 = AddLabelToMarker(currentLabel1, pointElement, anchorPoint, placement, currentLabel1Style);

                        //Add second label if any
                        if (currentLabel2 != null && currentLabel2 != string.Empty && currentLabel2 != " ")
                        {
                            if (currentLabel2Style == null || currentLabel2Style == " ")
                            {
                                currentLabel2Style = string.Empty;
                            }

                            AddLabelToMarker(currentLabel2, markerLabel1, anchorPoint, Constants.Styles.MarkerLabelPositioning.RightAboveCenter, currentLabel2Style, placement);
                        }
                    }

                    //Add Description
                    Element newDescriptionElement = AddDescription(currentDescription, pointElement, anchorPoint, lastElementType, true);
                    double descriptionHeight = newDescriptionElement.GetHeight();
                    if (descriptionHeight > smallDescriptionHeightLine)
                    {
                        //Reset anchor point for next element
                        double descriptionAdjustement = descriptionHeight - Constants.YSpacings.markerMeanHeight;

                        anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - descriptionAdjustement); //New anchor point with proper move inside it

                    }
                }
            }
            catch (Exception AddMarkersException)
            {
                new ErrorService(AddMarkersException).WriteToFile();
            }
        }

        /// <summary>
        /// Will add a line symbol
        /// </summary>
        /// <returns></returns>
        public async Task AddLine()
        {
            try
            {
                if (currentElementObject != null && (currentElementName == Constants.Graphics.beach || currentElementName == Constants.Graphics.moraines || 
                    currentElementName == Constants.Graphics.dunes || currentElementName == Constants.Graphics.landslide || 
                    currentElementName == Constants.Graphics.line || currentElementName == Constants.Graphics.wave || 
                    currentElementName == Constants.Graphics.lineDouble || currentElementName == Constants.Graphics.lineDoubleFLow || 
                    currentElementName == Constants.Graphics.lineDoubleFlip))
                {
                    //GLITCH - All line graphics are set to center point anchor and are copied as bottom left, force it back
                    currentElementObject.SetAnchor(Anchor.CenterPoint);

                    //Set new anchor
                    anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing); 
                    PositionElement(currentElementObject, anchorPoint.Item1, anchorPoint.Item2, Anchor.BottomLeftCorner);

                    //Move on x axis smaller symbols and those that have a different anchor
                    if (currentElementName == Constants.Graphics.beach || currentElementName == Constants.Graphics.dunes ||
                        currentElementName == Constants.Graphics.landslide || currentElementName == Constants.Graphics.moraines)
                    {
                        MoveElement(currentElementObject, (xSpacings.ELEMENT_WIDTH / 2.0) - currentElementObject.GetWidth() / 2.0, 0, Anchor.BottomLeftCorner);
                    }

                    //Set symbol
                    GraphicElement graphicLine = currentElementObject as GraphicElement;
                    if (currentStyle1 != null && lineSymbolDico.ContainsKey(currentStyle1))
                    {
                        SymbolStyleItem lineStyleItem = lineSymbolDico[currentStyle1];

                        GroupElement inGroupElement = currentElementObject as GroupElement;
                        if (inGroupElement != null)
                        {
                            //Check geometry of inner elements, if it's all lines
                            for (int el = 0; el < inGroupElement.Elements.Count(); el++)
                            {
                                Element currentLine = inGroupElement.Elements[el];
                                GraphicElement currentLineGraphic = currentLine as GraphicElement;
                                if (currentLineGraphic != null)
                                {
                                    CIMLineGraphic cimLine = currentLineGraphic.GetGraphic() as CIMLineGraphic;
                                    if (cimLine != null)
                                    {
                                        //Generic symbolize
                                        if (currentLine.Name != Constants.Graphics.subLineDoubleFLowBottom && currentLine.Name != Constants.Graphics.subLineDoubleFLowMiddle)
                                        {
                                            if (lineStyleItem != null)
                                            {
                                                CIMLineSymbol newLineSymbol = lineStyleItem.Symbol as CIMLineSymbol;
                                                cimLine.Symbol.Symbol = newLineSymbol;  
                                                currentLineGraphic.SetGraphic(cimLine);
                                            }
                                            else
                                            {
                                                Symbols.SetMissingLineSymbol(currentLineGraphic);
                                            }
                                        }

                                        // EDGE CASE 
                                        if (currentLine.Name == Constants.Graphics.subLineDoubleFLowBottom)
                                        {
                                            //If something isn't found in style2 revert to first one
                                            if (currentStyle2 != null && currentStyle2 != string.Empty && currentStyle2 != " " && lineSymbolDico.ContainsKey(currentStyle2))
                                            {
                                                // For double line, two style might be used if it's not inside a double line flow symbol
                                                SymbolStyleItem lineStyleItem2 = lineSymbolDico[currentStyle2];
                                                CIMLineSymbol newLineSymbol2 = lineStyleItem2.Symbol as CIMLineSymbol;
                                                cimLine.Symbol.Symbol = newLineSymbol2;
                                                currentLineGraphic.SetGraphic(cimLine);

                                                if (currentElementName == Constants.Graphics.lineDoubleFLow)
                                                {
                                                    //For double line flow symbol keep bottom line just like the top one and take style2 field for the flow symbol
                                                    CIMLineSymbol newLineSymbol = lineStyleItem.Symbol as CIMLineSymbol;
                                                    cimLine.Symbol.Symbol = newLineSymbol;
                                                    currentLineGraphic.SetGraphic(cimLine);
                                                }

                                            }
                                            else
                                            {
                                                Symbols.SetMissingLineSymbol(currentLineGraphic);
                                            }
                                        }

                                        // EDGE CASE 
                                        if (currentLine.Name == Constants.Graphics.subLineDoubleFLowMiddle)
                                        {
                                            if (currentStyle2 != null && lineSymbolDico.ContainsKey(currentStyle2))
                                            {
                                                SymbolStyleItem lineStyleItem2 = lineSymbolDico[currentStyle2];
                                                CIMLineSymbol newLineSymbol2 = lineStyleItem2.Symbol as CIMLineSymbol;
                                                cimLine.Symbol.Symbol = newLineSymbol2;
                                                currentLineGraphic.SetGraphic(cimLine);
                                            }
                                            else
                                            {
                                                Symbols.SetMissingLineSymbol(currentLineGraphic);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            CIMLineSymbol newLineSymbol = lineStyleItem.Symbol as CIMLineSymbol;
                            CIMLineGraphic cimLine = graphicLine.GetGraphic() as CIMLineGraphic;
                            if (cimLine != null)
                            {
                                cimLine.Symbol.Symbol = newLineSymbol;
                                graphicLine.SetGraphic(cimLine);
                            }
                        }
                    }
                    else
                    {
                        Symbols.SetMissingLineSymbol(currentElementObject);
                    }

                    //Add Description
                    //NOTES: Usually lines have anchors in the center, that needs special attention
                    Element newDescriptionElement = AddDescription(currentDescription, currentElementObject, anchorPoint, lastElementType, true);

                    //Move lines in y axis else AddDescription doesn't align them properly
                    MoveElement(newDescriptionElement, 0, currentElementObject.GetHeight() / 2.0, Anchor.TopLeftCorner);

                    double descriptionHeight = newDescriptionElement.GetHeight();
                    if (descriptionHeight >= smallDescriptionHeightLine)
                    {
                        double descriptionAdjustement = descriptionHeight;
                        if (currentElementObject.GetHeight() == 0)
                        {
                            descriptionAdjustement = descriptionAdjustement - Constants.YSpacings.lineHeight0DescriptionHeightAdjustement;
                        }
                        else
                        {
                            descriptionAdjustement = descriptionAdjustement - (currentElementObject.GetHeight() / 2.0);
                        }

                        //Reset anchor point for next element
                        anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - descriptionAdjustement); //New anchor point with proper move inside it
                    }

                    if (currentElementName == Constants.Graphics.line)
                    {
                        anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - 4.0);
                    }

                }
            }
            catch (Exception AddLineException)
            {
                new ErrorService(AddLineException).WriteToFile();
            }
        }

        /// <summary>
        /// Will add a polygon symbol tha represents overlays and blobs
        /// </summary>
        /// <returns></returns>
        public async Task AddOverlay()
        {
            try
            {
                if (currentElementObject != null && (currentElementName == Constants.Graphics.overlay || currentElementName == Constants.Graphics.blob))
                {

                    //Set new anchor
                    anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing); //New anchor point with proper move inside it
                    PositionElement(currentElementObject, anchorPoint.Item1, anchorPoint.Item2);

                    //Move on x axis smaller symbols and those that have a different anchor
                    if (currentElementName == Constants.Graphics.blob)
                    {
                        MoveElement(currentElementObject, (xSpacings.ELEMENT_WIDTH / 2.0), 0, Anchor.LeftMidPoint);
                    }

                    //Symbolize
                    SetPolygonFill(currentElementObject, currentStyle1, false);

                    //Add annotation or a marker in the middle if needed
                    if ((currentLabel1 != null && currentLabel1Style != string.Empty) || (currentStyle2 != null && currentStyle2 != string.Empty))
                    {
                        Tuple<double, double> offset = new Tuple<double, double>(0, 0);

                        //For a label in the middle
                        if (currentLabel1Style != null && textSymbolDico != null && textSymbolDico.Count() > 0 && textSymbolDico.ContainsKey(currentLabel1Style))
                        {

                            //Get appropriate element
                            Element symAreaPointElement = CopyElementObject(templateGraphicDico[Constants.Graphics.annotationBlob], currentOrder.ToString());
                            legendElementList.Add(symAreaPointElement);

                            SymbolStyleItem inStyleSymbol = textSymbolDico[currentLabel1Style];

                            GraphicElement pointGraphicElement = symAreaPointElement as GraphicElement;
                            CIMGraphic cimGraphic = pointGraphicElement.GetGraphic();
                            if (cimGraphic != null)
                            {
                                CIMMultiLayerSymbol cIMLayerSymbol = cimGraphic.Symbol.Symbol as CIMMultiLayerSymbol;

                                if (cIMLayerSymbol != null)
                                {
                                    CIMCharacterMarker cimCharact = cIMLayerSymbol.SymbolLayers[0] as CIMCharacterMarker;

                                    int labelCharset = 0;
                                    if (int.TryParse(currentLabel1, out labelCharset))
                                    {
                                        cimCharact.CharacterIndex = Convert.ToInt16(currentLabel1);

                                        //Make sure needed style font style is copied over new annotation, else leave default
                                        CIMTextSymbol inStyleTextSymbol = inStyleSymbol.Symbol as CIMTextSymbol;
                                        cimCharact.FontFamilyName = inStyleTextSymbol.FontFamilyName;
                                    }
                                    else
                                    {
                                        cimCharact.CharacterIndex = labelCharset;
                                    }

                                    pointGraphicElement.SetGraphic(cimGraphic);

                                }

                            }

                            //Set new anchor
                            SetPointFromAnchorType(symAreaPointElement, anchorPoint, offset);

                        }


                        //For a marker in the middle
                        if (currentStyle2 != null && markerSymbolDico != null && markerSymbolDico.Count > 0 && markerSymbolDico.ContainsKey(currentStyle2))
                        {
                            //Build marker element
                            Element symAreaPointElement = CopyElementObject(templateGraphicDico[Constants.Graphics.point], currentOrder.ToString());
                            symAreaPointElement = BuildMarker(symAreaPointElement, Constants.Graphics.point, currentOrder, currentStyle2, out offset);
                            legendElementList.Add(symAreaPointElement);

                            //Set new anchor
                            Tuple<double, double> overlayPointAnchor = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - currentElementObject.GetHeight() / 2.0);
                            SetPointFromAnchorType(symAreaPointElement, overlayPointAnchor, offset);
                        }

                        //For a new color fill on the symbol
                        if (currentStyle2 != null && fillSymbolDico != null && fillSymbolDico.Count > 0 && fillSymbolDico.ContainsKey(currentStyle2))
                        {
                            //Symbolize
                            Element symAreaPointElement = SetPolygonFill(currentElementObject, currentStyle1, false, false, anchorPoint, currentStyle2);
                        }

                    }

                    Element newDescriptionElement = AddDescription(currentDescription, currentElementObject, anchorPoint, currentElementName);
                    double descriptionHeight = currentElementObject.GetHeight();
                    if (descriptionHeight > smallDescriptionHeight)
                    {
                        //Reset anchor point for next element
                        anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - descriptionHeight); //New anchor point with proper move inside it

                    }
                }

            }
            catch (Exception AddOverlayException)
            {
                new ErrorService(AddOverlayException).WriteToFile();
            }
        }

        /// <summary>
        /// Will add a break line used mainly for headers
        /// </summary>
        /// <returns></returns>
        public async Task AddBreakLine()
        {
            try
            {
                if (currentElementObject != null && currentElementName == Constants.Graphics.breakLine)
                {
                    //Set new anchor
                    anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing); //New anchor point with proper move inside it
                    PositionElement(currentElementObject, anchorPoint.Item1, anchorPoint.Item2);

                    //Set symbol
                    if (currentStyle1 != null && currentStyle1 != string.Empty && currentStyle1 != " ")
                    {
                        GraphicElement graphicElement = currentElementObject as GraphicElement;
                        if (graphicElement != null)
                        {
                            CIMGraphic cimGraphic = graphicElement.GetGraphic();
                            if (cimGraphic != null)
                            {
                                CIMLineGraphic cIMLineGraphic = cimGraphic as CIMLineGraphic;
                                if (cIMLineGraphic != null)
                                {
                                    if (lineSymbolDico.ContainsKey(currentStyle1))
                                    {
                                        CIMLineSymbol newLineSymbol = lineSymbolDico[currentStyle1].Symbol as CIMLineSymbol;
                                        cIMLineGraphic.Symbol.Symbol = newLineSymbol;
                                        graphicElement.SetGraphic(cIMLineGraphic);
                                    }
                                    else
                                    {
                                        Symbols.SetMissingLineSymbol(graphicElement);
                                    }
                                }
                            }
                        }
                    }

                }
            }
            catch (Exception AddBreakLineException)
            {
                new ErrorService(AddBreakLineException).WriteToFile();
            }    
        }

        /// <summary>
        /// Will add a break line annotation if any needed
        /// </summary>
        /// <returns></returns>
        public async Task AddBreakLineAnnotation()
        {
            try
            {
                if ( currentElementObject != null && currentElementName == Constants.Graphics.annotationBreak)
                {
                    //Set new anchor
                    anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing);
                    PositionElement(currentElementObject, anchorPoint.Item1, anchorPoint.Item2);

                    //Move
                    MoveElement(currentElementObject, xSpacing, 0);

                    //Set text
                    TextElement tElement = currentElementObject as TextElement;
                    if (tElement != null)
                    {
                        tElement.SetTextProperties(new TextProperties(currentLabel1, tElement.TextProperties.Font, tElement.TextProperties.FontSize, tElement.TextProperties.FontStyle));
                        if (currentLabel1 == null || currentLabel1 == string.Empty || currentLabel1 == " ")
                        {
                            tElement = Symbols.SetMissingTextSymbol(tElement);
                        }
                    }
                }
            }
            catch (Exception AddBreakLineAnnotationException)
            {
                new ErrorService(AddBreakLineAnnotationException).WriteToFile();
            }
        }

        /// <summary>
        /// Will add a left bracket that wraps around a block a unit boxes
        /// </summary>
        /// <returns></returns>
        public async Task AddLeftBracket()
        {
            try
            {
                //Process left bracket that was waiting to have Y spacing
                if (waitingLeftBracket != null && !currentElementName.Contains(Constants.Graphics.keywordBracket))
                {
                    MoveElement(waitingLeftBracket, 0, -ySpacing);
                    waitingLeftBracket = null;
                }

                #region UPPER ELEMENT
                if (currentElementName == Constants.Graphics.bracketLeftUpper)
                {
                    //Get Y spacing
                    double upperElementY = anchorPoint.Item2;
                    if (lastElement != null)
                    {
                        upperElementY = lastElement.GetBounds().YMax;
                    }

                    //Set new anchor
                    Tuple<double, double> leftBracketAnchor = new Tuple<double, double>(anchorPoint.Item1, upperElementY);
                    PositionElement(currentElementObject, leftBracketAnchor.Item1, leftBracketAnchor.Item2, Anchor.TopRightCorner);

                    //Move
                    MoveElement(currentElementObject, GetXSpacing(currentElementName), 0);

                    //Keep in waiting line, to be move in Y axis
                    waitingLeftBracket = currentElementObject;
                    upLeftBracket = currentElementObject;
                }

                #endregion

                #region LOWER ELEMENT
                if (currentElementName == Constants.Graphics.bracketLeftLower && lastElement != null)
                {
                    #region END BRACKET NOTCH

                    Element leftEndBracketElement = currentElementObject;

                    //Set new anchor
                    Tuple<double, double> bottomBracketAnchor = new Tuple<double, double>(anchorPoint.Item1, lastElement.GetBounds().YMin + leftEndBracketElement.GetHeight());
                    PositionElement(leftEndBracketElement, bottomBracketAnchor.Item1, bottomBracketAnchor.Item2);

                    //Move
                    MoveElement(leftEndBracketElement, GetXSpacing(currentElementName), 0);

                    #endregion

                    #region MIDDLE BRACKET NOTCH
                    Element middleBracketElement = CopyElementObject(templateGraphicDico[Constants.Graphics.bracketLeftCenter], currentOrder.ToString());

                    //Set new anchor
                    double middleBracketY = upLeftBracket.GetBounds().YMin - Math.Abs(((leftEndBracketElement.GetBounds().YMax - upLeftBracket.GetBounds().YMin) / 2.0));
                    Tuple<double, double> middleBracketAnchor = new Tuple<double, double>(anchorPoint.Item1, middleBracketY);
                    PositionElement(middleBracketElement, middleBracketAnchor.Item1, middleBracketAnchor.Item2, Anchor.RightMidPoint);

                    //Move
                    MoveElement(middleBracketElement, GetXSpacing(Constants.Graphics.bracketLeftCenter), 0);

                    //Keep element to use with bracket annotation
                    waitingCenterLeftBracket = middleBracketElement;

                    //Add to legend list
                    legendElementList.Add(middleBracketElement);

                    #endregion

                    #region BRACKET SPINE 1
                    Element spine1BracketElement = CopyElementObject(templateGraphicDico[Constants.Graphics.bracketSpine], currentOrder.ToString(), "_UPPER");

                    //Resize and adjust straight line so that it touches the bracket curly lines
                    Coordinate2D startPointLine = new Coordinate2D(upLeftBracket.GetBounds().XMin, upLeftBracket.GetBounds().YMin);
                    Coordinate2D endPointLine = new Coordinate2D(middleBracketElement.GetBounds().XMax, middleBracketElement.GetBounds().YMax);
                    List<Coordinate2D> spineLineVertices = new List<Coordinate2D>() { startPointLine, endPointLine };
                    Polyline spineLine = PolylineBuilderEx.CreatePolyline(spineLineVertices);

                    spine1BracketElement.SetGeometry(spineLine);

                    //Add to legend list
                    legendElementList.Add(spine1BracketElement);

                    #endregion

                    #region BRACKET SPINE 2
                    Element spine1BracketElement2 = CopyElementObject(templateGraphicDico[Constants.Graphics.bracketSpine], currentOrder.ToString(), "_LOWER");

                    //Resize and adjust straight line so that it touches the bracket curly lines
                    Coordinate2D startPointLine2 = new Coordinate2D(middleBracketElement.GetBounds().XMax, middleBracketElement.GetBounds().YMin);
                    Coordinate2D endPointLine2 = new Coordinate2D(leftEndBracketElement.GetBounds().XMin, leftEndBracketElement.GetBounds().YMax);
                    List<Coordinate2D> spineLineVertices2 = new List<Coordinate2D>() { startPointLine2, endPointLine2 };
                    Polyline spineLine2 = PolylineBuilderEx.CreatePolyline(spineLineVertices2);

                    spine1BracketElement2.SetGeometry(spineLine2);

                    //Add to legend list
                    legendElementList.Add(spine1BracketElement2);

                    #endregion

                }

                #endregion

                #region ANNOTATION ELEMENT
                if (currentElementName == Constants.Graphics.annoBracket)
                {
                    
                    GraphicElement tGraphicElement = currentElementObject as GraphicElement;
                    if (tGraphicElement != null)
                    {
                        //Set rotation so the graphic is vertical instead of horizontal
                        tGraphicElement.SetRotation(90);

                        //Set text
                        TextElement tElement = currentElementObject as TextElement;
                        if (tElement != null)
                        {
                            //Reset text rotation, within the graphic, to 0 instead of 90.
                            //ArcMap used to have a different way of doing this
                            CIMGraphic cimGraphic = tElement.GetGraphic();
                            if (cimGraphic != null)
                            {
                                CIMTextSymbol cIMTextSymbol = cimGraphic.Symbol.Symbol as CIMTextSymbol;
                                cIMTextSymbol.Angle = 0;
                                tElement.SetGraphic(cimGraphic);
                            }

                            tElement.SetTextProperties(new TextProperties(Constants.TextConfiguration.tagAllCaps + currentLabel1 + Constants.TextConfiguration.endTagAllCaps, tElement.TextProperties.Font, tElement.TextProperties.FontSize, tElement.TextProperties.FontStyle));
                            if (currentLabel1 == null || currentLabel1 == string.Empty || currentLabel1 == " ")
                            {
                                tElement = Symbols.SetMissingTextSymbol(tElement);
                            }
                        }

                        

                        //Move to current anchor for a start
                        MoveElement(currentElementObject, anchorPoint.Item1, anchorPoint.Item2);

                        //Keep
                        annotationBracket = currentElementObject;
                    }
                }
                #endregion

                //Process left bracket annotation waiting to have the right anchor point
                if (waitingCenterLeftBracket != null && annotationBracket != null)
                {
                    //Set new anchor
                    Tuple<double, double> annoBracketAnchor = new Tuple<double, double>(waitingCenterLeftBracket.GetBounds().XMin, waitingCenterLeftBracket.GetBounds().YMax - waitingCenterLeftBracket.GetHeight() / 2.0);
                    PositionElement(annotationBracket, annoBracketAnchor.Item1, annoBracketAnchor.Item2, Anchor.BottomMidPoint); //Bottom mid because anchors rotates with the graphic
                    
                    //Move at the right anchor
                    MoveElement(annotationBracket, xSpacing, 0);

                    waitingCenterLeftBracket = null;
                    annotationBracket = null;
                }
            }
            catch (Exception AddLeftBracketException)
            {
                new ErrorService(AddLeftBracketException).WriteToFile();
            }
        }

        /// <summary>
        /// Will add a right bracket that wraps around a block of unit boxes and display another one a the far right
        /// </summary>
        /// <returns></returns>
        public async Task AddRightBracket()
        {
            try
            {
                //Process left bracket that was waiting to have Y spacing
                if (waitingRightBracket != null && !currentElementName.Contains(Constants.Graphics.keywordBracket))
                {
                    //Make sure to send bracket at the right place within the right column
                    double columnShift = 1;
                    if (currentColumn != 0)
                    {
                        columnShift = currentColumn;
                    }

                    MoveElement(waitingRightBracket, GetXSpacing(Constants.Graphics.bracketRightUpper) + (columnShift * GetXSpacing(Constants.Graphics.columnWidth)), -ySpacing);

                    waitingRightBracket = null;

                    howManyRightBrackets = howManyRightBrackets + 1;
                }

                #region UPPER ELEMENT
                if (currentElementName == Constants.Graphics.bracketRightUpper)
                {
                    //Get appropriate element
                    if (templateGraphicDico.ContainsKey(currentElementName) && lastElement != null)
                    {
                        Element rightBracketElement = currentElementObject;

                        //Set new anchor
                        Tuple<double, double> rightBracketAnchor = new Tuple<double, double>(anchorPoint.Item1, lastElement.GetBounds().YMax);
                        PositionElement(currentElementObject, rightBracketAnchor.Item1, rightBracketAnchor.Item2, Anchor.BottomLeftCorner);

                        //Keep in waiting line, to be move in Y axis
                        waitingRightBracket = rightBracketElement;
                        upRightBracket = rightBracketElement;

                    }
                }

                #endregion

                #region LOWER ELEMENT
                if (currentElementName == Constants.Graphics.bracketRightLower)
                {
                    //Get appropriate element
                    if (templateGraphicDico.ContainsKey(currentElementName) && upRightBracket != null)
                    {
                        #region END BRACKET NOTCH
                        Element rightEndBracketElement = currentElementObject;

                        //Set new anchor
                        Tuple<double, double> bottomBracketAnchor = new Tuple<double, double>(anchorPoint.Item1, lastElement.GetBounds().YMin + rightEndBracketElement.GetHeight());
                        PositionElement(rightEndBracketElement, bottomBracketAnchor.Item1, bottomBracketAnchor.Item2);

                        //Move
                        MoveElement(rightEndBracketElement, xSpacing + (currentColumn * GetXSpacing(Constants.Graphics.columnWidth)), 0);

                        #endregion

                        #region MIDDLE BRACKET NOTCH
                        Element rightMiddleBracketElement = CopyElementObject(templateGraphicDico[Constants.Graphics.bracketRightCenter], currentOrder.ToString());

                        //Set new anchor
                        double rightMiddleBracketY = upRightBracket.GetBounds().YMin - Math.Abs(((rightEndBracketElement.GetBounds().YMax - upRightBracket.GetBounds().YMin) / 2.0));
                        Tuple<double, double> rightMiddleBracketAnchor = new Tuple<double, double>(anchorPoint.Item1, rightMiddleBracketY);
                        PositionElement(rightMiddleBracketElement, rightMiddleBracketAnchor.Item1, rightMiddleBracketAnchor.Item2, Anchor.LeftMidPoint);

                        //Move
                        MoveElement(rightMiddleBracketElement, GetXSpacing(Constants.Graphics.bracketRightCenter) + (currentColumn * GetXSpacing(Constants.Graphics.columnWidth)), 0); //Move accordingly to x spacing if any

                        //Add to legend list
                        legendElementList.Add(rightMiddleBracketElement);

                        //Keep element to use with bracket annotation
                        waitingCenterRightBracket = rightMiddleBracketElement;

                        #endregion

                        #region BRACKET SPINE 1
                        Element rightSpine1BracketElement = CopyElementObject(templateGraphicDico[Constants.Graphics.bracketSpine], currentOrder.ToString(), "_UPPER");

                        //Set new anchor
                        Coordinate2D rightStartPointLine = new Coordinate2D(upRightBracket.GetBounds().XMax, upRightBracket.GetBounds().YMin);
                        Coordinate2D rightEndPointLine = new Coordinate2D(rightMiddleBracketElement.GetBounds().XMin, rightMiddleBracketElement.GetBounds().YMax);

                        List<Coordinate2D> spineLineVertices = new List<Coordinate2D>() { rightStartPointLine, rightEndPointLine };
                        Polyline spineLine = PolylineBuilderEx.CreatePolyline(spineLineVertices);

                        rightSpine1BracketElement.SetGeometry(spineLine);

                        //Add to legend list
                        legendElementList.Add(rightSpine1BracketElement);
                        #endregion

                        #region BRACKET SPINE 2
                        Element rightSpine2BracketElement = CopyElementObject(templateGraphicDico[Constants.Graphics.bracketSpine], currentOrder.ToString(), "_LOWER");

                        //Set new anchor
                        Coordinate2D rightStartPointLine2 = new Coordinate2D(rightMiddleBracketElement.GetBounds().XMin, rightMiddleBracketElement.GetBounds().YMin);
                        Coordinate2D rightEndPointLine2 = new Coordinate2D(rightEndBracketElement.GetBounds().XMax, rightEndBracketElement.GetBounds().YMax);

                        List<Coordinate2D> spineLineVertices2 = new List<Coordinate2D>() { rightStartPointLine2, rightEndPointLine2 };
                        Polyline spineLine2 = PolylineBuilderEx.CreatePolyline(spineLineVertices2);

                        rightSpine2BracketElement.SetGeometry(spineLine2);

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
                    double xUnitBoxBracketAnchor = waitingCenterRightBracket.GetBounds().XMax + GetXSpacing(Constants.Graphics.unitBoxBracket);
                    double yUnitBoxBracketAnchor = waitingCenterRightBracket.GetBounds().YMin + (waitingCenterRightBracket.GetBounds().YMax - waitingCenterRightBracket.GetBounds().YMin) / 2.0 + bracketMapUnit.Item1.GetHeight() / 2.0;
                    Tuple<double, double> mapUnitBracketAnchor = new Tuple<double, double>(xUnitBoxBracketAnchor, yUnitBoxBracketAnchor);
                    PositionElement(bracketMapUnit.Item1, mapUnitBracketAnchor.Item1, mapUnitBracketAnchor.Item2);
                    
                    //MoveItemToAnchorPoint(bracketMapUnit.Item1, mapUnitBracketAnchor);
                    //SetRectangularPolygonFromAnchorType(bracketMapUnit.Item1, mapUnitBracketAnchor);
                    //Set new anchor for unit box dem if any
                    if (bracketMapUnit.Item4 != null)
                    {
                        //MoveItemToAnchorPoint(bracketMapUnit.Item4, mapUnitBracketAnchor);
                        PositionElement(bracketMapUnit.Item4, mapUnitBracketAnchor.Item1, mapUnitBracketAnchor.Item2);
                    }

                    //Set new anchor for label
                    double xLabelBracketAnchor = mapUnitBracketAnchor.Item1 + bracketMapUnit.Item1.GetWidth() / 2.0;
                    double yLabelBracketAnchor = mapUnitBracketAnchor.Item2 - bracketMapUnit.Item1.GetHeight() / 2.0;
                    Tuple<double, double> labelBracketAnchor = new Tuple<double, double>(xLabelBracketAnchor, yLabelBracketAnchor);
                    PositionElement(bracketMapUnit.Item2, labelBracketAnchor.Item1, labelBracketAnchor.Item2);
                    //MoveItemToAnchorPoint(bracketMapUnit.Item2, labelBracketAnchor);

                    //Set new anchor for description
                    double xDescBracketAnchor = bracketMapUnit.Item1.GetBounds().XMax + GetXSpacing(Constants.Graphics.elementDescriptionGapWidth);
                    Tuple<double, double> descBracketAnchor = new Tuple<double, double>(xDescBracketAnchor, yUnitBoxBracketAnchor);
                    //AddDescriptionFromElement(bracketMapUnit.Item3, bracketMapUnit.Item1, currentDoc, mapUnitBracketAnchor, Constants.Graphics.unitBox);
                    //(bracketMapUnit.Item1.GetHeight() / 2.0) + bracketMapUnit.Item1.GetBounds().YMin
                    PositionElement(bracketMapUnit.Item3, descBracketAnchor.Item1, descBracketAnchor.Item2 - (bracketMapUnit.Item1.GetHeight() / 2.0), Anchor.LeftMidPoint);
                    //MoveElement(bracketMapUnit.Item3, descBracketAnchor.Item1, 0);

                    waitingCenterRightBracket = null;
                    bracketMapUnit = new Tuple<Element, Element, Element, Element>(null, null, null, null);
                }
                #endregion

            }
            catch (Exception AddRightBracketException)
            {
                new ErrorService(AddRightBracketException).WriteToFile();
            }
        }

        /// <summary>
        /// Will add a note or a top note element to the legend
        /// </summary>
        /// <returns></returns>
        public async Task AddNote()
        {
            try
            {
                if (currentElementName == Constants.Graphics.topNote || currentElementName == Constants.Graphics.note)
                {

                    //Get appropriate element
                    if (templateGraphicDico.ContainsKey(currentElementName))
                    {
                        //Set text
                        TextElement noteTextElement = currentElementObject as TextElement;
                        if (noteTextElement != null)
                        {
                            //Add description and header if needed
                            if (currentHeading != null && currentHeading != string.Empty && currentHeading != " ")
                            {
                                currentDescription = Constants.TextConfiguration.tagBold + currentHeading + Constants.TextConfiguration.endTagBold + " " + currentDescription;
                            }

                            if (currentDescription == null || currentDescription == string.Empty || currentDescription == " ")
                            {
                                noteTextElement = Symbols.SetMissingTextSymbol(noteTextElement, Properties.Resources.ErrorMissingLabel);
                            }
                            else
                            {
                                noteTextElement.SetTextProperties(new TextProperties(currentDescription, noteTextElement.TextProperties.Font,
                                    noteTextElement.TextProperties.FontSize, noteTextElement.TextProperties.FontStyle));
                            }
                        }

                        //Set width and height
                        double wantedTextHeight = GetTextHeight(currentDescription, columnWidth);

                        Coordinate2D lowerLeftNote = new Coordinate2D(anchorPoint.Item1, anchorPoint.Item2);
                        Coordinate2D upperRightNote = new Coordinate2D(anchorPoint.Item1 + columnWidth, anchorPoint.Item2 - wantedTextHeight);

                        Envelope llEnvelope = EnvelopeBuilderEx.CreateEnvelope(lowerLeftNote, upperRightNote);
                        ArcGIS.Core.Geometry.Polygon notePoly = PolygonBuilderEx.CreatePolygon(llEnvelope);
                        currentElementObject.SetGeometry(notePoly);

                        //Set new anchor
                        anchorPoint = new Tuple<double, double>(anchorPoint.Item1, anchorPoint.Item2 - ySpacing);
                        PositionElement(currentElementObject, anchorPoint.Item1, anchorPoint.Item2);

                        //Move
                        MoveElement(currentElementObject, xSpacing, 0);

                        //Keep name
                        double noteHeight = currentElementObject.GetHeight();
                        double noteYMin = currentElementObject.GetBounds().YMin;
                        double noteYMax = currentElementObject.GetBounds().YMax;

                    }

                }

            }
            catch (Exception AddNoteException)
            {
                new ErrorService(AddNoteException).WriteToFile();
            }
        }
        #endregion
    }
}
