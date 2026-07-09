using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Geometry;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Layouts;
using ArcGIS.Desktop.Mapping;
using GSCLegendRendererPro.Models;
using GSCLegendRendererPro.Services;
using GSCLegendRendererPro.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Media;
using static GSCLegendRendererPro.Utilities.Layers;
using Field = ArcGIS.Core.Data.Field;
using LinearUnit = ArcGIS.Core.Geometry.LinearUnit;

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
        //public Tuple<double, double> anchorPoint = GetAnchorPointStart(); //TODO Find if mxd is a CGM one or not.
        //public originalYSpacing = anchorPoint.Item2; //Synchronise with initial calculate anchor.
        public Tuple<double, double> anchorPointParent = new Tuple<double, double>(0, 0);
        public List<string> heading5Text = new List<string>(); //Init
        public double currentIteration = 0.0; //Will be used if user has forgot to enter an order.
        public bool nullOrderBreaker = false; //Will be used to show error message to user if null values are found, but only once.

        #endregion


        #region PROPERTIES

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
            CustomCombobox fieldElement = _legendElement.Where(x => x.Name.Contains(Constants.LegendTable.legendElementField, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();
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
            try
            {
                if (_legendSelectedLayerIndex != -1)
                {
                    //Setup prcedures
                    bool setupAddinCleared = await SetupAddinEnvironment();
                    bool setupLayoutCleared = await SetupLayoutAndGraphics();

                    if (setupAddinCleared && setupLayoutCleared)
                    {


                        //Show notication success
                        FrameworkApplication.AddNotification(new Notification()
                        {
                            Title = Properties.Resources.FormRendererTitle,
                            Message = Properties.Resources.GenericMessageCompleted,
                            ImageSource = System.Windows.Application.Current.Resources["Success_Toast48"] as ImageSource
                        });
                    }

                    //Close window
                    _view.Close();
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
                        if (styleItems == null || styleItems.Count() == 0)
                        {
                            //Load up the style coming from the default setup
                            string defaultStylePath = System.IO.Path.Combine(Properties.Settings.Default.WorkingEnvironmentPath, otherComponents.GEOLOGY_STYLE_NAME);
                            Project.Current.AddStyle(defaultStylePath);
                        }
                        gscStyle = Project.Current.GetItems<StyleProjectItem>().Where(x => x.Name == otherComponents.GEOLOGY_STYLE_NAME).FirstOrDefault();
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
        #endregion
    }
}
