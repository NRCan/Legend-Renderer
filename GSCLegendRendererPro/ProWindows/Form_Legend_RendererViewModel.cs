using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Core.Data.DDL;
using ArcGIS.Core.Internal.CIM;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Mapping;
using GSCLegendRendererPro.Utilities;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;
using System.Windows.Input;
using static GSCLegendRendererPro.Utilities.Layers;
using Field = ArcGIS.Core.Data.Field;

namespace GSCLegendRendererPro.ProWindows
{
    public class Form_Legend_RendererViewModel: PropertyChangedBase
    {
        #region INIT

        private Form_Legend_Renderer _view = null;
        private object _lock = new(); //For locking the threads to update obs. collection
        private Uri _legendTableWorkspaceUri = null;

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
        private void CreateLegend()
        {
            try
            {

            }
            catch (Exception CreateLegendException)
            {
                new ErrorService(CreateLegendException).WriteToFile();
            }
        }

        #endregion


    }
}
