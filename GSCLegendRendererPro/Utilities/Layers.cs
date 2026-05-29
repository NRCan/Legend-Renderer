using ArcGIS.Core.CIM;
using ArcGIS.Core.Data;
using ArcGIS.Desktop.Framework.Contracts;
using ArcGIS.Desktop.Framework.Threading.Tasks;
using ArcGIS.Desktop.Internal.Catalog;
using ArcGIS.Desktop.Internal.KnowledgeGraph;
using ArcGIS.Desktop.Mapping;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace GSCLegendRendererPro.Utilities
{
    /// <summary>
    /// A class that will help build some user interfaces with layers and their symbols as icons
    /// Viewmodel of the arc pro forms should inherit this class
    /// </summary>
    public class Layers: PropertyChangedBase
    {
        public class LayerDisplay
        {
            public string Name { get; set; }
            public FeatureLayer FLayer { get; set; }
            public BitmapSource Icon { get; set; }
            public bool IsChecked { get; set; }
            public StandaloneTable STable { get; set; }
        }

        /// <summary>
        /// Will create a combobox item with a layer file type along a little bitmap image of its symbols
        /// </summary>
        /// <param name="cimFeatureLayer"></param>
        /// <returns></returns>
        public static LayerDisplay MakeComboBoxItemWithSymbolIcons(CIMFeatureLayer cimFeatureLayer, FeatureLayer fl)
        {
            CIMSymbol sym = null;
            SymbolStyleItem si = null;
            BitmapSource bm = null;

            //Check for single renderer first
            CIMSimpleRenderer cimRenderer = cimFeatureLayer.Renderer as CIMSimpleRenderer;
            if (cimRenderer != null)
            {
                sym = cimRenderer.Symbol.Symbol;
            }
            else
            {
                //Get first symbol of first class instead
                CIMUniqueValueRenderer cimURenderer = cimFeatureLayer.Renderer as CIMUniqueValueRenderer;
                if (cimURenderer != null && cimURenderer.Groups != null && cimURenderer.Groups.Count() > 0)
                {
                    if (cimURenderer.Groups[0].Classes != null && cimURenderer.Groups[0].Classes.Count() > 0 &&
                    cimURenderer.Groups[0].Classes[0].Symbol != null)
                    {
                        sym = cimURenderer.Groups[0].Classes[0].Symbol.Symbol;
                    }
                    
                }
            }

            //Create a bitmap image for the icon in the combobox, if a symbol was detected
            if (sym != null)
            {
                si = new SymbolStyleItem()
                {
                    Symbol = sym,
                    PatchHeight = 15,
                    PatchWidth = 15
                };
                bm = si.PreviewImage as BitmapSource;
                bm.Freeze();
            }
            else
            {
                //TODO make this different if a line or a point is an input
                //Default value if nothing was found
                si = new SymbolStyleItem()
                {
                    Symbol = Symbols.GetDefaultPolygonSymbol() as CIMSymbol,
                    PatchHeight = 15,
                    PatchWidth = 15
                };
                bm = si.PreviewImage as BitmapSource;
                bm.Freeze();
            }

            //Create the combobox item
            LayerDisplay newLayerDisplay = new LayerDisplay
            {
                Name = fl.Name,
                FLayer = fl,
                Icon = bm,
                IsChecked = false
            };

            return newLayerDisplay;
        }

        /// <summary>
        /// Will fill the layer combobox with all feature layers in the map. By default, it will filter out layers
        /// that have GSC_SYMBOL or Label fields. You can add an extra field name to filter out more layers.
        /// </summary>
        /// <param name="geometryTypes">esriGeometryType to filter out layers</param>
        /// <param name="layerList">the observable collection list of layer to fill out</param>
        /// <param name="layerListName">the property name to notify of changes</param>
        /// <param name="selectedLayerIndex">the selected index for the list of layers</param>
        /// <param name="selectedLayerIndexName">the property name of the selected index</param>
        /// <param name="extraFieldFiltering">some extra field name to filter out possible layers</param>
        /// <returns></returns>
        public async Task<bool> UpdateLayerCombobox(List<esriGeometryType> geometryTypes, ObservableCollection<LayerDisplay> layerList, 
            string layerListName, int selectedLayerIndex, string selectedLayerIndexName, string extraFieldFiltering = "")
        {
            bool updated = false;

            try
            {
                await QueuedTask.Run(() =>
                {
                    if (MapView.Active != null && MapView.Active.Map != null)
                    {
                        List<FeatureLayer> layerEnum = MapView.Active.Map.GetLayersAsFlattenedList().OfType<FeatureLayer>().ToList();
                        if (layerEnum != null)
                        {
                            layerList.Clear();
                            NotifyPropertyChanged(nameof(layerListName));
                            foreach (FeatureLayer fl in layerEnum)
                            {
                                //System.Threading.Thread.Sleep(1);
                                if (geometryTypes.Contains(fl.ShapeType))
                                {
                                    //Get some definition to valide field and move with getting first symbol
                                    CIMFeatureLayer cIMFeatureLayer = fl.GetDefinition() as CIMFeatureLayer;
                                    List<FieldDescription> flDescriptions = fl.GetFieldDescriptions().ToList();

                                    if (cIMFeatureLayer != null && flDescriptions != null && flDescriptions.Count() > 0)
                                    {
                                        //Will need GSC_SYMBOL to work on or a label field
                                        bool styleFieldDescription = flDescriptions.Exists(x => x.Name == Constants.DatabaseFields.LegendSymbol);
                                        bool labelFieldDescription = flDescriptions.Exists(x => x.Alias == Constants.DatabaseFields.FLabelIDAlias);
                                        bool symbolFieldDescription = flDescriptions.Exists(x => x.Name == Constants.DatabaseFields.LegendSymbol_190101);

                                        //If some extra filtering is needed
                                        bool extraFiltering = false;
                                        if (extraFieldFiltering != string.Empty)
                                        {
                                            extraFiltering = flDescriptions.Exists(x => x.Name == extraFieldFiltering);
                                        }

                                        //Add if any options are true
                                        if (symbolFieldDescription || labelFieldDescription || extraFiltering || styleFieldDescription)
                                        {

                                            LayerDisplay layerItem = MakeComboBoxItemWithSymbolIcons(cIMFeatureLayer, fl);
                                            if (!layerList.Contains(layerItem))
                                            {
                                                layerList.Add(layerItem);
                                            }
                                            

                                        }

                                        NotifyPropertyChanged(nameof(layerListName));

                                    }
                                }
                            }

                            updated = true;
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                new ErrorService(ex).WriteToFile();
            }

            return updated;

        }

        /// <summary>
        /// Will fill the layer combobox with all feature layers in the map. By default, it will filter out layers
        /// that have GSC_SYMBOL or Label fields. You can add an extra field name to filter out more layers.
        /// </summary>
        /// <param name="geometryTypes">esriGeometryType to filter out layers</param>
        /// <param name="layerList">the observable collection list of layer to fill out</param>
        /// <param name="layerListName">the property name to notify of changes</param>
        /// <param name="selectedLayerIndex">the selected index for the list of layers</param>
        /// <param name="selectedLayerIndexName">the property name of the selected index</param>
        /// <param name="extraFieldFiltering">some extra field name to filter out possible layers</param>
        /// <returns></returns>
        public async Task<bool> UpdateTableViewCombobox(ObservableCollection<LayerDisplay> layerList,
            string layerListName, int selectedLayerIndex, string selectedLayerIndexName, string extraFieldFiltering = "")
        {
            bool updated = false;

            try
            {
                await QueuedTask.Run(() =>
                {
                    if (MapView.Active != null && MapView.Active.Map != null)
                    {
                        List<StandaloneTable> tableEnum = MapView.Active.Map.GetStandaloneTablesAsFlattenedList().OfType<StandaloneTable>().ToList();
                        if (tableEnum != null)
                        {

                            foreach (StandaloneTable st in tableEnum)
                            {
                                //Get some definition to validate fields
                                CIMStandaloneTable cIMTable= st.GetDefinition() as CIMStandaloneTable;
                                List<FieldDescription> flDescriptions = st.GetFieldDescriptions().ToList();

                                if (cIMTable != null && flDescriptions != null && flDescriptions.Count() > 0)
                                {
                                    LayerDisplay layerDisplay = new LayerDisplay()
                                    {
                                        Name = st.Name,
                                        FLayer = null,
                                        Icon = null,
                                        IsChecked = false,
                                        STable = st
                                    };

                                    //If some extra filtering is needed
                                    if (extraFieldFiltering != string.Empty)
                                    {
                                        if (flDescriptions.Exists(x => x.Name == extraFieldFiltering))
                                        {
                                            layerList.Add(layerDisplay);
                                        }
                                    }
                                    else
                                    {
                                        layerList.Add(layerDisplay);
                                    }


                                    NotifyPropertyChanged(nameof(layerListName));

                                }
                                
                            }

                            updated = true;
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                new ErrorService(ex).WriteToFile();
            }

            return updated;

        }

        /// <summary>
        /// Will return a list of feature layers if it exist within the map content from a given geodatabase
        /// </summary>
        /// <returns></returns>
        public async Task<List<FeatureLayer>> GetActiveFeatureLayerFromGeodatabase(Geodatabase geodatabase, string featureName)
        {
            List<FeatureLayer> outputFL = new List<FeatureLayer>();

            try
            {
                await QueuedTask.Run(() =>
                {
                    if (MapView.Active != null && MapView.Active.Map != null)
                    {
                        List<FeatureLayer> layerList = MapView.Active.Map.GetLayersAsFlattenedList().OfType<FeatureLayer>().ToList();
                        if (layerList != null)
                        {
                            foreach (FeatureLayer fl in layerList)
                            {
                                FeatureClass fc = fl.GetFeatureClass();
                                if (fc != null && fc.GetName() == featureName)
                                {
                                    Uri fcPath = fc.GetPath();
                                    if (fcPath != null && fcPath.OriginalString.Contains(geodatabase.GetPath().OriginalString))
                                    {
                                        outputFL.Add(fl);
                                    }

                                }
                            }
                            
                        }
                    }
                });
            }
            catch (Exception ex)
            {
                new ErrorService(ex).WriteToFile();
            }

            return outputFL;
        }
    }
}
