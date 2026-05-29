using ArcGIS.Desktop.Framework.Contracts;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static GSCLegendRendererPro.Utilities.Layers;

namespace GSCLegendRendererPro.ProWindows
{
    public class Form_Legend_RendererViewModel: PropertyChangedBase
    {
        #region INIT

        private Form_Legend_Renderer _view = null;

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

                //FillGeolineType();
            }
        }


        private ObservableCollection<ComboBoxItem> _legendOrder = new();
        public ObservableCollection<ComboBoxItem> LegendOrder
        {
            get { return _legendOrder; }
        }
        private int _legendSelectedOrder = -1;
        public int LegendSelectedOrder
        {
            get { return _legendSelectedOrder; }
            set
            {
                SetProperty(ref _legendSelectedOrder, value, () => _legendSelectedOrder);
            }
        }

        #endregion

        #region RELAYS
        #endregion

        public Form_Legend_RendererViewModel(Form_Legend_Renderer view)
        {
            _view = view;
        }

        #region METHODS
        #endregion

    }
}
