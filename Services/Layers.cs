using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.esriSystem;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GSC_Legend_Renderer.Services
{
    public class Layers
    {
        /// <summary>
        /// Will return a list of all layers inside table of content
        /// </summary>
        /// <param name="currentDoc">The document in which the list will be built</param>
        /// <param name="layerUID">The unique id of the type of layer to list</param>
        /// <returns></returns>
        public List<ILayer> GetListOfLayers(IMxDocument currentDoc, UID layerUID)
        {
            //Variables
            List<ILayer> layerList = new List<ILayer>();

            //Get all the layes
            IEnumLayer layers = currentDoc.ActiveView.FocusMap.get_Layers(layerUID, true);

            //Iterate through all layers inside the table of content
            ILayer currentLayer = layers.Next();
            while (currentLayer != null)
            {
                layerList.Add(currentLayer);
                currentLayer = layers.Next();
            }

            return layerList;
        }

        /// <summary>
        /// Will return a list of all table views inside table of content
        /// </summary>
        /// <param name="currentDoc">The document in which the list will be built</param>
        /// <returns></returns>
        public List<IStandaloneTable> GetListOfStandaloneTables(IMxDocument currentDoc)
        {
            //Variables
            List<IStandaloneTable> tableList = new List<IStandaloneTable>();

            //Get all the layes
            IStandaloneTableCollection tCollection = currentDoc.FocusMap as IStandaloneTableCollection;

            if (tCollection != null)
            {
                //Iterate through all layers inside the table of content
                for (int t = 0; t < tCollection.StandaloneTableCount; t++)
                {
                    IStandaloneTable currentTable = tCollection.StandaloneTable[t];
                    tableList.Add(currentTable);
                }

            }


            return tableList;
        }
    }
}
