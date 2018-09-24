using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GSC_Legend_Renderer.Services
{
    public class MXD
    {
        /// <summary>
        /// Will navigate to layout view, if user was in data view.
        /// </summary>
        /// <param name="actView"></param>
        public void ActivateLayoutView(IActiveView actView)
        {

            if (!(actView is IPageLayout))
            {
                
                //Get id for layer view command
                UID pUID = new UID();
                pUID.Value = Dictionaries.Constants.ESRI.UIDLayoutViewCommand;

                //Get commands and execute the layer view command.
                ICommandItem layoutItem = ArcMap.Application.Document.CommandBars.Find(pUID);
                layoutItem.Execute();
            }

        }
    }
}
