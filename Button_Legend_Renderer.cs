using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.SystemUI;

namespace GSC_Legend_Renderer
{
    public class Button_Legend_Renderer : ESRI.ArcGIS.Desktop.AddIns.Button
    {
        public Button_Legend_Renderer()
        {
        }

        protected override void OnClick()
        {

            //Validation - Get document and set layout view if not already set
            IMxDocument currentDoc = (IMxDocument)ArcMap.Application.Document;
            ActivateLayoutView(currentDoc.ActiveView);

            ArcMap.Application.CurrentTool = null;
        }
        protected override void OnUpdate()
        {
            Enabled = ArcMap.Application != null;
        }

        /// <summary>
        /// Will navigate to layout view, if user was in data view.
        /// </summary>
        /// <param name="actView"></param>
        public void ActivateLayoutView(IActiveView actView)
        {

            //Activate
            Services.MXD mxdService = new Services.MXD();
            mxdService.ActivateLayoutView(actView);

            //Issue #78 -  Code fails when user is within map, within layout. 
            //Enforce activate view to be set in layer.
            if (actView.IsMapActivated)
            {
                try
                {
                    actView.IsMapActivated = false;
                    actView.Refresh();
                }
                catch (Exception)
                {
                }
                
            }
            
            //Open renderer form
            Form_Legend_Renderer inputForm = new Form_Legend_Renderer();
            inputForm.Show();

        }
    }

}
