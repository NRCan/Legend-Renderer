using ArcGIS.Core.Data;
using ArcGIS.Desktop.Mapping;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GSCLegendRendererPro.Utilities
{
    public class Workspace
    {
        /// <summary>
        /// Will return the original database path of a feature layer or standalone table
        /// </summary>
        /// <param name="inObject">A feature layer object or a standalone table one</param>
        /// <returns></returns>
        public static Uri GetWorkspacePath(object inObject)
        {
            Uri outputWorkspaceUri = null;
            Uri objectUri = null;

            //Case - feature layer
            FeatureLayer fl = inObject as FeatureLayer;
            if (fl != null)
            {
                FeatureClass fc = fl.GetFeatureClass();

                if (fc != null)
                {
                    objectUri = fc.GetPath();
                }
            }

            //Case - standalone table
            StandaloneTable st = inObject as StandaloneTable;
            if (st != null)
            {
                Table t = st.GetTable();

                if (t != null)
                {
                    objectUri = t.GetPath();
                }
            }

            //Get workspace path 
            if (objectUri != null)
            {
                string path = objectUri.OriginalString;
                string outputWorkspacePath = Directory.GetParent(path).FullName;
                if (outputWorkspacePath != null && outputWorkspacePath != string.Empty)
                {
                    //Remove any feature dataset name at the end of the workspace path
                    if (outputWorkspacePath.Contains(".gdb") && (outputWorkspacePath.Contains(".gdb\\") || outputWorkspacePath.Contains(".gdb/")))
                    {
                        outputWorkspacePath = outputWorkspacePath.Substring(0, outputWorkspacePath.IndexOf(".gdb") + 4);
                    }

                    outputWorkspaceUri = new Uri(outputWorkspacePath);
                }

            }

            return outputWorkspaceUri;
        }
    }
}
