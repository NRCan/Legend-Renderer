using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ESRI.ArcGIS.esriSystem;

namespace GSC_Legend_Renderer.Services
{
    public class ObjectManagement
    {
        /// <summary>
        /// Will make a copy of the passed object. The object needs to implement IPersist or IPersistStream to work properly.
        /// </summary>
        /// <param name="inputOb">The object to get a copy rom</param>
        /// <returns></returns>
        public static object CopyInputObject(object inputOb)
        {
            //Get IObjectCopy interface
            IObjectCopy objectCopy = new ObjectCopy();

            //Get IUnknown interface (copied map)
            object copiedObj = objectCopy.Copy(inputOb);

            //return objectCopy.Copy(inputOb);
            return copiedObj;
        }

        /// <summary>
        /// Will release a given object from com
        /// </summary>
        /// <param name="objectToRelease">The object to release</param>
        public static void ReleaseObject(object objectToRelease)
        {
            try
            {
                if (System.Runtime.InteropServices.Marshal.IsComObject(objectToRelease))
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(objectToRelease);
                }
                
                objectToRelease = null;
            }
            catch (Exception ex)
            {
                objectToRelease = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.StackTrace);
            }

        }
    }
}
