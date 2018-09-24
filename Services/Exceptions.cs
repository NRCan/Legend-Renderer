using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GSC_Legend_Renderer.Services
{
    public static class Exceptions
    {
        ///USE STACK TRACE ONLY, INSTEAD OF THIS METHOD.
        /// <summary>
        /// A method to get a line number when something crashes. To be used with a try catch.
        /// </summary>
        /// <param name="e">The exception to retrieve line number from</param>
        /// <returns></returns>
        public static int LineNumber(this Exception e)
        {
            int lineNumber = 0;
            try
            {
                lineNumber = Convert.ToInt32(e.StackTrace.Substring(e.StackTrace.LastIndexOf(":line") + 5));
            }
            catch
            {

            }
            return lineNumber;
        }

    }
}
