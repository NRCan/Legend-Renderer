using ArcGIS.Desktop.Core.Geoprocessing;
using ArcGIS.Desktop.Framework;
using GSCLegendRendererPro.Utilities;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace GSCLegendRendererPro.Utilities
{
    internal class ErrorService
    {
        private WorkingEnvironment WorkingEnvironment
        {
            get
            {
                return new WorkingEnvironment();
            }
        }

        public string Message { get; set; }
        public Exception Exception { get; set; }
        public IGPResult GeoprocessingResult { get; set; }
        public string DefaultPath
        {
            get
            {

                var folder = WorkingEnvironment.WorkingEnvironmentPath;
                var fileName = Constants.Debug.debugFileName;

                return Path.Combine(folder, fileName);
            }
        }

        public ErrorService(string message) { Message = message; }
        public ErrorService(Exception ex) { Exception = ex; }
        public ErrorService(IGPResult gpResult) { GeoprocessingResult = gpResult; }

        public bool WriteToFile(string path = "", bool showToast = true)
        {

            if (string.IsNullOrEmpty(path))
            {
                path = DefaultPath;
            }

            //Show toast
            if (showToast)
            {
                FrameworkApplication.AddNotification(new Notification()
                {
                    Title = Properties.Resources.GenericMessageErrorTitle,
                    Message = Properties.Resources.GenericMessageError,
                    ImageSource = System.Windows.Application.Current.Resources["Error_Toast48"] as ImageSource,
                    Severity = Notification.SeverityLevel.High
                });
            }


            try
            {
                //Make sure the directory exists
                if (!Directory.Exists(Path.GetDirectoryName(path)))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(path));
                }

                //Write the error to the log file
                using (var writer = new StreamWriter(path, true))
                {
                    writer.WriteLine("-----------------------------------------------------------------------------");
                    writer.WriteLine("Date : " + DateTime.Now.ToString(CultureInfo.InvariantCulture));
                    writer.WriteLine();

                    if (Exception != null)
                    {
                        writer.WriteLine(Exception.GetType().FullName);
                        writer.WriteLine("Source : " + Exception.Source);
                        writer.WriteLine("Message : " + Exception.Message);
                        writer.WriteLine("StackTrace : " + Exception.StackTrace);
                        writer.WriteLine("InnerException : " + Exception.InnerException?.Message);
                    }

                    if (!string.IsNullOrEmpty(Message))
                    {
                        writer.WriteLine(Message);
                    }

                    if (GeoprocessingResult != null)
                    {
                        writer.WriteLine(GeoprocessingResult.GetType().FullName);
                        if (GeoprocessingResult.Parameters != null && GeoprocessingResult.Parameters.Count() > 0)
                        {
                            foreach (Tuple<string,string, string, bool> item in GeoprocessingResult.Parameters.ToList())
                            {
                                writer.WriteLine("Parameters : " + item.Item1 + ", " + item.Item2 + ", " + item.Item3);
                            }
                            
                        }
                        if (GeoprocessingResult.Messages != null && GeoprocessingResult.Messages.Count() > 0)
                        {
                            foreach (IGPMessage message in GeoprocessingResult.Messages)
                            {
                                if (message != null)
                                {
                                    writer.WriteLine("Message : " + message.Text);
                                }
                                
                            }
                            
                        }
                        
                        writer.WriteLine("ErrorCode : " + GeoprocessingResult.ErrorCode.ToString());

                    }

                    writer.Close();
                }
            }
            catch (Exception ex)
            {
                new ErrorService(ex).WriteToFile();
                return false;
            }

            return true;
        }
    }
}
