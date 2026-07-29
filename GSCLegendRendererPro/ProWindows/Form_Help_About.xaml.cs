using GSCLegendRendererPro.Utilities;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace GSCLegendRendererPro.ProWindows
{
    /// <summary>
    /// Interaction logic for Form_Help_About.xaml
    /// </summary>
    public partial class Form_Help_About : ArcGIS.Desktop.Framework.Controls.ProWindow
    {
        public Form_Help_About()
        {
            InitializeComponent();
            this.DataContext = new Form_Help_AboutViewModel(this);
        }

        /// <summary>
        /// Open hyperlink in default browser or mail client
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {


            try
            {
                // hack because of this: https://github.com/dotnet/corefx/issues/10361
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    string url = e.Uri.AbsoluteUri.Replace("&", "^&");
                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                    e.Handled = true;
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    Process.Start("xdg-open", e.Uri.AbsoluteUri);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    Process.Start("open", e.Uri.AbsoluteUri);
                }

            }
            catch (Exception ex)
            {
                new ErrorService(ex).WriteToFile();
            }
        }
    }
}
