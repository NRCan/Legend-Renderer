using ArcGIS.Desktop.Core;
using ArcGIS.Desktop.Framework;
using ArcGIS.Desktop.Framework.Contracts;
using GSCLegendRendererPro.Utilities;
using Octokit;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;

namespace GSCLegendRendererPro.ProWindows
{
    public class Form_Help_AboutViewModel: PropertyChangedBase
    {
        #region INIT
        private Form_Help_About _view = null;
        private WorkingEnvironment _workingEnvironment = new WorkingEnvironment();
        private enum webPageType { ReportIssue, ProjectPage }
        #endregion

        #region PROPERTIES
        private string _addinVersion = string.Empty;
        public string AddinVersion
        {
            get { return _addinVersion; }
            set
            {
                SetProperty(ref _addinVersion, value, () => _addinVersion);
            }
        }

        private string _addinVersionLatestOnline = string.Empty;
        public string AddinVersionLatestOnline
        {
            get { return _addinVersionLatestOnline; }
            set
            {
                SetProperty(ref _addinVersionLatestOnline, value, () => _addinVersionLatestOnline);
            }
        }

        /// <summary>
        /// Will be used to validate if online version of the addin is then latest or not, failing will color the font in red
        /// </summary>
        private bool _isValid = true;
        public bool IsValid
        {
            get { return _isValid; }
            set
            {
                SetProperty(ref _isValid, value, () => _isValid);
            }
        }
        #endregion

        #region RELAYS


        private ICommand _openOnlineIssue = null;
        public ICommand OpenOnlineIssue
        {
            get
            {
                if (_openOnlineIssue == null)
                {
                    _openOnlineIssue = new RelayCommand(() => OpenWebPage(webPageType.ReportIssue), () => true);
                }
                return _openOnlineIssue;
            }
        }

        private ICommand _openOnlineProject = null;
        public ICommand OpenOnlineProject
        {
            get
            {
                if (_openOnlineProject == null)
                {
                    _openOnlineProject = new RelayCommand(() => OpenWebPage(webPageType.ProjectPage), () => true);
                }
                return _openOnlineProject;
            }
        }

        #endregion

        #region METHODS
        public Form_Help_AboutViewModel(Form_Help_About aboutWindow)
        {
            try
            {
                _view = aboutWindow;

                // Get actual addin version
                _addinVersion = string.Format(Properties.Resources.FormHelpInstalledVersion, AddIn.GetConfigDamlAddInInfo(AddIn.GetAddInId()).Version);

                //Get latest online version
                _ = GetLatestRelease();
            }
            catch (Exception formHelpAboutViewModelException)
            {
                new ErrorService(formHelpAboutViewModelException).WriteToFile();
            }

        }

        public async Task GetLatestRelease()
        {
            try
            {
                _addinVersionLatestOnline = string.Format(Properties.Resources.FormHelpLatestVersion, string.Empty, string.Empty);

                //Access the web client, which is internally available for everyone (not need of authentication token or password)
                GitHubClient client = new GitHubClient(new ProductHeaderValue("LegendRenderer"));
                Release latestRelease = await client.Repository.Release.GetLatest("NRCan", "Legend-Renderer");
                if (latestRelease != null)
                {

                    string parsedOnlineVersion = latestRelease.TagName.Replace("V", "");
                    if (latestRelease.PublishedAt.HasValue)
                    {
                        DateTimeOffset releaseDate = latestRelease.PublishedAt.Value;
                        _addinVersionLatestOnline = string.Format(Properties.Resources.FormHelpLatestVersion, parsedOnlineVersion, releaseDate.ToString("yyyy-MM-dd"));
                        NotifyPropertyChanged(nameof(AddinVersionLatestOnline));
                    }

                    //Validate with current version
                    string parsedInstalledVersion = _addinVersion.Split(":")[1].Trim();
                    if (parsedOnlineVersion != parsedInstalledVersion || !parsedOnlineVersion.Contains(parsedInstalledVersion))
                    {
                        _isValid = false;
                        NotifyPropertyChanged(nameof(IsValid));
                    }
                    
                }

            }
            catch (Exception getLatestReleaseException)
            {
                new ErrorService(getLatestReleaseException).WriteToFile();
            }
        }

        /// <summary>
        /// Will open a web page in the default browser with the desire hyperlink passed as parameter.
        /// </summary>
        /// <param name="pageType"></param>
        private void OpenWebPage(webPageType pageType)
        {
            string url = "https://github.com/NRCan/GSC-Bedrock-Data-Model-and-Tools"; // Default URL

            switch (pageType)
            {

                case webPageType.ReportIssue:
                    url = "https://github.com/NRCan/Legend-Renderer/issues";
                    break;
                case webPageType.ProjectPage:
                    url = "https://github.com/NRCan/Legend-Renderer/releases";
                    break;
                default:
                    break;
            }

            try
            {
                // hack because of this: https://github.com/dotnet/corefx/issues/10361
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    url = url.Replace("&", "^&");
                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    Process.Start("xdg-open", url);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    Process.Start("open", url);
                }

            }
            catch (Exception ex)
            {
                new ErrorService(ex).WriteToFile();
            }
        }

        #endregion
    }
}
