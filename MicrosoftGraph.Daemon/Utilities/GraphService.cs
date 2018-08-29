using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using TERACC.WebJob.ResultsProcessor.Utilities;

namespace MicrosoftGraph.Daemon.Utilities
{
    /// <summary>
    /// Repsonsible for interactions with Microsoft Graph
    /// </summary>
    public class GraphService : IGraphService
    {
        private readonly string _clientId;
        private readonly string _clientSecret;
        private readonly string _tenantId;
        private readonly string _redirectUri;
        private readonly string _authorityFormat;
        private readonly string _graphScope;
        private readonly string _graphHostName;
        private readonly string _graphSiteRelativePath;
        private readonly string _graphBaseUrl;

        /// <summary>
        /// Creates a new instance of <see cref="GraphService"/>
        /// </summary>
        /// <param name="clientId">The Azure Active Directory client ID that will be used by this service</param>
        /// <param name="clientSecret">The Azure Active Directory client secret that will be used by this service</param>
        /// <param name="tenantId">The Office 365 tenant ID that will be used by this service</param>
        /// <param name="redirectUri">The redirect URI configured in the Azure Active Directory application</param>
        /// <param name="authorityFormat">The authority uri format that this service will conform to</param>
        /// <param name="graphScope">The scope that this service will request from Microsoft Graph</param>
        /// <param name="graphBaseUrl">The base url to query the graph from</param>
        /// <param name="graphHostName">The host name of the graph (e.g test123.sharepoint.com)</param>
        /// <param name="graphSiteRelativePath">The relative path of the site within SharePoint to begin querying from</param>
        public GraphService(string clientId, string clientSecret, string tenantId, string redirectUri, string authorityFormat, string graphScope, string graphBaseUrl, string graphHostName, string graphSiteRelativePath)
        {
            _clientId = clientId;
            _clientSecret = clientSecret;
            _tenantId = tenantId;
            _redirectUri = redirectUri;
            _authorityFormat = authorityFormat;
            _graphScope = graphScope;
            _graphHostName = graphHostName;
            _graphSiteRelativePath = graphSiteRelativePath;
            _graphBaseUrl = graphBaseUrl;
        }

        /// <summary>
        /// Uploads a file to Microsoft Graph asynchronously
        /// </summary>
        /// <param name="fileToUpload">The memory stream representation of the file that needs to be uploaded</param>
        /// <param name="fileNameWithExtension">The name the file should be called in the destination including the extension</param>
        /// <param name="driveName">The name of the Document Library / Drive within SharePoint to upload the file to</param>
        /// <returns></returns>
        public async Task<string> UploadFileAsync(MemoryStream fileToUpload, string fileNameWithExtension, string driveName)
        {
            try
            {
                var graphClient = await CreateGraphServiceClientAsync();

                var site = await graphClient.Sites.GetByPath(_graphSiteRelativePath, _graphHostName).Request().GetAsync();
            
                //NOTE: There are pages in the drives, this will only find drives on the first page
                var drives = await graphClient.Sites[site.Id].Drives.Request().GetAsync();
                var drive = drives.First(x => x.Name == driveName);           

                var file = await graphClient.Drives[drive.Id].Root.ItemWithPath(fileNameWithExtension).Content.Request().PutAsync<DriveItem>(fileToUpload);

                return file.Id;
            }
            catch (Exception e)
            {
                Trace.TraceError("An unexpected error occurred uploading the file to the Drive");
                throw;
            }            
        }

        /// <summary>
        /// Retrieves a token for the Microsoft Graph using the client secret and scope
        /// </summary>
        /// <returns></returns>
        private async Task<string> AuthenticateToGraphAsync()
        {
            MSALCache appTokenCache = new MSALCache(_clientId);

            ConfidentialClientApplication daemonClient = new ConfidentialClientApplication(_clientId, string.Format(_authorityFormat, _tenantId), _redirectUri, new ClientCredential(_clientSecret), null, appTokenCache.GetMsalCacheInstance());
            AuthenticationResult authResult = await daemonClient.AcquireTokenForClientAsync(new[] { _graphScope });

            if (string.IsNullOrEmpty(authResult?.AccessToken))
            {
                throw new Exception("A token could not be retrieve from Microsft Graph");
            }

            return authResult?.AccessToken;           
        }

        /// <summary>
        /// Creates a graph service client that can be used to interact with the graph
        /// </summary>
        /// <returns></returns>
        private async Task<GraphServiceClient> CreateGraphServiceClientAsync()
        {
            string authenticationToken = await AuthenticateToGraphAsync(); 

            GraphServiceClient graphClient = new GraphServiceClient(
                _graphBaseUrl,
                new DelegateAuthenticationProvider(
                    async (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("Bearer", authenticationToken);
                    }));

            return graphClient;
        }    
    }
}
