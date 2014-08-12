/* Release under GPL V3 License 
 * The GPL (V2 or V3) is a copyleft license that requires anyone who modifies or updates this code 
 * to make the source available under the same terms.
 */

using Microsoft.Office365.OAuth;
using Microsoft.Office365.SharePoint;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace O365APIContactsSample
{
    public class MyFile
    {
        public string Id { get; set; }

        public string Name { get; set; }

        public Stream ContentStream { get; set; }
    }

    public static class  MyFilesApiSample
    {
        const string MyFilesCapability = "MyFiles";

        // Do not make static in Web apps; store it in session or in a cookie instead
        static string _lastLoggedInUser;
        static DiscoveryContext _discoveryContext;
        static SharePointClient _spFilesClient;

        public static async Task<ObservableCollection<MyFile>> GetMyFiles()
        {
            ObservableCollection<MyFile> myFilesList = new ObservableCollection<MyFile>();

            var filesResults = await _spFilesClient.Files.ExecuteAsync();

            var files = filesResults.CurrentPage.OfType<Microsoft.Office365.SharePoint.File>().ToList();

            foreach (var file in files)
            {
                MyFile myFile = new MyFile();
                myFile.Id = file.Id;
                myFile.Name = file.Name;
                myFile.ContentStream = await GetFileContent(file.Id);

                myFilesList.Add(myFile);
            }

            return myFilesList;
        }        

        public static async Task<ObservableCollection<MyFile>> GetMyFiles(string strFolderName)
        {
            ObservableCollection<MyFile> myFilesList = new ObservableCollection<MyFile>();

            var filesResults = await _spFilesClient.Files[strFolderName].ToFolder().Children.ExecuteAsync();

            var files = filesResults.CurrentPage.OfType<Microsoft.Office365.SharePoint.File>().ToList();

            foreach (var file in files)
            {
                MyFile myFile = new MyFile();
                myFile.Id = file.Id;
                myFile.Name = file.Name;
                myFile.ContentStream = await GetFileContent(file.Id);

                myFilesList.Add(myFile);
            }

            return myFilesList;
        }

        public static async Task<Stream> GetFileContent(string strFileId)
        {
            var file = (IFileFetcher)(await _spFilesClient.Files.GetById(strFileId).ToFile().ExecuteAsync());

            var streamFileContent = await file.DownloadAsync();

            streamFileContent.Seek(0, System.IO.SeekOrigin.Begin);

            return streamFileContent;
        }

        public static async Task<MyFile> Create(string strFileName, Stream streamFileContent)
        {
            MyFile myFile = null;

            IFile file = await _spFilesClient.Files.AddAsync(strFileName, true, streamFileContent);

            if (file != null)
            {
                myFile = new MyFile();
                myFile.Id = file.Id;
                myFile.Name = file.Name;
                myFile.ContentStream = streamFileContent;
            }

            return myFile;
        }        

        public static async Task Update(MyFile myFile, Stream streamFileContent)
        {
            var file = await _spFilesClient.Files.GetById(myFile.Id).ToFile().ExecuteAsync();

            file.Name = myFile.Name;

            await file.UpdateAsync();           
        }

        public static async Task UpdateFileContent(string strFileId, Stream stremFileContent)
        {
            var file = (IFileFetcher)(await _spFilesClient.Files.GetById(strFileId).ToFile().ExecuteAsync());

            await file.UploadAsync(stremFileContent);
        }

        public static async Task Delete(string strFileId)
        {
            var file = await _spFilesClient.Files.GetById(strFileId).ToFile().ExecuteAsync();

            await file.DeleteAsync();
        }        
    
        public static async Task SignIn()
        {
            if (_discoveryContext == null)
            {
                _discoveryContext = await DiscoveryContext.CreateAsync();
            }

            var dcr = await _discoveryContext.DiscoverCapabilityAsync(MyFilesCapability);
            
            var ServiceResourceId = dcr.ServiceResourceId;
            var ServiceEndpointUri = dcr.ServiceEndpointUri;

            _lastLoggedInUser = dcr.UserId;

            // Create the MyFiles client proxy:
            _spFilesClient = new SharePointClient(ServiceEndpointUri, async () =>
            {
                return (await _discoveryContext.AuthenticationContext.AcquireTokenSilentAsync(ServiceResourceId, _discoveryContext.AppIdentity.ClientId, new Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier(dcr.UserId, Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifierType.UniqueId))).AccessToken;
            });
        }

        public static async Task SignOut()
        {
            if (string.IsNullOrEmpty(_lastLoggedInUser))
            {
                return;
            }

            if (_discoveryContext == null)
            {
                _discoveryContext = await DiscoveryContext.CreateAsync();
            }

            await _discoveryContext.LogoutAsync(_lastLoggedInUser);
        }
    }
}
