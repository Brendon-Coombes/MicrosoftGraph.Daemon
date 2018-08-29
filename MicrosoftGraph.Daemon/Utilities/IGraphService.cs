using System.IO;
using System.Threading.Tasks;

namespace TERACC.WebJob.ResultsProcessor.Utilities
{
    public interface IGraphService
    {
        /// <summary>
        /// Uploads a file to Microsoft Graph asynchronously
        /// </summary>
        /// <param name="fileToUpload">The memory stream representation of the file that needs to be uploaded</param>
        /// <param name="fileNameWithExtension">The name the file should be called in the destination including the extension</param>
        /// <param name="driveName">The name of the Document Library / Drive within SharePoint to upload the file to</param>
        /// <returns></returns>
        Task<string> UploadFileAsync(MemoryStream fileToUpload, string fileNameWithExtension, string driveName);
    }
}
