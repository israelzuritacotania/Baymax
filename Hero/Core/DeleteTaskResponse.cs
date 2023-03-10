using Newtonsoft.Json;

namespace LovePdf.Core
{
    /// <summary>
    /// Delete Task Response
    /// </summary>
    public class DeleteTaskResponse
    {
        /// <summary>
        /// Upload Status
        /// </summary>
        [JsonProperty("upload_status")]
        public string UploadStatus { get; set; }
    }
}
