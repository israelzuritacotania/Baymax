using LovePdf.Model.Enums;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace LovePdf.Core
{
    /// <summary>
    /// File Model
    /// </summary>
    public class FileModel
    {

        /// <summary>
        /// Server File name
        /// </summary>
        public string ServerFileName { get; set; }
        /// <summary>
        /// File name
        /// </summary>
        public string FileName { get; set; }
        /// <summary>
        /// Rotation
        /// </summary>
        [JsonConverter(typeof(StringEnumConverter))]
        public Rotate Rotate { get; set; }
        /// <summary>
        /// Password
        /// </summary>
        public string Password { get; set; }
    }
}