using LovePdf.Model.Enums;
using Newtonsoft.Json;
using Newtonsoft.Json.Converters;

namespace LovePdf.Model.TaskParams
{
    /// <summary>
    /// Image To Pdf Params
    /// </summary>
    public class ImageToPdfParams : BaseParams
    {
        /// <summary>
        /// Orientation
        /// </summary>
        [JsonConverter(typeof(StringEnumConverter))]
        [JsonProperty("orientation")]
        public Orientations Orientation { get; set; }

        /// <summary>
        /// Margin
        /// </summary>
        [JsonProperty("margin")]
        public int Margin { get; set; }

        /// <summary>
        /// Page Size
        /// </summary>
        [JsonConverter(typeof(StringEnumConverter))]
        [JsonProperty("pagesize")]
        public PageSizes PageSize { get; set; }

        /// <summary>
        /// Merge After
        /// </summary>
        [JsonProperty("merge_after")]
        public bool MergeAfter { get; set; }

        /// <summary>
        /// Constructor
        /// </summary>
        public ImageToPdfParams()
        {
            SetDefaultValues();
        }

        private void SetDefaultValues()
        {
            Orientation = Orientations.Portrait;
            Margin = 0;
            PageSize = PageSizes.Fit;
            MergeAfter = true;
        }
    }


}
