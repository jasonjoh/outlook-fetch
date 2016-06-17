using Newtonsoft.Json;
using System.Collections.Generic;

namespace outlook_fetch.Outlook
{
    class ItemCollection<T>
    {
        [JsonProperty(PropertyName = "value")]
        public List<T> Items { get; set; }
        [JsonProperty(PropertyName = "@odata.nextLink")]
        public string NextPageUrl { get; set; }
    }
}
