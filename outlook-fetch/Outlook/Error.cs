using Newtonsoft.Json;

namespace outlook_fetch.Outlook
{
    class Error
    {
        [JsonProperty(PropertyName = "code")]
        public string Code { get; set; }
        [JsonProperty(PropertyName = "message")]
        public string Message { get; set; }
    }

    class ErrorResponse
    {
        [JsonProperty(PropertyName = "error")]
        public Error Error { get; set; }
    }
}
