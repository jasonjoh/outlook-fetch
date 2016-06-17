using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace outlook_fetch.Outlook
{
    class ApiClient
    {
        // Used to set the base API endpoint, e.g. "https://outlook.office.com/api/beta"
        public string ApiEndpoint { get; set; }
        public string AccessToken { get; set; }

        public ApiClient()
        {
            // Set default endpoint
            ApiEndpoint = "https://outlook.office.com/api/v2.0";
            AccessToken = string.Empty;
        }

        public async Task<HttpResponseMessage> MakeApiCall(string method, string apiUrl, string userEmail, string payload, Dictionary<string, string> preferHeaders)
        {
            if (string.IsNullOrEmpty(AccessToken))
            {
                throw new ArgumentNullException("AccessToken", "You must supply an access token before making API calls.");
            }

            using (var httpClient = new HttpClient())
            {
                var request = new HttpRequestMessage(new HttpMethod(method), ApiEndpoint + apiUrl);

                // Headers
                // Add the access token in the Authorization header
                request.Headers.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", AccessToken);
                // Add a user agent (best practice)
                request.Headers.UserAgent.Add(new System.Net.Http.Headers.ProductInfoHeaderValue("outlook-fetch", "1.0"));
                // Indicate that we want JSON response
                request.Headers.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));
                // Set a unique ID on each request (best practice)
                request.Headers.Add("client-request-id", Guid.NewGuid().ToString());
                // Request that the unique ID also be included in the response (to make it easier to correlate request/response)
                request.Headers.Add("return-client-request-id", "true");
                // Set this header to optimize routing of request to appropriate server
                request.Headers.Add("X-AnchorMailbox", userEmail);

                if (preferHeaders != null)
                {
                    foreach (KeyValuePair<string, string> header in preferHeaders)
                    {
                        if (string.IsNullOrEmpty(header.Value))
                        {
                            // Some prefer headers only have a name, no value
                            request.Headers.Add("Prefer", header.Key);
                        }
                        else
                        {
                            request.Headers.Add("Prefer", string.Format("{0}=\"{1}\"", header.Key, header.Value));
                        }
                    }
                }

                // POST and PATCH should have a body
                if ((method.ToUpper() == "POST" || method.ToUpper() == "PATCH") &&
                    !string.IsNullOrEmpty(payload))
                {
                    request.Content = new StringContent(payload);
                    request.Content.Headers.ContentType.MediaType = "application/json";
                }

                var apiResult = await httpClient.SendAsync(request);
                return apiResult;
            }
        }

        public async Task<User> GetUser(string userEmail)
        {
            string requestUrl = string.Format("/Users/{0}", userEmail);

            HttpResponseMessage result = await MakeApiCall("GET", requestUrl, userEmail, null, null);

            if (result.IsSuccessStatusCode)
            {
                // Read the JSON response
                string response = await result.Content.ReadAsStringAsync();

                // Deserialize to a User object
                return JsonConvert.DeserializeObject<User>(response);
            }

            throw await GetExceptionFromError(result);
        }

        public async Task<ItemCollection<Contact>> GetContacts(string userEmail)
        {
            string requestUrl = string.Format("/Users/{0}/contacts", userEmail);
            // Sort the results by the date time created to get newest first
            requestUrl += "?$orderby=CreatedDateTime DESC";
            // Limit the results to the first 10
            requestUrl += "&$top=10";

            HttpResponseMessage result = await MakeApiCall("GET", requestUrl, userEmail, null, null);

            if (result.IsSuccessStatusCode)
            {
                // Read the JSON response
                string response = await result.Content.ReadAsStringAsync();

                // Deserialize to a collection of Contact objects
                return JsonConvert.DeserializeObject<ItemCollection<Contact>>(response);
            }

            throw await GetExceptionFromError(result);
        }

        private async Task<Exception> GetExceptionFromError(HttpResponseMessage errorResult)
        {
            StringBuilder errorBuilder = new StringBuilder();

            errorBuilder.AppendFormat("API Request failed with {0} ({1}).", (int)errorResult.StatusCode, errorResult.ReasonPhrase);

            string errorDetail = await errorResult.Content.ReadAsStringAsync();

            if (!string.IsNullOrEmpty(errorDetail))
            {
                ErrorResponse error = JsonConvert.DeserializeObject<ErrorResponse>(errorDetail);
                errorBuilder.AppendFormat("\nCode: {0}", error.Error.Code);
                errorBuilder.AppendFormat("\nMessage: {0}", error.Error.Message);
            }

            return new HttpRequestException(errorBuilder.ToString());
        }
    }
}
