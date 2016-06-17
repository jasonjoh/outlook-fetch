# Getting Started with the Outlook Mail API and the Client Credentials Grant Flow

## Registering the app

In order to use the client credentials flow, the Outlook APIs require that you use a certificate for app authentication, rather than a client secret. Before we create the app registration in Azure, we'll start by creating a certificate.

> As of this writing, the Azure v2.0 endpoints do not support the client credentials flow. So in order to use this flow, we need to register the app in the Azure Management portal and use the v1 token endpoint.

### Create the certificate

The certificate must have a key length of at least 2048 bits, but it can be a self-issued certificate. On Windows, we can use the **makecert.exe** tool that's included with Visual Studio to create a certificate.

> If makecert.exe is not on your PATH in a normal command prompt, you can open the **Developer Command Prompt for VS2015** and run the commands from there.

1. From a command prompt, enter the following command, which creates the certificate in the current user's personal certificate store.

  ```Shell
  makecert -r –pe -n "CN=Outlook Fetch Sample Cert" –ss my –len 2048
  ```

1. Find the certificate in the user's personal certificate store in the Certificates MMC snap-in.

  1. Press **Windows key + R**, then type `mmc` and hit enter.
  1. On the **File** menu, choose **Add/Remove Snap-in**.
  1. Select the **Certificates** snap-in and click **Add**.
  1. Choose **My user account** and click **Finish**.
  1. Click **OK**.
  1. Expand the **Certificates - Current User** item in the left-hand list.
  1. Expand **Personal**, then select **Certificates**.
  1. Locate the **Outlook Fetch Sample Cert** entry in the list and select it.

1. Export the certificate with the private key. The app will need the private key to sign token requests.

  1. On the **Action** menu, choose **All Tasks**, then **Export**.
  1. On the Welcome screen, click **Next**.
  1. Choose **Yes, export the private key** and click **Next**.
  1. Make sure **Personal Information Exchange - PKCS #12 (.PFX)** is selected and click **Next**.
  1. Choose **Password** and enter a strong password. Click **Next**.
  1. Choose a location and file name and click **Next**. In this example we'll call this file `outlook-fetch-priv.pfx`.
  1. Click **Finish**.

1. Export the certificate without the private key. We'll need this to attach to our app registration so Azure can use it to verify the app's signed token requests.

  1. Make sure you still have the **Outlook Fetch Sample Cert** selected.
  1. On the **Action** menu, choose **All Tasks**, then **Export**.
  1. On the Welcome screen, click **Next**.
  1. Choose **No, do not export the private key** and click **Next**.
  1. Choose **Base-64 encoded X.509 (.CER)** and click **Next**.
  1. Choose a location and file name and click **Next**. In this example we'll call this file `outlook-fetch-pub.cer`.
  1. Click **Finish**.

### Create the app registration

In this step we'll create an app registration in the Azure Active Directory associated with your Office 365 tenant. You'll need the credentials for an organizational administrator, and you'll need an Azure subscription associated with your Office 365 tenant. Instructions for associating an Azure subscription with Office 365 can be [found on MSDN](https://msdn.microsoft.com/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription).

1. Browse to https://manage.windowsazure.com and sign in as an organizational administrator.

1. In the **all items** list, choose your **Directory** entry.

1. In the navigation menu at the top, choose **APPLICATIONS**.

  ![](./readme-images/azure-portal-applications.PNG)

1. On the bottom toolbar, choose **ADD**.

  ![](./readme-images/azure-portal-add.PNG)

1. Choose **Add an application my organization is developing**.

1. Enter `Outlook Fetch Sample` for the name.

1. Make sure that **WEB APPLICATION AND/OR WEB API** is selected. Even though the app we are creating will be a console app, we still need to select this type. Click the **Next** button.

1. Enter `http://localhost` for the **SIGN-ON URL**.

1. Enter `https://<your Office 365 domain>/outlook-fetch` for the **APP ID URI**, replacing `<your Office 365 domain>` with your Office 365 domain. For example: `https://contoso.onmicrosoft.com/outlook-fetch`. Click the **Complete** button.

1. In the navigation menu at the top, choose **CONFIGURE**.

  ![](./readme-images/azure-portal-configure.PNG)

1. Locate the **CLIENT ID**. Copy this value and save it somewhere. We'll need this once we start developing the app.

1. Locate the **permissions to other applications** section. Click **Add application**.

1. Choose **Office 365 Exchange Online** and click the **Complete** button.

1. There should now be an entry for **Office 365 Exchange Online** in the **permissions to other applications** section. In the **Application Permissions** dropdown, select **Read mail in all mailboxes**, **Read contacts in all mailboxes**, and **Read calendars in all mailboxes**.

  ![](./readme-images/azure-portal-permissions.PNG)

1. In the bottom toolbar, click **SAVE**. Wait for the update to complete.

1. In the bottom toolbar, click **MANAGE MANIFEST**, then choose **Download Manifest**. Save the manifest in the same directory as your certificate files, and rename it to `appmanifest.json`.

### Upload public key to the app registration

In this step we'll insert the public key we exported earlier into the app manifest we downloaded, then upload the modified manifest back to Azure.

1. Download the Get-KeyCredentials.ps1 file and save it in the same directory as your certificate files.

1. Open **Windows Powershell** in the directory where you saved your certificate files.

1. Run the following command: 

  ```Shell
  .\Get-KeyCredentials.ps1 .\outlook-fetch-pub.cer
  ```

1. Open the `keyCredentials.txt` file and verify that values were generated. It should look similar to this (values truncated):

  ```JSON
  "keyCredentials": [
    {
      "customKeyIdentifier": "Vr3TRhO85mjjJevziQrBk2nQnpE=",
      "keyId": "79a41b2e-85d1-49a2-8a81-d378b0eb8f9a",
      "type": "AsymmetricX509Cert",
      "usage": "Verify",
      "value": "MIIDKzCCAhOgAwIBAgIQ...1iukV2yc52g="
    }
  ],
  ```

1. Open the `appmanifest.json` file in a text editor. Locate the `keyCredentials` value:

  ```JSON
  "keyCredentials": [],
  ```

1. Replace the `keyCredentials` line in `appmanifest.json` with the entire contents of the `keyCredentials.txt` file. Save the file.

1. In the Configure page for the app registration in Azure Management Portal, click the **MANAGE MANIFEST** button and choose **Upload Manifest**. Browse to `appmanifest.json` and click **OK** to upload the updated manifest.

1. Before leaving the Azure Management Portal, click the **VIEW ENDPOINTS** button in the bottom toolbar. Copy the value of **OAUTH 2.0 AUTHORIZATION ENDPOINT** and save it somewhere. We'll need this value later.

## Create the app

Let's start by creating a console app and adding the following libraries:

- [Active Directory Authentication Library](https://www.nuget.org/packages/Microsoft.IdentityModel.Clients.ActiveDirectory): We'll use this library to make token requests to the Azure OAuth 2.0 token endpoint.
- [Command Line Parser Library](https://www.nuget.org/packages/CommandLineParser): We'll use this to handle our command line arguments.
- [Json.NET](https://www.nuget.org/packages/Newtonsoft.Json): We'll use this to parse JSON responses from the Outlook API. 

These steps were written with Visual Studio 2015.

1. Create a C# console application named `outlook-fetch`.

1. On the **Tools** menu, choose **NuGet Package Manager**, then **Package Manager Console**.

1. In the **Package Manager Console**, enter the following command:

  ```Shell
  Install-Package Microsoft.IdentityModel.Clients.ActiveDirectory
  ```

1. In the **Package Manager Console**, enter the following command:

  ```Shell
  Install-Package CommandLineParser
  ```

1. In the **Package Manager Console**, enter the following command:

  ```Shell
  Install-Package Newtonsoft.Json
  ```

### Handle the command line

Right-click the project in **Solution Explorer** and choose **Add**, then **Class**. Name the new class `Options` and click **Add**. Replace the entire contents of the **Options.cs** file with the following code:

```C#
using CommandLine;
using CommandLine.Text;

namespace outlook_fetch
{
    class Options
    {
        [Option("certfile", Required = true, 
            HelpText = "The path to the PFX file containing the private key for your application")]
        public string CertFile { get; set; }

        [Option("certpass", Required = true, 
            HelpText = "The password for the PFX file containing the private key for your application")]
        public string CertPass { get; set; }

        [Option("user", Required = true,
            HelpText = "The SMTP address of the user to fetch data for")]
        public string UserEmail { get; set; }

        [Option("mail",
            HelpText = "If specified, the app will fetch the user's 10 most recent email messages")]
        public bool FetchMail { get; set; }

        [Option("calendar",
            HelpText = "If specified, the app will fetch the user's events for the next 7 days")]
        public bool FetchCalendar { get; set; }

        [Option("contacts",
            HelpText = "If specified, the app will fetch the user's 10 most recently added contacts")]
        public bool FetchContacts { get; set; }

        [HelpOption]
        public string GetUsage()
        {
            return HelpText.AutoBuild(this,
                (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
        }
    }
}
```

Open the **Program.cs** file. Add the following code to the `Main` function:

```C#
Options options = new Options();
if (CommandLine.Parser.Default.ParseArguments(args, options))
{
    Console.WriteLine("All required arguments present");
}

// Keep the window open when running in debugger
if (System.Diagnostics.Debugger.IsAttached)
{
    Console.WriteLine("Press any key to exit.");
    Console.ReadKey();
}
```

Now let's add a couple of helper functions to print errors and values. Add the following functions to the `Program` class:

```C#
static void PrintError(string[] errorlines)
{
    Console.ForegroundColor = ConsoleColor.Red;
    foreach (string line in errorlines)
    {
        Console.WriteLine(line);
    }
    Console.ResetColor();
}

static void PrintValue(string name, string value)
{
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.Write(name + ": ");
    Console.ResetColor();
    Console.WriteLine(value);
}
```

Save your changes and run the app in the debugger (press **F5** or click **Start**). You should see a listing of the available command line arguments. Now we're ready to do some real work.

### Get the access token

The first step for our application is to get an access token. This access token will be included in every Outlook API request we make.

Because token requests (and API requests for that matter) happen over HTTP, these operations are asynchronous. Our console app needs to be able to handle that. So to make things easy, we'll create an asynchronous "main" method called `MainAsync`. Add the following using statement to **Program.cs**:

```C#
using System.Threading.Tasks;
```

Add the following function to the `Program` class:

```C#
static async Task MainAsync(Options options)
{
}
```

Replace the `Console.WriteLine("All required arguments present");` line in `Main` with the following:

```C#
MainAsync(options).Wait();
```

Now we can call asynchronous methods from inside `MainAsync` and the `Main` method will wait until they are all completed.

Now let's add the client ID and authorization endpoint values you generated in the Azure Management Portal to the code. Add the following variables to the `Program` class, substituting your values for the placeholder text:

```C#
static string authority = "YOUR AUTHORIZATION ENDPOINT";
static string clientId = "YOUR CLIENT ID";
```

Now add a function to get the access token. Add the following using statements to **Program.cs**:

```C#
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
```

Then add the following function to the `Program` class:

```C#
static async Task<string> GetAppToken(string certFile, string certPass)
{
    try
    {
        // Load the certificate file
        X509Certificate2 cert = new X509Certificate2(certFile, certPass,
          X509KeyStorageFlags.MachineKeySet);

        // Create the ADAL auth context using the authorization endpoint
        // Because the endpoint has the Office 365 tenant ID in it, this allows
        // Azure to issue a token scoped to just that org
        AuthenticationContext authContext = new AuthenticationContext(authority);

        // Create a certificate-based client assertion
        ClientAssertionCertificate assertion = 
          new ClientAssertionCertificate(clientId, cert);

        // Request the token to access Outlook
        AuthenticationResult authResult = 
          await authContext.AcquireTokenAsync("https://outlook.office.com", assertion);

        return authResult.AccessToken;
    }
    catch (CryptographicException cex)
    {
        PrintError(new string[] {
            "There was a problem loading the specified certificate file.",
            string.Format("CERT FILE: {0}", certFile),
            string.Format("ERROR: {0}", cex.Message)
        });
    }
    catch (AdalServiceException aex)
    {
        PrintError(new string[] {
            "There was a problem getting the token from Azure.",
            string.Format("ERROR: {0}", aex.Message)
        });
    }

    return string.Empty;
}
```

Add the following code to the `MainAsync` function to get the token and display it.

```C#
try
{
    string accessToken = await GetAppToken(options.CertFile, options.CertPass);
    if (!string.IsNullOrEmpty(accessToken))
    {
        PrintValue("Access token", accessToken);
    }
    
}
catch (Exception ex)
{
    PrintError(new string[] {
        "There was an unhandled exception.",
        string.Format("MESSAGE: {0}", ex.Message),
        string.Format("STACK: {0}", ex.StackTrace)
    });
}
```

Let's test it out! We need to pass the required arguments `certfile`, `certpass`, and `user`. We can add these arguments to the project's debug settings so they are passed every time you run the app in the debugger. Right-click the project in **Solution Explorer** and choose **Properties**. Click the **Debug** item in the left-hand list. Enter the following into the **Command line arguments** box, replacing placeholders with actual values:

```Shell
--certfile <path to your outlook-fetch-priv.pfx file> --certpass <the password you specified when exporting the private key> --user <a user in your office 365 org>
```

![](.\readme-images\visual-studio-command-line.PNG)

Save the project and start debugging. You should see the access token printed to the console. Now that we can get an access token, we're ready to call the API.

### Call the Outlook API

Right-click the project in **Solution Explorer** and choose **Add**, then **New Folder**. Name the folder `Outlook` and click **Add**. We'll use this to contain all of our code to call the Outlook API.

Right-click the **Outlook** folder and choose **Add**, then **Class**. Name the class `ApiClient` and click **Add**. Replace the entire contents of the **ApiClient.cs** file with the following code:

```C#
using System;
using System.Collections.Generic;
using System.Net.Http;
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
                var request = new HttpRequestMessage(new HttpMethod(method), apiUrl);

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
    }
}
```

The `MakeApiCall` function handles the details of constructing the HTTP REST request and sending it. Let's test it out by making a call to get the user. 

Add the following function to the `Program` class:

```C#
static async Task GetUserInfo(Outlook.ApiClient client, string userEmail)
{
    Console.WriteLine();
    Console.WriteLine("Requesting user information...");

    string requestUrl = string.Format("/Users/{0}", userEmail)

    var result = await client.MakeApiCall("GET", requestUrl, userEmail, null, null);

    // Read the JSON response
    string response = await result.Content.ReadAsStringAsync();

    // Pretty-print the JSON
    string prettyResponse = Newtonsoft.Json.JsonConvert.SerializeObject(
                Newtonsoft.Json.JsonConvert.DeserializeObject(response), 
                Newtonsoft.Json.Formatting.Indented);

    Console.WriteLine(prettyResponse);
}
```

Save your changes and run the app. You should see a JSON response for the user you specified in the `--user` argument.

```JSON
{
  "@odata.context": "https://outlook.office.com/api/v2.0/$metadata#Users/$entity",
  "@odata.id": "https://outlook.office.com/api/v2.0/Users('f862f751-bde1-4d98-a2ff-1a16c3d4218a@5bd31d58-6591-4964-8798-4dab3e9832ba')",
  "Id": "f862f751-bde1-4d98-a2ff-1a16c3d4218a@5bd31d58-6591-4964-8798-4dab3e9832ba",
  "EmailAddress": "allieb@contoso.onmicrosoft.com",
  "DisplayName": "Allie Bellew",
  "Alias": "AllieB",
  "MailboxGuid": "4f011c8c-c42f-0f64-bd58-4d298a7c2598"
}
```

That shows us that the API calls are working, but we can do better! We used the Json.NET library to pretty-print the JSON response, but it can do much more useful things with the JSON. Let's create a class to represent the user, and use Json.NET to deserialize the JSON into our class.

Right-click the **Outlook** folder and choose **Add**, then **Class**. Name the class `User` and click **Add**. Replace the entire contents of the **User.cs** file with the following code:

```C#
namespace outlook_fetch.Outlook
{
    class User
    {
        public string Id { get; set; }
        public string DisplayName { get; set; }
        public string EmailAddress { get; set; }
        public string Alias { get; set; }
        public string MailboxGuid { get; set; }
    }
}
```

Now let's add a `GetUser` function to the `ApiClient` class:

```C#
public async Task<User> GetUser(string userEmail)
{
    string requestUrl = string.Format("/Users/{0}?$select=Foobar", userEmail);

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
```

Notice that if the HTTP request is not successful, the function calls `GetExceptionFromError`, which doesn't exist. Let's implement that now. Add the following to the `ApiClient` class:

```C#
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
```

We then need to implement the `ErrorResponse` class. Right-click the **Outlook** folder and choose **Add**, then **Class**. Name the class `Error` and click **Add**. Replace the entire contents of the **Error.cs** file with the following code:

```C#
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
```

Finally let's update the `GetUserInfo` function in the `Program` class to call our new function. Replace the existing function with this new one:

```C#
static async Task GetUserInfo(Outlook.ApiClient client, string userEmail)
{
    Console.WriteLine();
    Console.WriteLine("Requesting user information...");

    try
    {
        Outlook.User user = await client.GetUser(userEmail);

        PrintValue("User", user.DisplayName);
        PrintValue("Email Address", user.EmailAddress);
    }
    catch (HttpRequestException hex)
    {
        PrintError(new string[]
        {
            "There was an error making the API call to get the user info.",
            string.Format("ERROR: {0}", hex.Message)
        });
    }
}
```

Save your changes and run the app. You should see the user's display name and email address in the output instead of the raw JSON.

Now that we've established a pattern for making API calls and working with the response data, let's tackle implementing code to fetch the user's data. Remember the `--mail`, `--calendar`, and `--contacts` arguments we included in the `Options` class? Let's add some code to `MainAsync` to check those. Add the following code immediately after the `await GetUserInfo(client, options.UserEmail);` line:

```C#
if (options.FetchMail)
{
    // Fetch users's mail
}

if (options.FetchCalendar)
{
    // Fetch user's calendar
}

if (options.FetchContacts)
{
    // Fetch user's contacts
}
```

### Deserializing arrays in API responses

Before we look at specific API calls, let's look at how the API responds when a request is made for multiple items. For example, if we do the following:

```HTTP
GET https://outlook.office.com/api/v2.0/me/messages?$top=5
```

The API responds with JSON that is structured like so:

```JSON
{
  "value": [
    {
      // Message entity
    },
    {
      // Message entity
    },
    {
      // Message entity
    },
    {
      // Message entity
    },
    {
      // Message entity
    }
  ],
  "@odata.nextLink": "https://outlook.office.com/api/v2.0/me/messages/?%24top=5&%24skip=5"
}
```

The array of messages is in the `value` property, and there's also an `@odata.nextLink` property that gives us the request URL we can use to get the next page of results. The API uses this pattern anytime it returns a collection, which is good for us. In order to get Json.NET to deserialize the response properly, we need to define this structure. Since it's the same across all calls, we can do it just once!

Right-click the **Outlook** folder and choose **Add**, then **Class**. Name the class `ItemCollection` and click **Add**. Replace the entire contents of the **ItemCollection.cs** file with the following code:

```C#
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
```

We'll make use of this class soon.

### Call the Mail API

NYI

### Call the Calendar API

NYI

### Call the Contacts API

First let's create a `Contact` class and some supporting classes for the complex fields like `HomeAddress`. A full listing of the fields available on a contact can be found [here](https://msdn.microsoft.com/office/office365/api/complex-types-for-mail-contacts-calendar#RESTAPIResourcesContact). For this example we'll only use a subset.

Right-click the **Outlook** folder and choose **Add**, then **Class**. Name the class `Contact` and click **Add**. Replace the entire contents of the **Contact.cs** file with the following code:

```C#
using System;

namespace outlook_fetch.Outlook
{
    class Contact
    {
        public PhysicalAddress BusinessAddress { get; set; }
        public string[] BusinessPhones { get; set; }
        public DateTimeOffset CreatedDateTime { get; set; }
        public DateTimeOffset LastModifiedDateTime { get; set; }
        public string DisplayName { get; set; }
        public EmailAddress[] EmailAddresses { get; set; }
        public string Id { get; set; }
        public string MobilePhone1 { get; set; }
        public string PersonalNotes { get; set; }
        public string GivenName { get; set; }
        public string Surname { get; set; }
    }
}
```

Right-click the **Outlook** folder and choose **Add**, then **Class**. Name the class `PhysicalAddress` and click **Add**. Replace the entire contents of the **PhysicalAddress.cs** file with the following code:

```C#
namespace outlook_fetch.Outlook
{
    class PhysicalAddress
    {
        public string Street { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string CountryOrRegion { get; set; }
        public string PostalCode { get; set; }
    }
}
```

Right-click the **Outlook** folder and choose **Add**, then **Class**. Name the class `ConEmailAddresstact` and click **Add**. Replace the entire contents of the **EmailAddress.cs** file with the following code:

```C#
namespace outlook_fetch.Outlook
{
    class EmailAddress
    {
        public string Name { get; set; }
        public string Address { get; set; }
    }
}
```

Now let's add a function to the `ApiClient` class to get a user's contacts.

```C#
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
```

Now add the following function to the `Program` class:

```C#
static async Task GetUsersContacts(Outlook.ApiClient client, string userEmail)
{
    Console.WriteLine();
    Console.WriteLine("Fetching user's contacts...");

    try
    {
            Outlook.ItemCollection<Outlook.Contact> contacts = 
            await client.GetContacts(userEmail);

        Console.WriteLine("Fetched {0} contacts", contacts.Items.Count);
        foreach (Outlook.Contact contact in contacts.Items)
        {
            PrintValue("Created", contact.CreatedDateTime.ToString());
            PrintValue("Name", contact.DisplayName);
            if (null != contact.EmailAddresses && contact.EmailAddresses.Length > 0)
            {
                PrintValue("Email", contact.EmailAddresses[0].Address);
            }
            if (!string.IsNullOrEmpty(contact.PersonalNotes))
            {
                PrintValue("Notes", contact.PersonalNotes);
            }
            Console.WriteLine();
        }
    }
    catch (HttpRequestException hex)
    {
        PrintError(new string[]
        {
            "There was an error making the API call to get the user info.",
            string.Format("ERROR: {0}", hex.Message)
        });
    }
}
```

And finally insert the following code after the `// Fetch user's contacts` line in `MainAsync`:

```C#
await GetUsersContacts(client, options.UserEmail);
```

Update the **Command line arguments** in the project properties to include the `--contacts` argument. Save all of your changes and run the app. You should see a list of the user's contacts.