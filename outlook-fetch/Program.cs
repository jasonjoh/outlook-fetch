using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System;
using System.Net.Http;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Threading.Tasks;

namespace outlook_fetch
{
    class Program
    {
        static string authority = "YOUR AUTHORIZATION ENDPOINT";
        static string clientId = "YOUR CLIENT ID";
        
        static void Main(string[] args)
        {
            Options options = new Options();
            if (CommandLine.Parser.Default.ParseArguments(args, options))
            {
                MainAsync(options).Wait();
            }

            // Keep the window open when running in debugger
            if (System.Diagnostics.Debugger.IsAttached)
            {
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
            }
        }

        static async Task MainAsync(Options options)
        {
            try
            {
                string accessToken = await GetAppToken(options.CertFile, options.CertPass);
                if (!string.IsNullOrEmpty(accessToken))
                {
                    PrintValue("Access token", accessToken);

                    Outlook.ApiClient client = new Outlook.ApiClient();
                    client.AccessToken = accessToken;

                    await GetUserInfo(client, options.UserEmail);

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
                        await GetUsersContacts(client, options.UserEmail);
                    }
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
        }

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
    }
}
