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
