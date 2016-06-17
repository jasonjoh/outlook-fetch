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
