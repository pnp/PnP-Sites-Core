using System;
using System.Collections.Generic;

namespace OfficeDevPnP.Core.Framework.Graph.Model
{
    /// <summary>
    /// Defines a Microsoft Graph user
    /// </summary>
    public class User
    {
        /// <summary>
        /// Business phone numbers for the user
        /// </summary>
        public List<string> BusinessPhones { get; set; }

        /// <summary>
        /// Display name for the user
        /// </summary>
        public string DisplayName { get; set; }
        
        /// <summary>
        /// Given name of the user
        /// </summary>
        public string GivenName { get; set; }

        /// <summary>
        /// Job title of the user
        /// </summary>
        public string JobTitle { get; set; }

        /// <summary>
        /// Primary e-mail address of the user
        /// </summary>
        public string Mail { get; set; }

        /// <summary>
        /// Mobile phone number of the user
        /// </summary>
        public string MobilePhone { get; set; }
        
        /// <summary>
        /// Office location of the user
        /// </summary>
        public string OfficeLocation { get; set; }

        /// <summary>
        /// Preferred language of the user
        /// </summary>
        public string PreferredLanguage { get; set; }

        /// <summary>
        /// Surname of the user
        /// </summary>
        public string Surname { get; set; }

        /// <summary>
        /// User Principal Name (UPN) of the user
        /// </summary>
        public string UserPrincipalName { get; set; }

        /// <summary>
        /// Unique identifier of the user
        /// </summary>
        public Guid? Id { get; set; }
    }
}
