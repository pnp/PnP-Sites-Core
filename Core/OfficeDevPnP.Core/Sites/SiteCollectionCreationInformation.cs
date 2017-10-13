#if !ONPREMISES
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Sites
{
    public class CommunicationSiteCollectionCreationInformation
    {
        /// <summary>
        /// The fully qualified url (e.g. https://yourtenant.sharepoint.com/sites/mysitecollection) of the site.
        /// </summary>
        public string Url { get; set; }

        /// <summary>
        /// The title of the site to create
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// The owner of the site. Reserved for future use.
        /// </summary>
        public string Owner { get; set; }

        /// <summary>
        /// The Guid of the site design to be used. If specified will override the SiteDesign property
        /// </summary>
        public Guid SiteDesignId { get; set; }

        /// <summary>
        /// The built-in site design to used. If both SiteDesignId and SiteDesign have been specified, the GUID specified as SiteDesignId will be used.
        /// </summary>
        public CommunicationSiteDesign SiteDesign { get; set; }

        /// <summary>
        /// If set to true, file sharing for guest users will be allowed.
        /// </summary>
        public bool AllowFileSharingForGuestUsers { get; set; }

        /// <summary>
        /// The Site classification to use. For instance 'Contoso Classified'. See https://www.youtube.com/watch?v=E-8Z2ggHcS0 for more information
        /// </summary>
        public string Classification { get; set; }

        /// <summary>
        /// The description to use for the site.
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// The language to use for the site. If not specified will default to the language setting of the clientcontext.
        /// </summary>
        public uint Lcid { get; set; }


        public CommunicationSiteCollectionCreationInformation()
        {

        }

        public CommunicationSiteCollectionCreationInformation(string fullUrl, string title, string description = null)
        {
            this.Url = fullUrl;
            this.Title = title;
            this.Description = description;
        }
    }

    public class TeamSiteCollectionCreationInformation
    {
        //{"displayName":"test modernteamsite","alias":"testmodernteamsite","isPublic":true,"optionalParams":{"Description":"","CreationOptions":{"results":[]},"Classification":""}}

        /// <summary>
        /// The fully qualified url (e.g. https://yourtenant.sharepoint.com/sites/mysitecollection) of the site.
        /// </summary>
        public string Alias { get; set; }

        /// <summary>
        /// The title of the site to create
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// The owner of the site. Reserved for future use.
        /// </summary>
        public bool IsPublic { get; set; }

        /// <summary>
        /// The Guid of the site design to be used. If specified will override the SiteDesign property
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// The Site classification to use. For instance 'Contoso Classified'. See https://www.youtube.com/watch?v=E-8Z2ggHcS0 for more information
        /// </summary>
        public string Classification { get; set; }

        public TeamSiteCollectionCreationInformation()
        {

        }

        public TeamSiteCollectionCreationInformation(string alias, string displayName, string description = null)
        {
            this.Alias = alias;
            this.DisplayName = displayName;
            this.Description = description;
        }
    }

    public enum CommunicationSiteDesign
    {
        Topic = 0,
        Showcase = 1,
        Blank = 2,
    }

}
#endif