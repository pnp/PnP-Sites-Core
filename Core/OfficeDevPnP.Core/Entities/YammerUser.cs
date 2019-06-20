using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;

namespace OfficeDevPnP.Core.Entities
{

    /// <summary>
    /// Represents YammerUser
    /// Generated based on Yammer response on 30th of June 2014 and using http://json2csharp.com/ service 
    /// </summary>
    public class YammerUser
    {
        /// <summary>
        /// Represents yammer user type
        /// </summary>
        public string type { get; set; }
        /// <summary>
        /// Represents yammer user id
        /// </summary>
        public int id { get; set; }
        /// <summary>
        /// Represents yammer user network id
        /// </summary>
        public int network_id { get; set; }
        /// <summary>
        /// Represents yammer user state
        /// </summary>
        public string state { get; set; }
        /// <summary>
        /// Represents yammer user Guid
        /// </summary>
        public object guid { get; set; }
        /// <summary>
        /// Represents yammer user job title
        /// </summary>
        public string job_title { get; set; }
        /// <summary>
        /// Represents yammer user location
        /// </summary>
        public object location { get; set; }
        /// <summary>
        /// Represents yammer user other significant
        /// </summary>
        public object significant_other { get; set; }
        /// <summary>
        /// Represents yammer user kids names
        /// </summary>
        public object kids_names { get; set; }
        /// <summary>
        /// Represents yammer user interests
        /// </summary>
        public object interests { get; set; }
        /// <summary>
        /// Represents yammer user summary
        /// </summary>
        public object summary { get; set; }
        /// <summary>
        /// Represents yammer user expertise
        /// </summary>
        public object expertise { get; set; }
        /// <summary>
        /// Represents yammer user full name
        /// </summary>
        public string full_name { get; set; }
        /// <summary>
        /// Represents yammer user activated information
        /// </summary>
        public string activated_at { get; set; }
        /// <summary>
        /// Represents yammer user preferred show ask for photo option or not
        /// </summary>
        public bool show_ask_for_photo { get; set; }
        /// <summary>
        /// Represents yammer user first name
        /// </summary>
        public string first_name { get; set; }
        /// <summary>
        /// Represents yammer user last name
        /// </summary>
        public string last_name { get; set; }
        /// <summary>
        /// Represents yammer user network name
        /// </summary>
        public string network_name { get; set; }
        /// <summary>
        /// Represents yammer user list of network domains
        /// </summary>
        [SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public IList<string> network_domains { get; set; }
        /// <summary>
        /// Represents yammer user url
        /// </summary>
        public string url { get; set; }
        /// <summary>
        /// Represents yammer user web url
        /// </summary>
        public string web_url { get; set; }
        /// <summary>
        /// Represents yammer user name
        /// </summary>
        public string name { get; set; }
        /// <summary>
        /// Represents yammer user mugshot url
        /// </summary>
        public string mugshot_url { get; set; }
        /// <summary>
        /// Represents yammer user mugshot url template
        /// </summary>
        public string mugshot_url_template { get; set; }
        /// <summary>
        /// Represents yammer user hire date
        /// </summary>
        public object hire_date { get; set; }
        /// <summary>
        /// Represents yammer user birth date
        /// </summary>
        public string birth_date { get; set; }
        /// <summary>
        /// Represents yammer user time zone
        /// </summary>
        public string timezone { get; set; }
        /// <summary>
        /// Represents yammer user list of external urls
        /// </summary>
        [SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public IList<object> external_urls { get; set; }
        /// <summary>
        /// Represents yammer user admin
        /// </summary>
        public string admin { get; set; }
        /// <summary>
        /// Represents yammer user verified admin
        /// </summary>
        public string verified_admin { get; set; }
        /// <summary>
        /// Represents yammer user broadcast details
        /// </summary>
        public string can_broadcast { get; set; }
        /// <summary>
        /// Represents yammer user department
        /// </summary>
        public string department { get; set; }
        /// <summary>
        /// Represents yammer user list of previous companies
        /// </summary>
        [SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public IList<object> previous_companies { get; set; }
        /// <summary>
        /// Represents yammer user list of schools
        /// </summary>
        [SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public IList<object> schools { get; set; }
        /// <summary>
        /// Represents yammer user contact details
        /// </summary>
        public YammerUserContact contact { get; set; }
        /// <summary>
        /// Represents yammer user statistics information
        /// </summary>
        public YammerUserStats stats { get; set; }
        /// <summary>
        /// Represents yammer user settings information
        /// </summary>
        public YammerUserSettings settings { get; set; }
        /// <summary>
        /// Represents yammer user web preference information
        /// </summary>
        public YammerUserWebPreferences web_preferences { get; set; }
        /// <summary>
        /// Represents yammer user follows general messages or not 
        /// </summary>
        public bool follow_general_messages { get; set; }
        /// <summary>
        /// Represents yammer user web auth access token
        /// </summary>
        public string web_oauth_access_token { get; set; }
    }

    /// <summary>
    /// Holds Yammer user properties
    /// </summary>
    public class YammerUserIm
    {
        /// <summary>
        /// Yammer user provider name
        /// </summary>
        public string provider { get; set; }
        /// <summary>
        /// Yammer user name
        /// </summary>
        public string username { get; set; }
    }

    /// <summary>
    /// Holds yammer user email address
    /// </summary>
    public class YammerUserEmailAddress
    {
        /// <summary>
        /// Type of email address
        /// </summary>
        public string type { get; set; }
        /// <summary>
        /// Yammer user email address
        /// </summary>
        public string address { get; set; }
    }

    /// <summary>
    /// Holds yammer user contact details
    /// </summary>
    public class YammerUserContact
    {
        /// <summary>
        /// Yammer user details
        /// </summary>
        public YammerUserIm im { get; set; }
        /// <summary>
        /// List of yammer user phone numbers
        /// </summary>
        [SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public IList<object> phone_numbers { get; set; }
        /// <summary>
        /// List of yammer user email addresses
        /// </summary>
        [SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public IList<YammerUserEmailAddress> email_addresses { get; set; }
        /// <summary>
        /// Specifies whether yammer user has fake email or not
        /// </summary>
        public bool has_fake_email { get; set; }
    }

    /// <summary>
    /// Holds yammer user statistics information
    /// </summary>
    public class YammerUserStats
    {
        /// <summary>
        /// Number of memebrs the yammer user following
        /// </summary>
        public int following { get; set; }
        /// <summary>
        /// Number of members following the user
        /// </summary>
        public int followers { get; set; }
        /// <summary>
        /// Number of updates of the user
        /// </summary>
        public int updates { get; set; }
    }

    /// <summary>
    /// Holds yammer user settings
    /// </summary>
    public class YammerUserSettings
    {
        /// <summary>
        /// Represents XDR Proxy
        /// </summary>
        public string xdr_proxy { get; set; }
    }

    /// <summary>
    /// Holds yammer user network settings
    /// </summary>
    public class YammerUserNetworkSettings
    {
        /// <summary>
        /// Yammer user message promt
        /// </summary>
        public string message_prompt { get; set; }
        /// <summary>
        /// Yammer user attachments
        /// </summary>
        public string allow_attachments { get; set; }
        /// <summary>
        /// Represents boolean value to dispay communities directory
        /// </summary>
        public bool show_communities_directory { get; set; }
        /// <summary>
        /// Represents boolean value to enable groups to user
        /// </summary>
        public bool enable_groups { get; set; }
        /// <summary>
        /// Represents boolean value to allow yammer application
        /// </summary>
        public bool allow_yammer_apps { get; set; }
        /// <summary>
        /// Represents admin delegate messages
        /// </summary>
        public string admin_can_delete_messages { get; set; }
        /// <summary>
        /// Represents boolean value to allow inline document view
        /// </summary>
        public bool allow_inline_document_view { get; set; }
        /// <summary>
        /// Represents boolean value to allow inline video
        /// </summary>
        public bool allow_inline_video { get; set; }
        /// <summary>
        /// Represents boolean value to enable private messages
        /// </summary>
        public bool enable_private_messages { get; set; }
        /// <summary>
        /// Represents boolean value to allow external sharing
        /// </summary>
        public bool allow_external_sharing { get; set; }
        /// <summary>
        /// Represents boolean value enable chat
        /// </summary>
        public bool enable_chat { get; set; }
    }

    /// <summary>
    /// Holds yammer user home tab details
    /// </summary>
    public class YammerUserHomeTab
    {
        /// <summary>
        /// Represents name of user
        /// </summary>
        public string name { get; set; }
        /// <summary>
        /// Represents select name of user
        /// </summary>
        public string select_name { get; set; }
        /// <summary>
        /// Represents user type
        /// </summary>
        public string type { get; set; }
        /// <summary>
        /// Represents description of user
        /// </summary>
        public string feed_description { get; set; }
        /// <summary>
        /// Represents index of user
        /// </summary>
        public string ordering_index { get; set; }
        /// <summary>
        /// Represents url of user
        /// </summary>
        public string url { get; set; }
        /// <summary>
        /// Represents group id of user
        /// </summary>
        public int? group_id { get; set; }
        /// <summary>
        /// Represents boolean value to make user information as private
        /// </summary>
        public bool? @private { get; set; }
    }

    /// <summary>
    /// Holds yammer user web preferences
    /// </summary>
    public class YammerUserWebPreferences
    {
        /// <summary>
        /// Represents yammer user full names display
        /// </summary>
        public string show_full_names { get; set; }
        /// <summary>
        /// Represents yammer user absolute time stamps
        /// </summary>
        public string absolute_timestamps { get; set; }
        /// <summary>
        /// Represents yammer user threaded mode
        /// </summary>
        public string threaded_mode { get; set; }
        /// <summary>
        /// Represents yammer user network settings
        /// </summary>
        public YammerUserNetworkSettings network_settings { get; set; }
        /// <summary>
        /// Represents yammer user list of home tabs
        /// </summary>
        [SuppressMessage("Microsoft.Usage", "CA2227:CollectionPropertiesShouldBeReadOnly")]
        public IList<YammerUserHomeTab> home_tabs { get; set; }
        /// <summary>
        /// Represents yammer user message not to be submitted
        /// </summary>
        public string enter_does_not_submit_message { get; set; }
        /// <summary>
        /// Represents yammer user preferred feed
        /// </summary>
        public string preferred_my_feed { get; set; }
        /// <summary>
        /// Represents yammer user prescribed feed
        /// </summary>
        public string prescribed_my_feed { get; set; }
        /// <summary>
        /// Represents yammer user sticky feed
        /// </summary>
        public bool sticky_my_feed { get; set; }
        /// <summary>
        /// Represents yammer user chat enable information
        /// </summary>
        public string enable_chat { get; set; }
        /// <summary>
        /// Represents yammer user dismissed feed tooltip or not
        /// </summary>
        public bool dismissed_feed_tooltip { get; set; }
        /// <summary>
        /// Represents yammer user dismissed group tooltip or not
        /// </summary>
        public bool dismissed_group_tooltip { get; set; }
        /// <summary>
        /// Represents yammer user dismissed profile prompt or not
        /// </summary>
        public bool dismissed_profile_prompt { get; set; }
        /// <summary>
        /// Represents yammer user dismissed tool tip invitation or not
        /// </summary>
        public bool dismissed_invite_tooltip { get; set; }
        /// <summary>
        /// Represents yammer user dismissed apps tooltip or not
        /// </summary>
        public bool dismissed_apps_tooltip { get; set; }
        /// <summary>
        /// Represents yammer user dismissed tooltip invitation location
        /// </summary>
        public string dismissed_invite_tooltip_at { get; set; }
        /// <summary>
        /// Represents yammer user locale
        /// </summary>
        public string locale { get; set; }
        /// <summary>
        /// Represents yammer user current app id
        /// </summary>
        public int yammer_now_app_id { get; set; }
        /// <summary>
        /// Represents yammer user has yammer now or not
        /// </summary>
        public bool has_yammer_now { get; set; }
    }
}
