namespace OfficeDevPnP.Core.Entities
{
    /// <summary>
    /// Represents Yammer Group statistics
    /// </summary>
    public class YammerGroupStats
    {
        /// <summary>
        /// Number of users in the yammer group
        /// </summary>
        public int members { get; set; }
        /// <summary>
        /// Number of updates of the yammer group
        /// </summary>
        public int updates { get; set; }
        /// <summary>
        /// yammer group last meesage id
        /// </summary>
        public int? last_message_id { get; set; }
        /// <summary>
        /// yammer group last message location
        /// </summary>
        public string last_message_at { get; set; }
    }

    /// <summary>
    /// Represents Yammer Group information
    /// Generated based on Yammer response on 30th of June 2014 and using http://json2csharp.com/ service 
    /// </summary>
    public class YammerGroup
    {
        /// <summary>
        /// Type of yammer group
        /// </summary>
        public string type { get; set; }
        /// <summary>
        /// Id of yammer group
        /// </summary>
        public int id { get; set; }
        /// <summary>
        /// full name of yammer group
        /// </summary>
        public string full_name { get; set; }
        /// <summary>
        /// yammer group name
        /// </summary>
        public string name { get; set; }
        /// <summary>
        /// yammer group description
        /// </summary>
        public object description { get; set; }
        /// <summary>
        /// privacy of yammer group
        /// </summary>
        public string privacy { get; set; }
        /// <summary>
        /// url of yammer group
        /// </summary>
        public string url { get; set; }
        /// <summary>
        /// web url of yammer group
        /// </summary>
        public string web_url { get; set; }
        /// <summary>
        /// Mugshot url of yammer group
        /// </summary>
        public string mugshot_url { get; set; }
        /// <summary>
        /// Mugshot url temlate of yammer group
        /// </summary>
        public string mugshot_url_template { get; set; }
        /// <summary>
        /// Mugshot id of yammer group
        /// </summary>
        public object mugshot_id { get; set; }
        /// <summary>
        /// string value to be displayed in the directory of yammer value
        /// </summary>
        public string show_in_directory { get; set; }
        /// <summary>
        /// yammer group office 365 url
        /// </summary>
        public object office365_url { get; set; }
        /// <summary>
        /// DateTime of yammer group created
        /// </summary>
        public string created_at { get; set; }
        /// <summary>
        /// yammer group creator type
        /// </summary>
        public string creator_type { get; set; }
        /// <summary>
        /// yammer group creator id
        /// </summary>
        public int creator_id { get; set; }
        /// <summary>
        /// Yammer group state
        /// </summary>
        public string state { get; set; }
        /// <summary>
        /// Yammer group statistics
        /// </summary>
        public YammerGroupStats stats { get; set; }
        // Added manually as extended property which can be set if needed in the code. Set in YammerUtility class code automatically
        /// <summary>
        /// yammer group network id
        /// </summary>
        public int network_id { get; set; }
        /// <summary>
        /// yammer group network name
        /// </summary>
        public string network_name { get; set; }
    }
}
