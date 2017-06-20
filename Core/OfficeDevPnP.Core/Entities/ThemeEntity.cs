namespace OfficeDevPnP.Core.Entities
{
    /// <summary>
    /// Holds properties for SharePoint Theme
    /// </summary>
    public class ThemeEntity
    {
        /// <summary>
        /// Name of the Theme
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Sets the theme as custom composed look
        /// </summary>
        public bool IsCustomComposedLook { get; set; }
        /// <summary>
        /// Master page url
        /// </summary>
        public string MasterPage { get; set; }
        /// <summary>
        /// Custom master page url
        /// </summary>
        public string CustomMasterPage { get; set; }
        /// <summary>
        /// Theme url
        /// </summary>
        public string Theme { get; set; }
        /// <summary>
        /// Background image url
        /// </summary>
        public string BackgroundImage { get; set; }
        /// <summary>
        /// Font scheme url
        /// </summary>
        public string Font { get; set; }
    }
}
