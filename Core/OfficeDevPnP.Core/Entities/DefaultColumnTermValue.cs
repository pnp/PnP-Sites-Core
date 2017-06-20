using System.Collections.Generic;
using Microsoft.SharePoint.Client.Taxonomy;

namespace OfficeDevPnP.Core.Entities
{
    /// <summary>
    /// Holds default column value properties
    /// </summary>
    public interface IDefaultColumnValue
    {
        /// <summary>
        /// Folder relative path
        /// </summary>
        string FolderRelativePath { get; set; }

        /// <summary>
        /// Field internal name
        /// </summary>
        string FieldInternalName { get; set; }
    }
    
    /// <summary>
    /// Specifies a default column value for a document library
    /// </summary>
    public class DefaultColumnTermValue : DefaultColumnValue
    {
      
        /// <summary>
        /// Taxonomy paths in the shape of "TermGroup|TermSet|Term"
        /// </summary>
        public IList<Term> Terms { get; private set; }

        /// <summary>
        /// Default Constructor
        /// </summary>
        public DefaultColumnTermValue()
        {
            Terms = new List<Term>();
        }
    }

    /// <summary>
    /// Holds DefaultColumnTextValue Properties
    /// </summary>
    public class DefaultColumnTextValue : DefaultColumnValue
    {
        /// <summary>
        /// Default column text value
        /// </summary>
        public string Text { get; set; }
    }

    /// <summary>
    /// Holds DefalutColumnValue Properties
    /// </summary>
    public class DefaultColumnValue : IDefaultColumnValue
    {
        /// <summary>
        /// The Path of the folder, Rootfolder of the document library is "/" 
        /// </summary>
        public string FolderRelativePath { get; set; }

        /// <summary>
        /// The internal name of the field
        /// </summary>
        public string FieldInternalName { get; set; }

    }
}
