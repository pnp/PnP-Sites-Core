using System;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Entities
{
    /// <summary>
    /// Class that describes the field creation information
    /// </summary>
    public class FieldCreationInformation
    {
        /// <summary>
        /// Guid of the field
        /// </summary>
        public Guid Id { get; set; }
        /// <summary>
        /// Field display name
        /// </summary>
        public string DisplayName { get; set; }
        /// <summary>
        /// Field internal name
        /// </summary>
        public string InternalName { get; set; }
        /// <summary>
        /// Adds fields to default view if value is true.
        /// </summary>
        public bool AddToDefaultView { get; set;}
        /// <summary>
        /// List of additional properties that need to be applied to the field on creation
        /// </summary>
        public IEnumerable<KeyValuePair<string, string>> AdditionalAttributes { get; set; }
        /// <summary>
        /// List of additional child nodes that need to be included in the CAML field on creation
        /// </summary>
        public IEnumerable<KeyValuePair<string, string>> AdditionalChildNodes { get; set; }
        /// <summary>
        /// Type of the field
        /// </summary>
        public string FieldType { get; protected set; }
        /// <summary>
        /// Group of the field
        /// </summary>
        public string Group { get; set; }
        /// <summary>
        /// Specifies filds is required to enter vlaue or not.
        /// </summary>
        public bool Required { get; set; }

#if !SP2013
        /// <summary>
        /// Ignored currently for SP2016
        /// </summary>
        public Guid ClientSideComponentId { get; set; }
        /// <summary>
        /// Ignored currently for SP2016
        /// </summary>
        public string ClientSideComponentProperties { get; set; }
#endif
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="fieldType">Type of the field</param>
        public FieldCreationInformation(string fieldType)
        {
            FieldType = fieldType;
        }

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="fieldType">Type of the field</param>
        public FieldCreationInformation(FieldType fieldType)
        {
            FieldType = fieldType.ToString();
        }
    }

}
