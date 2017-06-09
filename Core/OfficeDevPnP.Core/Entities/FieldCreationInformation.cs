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
        public Guid Id { get; set; }
        public string DisplayName { get; set; }
        public string InternalName { get; set; }
        public bool AddToDefaultView { get; set;}
        /// <summary>
        /// List of additional properties that need to be applied to the field on creation
        /// </summary>
        public IEnumerable<KeyValuePair<string, string>> AdditionalAttributes { get; set; }
        public string FieldType { get; protected set; }
        public string Group { get; set; }
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
        public FieldCreationInformation(string fieldType)
        {
            FieldType = fieldType;
        }

        public FieldCreationInformation(FieldType fieldType)
        {
            FieldType = fieldType.ToString();
        }
    }

}
