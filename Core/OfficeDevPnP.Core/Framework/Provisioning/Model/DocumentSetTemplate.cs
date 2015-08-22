using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    /// <summary>
    /// Defines a DocumentSet Template for creating multiple DocumentSet instances
    /// </summary>
    public class DocumentSetTemplate : IEquatable<DocumentSetTemplate>
    {
        /// <summary>
        /// The list of allowed Content Types for the Document Set
        /// </summary>
        public List<String> AllowedContentTypes { get; set; }

        /// <summary>
        /// The list of default Documents for the Document Set
        /// </summary>
        public List<DefaultDocument> DefaultDocuments { get; set; }

        /// <summary>
        /// The list of Shared Fields for the Document Set
        /// </summary>
        public List<Guid> SharedFields { get; set; }

        /// <summary>
        /// The list of Welcome Page Fields for the Document Set
        /// </summary>
        public List<Guid> WelcomePageFields { get; set; }

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|",
                this.AllowedContentTypes.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                this.DefaultDocuments.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                this.SharedFields.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                this.WelcomePageFields.Aggregate(0, (acc, next) => acc += next.GetHashCode())
               ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is DocumentSetTemplate))
            {
                return (false);
            }
            return (Equals((DocumentSetTemplate)obj));
        }

        public bool Equals(DocumentSetTemplate other)
        {
            return (this.AllowedContentTypes.DeepEquals(other.AllowedContentTypes) &&
                    this.DefaultDocuments.DeepEquals(other.DefaultDocuments) &&
                    this.SharedFields.DeepEquals(other.SharedFields) &&
                    this.WelcomePageFields.DeepEquals(other.WelcomePageFields)
                );
        }

        #endregion
    }
}
