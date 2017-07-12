using System;
using System.Collections.Generic;
using System.Linq;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class TermSet : BaseModel, IEquatable<TermSet>
    {
        #region Private Members
        private TermCollection _terms;
        private Guid _id;
        private Dictionary<string, string> _properties = new Dictionary<string, string>();
        #endregion

        #region Public Members
        /// <summary>
        /// Gets or sets the termset id
        /// </summary>
        public Guid Id
        {
            get { return _id; }
            set { _id = value; }
        }
        /// <summary>
        /// Gets or sets the termset name
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Gets or sets the termset description
        /// </summary>
        public string Description { get; set; }
        /// <summary>
        /// Gets or sets the termset language
        /// </summary>
        public int? Language { get; set; }
        /// <summary>
        /// Gets or sets the IsOpenForTermCreation flag for the termset
        /// </summary>
        public bool IsOpenForTermCreation { get; set; }
        /// <summary>
        /// Gets or sets the IsAvailableForTagging flag for the termset
        /// </summary>
        public bool IsAvailableForTagging { get; set; }
        /// <summary>
        /// Gets or sets the termset owner
        /// </summary>
        public string Owner { get; set; }
        /// <summary>
        /// Gets or sets the terms
        /// </summary>
        public TermCollection Terms
        {
            get { return _terms; }
            private set { _terms = value; }
        }
        /// <summary>
        /// Gets or sets the termset properties
        /// </summary>
        public Dictionary<string, string> Properties
        {
            get { return _properties; }
            private set { _properties = value; }
        }

        #endregion

        #region Constructors
        /// <summary>
        /// Constructor for TermSet class
        /// </summary>
        public TermSet()
        {
            this._terms = new TermCollection(this.ParentTemplate);
        }

        /// <summary>
        /// Constructor for Termset class
        /// </summary>
        /// <param name="id">Id of the termset</param>
        /// <param name="name">Name of the termset</param>
        /// <param name="language">Language of the termset</param>
        /// <param name="isAvailableForTagging">IsAvailableForTagging flag for termset</param>
        /// <param name="isOpenForTermCreation">IsOpenForTermCreation flag for termset</param>
        /// <param name="terms">Temset terms</param>
        /// <param name="properties">Termset properties</param>
        public TermSet(Guid id, string name, int? language, bool isAvailableForTagging, bool isOpenForTermCreation, List<Term> terms, Dictionary<string, string> properties): 
            this()
        {
            this.Id = id;
            this.Name = name;
            this.Language = language;
            this.IsAvailableForTagging = isAvailableForTagging;
            this.IsOpenForTermCreation = isOpenForTermCreation;
            this.Terms.AddRange(terms);
            if (properties != null)
            {
                foreach (var property in properties)
                {
                    this.Properties.Add(property.Key, property.Value);
                }
            }
        }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}",
                (this.Id != null ? this.Id.GetHashCode() : 0),
                (this.Name != null ? this.Name.GetHashCode() : 0),
                (this.Description != null ? this.Description.GetHashCode() : 0),
                (this.Language != null ? this.Language.GetHashCode() : 0),
                this.IsOpenForTermCreation.GetHashCode(),
                this.IsAvailableForTagging.GetHashCode(),
                (this.Owner != null ? this.Owner.GetHashCode() : 0),
                this.Terms.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Properties.Aggregate(0, (acc, next) => acc += next.GetHashCode())
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with TermSet
        /// </summary>
        /// <param name="obj">Object that represents TermSet</param>
        /// <returns>true if the current object is equal to the TermSet</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is TermSet))
            {
                return (false);
            }
            return (Equals((TermSet)obj));
        }

        /// <summary>
        /// Compares TermSet object based on Id, Name, Description, Language, IsOpenForTermCreation, IsAvailableForTagging, Owner, Terms and Properties.
        /// </summary>
        /// <param name="other">TermSet object</param>
        /// <returns>true if the TermSet object is equal to the current object; otherwise, false.</returns>
        public bool Equals(TermSet other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Id == other.Id &&
                    this.Name == other.Name &&
                    this.Description == other.Description &&
                    this.Language == other.Language &&
                    this.IsOpenForTermCreation == other.IsOpenForTermCreation &&
                    this.IsAvailableForTagging == other.IsAvailableForTagging &&
                    this.Owner == other.Owner &&
                    this.Terms.DeepEquals(other.Terms) &&
                    this.Properties.DeepEquals(other.Properties));
        }

        #endregion
    }
}
