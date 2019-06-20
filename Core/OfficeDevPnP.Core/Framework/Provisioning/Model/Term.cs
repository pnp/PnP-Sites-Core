using System;
using System.Collections.Generic;
using System.IdentityModel.Tokens;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class Term : BaseModel, IEquatable<Term>
    {
        #region Private Members
        private TermCollection _terms;
        private TermLabelCollection _labels;
        private Dictionary<string, string> _properties = new Dictionary<string, string>();
        private Dictionary<string, string> _localProperties = new Dictionary<string, string>();
        private Guid _id;
        #endregion

        #region Public Members
        /// <summary>
        /// Gets or sets the term id
        /// </summary>
        public Guid Id { get { return _id; } set { _id = value; } }
        /// <summary>
        /// Gets or sets the term name
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Gets or sets the term description
        /// </summary>
        public String Description { get; set; }
        /// <summary>
        /// Gets or sets the term owner
        /// </summary>
        public String Owner { get; set; }
        /// <summary>
        /// Gets or sets the IsAvailableForTagging flag for the term
        /// </summary>
        public Boolean IsAvailableForTagging { get; set; }
        /// <summary>
        /// Gets or sets the IsReused flag for the term
        /// </summary>
        public Boolean IsReused { get; set; }
        /// <summary>
        /// Gets or sets the IsSourceTerm flag for the term
        /// </summary>
        public Boolean IsSourceTerm { get; set; }
        /// <summary>
        /// Gets or sets the term source id
        /// </summary>
        public Guid SourceTermId { get; set; }
        /// <summary>
        /// Gets or sets the IsDeprecated flag for the term
        /// </summary>
        public Boolean IsDeprecated { get; set; }
        /// <summary>
        /// Gets or sets Language of the term
        /// </summary>
        public int? Language { get; set; }
        /// <summary>
        /// Gets or sets the order of the term
        /// </summary>
        public int CustomSortOrder { get; set; }
        /// <summary>
        /// Gets or sets terms
        /// </summary>
        public TermCollection Terms
        {
            get { return _terms; }
            private set { _terms = value; }
        }
        /// <summary>
        /// Gets or sets term labels
        /// </summary>
        public TermLabelCollection Labels
        {
            get { return _labels; }
            private set { _labels = value; }
        }
        /// <summary>
        /// Gets or sets the properties of the term
        /// </summary>
        public Dictionary<string, string> Properties
        {
            get { return _properties; }
            private set { _properties = value; }
        }
        /// <summary>
        /// Gets or sets local properties for the term
        /// </summary>
        public Dictionary<string, string> LocalProperties
        {
            get { return _localProperties; }
            private set { _localProperties = value; }
        }
        #endregion

        #region Constructors
        /// <summary>
        /// Constructor for Term class
        /// </summary>
        public Term()
        {
            this._terms = new TermCollection(this.ParentTemplate);
            this._labels = new TermLabelCollection(this.ParentTemplate);
        }

        /// <summary>
        /// Constructor for Term class
        /// </summary>
        /// <param name="id">Id of the term</param>
        /// <param name="name">Name of the term</param>
        /// <param name="language">Language of the term</param>
        /// <param name="terms">Terms</param>
        /// <param name="labels">Labels of the term</param>
        /// <param name="properties">Properties of the term</param>
        /// <param name="localProperties">LocalProperties of the term</param>
        public Term(Guid id, string name, int? language, List<Term> terms, List<TermLabel> labels, Dictionary<string, string> properties, Dictionary<string, string> localProperties):
            this()
        {
            this.Id = id;
            this.Name = name;
            if (language.HasValue)
            {
                this.Language = language;
            }

            this.Terms.AddRange(terms);
            this.Labels.AddRange(labels);

            if (properties != null)
            {
                foreach (var property in properties)
                {
                    this.Properties.Add(property.Key, property.Value);
                }
            }
            if (localProperties != null)
            {
                foreach (var property in localProperties)
                {
                    this.LocalProperties.Add(property.Key, property.Value);
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
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|{11}|{12}|{13}|{14}",
                (this.Id != null ? this.Id.GetHashCode() : 0),
                (this.Name != null ? this.Name.GetHashCode() : 0),
                (this.Description != null ? this.Description.GetHashCode() : 0),
                (this.Language != null ? this.Language.GetHashCode() : 0),
                (this.Owner != null ? this.Owner.GetHashCode() : 0),
                this.IsAvailableForTagging.GetHashCode(),
                this.IsReused.GetHashCode(),
                this.IsSourceTerm.GetHashCode(),
                this.SourceTermId.GetHashCode(),
                this.IsDeprecated.GetHashCode(),
                this.CustomSortOrder.GetHashCode(),
                this.Labels.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Terms.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Properties.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                this.LocalProperties.Aggregate(0, (acc, next) => acc += next.GetHashCode())
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with Term
        /// </summary>
        /// <param name="obj">Object that represents Term</param>
        /// <returns>true if the current object is equal to the Term</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is Term))
            {
                return (false);
            }
            return (Equals((Term)obj));
        }

        /// <summary>
        /// Compares Term object based on Id, Name, Description, Language, Owner, IsAvailableForTagging, IsReused, IsSourceTerm, SourceTermId, 
        /// IsDeprecated, CustomSortOrder, Labels, Terms, Properties and LocalProperties.
        /// </summary>
        /// <param name="other">Term object</param>
        /// <returns>true if the Term object is equal to the current object; otherwise, false.</returns>
        public bool Equals(Term other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Id == other.Id &&
                this.Name == other.Name &&
                this.Description == other.Description &&
                this.Language == other.Language &&
                this.Owner == other.Owner &&
                this.IsAvailableForTagging == other.IsAvailableForTagging &&
                this.IsReused == other.IsReused &&
                this.IsSourceTerm == other.IsSourceTerm &&
                this.SourceTermId == other.SourceTermId &&
                this.IsDeprecated == other.IsDeprecated &&
                this.CustomSortOrder == other.CustomSortOrder &&
                this.Labels.DeepEquals(other.Labels) &&
                this.Terms.DeepEquals(other.Terms) &&
                this.Properties.DeepEquals(other.Properties) &&
                this.LocalProperties.DeepEquals(other.LocalProperties));
        }

        #endregion
    }
}
