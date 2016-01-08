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
        public Guid Id { get { return _id; } set { _id = value; } }
        public string Name { get; set; }
        public String Description { get; set; }
        public String Owner { get; set; }
        public Boolean IsAvailableForTagging { get; set; }
        public int? Language { get; set; }
        public int CustomSortOrder { get; set; }

        public TermCollection Terms
        {
            get { return _terms; }
            private set { _terms = value; }
        }

        public TermLabelCollection Labels
        {
            get { return _labels; }
            private set { _labels = value; }
        }

        public Dictionary<string, string> Properties
        {
            get { return _properties; }
            private set { _properties = value; }
        }

        public Dictionary<string, string> LocalProperties
        {
            get { return _localProperties; }
            private set { _localProperties = value; }
        }
        #endregion

        #region Constructors

        public Term()
        {
            this._terms = new TermCollection(this.ParentTemplate);
            this._labels = new TermLabelCollection(this.ParentTemplate);
        }

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

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}|{8}|{9}|{10}|",
                (this.Id != null ? this.Id.GetHashCode() : 0),
                (this.Name != null ? this.Name.GetHashCode() : 0),
                (this.Description != null ? this.Description.GetHashCode() : 0),
                (this.Language != null ? this.Language.GetHashCode() : 0),
                (this.Owner != null ? this.Owner.GetHashCode() : 0),
                this.IsAvailableForTagging.GetHashCode(),
                this.CustomSortOrder.GetHashCode(),
                this.Labels.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Terms.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0)),
                this.Properties.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                this.LocalProperties.Aggregate(0, (acc, next) => acc += next.GetHashCode())
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is Term))
            {
                return (false);
            }
            return (Equals((Term)obj));
        }

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
                this.CustomSortOrder == other.CustomSortOrder &&
                this.Labels.DeepEquals(other.Labels) &&
                this.Terms.DeepEquals(other.Terms) &&
                this.Properties.DeepEquals(other.Properties) &&
                this.LocalProperties.DeepEquals(other.LocalProperties));
        }

        #endregion
    }
}
