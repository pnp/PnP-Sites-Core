using System;
using System.Collections.Generic;
using System.Linq;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class TermGroup : BaseModel, IEquatable<TermGroup>
    {
        #region Private Members
        private TermSetCollection _termSets;
        private Guid _id;
        #endregion

        #region Public Members
        public Guid Id { get { return _id; } set { _id = value; } }
        public string Name { get; set; }
        public string Description { get; set; }

        public TermSetCollection TermSets
        {
            get { return _termSets; }
            private set { _termSets = value; }
        }

        #endregion

        #region Constructors

        public TermGroup()
        {
            this._termSets = new TermSetCollection(this.ParentTemplate);
        }

        public TermGroup(Guid id, string name, List<TermSet> termSets):
            this()
        {
            this.Id = id;
            this.Name = name;
            this.TermSets.AddRange(termSets);
        }
        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}|{3}",
                (this.Id != null ? this.Id.GetHashCode() : 0),
                (this.Name != null ? this.Name.GetHashCode() : 0),
                (this.Description != null ? this.Description.GetHashCode() : 0),
                this.TermSets.Aggregate(0, (acc, next) => acc += (next != null ? next.GetHashCode() : 0))
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is TermGroup))
            {
                return (false);
            }
            return (Equals((TermGroup)obj));
        }

        public bool Equals(TermGroup other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Id == other.Id &&
                this.Name == other.Name &&
                this.Description == other.Description &&
                this.TermSets.DeepEquals(other.TermSets));
        }

        #endregion
    }
}
