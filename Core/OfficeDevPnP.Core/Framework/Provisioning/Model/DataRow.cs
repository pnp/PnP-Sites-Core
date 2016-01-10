using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Extensions;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public partial class DataRow : BaseModel, IEquatable<DataRow>
    {
        #region Private members
        private Dictionary<string, string> _values = new Dictionary<string, string>();
        private ObjectSecurity _objectSecurity;
        #endregion

        #region Public Members

        /// <summary>
        /// Defines the fields to provision within a row that will be added to the List Instance
        /// </summary>
        public Dictionary<string, string> Values
        {
            get { return _values; }
            private set { _values = value; }
        }

        /// <summary>
        /// Defines the security rules for the row that will be added to the List Instance
        /// </summary>
        public ObjectSecurity Security
        {
            get { return _objectSecurity; }
            private set
            {
                if (this._objectSecurity != null)
                {
                    this._objectSecurity.ParentTemplate = null;
                }
                this._objectSecurity = value;
                if (this._objectSecurity != null)
                {
                    this._objectSecurity.ParentTemplate = this.ParentTemplate;
                }
            }
        }

        #endregion

        #region constructors
        public DataRow()
        {
            this.Security = new ObjectSecurity();
        }

        public DataRow(Dictionary<string, string> values) : this(values, null)
        {
        }

        public DataRow(Dictionary<string, string> values, ObjectSecurity security) :
            this()
        {
            if (values != null)
            {
                foreach (var key in values.Keys)
                {
                    Values.Add(key, values[key]);
                }
            }

            this.Security = security;
        }

        #endregion

        #region Comparison code

        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|",
                this.Values.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                (this.Security != null ? this.Security.GetHashCode() : 0)
            ).GetHashCode());
        }

        public override bool Equals(object obj)
        {
            if (!(obj is DataRow))
            {
                return (false);
            }
            return (Equals((DataRow)obj));
        }

        public bool Equals(DataRow other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Values.DeepEquals(other.Values) &&
                    (this.Security != null ? this.Security.Equals(other.Security) : true)
                );
        }

        #endregion
    }
}
