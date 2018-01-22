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
        private string _keyValue;
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

        public string Key
        {
            get { return _keyValue; }
            set { _keyValue = value; }
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
        /// <summary>
        /// Constructor for DataRow class
        /// </summary>
        public DataRow()
        {
            this.Security = new ObjectSecurity();
        }

        /// <summary>
        /// Constructor for DataRow class
        /// </summary>
        /// <param name="values">DataRow Values</param>
        public DataRow(Dictionary<string, string> values) : this(values, null, null)
        { }

        /// <summary>
        /// Constructor for DataRow class
        /// </summary>
        /// <param name="values">DataRow Values</param>
        /// <param name="key">Key column value in case of KeyColumn it set on collection</param>
        public DataRow(Dictionary<string, string> values, string keyValue) : this(values, null, keyValue)
        { }

        /// <summary>
        /// Constructor for DataRow class
        /// </summary>
        /// <param name="values">DataRow Values</param>
        /// <param name="security">ObjectSecurity object</param>
        public DataRow(Dictionary<string, string> values, ObjectSecurity security) : this(values, security, null)
        { }

        /// <summary>
        /// Constructor for DataRow class
        /// </summary>
        /// <param name="values">DataRow Values</param>
        /// <param name="security">ObjectSecurity object</param>
        /// <param name="keyValue">Key column value in case of KeyColumn it set on collection</param>
        public DataRow(Dictionary<string, string> values, ObjectSecurity security, string keyValue) :
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
            this.Key = keyValue;
        }

        #endregion

        #region Comparison code
        /// <summary>
        /// Gets the hash code
        /// </summary>
        /// <returns>Returns HashCode</returns>
        public override int GetHashCode()
        {
            return (String.Format("{0}|{1}|{2}",
                this.Values.Aggregate(0, (acc, next) => acc += next.GetHashCode()),
                (this.Security != null ? this.Security.GetHashCode() : 0),
                (this.Key != null ? this.Key.GetHashCode() : 0)
            ).GetHashCode());
        }

        /// <summary>
        /// Compares object with DataRow
        /// </summary>
        /// <param name="obj">Object that represents DataRow</param>
        /// <returns>true if the current object is equal to the DataRow</returns>
        public override bool Equals(object obj)
        {
            if (!(obj is DataRow))
            {
                return (false);
            }
            return (Equals((DataRow)obj));
        }

        /// <summary>
        /// Compares DataRow object based on values and Security properties.
        /// </summary>
        /// <param name="other">DataRow object</param>
        /// <returns>true if the DataRow object is equal to the current object; otherwise, false.</returns>
        public bool Equals(DataRow other)
        {
            if (other == null)
            {
                return (false);
            }

            return (this.Values.DeepEquals(other.Values) &&
                    (this.Security != null ? this.Security.Equals(other.Security) : true) &&
                    (this.Key != null ? this.Key.Equals(other.Key) : true)
                );
        }

        #endregion
    }
}
