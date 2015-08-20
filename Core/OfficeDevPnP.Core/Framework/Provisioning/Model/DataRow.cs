using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Framework.Provisioning.Model
{
    public class DataRow
    {
        #region Private members
        private Dictionary<string, string> _values = new Dictionary<string, string>();
        private ObjectSecurity _objectSecurity = new ObjectSecurity();
        #endregion

        #region public members

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
        public ObjectSecurity ObjectSecurity
        {
            get { return _objectSecurity; }
            private set { _objectSecurity = value; }
        }
        #endregion

        #region constructors
        public DataRow()
        {

        }

        public DataRow(Dictionary<string, string> values)
        {
            foreach (var key in values.Keys)
            {
                Values.Add(key, values[key]);
            }
        }
        #endregion
    }
}
