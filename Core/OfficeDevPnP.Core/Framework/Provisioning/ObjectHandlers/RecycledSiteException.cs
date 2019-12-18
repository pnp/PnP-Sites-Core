using System;
using System.Runtime.Serialization;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Initializes a new instance of the RecycledSiteException class. This Exception occurs when the provisioning
    /// engine targets a site that is in the recycle bin
    /// </summary>
    [Serializable]
    public sealed class RecycledSiteException : Exception
    {
        /// <summary>
        /// Initializes a new instance of the RecycledSiteException class with a system supplied message
        /// </summary>
        public RecycledSiteException() : base()
        {
        }

        /// <summary>
        /// Initializes a new instance of the RecycledSiteException class with the specified message string.
        /// </summary>
        /// <param name="message"> A string that describes the exception.</param>
        public RecycledSiteException(string message) : base(message)
        {
        }

        /// <summary>
        /// Initializes a new instance of the RecycledSiteException class with a specified error message and a reference to the inner exception that
        /// is the cause of this exception.
        /// </summary>
        /// <param name="message">A string that describes the exception.</param>
        /// <param name="innerException">The exception that is the cause of the current exception.</param>
        public RecycledSiteException(string message, Exception innerException) : base(message, innerException)
        {
        }

        /// <summary>
        /// Initializes a new instance of the RecycledSiteException class from serialized data.
        /// </summary>
        /// <param name="info">The object that contains the serialized data.</param>
        /// <param name="context">The stream that contains the serialized data.</param>
        /// <exception cref="System.ArgumentNullException">The info parameter is null.-or-The context parameter is null.</exception>
        private RecycledSiteException(SerializationInfo info, StreamingContext context) : base(info, context)
        {
        }
    }
}
