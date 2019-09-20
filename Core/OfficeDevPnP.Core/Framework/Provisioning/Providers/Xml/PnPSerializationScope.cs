using System;
using System.Collections.Generic;
using System.Linq;
#if !NETSTANDARD2_0
using System.Runtime.Remoting.Messaging;
#endif
using System.Text;
using System.Threading.Tasks;
#if NETSTANDARD2_0
using System.Threading;
#endif

namespace OfficeDevPnP.Core.Framework.Provisioning.Providers.Xml
{
    /// <summary>
    /// Internal class to handle a Provisioning Template serialization scope
    /// </summary>
    internal class PnPSerializationScope : IDisposable
    {
        private String _baseSchemaNamespace;
        private String _baseSchemaAssemblyName;
        private readonly PnPSerializationScope _previous;

        public String BaseSchemaNamespace => this._baseSchemaNamespace;
        public String BaseSchemaAssemblyName => this._baseSchemaAssemblyName;

        public PnPSerializationScope(Type schemaTemplateType)
        {
            // Save the scope information
            this._baseSchemaNamespace = schemaTemplateType.Namespace;
            this._baseSchemaAssemblyName = schemaTemplateType.Assembly.FullName;

            // Save the previous scope, if any
            this._previous = Current;

            // Set the new scope to this
            Current = this;
        }

        ~PnPSerializationScope()
        {
            Dispose(false);
        }

        /// <summary>Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.</summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (disposing)
            {
                Current = this._previous;
            }
        }

#if !NETSTANDARD2_0
        public static PnPSerializationScope Current
        {
            get { return CallContext.LogicalGetData(nameof(PnPSerializationScope)) as PnPSerializationScope; }
            set
            {
                System.Configuration.ConfigurationManager.GetSection("system.xml/xmlReader");
                CallContext.LogicalSetData(nameof(PnPSerializationScope), value);
            }
        }
#else
        private static AsyncLocal<PnPSerializationScope> _pnpSerializationScope = new AsyncLocal<PnPSerializationScope>();

        public static PnPSerializationScope Current
        {
            get { return _pnpSerializationScope.Value; }
            set
            {
                _pnpSerializationScope.Value = value;
            }
        }
#endif
    }
}
