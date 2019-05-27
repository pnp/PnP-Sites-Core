using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    /// <summary>
    /// Defines a provisioning engine Token. Make sure to only use the TokenContext property to execute queries in token methods.
    /// </summary>
    public abstract class SimpleTokenDefinition
    {
        protected string CacheValue;
        private readonly string[] _tokens;

        /// <summary>
        /// Constructor
        /// </summary>        
        /// <param name="token">token</param>
        public SimpleTokenDefinition(params string[] token)
        {
            this._tokens = token;
        }

        /// <summary>
        /// Gets tokens
        /// </summary>
        /// <returns>Returns array string of tokens</returns>
        public string[] GetTokens()
        {
            return _tokens;
        }

        /// <summary>
        /// Gets token length in integer
        /// </summary>
        /// <returns>token length in integer</returns>
        public int GetTokenLength()
        {
            return _tokens.Select(t => t.Length).Concat(new[] { 0 }).Max();
        }

        /// <summary>
        /// abstract method
        /// </summary>
        /// <returns>Returns string</returns>
        public abstract string GetReplaceValue();

        /// <summary>
        /// Clears cache
        /// </summary>
        public void ClearCache()
        {
            this.CacheValue = null;
        }
    }
}