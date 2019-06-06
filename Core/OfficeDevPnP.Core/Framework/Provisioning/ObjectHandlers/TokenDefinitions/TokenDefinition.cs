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
    public abstract class TokenDefinition
    {
        private bool _isCacheable = true;
        private ClientContext _context;
        protected string CacheValue;
        private readonly string[] _tokens;

        /// <summary>
        /// Defines if a token is cacheable and should be added to the token cache during initialization of the token parser. This means that the value for a token will be returned from the cache instead from the GetReplaceValue during the provisioning run. Defaults to true.
        /// </summary>
        public bool IsCacheable
        {
            get
            {
                return _isCacheable;
            }
            set
            {
                _isCacheable = value;
            }

        }
        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="web">Current site/subsite</param>
        /// <param name="token">token</param>
        public TokenDefinition(Web web, params string[] token)
        {
            this._tokens = token;
            this.Web = web;
        }

        /// <summary>
        /// Returns a cloned context which is separate from the current context, not affecting ongoing queries.
        /// </summary>
        public ClientContext TokenContext
        {
            get
            {
                if (_context == null)
                {
                    // Make sure that the Url property has been loaded on the web in the constructor
                    _context = Web.Context.Clone(Web.Url);
                }
                return _context;
            }
        }

        /// <summary>
        /// Gets tokens
        /// </summary>
        /// <returns>Returns array string of tokens</returns>
        public string[] GetTokens()
        {
            return _tokens;
        }

        // public string[] Token { get; private set; }
        /// <summary>
        /// Web is a SiteCollection or SubSite
        /// </summary>
        public Web Web { get; set; }

        /// <summary>
        /// Gets array of regular expressions
        /// </summary>
        /// <returns>Returns all Regular Expressions</returns>
        [Obsolete("No longer in use")]
        public Regex[] GetRegex()
        {
            var regexs = new Regex[this._tokens.Length];
            for (var q = 0; q < this._tokens.Length; q++)
            {
                regexs[q] = new Regex(this._tokens[q], RegexOptions.IgnoreCase);
            }
            return regexs;
        }

        /// <summary>
        /// Gets regular expressionf for the given token
        /// </summary>
        /// <param name="token">token string</param>
        /// <returns>Returns RegularExpression</returns>
        [Obsolete("No longer in use")]
        public Regex GetRegexForToken(string token)
        {
            return new Regex(token, RegexOptions.IgnoreCase);
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