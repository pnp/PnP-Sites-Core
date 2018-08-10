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
        private ClientContext _context;
        protected string CacheValue;
        private readonly string[] _tokens;
        private static readonly Dictionary<string, Regex> _tokeRegexes = new Dictionary<string, Regex>(1500);

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
                    var webUrl = Web.EnsureProperty(w => w.Url);
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
        public Regex[] GetRegex()
        {
            if (_tokeRegexes.Count == this._tokens.Length) return _tokeRegexes.Values.ToArray();

            _tokeRegexes.Clear();
            var regexs = new Regex[this._tokens.Length];
            for (var q = 0; q < this._tokens.Length; q++)
            {
                regexs[q] = new Regex(this._tokens[q], RegexOptions.IgnoreCase | RegexOptions.Compiled);
                _tokeRegexes.Add(this._tokens[q], regexs[q]);
            }
            return regexs;
        }

        /// <summary>
        /// Gets regular expressionf for the given token
        /// </summary>
        /// <param name="token">token string</param>
        /// <returns>Returns RegularExpression</returns>
        public Regex GetRegexForToken(string token)
        {
            if (!_tokeRegexes.TryGetValue(token, out var regEx))
            {
                regEx = new Regex(token, RegexOptions.IgnoreCase | RegexOptions.Compiled);
                _tokeRegexes[token] = regEx;
            }
            return regEx;
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