using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.SharePoint.Client;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions
{
    /// <summary>
    /// Handles Token methods
    /// </summary>
    public abstract class TokenDefinition
    {
        protected string CacheValue;
        private readonly string[] _tokens;

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