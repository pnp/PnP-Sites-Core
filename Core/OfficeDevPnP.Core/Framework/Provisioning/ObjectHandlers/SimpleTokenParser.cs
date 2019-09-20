using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers.TokenDefinitions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers
{
    /// <summary>
    /// Handles methods for token parser
    /// </summary>
    internal class SimpleTokenParser
    {
        private List<SimpleTokenDefinition> _tokens;

        public SimpleTokenParser()
        {
            _tokens = new List<SimpleTokenDefinition>();
        }

        /// <summary>
        /// List of token definitions
        /// </summary>
        public List<SimpleTokenDefinition> Tokens
        {
            get { return _tokens; }
            private set
            {
                _tokens = value;
            }
        }

        /// <summary>
        /// adds token definition
        /// </summary>
        /// <param name="tokenDefinition">A TokenDefinition object</param>
        public void AddToken(SimpleTokenDefinition tokenDefinition)
        {

            _tokens.Add(tokenDefinition);
            // ORDER IS IMPORTANT!
            var sortedTokens = from t in _tokens
                               orderby t.GetTokenLength() descending
                               select t;

            _tokens = sortedTokens.ToList();
            BuildTokenCache();
        }

        /// <summary>
        /// Parses the string
        /// </summary>
        /// <param name="input">input string to parse</param>
        /// <returns>Returns parsed string</returns>
        public string ParseString(string input)
        {
            return ParseString(input, null);
        }

        static readonly Regex ReGuid = new Regex("(?<guid>\\{\\S{8}-\\S{4}-\\S{4}-\\S{4}-\\S{12}?\\})", RegexOptions.Compiled);
        /// <summary>
        /// Gets left over tokens
        /// </summary>
        /// <param name="input">input string</param>
        /// <returns>Returns collections of left over tokens</returns>
        public IEnumerable<string> GetLeftOverTokens(string input)
        {
            List<string> values = new List<string>();
            var matches = ReGuid.Matches(input).OfType<Match>().Select(m => m.Value);
            foreach (var match in matches)
            {
                Guid gout;
                if (!Guid.TryParse(match, out gout))
                {
                    values.Add(match);
                }
            }
            return values;
        }


        private void BuildTokenCache()
        {
            foreach (var tokenDefinition in _tokens)
            {
                foreach (string token in tokenDefinition.GetTokens())
                {
                    var tokenKey = Regex.Unescape(token);
                    if (TokenDictionary.ContainsKey(tokenKey)) continue;

                    string value = tokenDefinition.GetReplaceValue();

                    TokenDictionary[tokenKey] = value;
                }
            }
        }

        private static readonly Regex ReToken = new Regex(@"(?:(\{(?:\1??[^{]*?\})))", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ReTokenFallback = new Regex(@"\{.*?\}", RegexOptions.Compiled);

        private static readonly char[] TokenChars = { '{', '~' };
        private readonly Dictionary<string, string> TokenDictionary = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Parses given string
        /// </summary>
        /// <param name="input">input string</param>
        /// <param name="tokensToSkip">array of tokens to skip</param>
        /// <returns>Returns parsed string</returns>
        public string ParseString(string input, params string[] tokensToSkip)
        {
            if (string.IsNullOrWhiteSpace(input)) return input;

            if (string.IsNullOrEmpty(input) || input.IndexOfAny(TokenChars) == -1) return input;

            BuildTokenCache();

            // Optimize for direct match with string search
            if (TokenDictionary.TryGetValue(input, out string directMatch))
            {
                return directMatch;
            }

            string output = input;
            bool hasMatch = false;

            do
            {
                hasMatch = false;
                output = ReToken.Replace(output, match =>
                {
                    string tokenString = match.Groups[0].Value;
                    if (TokenDictionary.TryGetValue(tokenString, out string val))
                    {
                        hasMatch = true;
                        return val;
                    }
                    return match.Groups[0].Value;
                });
            } while (hasMatch && input != output);

            if (hasMatch) return output;

            var fallbackMatches = ReTokenFallback.Matches(output);
            if (fallbackMatches.Count == 0) return output;

            // If all token constructs {...} are GUID's, we can skip the expensive fallback
            bool needFallback = false;
            foreach (Match match in fallbackMatches)
            {
                if (!ReGuid.IsMatch(match.Value)) needFallback = true;
            }

            if (!needFallback) return output;
            // Fallback for tokens which may contain { or } as part of their name
            foreach (var pair in TokenDictionary)
            {
                int idx = output.IndexOf(pair.Key, StringComparison.CurrentCultureIgnoreCase);
                if (idx != -1)
                {
                    output = output.Remove(idx, pair.Key.Length).Insert(idx, pair.Value);
                }
                if (!ReTokenFallback.IsMatch(output)) break;
            }
            return output;
        }

        internal void RemoveToken<T>(T oldToken) where T : TokenDefinition
        {
            for (int i = 0; i < _tokens.Count; i++)
            {
                var tokenDefinition = _tokens[i];
                if (tokenDefinition.GetTokens().SequenceEqual(oldToken.GetTokens()))
                {
                    _tokens.RemoveAt(i);

                    foreach (string token in tokenDefinition.GetTokens())
                    {
                        var tokenKey = Regex.Unescape(token);
                        TokenDictionary.Remove(tokenKey);
                    }

                    break;
                }
            }
        }
    }
}

