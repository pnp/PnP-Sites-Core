using System;
using System.Collections.Generic;
using System.Globalization;

namespace OfficeDevPnP.Core.Pages
{
    /// <summary>
    /// Class that defines the languages for which a translation must be generated
    /// </summary>
    public sealed class TranslationStatusCreationRequest
    {
        /// <summary>
        /// List of languages to generate a translation for
        /// </summary>
        public List<String> LanguageCodes { get; private set; }

        /// <summary>
        /// Add a new language to the list of langauges to be generated. Note that this language must be a language configured for multi-lingual pages on the site
        /// </summary>
        /// <param name="culture"><see cref="CultureInfo"/> object defining the language to add</param>
        public void AddLanguage(CultureInfo culture)
        {
            if (culture == null)
            {
                throw new ArgumentNullException("culture");
            }

            if (LanguageCodes == null)
            {
                LanguageCodes = new List<string>();
            }

            string code = culture.Name.ToLowerInvariant();

            if (!LanguageCodes.Contains(code))
            {
                LanguageCodes.Add(code);
            }
        }

    }
}
