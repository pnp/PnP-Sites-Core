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
        /// <param name="languageId">Id defining the language to add</param>
        public void AddLanguage(int languageId)
        {
            if (languageId == 0)
            {
                throw new ArgumentException("culture");
            }

            CultureInfo culture = new CultureInfo(languageId);
            string code = culture.Name.ToLowerInvariant();

            if (LanguageCodes == null)
            {
                LanguageCodes = new List<string>();
            }

            if (!LanguageCodes.Contains(code))
            {
                LanguageCodes.Add(code);
            }
        }

    }
}
