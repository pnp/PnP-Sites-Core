using OfficeDevPnP.Core.Utilities.Themes.Palettes;
using OfficeDevPnP.Core.Utilities.Themes.Generator;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Utilities.Themes
{
    public static class ThemeUtility
    {
        /// <summary>
        /// Apply three custom colors to the standard Office UI Fabric template and get back the JSON with the CSS rules
        /// </summary>
        /// <param name="primaryColor">Primary Color for the Theme</param>
        /// <param name="bodyTextColor">Body Text Color for the Theme</param>
        /// <param name="bodyBackgroundColor">Body Background Color for the Theme</param>
        /// <returns>THe JSON representation of the Theme</returns>
        public static String GetThemeAsJSON(String primaryColor, String bodyTextColor, String bodyBackgroundColor)
        {
            // Validate inputs
            if (String.IsNullOrEmpty(primaryColor))
                throw new ArgumentNullException(nameof(primaryColor));

            if (String.IsNullOrEmpty(bodyTextColor))
                throw new ArgumentNullException(nameof(bodyTextColor));

            if (String.IsNullOrEmpty(bodyBackgroundColor))
                throw new ArgumentNullException(nameof(bodyBackgroundColor));

            // Parse the custom colors
            var primaryCustomColor = Colors.getColorFromString(primaryColor);
            var bodyTextCustomColor = Colors.getColorFromString(bodyTextColor);
            var bodyBackgroundCustomColor = Colors.getColorFromString(bodyBackgroundColor);

            // Get the standard template
            var standardRules = new ThemeRulesStandard();
            ThemeGenerator.InsureSlots(standardRules, false);

            // Set the custom colors to the template
            ThemeGenerator.SetSlot(standardRules[BaseSlots.primaryColor.ToString()], primaryCustomColor);
            ThemeGenerator.SetSlot(standardRules[BaseSlots.foregroundColor.ToString()], bodyTextCustomColor);
            ThemeGenerator.SetSlot(standardRules[BaseSlots.backgroundColor.ToString()], bodyBackgroundCustomColor);

            // Get the JSON string
            String json = ThemeGenerator.GetThemeAsJson(standardRules);

            return json;
        }

        /// <summary>
        /// Return an out of the box modern theme as a JSON string
        /// </summary>
        /// <param name="theme">Name of the out of the box theme</param>
        /// <returns></returns>
        public static String GetOOBModernThemeJson(OOBTheme theme)
        {
            switch (theme)
            {
                case OOBTheme.Blue:
                    return "{\"themePrimary\":{\"R\":0,\"G\":120,\"B\":212,\"A\":255},\"themeLighterAlt\":{\"R\":239,\"G\":246,\"B\":252,\"A\":255},\"themeLighter\":{\"R\":222,\"G\":236,\"B\":249,\"A\":255},\"themeLight\":{\"R\":199,\"G\":224,\"B\":244,\"A\":255},\"themeTertiary\":{\"R\":113,\"G\":175,\"B\":229,\"A\":255},\"themeSecondary\":{\"R\":43,\"G\":136,\"B\":216,\"A\":255},\"themeDarkAlt\":{\"R\":16,\"G\":110,\"B\":190,\"A\":255},\"themeDark\":{\"R\":0,\"G\":90,\"B\":158,\"A\":255},\"themeDarker\":{\"R\":0,\"G\":69,\"B\":120,\"A\":255},\"accent\":{\"R\":135,\"G\":100,\"B\":184,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}}";
                case OOBTheme.Orange:
                    return "{\"themePrimary\":{\"R\":202,\"G\":80,\"B\":16,\"A\":255},\"themeLighterAlt\":{\"R\":253,\"G\":247,\"B\":244,\"A\":255},\"themeLighter\":{\"R\":246,\"G\":223,\"B\":210,\"A\":255},\"themeLight\":{\"R\":239,\"G\":196,\"B\":173,\"A\":255},\"themeTertiary\":{\"R\":223,\"G\":143,\"B\":100,\"A\":255},\"themeSecondary\":{\"R\":208,\"G\":98,\"B\":40,\"A\":255},\"themeDarkAlt\":{\"R\":181,\"G\":73,\"B\":15,\"A\":255},\"themeDark\":{\"R\":153,\"G\":62,\"B\":12,\"A\":255},\"themeDarker\":{\"R\":113,\"G\":45,\"B\":9,\"A\":255},\"accent\":{\"R\":152,\"G\":111,\"B\":11,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}}";
                case OOBTheme.Red:
                    return "{\"themePrimary\":{\"R\":164,\"G\":38,\"B\":44,\"A\":255},\"themeLighterAlt\":{\"R\":251,\"G\":244,\"B\":244,\"A\":255},\"themeLighter\":{\"R\":240,\"G\":211,\"B\":212,\"A\":255},\"themeLight\":{\"R\":227,\"G\":175,\"B\":178,\"A\":255},\"themeTertiary\":{\"R\":200,\"G\":108,\"B\":112,\"A\":255},\"themeSecondary\":{\"R\":174,\"G\":56,\"B\":62,\"A\":255},\"themeDarkAlt\":{\"R\":147,\"G\":34,\"B\":39,\"A\":255},\"themeDark\":{\"R\":124,\"G\":29,\"B\":33,\"A\":255},\"themeDarker\":{\"R\":91,\"G\":21,\"B\":25,\"A\":255},\"accent\":{\"R\":202,\"G\":80,\"B\":16,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}}";
                case OOBTheme.Purple:
                    return "{\"themePrimary\":{\"R\":135,\"G\":100,\"B\":184,\"A\":255},\"themeLighterAlt\":{\"R\":249,\"G\":248,\"B\":252,\"A\":255},\"themeLighter\":{\"R\":233,\"G\":226,\"B\":244,\"A\":255},\"themeLight\":{\"R\":215,\"G\":201,\"B\":234,\"A\":255},\"themeTertiary\":{\"R\":178,\"G\":154,\"B\":212,\"A\":255},\"themeSecondary\":{\"R\":147,\"G\":114,\"B\":192,\"A\":255},\"themeDarkAlt\":{\"R\":121,\"G\":89,\"B\":165,\"A\":255},\"themeDark\":{\"R\":102,\"G\":75,\"B\":140,\"A\":255},\"themeDarker\":{\"R\":75,\"G\":56,\"B\":103,\"A\":255},\"accent\":{\"R\":3,\"G\":131,\"B\":135,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}}";
                case OOBTheme.Green:
                    return "{\"themePrimary\":{\"R\":73,\"G\":130,\"B\":5,\"A\":255},\"themeLighterAlt\":{\"R\":246,\"G\":250,\"B\":240,\"A\":255},\"themeLighter\":{\"R\":219,\"G\":235,\"B\":199,\"A\":255},\"themeLight\":{\"R\":189,\"G\":218,\"B\":155,\"A\":255},\"themeTertiary\":{\"R\":133,\"G\":180,\"B\":76,\"A\":255},\"themeSecondary\":{\"R\":90,\"G\":145,\"B\":23,\"A\":255},\"themeDarkAlt\":{\"R\":66,\"G\":117,\"B\":5,\"A\":255},\"themeDark\":{\"R\":56,\"G\":99,\"B\":4,\"A\":255},\"themeDarker\":{\"R\":41,\"G\":73,\"B\":3,\"A\":255},\"accent\":{\"R\":3,\"G\":131,\"B\":135,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}}";
                case OOBTheme.Gray:
                    return "{\"themePrimary\":{\"R\":105,\"G\":121,\"B\":126,\"A\":255},\"themeLighterAlt\":{\"R\":248,\"G\":249,\"B\":250,\"A\":255},\"themeLighter\":{\"R\":228,\"G\":233,\"B\":234,\"A\":255},\"themeLight\":{\"R\":205,\"G\":213,\"B\":216,\"A\":255},\"themeTertiary\":{\"R\":159,\"G\":173,\"B\":177,\"A\":255},\"themeSecondary\":{\"R\":120,\"G\":136,\"B\":141,\"A\":255},\"themeDarkAlt\":{\"R\":93,\"G\":108,\"B\":112,\"A\":255},\"themeDark\":{\"R\":79,\"G\":91,\"B\":95,\"A\":255},\"themeDarker\":{\"R\":58,\"G\":67,\"B\":70,\"A\":255},\"accent\":{\"R\":0,\"G\":120,\"B\":212,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}}";
                case OOBTheme.DarkYellow:
                    return "{\"themePrimary\":{\"R\":255,\"G\":200,\"B\":61,\"A\":255},\"themeLighterAlt\":{\"R\":10,\"G\":8,\"B\":2,\"A\":255},\"themeLighter\":{\"R\":41,\"G\":32,\"B\":10,\"A\":255},\"themeLight\":{\"R\":77,\"G\":60,\"B\":18,\"A\":255},\"themeTertiary\":{\"R\":153,\"G\":120,\"B\":37,\"A\":255},\"themeSecondary\":{\"R\":224,\"G\":176,\"B\":54,\"A\":255},\"themeDarkAlt\":{\"R\":255,\"G\":206,\"B\":81,\"A\":255},\"themeDark\":{\"R\":255,\"G\":213,\"B\":108,\"A\":255},\"themeDarker\":{\"R\":255,\"G\":224,\"B\":146,\"A\":255},\"accent\":{\"R\":255,\"G\":200,\"B\":61,\"A\":255},\"neutralLighterAlt\":{\"R\":40,\"G\":40,\"B\":40,\"A\":255},\"neutralLighter\":{\"R\":49,\"G\":49,\"B\":49,\"A\":255},\"neutralLight\":{\"R\":63,\"G\":63,\"B\":63,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":72,\"G\":72,\"B\":72,\"A\":255},\"neutralQuaternary\":{\"R\":79,\"G\":79,\"B\":79,\"A\":255},\"neutralTertiaryAlt\":{\"R\":109,\"G\":109,\"B\":109,\"A\":255},\"neutralTertiary\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralSecondary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralPrimaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralPrimary\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"neutralDark\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"black\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"white\":{\"R\":31,\"G\":31,\"B\":31,\"A\":255},\"primaryBackground\":{\"R\":31,\"G\":31,\"B\":31,\"A\":255},\"primaryText\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255}}";
                case OOBTheme.DarkBlue:
                    return "{\"themePrimary\":{\"R\":58,\"G\":150,\"B\":221,\"A\":255},\"themeLighterAlt\":{\"R\":2,\"G\":6,\"B\":9,\"A\":255},\"themeLighter\":{\"R\":9,\"G\":24,\"B\":35,\"A\":255},\"themeLight\":{\"R\":17,\"G\":45,\"B\":67,\"A\":255},\"themeTertiary\":{\"R\":35,\"G\":90,\"B\":133,\"A\":255},\"themeSecondary\":{\"R\":51,\"G\":133,\"B\":195,\"A\":255},\"themeDarkAlt\":{\"R\":75,\"G\":160,\"B\":225,\"A\":255},\"themeDark\":{\"R\":101,\"G\":174,\"B\":230,\"A\":255},\"themeDarker\":{\"R\":138,\"G\":194,\"B\":236,\"A\":255},\"accent\":{\"R\":58,\"G\":150,\"B\":221,\"A\":255},\"neutralLighterAlt\":{\"R\":29,\"G\":43,\"B\":60,\"A\":255},\"neutralLighter\":{\"R\":34,\"G\":50,\"B\":68,\"A\":255},\"neutralLight\":{\"R\":43,\"G\":61,\"B\":81,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":50,\"G\":68,\"B\":89,\"A\":255},\"neutralQuaternary\":{\"R\":55,\"G\":74,\"B\":95,\"A\":255},\"neutralTertiaryAlt\":{\"R\":79,\"G\":99,\"B\":122,\"A\":255},\"neutralTertiary\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralSecondary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralPrimaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralPrimary\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"neutralDark\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"black\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"white\":{\"R\":24,\"G\":37,\"B\":52,\"A\":255},\"primaryBackground\":{\"R\":24,\"G\":37,\"B\":52,\"A\":255},\"primaryText\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255}}";
                default:
                    return "{\"themePrimary\":{\"R\":202,\"G\":80,\"B\":16,\"A\":255},\"themeLighterAlt\":{\"R\":253,\"G\":247,\"B\":244,\"A\":255},\"themeLighter\":{\"R\":246,\"G\":223,\"B\":210,\"A\":255},\"themeLight\":{\"R\":239,\"G\":196,\"B\":173,\"A\":255},\"themeTertiary\":{\"R\":223,\"G\":143,\"B\":100,\"A\":255},\"themeSecondary\":{\"R\":208,\"G\":98,\"B\":40,\"A\":255},\"themeDarkAlt\":{\"R\":181,\"G\":73,\"B\":15,\"A\":255},\"themeDark\":{\"R\":153,\"G\":62,\"B\":12,\"A\":255},\"themeDarker\":{\"R\":113,\"G\":45,\"B\":9,\"A\":255},\"accent\":{\"R\":152,\"G\":111,\"B\":11,\"A\":255},\"neutralLighterAlt\":{\"R\":248,\"G\":248,\"B\":248,\"A\":255},\"neutralLighter\":{\"R\":244,\"G\":244,\"B\":244,\"A\":255},\"neutralLight\":{\"R\":234,\"G\":234,\"B\":234,\"A\":255},\"neutralQuaternaryAlt\":{\"R\":218,\"G\":218,\"B\":218,\"A\":255},\"neutralQuaternary\":{\"R\":208,\"G\":208,\"B\":208,\"A\":255},\"neutralTertiaryAlt\":{\"R\":200,\"G\":200,\"B\":200,\"A\":255},\"neutralTertiary\":{\"R\":166,\"G\":166,\"B\":166,\"A\":255},\"neutralSecondary\":{\"R\":102,\"G\":102,\"B\":102,\"A\":255},\"neutralPrimaryAlt\":{\"R\":60,\"G\":60,\"B\":60,\"A\":255},\"neutralPrimary\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255},\"neutralDark\":{\"R\":33,\"G\":33,\"B\":33,\"A\":255},\"black\":{\"R\":0,\"G\":0,\"B\":0,\"A\":255},\"white\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryBackground\":{\"R\":255,\"G\":255,\"B\":255,\"A\":255},\"primaryText\":{\"R\":51,\"G\":51,\"B\":51,\"A\":255}}";
            }
        }
    }

}
