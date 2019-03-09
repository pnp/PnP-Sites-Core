using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;
using OfficeDevPnP.Core.Utilities.Themes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OfficeDevPnP.Core.Tests.Utilities
{
    [TestClass]
    public class ThemeUtilityTests
    {
        [TestMethod]
        public void GetThemeAsJSONTest()
        {
            var expectedThemeJson = "{" +
                "  'themePrimary': '#425463'," +
                "  'themeLighterAlt': '#F5F7F9'," +
                "  'themeLighter': '#DAE1E6'," +
                "  'themeLight': '#BCC7D0'," +
                "  'themeTertiary': '#8293A2'," +
                "  'themeSecondary': '#546776'," +
                "  'themeDarkAlt': '#3C4C5A'," +
                "  'themeDark': '#33404C'," +
                "  'themeDarker': '#252F38'," +
                "  'neutralLighterAlt': '#F8CB68'," +
                "  'neutralLighter': '#F4C766'," +
                "  'neutralLight': '#EABF62'," +
                "  'neutralQuaternaryAlt': '#DAB25C'," +
                "  'neutralQuaternary': '#D0AA57'," +
                "  'neutralTertiaryAlt': '#C8A354'," +
                "  'neutralTertiary': '#595959'," +
                "  'neutralSecondary': '#373737'," +
                "  'neutralPrimaryAlt': '#2F2F2F'," +
                "  'neutralPrimary': '#000000'," +
                "  'neutralDark': '#151515'," +
                "  'black': '#0B0B0B'," +
                "  'white': '#ffd06a'," +
                "  'bodyBackground': '#ffd06a'," +
                "  'bodyText': '#000000'" +
            "}";
            var expectedTheme = JsonConvert.DeserializeObject(expectedThemeJson);

            var generatedThemeJson = ThemeUtility.GetThemeAsJSON("#425463", "#000000", "#ffd06a");
            var generatedTheme = JsonConvert.DeserializeObject(generatedThemeJson);

            Assert.AreEqual(
                JsonConvert.SerializeObject(expectedTheme),
                JsonConvert.SerializeObject(generatedTheme));
        }
    }
}
