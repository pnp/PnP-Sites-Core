using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Utilities.Themes.Palettes;

namespace OfficeDevPnP.Core.Utilities.Themes.Generator
{

    /* The most minimal set of slots we start with. All other ones can be generated based on rules.
     * This is not so much an enum as it is a list. The enum is used to insure "type"-safety.
     * For now, we are only dealing with color. */
    public enum BaseSlots
    {
        primaryColor,
        backgroundColor,
        foregroundColor
    }

    /* The original Fabric palette, only for back-compat. */
    public enum FabricSlots
    {
        themePrimary, // BaseSlots.primaryColor, Shade[Shade.Unshaded]);
        themeLighterAlt, // BaseSlots.primaryColor, Shade[Shade.Shade1]);
        themeLighter, // BaseSlots.primaryColor, Shade[Shade.Shade2]);
        themeLight, // BaseSlots.primaryColor, Shade[Shade.Shade3]);
        themeTertiary, // BaseSlots.primaryColor, Shade[Shade.Shade4]);
        themeSecondary, // BaseSlots.primaryColor, Shade[Shade.Shade5]);
        themeDarkAlt, // BaseSlots.primaryColor, Shade[Shade.Shade6]);
        themeDark, // BaseSlots.primaryColor, Shade[Shade.Shade7]);
        themeDarker, // BaseSlots.primaryColor, Shade[Shade.Shade8]);

        neutralLighterAlt, // BaseSlots.backgroundColor, Shade[Shade.Shade1]);
        neutralLighter, // BaseSlots.backgroundColor, Shade[Shade.Shade2]);
        neutralLight, // BaseSlots.backgroundColor, Shade[Shade.Shade3]);
        neutralQuaternaryAlt, // BaseSlots.backgroundColor, Shade[Shade.Shade4]);
        neutralQuaternary, // BaseSlots.backgroundColor, Shade[Shade.Shade5]);
        neutralTertiaryAlt, // BaseSlots.backgroundColor, Shade[Shade.Shade6]); // bg6 or fg2
        neutralTertiary, // BaseSlots.foregroundColor, Shade[Shade.Shade3]);
                         // deprecated: neutralSecondaryAlt, // BaseSlots.foregroundColor, Shade[Shade.Shade4]);
        neutralSecondary, // BaseSlots.foregroundColor, Shade[Shade.Shade5]);
        neutralPrimaryAlt, // BaseSlots.foregroundColor, Shade[Shade.Shade6]);
        neutralPrimary, // BaseSlots.foregroundColor, Shade[Shade.Unshaded]);
        neutralDark, // BaseSlots.foregroundColor, Shade[Shade.Shade7]);

        black, // BaseSlots.foregroundColor, Shade[Shade.Shade8]);
        white // BaseSlots.backgroundColor, Shade[Shade.Unshaded]);
    }

    /* List of all the semantic color slots for this theme.
     * This is not so much an enum as it is a list. The enum is used to insure "type"-safety. */
    public enum SemanticColorSlots
    {
        bodyBackground,
        bodyText,
        disabledBackground,
        disabledText
    }

    public class ThemeRulesStandard : ThemeRules
    {
        public ThemeRulesStandard()
        {
            /*** BASE COLORS and their SHADES */
            // iterate through each base slot and make the SlotRules for those
            var baseSlotsNames = Enum.GetNames(typeof(BaseSlots));

            foreach (var baseSlot in baseSlotsNames)
            {
                this[baseSlot] = new ThemeSlotRule()
                {
                    name = baseSlot,
                    isCustomized = true,
                    dependentRules = new List<IThemeSlotRule>()
                };

                var shadeNames = Enum.GetNames(typeof(Shade));

                foreach (var shadeName in shadeNames.Where(s => s != Shade.Unshaded.ToString()))
                {
                    var slotRuleName = baseSlot + shadeName;

                    var thisSlotRule = new ThemeSlotRule()
                    {
                        name = slotRuleName,
                        inherits = this[baseSlot],
                        asShade = (Shade)Enum.Parse(typeof(Shade), shadeName),
                        isCustomized = false,
                        isBackgroundShade = (baseSlot == BaseSlots.backgroundColor.ToString()) ? true : false,
                        dependentRules = new List<IThemeSlotRule>()
                    };

                    this[slotRuleName] = thisSlotRule;
                    this[baseSlot].dependentRules.Add(thisSlotRule);
                }
            }

            // set default colors for the base colors
            this[BaseSlots.primaryColor.ToString()].color = Colors.getColorFromString("#0078d4");
            this[BaseSlots.backgroundColor.ToString()].color = Colors.getColorFromString("#fff");
            this[BaseSlots.foregroundColor.ToString()].color = Colors.getColorFromString("#333");

            // set default colors for shades (the slot rules were already created above and will be used if the base colors ever change)
            this[BaseSlots.primaryColor.ToString() + Shade.Shade1.ToString()].color = Colors.getColorFromString("#eff6fc");
            this[BaseSlots.primaryColor.ToString() + Shade.Shade2.ToString()].color = Colors.getColorFromString("#deecf9");
            this[BaseSlots.primaryColor.ToString() + Shade.Shade3.ToString()].color = Colors.getColorFromString("#c7e0f4");
            this[BaseSlots.primaryColor.ToString() + Shade.Shade4.ToString()].color = Colors.getColorFromString("#71afe5");
            this[BaseSlots.primaryColor.ToString() + Shade.Shade5.ToString()].color = Colors.getColorFromString("#2b88d8");
            this[BaseSlots.primaryColor.ToString() + Shade.Shade6.ToString()].color = Colors.getColorFromString("#106ebe");
            this[BaseSlots.primaryColor.ToString() + Shade.Shade7.ToString()].color = Colors.getColorFromString("#005a9e");
            this[BaseSlots.primaryColor.ToString() + Shade.Shade8.ToString()].color = Colors.getColorFromString("#004578");

            // set default colors for shades (the slot rules were already created above and will be used if the base colors ever change)
            this[BaseSlots.foregroundColor.ToString() + Shade.Shade1.ToString()].color = Colors.getColorFromString("#eaeaea");
            this[BaseSlots.foregroundColor.ToString() + Shade.Shade2.ToString()].color = Colors.getColorFromString("#c8c8c8");
            this[BaseSlots.foregroundColor.ToString() + Shade.Shade3.ToString()].color = Colors.getColorFromString("#a6a6a6");
            this[BaseSlots.foregroundColor.ToString() + Shade.Shade4.ToString()].color = Colors.getColorFromString("#767676");
            this[BaseSlots.foregroundColor.ToString() + Shade.Shade5.ToString()].color = Colors.getColorFromString("#666666");
            this[BaseSlots.foregroundColor.ToString() + Shade.Shade6.ToString()].color = Colors.getColorFromString("#3c3c3c");
            this[BaseSlots.foregroundColor.ToString() + Shade.Shade7.ToString()].color = Colors.getColorFromString("#212121");
            this[BaseSlots.foregroundColor.ToString() + Shade.Shade8.ToString()].color = Colors.getColorFromString("#000000");


            _makeFabricSlotRule(FabricSlots.themePrimary.ToString(), BaseSlots.primaryColor, Shade.Unshaded);
            _makeFabricSlotRule(FabricSlots.themeLighterAlt.ToString(), BaseSlots.primaryColor, Shade.Shade1);
            _makeFabricSlotRule(FabricSlots.themeLighter.ToString(), BaseSlots.primaryColor, Shade.Shade2);
            _makeFabricSlotRule(FabricSlots.themeLight.ToString(), BaseSlots.primaryColor, Shade.Shade3);
            _makeFabricSlotRule(FabricSlots.themeTertiary.ToString(), BaseSlots.primaryColor, Shade.Shade4);
            _makeFabricSlotRule(FabricSlots.themeSecondary.ToString(), BaseSlots.primaryColor, Shade.Shade5);
            _makeFabricSlotRule(FabricSlots.themeDarkAlt.ToString(), BaseSlots.primaryColor, Shade.Shade6);
            _makeFabricSlotRule(FabricSlots.themeDark.ToString(), BaseSlots.primaryColor, Shade.Shade7);
            _makeFabricSlotRule(FabricSlots.themeDarker.ToString(), BaseSlots.primaryColor, Shade.Shade8);

            _makeFabricSlotRule(FabricSlots.neutralLighterAlt.ToString(), BaseSlots.backgroundColor, Shade.Shade1, true);
            _makeFabricSlotRule(FabricSlots.neutralLighter.ToString(), BaseSlots.backgroundColor, Shade.Shade2, true);
            _makeFabricSlotRule(FabricSlots.neutralLight.ToString(), BaseSlots.backgroundColor, Shade.Shade3, true);
            _makeFabricSlotRule(FabricSlots.neutralQuaternaryAlt.ToString(), BaseSlots.backgroundColor, Shade.Shade4, true);
            _makeFabricSlotRule(FabricSlots.neutralQuaternary.ToString(), BaseSlots.backgroundColor, Shade.Shade5, true);
            _makeFabricSlotRule(FabricSlots.neutralTertiaryAlt.ToString(), BaseSlots.backgroundColor, Shade.Shade6, true); // bg6 or fg2
            _makeFabricSlotRule(FabricSlots.neutralTertiary.ToString(), BaseSlots.foregroundColor, Shade.Shade3);
            _makeFabricSlotRule(FabricSlots.neutralSecondary.ToString(), BaseSlots.foregroundColor, Shade.Shade4);
            _makeFabricSlotRule(FabricSlots.neutralPrimaryAlt.ToString(), BaseSlots.foregroundColor, Shade.Shade5);
            _makeFabricSlotRule(FabricSlots.neutralPrimary.ToString(), BaseSlots.foregroundColor, Shade.Unshaded);
            _makeFabricSlotRule(FabricSlots.neutralDark.ToString(), BaseSlots.foregroundColor, Shade.Shade7);

            _makeFabricSlotRule(FabricSlots.black.ToString(), BaseSlots.foregroundColor, Shade.Shade8);
            _makeFabricSlotRule(FabricSlots.white.ToString(), BaseSlots.backgroundColor, Shade.Unshaded, true);

            // manually set initial colors for the primary-based Fabric slots to match the default theme
            this[FabricSlots.themeLighterAlt.ToString()].color = Colors.getColorFromString("#eff6fc");
            this[FabricSlots.themeLighter.ToString()].color = Colors.getColorFromString("#deecf9");
            this[FabricSlots.themeLight.ToString()].color = Colors.getColorFromString("#c7e0f4");
            this[FabricSlots.themeTertiary.ToString()].color = Colors.getColorFromString("#71afe5");
            this[FabricSlots.themeSecondary.ToString()].color = Colors.getColorFromString("#2b88d8");
            this[FabricSlots.themeDarkAlt.ToString()].color = Colors.getColorFromString("#106ebe");
            this[FabricSlots.themeDark.ToString()].color = Colors.getColorFromString("#005a9e");
            this[FabricSlots.themeDarker.ToString()].color = Colors.getColorFromString("#004578");
            this[FabricSlots.themeLighterAlt.ToString()].isCustomized = true;
            this[FabricSlots.themeLighter.ToString()].isCustomized = true;
            this[FabricSlots.themeLight.ToString()].isCustomized = true;
            this[FabricSlots.themeTertiary.ToString()].isCustomized = true;
            this[FabricSlots.themeSecondary.ToString()].isCustomized = true;
            this[FabricSlots.themeDarkAlt.ToString()].isCustomized = true;
            this[FabricSlots.themeDark.ToString()].isCustomized = true;
            this[FabricSlots.themeDarker.ToString()].isCustomized = true;

            // Basic simple slots
            _makeSemanticSlotRule(SemanticColorSlots.bodyBackground, FabricSlots.white);
            _makeSemanticSlotRule(SemanticColorSlots.bodyText, FabricSlots.neutralPrimary);
        }

        private void _makeSemanticSlotRule(SemanticColorSlots semanticSlot, FabricSlots inheritedFabricSlot)
        {
            var inherits = this[inheritedFabricSlot.ToString()];
            var thisSlotRule = new ThemeSlotRule
            {
                name = semanticSlot.ToString(),
                inherits = this[inheritedFabricSlot.ToString()],
                isCustomized = false,
                dependentRules = new List<IThemeSlotRule>()
            };
            this[semanticSlot.ToString()] = thisSlotRule;
            inherits.dependentRules.Add(thisSlotRule);
        }

        private void _makeFabricSlotRule(string slotName, BaseSlots inheritedBase, Shade inheritedShade, bool isBackgroundShade = false)
        {
            var inherits = this[inheritedBase.ToString()];
            var thisSlotRule = new ThemeSlotRule
            {
                name = slotName,
                inherits = inherits,
                asShade = inheritedShade,
                isCustomized = false,
                isBackgroundShade = isBackgroundShade,
                dependentRules = new List<IThemeSlotRule>()
            };
            this[slotName] = thisSlotRule;
            inherits.dependentRules.Add(thisSlotRule);
        }
    }
}
