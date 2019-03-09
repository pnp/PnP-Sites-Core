using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeDevPnP.Core.Utilities.Themes.Palettes;

namespace OfficeDevPnP.Core.Utilities.Themes.Generator
{
    /// <summary>
    /// Sets an IThemeSlotRule to the given color and cascades it to the rest of the theme, updating other IThemeSlotRules in the theme that
    /// inherit from that color.
    /// isInverted: whether it's a dark theme or not, which affects the algorithm used to generate shades
    /// isCustomization should be true only if it's a user action, and indicates overwriting the slot's inheritance(if any)
    /// overwriteCustomColor: a slot could have a generated color based on its inheritance rules(isCustomized is false), or a custom color
    /// based on user input(isCustomized is true), this bool tells us whether to override existing customized colors
    /// </summary>
    public class ThemeGenerator
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="rule"></param>
        /// <param name="color"></param>
        /// <param name="isInverted"></param>
        /// <param name="isCustomization"></param>
        /// <param name="overwriteCustomColor"></param>
        public static void SetSlot(IThemeSlotRule rule, IColor color, Boolean isInverted = false, Boolean isCustomization = false, Boolean overwriteCustomColor = true)
        {
            if (rule.color == null && String.IsNullOrEmpty(rule.value))
            {
                // not a color rule
                return;
            }

            if (overwriteCustomColor)
            {
                IColor colorAsIColor = color;
                if (colorAsIColor == null)
                {
                    throw new ArgumentNullException(nameof(color), "Color is invalid in setSlot(): " + color.ToString());
                }
                
                ThemeGenerator.SetSlotInternal(rule, colorAsIColor, isInverted, isCustomization, overwriteCustomColor);
            }
            else if (rule.color != null)
            {
                ThemeGenerator.SetSlotInternal(rule, rule.color, isInverted, isCustomization, overwriteCustomColor);
            }
        }

        /// <summary>
        /// Sets the color of each slot based on its rule. Slots that don't inherit must have a color already.
        /// If this completes without error, then the theme is ready to use. (All slots will have a color.)
        /// setSlot() can be called before this, but this must be called before getThemeAs*().
        /// Does not override colors of rules where isCustomized is true (i.e. doesn't override existing customizations).
        /// </summary>
        /// <param name="slotRules"></param>
        /// <param name="isInverted"></param>
        public static void InsureSlots(IThemeRules slotRules, Boolean isInverted)
        {
            // Get all the "root" rules, the ones which don't inherit. Then "set" them to trigger updating dependent slots.
            foreach (var ruleName in slotRules) 
            {
                IThemeSlotRule rule = slotRules[ruleName];
                if (rule.inherits == null && String.IsNullOrEmpty(rule.value))
                {
                    if (rule.color == null)
                    {
                        throw new InvalidOperationException("A color slot rule that does not inherit must provide its own color.");
                    }
                    ThemeGenerator.SetSlotInternal(rule, rule.color, isInverted, false, false);
                }
            }
        }

        /// <summary>
        /// Gets the JSON-formatted blob that describes the theme, usable with the REST request endpoints
        /// { [theme slot name as string] : [color as string],
        ///  "tokenName": "#f00f00",
        ///  "tokenName2": "#ba2ba2",
        ///   ... }
        /// </summary>
        /// <param name="slotRules"></param>
        public static string GetThemeAsJson(IThemeRules slotRules)
        {
            var theme = new Dictionary<string,string>();
            foreach (var ruleName in slotRules)
            {
                // strip out the unnecessary shade slots from the final output theme
                if (ruleName.IndexOf("ColorShade") == -1 && ruleName != BaseSlots.primaryColor.ToString() && ruleName != BaseSlots.backgroundColor.ToString() && ruleName != BaseSlots.foregroundColor.ToString())
                {
                    IThemeSlotRule rule = slotRules[ruleName];
                    theme[rule.name] = (rule.color != null) ? rule.color.str : rule.value;
                }
            }

            var json = Newtonsoft.Json.JsonConvert.SerializeObject(theme);

            return json;
        }

        /// <summary>
        /// Sets the given slot's color to the appropriate color, shading it if necessary.
        /// Then, iterates through all other rules(that are this rule's dependents) to update them accordingly.
        /// isCustomization= true means it's a user provided color, set it to that raw color
        /// isCustomization= false means the rule it's inheriting from changed, so updated using asShade
        /// </summary>
        /// <param name="rule"></param>
        /// <param name="color"></param>
        /// <param name="isInverted"></param>
        /// <param name="isCustomization"></param>
        /// <param name="overwriteCustomColor"></param>
        private static void SetSlotInternal(IThemeSlotRule rule, IColor color, Boolean isInverted, Boolean isCustomization, Boolean overwriteCustomColor = true)
        {
            //if (rule.color == null && String.IsNullOrEmpty(rule.value))
            //{
            //    // not a color rule
            //    return;
            //}

            if (overwriteCustomColor || rule.color == null || !rule.isCustomized || rule.inherits == null)
            {
                // set the rule's color under these conditions
                if ((overwriteCustomColor || !rule.isCustomized) && !isCustomization && rule.inherits != null && Shades.IsValidShade(rule.asShade))
                {
                    // it's inheriting by shade
                    if (rule.isBackgroundShade)
                    {
                        rule.color = Shades.GetBackgroundShade(color, rule.asShade, isInverted);
                    }
                    else
                    {
                        rule.color = Shades.GetShade(color, rule.asShade, isInverted);
                    }
                    rule.isCustomized = false;
                }
                else
                {
                    rule.color = color;
                    rule.isCustomized = true;
                }

                // then update dependent colors
                foreach (var ruleToUpdate in rule.dependentRules)
                {
                    ThemeGenerator.SetSlotInternal(ruleToUpdate, rule.color, isInverted, false, overwriteCustomColor);
                }
            }
        }
    }
}
