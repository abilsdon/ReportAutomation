using System;
using System.Collections.Generic;

namespace ReportAutomation.Vsto
{
    internal static class GenderRules
    {
        internal static readonly IReadOnlyDictionary<string, string> Replacements =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "he", "her" },
                { "she", "him" },
                { "him", "her" },
                { "her", "him" },
                { "his", "hers" },
                { "hers", "his" },
                { "himself", "herself" },
                { "herself", "himself" },
                { "man", "woman" },
                { "woman", "man" },
                { "men", "women" },
                { "women", "men" }
            };

        internal static string ReplacementFor(string source)
        {
            string target;
            if (string.IsNullOrEmpty(source) || !Replacements.TryGetValue(source, out target))
            {
                return source;
            }

            if (source == source.ToUpperInvariant())
            {
                return target.ToUpperInvariant();
            }

            if (char.IsUpper(source[0]))
            {
                return char.ToUpperInvariant(target[0]) + target.Substring(1);
            }

            return target;
        }
    }
}
