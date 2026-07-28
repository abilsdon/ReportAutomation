using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace ReportAutomation.Vsto
{
    internal sealed class GenderTextMatch
    {
        internal int Index { get; set; }
        internal int Length { get; set; }
        internal string Value { get; set; }
    }

    internal static class GenderRules
    {
        private static readonly Regex SupportedWordPattern = new Regex(
            @"\b(himself|herself|woman|women|hers|his|she|him|her|men|man|he)\b",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        internal static readonly IReadOnlyDictionary<string, string> Replacements =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "he", "her" }, { "she", "him" }, { "him", "her" },
                { "her", "him" }, { "his", "hers" }, { "hers", "his" },
                { "himself", "herself" }, { "herself", "himself" },
                { "man", "woman" }, { "woman", "man" },
                { "men", "women" }, { "women", "men" }
            };

        internal static IEnumerable<GenderTextMatch> FindMatches(string text)
        {
            if (string.IsNullOrEmpty(text)) yield break;

            foreach (Match match in SupportedWordPattern.Matches(text))
            {
                yield return new GenderTextMatch
                {
                    Index = match.Index,
                    Length = match.Length,
                    Value = match.Value
                };
            }
        }

        internal static string ReplacementFor(string source)
        {
            string target;
            if (string.IsNullOrEmpty(source) || !Replacements.TryGetValue(source, out target))
            {
                return source;
            }

            if (source == source.ToUpperInvariant()) return target.ToUpperInvariant();
            if (char.IsUpper(source[0]))
            {
                return char.ToUpperInvariant(target[0]) + target.Substring(1);
            }

            return target;
        }
    }
}
