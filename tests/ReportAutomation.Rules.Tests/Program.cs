using System;
using System.Collections.Generic;

namespace ReportAutomation.Vsto
{
    internal static class Program
    {
        private static int Main()
        {
            AssertEqual("her", GenderRules.ReplacementFor("he"), "he");
            AssertEqual("Him", GenderRules.ReplacementFor("She"), "She");
            AssertEqual("HIMSELF", GenderRules.ReplacementFor("HERSELF"), "HERSELF");
            AssertEqual("woman", GenderRules.ReplacementFor("man"), "man");
            AssertEqual("WOMEN", GenderRules.ReplacementFor("MEN"), "MEN");
            AssertEqual("human", GenderRules.ReplacementFor("human"), "unsupported word");
            AssertMatches("He met a woman and himself.", "He", "woman", "himself");
            AssertMatches("human womanhood amen", new string[0]);
            Console.WriteLine("All GenderRules tests passed.");
            return 0;
        }

        private static void AssertEqual(string expected, string actual, string scenario)
        {
            if (!string.Equals(expected, actual, StringComparison.Ordinal))
            {
                throw new InvalidOperationException(
                    string.Format("{0}: expected '{1}', received '{2}'.", scenario, expected, actual));
            }
        }

        private static void AssertMatches(string text, params string[] expected)
        {
            var actual = new List<string>();
            foreach (GenderTextMatch match in GenderRules.FindMatches(text))
            {
                actual.Add(match.Value);
            }

            AssertEqual(string.Join("|", expected), string.Join("|", actual), "matched words");
        }
    }
}
