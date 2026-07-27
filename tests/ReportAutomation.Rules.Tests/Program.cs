using System;

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
    }
}
