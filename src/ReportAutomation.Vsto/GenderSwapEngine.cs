using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace ReportAutomation.Vsto
{
    internal sealed class GenderSwapEngine
    {
        private sealed class Match
        {
            internal Word.Range Range { get; set; }
            internal string Original { get; set; }
        }

        internal int CountDocument(Word.Document document)
        {
            return FindDocumentMatches(document).Count;
        }

        internal int CountRange(Word.Range range)
        {
            return FindMatches(new[] { range }).Count;
        }

        internal int SwapDocument(Word.Document document)
        {
            return Apply(FindDocumentMatches(document));
        }

        internal int SwapRange(Word.Range range)
        {
            return Apply(FindMatches(new[] { range }));
        }

        private static List<Match> FindDocumentMatches(Word.Document document)
        {
            var stories = new List<Word.Range>();
            foreach (Word.Range firstStory in document.StoryRanges)
            {
                Word.Range story = firstStory;
                while (story != null)
                {
                    stories.Add(story.Duplicate);
                    story = story.NextStoryRange;
                }
            }

            return FindMatches(stories);
        }

        private static List<Match> FindMatches(IEnumerable<Word.Range> scopes)
        {
            var matches = new List<Match>();
            foreach (Word.Range scope in scopes)
            {
                int scopeEnd = scope.End;
                foreach (string source in GenderRules.Replacements.Keys)
                {
                    Word.Range search = scope.Duplicate;
                    search.Find.ClearFormatting();
                    search.Find.Text = source;
                    search.Find.Forward = true;
                    search.Find.Wrap = Word.WdFindWrap.wdFindStop;
                    search.Find.Format = false;
                    search.Find.MatchCase = false;
                    search.Find.MatchWholeWord = true;

                    while (search.Find.Execute())
                    {
                        matches.Add(new Match { Range = search.Duplicate, Original = search.Text });
                        search.Start = search.End;
                        search.End = scopeEnd;
                    }
                }
            }

            return matches;
        }

        private static int Apply(List<Match> matches)
        {
            for (int index = matches.Count - 1; index >= 0; index--)
            {
                Match match = matches[index];
                match.Range.Text = GenderRules.ReplacementFor(match.Original);
            }

            return matches.Count;
        }
    }
}
