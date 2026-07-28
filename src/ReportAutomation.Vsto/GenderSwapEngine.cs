using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace ReportAutomation.Vsto
{
    internal sealed class GenderSwapCandidate
    {
        internal Word.Range Range { get; set; }
        internal string Original { get; set; }
        internal string Replacement { get; set; }
    }

    internal sealed class GenderSwapEngine
    {
        internal List<GenderSwapCandidate> FindDocumentCandidates(Word.Document document)
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

            return FindCandidates(stories);
        }

        internal List<GenderSwapCandidate> FindRangeCandidates(Word.Range range)
        {
            return FindCandidates(new[] { range });
        }

        private static List<GenderSwapCandidate> FindCandidates(IEnumerable<Word.Range> scopes)
        {
            var candidates = new List<GenderSwapCandidate>();
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
                        candidates.Add(new GenderSwapCandidate
                        {
                            Range = search.Duplicate,
                            Original = search.Text,
                            Replacement = GenderRules.ReplacementFor(search.Text)
                        });
                        search.Start = search.End;
                        search.End = scopeEnd;
                    }
                }
            }

            candidates.Sort((left, right) =>
            {
                int storyComparison = left.Range.StoryType.CompareTo(right.Range.StoryType);
                return storyComparison != 0 ? storyComparison : left.Range.Start.CompareTo(right.Range.Start);
            });

            return candidates;
        }
    }
}
