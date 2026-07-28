using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Word = Microsoft.Office.Interop.Word;

namespace ReportAutomation.Vsto
{
    internal sealed class GenderSwapCandidate
    {
        internal int ScopeIndex { get; set; }
        internal int Start { get; set; }
        internal int Length { get; set; }
        internal Word.WdStoryType StoryType { get; set; }
        internal string Original { get; set; }
        internal string Replacement { get; set; }
    }

    internal sealed class GenderSwapScanResult : IDisposable
    {
        private readonly List<Word.Range> scopes;

        internal GenderSwapScanResult(List<Word.Range> scopes, List<GenderSwapCandidate> candidates)
        {
            this.scopes = scopes;
            Candidates = candidates;
        }

        internal List<GenderSwapCandidate> Candidates { get; private set; }

        internal Word.Range CreateRange(GenderSwapCandidate candidate, int positionOffset)
        {
            Word.Range range = scopes[candidate.ScopeIndex].Duplicate;
            int start = candidate.Start + positionOffset;
            range.SetRange(start, start + candidate.Length);
            return range;
        }

        public void Dispose()
        {
            foreach (Word.Range scope in scopes)
            {
                ReleaseComObject(scope);
            }

            scopes.Clear();
            Candidates.Clear();
        }

        internal static void ReleaseComObject(object value)
        {
            if (value != null && Marshal.IsComObject(value))
            {
                Marshal.FinalReleaseComObject(value);
            }
        }
    }

    internal sealed class GenderSwapEngine
    {
        internal GenderSwapScanResult FindDocumentCandidates(Word.Document document)
        {
            var scopes = new List<Word.Range>();
            foreach (Word.Range firstStory in document.StoryRanges)
            {
                Word.Range story = firstStory;
                try
                {
                    while (story != null)
                    {
                        scopes.Add(story.Duplicate);
                        Word.Range nextStory = story.NextStoryRange;
                        if (!ReferenceEquals(story, firstStory))
                        {
                            GenderSwapScanResult.ReleaseComObject(story);
                        }

                        story = nextStory;
                    }
                }
                finally
                {
                    if (story != null && !ReferenceEquals(story, firstStory))
                    {
                        GenderSwapScanResult.ReleaseComObject(story);
                    }

                    GenderSwapScanResult.ReleaseComObject(firstStory);
                }
            }

            return FindCandidates(scopes);
        }

        internal GenderSwapScanResult FindRangeCandidates(Word.Range range)
        {
            return FindCandidates(new List<Word.Range> { range.Duplicate });
        }

        private static GenderSwapScanResult FindCandidates(List<Word.Range> scopes)
        {
            var candidates = new List<GenderSwapCandidate>();
            try
            {
                for (int scopeIndex = 0; scopeIndex < scopes.Count; scopeIndex++)
                {
                    Word.Range scope = scopes[scopeIndex];
                    string text = scope.Text ?? string.Empty;
                    foreach (GenderTextMatch match in GenderRules.FindMatches(text))
                    {
                        candidates.Add(new GenderSwapCandidate
                        {
                            ScopeIndex = scopeIndex,
                            Start = scope.Start + match.Index,
                            Length = match.Length,
                            StoryType = scope.StoryType,
                            Original = match.Value,
                            Replacement = GenderRules.ReplacementFor(match.Value)
                        });
                    }
                }

                candidates.Sort((left, right) =>
                {
                    int storyComparison = left.StoryType.CompareTo(right.StoryType);
                    return storyComparison != 0 ? storyComparison : left.Start.CompareTo(right.Start);
                });

                return new GenderSwapScanResult(scopes, candidates);
            }
            catch
            {
                foreach (Word.Range scope in scopes)
                {
                    GenderSwapScanResult.ReleaseComObject(scope);
                }

                throw;
            }
        }
    }
}
