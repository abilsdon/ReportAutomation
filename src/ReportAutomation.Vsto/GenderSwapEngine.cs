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
        internal GenderSwapScanResult FindDocumentCandidates(
            Word.Document document,
            Action<int, int> reportProgress)
        {
            return FindCandidates(new List<Word.Range> { document.Content }, reportProgress);
        }

        internal GenderSwapScanResult FindRangeCandidates(
            Word.Range range,
            Action<int, int> reportProgress)
        {
            return FindCandidates(new List<Word.Range> { range.Duplicate }, reportProgress);
        }

        private static GenderSwapScanResult FindCandidates(
            List<Word.Range> scopes,
            Action<int, int> reportProgress)
        {
            var candidates = new List<GenderSwapCandidate>();
            try
            {
                for (int scopeIndex = 0; scopeIndex < scopes.Count; scopeIndex++)
                {
                    Word.Range scope = scopes[scopeIndex];
                    var expectedMatches = new List<GenderTextMatch>(
                        GenderRules.FindMatches(scope.Text ?? string.Empty));
                    int total = expectedMatches.Count;
                    ReportProgress(reportProgress, 0, total);

                    int nextStart = scope.Start;
                    int scopeEnd = scope.End;
                    Word.Range search = scope.Duplicate;
                    Word.Find find = null;
                    try
                    {
                        find = search.Find;
                        find.ClearFormatting();
                        find.Forward = true;
                        find.Wrap = Word.WdFindWrap.wdFindStop;
                        find.Format = false;
                        find.MatchCase = false;
                        find.MatchWholeWord = true;
                        find.MatchWildcards = false;
                        find.MatchSoundsLike = false;
                        find.MatchAllWordForms = false;

                        for (int index = 0; index < total && nextStart < scopeEnd; index++)
                        {
                            GenderTextMatch expected = expectedMatches[index];
                            search.SetRange(nextStart, scopeEnd);
                            find.Text = expected.Value;
                            if (find.Execute())
                            {
                                string original = search.Text;
                                candidates.Add(new GenderSwapCandidate
                                {
                                    ScopeIndex = scopeIndex,
                                    Start = search.Start,
                                    Length = search.End - search.Start,
                                    Original = original,
                                    Replacement = GenderRules.ReplacementFor(original)
                                });

                                nextStart = search.End;
                            }

                            int completed = index + 1;
                            if (completed == total || completed % 10 == 0)
                            {
                                ReportProgress(reportProgress, completed, total);
                            }
                        }
                    }
                    finally
                    {
                        GenderSwapScanResult.ReleaseComObject(find);
                        GenderSwapScanResult.ReleaseComObject(search);
                    }
                }

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

        private static void ReportProgress(Action<int, int> reportProgress, int completed, int total)
        {
            if (reportProgress != null)
            {
                reportProgress(completed, total);
            }
        }
    }
}
