using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;

namespace ReportAutomation.Vsto
{
    [ComVisible(true)]
    public sealed class GenderSwapRibbon : Office.IRibbonExtensibility
    {
        private static bool operationInProgress;

        private const string RibbonXml = @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab idMso='TabHome'>
        <group id='ReportAutomation.Group' label='Report Automation'>
          <button id='ReportAutomation.SwapDocument' label='Gender Swap' size='large'
                  imageMso='ReplaceDialog' screentip='Swap gendered words'
                  supertip='Review and approve each supported gendered-word change in the document body.'
                  onAction='SwapDocument' />
          <button id='ReportAutomation.SwapSelection' label='Swap Selection'
                  imageMso='SelectAll' screentip='Swap the selected text'
                  onAction='SwapSelection' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";

        public string GetCustomUI(string ribbonId)
        {
            return RibbonXml;
        }

        public void SwapDocument(Office.IRibbonControl control)
        {
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            ReviewChanges(
                "the document body",
                progress => new GenderSwapEngine().FindDocumentCandidates(document, progress));
        }

        public void SwapSelection(Office.IRibbonControl control)
        {
            Word.Selection selection = Globals.ThisAddIn.Application.Selection;
            if (selection == null || selection.Range.Start == selection.Range.End)
            {
                MessageBox.Show("Select some text first.", "Report Automation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Word.Range selectedRange = selection.Range.Duplicate;
            try
            {
                ReviewChanges(
                    "the current selection",
                    progress => new GenderSwapEngine().FindRangeCandidates(selectedRange, progress));
            }
            finally
            {
                GenderSwapScanResult.ReleaseComObject(selectedRange);
            }
        }

        private static void ReviewChanges(
            string scopeDescription,
            Func<Action<int, int>, GenderSwapScanResult> findCandidates)
        {
            if (operationInProgress)
            {
                return;
            }

            operationInProgress = true;
            Word.Application application = Globals.ThisAddIn.Application;
            try
            {
                application.StatusBar = "Gender Swap: scanning " + scopeDescription + "...";
                using (GenderSwapScanResult scan = findCandidates((completed, total) =>
                {
                    application.StatusBar = total == 0
                        ? "Gender Swap: no candidate words found."
                        : string.Format("Gender Swap: locating change {0} of {1}...", completed, total);
                    System.Windows.Forms.Application.DoEvents();
                }))
                {
                    List<GenderSwapCandidate> candidates = scan.Candidates;
                    if (candidates.Count == 0)
                    {
                        MessageBox.Show("No supported gendered words were found in " + scopeDescription + ".",
                            "Report Automation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    DialogResult confirmation = MessageBox.Show(
                        string.Format("Found {0} proposed {1} in {2}.\n\nReview each change?",
                            candidates.Count, candidates.Count == 1 ? "change" : "changes", scopeDescription),
                        "Gender Swap Preview", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (confirmation != DialogResult.Yes)
                    {
                        return;
                    }

                    var positionOffsets = new Dictionary<int, int>();
                    application.UndoRecord.StartCustomRecord("Gender Swap");
                    int approved = 0;
                    int skipped = 0;
                    bool cancelled = false;
                    try
                    {
                        for (int index = 0; index < candidates.Count; index++)
                        {
                            application.StatusBar = string.Format(
                                "Gender Swap: reviewing change {0} of {1}...", index + 1, candidates.Count);
                            GenderSwapCandidate candidate = candidates[index];
                            int positionOffset;
                            positionOffsets.TryGetValue(candidate.ScopeIndex, out positionOffset);
                            Word.Range candidateRange = scan.CreateRange(candidate, positionOffset);
                            try
                            {
                                candidateRange.Select();
                                application.ActiveWindow.ScrollIntoView(candidateRange, true);

                                DialogResult decision = MessageBox.Show(
                                    string.Format("Change {0} of {1}\n\n{2}  ->  {3}\n\nYes = approve   No = skip   Cancel = stop",
                                        index + 1, candidates.Count, candidate.Original, candidate.Replacement),
                                    "Review Gender Swap", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                                if (decision == DialogResult.Cancel)
                                {
                                    cancelled = true;
                                    break;
                                }

                                if (decision == DialogResult.Yes)
                                {
                                    candidateRange.Text = candidate.Replacement;
                                    positionOffsets[candidate.ScopeIndex] = positionOffset
                                        + candidate.Replacement.Length - candidate.Length;
                                    approved++;
                                }
                                else
                                {
                                    skipped++;
                                }
                            }
                            finally
                            {
                                GenderSwapScanResult.ReleaseComObject(candidateRange);
                            }
                        }
                    }
                    finally
                    {
                        application.UndoRecord.EndCustomRecord();
                    }

                    MessageBox.Show(string.Format(
                        "{0}\n\nApproved: {1}\nSkipped: {2}\n\nApproved changes can be undone as one Word action.",
                        cancelled ? "Review stopped." : "Review complete.", approved, skipped), "Report Automation",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show("Gender Swap could not complete:\n\n" + exception.Message,
                    "Report Automation", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                application.StatusBar = string.Empty;
                operationInProgress = false;
            }
        }
    }
}
