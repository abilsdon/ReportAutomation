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
        private const string RibbonXml = @"
<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
  <ribbon>
    <tabs>
      <tab idMso='TabHome'>
        <group id='ReportAutomation.Group' label='Report Automation'>
          <button id='ReportAutomation.SwapDocument' label='Gender Swap' size='large'
                  imageMso='ReplaceDialog' screentip='Swap gendered words'
                  supertip='Review and approve each supported gendered-word change throughout the document.'
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
                "the entire document, including headers, footers and footnotes",
                () => new GenderSwapEngine().FindDocumentCandidates(document));
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
            ReviewChanges(
                "the current selection",
                () => new GenderSwapEngine().FindRangeCandidates(selectedRange));
        }

        private static void ReviewChanges(
            string scopeDescription,
            Func<List<GenderSwapCandidate>> findCandidates)
        {
            try
            {
                List<GenderSwapCandidate> candidates = findCandidates();
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

                Word.Application application = Globals.ThisAddIn.Application;
                application.UndoRecord.StartCustomRecord("Gender Swap");
                int approved = 0;
                int skipped = 0;
                bool cancelled = false;
                try
                {
                    for (int index = 0; index < candidates.Count; index++)
                    {
                        GenderSwapCandidate candidate = candidates[index];
                        candidate.Range.Select();
                        application.ActiveWindow.ScrollIntoView(candidate.Range, true);

                        DialogResult decision = MessageBox.Show(
                            string.Format("Change {0} of {1}\n\n{2}  →  {3}\n\nYes = approve   No = skip   Cancel = stop",
                                index + 1, candidates.Count, candidate.Original, candidate.Replacement),
                            "Review Gender Swap", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                        if (decision == DialogResult.Cancel)
                        {
                            cancelled = true;
                            break;
                        }

                        if (decision == DialogResult.Yes)
                        {
                            candidate.Range.Text = candidate.Replacement;
                            approved++;
                        }
                        else
                        {
                            skipped++;
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
            catch (Exception exception)
            {
                MessageBox.Show("Gender Swap could not complete:\n\n" + exception.Message,
                    "Report Automation", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
