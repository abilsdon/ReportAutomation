using System;
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
                  supertip='Preview and swap supported gendered words throughout the document.'
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
            RunSwap(
                "the entire document, including headers, footers and footnotes",
                () => new GenderSwapEngine().CountDocument(document),
                () => new GenderSwapEngine().SwapDocument(document));
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
            RunSwap(
                "the current selection",
                () => new GenderSwapEngine().CountRange(selectedRange),
                () => new GenderSwapEngine().SwapRange(selectedRange));
        }

        private static void RunSwap(string scopeDescription, Func<int> countAction, Func<int> swapAction)
        {
            try
            {
                int previewCount = countAction();
                if (previewCount == 0)
                {
                    MessageBox.Show("No supported gendered words were found in " + scopeDescription + ".",
                        "Report Automation", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                DialogResult confirmation = MessageBox.Show(
                    string.Format("Found {0} {1} to change in {2}.\n\nContinue?",
                        previewCount, previewCount == 1 ? "word" : "words", scopeDescription),
                    "Gender Swap Preview", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (confirmation != DialogResult.Yes)
                {
                    return;
                }

                Word.Application application = Globals.ThisAddIn.Application;
                bool previousScreenUpdating = application.ScreenUpdating;
                application.ScreenUpdating = false;
                application.UndoRecord.StartCustomRecord("Gender Swap");
                int changed;
                try
                {
                    changed = swapAction();
                }
                finally
                {
                    application.UndoRecord.EndCustomRecord();
                    application.ScreenUpdating = previousScreenUpdating;
                }

                MessageBox.Show(string.Format("Complete — changed {0} {1}. You can undo this as one Word action.",
                    changed, changed == 1 ? "word" : "words"), "Report Automation",
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
