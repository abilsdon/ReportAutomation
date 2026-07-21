/* global Office, Word, GenderSwapCore */
"use strict";
const terms = Object.keys(GenderSwapCore.replacements);
let previewState = null;

Office.onReady(info => {
  if (info.host !== Office.HostType.Word) return;
  document.getElementById("preview").addEventListener("click", previewChanges);
  document.getElementById("apply").addEventListener("click", applyChanges);
  document.querySelectorAll('input[name="scope"]').forEach(input => input.addEventListener("change", invalidatePreview));
});

function scopeName() { return document.querySelector('input[name="scope"]:checked').value; }
function scopeRange(context, scope) { return scope === "selection" ? context.document.getSelection() : context.document.body; }
function findMatches(range) { return terms.map(term => range.search(term, { matchCase: false, matchWholeWord: true })); }

async function previewChanges() {
  setBusy(true);
  try {
    const state = await Word.run(async context => {
      const scope = scopeName();
      const collections = findMatches(scopeRange(context, scope));
      collections.forEach(collection => collection.load("items/text"));
      await context.sync();
      return { scope, count: collections.reduce((sum, collection) => sum + collection.items.length, 0) };
    });
    previewState = state;
    document.getElementById("apply").disabled = state.count === 0;
    setStatus(state.count ? `Found ${state.count} ${state.count === 1 ? "word" : "words"} to replace.` : "No supported gendered words were found.");
  } catch (error) { showError(error); } finally { setBusy(false); }
}

async function applyChanges() {
  if (!previewState || previewState.scope !== scopeName()) return;
  setBusy(true);
  try {
    const count = await Word.run(async context => {
      const collections = findMatches(scopeRange(context, previewState.scope));
      collections.forEach(collection => collection.load("items/text"));
      await context.sync();
      let changed = 0;
      collections.forEach(collection => collection.items.forEach(range => {
        range.insertText(GenderSwapCore.replacementFor(range.text), Word.InsertLocation.replace);
        changed += 1;
      }));
      await context.sync();
      return changed;
    });
    previewState = null;
    document.getElementById("apply").disabled = true;
    setStatus(`Complete — replaced ${count} ${count === 1 ? "word" : "words"}. Please review the document.`);
  } catch (error) { showError(error); } finally { setBusy(false); }
}

function invalidatePreview() {
  previewState = null; document.getElementById("apply").disabled = true;
  setStatus("Scope changed. Preview again before applying changes.");
}
function setBusy(busy) {
  document.getElementById("preview").disabled = busy;
  if (busy) document.getElementById("apply").disabled = true;
}
function setStatus(message, error = false) {
  const element = document.getElementById("status");
  element.textContent = message; element.classList.toggle("error", error);
}
function showError(error) {
  console.error(error); setStatus(`Word could not complete the operation: ${error.message || error}`, true);
}
