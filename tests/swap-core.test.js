const test = require("node:test");
const assert = require("node:assert/strict");
const { replacementFor } = require("../swap-core");

test("uses requested replacement rules", () => {
  assert.equal(replacementFor("he"), "her");
  assert.equal(replacementFor("she"), "him");
  assert.equal(replacementFor("him"), "her");
  assert.equal(replacementFor("her"), "him");
  assert.equal(replacementFor("his"), "hers");
  assert.equal(replacementFor("hers"), "his");
  assert.equal(replacementFor("himself"), "herself");
});

test("preserves common capitalisation", () => {
  assert.equal(replacementFor("He"), "Her");
  assert.equal(replacementFor("SHE"), "HIM");
  assert.equal(replacementFor("Herself"), "Himself");
});

test("leaves unsupported words unchanged", () => {
  assert.equal(replacementFor("hero"), "hero");
});
