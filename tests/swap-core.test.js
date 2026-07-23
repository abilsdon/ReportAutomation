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
  assert.equal(replacementFor("man"), "woman");
  assert.equal(replacementFor("woman"), "man");
  assert.equal(replacementFor("men"), "women");
  assert.equal(replacementFor("women"), "men");
});

test("preserves common capitalisation", () => {
  assert.equal(replacementFor("He"), "Her");
  assert.equal(replacementFor("SHE"), "HIM");
  assert.equal(replacementFor("Herself"), "Himself");
  assert.equal(replacementFor("MAN"), "WOMAN");
  assert.equal(replacementFor("Women"), "Men");
});

test("leaves unsupported words unchanged", () => {
  assert.equal(replacementFor("hero"), "hero");
});
