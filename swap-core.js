(function (root, factory) {
  const api = factory();
  if (typeof module === "object" && module.exports) module.exports = api;
  root.GenderSwapCore = api;
})(typeof globalThis !== "undefined" ? globalThis : this, function () {
  "use strict";
  const replacements = Object.freeze({
    he: "her", she: "him", him: "her", her: "him",
    his: "hers", hers: "his", himself: "herself", herself: "himself",
    man: "woman", woman: "man", men: "women", women: "men"
  });
  function matchCase(source, target) {
    if (source === source.toUpperCase()) return target.toUpperCase();
    if (source[0] === source[0].toUpperCase()) return target[0].toUpperCase() + target.slice(1);
    return target;
  }
  function replacementFor(word) {
    const replacement = replacements[word.toLowerCase()];
    return replacement ? matchCase(word, replacement) : word;
  }
  return { replacements, matchCase, replacementFor };
});
