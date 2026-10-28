// Porta alla versione italiana o inglese in base alla lingua del browser.
(function () {
  var lang = (navigator.language || 'en').toLowerCase();
  window.location.replace(lang.indexOf('it') === 0 ? 'it/' : 'en/');
})();
