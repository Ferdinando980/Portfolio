// Filtri della pagina Cybersecurity: senza JavaScript restano visibili tutte le sezioni.
(function () {
  var chips = document.querySelectorAll('.chip[data-filter]');
  var sections = document.querySelectorAll('.fn[data-fn]');
  if (!chips.length || !sections.length) return;
  chips.forEach(function (chip) {
    chip.addEventListener('click', function () {
      var f = chip.getAttribute('data-filter');
      chips.forEach(function (c) { c.setAttribute('aria-pressed', c === chip ? 'true' : 'false'); });
      sections.forEach(function (s) { s.hidden = !(f === 'all' || s.getAttribute('data-fn') === f); });
    });
  });
})();

// Protezione anti-clickjacking: GitHub Pages non permette di impostare gli header, quindi la pagina non resta dentro un frame altrui.
(function () {
  if (window.top !== window.self) {
    try { window.top.location.href = window.location.href; }
    catch (e) { document.documentElement.hidden = true; }
  }
})();
