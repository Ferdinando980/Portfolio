// Filtri della home: funziona anche senza JavaScript, che mostra comunque tutte le schede.
(function () {
  var chips = document.querySelectorAll('.chip[data-filter]');
  var cards = document.querySelectorAll('.grid [data-area]');
  if (!chips.length || !cards.length) return;
  chips.forEach(function (chip) {
    chip.addEventListener('click', function () {
      var f = chip.getAttribute('data-filter');
      chips.forEach(function (c) { c.setAttribute('aria-pressed', c === chip ? 'true' : 'false'); });
      var note = document.getElementById('filter-note-recover');
      if (note) note.hidden = f !== 'recover';
      cards.forEach(function (card) {
        var also = (card.getAttribute('data-also') || '').split(' ');
        card.hidden = !(f === 'all' || card.getAttribute('data-area') === f || also.indexOf(f) !== -1);
        card.classList.toggle('is-ref', f !== 'all' && card.getAttribute('data-area') !== f && also.indexOf(f) !== -1);
      });
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
