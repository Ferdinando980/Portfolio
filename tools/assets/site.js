// Filtri della home: funziona anche senza JavaScript, che mostra comunque tutte le schede.
(function () {
  var chips = document.querySelectorAll('.chip[data-filter]');
  var cards = document.querySelectorAll('.grid [data-area]');
  if (!chips.length || !cards.length) return;
  var grid = document.querySelector('.grid');
  var divider = document.getElementById('ref-divider');
  cards.forEach(function (card, i) { card.setAttribute('data-i', i); });
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
      var ordered = Array.prototype.slice.call(cards).sort(function (a, b) {
        var ra = a.classList.contains('is-ref') ? 1 : 0, rb = b.classList.contains('is-ref') ? 1 : 0;
        return (ra - rb) || (a.getAttribute('data-i') - b.getAttribute('data-i'));
      });
      ordered.forEach(function (c) { grid.appendChild(c); });
      var firstRef = ordered.filter(function (c) { return !c.hidden && c.classList.contains('is-ref'); })[0];
      var hasMain = ordered.some(function (c) { return !c.hidden && !c.classList.contains('is-ref'); });
      if (divider) {
        if (firstRef && hasMain) { grid.insertBefore(divider, firstRef); divider.hidden = false; }
        else { divider.hidden = true; }
      }
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
