// Sheet tabs + cell selection with name box and formula bar (read-only viewer).
(() => {
  const tabs = document.querySelectorAll('.tab');
  const sheets = document.querySelectorAll('.sheet');
  const namebox = document.getElementById('namebox');
  const fcontent = document.getElementById('fcontent');
  let selected = null;

  const colName = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

  function show(i) {
    tabs.forEach((t) => { const on = t.dataset.i === String(i); t.classList.toggle('on', on); t.setAttribute('aria-selected', on); });
    sheets.forEach((s) => { s.hidden = s.dataset.i !== String(i); });
    if (selected) selected.classList.remove('sel');
    selected = null; namebox.textContent = 'A1'; fcontent.textContent = '';
    history.replaceState(null, '', '#' + i);
  }
  tabs.forEach((t) => t.addEventListener('click', () => show(t.dataset.i)));
  const start = parseInt(location.hash.slice(1), 10);
  if (!Number.isNaN(start) && start < tabs.length) show(start);

  // Address of a td: count the physical column, accounting for colspan/rowspan of previous cells.
  function address(td) {
    const tr = td.parentElement;
    const table = tr.closest('table');
    const rowIndex = Array.prototype.indexOf.call(tr.parentElement.children, tr);
    const occupied = new Map();
    const rows = tr.parentElement.children;
    for (let r = 0; r <= rowIndex; r++) {
      let c = 1;
      for (const cell of rows[r].children) {
        if (cell.tagName !== 'TD') continue;
        while ((occupied.get(r) || new Set()).has(c)) c++;
        const rs = cell.rowSpan || 1, cs = cell.colSpan || 1;
        if (cell === td) return colName(c) + (rowIndex + 1);
        for (let dr = 0; dr < rs; dr++) {
          if (!occupied.has(r + dr)) occupied.set(r + dr, new Set());
          for (let dc = 0; dc < cs; dc++) occupied.get(r + dr).add(c + dc);
        }
        c += cs;
      }
    }
    return '';
  }

  document.addEventListener('click', (e) => {
    const td = e.target.closest('table.xl td');
    if (!td) return;
    if (selected) selected.classList.remove('sel');
    selected = td; td.classList.add('sel');
    namebox.textContent = address(td);
    fcontent.textContent = td.dataset.f || td.textContent;
  });
})();
