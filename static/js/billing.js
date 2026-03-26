/* ============================================================
   billing.js — Dynamic calculation logic for Create Bill page
   ============================================================ */

document.addEventListener('DOMContentLoaded', () => {

  /* ── Toast utility ─────────────────────────────────────── */
  window.showToast = function(msg, type = 'info') {
    const container = document.getElementById('toast-container');
    if (!container) return;
    const toast = document.createElement('div');
    toast.className = `toast-msg ${type}`;
    const icons = { success:'✅', error:'❌', warning:'⚠️', info:'ℹ️' };
    toast.innerHTML = `<span>${icons[type] || 'ℹ️'}</span> ${msg}`;
    container.appendChild(toast);
    requestAnimationFrame(() => toast.classList.add('show'));
    setTimeout(() => {
      toast.classList.remove('show');
      setTimeout(() => toast.remove(), 350);
    }, 3500);
  };

  /* ── Loading overlay ────────────────────────────────────── */
  function showLoading()  { document.getElementById('loading-overlay').classList.add('show'); }
  function hideLoading()  { document.getElementById('loading-overlay').classList.remove('show'); }

  /* ── Flower row toggle ──────────────────────────────────── */
  document.querySelectorAll('.flower-checkbox').forEach(cb => {
    cb.addEventListener('change', () => {
      const item = cb.closest('.flower-item');
      const controls = item.querySelector('.flower-controls');
      if (cb.checked) {
        item.classList.add('selected');
        controls.classList.add('visible');
        const qtyInput = controls.querySelector('.qty-input');
        if (qtyInput && !qtyInput.value) qtyInput.focus();
      } else {
        item.classList.remove('selected');
        controls.classList.remove('visible');
        controls.querySelector('.qty-input').value = '';
        controls.querySelector('.rate-input').value = '';
        recalcItem(item);
        recalcTotals();
      }
    });
  });

  /* ── Per-item total ──────────────────────────────────────── */
  function recalcItem(item) {
    const qty  = parseFloat(item.querySelector('.qty-input')?.value)  || 0;
    const rate = parseFloat(item.querySelector('.rate-input')?.value) || 0;
    const total = qty * rate;
    const badge = item.querySelector('.item-total-badge');
    if (badge) badge.textContent = '₹' + total.toFixed(2);
    return total;
  }

  function recalcTotals() {
    let subtotal = 0;
    document.querySelectorAll('.flower-item').forEach(item => {
      if (item.querySelector('.flower-checkbox')?.checked) {
        subtotal += recalcItem(item);
      }
    });

    // Luggage calculation (qty * rate)
    const lugQty = parseFloat(document.getElementById('luggage-qty')?.value) || 0;
    const lugRate = parseFloat(document.getElementById('luggage-rate')?.value) || 0;
    const lugTotal = lugQty * lugRate;
    const lugEl = document.getElementById('luggage-total-display');
    if (lugEl) lugEl.textContent = '₹' + lugTotal.toFixed(2);

    // Include luggage in subtotal
    subtotal += lugTotal;

    const gstPct   = parseFloat(document.getElementById('gst-percent')?.value) || 0;
    const gstAmt   = subtotal * gstPct / 100;
    // Old balance (carryover) - added after GST
    const oldBal = parseFloat(document.getElementById('old-balance')?.value) || 0;

    // Pre-paid/paid (cash)
    const cashPaid = parseFloat(document.getElementById('cash-paid')?.value) || 0;

    const preGrand = subtotal + gstAmt + oldBal;
    // Subtract cash paid from total to get final payable (due)
    const dueRaw   = preGrand - cashPaid;
    const due      = Math.round(dueRaw * 100) / 100;
    const grand    = due >= 0 ? due : 0; // displayed grand total (after cash)

    setText('subtotal-display',  '₹' + subtotal.toFixed(2));
    setText('gst-amount-display','₹' + gstAmt.toFixed(2));
    setText('grand-total-display','₹' + grand.toFixed(2));
    // show old balance in totals area if present
    const oldRow = document.getElementById('old-balance-display');
    if (oldRow) oldRow.textContent = oldBal ? '₹' + oldBal.toFixed(2) : '₹0.00';

    // hidden fields for form submit
    setVal('hidden-subtotal',  subtotal.toFixed(2));
    setVal('hidden-gst-amount', gstAmt.toFixed(2));
    // hidden-grand-total keeps original invoice total before cash
    setVal('hidden-grand-total', preGrand.toFixed(2));
    setVal('hidden-paid',       cashPaid.toFixed(2));
    setVal('hidden-due',        (due >= 0 ? due : 0).toFixed(2));

    // Grand total in words (use displayed grand after cash)
    const wordsEl = document.getElementById('grand-total-words');
    if (wordsEl) {
      const words = numberToWords(grand);
      wordsEl.textContent = words ? `(${words})` : '';
    }
  }

  // Convert number to words (Rupees and Paise)
  function numberToWords(amount) {
    if (isNaN(amount)) return '';
    const n = Math.abs(Number(amount));
    const rupees = Math.floor(n);
    const paise = Math.round((n - rupees) * 100);

    function oneToWords(num) {
      const a = ['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine','Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen','Seventeen','Eighteen','Nineteen'];
      const b = ['', '', 'Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety'];
      if (num < 20) return a[num];
      if (num < 100) return b[Math.floor(num/10)] + (num%10 ? ' ' + a[num%10] : '');
      if (num < 1000) return a[Math.floor(num/100)] + ' Hundred' + (num%100 ? ' ' + oneToWords(num%100) : '');
      return '';
    }

    function convert(num) {
      if (num === 0) return 'Zero';
      const parts = [];
      const crore = Math.floor(num / 10000000);
      if (crore) { parts.push(convert(crore) + ' Crore'); num = num % 10000000; }
      const lakh = Math.floor(num / 100000);
      if (lakh) { parts.push(convert(lakh) + ' Lakh'); num = num % 100000; }
      const thousand = Math.floor(num / 1000);
      if (thousand) { parts.push(convert(thousand) + ' Thousand'); num = num % 1000; }
      const hundred = Math.floor(num / 100);
      if (hundred) { parts.push(oneToWords(hundred) + (num%100 ? ' ' + oneToWords(num%100) : '') ); return parts.join(' '); }
      if (num) parts.push(oneToWords(num));
      return parts.join(' ');
    }

    const rupeesText = convert(rupees) + (rupees === 1 ? ' Rupee' : ' Rupees');
    const paiseText = paise ? (convert(paise) + (paise === 1 ? ' Paise' : ' Paise')) : '';
    return paise ? `${rupeesText} and ${paiseText}` : rupeesText;
  }

  function setText(id, val) {
    const el = document.getElementById(id);
    if (el) el.textContent = val;
  }
  function setVal(id, val) {
    const el = document.getElementById(id);
    if (el) el.value = val;
  }

  /* ── Listen for qty/rate/gst changes ────────────────────── */
  document.querySelectorAll('.qty-input, .rate-input').forEach(input => {
    input.addEventListener('input', () => {
      // Prevent negative
      if (parseFloat(input.value) < 0) input.value = '';
      recalcTotals();
    });
  });

  // Also listen for luggage and old-balance specific inputs
  const lugQty = document.getElementById('luggage-qty');
  const lugRate = document.getElementById('luggage-rate');
  if (lugQty) lugQty.addEventListener('input', () => { if (parseFloat(lugQty.value) < 0) lugQty.value = ''; recalcTotals(); });
  if (lugRate) lugRate.addEventListener('input', () => { if (parseFloat(lugRate.value) < 0) lugRate.value = ''; recalcTotals(); });

  const oldBalInput = document.getElementById('old-balance');
  if (oldBalInput) oldBalInput.addEventListener('input', () => { if (parseFloat(oldBalInput.value) < 0) oldBalInput.value = ''; recalcTotals(); });

  const cashInput = document.getElementById('cash-paid');
  if (cashInput) cashInput.addEventListener('input', () => { if (parseFloat(cashInput.value) < 0) cashInput.value = ''; recalcTotals(); });

  const gstInput = document.getElementById('gst-percent');
  if (gstInput) {
    gstInput.addEventListener('input', () => {
      if (parseFloat(gstInput.value) < 0) gstInput.value = 0;
      recalcTotals();
    });
  }

  /* ── Form submission ─────────────────────────────────────── */
  const form = document.getElementById('bill-form');
  if (form) {
    form.addEventListener('submit', async (e) => {
      e.preventDefault();

      // Basic validation
      const name = document.getElementById('customer-name').value.trim();
      if (!name) {
        showToast('Customer name is required.', 'warning');
        document.getElementById('customer-name').focus();
        return;
      }

      let hasItem = false;
      document.querySelectorAll('.flower-checkbox').forEach(cb => {
        if (cb.checked) hasItem = true;
      });
      if (!hasItem) {
        showToast('Please select at least one flower item.', 'warning');
        return;
      }

      showLoading();

      const data = new FormData(form);

      try {
        const res = await fetch(form.action, {
          method: 'POST',
          body: data,
        });
        const json = await res.json();
        if (json.success) {
          showToast('Bill saved successfully! 🌸', 'success');
          setTimeout(() => {
            window.location.href = `/invoice/${json.bill_id}`;
          }, 800);
        } else {
          hideLoading();
          showToast(json.error || 'Failed to save bill.', 'error');
        }
      } catch (err) {
        hideLoading();
        showToast('Network error. Please try again.', 'error');
      }
    });
  }

  // Initial calculation on page load
  recalcTotals();

  /* ── Stagger-animate bill cards ─────────────────────────── */
  document.querySelectorAll('.bill-card').forEach((card, i) => {
    card.style.opacity = '0';
    card.style.transform = 'translateY(22px)';
    setTimeout(() => {
      card.style.transition = 'opacity .38s ease, transform .38s cubic-bezier(.4,0,.2,1)';
      card.style.opacity = '1';
      card.style.transform = 'translateY(0)';
    }, 60 + i * 55);
  });

  /* ── Sparkle burst on primary buttons ──────────────────── */
  document.querySelectorAll('.btn-primary, .btn-accent').forEach(btn => {
    btn.addEventListener('click', function(e) {
      const symbols = ['🌸','✨','💫','🌺','⭐','🪷'];
      for (let k = 0; k < 4; k++) {
        const dot = document.createElement('div');
        dot.className = 'sparkle-dot';
        dot.textContent = symbols[Math.floor(Math.random() * symbols.length)];
        const angle = (k / 4) * 2 * Math.PI + Math.random() * .8;
        const dist  = 48 + Math.random() * 40;
        dot.style.left = (e.clientX + Math.cos(angle) * dist * .3) + 'px';
        dot.style.top  = (e.clientY + Math.sin(angle) * dist * .3) + 'px';
        dot.style.animationDelay = (k * 60) + 'ms';
        document.body.appendChild(dot);
        setTimeout(() => dot.remove(), 1200);
      }
    });
  });
});

/* ── Floating petals background (all pages) ─────────────── */
(function() {
  const PETALS = ['🌸','🌺','🌷','🌼','🪷','✿'];
  function spawnPetal() {
    const el = document.createElement('div');
    el.className = 'petal';
    el.textContent = PETALS[Math.floor(Math.random() * PETALS.length)];
    el.style.left   = (5 + Math.random() * 90) + 'vw';
    el.style.bottom = '-2rem';
    const dur = 10 + Math.random() * 10;
    el.style.animationDuration  = dur + 's';
    el.style.animationDelay     = (Math.random() * 2) + 's';
    el.style.fontSize = (.7 + Math.random() * .7) + 'rem';
    document.body.appendChild(el);
    setTimeout(() => el.remove(), (dur + 3) * 1000);
  }
  // Initial burst
  if (!window.matchMedia('(prefers-reduced-motion: reduce)').matches) {
    for (let i = 0; i < 4; i++) setTimeout(spawnPetal, i * 800);
    setInterval(spawnPetal, 3800);
  }
})();
