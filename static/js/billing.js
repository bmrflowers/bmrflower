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

    const gstPct   = parseFloat(document.getElementById('gst-percent')?.value) || 0;
    const gstAmt   = subtotal * gstPct / 100;
    const grand    = subtotal + gstAmt;

    setText('subtotal-display',  '₹' + subtotal.toFixed(2));
    setText('gst-amount-display','₹' + gstAmt.toFixed(2));
    setText('grand-total-display','₹' + grand.toFixed(2));

    // hidden fields for form submit
    setVal('hidden-subtotal',  subtotal.toFixed(2));
    setVal('hidden-gst-amount', gstAmt.toFixed(2));
    setVal('hidden-grand-total', grand.toFixed(2));
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
