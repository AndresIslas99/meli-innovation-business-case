/* =====================================================================
   DECK NAVIGATION · keyboard + click controls
   ===================================================================== */

(function() {
  'use strict';

  const slides = Array.from(document.querySelectorAll('.slide'));
  const total = slides.length;
  let current = 0;

  // Try to restore from URL hash
  const hashMatch = window.location.hash.match(/^#slide-(\d+)$/);
  if (hashMatch) {
    const fromHash = parseInt(hashMatch[1], 10);
    if (fromHash >= 0 && fromHash < total) {
      current = fromHash;
    }
  }

  function update() {
    slides.forEach((s, i) => {
      s.classList.toggle('is-active', i === current);
    });
    const counter = document.querySelector('.deck-header__counter');
    if (counter) {
      counter.textContent = `${current + 1} / ${total}`;
    }
    const prev = document.querySelector('[data-deck-prev]');
    const next = document.querySelector('[data-deck-next]');
    if (prev) prev.disabled = current === 0;
    if (next) next.disabled = current === total - 1;

    // Update URL hash without page jump
    history.replaceState(null, '', `#slide-${current}`);
    window.scrollTo({ top: 0, behavior: 'smooth' });

    // Fire a custom event for widgets that may want to lazy-init
    document.dispatchEvent(new CustomEvent('deck:navigate', {
      detail: { current, slideId: slides[current].id }
    }));
  }

  function go(dir) {
    const next = current + dir;
    if (next >= 0 && next < total) {
      current = next;
      update();
    }
  }

  function goTo(i) {
    if (i >= 0 && i < total) {
      current = i;
      update();
    }
  }

  // Keyboard navigation
  document.addEventListener('keydown', (e) => {
    // Don't intercept when user is typing in an input/textarea
    if (['INPUT', 'TEXTAREA', 'SELECT'].includes(e.target.tagName)) return;
    if (e.target.isContentEditable) return;
    if (e.ctrlKey || e.metaKey || e.altKey) return;

    switch (e.key) {
      case 'ArrowRight':
      case 'PageDown':
      case ' ':
        e.preventDefault();
        go(+1);
        break;
      case 'ArrowLeft':
      case 'PageUp':
        e.preventDefault();
        go(-1);
        break;
      case 'Home':
        e.preventDefault();
        goTo(0);
        break;
      case 'End':
        e.preventDefault();
        goTo(total - 1);
        break;
      case 'h':
        e.preventDefault();
        goTo(0);
        break;
      case 'p':
        e.preventDefault();
        window.print();
        break;
      case 't':
        e.preventDefault();
        toggleTOC();
        break;
      case 'Escape':
        closeTOC();
        break;
    }
  });

  // Wire nav buttons
  document.addEventListener('click', (e) => {
    if (e.target.closest('[data-deck-prev]')) {
      e.preventDefault();
      go(-1);
    } else if (e.target.closest('[data-deck-next]')) {
      e.preventDefault();
      go(+1);
    } else if (e.target.closest('[data-deck-toc]')) {
      e.preventDefault();
      toggleTOC();
    } else if (e.target.closest('[data-deck-print]')) {
      e.preventDefault();
      window.print();
    } else {
      const tocItem = e.target.closest('[data-goto]');
      if (tocItem) {
        e.preventDefault();
        const idx = parseInt(tocItem.dataset.goto, 10);
        goTo(idx);
        closeTOC();
      }
    }
  });

  // TOC overlay
  function toggleTOC() {
    const overlay = document.querySelector('.toc-overlay');
    if (!overlay) return;
    overlay.classList.toggle('is-open');
  }
  function closeTOC() {
    const overlay = document.querySelector('.toc-overlay');
    if (!overlay) return;
    overlay.classList.remove('is-open');
  }
  document.addEventListener('click', (e) => {
    if (e.target.classList && e.target.classList.contains('toc-overlay')) {
      closeTOC();
    }
  });

  // Tabs (inside slides)
  document.addEventListener('click', (e) => {
    const tab = e.target.closest('.tab-btn');
    if (!tab) return;
    e.preventDefault();
    const group = tab.dataset.tabGroup;
    const target = tab.dataset.tabTarget;
    document.querySelectorAll(`.tab-btn[data-tab-group="${group}"]`).forEach(b => b.classList.remove('is-active'));
    document.querySelectorAll(`.tab-content[data-tab-group="${group}"]`).forEach(c => c.classList.remove('is-active'));
    tab.classList.add('is-active');
    const targetEl = document.querySelector(`.tab-content[data-tab-group="${group}"][data-tab-name="${target}"]`);
    if (targetEl) targetEl.classList.add('is-active');
  });

  // Initial paint
  update();

  // Expose for widgets
  window.MELI_DECK = { go, goTo, current: () => current, total };
})();
