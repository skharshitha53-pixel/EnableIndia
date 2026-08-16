/* Incluzza — micro-interactions */
(() => {
  const reduce = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  const $$ = (s, c = document) => Array.from(c.querySelectorAll(s));

  // ── Reveal on scroll ─────────────────────────────────────────────
  const items = $$('[data-reveal]');
  if (!reduce && 'IntersectionObserver' in window) {
    const io = new IntersectionObserver((entries) => {
      entries.forEach((e) => {
        if (e.isIntersecting) {
          const siblings = Array.from(e.target.parentElement?.children || [])
            .filter((c) => c.hasAttribute('data-reveal'));
          const idx = siblings.indexOf(e.target);
          const delay = Math.max(0, Math.min(idx * 70, 280));
          e.target.style.transitionDelay = `${delay}ms`;
          e.target.classList.add('is-visible');
          io.unobserve(e.target);
        }
      });
    }, { rootMargin: '0px 0px -8% 0px', threshold: 0.08 });
    items.forEach((el) => io.observe(el));
  } else {
    items.forEach((el) => el.classList.add('is-visible'));
  }

  // ── Active nav link based on which "part" you're in ─────────────
  const navLinks = $$('.nav-links a[href^="#"]');
  const partSections = navLinks
    .map((a) => document.querySelector(a.getAttribute('href')))
    .filter(Boolean);

  if ('IntersectionObserver' in window) {
    const activeIo = new IntersectionObserver((entries) => {
      entries.forEach((e) => {
        if (!e.isIntersecting) return;
        const id = '#' + e.target.id;
        navLinks.forEach((a) => {
          a.style.color = a.getAttribute('href') === id ? 'var(--ink)' : '';
        });
      });
    }, { rootMargin: '-30% 0px -60% 0px', threshold: 0 });
    partSections.forEach((s) => activeIo.observe(s));
  }

  // ── Parallax the ambient orbs (very subtle) ──────────────────────
  if (!reduce) {
    const orbs = $$('.orb');
    let raf = null;
    const onScroll = () => {
      if (raf) return;
      raf = requestAnimationFrame(() => {
        const y = window.scrollY;
        orbs.forEach((o, i) => {
          const k = (i + 1) * 0.05;
          o.style.translate = `0 ${y * k * -0.2}px`;
        });
        raf = null;
      });
    };
    window.addEventListener('scroll', onScroll, { passive: true });
  }

  // ── Service cards: cycle accent colors ───────────────────────────
  const services = $$('.service');
  services.forEach((el, i) => {
    const c = `var(--c${(i % 8) + 1})`;
    el.style.setProperty('--svc-color', c);
  });
})();