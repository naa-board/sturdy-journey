(function () {
  var header = document.querySelector('.site-header');
  var toggle = document.querySelector('.nav-toggle');
  if (!toggle) return;

  function open()  { header.classList.add('nav-open');    toggle.setAttribute('aria-expanded', 'true');  toggle.setAttribute('aria-label', 'Close navigation'); }
  function close() { header.classList.remove('nav-open'); toggle.setAttribute('aria-expanded', 'false'); toggle.setAttribute('aria-label', 'Open navigation'); }

  toggle.addEventListener('click', function () {
    header.classList.contains('nav-open') ? close() : open();
  });

  // Close when a nav link is followed
  document.querySelectorAll('nav a').forEach(function (a) {
    a.addEventListener('click', close);
  });

  // Close on outside click
  document.addEventListener('click', function (e) {
    if (!header.contains(e.target)) close();
  });

  // Close when Escape is pressed
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') close();
  });
})();
