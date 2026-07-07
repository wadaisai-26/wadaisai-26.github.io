document.addEventListener('DOMContentLoaded', function () {
  var btn = document.querySelector('.mobile-menu-btn');
  var nav = document.querySelector('.global-nav');
  if (!btn || !nav) return;

  btn.addEventListener('click', function () {
    nav.classList.toggle('open');
    btn.classList.toggle('active');
  });

  document.addEventListener('click', function (e) {
    if (!btn.contains(e.target) && !nav.contains(e.target)) {
      nav.classList.remove('open');
      btn.classList.remove('active');
    }
  });

  nav.querySelectorAll('a').forEach(function (a) {
    a.addEventListener('click', function () {
      nav.classList.remove('open');
      btn.classList.remove('active');
    });
  });
});
