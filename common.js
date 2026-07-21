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

  var modal = document.getElementById('program-modal');
  if (!modal) return;
  var modalHead = document.getElementById('program-modal-head');
  var modalSection = document.getElementById('program-modal-section');
  var modalTitle = document.getElementById('program-modal-title');
  var modalBody = document.getElementById('program-modal-body');
  var modalSchedule = document.getElementById('program-modal-schedule');
  var modalClose = document.getElementById('program-modal-close');

  function openProgramModal(card) {
    var ph = card.querySelector('.ph');
    var section = card.querySelector('.ph-section');
    var title = card.querySelector('.ph h3');
    var body = card.querySelector('.pb');
    var schedule = card.querySelector('.pt');
    if (ph) {
      var phStyle = getComputedStyle(ph);
      modalHead.style.backgroundColor = phStyle.backgroundColor;
      modalHead.style.backgroundImage = phStyle.backgroundImage;
    }
    modalSection.textContent = section ? section.textContent : '';
    modalTitle.textContent = title ? title.textContent : '';
    modalBody.innerHTML = body ? body.innerHTML : '';
    modalSchedule.innerHTML = schedule ? schedule.innerHTML : '';

    var noteHeight = modal.offsetHeight;
    var cardRect = card.getBoundingClientRect();
    var spaceBelow = window.innerHeight - cardRect.bottom;
    var top;
    if (noteHeight + 8 <= spaceBelow) {
      top = card.offsetTop + card.offsetHeight + 8;
    } else {
      top = Math.max(0, card.offsetTop - noteHeight - 8);
    }
    modal.style.top = top + 'px';
    modal.classList.add('open');
  }

  function closeProgramModal() {
    modal.classList.remove('open');
  }

  var isMobile = function () { return window.matchMedia('(max-width: 768px)').matches; };

  document.querySelectorAll('.program-card').forEach(function (card) {
    card.addEventListener('click', function (e) {
      if (!isMobile()) return;
      e.stopPropagation();
      openProgramModal(card);
    });
  });

  modal.addEventListener('click', function (e) { e.stopPropagation(); });
  modalClose.addEventListener('click', closeProgramModal);
  document.addEventListener('click', function () { closeProgramModal(); });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') closeProgramModal();
  });
});
