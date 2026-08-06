/**
 * pdf-viewer.js - shared inline PDF viewer (Bootstrap modal).
 *
 * Renders PDF pages as images server-side. Default zoom = fit-to-width, with a
 * page thumbnail strip for quick jumping. Any element with class "pdf-open" and
 * a data-src attribute opens it on click. Pages load from:
 *   <data-src>/pages  -> {"pages": N}
 *   <data-src>/page/N -> rendered page image
 */
(function () {
  function init() {
    var modal = document.getElementById('pdfModal');
    if (!modal) return;
    var title = document.getElementById('pdfTitle');
    var view = document.getElementById('pdfZoomView');
    var img = document.getElementById('pdfPageImg');
    var pct = document.getElementById('pdfZoomPct');
    var pageInput = document.getElementById('pdfPageNum');
    var pageTotalEl = document.getElementById('pdfPageTotal');
    var loadingEl = document.getElementById('pdfLoading');
    var thumbsEl = document.getElementById('pdfThumbs');

    var zoom = 1;
    var baseImgW = 800;
    var pages = 1, currentPage = 1, currentSrc = '';
    var fitMode = false;  // default = 100% (not fit-to-width)

    function showLoading() {
      if (loadingEl) loadingEl.style.display = 'flex';
    }
    function hideLoading() {
      if (loadingEl) loadingEl.style.display = 'none';
    }
    function applyZoom() {
      img.style.width = Math.round(baseImgW * zoom) + 'px';
      pct.textContent = Math.round(zoom * 100) + '%';
    }
    function fitToWidth() {
      var natural = img.naturalWidth || baseImgW;
      var cw = view.clientWidth || 800;
      zoom = Math.max(cw / natural, 0.1);
      applyZoom();
    }
    function updatePageUI() {
      pageTotalEl.textContent = pages;
      pageInput.max = pages;
      pageInput.value = currentPage;
    }

    function buildThumbs() {
      if (!thumbsEl) return;
      thumbsEl.innerHTML = '';
      for (var i = 1; i <= pages; i++) {
        (function (n) {
          var t = document.createElement('div');
          t.className = 'pdf-thumb' + (n === currentPage ? ' active' : '');
          t.innerHTML = '<img src="' + currentSrc + '/page/' + n + '" alt="page ' + n + '">' +
                        '<span>' + n + '</span>';
          t.addEventListener('click', function () { loadPage(n); });
          thumbsEl.appendChild(t);
        })(i);
      }
    }
    function highlightThumbs() {
      if (!thumbsEl) return;
      Array.prototype.forEach.call(thumbsEl.children, function (t, idx) {
        t.classList.toggle('active', (idx + 1) === currentPage);
      });
    }

    function loadPage(n) {
      n = Math.min(Math.max(parseInt(n) || 1, 1), pages);
      currentPage = n;
      updatePageUI();
      highlightThumbs();
      var next = currentSrc + '/page/' + n;
      if (img.src !== next) {
        loading = true;
        showLoading();
        img.src = next;
        img.onload = function () {
          loading = false; hideLoading();
          if (fitMode) fitToWidth(); else applyZoom();
        };
        img.onerror = function () { loading = false; hideLoading(); };
      }
      view.scrollTop = 0;
      view.scrollLeft = 0;
    }

    function openPdf(src, name) {
      currentSrc = src;
      title.textContent = name || 'Quotation';
      var bs = window.bootstrap && bootstrap.Modal.getOrCreateInstance(modal);
      if (bs) bs.show();
      currentPage = 1;
      fitMode = false;  // open at 100%
      pages = 1;
      updatePageUI();
      showLoading();
      fetch(src + '/pages').then(function (r) { return r.json(); })
        .then(function (d) {
          pages = d.pages || 1;
          updatePageUI();
          buildThumbs();
          loadPage(1);
        })
        .catch(function () { pages = 1; updatePageUI(); buildThumbs(); loadPage(1); });
      img.src = src + '/page/1';
      img.onload = function () {
        baseImgW = img.naturalWidth || 800;
        loading = false; hideLoading();
        if (fitMode) fitToWidth(); else applyZoom();
      };
      img.onerror = function () { loading = false; hideLoading(); };
    }

    // Delegated click on any .pdf-open trigger
    document.addEventListener('click', function (e) {
      var el = e.target;
      var t = el.closest ? el.closest('.pdf-open') : null;
      if (!t) return;
      var control = el.closest ? el.closest('button, a, input, select, textarea, form') : null;
      if (control && control !== t && t.contains(control)) return;
      openPdf(t.getAttribute('data-src'), t.getAttribute('data-name'));
    });

    document.getElementById('pdfZoomIn').onclick = function () {
      fitMode = false; zoom = Math.min(zoom + 0.25, 5); applyZoom();
    };
    document.getElementById('pdfZoomOut').onclick = function () {
      fitMode = false; zoom = Math.max(zoom - 0.25, 0.25); applyZoom();
    };
    document.getElementById('pdfZoomReset').onclick = function () {
      fitMode = false; zoom = 1; applyZoom();
    };
    document.getElementById('pdfZoomFit').onclick = function () {
      fitMode = true; fitToWidth();
    };
    document.getElementById('pdfPrev').onclick = function () { loadPage(currentPage - 1); };
    document.getElementById('pdfNext').onclick = function () { loadPage(currentPage + 1); };
    pageInput.addEventListener('change', function () { loadPage(pageInput.value); });
    pageInput.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); loadPage(pageInput.value); }
    });
    view.addEventListener('wheel', function (e) {
      if (!e.ctrlKey) return;
      e.preventDefault();
      fitMode = false;
      zoom += e.deltaY < 0 ? 0.1 : -0.1;
      zoom = Math.min(Math.max(zoom, 0.1), 5);
      applyZoom();
    }, { passive: false });

    modal.addEventListener('hidden.bs.modal', function () {
      img.src = '';
      hideLoading();
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
