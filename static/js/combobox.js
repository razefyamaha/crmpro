/**
 * customer-combobox - a searchable dropdown for the Company/Customer fields.
 * Wraps any <input class="customer-combobox"> and lets the user:
 *   - type to filter the existing customer list
 *   - click (or arrow+Enter) to pick an existing customer
 *   - keep typing freely to enter a brand-new customer
 * The value is written straight into the wrapped input, so normal form
 * submission works unchanged.
 */
(function () {
  function escapeHtml(s) {
    return (s || "").replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }

  function setup(input) {
    if (input.dataset.cbInit) return;
    input.dataset.cbInit = "1";

    var names = window.ALL_CUSTOMERS || [];
    // If data-submit-on-pick is set (search/filter boxes), picking a suggestion
    // submits the surrounding form immediately.
    var submitOnPick = input.hasAttribute("data-submit-on-pick");
    var form = input.closest("form");

    // Build a wrapper for positioning the dropdown
    var wrap = document.createElement("div");
    wrap.className = "cb-wrap";
    input.parentNode.insertBefore(wrap, input);
    wrap.appendChild(input);

    // Dropdown list
    var dd = document.createElement("ul");
    dd.className = "cb-dropdown";
    dd.style.display = "none";
    wrap.appendChild(dd);

    var currentSel = -1;
    var items = [];

    function hide() {
      dd.style.display = "none";
      currentSel = -1;
    }

    function render(filter) {
      filter = (filter || "").toLowerCase().trim();
      var list = names.filter(function (n) {
        return n && (!filter || n.toLowerCase().indexOf(filter) !== -1);
      });
      list.sort();
      if (list.length > 200) list = list.slice(0, 200);
      dd.innerHTML = "";
      items = list;
      if (!list.length) {
        var li = document.createElement("li");
        li.className = "cb-empty";
        li.textContent = "No matches — keep typing to add a new customer";
        dd.appendChild(li);
      } else {
        list.forEach(function (n, i) {
          var li = document.createElement("li");
          li.textContent = n;
          li.addEventListener("mousedown", function (e) {
            e.preventDefault(); // avoid blur before click
            input.value = n;
            hide();
            if (submitOnPick && form) form.submit();
          });
          dd.appendChild(li);
        });
      }
      dd.style.display = "block";
    }

    function highlight() {
      Array.prototype.forEach.call(dd.children, function (li, i) {
        if (li.classList.contains("cb-empty")) return;
        if (i === currentSel) li.classList.add("active");
        else li.classList.remove("active");
      });
    }

    input.addEventListener("focus", function () { render(input.value); });
    input.addEventListener("input", function () { render(input.value); });
    input.addEventListener("keydown", function (e) {
      if (dd.style.display === "none") return;
      if (e.key === "ArrowDown") {
        e.preventDefault();
        currentSel = Math.min(currentSel + 1, items.length - 1);
        highlight();
      } else if (e.key === "ArrowUp") {
        e.preventDefault();
        currentSel = Math.max(currentSel - 1, 0);
        highlight();
      } else if (e.key === "Enter") {
        if (currentSel >= 0 && currentSel < items.length) {
          e.preventDefault();
          input.value = items[currentSel];
          hide();
          if (submitOnPick && form) form.submit();
        }
      } else if (e.key === "Escape") {
        hide();
      }
    });
    input.addEventListener("blur", function () {
      // let mousedown selection finish first
      setTimeout(hide, 120);
    });
  }

  function init() {
    document.querySelectorAll("input.customer-combobox").forEach(setup);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
