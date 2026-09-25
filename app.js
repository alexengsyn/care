/* Alexandria medical network — client-side search */
(function () {
  'use strict';

  var DATA = [], view = [], shown = 0, PAGE = 60;
  var q = document.getElementById('q'),
      clearBtn = document.getElementById('clear'),
      selGroup = document.getElementById('group'),
      selSection = document.getElementById('section'),
      selZone = document.getElementById('zone'),
      selDistrict = document.getElementById('district'),
      selSort = document.getElementById('sort'),
      results = document.getElementById('results'),
      count = document.getElementById('count'),
      empty = document.getElementById('empty'),
      more = document.getElementById('more');

  /* normalise Arabic for tolerant matching (alef/ya/ta-marbuta/diacritics) */
  function norm(s) {
    return (s || '')
      .replace(/[\u064b-\u0652\u0670\u0640]/g, '')
      .replace(/[أإآٱ]/g, 'ا')
      .replace(/ى/g, 'ي')
      .replace(/ة/g, 'ه')
      .replace(/[ؤئ]/g, function (m) { return m === 'ؤ' ? 'و' : 'ي'; })
      .replace(/[٠-٩]/g, function (d) { return '٠١٢٣٤٥٦٧٨٩'.indexOf(d); })
      .replace(/[^\w\s\u0600-\u06ff]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .toLowerCase();
  }

  function esc(s) {
    return (s || '').replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }

  /* highlight the matched terms inside the original (un-normalised) text */
  function hi(text, terms) {
    var out = esc(text);
    if (!terms.length) return out;
    var n = norm(text);
    if (!n) return out;
    // map: highlight by locating each term in the normalised string, then
    // approximate back by splitting on whitespace-aligned words.
    var words = out.split(/(\s+)/);
    return words.map(function (w) {
      var nw = norm(w);
      if (!nw) return w;
      for (var i = 0; i < terms.length; i++) {
        if (terms[i] && nw.indexOf(terms[i]) !== -1) return '<mark>' + w + '</mark>';
      }
      return w;
    }).join('');
  }

  var ICON = {
    pin: '<svg viewBox="0 0 24 24"><path d="M12 2a7 7 0 00-7 7c0 5.25 7 13 7 13s7-7.75 7-13a7 7 0 00-7-7zm0 9.5A2.5 2.5 0 1112 6.5a2.5 2.5 0 010 5z"/></svg>',
    tel: '<svg viewBox="0 0 24 24"><path d="M6.6 10.8a15.1 15.1 0 006.6 6.6l2.2-2.2a1 1 0 011-.25 11.4 11.4 0 003.6.57 1 1 0 011 1V20a1 1 0 01-1 1A17 17 0 013 4a1 1 0 011-1h3.5a1 1 0 011 1c0 1.25.2 2.46.57 3.6a1 1 0 01-.25 1z"/></svg>'
  };

  function card(r, terms) {
    var tels = (r.phones || []).map(function (p) {
      var digits = p.replace(/[^\d+]/g, '');
      return '<a class="tel" href="tel:' + esc(digits) + '">' + ICON.tel + '<span>' + esc(p) + '</span></a>';
    }).join('');

    return '<article class="card">' +
      '<h3>' + hi(r.name, terms) + '</h3>' +
      '<div class="tags">' +
        '<span class="tag">' + esc(r.section) + '</span>' +
        (r.degree ? '<span class="tag deg">' + esc(r.degree) + '</span>' : '') +
      '</div>' +
      (r.district && r.district !== 'غير محدد'
        ? '<div class="row"><span class="pill" data-d="' + esc(r.district) + '">' +
          ICON.pin + esc(r.district) +
          (r.zone && r.zone !== 'غير محدد' ? '<i>' + esc(r.zone) + '</i>' : '') +
          '</span></div>'
        : '') +
      (r.address ? '<div class="row">' + ICON.pin + '<span>' + hi(r.address, terms) + '</span></div>' : '') +
      (tels ? '<div class="tels">' + tels + '</div>' : '') +
    '</article>';
  }

  function currentTerms() {
    return norm(q.value).split(' ').filter(Boolean);
  }

  function apply() {
    var terms = currentTerms(),
        g = selGroup.value,
        s = selSection.value,
        z = selZone.value,
        dd = selDistrict.value;

    view = DATA.filter(function (r) {
      if (g && r.group !== g) return false;
      if (s && r.section !== s) return false;
      if (z && r.zone !== z) return false;
      if (dd && (r.districts || []).indexOf(dd) === -1) return false;
      for (var i = 0; i < terms.length; i++) {
        if (r.q.indexOf(terms[i]) === -1) return false;
      }
      return true;
    });

    var by = selSort.value;
    if (by === 'name') view.sort(function (a, b) { return a.name.localeCompare(b.name, 'ar'); });
    else if (by === 'section') view.sort(function (a, b) {
      return a.section.localeCompare(b.section, 'ar') || a.name.localeCompare(b.name, 'ar');
    });
    else if (by === 'district') view.sort(function (a, b) {
      return a.district.localeCompare(b.district, 'ar') || a.name.localeCompare(b.name, 'ar');
    });
    else view.sort(function (a, b) { return a.id - b.id; });

    shown = 0;
    results.innerHTML = '';
    empty.hidden = view.length > 0;
    count.innerHTML = view.length
      ? 'عدد النتائج: <b>' + view.length + '</b> من إجمالي ' + DATA.length + ' جهة'
      : '';
    clearBtn.style.display = q.value ? 'block' : 'none';
    ['dl-pdf', 'dl-xlsx'].forEach(function (id) {
      var b = document.getElementById(id);
      if (b) b.disabled = view.length === 0;
    });
    render();
    fillSections();
    fillDistricts();
  }

  function render() {
    var terms = currentTerms(),
        slice = view.slice(shown, shown + PAGE);
    results.insertAdjacentHTML('beforeend', slice.map(function (r) { return card(r, terms); }).join(''));
    shown += slice.length;
    more.hidden = shown >= view.length;
    more.textContent = 'عرض المزيد (' + (view.length - shown) + ')';
  }

  function fillSections() {
    var g = selGroup.value, keep = selSection.value, seen = {}, list = [];
    DATA.forEach(function (r) {
      if (g && r.group !== g) return;
      if (!seen[r.section]) { seen[r.section] = 0; list.push(r.section); }
      seen[r.section]++;
    });
    list.sort(function (a, b) { return a.localeCompare(b, 'ar'); });
    selSection.innerHTML = '<option value="">كل التخصصات</option>' +
      list.map(function (s) {
        return '<option value="' + esc(s) + '">' + esc(s) + ' (' + seen[s] + ')</option>';
      }).join('');
    if (list.indexOf(keep) !== -1) selSection.value = keep;
  }

  function fillDistricts() {
    var g = selGroup.value, sec = selSection.value, z = selZone.value,
        keep = selDistrict.value, seen = {}, list = [];
    DATA.forEach(function (r) {
      if (g && r.group !== g) return;
      if (sec && r.section !== sec) return;
      if (z && r.zone !== z) return;
      (r.districts && r.districts.length ? r.districts : [r.district]).forEach(function (d) {
        if (!d || d === 'غير محدد') return;
        if (!seen[d]) { seen[d] = 0; list.push(d); }
        seen[d]++;
      });
    });
    list.sort(function (a, b) { return seen[b] - seen[a] || a.localeCompare(b, 'ar'); });
    selDistrict.innerHTML = '<option value="">كل المناطق</option>' +
      list.map(function (d) {
        return '<option value="' + esc(d) + '">' + esc(d) + ' (' + seen[d] + ')</option>';
      }).join('');
    if (list.indexOf(keep) !== -1) selDistrict.value = keep;
  }

  function boot(data) {
    DATA = data;
    var groups = {}, order = [];
    DATA.forEach(function (r) {
      if (!groups[r.group]) { groups[r.group] = 0; order.push(r.group); }
      groups[r.group]++;
    });


    var zones = {}, zorder = [];
    DATA.forEach(function (r) {
      if (!r.zone || r.zone === 'غير محدد') return;
      if (!zones[r.zone]) { zones[r.zone] = 0; zorder.push(r.zone); }
      zones[r.zone]++;
    });
    zorder.sort(function (a, b) { return zones[b] - zones[a]; });
    selZone.innerHTML = '<option value="">كل الأحياء</option>' +
      zorder.map(function (z) {
        return '<option value="' + esc(z) + '">' + esc(z) + ' (' + zones[z] + ')</option>';
      }).join('');

    selGroup.innerHTML = '<option value="">كل الأقسام</option>' +
      order.map(function (g) { return '<option value="' + esc(g) + '">' + esc(g) + ' (' + groups[g] + ')</option>'; }).join('');

    // ---- branded exports of the currently filtered view
    function subtitle() {
      var bits = [];
      if (selGroup.value) bits.push(selGroup.value);
      if (selSection.value) bits.push(selSection.value);
      if (selZone.value) bits.push(selZone.value);
      if (selDistrict.value) bits.push(selDistrict.value);
      if (q.value.trim()) bits.push('بحث: ' + q.value.trim());
      return bits.length ? bits.join(' • ') : 'كل الجهات';
    }
    function slug() {
      var s = subtitle().replace(/[•:]/g, '-').replace(/\s+/g, '-');
      return 'الشبكة-الطبية-2026-' + s.slice(0, 60) + '-' + Exporter.stamp();
    }
    document.getElementById('dl-xlsx').addEventListener('click', function () {
      Exporter.xlsx(view, subtitle(), slug() + '.xlsx');
    });
    document.getElementById('dl-pdf').addEventListener('click', function () {
      var img = document.querySelector('.brand img');
      Exporter.pdf(view, subtitle(), img ? img.src : '');
    });

    var t;
    q.addEventListener('input', function () { clearTimeout(t); t = setTimeout(apply, 120); });
    clearBtn.addEventListener('click', function () { q.value = ''; q.focus(); apply(); });
    selGroup.addEventListener('change', function () { selSection.value = ''; apply(); });
    selSection.addEventListener('change', apply);
    selZone.addEventListener('change', function () { selDistrict.value = ''; apply(); });
    selDistrict.addEventListener('change', apply);

    // clicking the district pill on a card filters by it
    results.addEventListener('click', function (e) {
      var p = e.target.closest('.pill');
      if (!p) return;
      selDistrict.value = p.dataset.d;
      selZone.value = '';
      apply();
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
    selSort.addEventListener('change', apply);
    more.addEventListener('click', render);

    // deep link:  ?q=...&g=...
    var p = new URLSearchParams(location.search);
    if (p.get('q')) q.value = p.get('q');
    if (p.get('g')) selGroup.value = p.get('g');
    if (p.get('z')) selZone.value = p.get('z');
    if (p.get('d')) { fillDistricts(); selDistrict.value = p.get('d'); }

    apply();
  }

  fetch('data.json')
    .then(function (r) { return r.json(); })
    .then(boot)
    .catch(function () {
      count.textContent = 'تعذّر تحميل البيانات. شغّل الملفات عبر خادم محلي (وليس file://).';
    });
})();
