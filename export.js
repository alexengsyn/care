/* Branded exports — real .xlsx (no dependencies) + print-to-PDF view.
   Dependency-free so it also works inside the single-file standalone build. */
(function (global) {
  'use strict';

  var BRAND = {
    title: 'الشبكة الطبية بالإسكندرية 2026',
    org: 'نقابة المهندسين بالإسكندرية — لجنة الرعاية الصحية والمعاشات',
    credit: 'إعداد وتطوير: م. يوسف أسامة — تطوعاً لنقابتنا الحبيبة',
    note: 'المصدر: كتيب الشبكة الطبية 2026. الخصومات تُصرف بالكارنيه الساري للمهندس وذويه.'
  };

  /* ------------------------------------------------------------ CRC32 */
  var CRC = (function () {
    var t = new Uint32Array(256), c, i, k;
    for (i = 0; i < 256; i++) {
      c = i;
      for (k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[i] = c >>> 0;
    }
    return t;
  })();

  function crc32(buf) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < buf.length; i++) c = CRC[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  var enc = new TextEncoder();

  /* -------------------------------------------- minimal STORE-only ZIP */
  function zip(files) {
    var parts = [], central = [], offset = 0;

    files.forEach(function (f) {
      var name = enc.encode(f.name),
          data = typeof f.data === 'string' ? enc.encode(f.data) : f.data,
          crc = crc32(data);

      var lh = new Uint8Array(30 + name.length), dv = new DataView(lh.buffer);
      dv.setUint32(0, 0x04034b50, true);
      dv.setUint16(4, 20, true);          // version needed
      dv.setUint16(6, 0x0800, true);      // UTF-8 filename flag
      dv.setUint16(8, 0, true);           // STORE
      dv.setUint32(14, crc, true);
      dv.setUint32(18, data.length, true);
      dv.setUint32(22, data.length, true);
      dv.setUint16(26, name.length, true);
      lh.set(name, 30);
      parts.push(lh, data);

      var cd = new Uint8Array(46 + name.length), cv = new DataView(cd.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true);
      cv.setUint16(6, 20, true);
      cv.setUint16(8, 0x0800, true);
      cv.setUint16(10, 0, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, data.length, true);
      cv.setUint32(24, data.length, true);
      cv.setUint16(28, name.length, true);
      cv.setUint32(42, offset, true);
      cd.set(name, 46);
      central.push(cd);

      offset += lh.length + data.length;
    });

    var cdSize = central.reduce(function (a, b) { return a + b.length; }, 0);
    var eocd = new Uint8Array(22), ev = new DataView(eocd.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(8, files.length, true);
    ev.setUint16(10, files.length, true);
    ev.setUint32(12, cdSize, true);
    ev.setUint32(16, offset, true);

    return new Blob(parts.concat(central, [eocd]),
      { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  /* ------------------------------------------------------------ helpers */
  function xesc(s) {
    return String(s == null ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&apos;')
      // strip control chars Excel rejects
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '');
  }

  function colName(n) {                      // 1 -> A
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = (n - m - 1) / 26; }
    return s;
  }

  function download(blob, filename) {
    var url = URL.createObjectURL(blob),
        a = document.createElement('a');
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(url); a.remove(); }, 1500);
  }

  function stamp() {
    var d = new Date(), p = function (n) { return (n < 10 ? '0' : '') + n; };
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
  }

  /* ------------------------------------------------------------- XLSX */
  var HEAD = ['#', 'الاسم', 'القسم', 'التخصص', 'الدرجة العلمية', 'المنطقة',
              'الحي', 'كل المناطق', 'العنوان', 'التليفونات', 'صفحة الكتيب'];
  var WIDTHS = [5, 34, 14, 22, 14, 15, 12, 24, 55, 24, 8];

  function rowsFrom(records) {
    return records.map(function (r, i) {
      return [i + 1, r.name, r.group, r.section, r.degree, r.district, r.zone,
              (r.districts || []).join(' / '), r.address, (r.phones || []).join(' / '), r.page];
    });
  }

  // style ids: 0 normal | 1 title | 2 subtitle | 3 credit | 4 header | 5 body | 6 num
  var STYLES =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<fonts count="6">' +
      '<font><sz val="11"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="18"/><color rgb="FF19350C"/><name val="Calibri"/></font>' +
      '<font><sz val="11"/><color rgb="FF44553C"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="11"/><color rgb="FF550C0C"/><name val="Calibri"/></font>' +
      '<font><b/><sz val="12"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>' +
      '<font><sz val="11"/><color rgb="FF1C2118"/><name val="Calibri"/></font>' +
    '</fonts>' +
    '<fills count="3">' +
      '<fill><patternFill patternType="none"/></fill>' +
      '<fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FF19350C"/><bgColor indexed="64"/></patternFill></fill>' +
    '</fills>' +
    '<borders count="2"><border/>' +
      '<border><left style="thin"><color rgb="FFD8DFD2"/></left><right style="thin"><color rgb="FFD8DFD2"/></right>' +
      '<top style="thin"><color rgb="FFD8DFD2"/></top><bottom style="thin"><color rgb="FFD8DFD2"/></bottom></border>' +
    '</borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="7">' +
      '<xf xfId="0" numFmtId="0" fontId="0" fillId="0" borderId="0"/>' +
      '<xf xfId="0" numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"><alignment vertical="center"/></xf>' +
      '<xf xfId="0" numFmtId="0" fontId="2" fillId="0" borderId="0" applyFont="1"><alignment vertical="center"/></xf>' +
      '<xf xfId="0" numFmtId="0" fontId="3" fillId="0" borderId="0" applyFont="1"><alignment vertical="center"/></xf>' +
      '<xf xfId="0" numFmtId="0" fontId="4" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' +
        '<alignment horizontal="center" vertical="center" wrapText="1"/></xf>' +
      '<xf xfId="0" numFmtId="0" fontId="5" fillId="0" borderId="1" applyFont="1" applyBorder="1" applyAlignment="1">' +
        '<alignment vertical="top" wrapText="1"/></xf>' +
      '<xf xfId="0" numFmtId="0" fontId="5" fillId="0" borderId="1" applyFont="1" applyBorder="1" applyAlignment="1">' +
        '<alignment horizontal="center" vertical="top"/></xf>' +
    '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    '<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>' +
    '</styleSheet>';

  function sheetXml(records, subtitle) {
    var rows = rowsFrom(records), out = [], rn = 0, last = colName(HEAD.length), lastDataRow = 0;

    function cell(ref, val, style) {
      if (typeof val === 'number' && isFinite(val)) {
        return '<c r="' + ref + '" s="' + style + '"><v>' + val + '</v></c>';
      }
      var t = xesc(val);
      if (t === '') return '<c r="' + ref + '" s="' + style + '"/>';
      return '<c r="' + ref + '" s="' + style + '" t="inlineStr"><is><t xml:space="preserve">' + t + '</t></is></c>';
    }

    function banner(text, style) {
      rn++;
      out.push('<row r="' + rn + '" ht="' + (style === 1 ? 26 : 17) + '" customHeight="1">' +
               cell('A' + rn, text, style) + '</row>');
    }

    banner(BRAND.title, 1);
    banner(BRAND.org, 2);
    banner(subtitle + '  •  تاريخ التصدير: ' + stamp(), 2);
    rn++; out.push('<row r="' + rn + '"/>');            // spacer

    var headRow = ++rn, cells = '';
    HEAD.forEach(function (h, i) { cells += cell(colName(i + 1) + rn, h, 4); });
    out.push('<row r="' + rn + '" ht="24" customHeight="1">' + cells + '</row>');

    rows.forEach(function (r) {
      rn++;
      var s = '';
      r.forEach(function (v, i) {
        s += cell(colName(i + 1) + rn, v, (i === 0 || i === 10) ? 6 : 5);
      });
      out.push('<row r="' + rn + '">' + s + '</row>');
    });
    lastDataRow = rn;

    // footer credit — one line, after the data
    rn++; out.push('<row r="' + rn + '"/>');            // spacer
    rn++;
    out.push('<row r="' + rn + '" ht="18" customHeight="1">' +
             cell('A' + rn, BRAND.credit, 3) + '</row>');

    var cols = WIDTHS.map(function (w, i) {
      return '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + w + '" customWidth="1"/>';
    }).join('');

    return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<sheetPr><outlinePr summaryBelow="1" summaryRight="1"/></sheetPr>' +
      '<sheetViews><sheetView rightToLeft="1" tabSelected="1" workbookViewId="0">' +
      '<pane ySplit="' + headRow + '" topLeftCell="A' + (headRow + 1) + '" activePane="bottomLeft" state="frozen"/>' +
      '</sheetView></sheetViews>' +
      '<sheetFormatPr defaultRowHeight="15"/>' +
      '<cols>' + cols + '</cols>' +
      '<sheetData>' + out.join('') + '</sheetData>' +
      '<autoFilter ref="A' + headRow + ':' + last + lastDataRow + '"/>' +
      '<pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>' +
      '</worksheet>';
  }

  function toXlsx(records, subtitle, filename) {
    var files = [
      { name: '[Content_Types].xml', data:
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
        '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
        '</Types>' },
      { name: '_rels/.rels', data:
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
        '</Relationships>' },
      { name: 'xl/workbook.xml', data:
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
        '<sheets><sheet name="الشبكة الطبية 2026" sheetId="1" r:id="rId1"/></sheets></workbook>' },
      { name: 'xl/_rels/workbook.xml.rels', data:
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
        '</Relationships>' },
      { name: 'xl/styles.xml', data: STYLES },
      { name: 'xl/worksheets/sheet1.xml', data: sheetXml(records, subtitle) }
    ];
    download(zip(files), filename || ('الشبكة-الطبية-2026-' + stamp() + '.xlsx'));
  }

  /* --------------------------------------------------------------- CSV */
  function toCsv(records, subtitle, filename) {
    var q = function (v) { return '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"'; };
    var lines = [
      [BRAND.title], [BRAND.org],
      [subtitle + ' — تاريخ التصدير: ' + stamp()], [], HEAD
    ].map(function (r) { return r.map(q).join(','); });
    rowsFrom(records).forEach(function (r) { lines.push(r.map(q).join(',')); });
    lines.push('');
    lines.push(q(BRAND.credit));
    download(new Blob(['\uFEFF' + lines.join('\r\n')], { type: 'text/csv;charset=utf-8' }),
             filename || ('الشبكة-الطبية-2026-' + stamp() + '.csv'));
  }

  /* --------------------------------------------------------------- PDF
     Opens a branded, print-optimised window; the browser's own
     "Save as PDF" handles Arabic shaping perfectly (no font embedding). */
  function toPdf(records, subtitle, logoSrc) {
    var esc = function (s) {
      return String(s == null ? '' : s)
        .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    };

    var body = records.map(function (r, i) {
      return '<tr>' +
        '<td class="n">' + (i + 1) + '</td>' +
        '<td><b>' + esc(r.name) + '</b>' +
          (r.degree ? '<span class="dg">' + esc(r.degree) + '</span>' : '') + '</td>' +
        '<td>' + esc(r.section) + '</td>' +
        '<td>' + esc(r.district || '') + '</td>' +
        '<td>' + esc(r.address) + '</td>' +
        '<td class="tel">' + esc((r.phones || []).join(' / ')) + '</td>' +
      '</tr>';
    }).join('');

    var html =
'<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="utf-8">' +
'<title>' + BRAND.title + '</title>' +
'<link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;900&display=swap" rel="stylesheet">' +
'<style>' +
'@page{size:A4 landscape;margin:11mm 9mm}' +
'*{box-sizing:border-box}' +
'body{font-family:"Cairo",Tahoma,Arial,sans-serif;margin:0;color:#1c2118;font-size:10.5px}' +
'.hd{display:flex;align-items:center;gap:12px;border-bottom:3px solid #19350C;padding-bottom:9px;margin-bottom:6px}' +
'.hd img{width:48px;height:48px;object-fit:contain}' +
'.hd h1{margin:0;font-size:19px;color:#19350C;font-weight:900}' +
'.hd p{margin:2px 0 0;font-size:11px;color:#55603f}' +
'.hd .meta{margin-inline-start:auto;text-align:left;font-size:10px;color:#6b7565;line-height:1.7}' +
'.hd .meta b{color:#19350C;font-size:13px}' +
'.note{background:#eef5e9;border-inline-start:4px solid #88DA62;padding:5px 10px;margin-bottom:9px;' +
  'font-size:10.5px;font-weight:600;color:#4b5745}' +
'table{width:100%;border-collapse:collapse}' +
'thead{display:table-header-group}' +
'th{background:#19350C;color:#fff;font-size:10.5px;padding:6px 5px;text-align:right;font-weight:700}' +
'td{border:1px solid #dde3d7;padding:5px;vertical-align:top;line-height:1.55}' +
'tr:nth-child(even) td{background:#f8faf6}' +
'tr{break-inside:avoid}' +
'td.n{text-align:center;color:#7d8a74;width:26px}' +
'td.tel{direction:ltr;text-align:left;white-space:nowrap;font-weight:700;width:105px}' +
'.dg{display:block;font-size:9px;color:#550C0C;font-weight:600}' +
'.ft{margin-top:10px;padding-top:7px;border-top:1px solid #dde3d7;font-size:9.5px;color:#6b7565;' +
  'display:flex;justify-content:space-between;gap:12px}' +
'.bar{position:fixed;top:0;right:0;left:0;background:#19350C;color:#fff;padding:9px 16px;' +
  'display:flex;gap:10px;align-items:center;font-family:Cairo,Tahoma,sans-serif;z-index:9}' +
'.bar button{font:inherit;font-weight:700;font-size:13px;border:0;border-radius:8px;padding:8px 20px;cursor:pointer}' +
'.bar .go{background:#88DA62;color:#19350C}.bar .cl{background:#3c5a30;color:#fff}' +
'.bar span{font-size:12.5px;opacity:.9}' +
'@media print{.bar{display:none}body{padding-top:0}}' +
'@media screen{body{padding:76px 18px 28px;background:#eef1ea}' +
  '.sheet{background:#fff;max-width:1180px;margin:0 auto;padding:22px;box-shadow:0 6px 26px rgba(0,0,0,.13)}}' +
'</style></head><body>' +
'<div class="bar">' +
  '<button class="go" onclick="window.print()">🖨️ حفظ كـ PDF / طباعة</button>' +
  '<button class="cl" onclick="window.close()">إغلاق</button>' +
  '<span>في نافذة الطباعة اختر «الوجهة → حفظ كـ PDF»</span>' +
'</div><div class="sheet">' +
'<div class="hd">' +
  (logoSrc ? '<img src="' + logoSrc + '" alt="">' : '') +
  '<div><h1>' + BRAND.title + '</h1><p>' + BRAND.org + '</p></div>' +
  '<div class="meta"><b>' + records.length + '</b> جهة<br>' + esc(subtitle) + '<br>' + stamp() + '</div>' +
'</div>' +
'<div class="note">' + BRAND.note + '</div>' +
'<table><thead><tr><th>#</th><th>الاسم</th><th>التخصص</th><th>المنطقة</th><th>العنوان</th><th>التليفون</th></tr></thead>' +
'<tbody>' + body + '</tbody></table>' +
'<div class="ft"><span>' + BRAND.org + '</span><span>' + BRAND.credit + '</span></div>' +
'</div></body></html>';

    var w = window.open('', '_blank');
    if (!w) { alert('يرجى السماح بالنوافذ المنبثقة لتصدير PDF'); return; }
    w.document.write(html);
    w.document.close();
  }

  global.Exporter = { xlsx: toXlsx, csv: toCsv, pdf: toPdf, brand: BRAND, stamp: stamp };
})(window);
