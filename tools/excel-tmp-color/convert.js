/* Excel 字体颜色 → Unity TextMeshPro <color> 标签 */
(function (root, factory) {
  var api = factory();
  if (typeof module === 'object' && module.exports) module.exports = api;
  root.TmpExcel = api;
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  // SpreadsheetML theme index is not the XML element order.
  // 0 lt1, 1 dk1, 2 lt2, 3 dk2, 4 accent1 ... 10 hlink, 11 folHlink
  var THEME_NAMES = [
    'lt1', 'dk1', 'lt2', 'dk2',
    'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6',
    'hlink', 'folHlink'
  ];

  var DEFAULT_THEME = [
    'FFFFFF', '000000', 'EEECE1', '1F497D',
    '4F81BD', 'C0504D', '9BBB59', '8064A2', '4BACC6', 'F79646',
    '0000FF', '800080'
  ];

  // BIFF8 / SpreadsheetML default indexed palette. Index 64 is automatic.
  var DEFAULT_INDEXED = [
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
    '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
    '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
    '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'
  ];

  var COLOR_TAG_SPLIT = /(<color\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)\s*>|<\/color\s*>)/gi;
  var COLOR_TAG_TEST = /<color\s*=/i;
  var COLOR_OPEN_CANON = /<color\s*=\s*(?:"[^"]*"|'[^']*'|[^\s>]+)\s*>/gi;

  function getLib(globalName, moduleName) {
    var fromGlobal = globalThis[globalName];
    if (fromGlobal) return fromGlobal;
    if (typeof require === 'function') {
      try { return require(moduleName); } catch (err) { /* fall through */ }
    }
    throw new Error('缺少 ' + moduleName + '，请确认脚本已加载');
  }

  function clamp01(n) {
    return Math.min(1, Math.max(0, n));
  }

  function rgbToHls(r, g, b) {
    var max = Math.max(r, g, b);
    var min = Math.min(r, g, b);
    var l = (max + min) / 2;
    if (max === min) return [0, l, 0];
    var d = max - min;
    var s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    var h;
    if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
    else if (max === g) h = (b - r) / d + 2;
    else h = (r - g) / d + 4;
    return [h / 6, l, s];
  }

  function hue2rgb(p, q, t) {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  }

  function hlsToRgb(h, l, s) {
    if (s === 0) return [l, l, l];
    var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    var p = 2 * l - q;
    return [hue2rgb(p, q, h + 1 / 3), hue2rgb(p, q, h), hue2rgb(p, q, h - 1 / 3)];
  }

  function channelToHex(unit) {
    return Math.round(clamp01(unit) * 255).toString(16).padStart(2, '0').toUpperCase();
  }

  // Excel tint: adjust HLS luminance. tint < 0 darkens, tint > 0 lightens.
  function applyTint(rgb, tint) {
    var hex = String(rgb || '').replace(/^#/, '').toUpperCase();
    if (!/^[0-9A-F]{6}$/.test(hex)) return hex;
    var t = Number(tint) || 0;
    if (!t) return hex;
    var r = parseInt(hex.slice(0, 2), 16) / 255;
    var g = parseInt(hex.slice(2, 4), 16) / 255;
    var b = parseInt(hex.slice(4, 6), 16) / 255;
    var hls = rgbToHls(r, g, b);
    var l = hls[1];
    if (t < 0) l = l * (1 + t);
    else l = l * (1 - t) + t;
    var rgbUnit = hlsToRgb(hls[0], clamp01(l), hls[2]);
    return channelToHex(rgbUnit[0]) + channelToHex(rgbUnit[1]) + channelToHex(rgbUnit[2]);
  }

  function parseThemeXml(xml) {
    if (!xml) return null;
    var schemeMatch = xml.match(/<(?:\w+:)?clrScheme\b[^>]*>([\s\S]*?)<\/(?:\w+:)?clrScheme>/i);
    var scope = schemeMatch ? schemeMatch[1] : xml;
    var theme = [];
    var found = 0;
    for (var i = 0; i < THEME_NAMES.length; i++) {
      var name = THEME_NAMES[i];
      var re = new RegExp('<(?:\\w+:)?' + name + '\\b[^>]*>([\\s\\S]*?)<\\/(?:\\w+:)?' + name + '>', 'i');
      var block = scope.match(re);
      var color = DEFAULT_THEME[i];
      if (block) {
        var inner = block[1];
        var srgb = inner.match(/<(?:\w+:)?srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/i);
        var last = inner.match(/\blastClr="([0-9A-Fa-f]{6})"/i);
        var sys = inner.match(/<(?:\w+:)?sysClr\b[^>]*\bval="([^"]+)"/i);
        if (srgb) color = srgb[1].toUpperCase();
        else if (last) color = last[1].toUpperCase();
        else if (sys && /windowText/i.test(sys[1])) color = '000000';
        else if (sys && /window/i.test(sys[1])) color = 'FFFFFF';
        found++;
      }
      theme.push(color);
    }
    return found ? theme : null;
  }

  function parseIndexedXml(xml) {
    if (!xml) return null;
    var block = xml.match(/<(?:\w+:)?indexedColors\b[^>]*>([\s\S]*?)<\/(?:\w+:)?indexedColors>/i);
    if (!block) return null;
    var colors = DEFAULT_INDEXED.slice();
    var re = /<(?:\w+:)?rgbColor\b[^>]*\brgb="([0-9A-Fa-f]{6,8})"/gi;
    var match;
    var count = 0;
    while ((match = re.exec(block[1]))) {
      var hex = match[1].toUpperCase();
      colors[count++] = hex.length === 8 ? hex.slice(2) : hex;
    }
    return count ? colors : null;
  }

  function toUint8Array(data) {
    if (data instanceof ArrayBuffer) return new Uint8Array(data.slice(0));
    if (ArrayBuffer.isView(data)) {
      var copy = new Uint8Array(data.byteLength);
      copy.set(new Uint8Array(data.buffer, data.byteOffset, data.byteLength));
      return copy;
    }
    throw new Error('无法读取文件内容');
  }

  async function readColorContext(bytes) {
    var zip = await getLib('JSZip', 'jszip').loadAsync(bytes);
    var theme = DEFAULT_THEME.slice();
    var indexed = DEFAULT_INDEXED.slice();
    var themeEntry = firstZip(zip, /xl\/theme\/theme\d+\.xml$/i);
    if (themeEntry) {
      var parsedTheme = parseThemeXml(await themeEntry.async('string'));
      if (parsedTheme) theme = parsedTheme;
    }
    var stylesEntry = firstZip(zip, /^xl\/styles\.xml$/i);
    if (stylesEntry) {
      var parsedIndexed = parseIndexedXml(await stylesEntry.async('string'));
      if (parsedIndexed) indexed = parsedIndexed;
    }
    return { theme: theme, indexed: indexed };
  }

  function firstZip(zip, pattern) {
    var found = zip.file(pattern);
    if (!found || !found.length) return null;
    return found[0];
  }

  function hasColorTag(text) {
    return COLOR_TAG_TEST.test(text);
  }

  function canonColorTags(text) {
    return String(text).replace(new RegExp(COLOR_OPEN_CANON.source, 'gi'), function (tag) {
      var hex = tag.match(/#([0-9A-Fa-f]{6,8})/);
      if (!hex) return tag.toLowerCase();
      return '<color=#' + hex[1].toUpperCase() + '>';
    });
  }

  function applyTmpColor(text, hex) {
    var source = text == null ? '' : String(text);
    if (!source) return '';
    var upper = String(hex).replace(/^#/, '').toUpperCase();
    var open = '<color=#' + upper + '>';
    var close = '</color>';
    var parts = source.split(new RegExp(COLOR_TAG_SPLIT.source, 'gi'));
    var depth = 0;
    var out = '';
    for (var i = 0; i < parts.length; i++) {
      var part = parts[i];
      if (!part) continue;
      if (/^<color\b/i.test(part)) {
        out += open;
        depth++;
      } else if (/^<\/color\b/i.test(part)) {
        if (depth > 0) depth--;
        out += close;
      } else if (depth > 0) {
        out += part;
      } else {
        out += open + part + close;
      }
    }
    while (depth > 0) {
      out += close;
      depth--;
    }
    return out;
  }

  function resolveColor(color, ctx, options) {
    if (!color || typeof color !== 'object') return null;
    var indexed = color.indexed == null ? null : Number(color.indexed);
    // 64 = automatic foreground. It is not an explicit font color.
    if (indexed === 64) return null;

    var themeIndex = color.theme == null ? null : Number(color.theme);
    var tint = color.tint ? Number(color.tint) : 0;
    var hasArgb = !!color.argb;

    // Theme 1 is Text 1 / dk1, Excel's default font color, even when a style
    // such as a number format materializes it onto the cell.
    if (!options.includeBlack && themeIndex === 1 && !tint && indexed == null && !hasArgb) {
      return null;
    }

    var rgb = null;
    var alpha = 'FF';
    if (themeIndex != null && ctx.theme[themeIndex]) {
      rgb = ctx.theme[themeIndex];
    } else if (hasArgb) {
      var raw = String(color.argb).replace(/^#/, '').toUpperCase();
      if (raw.length === 8) {
        alpha = raw.slice(0, 2);
        rgb = raw.slice(2);
      } else if (raw.length === 6) {
        rgb = raw;
      }
    } else if (indexed === 65) {
      rgb = 'FFFFFF';
    } else if (indexed != null) {
      rgb = ctx.indexed[indexed] || null;
    }

    if (!rgb || !/^[0-9A-F]{6}$/.test(rgb)) return null;
    if (tint) rgb = applyTint(rgb, tint);
    // Excel sometimes stores opaque colors with alpha 00.
    if (alpha === '00') alpha = 'FF';
    if (!/^[0-9A-F]{2}$/.test(alpha)) alpha = 'FF';
    if (!options.includeBlack && rgb === '000000' && alpha === 'FF') return null;

    var tmp = (!options.keepAlpha || alpha === 'FF') ? rgb : rgb + alpha;
    return { rgb: rgb, alpha: alpha, tmp: tmp };
  }

  function cellFontColor(cell) {
    return cell && cell.font && cell.font.color ? cell.font.color : null;
  }

  function formatExcelDate(value) {
    var y = value.getUTCFullYear();
    var m = String(value.getUTCMonth() + 1).padStart(2, '0');
    var d = String(value.getUTCDate()).padStart(2, '0');
    if (value.getUTCHours() || value.getUTCMinutes() || value.getUTCSeconds()) {
      return y + '-' + m + '-' + d + ' ' +
        String(value.getUTCHours()).padStart(2, '0') + ':' +
        String(value.getUTCMinutes()).padStart(2, '0') + ':' +
        String(value.getUTCSeconds()).padStart(2, '0');
    }
    return y + '-' + m + '-' + d;
  }

  function piecesFromRichText(runs, fallbackColor) {
    return runs.map(function (run) {
      var runColor = run.font && run.font.color ? run.font.color : fallbackColor;
      return { text: run.text == null ? '' : String(run.text), color: runColor };
    });
  }

  function getPieces(cell) {
    var value = cell.value;
    var fallback = cellFontColor(cell);
    var meta = { formula: false, rich: false };
    if (value && typeof value === 'object') {
      if (Array.isArray(value.richText)) {
        meta.rich = true;
        return { pieces: piecesFromRichText(value.richText, fallback), meta: meta };
      }
      if (Object.prototype.hasOwnProperty.call(value, 'formula') ||
          Object.prototype.hasOwnProperty.call(value, 'sharedFormula')) {
        meta.formula = true;
        var result = value.result;
        if (result && typeof result === 'object' && Array.isArray(result.richText)) {
          meta.rich = true;
          return { pieces: piecesFromRichText(result.richText, fallback), meta: meta };
        }
        if (result instanceof Date) {
          return { pieces: [{ text: formatExcelDate(result), color: fallback }], meta: meta };
        }
        var formulaText = '';
        if (result && typeof result === 'object' && result.error) formulaText = String(result.error);
        else if (cell.text != null && cell.text !== '') formulaText = String(cell.text);
        else if (result != null) formulaText = String(result);
        return { pieces: [{ text: formulaText, color: fallback }], meta: meta };
      }
      if (value instanceof Date) {
        return { pieces: [{ text: formatExcelDate(value), color: fallback }], meta: meta };
      }
      if (typeof value.text === 'string') {
        return { pieces: [{ text: value.text, color: fallback }], meta: meta };
      }
      if (value.error) {
        return { pieces: [{ text: String(cell.text || value.error), color: fallback }], meta: meta };
      }
    }
    if (value == null) return { pieces: [{ text: '', color: fallback }], meta: meta };
    var text = cell.text != null && cell.text !== '' ? String(cell.text) : String(value);
    return { pieces: [{ text: text, color: fallback }], meta: meta };
  }

  function convertPieces(pieces, ctx, options) {
    var resolved = pieces.map(function (piece) {
      return {
        text: piece.text,
        color: resolveColor(piece.color, ctx, options)
      };
    });
    var coalesced = [];
    for (var i = 0; i < resolved.length; i++) {
      var piece = resolved[i];
      var prev = coalesced[coalesced.length - 1];
      var canJoin = !!(piece.color && !hasColorTag(piece.text));
      var prevCanJoin = !!(prev && prev.color && piece.color && !hasColorTag(prev.text) && prev.color.tmp === piece.color.tmp);
      if (canJoin && prevCanJoin) prev.text += piece.text;
      else coalesced.push({ text: piece.text, color: piece.color });
    }
    var colors = [];
    var tagged = coalesced.map(function (piece) {
      if (!piece.color || piece.text === '') return piece.text;
      colors.push(piece.color.tmp);
      return applyTmpColor(piece.text, piece.color.tmp);
    }).join('');
    return { tagged: tagged, colors: unique(colors) };
  }

  function unique(list) {
    var seen = Object.create(null);
    var out = [];
    for (var i = 0; i < list.length; i++) {
      if (seen[list[i]]) continue;
      seen[list[i]] = true;
      out.push(list[i]);
    }
    return out;
  }

  function convertWorkbook(workbook, ctx, options) {
    var opts = normalizeOptions(options);
    var cells = [];
    var scanned = 0;
    var formulaCount = 0;
    workbook.eachSheet(function (sheet) {
      sheet.eachRow({ includeEmpty: false }, function (row) {
        row.eachCell({ includeEmpty: false }, function (cell) {
          var extracted = getPieces(cell);
          var original = extracted.pieces.map(function (piece) { return piece.text; }).join('');
          if (!original) return;
          scanned++;
          var converted = convertPieces(extracted.pieces, ctx, opts);
          if (!converted.tagged || canonColorTags(converted.tagged) === canonColorTags(original)) return;
          if (extracted.meta.formula) formulaCount++;
          cells.push({
            sheet: sheet.name,
            address: cell.address,
            original: original,
            tagged: converted.tagged,
            colors: converted.colors,
            rich: extracted.meta.rich,
            formula: extracted.meta.formula,
            hidden: sheet.state === 'hidden' || sheet.state === 'veryHidden'
          });
        });
      });
    });
    return {
      cells: cells,
      scanned: scanned,
      formulaCount: formulaCount,
      sheetNames: workbook.worksheets.map(function (sheet) { return sheet.name; }),
      sheetCount: workbook.worksheets.length
    };
  }

  function normalizeOptions(options) {
    var source = options || {};
    return {
      includeBlack: !!source.includeBlack,
      keepAlpha: source.keepAlpha !== false
    };
  }

  function isZip(data) {
    var bytes = toUint8Array(data);
    return bytes.length >= 2 && bytes[0] === 0x50 && bytes[1] === 0x4B;
  }

  function isOleCompound(data) {
    var bytes = toUint8Array(data);
    return bytes.length >= 4 && bytes[0] === 0xD0 && bytes[1] === 0xCF && bytes[2] === 0x11 && bytes[3] === 0xE0;
  }

  async function convertArrayBuffer(data, options) {
    if (isOleCompound(data)) {
      throw new Error('这是旧版 .xls 文件。请在 Excel 中另存为 .xlsx 后再拖进来。');
    }
    if (!isZip(data)) {
      throw new Error('无法识别这个文件。请使用 .xlsx 或 .xlsm。');
    }
    var bytes = toUint8Array(data);
    var ctx = await readColorContext(bytes.slice());
    var workbook = new (getLib('ExcelJS', 'exceljs')).Workbook();
    await workbook.xlsx.load(bytes.slice());
    return convertWorkbook(workbook, ctx, options);
  }

  function applyChanges(workbook, cells) {
    var count = 0;
    (cells || []).forEach(function (item) {
      var sheet = workbook.getWorksheet(item.sheet);
      if (!sheet) return;
      sheet.getCell(item.address).value = item.tagged;
      count++;
    });
    return count;
  }

  return {
    applyTint: applyTint,
    applyTmpColor: applyTmpColor,
    parseThemeXml: parseThemeXml,
    parseIndexedXml: parseIndexedXml,
    canonColorTags: canonColorTags,
    convertArrayBuffer: convertArrayBuffer,
    convertWorkbook: convertWorkbook,
    applyChanges: applyChanges,
    DEFAULT_THEME: DEFAULT_THEME,
    DEFAULT_INDEXED: DEFAULT_INDEXED
  };
});
