(function () {
  'use strict';

  var state = {
    buffer: null,
    fileName: '',
    result: null,
    sheet: '全部',
    query: '',
    limit: 200
  };

  var drop = document.getElementById('drop');
  var fileInput = document.getElementById('file');
  var sampleButton = document.getElementById('sample');
  var includeBlack = document.getElementById('includeBlack');
  var keepAlpha = document.getElementById('keepAlpha');
  var errorBox = document.getElementById('error');
  var loading = document.getElementById('loading');
  var results = document.getElementById('results');
  var fileName = document.getElementById('fileName');
  var summary = document.getElementById('summary');
  var warn = document.getElementById('warn');
  var tabs = document.getElementById('tabs');
  var search = document.getElementById('search');
  var empty = document.getElementById('empty');
  var tableWrap = document.getElementById('tableWrap');
  var rows = document.getElementById('rows');
  var toast = document.getElementById('toast');
  var toastTimer = 0;

  fileInput.addEventListener('change', function () {
    if (fileInput.files && fileInput.files[0]) readFile(fileInput.files[0]);
  });

  ['dragenter', 'dragover'].forEach(function (name) {
    drop.addEventListener(name, function (event) {
      event.preventDefault();
      drop.classList.add('is-over');
    });
  });
  ['dragleave', 'drop'].forEach(function (name) {
    drop.addEventListener(name, function (event) {
      event.preventDefault();
      drop.classList.remove('is-over');
    });
  });
  drop.addEventListener('drop', function (event) {
    var files = event.dataTransfer && event.dataTransfer.files;
    if (files && files[0]) readFile(files[0]);
  });

  sampleButton.addEventListener('click', function (event) {
    event.preventDefault();
    event.stopPropagation();
    loadSample();
  });
  includeBlack.addEventListener('change', rerun);
  keepAlpha.addEventListener('change', rerun);
  search.addEventListener('input', function () {
    state.query = search.value.trim().toLowerCase();
    state.limit = 200;
    renderRows();
  });
  document.getElementById('copyAll').addEventListener('click', copyAll);
  document.getElementById('downloadXlsx').addEventListener('click', downloadXlsx);
  document.getElementById('downloadCsv').addEventListener('click', downloadCsv);
  rows.addEventListener('click', function (event) {
    var button = event.target.closest('button[data-i]');
    if (!button || !state.result) return;
    var item = visibleCells()[Number(button.getAttribute('data-i'))];
    if (item) copyText(item.tagged);
  });

  function options() {
    return {
      includeBlack: includeBlack.checked,
      keepAlpha: keepAlpha.checked
    };
  }

  function readFile(file) {
    var reader = new FileReader();
    reader.onload = function () {
      state.fileName = file.name;
      convertBuffer(reader.result);
    };
    reader.onerror = function () {
      showError('读取文件失败。');
    };
    reader.readAsArrayBuffer(file);
  }

  function rerun() {
    if (state.buffer) convertBuffer(state.buffer);
  }

  async function convertBuffer(buffer) {
    showError('');
    loading.hidden = false;
    try {
      if (typeof TmpExcel === 'undefined' || typeof ExcelJS === 'undefined') {
        throw new Error('转换脚本没有加载完成。请确认 vendor 和 convert.js 与本页在同一目录。');
      }
      state.buffer = buffer;
      state.result = await TmpExcel.convertArrayBuffer(buffer, options());
      state.sheet = '全部';
      state.query = '';
      state.limit = 200;
      search.value = '';
      render();
    } catch (err) {
      results.hidden = true;
      showError(err && err.message ? err.message : '转换失败。');
    } finally {
      loading.hidden = true;
    }
  }

  function render() {
    var result = state.result;
    results.hidden = false;
    fileName.textContent = state.fileName || '表格';
    summary.textContent = '扫描 ' + result.scanned + ' 个有内容的单元格，转换 ' + result.cells.length + ' 处。';
    if (result.formulaCount) {
      warn.hidden = false;
      warn.textContent = '有 ' + result.formulaCount + ' 个公式单元格被写成了带标签的计算结果。下载的 Excel 里，这些格子不再保留公式。';
    } else {
      warn.hidden = true;
    }
    renderTabs();
    renderRows();
  }

  function renderTabs() {
    tabs.textContent = '';
    var names = ['全部'].concat(state.result.sheetNames);
    names.forEach(function (name) {
      var count = name === '全部'
        ? state.result.cells.length
        : state.result.cells.filter(function (cell) { return cell.sheet === name; }).length;
      var button = document.createElement('button');
      button.type = 'button';
      button.className = 'tab';
      button.setAttribute('aria-pressed', name === state.sheet ? 'true' : 'false');
      button.textContent = name + ' · ' + count;
      button.addEventListener('click', function () {
        state.sheet = name;
        state.limit = 200;
        renderTabs();
        renderRows();
      });
      tabs.appendChild(button);
    });
  }

  function visibleCells() {
    var query = state.query;
    return state.result.cells.filter(function (cell) {
      if (state.sheet !== '全部' && cell.sheet !== state.sheet) return false;
      if (!query) return true;
      var hay = (cell.sheet + ' ' + cell.address + ' ' + cell.original + ' ' + cell.tagged).toLowerCase();
      return hay.indexOf(query) !== -1;
    });
  }

  function renderRows() {
    var cells = visibleCells();
    rows.textContent = '';
    var hasRows = cells.length > 0;
    empty.hidden = hasRows;
    tableWrap.hidden = !hasRows;
    document.getElementById('copyAll').disabled = !hasRows;
    document.getElementById('downloadXlsx').disabled = !state.result.cells.length;
    document.getElementById('downloadCsv').disabled = !state.result.cells.length;
    if (!hasRows) return;

    var shown = cells.slice(0, state.limit);
    var fragment = document.createDocumentFragment();
    shown.forEach(function (cell, index) {
      fragment.appendChild(renderRow(cell, index));
    });
    if (cells.length > shown.length) {
      var more = document.createElement('tr');
      var td = document.createElement('td');
      td.colSpan = 4;
      td.className = 'more-row';
      var button = document.createElement('button');
      button.type = 'button';
      button.className = 'ghost';
      button.textContent = '还有 ' + (cells.length - shown.length) + ' 处，继续显示';
      button.addEventListener('click', function () {
        state.limit += 200;
        renderRows();
      });
      td.appendChild(button);
      more.appendChild(td);
      fragment.appendChild(more);
    }
    rows.appendChild(fragment);
  }

  function renderRow(cell, index) {
    var tr = document.createElement('tr');

    var where = document.createElement('td');
    var sheet = document.createElement('div');
    sheet.className = 'sheet-name';
    sheet.textContent = cell.sheet;
    if (cell.hidden) {
      var flag = document.createElement('span');
      flag.className = 'hidden-flag';
      flag.textContent = '隐藏表';
      sheet.appendChild(flag);
    }
    var addr = document.createElement('div');
    addr.className = 'addr';
    addr.textContent = cell.address;
    where.appendChild(sheet);
    where.appendChild(addr);

    var colorCell = document.createElement('td');
    var swatches = document.createElement('div');
    swatches.className = 'swatches';
    cell.colors.forEach(function (color) { swatches.appendChild(swatch(color)); });
    colorCell.appendChild(swatches);

    var original = document.createElement('td');
    var originalText = document.createElement('div');
    originalText.className = 'raw';
    originalText.textContent = cell.original;
    original.appendChild(originalText);

    var tagged = document.createElement('td');
    var raw = document.createElement('div');
    raw.className = 'raw';
    raw.textContent = cell.tagged;
    var preview = document.createElement('div');
    preview.className = 'preview';
    fillPreview(preview, cell.tagged);
    var actions = document.createElement('div');
    actions.className = 'cell-actions';
    var copy = document.createElement('button');
    copy.type = 'button';
    copy.className = 'mini';
    copy.setAttribute('data-i', String(index));
    copy.textContent = '复制这段';
    actions.appendChild(copy);
    tagged.appendChild(raw);
    tagged.appendChild(preview);
    tagged.appendChild(actions);

    tr.appendChild(where);
    tr.appendChild(colorCell);
    tr.appendChild(original);
    tr.appendChild(tagged);
    return tr;
  }

  function swatch(color) {
    var row = document.createElement('div');
    row.className = 'swatch';
    var dot = document.createElement('span');
    dot.className = 'dot';
    var rgb = color.slice(0, 6);
    var alpha = color.length === 8 ? parseInt(color.slice(6), 16) / 255 : 1;
    dot.style.background = '#' + rgb;
    dot.style.opacity = String(alpha);
    var label = document.createElement('span');
    label.textContent = '#' + color;
    row.appendChild(dot);
    row.appendChild(label);
    return row;
  }

  function fillPreview(container, text) {
    var re = /<color=#([0-9A-Fa-f]{6})([0-9A-Fa-f]{2})?>([\s\S]*?)<\/color>/gi;
    var last = 0;
    var match;
    while ((match = re.exec(text))) {
      if (match.index > last) container.appendChild(document.createTextNode(text.slice(last, match.index)));
      var span = document.createElement('span');
      span.style.color = '#' + match[1];
      if (match[2]) span.style.opacity = String(parseInt(match[2], 16) / 255);
      appendInlineMarks(span, match[3]);
      container.appendChild(span);
      last = match.index + match[0].length;
    }
    if (last < text.length) container.appendChild(document.createTextNode(text.slice(last)));
  }

  function appendInlineMarks(parent, text) {
    var parts = String(text).split(/(<\/?b\s*>|<\/?i\s*>)/gi);
    var stack = [parent];
    parts.forEach(function (part) {
      if (!part) return;
      var current = stack[stack.length - 1];
      if (/^<b\b/i.test(part)) {
        var strong = document.createElement('strong');
        current.appendChild(strong);
        stack.push(strong);
      } else if (/^<\/b\b/i.test(part)) {
        if (stack.length > 1) stack.pop();
      } else if (/^<i\b/i.test(part)) {
        var em = document.createElement('em');
        current.appendChild(em);
        stack.push(em);
      } else if (/^<\/i\b/i.test(part)) {
        if (stack.length > 1) stack.pop();
      } else {
        current.appendChild(document.createTextNode(part));
      }
    });
  }

  function filteredOrAll() {
    var cells = visibleCells();
    return cells.length ? cells : [];
  }

  function copyAll() {
    var cells = state.sheet === '全部' && !state.query ? state.result.cells : filteredOrAll();
    var text = cells.map(function (cell) {
      return [cell.sheet, cell.address, cell.tagged].join('\t');
    }).join('\n');
    copyText(text);
  }

  async function downloadXlsx() {
    try {
      var workbook = new ExcelJS.Workbook();
      var bytes = state.buffer instanceof ArrayBuffer ? state.buffer.slice(0) : state.buffer.slice();
      await workbook.xlsx.load(bytes);
      TmpExcel.applyChanges(workbook, state.result.cells);
      var out = await workbook.xlsx.writeBuffer();
      downloadBlob(out, baseName() + '.tmp.xlsx', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    } catch (err) {
      showError(err && err.message ? err.message : '导出 Excel 失败。');
    }
  }

  function downloadCsv() {
    var lines = ['工作表,单元格,颜色,原文,富文本'];
    state.result.cells.forEach(function (cell) {
      lines.push([
        csv(cell.sheet),
        csv(cell.address),
        csv(cell.colors.map(function (color) { return '#' + color; }).join(' ')),
        csv(cell.original),
        csv(cell.tagged)
      ].join(','));
    });
    downloadBlob('\uFEFF' + lines.join('\n'), baseName() + '.tmp.csv', 'text/csv;charset=utf-8');
  }

  function csv(value) {
    var text = value == null ? '' : String(value);
    if (/[",\n\r]/.test(text)) return '"' + text.replace(/"/g, '""') + '"';
    return text;
  }

  function baseName() {
    return String(state.fileName || '表格').replace(/\.[^.]+$/, '').replace(/[\\/:*?"<>|]+/g, '_') || '表格';
  }

  function downloadBlob(data, name, type) {
    var blob = new Blob([data], { type: type });
    var link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = name;
    document.body.appendChild(link);
    link.click();
    link.remove();
    setTimeout(function () { URL.revokeObjectURL(link.href); }, 1000);
  }

  async function copyText(text) {
    try {
      await navigator.clipboard.writeText(text);
      showToast('已复制');
    } catch (err) {
      var area = document.createElement('textarea');
      area.value = text;
      document.body.appendChild(area);
      area.select();
      var ok = document.execCommand('copy');
      area.remove();
      showToast(ok ? '已复制' : '复制失败，请手动选择文本');
    }
  }

  function showToast(message) {
    toast.textContent = message;
    toast.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { toast.classList.remove('show'); }, 1600);
  }

  function showError(message) {
    if (!message) {
      errorBox.hidden = true;
      errorBox.textContent = '';
      return;
    }
    errorBox.hidden = false;
    errorBox.textContent = message;
  }

  async function loadSample() {
    var workbook = new ExcelJS.Workbook();
    var sheet = workbook.addWorksheet('文案');
    sheet.getCell('A1').value = '危险警告';
    sheet.getCell('A1').font = { color: { argb: 'FFFF0000' }, bold: true };
    sheet.getCell('A2').value = '这行是黑色，不会加标签';
    sheet.getCell('A2').font = { color: { argb: 'FF000000' } };
    sheet.getCell('A3').value = '<color=#00FF00>治疗完成</color>';
    sheet.getCell('A3').font = { color: { argb: 'FF0000FF' } };
    sheet.getCell('A4').value = {
      richText: [
        { text: '获得' },
        { text: '稀有道具', font: { color: { argb: 'FFD4A017' } } },
        { text: '！', font: { color: { argb: 'FF00AAFF' } } }
      ]
    };
    sheet.getCell('A5').value = '主题色强调';
    sheet.getCell('A5').font = { color: { theme: 4 } };
    sheet.getCell('A6').value = '<b>暴击</b>';
    sheet.getCell('A6').font = { color: { argb: 'FFFF8800' } };
    sheet.getCell('A7').value = "获得<color='#00FF00'>金币</color>奖励";
    sheet.getCell('A7').font = { color: { argb: 'FFFF0000' } };
    sheet.getCell('A8').value = '<color=#FF0000>生命值</color>';
    sheet.getCell('A8').font = { color: { argb: 'FFFF0000' } };
    sheet.getCell('A9').value = '只有黄色底，不转换';
    sheet.getCell('A9').fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF99' } };
    var notes = workbook.addWorksheet('备注');
    notes.getCell('B2').value = '查看详情';
    notes.getCell('B2').font = { color: { theme: 10 }, underline: true };
    var buffer = await workbook.xlsx.writeBuffer();
    state.fileName = '示例表格.xlsx';
    convertBuffer(buffer);
  }
})();
