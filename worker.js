let terminated = false;

self.onmessage = function (e) {
  const msg = e.data;

  if (msg.cmd === 'terminate') {
    terminated = true;
    self.postMessage({ cmd: 'terminated' });
    self.close();
    return;
  }

  if (terminated || msg.cmd !== 'validate') return;

  let rows = msg.data;
  const fileName = msg.filename || '';
  const expectedB = fileName.split('_')[1]?.split('.')[0]?.trim() || '';

  rows = rows.filter(row => row?.some(cell => String(cell ?? '').trim() !== ''));

  const errors = [];
  const total = rows.length;

  function excelDateToJSDate(serial) {
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);
    const fractional_day = serial - Math.floor(serial) + 0.0000001;
    const total_seconds = Math.floor(86400 * fractional_day);
    const seconds = total_seconds % 60;
    const hours = Math.floor(total_seconds / 3600);
    const minutes = Math.floor((total_seconds % 3600) / 60);
    return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
  }

  function stripTime(date) {
    return new Date(date.getFullYear(), date.getMonth(), date.getDate());
  }

  function parseDate(raw) {
    if (raw === undefined || raw === null || String(raw).trim() === '') return null;

    let date = null;

    if (typeof raw === 'number') {
      date = excelDateToJSDate(raw);
    } else if (typeof raw === 'string') {
      const cleaned = raw.trim().replace(/[.\/]/g, '-');
      const parts = cleaned.split('-');
      if (parts.length === 3) {
        const [y, m, d] = parts.map(p => parseInt(p, 10));
        if (!isNaN(y) && !isNaN(m) && !isNaN(d)) {
          date = new Date(y, m - 1, d);
        }
      }
    }

    return date ? stripTime(date) : null;
  }

  const today = stripTime(new Date());

  if (expectedB && rows.length > 4 && String(rows[4][1] ?? '').trim() !== expectedB) {
    errors.push(`5행 오류 : B열 값 "${rows[4][1]}"과 파일명 "${expectedB}" 불일치.`);
  }

  const hashVal = Number(rows[3]?.[0]);
  const validHash = !isNaN(hashVal);
  const EXCEL_ROW_OFFSET = 1;

  for (let i = 4; i < rows.length; i++) {
    const row = rows[i];
    if (!row) continue;

    const rowNum = i + EXCEL_ROW_OFFSET;
    const errs = [];

    const aVal = Number(row[0]);
    if (!Number.isInteger(aVal)) errs.push(`A열(${rowNum}행) 정수 아님`);

    if (!validHash) {
      errs.push('4행 A열 "#" 값이 유효한 정수가 아님');
    } else {
      if (!row[1] || String(row[1]).trim() === '') errs.push(`B열(${rowNum}행) 비어 있음`);
      if (!row[2] || String(row[2]).trim() === '') errs.push(`C열(${rowNum}행) 비어 있음`);
    }

    const startDate = parseDate(row[4]);
    const endDate = parseDate(row[5]);

    if (!startDate) {
      errs.push(`E열(${rowNum}행) 시작일 변환 실패`);
    } else if (startDate < today) {
      errs.push(`E열(${rowNum}행) 오늘 이전`);
    }

    if (!endDate) {
      errs.push(`F열(${rowNum}행) 종료일 변환 실패`);
    } else {
      if (startDate && endDate < startDate) errs.push(`F열(${rowNum}행) 시작일보다 이전`);
      if (endDate < today) errs.push(`F열(${rowNum}행) 오늘 이전`);
    }

    const gVal = Number(String(row[6]).trim());
    const hVal = Number(String(row[7]).trim());

    if (isNaN(gVal)) errs.push(`G열(${rowNum}행) 숫자 아님`);
    if (isNaN(hVal)) errs.push(`H열(${rowNum}행) 숫자 아님`);

    if (!isNaN(gVal) && !isNaN(hVal)) {
      if (gVal < hVal) errs.push(`G열(${rowNum}행) 전체 인원이 일 최대 인원보다 작음`);
      if (startDate && endDate) {
        const diffDays = Math.round((endDate - startDate) / (1000 * 60 * 60 * 24)) + 1;
        if (gVal !== diffDays * hVal) {
          errs.push(`G열(${rowNum}행) ≠ (기간 * H열)`);
        }
      }
    }

    if (!row[10] || String(row[10]).trim() === '') errs.push(`K열(${rowNum}행) 비어 있음`);
    if (!row[11] || String(row[11]).trim() === '') errs.push(`L열(${rowNum}행) 비어 있음`);
    if (!row[12] || String(row[12]).trim() === '') errs.push(`M열(${rowNum}행) 비어 있음`);
    if (!row[17] || String(row[17]).trim() === '') errs.push(`R열(${rowNum}행) 비어 있음`);
    if (!row[18] || String(row[18]).trim() === '') errs.push(`S열(${rowNum}행) 비어 있음`);

    if (errs.length > 0) {
      errors.push(`${rowNum}행 오류: ${errs.join(', ')}`);
    }

    if (i % 10 === 0) {
      self.postMessage({ cmd: 'progress', current: i, total });
    }
  }

  self.postMessage({ cmd: 'progress', current: total, total });
  self.postMessage({ cmd: 'result', errors });
};
