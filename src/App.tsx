import { useState, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';

interface ParsedItem {
  code: string;
  name: string;
  unit: string;
  qty: number;
  price: number;
  sum: number;
  type: string;
}

interface ParsedData {
  [sheetName: string]: {
    [code: string]: ParsedItem;
  };
}

interface Checksums {
  before: { smr: number; tmc: number; total: number };
  after: { smr: number; tmc: number; total: number };
}

interface ReportRow {
  isSectionHeader?: boolean;
  isTotal?: boolean;
  text?: string;
  type?: string;
  code?: string;
  name?: string;
  unit?: string;
  qty?: number;
  sum?: number;
  label?: string;
  rowClass?: string;
  diffVal?: number;
  statusIcon?: string;
  isChange?: boolean;
  sumWas?: number;
  sumBecame?: number;
  sumDiff?: number;
}

const fmtMoney = new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

function parseSheet(rows: (string | number | undefined)[][], sheetName: string): { [code: string]: ParsedItem } {
  const items: { [code: string]: ParsedItem } = {};
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || row.length === 0) continue;

    let code = row[0];
    if (!code) continue;
    code = String(code).trim();

    items[code] = {
      code: code,
      name: String(row[1] || ''),
      unit: String(row[2] || ''),
      qty: parseFloat(String(row[3])) || 0,
      price: parseFloat(String(row[4])) || 0,
      sum: parseFloat(String(row[5])) || 0,
      type: sheetName.includes('СМР') ? 'КЕР' : 'ТМЦ'
    };
  }
  return items;
}

export default function App() {
  const [step, setStep] = useState(1);
  const [dataBefore, setDataBefore] = useState<ParsedData>({});
  const [dataAfter, setDataAfter] = useState<ParsedData>({});
  const [fileNameBefore, setFileNameBefore] = useState('');
  const [fileNameAfter, setFileNameAfter] = useState('');
  const [checksums, setChecksums] = useState<Checksums>({ before: { smr: 0, tmc: 0, total: 0 }, after: { smr: 0, tmc: 0, total: 0 } });
  const [reportRows, setReportRows] = useState<ReportRow[]>([]);
  const [loading, setLoading] = useState(false);
  const [showResults, setShowResults] = useState(false);
  const [showVersionModal, setShowVersionModal] = useState(false);
  const [totalCount, setTotalCount] = useState(0);
  const [changedCount, setChangedCount] = useState(0);
  const [totalDiff, setTotalDiff] = useState(0);

  const fileBeforeRef = useRef<HTMLInputElement>(null);
  const fileAfterRef = useRef<HTMLInputElement>(null);

  const handleFileSelect = useCallback((e: React.ChangeEvent<HTMLInputElement>, type: 'before' | 'after') => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (type === 'before') setFileNameBefore(file.name);
    else setFileNameAfter(file.name);

    setStep(2);

    const reader = new FileReader();
    reader.onload = function (ev) {
      const data = new Uint8Array(ev.target!.result as ArrayBuffer);
      const workbook = XLSX.read(data, { type: 'array' });

      const result: ParsedData = {};
      let sumSmr = 0, sumTmc = 0;

      ['СМР уник', 'ТМЦ уник'].forEach(sheetName => {
        if (workbook.SheetNames.includes(sheetName)) {
          const worksheet = workbook.Sheets[sheetName];
          const json = XLSX.utils.sheet_to_json<(string | number | undefined)[]>(worksheet, { header: 1 });
          const parsed = parseSheet(json, sheetName);
          result[sheetName] = parsed;

          Object.values(parsed).forEach(item => {
            if (sheetName.includes('СМР')) sumSmr += item.sum;
            else sumTmc += item.sum;
          });
        }
      });

      if (type === 'before') {
        setDataBefore(result);
        setChecksums(prev => ({ ...prev, before: { smr: sumSmr, tmc: sumTmc, total: sumSmr + sumTmc } }));
      } else {
        setDataAfter(result);
        setChecksums(prev => ({ ...prev, after: { smr: sumSmr, tmc: sumTmc, total: sumSmr + sumTmc } }));
      }

      setStep(3);
    };
    reader.readAsArrayBuffer(file);
  }, []);

  const processFiles = useCallback(() => {
    setLoading(true);
    setStep(4);

    setTimeout(() => {
      const rows: ReportRow[] = [];
      const types = ['КЕР', 'ТМЦ'];
      let grandTotalWas = 0;
      let grandTotalBecame = 0;
      let grandTotalDiff = 0;
      let changed = 0;

      types.forEach(type => {
        const sheetName = type === 'КЕР' ? 'СМР уник' : 'ТМЦ уник';
        const dataB = dataBefore[sheetName] || {};
        const dataA = dataAfter[sheetName] || {};

        const allCodes = new Set([...Object.keys(dataB), ...Object.keys(dataA)]);
        const sortedCodes = Array.from(allCodes).sort();

        let sectionTotalWas = 0;
        let sectionTotalBecame = 0;
        let sectionTotalDiff = 0;

        rows.push({ isSectionHeader: true, text: `=== ${type} ===` });

        sortedCodes.forEach(code => {
          const itemB = dataB[code];
          const itemA = dataA[code];

          const qtyB = itemB ? itemB.qty : 0;
          const sumB = itemB ? itemB.sum : 0;
          const qtyA = itemA ? itemA.qty : 0;
          const sumA = itemA ? itemA.sum : 0;
          const diff = sumA - sumB;

          sectionTotalWas += sumB;
          sectionTotalBecame += sumA;
          sectionTotalDiff += diff;

          if (diff !== 0) changed++;

          const name = itemA ? itemA.name : (itemB ? itemB.name : '');
          const unit = itemA ? itemA.unit : (itemB ? itemB.unit : '');

          rows.push({
            type, code, name, unit,
            qty: qtyB, sum: sumB, label: 'Было', rowClass: 'row-was', isChange: false
          });
          rows.push({
            type, code, name, unit,
            qty: qtyA, sum: sumA, label: 'Стало', rowClass: 'row-became', isChange: false
          });

          let statusIcon = '';
          if (diff > 0) statusIcon = '▲';
          else if (diff < 0) statusIcon = '▼';

          rows.push({
            type, code, name: '', unit: '',
            qty: qtyA - qtyB, sum: diff, label: 'Разница', rowClass: 'row-diff',
            diffVal: diff, statusIcon, isChange: (diff !== 0)
          });
        });

        rows.push({
          isTotal: true, type,
          text: `ВСЕГО ${type}`,
          sumWas: sectionTotalWas, sumBecame: sectionTotalBecame, sumDiff: sectionTotalDiff,
          rowClass: 'row-total'
        });

        grandTotalWas += sectionTotalWas;
        grandTotalBecame += sectionTotalBecame;
        grandTotalDiff += sectionTotalDiff;
      });

      rows.push({
        isTotal: true, type: 'ALL',
        text: 'ВСЕГО ОТЧЕТ',
        sumWas: grandTotalWas, sumBecame: grandTotalBecame, sumDiff: grandTotalDiff,
        rowClass: 'row-total'
      });

      setReportRows(rows);
      setTotalCount(new Set(rows.filter(r => r.code).map(r => r.code)).size);
      setChangedCount(changed);
      setTotalDiff(grandTotalDiff);
      setLoading(false);
      setShowResults(true);
    }, 800);
  }, [dataBefore, dataAfter]);

  const exportToExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Было-Стало');

    worksheet.columns = [
      { key: 'id', width: 8 }, { key: 'name', width: 54 }, { key: 'unit', width: 7 },
      { key: 'qty_was', width: 13 }, { key: 'price_was', width: 13 }, { key: 'sum_was', width: 13 },
      { key: 'qty_bec', width: 13 }, { key: 'price_bec', width: 13 }, { key: 'sum_bec', width: 13 },
      { key: 'diff', width: 13 }
    ];

    const C_WAS = 'FFFFCC', C_BECAME = 'FFCCFFCC', C_DIFF_HEADER = 'FFFF0000',
      C_HEADER_GRAY = 'FFC0C0C0', C_TOTAL_MAIN = 'FFCCFFCC', C_SECTION = 'FF00B0F0';
    const MONEY_FMT = '#,##0.00_ ;[Red]-#,##0.00';

    // 1. Шапка
    worksheet.mergeCells('A1:B1');
    const cellTitle = worksheet.getCell('A1');
    cellTitle.value = `${fileNameBefore}\n${fileNameAfter}`;
    cellInfo(cellTitle, { font: { name: 'Arial', size: 10 }, align: 'left', wrap: true, border: true });
    worksheet.getRow(1).height = 44;

    worksheet.mergeCells('C1:F1');
    cellInfo(worksheet.getCell('C1'), { value: 'БЫЛО', font: { name: 'Arial', size: 14, bold: true }, align: 'center', bg: C_WAS, border: true });

    worksheet.mergeCells('G1:I1');
    cellInfo(worksheet.getCell('G1'), { font: { name: 'Arial', size: 14, bold: true }, align: 'center', bg: C_BECAME, border: true });
    worksheet.getCell('G1').value = 'СТАЛО';

    cellInfo(worksheet.getCell('J1'), { value: 'РАЗНИЦА', font: { name: 'Arial', size: 14, bold: true }, align: 'center', bg: C_DIFF_HEADER, fontColor: 'FFFFFFFF', border: true });

    // 2. Заголовки
    const headers = ['ИД.КЕР/ТМЦ', 'Наименование КЕР/ТМЦ', 'ЕдИзм КЕР/ТМЦ', 'Объем СМР/ТМЦ Всего в СР', 'СМР/ТМЦ за ЕдИзм', 'СМР/ТМЦ Всего', 'Объем СМР/ТМЦ Всего в СР', 'СМР/ТМЦ за ЕдИзм', 'СМР/ТМЦ Всего', 'Стоимость всего руб с НДС'];
    worksheet.getRow(2).values = headers;
    worksheet.getRow(2).height = 34;
    for (let i = 1; i <= 10; i++) {
      cellInfo(worksheet.getCell(2, i), { font: { name: 'Aptos Narrow', size: 8 }, align: 'center', wrap: true, bg: C_HEADER_GRAY, border: true });
    }

    let currentRow = 3;

    // --- ВСЕГО СМР+ТМЦ ---
    worksheet.mergeCells(`A${currentRow}:B${currentRow}`);
    cellInfo(worksheet.getCell(`B${currentRow}`), { value: 'ВСЕГО СМР+ТМЦ', font: { name: 'Aptos Narrow', size: 10 }, align: 'center', bg: C_TOTAL_MAIN, border: true });

    let gWas = 0, gBec = 0;
    Object.values(dataBefore['СМР уник'] || {}).forEach(i => gWas += i.sum);
    Object.values(dataBefore['ТМЦ уник'] || {}).forEach(i => gWas += i.sum);
    Object.values(dataAfter['СМР уник'] || {}).forEach(i => gBec += i.sum);
    Object.values(dataAfter['ТМЦ уник'] || {}).forEach(i => gBec += i.sum);

    setMoneyCell(worksheet, currentRow, 6, gWas, MONEY_FMT);
    setMoneyCell(worksheet, currentRow, 9, gBec, MONEY_FMT);
    setMoneyCell(worksheet, currentRow, 10, gBec - gWas, MONEY_FMT);

    for (let c = 1; c <= 10; c++) {
      if (c !== 6 && c !== 9 && c !== 10) worksheet.getCell(currentRow, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C_TOTAL_MAIN } };
    }
    worksheet.getRow(currentRow).height = 13;
    currentRow++;

    // --- СМР ---
    cellInfo(worksheet.getCell(`B${currentRow}`), { value: 'СМР', font: { name: 'Aptos Narrow', size: 10 }, align: 'center', bg: C_SECTION, border: true });
    let sumSmrWas = 0, sumSmrBec = 0;
    Object.values(dataBefore['СМР уник'] || {}).forEach(i => sumSmrWas += i.sum);
    Object.values(dataAfter['СМР уник'] || {}).forEach(i => sumSmrBec += i.sum);

    setMoneyCell(worksheet, currentRow, 6, sumSmrWas, MONEY_FMT);
    setMoneyCell(worksheet, currentRow, 9, sumSmrBec, MONEY_FMT);
    setMoneyCell(worksheet, currentRow, 10, sumSmrBec - sumSmrWas, MONEY_FMT);

    for (let c = 1; c <= 10; c++) {
      if (c !== 6 && c !== 9 && c !== 10) worksheet.getCell(currentRow, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C_SECTION } };
    }
    worksheet.getRow(currentRow).height = 13;
    currentRow++;

    // --- ДАННЫЕ СМР ---
    const codesSmr = new Set([...Object.keys(dataBefore['СМР уник'] || {}), ...Object.keys(dataAfter['СМР уник'] || {})]);
    Array.from(codesSmr).sort().forEach(code => {
      const itemB = dataBefore['СМР уник']?.[code];
      const itemA = dataAfter['СМР уник']?.[code];
      const wasSum = itemB ? itemB.sum : 0;
      const becSum = itemA ? itemA.sum : 0;
      const diff = becSum - wasSum;
      const isNew = !itemB && itemA;

      const row = worksheet.getRow(currentRow);
      row.values = [code, itemA ? itemA.name : itemB?.name, itemA ? itemA.unit : itemB?.unit, itemB ? itemB.qty : 0, itemB ? itemB.price : 0, wasSum, itemA ? itemA.qty : 0, itemA ? itemA.price : 0, becSum, diff];

      for (let i = 1; i <= 10; i++) {
        const cell = row.getCell(i);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

        const style = { font: { name: 'Aptos Narrow', size: 8 }, border: true, align: (i === 1 || i === 3) ? 'center' : (i === 2 ? 'left' : 'right'), wrap: (i === 2) };
        if (!isNew) cellInfo(cell, style);
        else {
          cell.font = style.font;
          cell.alignment = { horizontal: style.align === 'left' ? 'left' : (style.align === 'right' ? 'right' : 'center'), vertical: 'middle', wrapText: !!style.wrap };
          if (style.border) {
            cell.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
          }
        }

        if ([6, 9, 10].includes(i)) cell.numFmt = MONEY_FMT;
      }
      currentRow++;
    });

    // --- ТМЦ ---
    const tmciRow = currentRow;    cellInfo(worksheet.getCell(`B${tmciRow}`), { value: 'ТМЦ', font: { name: 'Aptos Narrow', size: 10 }, align: 'center', bg: C_SECTION, border: true });
    let sumTmcWas = 0, sumTmcBec = 0;
    Object.values(dataBefore['ТМЦ уник'] || {}).forEach(i => sumTmcWas += i.sum);
    Object.values(dataAfter['ТМЦ уник'] || {}).forEach(i => sumTmcBec += i.sum);

    setMoneyCell(worksheet, tmciRow, 6, sumTmcWas, MONEY_FMT);
    setMoneyCell(worksheet, tmciRow, 9, sumTmcBec, MONEY_FMT);
    setMoneyCell(worksheet, tmciRow, 10, sumTmcBec - sumTmcWas, MONEY_FMT);

    for (let c = 1; c <= 10; c++) {
      if (c !== 6 && c !== 9 && c !== 10) worksheet.getCell(tmciRow, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: C_SECTION } };
    }
    worksheet.getRow(tmciRow).height = 13;
    currentRow++;

    // --- ДАННЫЕ ТМЦ ---
    const codesTmc = new Set([...Object.keys(dataBefore['ТМЦ уник'] || {}), ...Object.keys(dataAfter['ТМЦ уник'] || {})]);
    Array.from(codesTmc).sort().forEach(code => {
      const itemB = dataBefore['ТМЦ уник']?.[code];
      const itemA = dataAfter['ТМЦ уник']?.[code];
      const wasSum = itemB ? itemB.sum : 0;
      const becSum = itemA ? itemA.sum : 0;
      const diff = becSum - wasSum;
      const isNew = !itemB && itemA;

      const row = worksheet.getRow(currentRow);
      row.values = [code, itemA ? itemA.name : itemB?.name, itemA ? itemA.unit : itemB?.unit, itemB ? itemB.qty : 0, itemB ? itemB.price : 0, wasSum, itemA ? itemA.qty : 0, itemA ? itemA.price : 0, becSum, diff];

      for (let i = 1; i <= 10; i++) {
        const cell = row.getCell(i);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFFFF' } };

        const style = { font: { name: 'Aptos Narrow', size: 8 }, border: true, align: (i === 1 || i === 3) ? 'center' : (i === 2 ? 'left' : 'right'), wrap: (i === 2) };
        if (!isNew) cellInfo(cell, style);
        else {
          cell.font = style.font;
          cell.alignment = { horizontal: style.align === 'left' ? 'left' : (style.align === 'right' ? 'right' : 'center'), vertical: 'middle', wrapText: !!style.wrap };
          if (style.border) {
            cell.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
          }
        }
        if ([6, 9, 10].includes(i)) cell.numFmt = MONEY_FMT;
      }
      currentRow++;
    });

    // Границы
    const lastRow = currentRow - 1;
    worksheet.eachRow(row => {
      if (row.number <= lastRow) {
        row.eachCell(cell => {
          if (!cell.border) {
            cell.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
          }
        });
      }
    });

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    saveAs(blob, "StateDiff_Report.xlsx");
  };

  const resetApp = () => {
    window.location.reload();
  };

  const bothLoaded = Object.keys(dataBefore).length > 0 && Object.keys(dataAfter).length > 0;

  return (
    <div className="app-container">
      <div className="floating-icon" style={{ top: '10%', left: '5%', fontSize: '4rem' }}>Σ</div>
      <div className="floating-icon" style={{ top: '20%', right: '10%', fontSize: '3rem' }}>₽</div>
      <div className="floating-icon" style={{ bottom: '15%', left: '15%', fontSize: '5rem' }}>=</div>

      <div className="container mx-auto px-4 py-8 relative z-10 flex-grow flex flex-col">
        <header className="text-center mb-10">
          <div className="flex justify-center items-center gap-6 mb-4">
            <img src="/StateDiffGHP/logo.png" alt="Логотип" className="h-20 w-auto object-contain drop-shadow-sm" />
            <h1 className="text-4xl md:text-5xl font-bold text-green-900 tracking-tight">
              Форма БылоСтало (для СУ)
            </h1>
          </div>
          <p className="text-xl text-green-700 font-medium bg-white/70 inline-block px-6 py-2 rounded-full border border-green-300 shadow-sm">
            Мастер сравнения КЕР и ТМЦ
          </p>
        </header>

        {/* Step indicators */}
        <div className="max-w-4xl mx-auto w-full mb-8">
          <div className="flex justify-between items-center relative">
            <div className="absolute top-1/2 left-0 w-full h-1 bg-green-200 -z-10 rounded"></div>
            {[1, 2, 3, 4].map(s => (
              <div key={s} className={`step-indicator w-10 h-10 rounded-full border-2 bg-white flex items-center justify-center font-bold transition-all duration-300 ${s < step || (s === step && (s === 1 || s === 4 || (s === 3 && bothLoaded))) ? 'step-completed' : s === step ? 'step-active' : 'border-gray-300 text-gray-400'}`}>
                {s}
              </div>
            ))}
          </div>
          <div className="flex justify-between mt-2 text-xs font-semibold text-green-800 uppercase tracking-wide">
            <span>Загрузка</span><span>Проверка</span><span>Сравнение</span><span>Отчет</span>
          </div>
        </div>

        <main className="flex-grow max-w-7xl mx-auto w-full">
          {/* Upload screen */}
          {!showResults && !loading && (
            <>
              <div className="grid md:grid-cols-2 gap-8">
                {/* Before */}
                <div className="card-shadow bg-white rounded-xl p-8 flex flex-col items-center justify-center text-center h-80 hover:bg-green-50 cursor-pointer relative group" onClick={() => fileBeforeRef.current?.click()}>
                  <input ref={fileBeforeRef} type="file" className="hidden" accept=".xlsx,.xls" onChange={(e) => handleFileSelect(e, 'before')} />
                  <div className="absolute inset-0 border-4 border-dashed border-green-300 rounded-xl m-2 pointer-events-none group-hover:border-green-500 transition-colors"></div>
                  <svg className="w-16 h-16 text-green-600 mb-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path></svg>
                  <h3 className="text-2xl font-bold text-green-900 mb-2">Файл "БЫЛО"</h3>
                  <p className="text-green-600">Перетащите или кликните</p>
                  {fileNameBefore && <div className="mt-4 px-3 py-1 bg-green-100 text-green-800 rounded-full text-sm font-semibold break-all">{fileNameBefore}</div>}
                </div>

                {/* After */}
                <div className="card-shadow bg-white rounded-xl p-8 flex flex-col items-center justify-center text-center h-80 hover:bg-green-50 cursor-pointer relative group" onClick={() => fileAfterRef.current?.click()}>
                  <input ref={fileAfterRef} type="file" className="hidden" accept=".xlsx,.xls" onChange={(e) => handleFileSelect(e, 'after')} />
                  <div className="absolute inset-0 border-4 border-dashed border-green-300 rounded-xl m-2 pointer-events-none group-hover:border-green-500 transition-colors"></div>
                  <svg className="w-16 h-16 text-blue-600 mb-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"></path></svg>
                  <h3 className="text-2xl font-bold text-green-900 mb-2">Файл "СТАЛО"</h3>
                  <p className="text-green-600">Перетащите или кликните</p>
                  {fileNameAfter && <div className="mt-4 px-3 py-1 bg-blue-100 text-blue-800 rounded-full text-sm font-semibold break-all">{fileNameAfter}</div>}
                </div>
              </div>

              {/* Checksums */}
              {bothLoaded && (
                <div className="mt-8 max-w-4xl mx-auto">
                  <div className="bg-white rounded-xl border-2 border-green-600 shadow-lg p-6">
                    <h3 className="text-xl font-bold text-green-900 mb-4 text-center">Проверочные суммы (Контроль)</h3>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                      <div className="space-y-2">
                        <p className="font-bold text-green-700 border-b border-green-200 pb-1">Файл "БЫЛО"</p>
                        <div className="flex justify-between text-sm"><span>СМР:</span> <span className="font-mono">{fmtMoney.format(checksums.before.smr)}</span></div>
                        <div className="flex justify-between text-sm"><span>ТМЦ:</span> <span className="font-mono">{fmtMoney.format(checksums.before.tmc)}</span></div>
                        <div className="flex justify-between font-bold text-green-900 pt-2 border-t border-green-200"><span>ВСЕГО:</span> <span className="font-mono">{fmtMoney.format(checksums.before.total)}</span></div>
                      </div>
                      <div className="space-y-2">
                        <p className="font-bold text-blue-700 border-b border-blue-200 pb-1">Файл "СТАЛО"</p>
                        <div className="flex justify-between text-sm"><span>СМР:</span> <span className="font-mono">{fmtMoney.format(checksums.after.smr)}</span></div>
                        <div className="flex justify-between text-sm"><span>ТМЦ:</span> <span className="font-mono">{fmtMoney.format(checksums.after.tmc)}</span></div>
                        <div className="flex justify-between font-bold text-blue-900 pt-2 border-t border-blue-200"><span>ВСЕГО:</span> <span className="font-mono">{fmtMoney.format(checksums.after.total)}</span></div>
                      </div>
                    </div>
                  </div>
                </div>
              )}

              {/* Action button */}
              {bothLoaded && (
                <div className="text-center mt-10">
                  <button onClick={processFiles} className="btn-primary px-10 py-4 rounded-lg text-lg uppercase tracking-wider hover:bg-green-600">
                    Сравнить данные
                  </button>
                  <p className="mt-4 text-sm text-green-700 flex items-center justify-center gap-2">
                    <svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z"></path></svg>
                    файлы обрабатываются локально, никуда не отправляются
                  </p>
                </div>
              )}
            </>
          )}

          {/* Loading */}
          {loading && (
            <div className="text-center py-20">
              <div className="inline-block animate-spin rounded-full h-16 w-16 border-t-4 border-b-4 border-green-600 mb-4"></div>
              <h3 className="text-2xl font-bold text-green-800">Анализ данных...</h3>
              <p className="text-green-600">Сопоставление кодов и расчет разниц</p>
            </div>
          )}

          {/* Results */}
          {showResults && (
            <div className="space-y-6">
              <div className="flex justify-between items-center flex-wrap gap-4">
                <h2 className="text-2xl font-bold text-green-900">Результаты сравнения</h2>
                <button onClick={exportToExcel} className="btn-download px-6 py-3 rounded-lg text-sm uppercase tracking-wider hover:bg-blue-800 flex items-center gap-2">
                  <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"></path></svg>
                  Скачать отчет Excel (Было-Стало)
                </button>
              </div>

              {/* Summary */}
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div className="bg-white p-6 rounded-lg border-2 border-green-200 shadow-sm">
                  <p className="text-sm text-gray-500 uppercase font-bold">Позиций всего</p>
                  <p className="text-3xl font-bold text-green-800">{totalCount}</p>
                </div>
                <div className="bg-white p-6 rounded-lg border-2 border-green-200 shadow-sm">
                  <p className="text-sm text-gray-500 uppercase font-bold">Изменено позиций</p>
                  <p className="text-3xl font-bold text-orange-600">{changedCount}</p>
                </div>
                <div className="bg-white p-6 rounded-lg border-2 border-green-200 shadow-sm">
                  <p className="text-sm text-gray-500 uppercase font-bold">Итоговая разница</p>
                  <p className={`text-3xl font-bold ${totalDiff >= 0 ? 'text-green-600' : 'text-red-600'}`}>{fmtMoney.format(totalDiff)} ₽</p>
                </div>
              </div>

              {/* Table */}
              <div className="bg-white rounded-xl shadow-lg border-2 border-green-800 overflow-hidden">
                <div className="bg-green-900 text-white px-6 py-4">
                  <h3 className="text-xl font-bold">Детальный отчет</h3>
                </div>

                <div className="overflow-x-auto max-h-[700px]">
                  <table>
                    <thead>
                      <tr>
                        <th style={{ width: '10%' }}>Тип</th>
                        <th style={{ width: '15%' }}>Код</th>
                        <th style={{ width: '35%' }}>Наименование</th>
                        <th style={{ width: '10%' }}>Ед.</th>
                        <th style={{ width: '10%' }} className="text-right">Кол-во</th>
                        <th style={{ width: '15%' }} className="text-right">Сумма</th>
                        <th style={{ width: '5%' }} className="text-center">Статус</th>
                      </tr>
                    </thead>
                    <tbody>
                      {reportRows.map((row, idx) => {
                        if (row.isSectionHeader) {
                          return (
                            <tr key={idx}>
                              <td colSpan={7} className="bg-green-100 font-bold text-green-900 text-center py-2">{row.text}</td>
                            </tr>
                          );
                        }

                        if (row.isTotal) {
                          return (
                            <tr key={idx} className="row-total">
                              <td colSpan={4} className="text-right font-bold text-white">{row.text}</td>
                              <td className="text-right text-xs text-green-200">{fmtMoney.format(row.sumWas || 0)}</td>
                              <td className="text-right text-xs text-green-200">{fmtMoney.format(row.sumBecame || 0)}</td>
                              <td className="text-right font-bold text-lg">{fmtMoney.format(row.sumDiff || 0)}</td>
                            </tr>
                          );
                        }

                        let diffColor = '';
                        if (row.label === 'Разница') {
                          if ((row.diffVal || 0) > 0) diffColor = 'text-green-700';
                          if ((row.diffVal || 0) < 0) diffColor = 'text-red-600';
                        }

                        return (
                          <tr key={idx} className={row.rowClass}>
                            <td className="text-xs font-bold text-gray-500">{row.type}</td>
                            <td className="font-mono text-xs">{row.code}</td>
                            <td className="font-medium truncate max-w-xs" title={row.name}>{row.name}</td>
                            <td className="text-xs">{row.unit}</td>
                            <td className="text-right font-mono text-xs">{fmtMoney.format(row.qty || 0)}</td>
                            <td className={`text-right font-mono ${diffColor}`}>{fmtMoney.format(row.sum || 0)}</td>
                            <td className="text-center text-lg">{row.statusIcon || (row.label === 'Разница' ? '•' : '')}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>

              <div className="text-center pb-10">
                <button onClick={resetApp} className="text-green-700 underline hover:text-green-900 font-medium">
                  Начать заново
                </button>
              </div>
            </div>
          )}
        </main>

        {/* Footer */}
        <footer className="mt-12 mb-6 text-center">
          <span className="version-link cursor-pointer" onClick={() => setShowVersionModal(true)}>Версия 1.0 (14.09.2026)</span>
        </footer>
      </div>

      {/* Version modal */}
      {showVersionModal && (
        <div className="fixed inset-0 modal-overlay z-50 flex items-center justify-center p-4" onClick={(e) => { if (e.target === e.currentTarget) setShowVersionModal(false); }}>
          <div className="bg-white rounded-xl shadow-2xl max-w-lg w-full max-h-[80vh] overflow-y-auto border-2 border-green-800">
            <div className="bg-green-900 text-white px-6 py-4 flex justify-between items-center sticky top-0">
              <h3 className="text-xl font-bold">История стабильных версий</h3>
              <button onClick={() => setShowVersionModal(false)} className="text-white hover:text-green-200 text-2xl">&times;</button>
            </div>
            <div className="p-6 space-y-6">
              <div className="border-l-4 border-green-500 pl-4">
                <div className="flex justify-between items-baseline mb-1">
                  <h4 className="font-bold text-lg text-green-900">Версия 1.0</h4>
                  <span className="text-xs text-gray-500 font-mono">14.09.2026</span>
                </div>
                <ul className="list-disc list-inside text-sm text-gray-700 space-y-1">
                  <li>Исправлено форматирование чисел в Excel (русская локаль).</li>
                  <li>Добавлена проверка контрольных сумм перед сравнением.</li>
                  <li>Реализована полная заливка ячеек в строках итогов (ВСЕГО, СМР, ТМЦ) в Excel.</li>
                  <li>Добавлено предупреждение о локальной обработке данных.</li>
                  <li>Добавлена кнопка версии с историей изменений.</li>
                  <li>Устранены проблемы с отображением пустых ячеек в Excel.</li>
                </ul>
              </div>
            </div>
            <div className="px-6 py-4 bg-gray-50 border-t border-gray-200 text-center">
              <button onClick={() => setShowVersionModal(false)} className="text-green-700 font-semibold hover:text-green-900">Закрыть</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// Helper functions for Excel export
function cellInfo(cell: ExcelJS.Cell, opts: { value?: string; font?: Partial<ExcelJS.Font>; align?: string; bg?: string; fontColor?: string; border?: boolean; wrap?: boolean }) {
  if (opts.value) cell.value = opts.value;
  if (opts.font) cell.font = opts.font as ExcelJS.Font;
  if (opts.align) {
    cell.alignment = { horizontal: opts.align === 'left' ? 'left' : (opts.align === 'right' ? 'right' : 'center'), vertical: 'middle', wrapText: !!opts.wrap };
  }
  if (opts.bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: opts.bg } };
  if (opts.fontColor) cell.font = { ...cell.font as ExcelJS.Font, color: { argb: opts.fontColor } };
  if (opts.border) {
    cell.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
  }
}

function setMoneyCell(ws: ExcelJS.Worksheet, row: number, col: number, value: number, fmt: string) {
  const cell = ws.getCell(row, col);
  cell.value = value;
  cell.numFmt = fmt;
  cell.alignment = { horizontal: 'right', vertical: 'middle' };
  cell.border = { top: { style: 'thin', color: { argb: 'FF000000' } }, left: { style: 'thin', color: { argb: 'FF000000' } }, bottom: { style: 'thin', color: { argb: 'FF000000' } }, right: { style: 'thin', color: { argb: 'FF000000' } } };
}
