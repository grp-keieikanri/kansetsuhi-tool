// v2.0.1 - 2026-04-14
import { useState, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ============================================================
// CONSTANTS
// ============================================================
const EMPLOYMENT_TYPES = ["正社員", "サポート", "アルバイト", "派遣社員", "社内外注"];
const BASE_RATES = { 正社員: 200000, サポート: 150000, アルバイト: 120000, 派遣社員: 50000, 社内外注: 30000 };

function getConversionRate(emp, value) {
  if (emp === "正社員" || emp === "サポート") return 1;
  if (emp === "アルバイト") {
    if (value < 30) return 0;
    if (value < 50) return 0.3;
    if (value < 100) return 0.5;
    return 1;
  }
  if (emp === "派遣社員") return value <= 170000 ? 0.5 : 1;
  if (emp === "社内外注") {
    if (value < 50000) return 0;
    if (value < 100000) return 0.3;
    if (value < 150000) return 0.5;
    if (value < 200000) return 0.7;
    return 1;
  }
  return 1;
}

function calcEmployee(emp) {
  const rate = getConversionRate(emp.employmentType, parseFloat(emp.criteriaValue) || 0);
  const baseAmount = rate * BASE_RATES[emp.employmentType];
  const allocations = emp.departments
    .filter(d => d.name && d.ratio > 0)
    .map(d => ({ department: d.name, amount: baseAmount * (d.ratio / 100), ratio: d.ratio }));
  return { rate, baseAmount, allocations };
}

function fmt(n) { return Math.round(n).toLocaleString("ja-JP"); }

function normalizeSpaces(str) {
  // スペース除去 + NFKC正規化 + 旧字体→新字体変換
  const kanjiMap = {
    '郞': '郎', // 郞→郎
    '圓': '円', // 圓→円
    '濵': '浜', // 濱→浜
    '澤': '沢', // 澤→沢
    '齋': '斎', // 齋→斎
    '齊': '斉', // 齊→斉
    '廣': '広', // 廣→広
    '德': '徳', // 德→徳
    '國': '国', // 國→国
    '會': '会', // 會→会
    '實': '実', // 實→実
    '關': '関', // 關→関
    '變': '変', // 變→変
  };
  let s = (str || "").replace(/[\s\u3000\u00a0]/g, "").normalize("NFKC");
  for (const [old, nw] of Object.entries(kanjiMap)) {
    s = s.split(old).join(nw);
  }
  return s;
}

// ============================================================
// EXCEL IMPORT HELPER
// ============================================================
function parseExcelToEmployees(data) {
  const ws = data.Sheets[data.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
  return rows.map(row => {
    const departments = [];
    for (let i = 1; i <= 10; i++) {
      const name = row[`部門${i}`] || (i === 1 ? row["部門"] : "");
      let ratio = parseFloat(row[`割合${i}`] || (i === 1 ? row["割合"] : 0)) || 0;
      if (ratio > 1.5) ratio = ratio / 100;
      if (name) departments.push({ name: String(name), ratio: ratio * 100 });
    }
    if (departments.length === 0) departments.push({ name: "", ratio: 100 });
    return {
      id: Math.random().toString(36).slice(2),
      name: String(row["氏名"] || ""),
      employmentType: String(row["雇用形態"] || "正社員"),
      criteriaValue: String(row["判定基準値"] || ""),
      departments,
    };
  }).filter(e => e.name);
}

// ============================================================
// COMPONENTS
// ============================================================

function Tag({ color, children }) {
  const colors = {
    blue: "bg-blue-100 text-blue-700",
    green: "bg-emerald-100 text-emerald-700",
    orange: "bg-orange-100 text-orange-700",
    purple: "bg-purple-100 text-purple-700",
    pink: "bg-pink-100 text-pink-700",
  };
  return <span className={`text-xs px-2 py-0.5 rounded-full font-medium ${colors[color] || colors.blue}`}>{children}</span>;
}

// ============================================================
// TAB: データ入力
// ============================================================
function DataInputTab({ employees, setEmployees }) {
  const fileRef = useRef();

  const addEmployee = () => {
    setEmployees(prev => [...prev, {
      id: Math.random().toString(36).slice(2),
      name: "", employmentType: "正社員", criteriaValue: "",
      departments: [{ name: "", ratio: 100 }]
    }]);
  };

  const removeEmployee = (id) => setEmployees(prev => prev.filter(e => e.id !== id));

  const updateEmployee = (id, field, value) => {
    setEmployees(prev => prev.map(e => e.id === id ? { ...e, [field]: value } : e));
  };

  const updateDept = (empId, idx, field, value) => {
    setEmployees(prev => prev.map(e => {
      if (e.id !== empId) return e;
      const depts = [...e.departments];
      depts[idx] = { ...depts[idx], [field]: field === "ratio" ? parseFloat(value) || 0 : value };
      return { ...e, departments: depts };
    }));
  };

  const addDept = (empId) => {
    setEmployees(prev => prev.map(e => e.id === empId ? { ...e, departments: [...e.departments, { name: "", ratio: 0 }] } : e));
  };

  const removeDept = (empId, idx) => {
    setEmployees(prev => prev.map(e => {
      if (e.id !== empId) return e;
      const depts = e.departments.filter((_, i) => i !== idx);
      return { ...e, departments: depts.length ? depts : [{ name: "", ratio: 100 }] };
    }));
  };

  const handleImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: "binary" });
      const imported = parseExcelToEmployees(wb);
      setEmployees(prev => {
        const updated = [...prev];
        imported.forEach(imp => {
          const idx = updated.findIndex(u => normalizeSpaces(u.name) === normalizeSpaces(imp.name) && u.employmentType === imp.employmentType);
          if (idx >= 0) {
            updated[idx] = {
              ...updated[idx],
              ...(imp.criteriaValue ? { criteriaValue: imp.criteriaValue } : {}),
              ...(imp.departments[0]?.name ? { departments: imp.departments } : {}),
            };
          } else {
            updated.push(imp);
          }
        });
        return updated;
      });
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const empTagColor = { 正社員: "blue", サポート: "purple", アルバイト: "green", 派遣社員: "orange", 社内外注: "pink" };

  return (
    <div className="space-y-4">
      {/* Guide */}
      <div className="bg-indigo-50 border border-indigo-200 rounded-xl p-4 text-sm text-indigo-800">
        <div className="font-bold mb-1">📥 Excelインポートの形式</div>
        <div className="mb-1">列名：<code className="bg-white px-1 rounded">氏名</code> <code className="bg-white px-1 rounded">雇用形態</code> <code className="bg-white px-1 rounded">判定基準値</code> <code className="bg-white px-1 rounded">部門1</code> <code className="bg-white px-1 rounded">割合1</code> ...</div>
        <div className="mb-1">割合は <code className="bg-white px-1 rounded">100</code>・<code className="bg-white px-1 rounded">50%</code>・<code className="bg-white px-1 rounded">0.5</code> いずれでも可。</div>
        <div className="mt-2 font-semibold">💡 活用例（バラバラのデータを取り込む場合）</div>
        <ol className="list-decimal ml-4 mt-1 space-y-0.5">
          <li>まず「配賦先・割合」のExcelをインポート（土台作成）</li>
          <li>次に「アルバイトの勤務時間」Excelをインポート（基準値の更新）</li>
          <li>最後に「社内外注の費用」Excelをインポート（基準値の更新）</li>
        </ol>
        <div className="mt-1 text-indigo-600">※ 氏名と雇用形態が一致していれば、既存データに統合されます。</div>
      </div>

      <div className="flex justify-between items-center">
        <div className="text-sm text-gray-500">{employees.length} 名登録</div>
        <div className="flex gap-2">
          <button onClick={() => fileRef.current.click()}
            className="flex items-center gap-1.5 px-3 py-2 bg-emerald-600 text-white text-sm rounded-lg hover:bg-emerald-700 transition-colors">
            <span>📂</span> Excelインポート
          </button>
          <input ref={fileRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleImport} />
          <button onClick={addEmployee}
            className="flex items-center gap-1.5 px-3 py-2 bg-indigo-600 text-white text-sm rounded-lg hover:bg-indigo-700 transition-colors">
            <span>＋</span> 従業員追加
          </button>
        </div>
      </div>

      <div className="space-y-3">
        {employees.map((emp) => {
          const { rate, baseAmount } = calcEmployee(emp);
          const totalRatio = emp.departments.reduce((s, d) => s + (d.ratio || 0), 0);
          const ratioOk = Math.abs(totalRatio - 100) < 0.1;
          return (
            <div key={emp.id} className="border border-gray-200 rounded-xl p-4 bg-white shadow-sm">
              <div className="flex flex-wrap gap-3 items-start mb-3">
                <input value={emp.name} onChange={e => updateEmployee(emp.id, "name", e.target.value)}
                  placeholder="氏名" className="border rounded-lg px-3 py-1.5 text-sm w-32 focus:ring-2 focus:ring-indigo-300 outline-none" />
                <select value={emp.employmentType} onChange={e => updateEmployee(emp.id, "employmentType", e.target.value)}
                  className="border rounded-lg px-3 py-1.5 text-sm focus:ring-2 focus:ring-indigo-300 outline-none">
                  {EMPLOYMENT_TYPES.map(t => <option key={t}>{t}</option>)}
                </select>
                {(emp.employmentType === "アルバイト" || emp.employmentType === "派遣社員" || emp.employmentType === "社内外注") && (
                  <div className="flex items-center gap-1">
                    <input value={emp.criteriaValue} onChange={e => updateEmployee(emp.id, "criteriaValue", e.target.value)}
                      placeholder={emp.employmentType === "アルバイト" ? "時間" : "金額（税抜）"}
                      className="border rounded-lg px-3 py-1.5 text-sm w-32 focus:ring-2 focus:ring-indigo-300 outline-none" />
                    <span className="text-xs text-gray-400">{emp.employmentType === "アルバイト" ? "時間" : "円（税抜）"}</span>
                  </div>
                )}
                <div className="ml-auto flex items-center gap-2">
                  <Tag color={empTagColor[emp.employmentType]}>{emp.employmentType}</Tag>
                  <span className="text-xs text-gray-500">換算: {rate}人 → <strong className="text-indigo-700">¥{fmt(baseAmount)}</strong></span>
                  <button onClick={() => removeEmployee(emp.id)} className="text-red-400 hover:text-red-600 text-lg leading-none">×</button>
                </div>
              </div>
              <div className="space-y-1.5">
                {emp.departments.map((dept, idx) => (
                  <div key={idx} className="flex gap-2 items-center">
                    <input value={dept.name} onChange={e => updateDept(emp.id, idx, "name", e.target.value)}
                      placeholder="部門名" className="border rounded-lg px-3 py-1.5 text-sm w-40 focus:ring-2 focus:ring-indigo-300 outline-none" />
                    <input type="number" value={dept.ratio} onChange={e => updateDept(emp.id, idx, "ratio", e.target.value)}
                      placeholder="割合%" className="border rounded-lg px-3 py-1.5 text-sm w-20 focus:ring-2 focus:ring-indigo-300 outline-none" />
                    <span className="text-xs text-gray-400">%</span>
                    {emp.departments.length > 1 && (
                      <button onClick={() => removeDept(emp.id, idx)} className="text-gray-400 hover:text-red-500 text-sm">削除</button>
                    )}
                  </div>
                ))}
                <div className="flex items-center gap-3 mt-1">
                  <button onClick={() => addDept(emp.id)} className="text-xs text-indigo-600 hover:underline">＋ 部門追加</button>
                  {!ratioOk && <span className="text-xs text-red-500">⚠ 割合合計: {totalRatio.toFixed(1)}%（100%にしてください）</span>}
                  {ratioOk && emp.departments[0].name && <span className="text-xs text-emerald-600">✓ 割合OK</span>}
                </div>
              </div>
            </div>
          );
        })}
        {employees.length === 0 && (
          <div className="text-center py-12 text-gray-400 border-2 border-dashed border-gray-200 rounded-xl">
            従業員を追加するか、Excelをインポートしてください
          </div>
        )}
      </div>
    </div>
  );
}

// ============================================================
// TAB: 部門別集計
// ============================================================
function DeptSummaryTab({ employees }) {
  const deptMap = {};
  employees.forEach(emp => {
    const { rate, allocations } = calcEmployee(emp);
    allocations.forEach(({ department, amount }) => {
      if (!deptMap[department]) deptMap[department] = { total: 0, breakdown: {} };
      deptMap[department].total += amount;
      const et = emp.employmentType;
      deptMap[department].breakdown[et] = (deptMap[department].breakdown[et] || 0) + rate * (emp.departments.find(d => d.name === department)?.ratio || 0) / 100;
    });
  });

  const depts = Object.entries(deptMap).sort((a, b) => b[1].total - a[1].total);
  const grandTotal = depts.reduce((s, [, v]) => s + v.total, 0);
  const maxTotal = depts[0]?.[1].total || 1;

  const handleExport = () => {
    const rows = depts.map(([name, { total, breakdown }]) => {
      const row = { 部門名: name, 合計請求額: total };
      EMPLOYMENT_TYPES.forEach(et => { row[`${et}換算人数`] = breakdown[et] ? parseFloat(breakdown[et].toFixed(2)) : 0; });
      row["換算人数合計"] = parseFloat(Object.values(breakdown).reduce((s, v) => s + v, 0).toFixed(2));
      return row;
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    ws["!cols"] = [{ wch: 20 }, { wch: 15 }, ...EMPLOYMENT_TYPES.map(() => ({ wch: 14 })), { wch: 14 }];
    // Format 合計請求額 as number with comma
    const range = XLSX.utils.decode_range(ws["!ref"]);
    for (let R = 1; R <= range.e.r; R++) {
      const cell = ws[XLSX.utils.encode_cell({ r: R, c: 1 })];
      if (cell) { cell.t = "n"; cell.z = "#,##0"; }
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "部門別集計");
    XLSX.writeFile(wb, "部門別集計結果.xlsx");
  };

  const empTagColor = { 正社員: "#6366f1", サポート: "#8b5cf6", アルバイト: "#10b981", 派遣社員: "#f59e0b", 社内外注: "#ec4899" };

  return (
    <div className="space-y-4">
      <div className="flex justify-between items-center">
        <div className="text-sm text-gray-500">全社合計: <strong className="text-indigo-700 text-base">¥{fmt(grandTotal)}</strong></div>
        <button onClick={handleExport}
          className="flex items-center gap-1.5 px-3 py-2 bg-emerald-600 text-white text-sm rounded-lg hover:bg-emerald-700 transition-colors">
          📥 Excelエクスポート
        </button>
      </div>
      {depts.length === 0 ? (
        <div className="text-center py-12 text-gray-400 border-2 border-dashed border-gray-200 rounded-xl">データを入力してください</div>
      ) : (
        <div className="overflow-x-auto">
          <table className="w-full text-sm border-collapse">
            <thead>
              <tr className="bg-gray-50">
                <th className="text-left px-4 py-3 border-b border-gray-200 font-semibold text-gray-700">部門名</th>
                <th className="text-right px-4 py-3 border-b border-gray-200 font-semibold text-gray-700">合計請求額</th>
                <th className="px-4 py-3 border-b border-gray-200 font-semibold text-gray-700">構成比</th>
              </tr>
            </thead>
            <tbody>
              {depts.map(([name, { total, breakdown }]) => {
                const totalPeople = Object.values(breakdown).reduce((s, v) => s + v, 0);
                return (
                  <tr key={name} className="border-b border-gray-100 hover:bg-gray-50 transition-colors">
                    <td className="px-4 py-3 align-top">
                      <div className="font-medium text-gray-800">{name}</div>
                      <div className="flex flex-wrap gap-1 mt-1">
                        {EMPLOYMENT_TYPES.filter(et => breakdown[et]).map(et => (
                          <span key={et} className="text-xs px-2 py-0.5 rounded-full text-white" style={{ backgroundColor: empTagColor[et] }}>
                            {et}: {breakdown[et].toFixed(2)}人
                          </span>
                        ))}
                        <span className="text-xs px-2 py-0.5 rounded-full bg-gray-200 text-gray-600">合計: {totalPeople.toFixed(2)}人</span>
                      </div>
                    </td>
                    <td className="px-4 py-3 text-right font-mono font-semibold text-gray-800">¥{fmt(total)}</td>
                    <td className="px-4 py-3">
                      <div className="flex items-center gap-2">
                        <div className="flex-1 bg-gray-200 rounded-full h-2">
                          <div className="bg-indigo-500 h-2 rounded-full" style={{ width: `${(total / maxTotal) * 100}%` }} />
                        </div>
                        <span className="text-xs text-gray-500 w-10 text-right">{((total / grandTotal) * 100).toFixed(1)}%</span>
                      </div>
                    </td>
                  </tr>
                );
              })}
            </tbody>
            <tfoot>
              <tr className="bg-indigo-50">
                <td className="px-4 py-3 font-bold text-indigo-800">全社合計</td>
                <td className="px-4 py-3 text-right font-mono font-bold text-indigo-800">¥{fmt(grandTotal)}</td>
                <td className="px-4 py-3" />
              </tr>
            </tfoot>
          </table>
        </div>
      )}
    </div>
  );
}

// ============================================================
// TAB: 個人別明細
// ============================================================
function IndividualTab({ employees }) {
  return (
    <div className="space-y-3">
      {employees.length === 0 ? (
        <div className="text-center py-12 text-gray-400 border-2 border-dashed border-gray-200 rounded-xl">データを入力してください</div>
      ) : (
        employees.map(emp => {
          const v = parseFloat(emp.criteriaValue) || 0;
          const { rate, baseAmount, allocations } = calcEmployee(emp);
          return (
            <div key={emp.id} className="border border-gray-200 rounded-xl p-4 bg-white shadow-sm">
              <div className="flex items-center gap-2 mb-2">
                <span className="font-semibold text-gray-800">{emp.name || "（未入力）"}</span>
                <span className="text-xs bg-gray-100 text-gray-600 px-2 py-0.5 rounded-full">{emp.employmentType}</span>
              </div>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3 text-sm mb-3">
                <div className="bg-gray-50 rounded-lg p-2">
                  <div className="text-xs text-gray-500">判定基準値</div>
                  <div className="font-medium">{v || "—"} {emp.employmentType === "アルバイト" ? "h" : emp.criteriaValue ? "円" : ""}</div>
                </div>
                <div className="bg-gray-50 rounded-lg p-2">
                  <div className="text-xs text-gray-500">換算人数</div>
                  <div className="font-medium">{rate} 人</div>
                </div>
                <div className="bg-gray-50 rounded-lg p-2">
                  <div className="text-xs text-gray-500">基本単価</div>
                  <div className="font-medium">¥{fmt(BASE_RATES[emp.employmentType])}</div>
                </div>
                <div className="bg-indigo-50 rounded-lg p-2">
                  <div className="text-xs text-indigo-600">算出額合計</div>
                  <div className="font-bold text-indigo-700">¥{fmt(baseAmount)}</div>
                </div>
              </div>
              {allocations.length > 0 && (
                <div className="space-y-1">
                  {allocations.map((a, i) => (
                    <div key={i} className="flex items-center justify-between text-sm bg-emerald-50 rounded-lg px-3 py-1.5">
                      <span className="text-emerald-800">{a.department}</span>
                      <span className="text-emerald-600 font-medium">{a.ratio}% → ¥{fmt(a.amount)}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          );
        })
      )}
    </div>
  );
}

// ============================================================
// PDF TEXT EXTRACTION using PDF.js
// ============================================================
async function extractTextFromPdf(file) {
  return new Promise((resolve) => {
    const reader = new FileReader();
    reader.onload = async (ev) => {
      try {
        const pdfjsLib = window["pdfjs-dist/build/pdf"];
        pdfjsLib.GlobalWorkerOptions.workerSrc =
          "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
        const typedArray = new Uint8Array(ev.target.result);
        const pdf = await pdfjsLib.getDocument({ data: typedArray }).promise;
        let fullText = "";
        for (let i = 1; i <= pdf.numPages; i++) {
          const page = await pdf.getPage(i);
          const content = await page.getTextContent();
          const items = content.items;

          // 方式1: Y座標でグループ化（通常のPDF）
          const lineMap = {};
          items.forEach(item => {
            if (!item.str || item.str.trim() === "") return;
            const y = Math.round(item.transform[5]);
            if (!lineMap[y]) lineMap[y] = [];
            lineMap[y].push({ x: item.transform[4], str: item.str });
          });
          const sortedYs = Object.keys(lineMap).map(Number).sort((a, b) => b - a);
          const lineTexts = sortedYs.map(y =>
            lineMap[y].sort((a, b) => a.x - b.x).map(i => i.str).join(" ")
          ).filter(l => l.trim());

          // 方式2: テキストをそのまま連結（Y座標が取れない特殊PDFのフォールバック）
          const rawText = items.map(i => i.str).join(" ");

          // 行数が極端に少ない場合（Y座標方式で失敗）はrawTextを行分割して使用
          if (lineTexts.length < 3 && rawText.trim().length > 50) {
            // スペースや区切り文字で行を推定
            const fallbackLines = rawText
              .replace(/([。．！？\.!?])\s*/g, "$1\n")
              .replace(/(\d{3,})\s+/g, "$1\n")
              .split("\n")
              .map(l => l.trim())
              .filter(l => l.length > 0);
            fullText += fallbackLines.join("\n") + "\n";
          } else {
            fullText += lineTexts.join("\n") + "\n";
          }
        }
        resolve(fullText || "（テキスト抽出失敗）");
      } catch (err) {
        resolve("（エラー: " + err.message + "）");
      }
    };
    reader.readAsArrayBuffer(file);
  });
}

// PDFテキストから氏名と金額を解析するロジック（改良版v4）
function parsePdfText(text, masterData) {
  const results = [];
  const rawLines = text.split("\n").map(l => l.trim()).filter(Boolean);

  // 行から金額らしい数値を抽出（3桁以上）
  const extractNums = (str) => {
    const matches = str.match(/[\d,，]{3,}/g) || [];
    return matches.map(n => parseInt(n.replace(/[,，]/g, ""), 10)).filter(n => n >= 100);
  };

  // 金額明細キーワード
  const amountLineKeyword = /時間内|時間外残業|時間外|深夜|早朝|休日|通勤|交通費|交通|基本|普通残|割増|残業|深夜残|手当/;

  // 除外キーワード（合算しない行）
  const excludeKeyword = /消費税|内税|外税|税込|税抜|立替|相殺|業務料合計|取引銀行|振込|通勤手当|当座|銀行|みずほ|三菱|三井|りそな|北洋|千葉|口座名義|登録番号|請求書No|お得意様|ページ|稼働年月|出勤日数|請求内訳|時間内時間外|源泉徴収|源泉税/;

  // 部署名行の判定（「課」「部」「室」「センター」で終わる短い行）
  const isDeptLine = (line) => {
    const t = line.trim();
    return t.length <= 15 && /[課部室係班センター棟]$/.test(t) && !/\d/.test(t);
  };

  // 請求明細行の判定（年月日パターンを含む行を請求明細とみなす）
  // 例: 「伊藤 芳恵 26/03/01～26/03/31 普通 ...」
  const isBillingDetailLine = (line) => /\d{2}\s*\/\s*\d{2}\s*\/\s*\d{2}/.test(line);

  // 勤怠一覧行の除外判定（曜日・出勤・休日などの勤怠キーワードを含む行）
  const isAttendanceLine = (line) => /出勤|休日|有休|欠勤|曜日|開始|終了|休憩|実働|勤怠|稼働年月|スタッフNo|JOB No/.test(line);

  // 勤怠管理番号行の判定（「氏名 数字6桁 数字7桁」のパターン）
  const isAttendanceIdLine = (line) => /[一-龥]{1,4}\s[一-龥]{1,4}\s+\d{6}\s+\d{7}/.test(line);

  // ブロック終端判定（合計・小計行）
  const isBlockEnd = (line) => {
    const t = line.trim();
    // 「小計 消費税 請求金額」のような複合行は終端にしない（単独の合計・小計行のみ）
    const isSimpleTotal = /^合[\s　]*計[\s　]*$/.test(t) || /^小[\s　]*計[\s　]*$/.test(t) ||
           /^合[\s　]*計[\s　]*\d/.test(t) || /^小[\s　]*計[\s　]*\d/.test(t);
    return isSimpleTotal;
  };

  // マスタ氏名を正規化
  const masterNames = masterData.map(m => ({
    ...m,
    norm: normalizeSpaces(m.name),
  })).filter(m => m.norm.length > 0);

  if (masterNames.length > 0) {
    masterNames.forEach(master => {
      // 氏名が含まれる行を収集（勤怠一覧・部署名行は除外）
      const billingLines = []; // 請求明細行（年月日含む・明細キーワード含む）
      const otherLines  = []; // その他の行

      rawLines.forEach((line, idx) => {
        const normLine = normalizeSpaces(line);
        // 完全一致のみ（部分マッチは誤マッチの原因になるため使用しない）
        if (!normLine.includes(master.norm)) return;
        if (isAttendanceLine(line)) return;
        if (isDeptLine(line)) return;
        if (isAttendanceIdLine(line)) return;
        // エキスパート型: スタッフ番号(5〜10桁)で始まる氏名行もbillingLinesに追加
        const isExpertStaffLine = /^\d{5,10}\s+[一-龥]/.test(line);
        if (isBillingDetailLine(line) || amountLineKeyword.test(line) || isExpertStaffLine) {
          billingLines.push(idx);
        } else {
          otherLines.push(idx);
        }
      });

      // 請求明細行を優先、なければその他行を使用
      const targetLines = billingLines.length > 0 ? billingLines : otherLines;
      if (targetLines.length === 0) return;

      let totalAmount = 0;

      targetLines.forEach(startIdx => {
        // ブロック終端を決定（最大15行）
        let endIdx = Math.min(rawLines.length, startIdx + 15);
        for (let i = startIdx + 1; i < endIdx; i++) {
          const l = rawLines[i];
          const normL = normalizeSpaces(l);
          // 別のマスタ氏名が出たら終了
          if (masterNames.some(m => m.norm !== master.norm && normL.includes(m.norm))) {
            endIdx = i; break;
          }
          // 合計・小計行で終了
          if (isBlockEnd(l)) { endIdx = i; break; }
          // スタッフ番号行（エキスパート型の次の人）で終了
          if (/^\d{5,10}\s+[\u4e00-\u9fa5]/.test(l)) { endIdx = i; break; }
        }

        // ブロック内の金額を集計
        const blockLines = rawLines.slice(startIdx, endIdx);
        let blockAmount = 0;

        // PDF全体に「源泉徴収税額」がある場合はrawLines全体から小計+消費税を取得
        // 「源泉税率」などは除外（スタッフサービスの支払依頼書に含まれるため）
        const hasSoukeiInPdf = rawLines.some(l => /源泉徴収税額/.test(l));
        if (hasSoukeiInPdf) {
          let shoukei = 0, zeikin = 0;
          // startIdxより後の全行から小計と消費税を探す
          rawLines.slice(startIdx).forEach(l => {
            const sm = l.match(/^小[\s　]*計[\s　]*([\d,]+)/);
            if (sm) shoukei = parseInt(sm[1].replace(/,/g, ""), 10);
            const tm = l.match(/消費税[^\n]*?(\d[\d,]*)$/);
            if (tm && !/源泉/.test(l)) zeikin = parseInt(tm[1].replace(/,/g, ""), 10);
          });
          if (shoukei > 0) {
            blockAmount = shoukei + zeikin;
            totalAmount += blockAmount;
            return;
          }
        }

        // ¥マーク・円マーク付き金額を優先採用（社内外注型）
        // 社内外注は小額案件があるため1000円以上、派遣は10000円以上
        const isShaInGaichuu = master?.employmentType === "社内外注";
        const yenMinAmt = isShaInGaichuu ? 1000 : 10000;
        let yenAmount = 0;
        blockLines.forEach(line => {
          if (excludeKeyword.test(line) || isBlockEnd(line)) return;
          const yenMatch = line.match(/[¥￥]\s*([\d,，]+)/);
          if (yenMatch) {
            const amt = parseInt(yenMatch[1].replace(/[,，]/g, ""), 10);
            if (amt >= yenMinAmt && amt > yenAmount) yenAmount = amt;
          }
          const enMatches = line.matchAll(/(\d[\d,，]*)\s*円/g);
          for (const m of enMatches) {
            const amt = parseInt(m[1].replace(/[,，]/g, ""), 10);
            if (amt >= yenMinAmt && amt > yenAmount) yenAmount = amt;
          }
        });
        if (yenAmount < 10000) {
          const requestKwLocal = /請求金額|合計金額|ご請求金額|税込合計/;
          for (let bi = 0; bi < blockLines.length; bi++) {
            if (requestKwLocal.test(blockLines[bi])) {
              for (let bj = bi; bj <= Math.min(bi + 3, blockLines.length - 1); bj++) {
                const yenM = blockLines[bj].match(/[¥￥]\s*([\d,，]+)/);
                if (yenM) {
                  const amt = parseInt(yenM[1].replace(/[,，]/g, ""), 10);
                  if (amt >= yenMinAmt) { yenAmount = amt; break; }
                }
                const enM = blockLines[bj].match(/(\d[\d,，]*)\s*円/);
                if (enM) {
                  const amt = parseInt(enM[1].replace(/[,，]/g, ""), 10);
                  if (amt >= yenMinAmt && amt > yenAmount) yenAmount = amt;
                }
              }
            }
          }
        }
        if (yenAmount >= yenMinAmt) {
          blockAmount = yenAmount;
          totalAmount += blockAmount;
          return;
        }

        blockLines.forEach(line => {
          if (excludeKeyword.test(line) || isBlockEnd(line) || isAttendanceLine(line) || isDeptLine(line) || isAttendanceIdLine(line)) return;

          // 小数・契約番号を除いた整数リスト
          const lineNoDecimal = line.replace(/\d+\.\d+/g, "");
          const cleanNums = (lineNoDecimal.match(/[\d,]{3,}/g) || [])
            .map(n => parseInt(n.replace(/,/g, ""), 10))
            .filter(n => n >= 100 && n <= 999999); // 100万未満に制限（勤怠管理番号等を除外）

          if (cleanNums.length === 0) return;

          if (amountLineKeyword.test(line)) {
            const val = cleanNums[cleanNums.length - 1];
            blockAmount += val;
          } else if (cleanNums.length === 1 && cleanNums[0] >= 50000 && cleanNums[0] < 500000) {
            // 単独数値行は5万円以上のみ（小額は書類番号・社員番号として除外）
            blockAmount += cleanNums[0];
          } else if (cleanNums.length >= 2) {
            const val = cleanNums[cleanNums.length - 1];
            if (val >= yenMinAmt) {
              blockAmount += val;
            }
          }
        });

        totalAmount += blockAmount;
      });

      // 金額が取れなかった場合は「今回ご請求額」フォールバックを使用（リクルート型対応）
      const minAmt = master?.employmentType === "社内外注" ? 1000 : 10000;
      if (totalAmount < minAmt) {
        // ￥マーク直後または今回ご請求額キーワード直後の金額を探す
        const requestKw = /今回.*請求額|今回.*ご請求|御請求額|ご請求額/;
        for (let i = 0; i < rawLines.length; i++) {
          const line = rawLines[i];
          if (/^[￥¥$]\s*$/.test(line.trim())) {
            for (let j = i + 1; j <= Math.min(i + 2, rawLines.length - 1); j++) {
              const nums = extractNums(rawLines[j]).filter(n => n >= minAmt);
              if (nums.length > 0) { totalAmount = Math.max(totalAmount, ...nums); break; }
            }
          }
          const yenMatch = line.match(/[￥¥$]\s*([\d,，]+)/);
          if (yenMatch) {
            const amt = parseInt(yenMatch[1].replace(/[,，]/g, ""), 10);
            if (amt >= minAmt && amt > totalAmount) totalAmount = amt;
          }
        }
        if (totalAmount < minAmt) {
          for (let i = 0; i < rawLines.length; i++) {
            if (requestKw.test(rawLines[i])) {
              for (let j = i; j <= Math.min(i + 5, rawLines.length - 1); j++) {
                const nums = extractNums(rawLines[j]).filter(n => n >= minAmt);
                if (nums.length > 0) { totalAmount = Math.max(totalAmount, ...nums); break; }
              }
            }
          }
        }
      }
      if (totalAmount >= (master?.employmentType === "社内外注" ? 1000 : 10000)) {
        results.push({ name: master.name, amount: totalAmount, matched: true, master });
      }
    });
  }

  // マスタ照合済みだが金額が取れなかった人のフォールバック
  // 「今回ご請求額」直後の金額で補完（リクルート型対応）
  if (masterNames.length > 0 && results.length === 0) {
    // ￥マークまたは今回ご請求額キーワードから金額を取得
    let fallbackAmount = 0;
    const requestKeyword2 = /今回.*請求額|今回.*ご請求|御請求額|ご請求額/;
    for (let i = 0; i < rawLines.length; i++) {
      const line = rawLines[i];
      if (/^[￥¥$]\s*$/.test(line.trim())) {
        for (let j = i + 1; j <= Math.min(i + 2, rawLines.length - 1); j++) {
          const nums = extractNums(rawLines[j]).filter(n => n >= 1000);
          if (nums.length > 0) { fallbackAmount = Math.max(fallbackAmount, ...nums); break; }
        }
      }
      const yenMatch = line.match(/[￥¥$]\s*([\d,，]+)/);
      if (yenMatch) {
        const amt = parseInt(yenMatch[1].replace(/[,，]/g, ""), 10);
        if (amt >= 1000 && amt > fallbackAmount) fallbackAmount = amt;
      }
    }
    if (fallbackAmount === 0) {
      for (let i = 0; i < rawLines.length; i++) {
        if (requestKeyword2.test(rawLines[i])) {
          for (let j = i; j <= Math.min(i + 5, rawLines.length - 1); j++) {
            const nums = extractNums(rawLines[j]).filter(n => n >= 1000);
            if (nums.length > 0) { fallbackAmount = Math.max(fallbackAmount, ...nums); break; }
          }
        }
      }
    }
    if (fallbackAmount > 0) {
      // マスタ内で最初の人に紐づける（1人しかいない場合）
      // 複数人いる場合は手動修正に委ねる
      if (masterNames.length === 1) {
        results.push({ name: masterNames[0].name, amount: fallbackAmount, matched: true, master: masterNames[0] });
      } else {
        results.push({ name: "", amount: fallbackAmount, matched: false, master: null });
      }
    }
  }

  // フォールバック: マスタ未登録時、氏名パターンで検索
  if (results.length === 0) {
    const namePattern = /[\u4e00-\u9fa5]{1,4}[\s\u3000][\u4e00-\u9fa5]{1,4}/g;
    // 氏名として不適切な一般語を除外
    const invalidNameWords = /合計|請求|消費|内訳|小計|明細|御中|様|会社|株式|品目|単価|数量|価格|納品|原稿|整理|作成|備考|チェック|件数|振込|支払|登録|番号|氏名|住所|銀行|口座|支店|期日|摘要|金額|基本|超過|不足|交通|出張|経費|課税|消費|対象|内訳|手当|通信|業務|インフラ|提供|受注|機種|代表|事業|部門|担当|締め|翌月|登録番号|請求書|年月日|件名|補足|その他/;
    const foundNames = new Set();
    rawLines.forEach((line, startIdx) => {
      (line.match(namePattern) || []).forEach(rawName => {
        const name = rawName.trim();
        if (foundNames.has(name) || invalidNameWords.test(name)) return;
        foundNames.add(name);
        let endIdx = Math.min(rawLines.length, startIdx + 20);
        for (let i = startIdx + 1; i < endIdx; i++) {
          if (isBlockEnd(rawLines[i])) { endIdx = i; break; }
        }
        let amount = 0;
        rawLines.slice(startIdx, endIdx).forEach(l => {
          if (excludeKeyword.test(l) || isBlockEnd(l)) return;
          const nums = extractNums(l);
          if (amountLineKeyword.test(l) && nums.length > 0) amount += nums[nums.length - 1];
        });
        if (amount === 0) {
          const allNums = rawLines.slice(startIdx, endIdx).filter(l => !excludeKeyword.test(l)).flatMap(l => extractNums(l));
          if (allNums.length > 0) amount = Math.max(...allNums);
        }
        if (amount > 0) results.push({ name, amount, matched: false, master: null });
      });
    });
  }

  // 最終フォールバック: 御請求額キーワードで金額を取得
  if (results.length === 0) {
    // 「今回ご請求額」「御請求額」などのキーワードの直後行から金額を探す
    const requestKeyword = /今回.*請求額|今回.*ご請求|御請求額|ご請求額|請求金額|合計金額|支払額/;
    let foundAmount = 0;

    // まず全文から「￥」マークの直後にある金額を探す（リクルート型対応）
    for (let i = 0; i < rawLines.length; i++) {
      const line = rawLines[i];
      // 「￥」単体行の次の行が金額
      if (/^[￥¥$]\s*$/.test(line.trim())) {
        for (let j = i + 1; j <= Math.min(i + 2, rawLines.length - 1); j++) {
          const nums = extractNums(rawLines[j]).filter(n => n >= 1000);
          if (nums.length > 0) {
            foundAmount = Math.max(foundAmount, ...nums);
            break;
          }
        }
      }
      // 「￥528,512」のように同行に金額がある場合
      const yenMatch = line.match(/[￥¥$]\s*([\d,，]+)/);
      if (yenMatch) {
        const amt = parseInt(yenMatch[1].replace(/[,，]/g, ""), 10);
        if (amt >= 1000 && amt > foundAmount) foundAmount = amt;
      }
    }

    // 次に「今回ご請求額」等キーワード直後の金額を探す
    if (foundAmount === 0) {
      for (let i = 0; i < rawLines.length; i++) {
        if (requestKeyword.test(rawLines[i])) {
          for (let j = i; j <= Math.min(i + 5, rawLines.length - 1); j++) {
            const nums = extractNums(rawLines[j]).filter(n => n >= 1000);
            if (nums.length > 0) {
              foundAmount = Math.max(foundAmount, ...nums);
              break;
            }
          }
        }
      }
    }

    // キーワードで見つからない場合は最大金額をフォールバック
    if (foundAmount === 0) {
      rawLines.forEach(line => {
        if (/御請求|ご請求|請求金額|総合計|合計金額/.test(line)) {
          const nums = extractNums(line);
          if (nums.length > 0) foundAmount = Math.max(foundAmount, ...nums);
        }
      });
    }

    if (foundAmount > 0) results.push({ name: "", amount: foundAmount, matched: false, master: null });
  }

  // 同一人物（氏名一致）の金額を合算して1件にまとめる
  const mergedResults = [];
  results.forEach(r => {
    const existing = mergedResults.find(m => m.name === r.name && m.matched === r.matched);
    if (existing) {
      existing.amount += r.amount;
    } else {
      mergedResults.push({ ...r });
    }
  });

  return mergedResults;
}
// ============================================================
// TAB: インポート準備
// ============================================================
function ImportPrepTab() {
  const [masterData, setMasterData] = useState([]);
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(false);
  const [status, setStatus] = useState("");
  const [debugText, setDebugText] = useState(""); // デバッグ用
  const [showDebug, setShowDebug] = useState(false);
  const masterRef = useRef();
  const pdfRef = useRef();

  const handleMasterImport = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      const wb = XLSX.read(ev.target.result, { type: "binary" });
      const ws = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { defval: "" });
      const parsed = data.map(r => {
        const dept =
          r["部門"] || r["部門1"] || r["所属部門"] || r["配賦先"] ||
          r["配賦先部門"] || r["部署"] || r["部署名"] || r["部門名"] || "";
        return {
          name: String(r["氏名"] || ""),
          employmentType: String(r["雇用形態"] || ""),
          department: String(dept),
        };
      }).filter(r => r.name);
      setMasterData(parsed);
      const cols = data.length > 0 ? Object.keys(data[0]).join("、") : "";
      setStatus(`マスタ読込完了: ${parsed.length}件　検出列名: ${cols}`);
    };
    reader.readAsBinaryString(file);
    e.target.value = "";
  };

  const handlePdfImport = async (e) => {
    const files = Array.from(e.target.files);
    if (!files.length) return;
    setLoading(true);
    const newRows = [];
    for (const file of files) {
      setStatus(`解析中: ${file.name}`);
      const text = await extractTextFromPdf(file);
      setDebugText(prev => prev + "\n\n=== " + file.name + " ===\n" + text);
      const parsed = parsePdfText(text, masterData);
      parsed.forEach(item => {
        const extractedName = item.name;
        const amount = item.amount;
        const master = item.master;
        const norm = normalizeSpaces(extractedName);
        const match = master || masterData.find(m => normalizeSpaces(m.name) === norm);
        
        // 同一氏名がすでにnewRowsにある場合は金額を合算（複数PDF対応）
        const existingRow = match ? newRows.find(r => normalizeSpaces(r.name) === norm && r.matched) : null;
        if (existingRow) {
          const addAmount = amount;
          if (match?.employmentType === "社内外注") {
            existingRow.criteriaValue = Math.round((existingRow.amount + addAmount) / 1.1);
          } else if (match?.employmentType === "派遣社員") {
            existingRow.criteriaValue = existingRow.taxType === "taxIncluded"
              ? existingRow.amount + addAmount
              : Math.round((existingRow.amount + addAmount) * 1.1);
          }
          existingRow.amount += addAmount;
          existingRow.fileName += ", " + file.name;
          return; // 新しい行は追加しない
        }
        // 社内外注: PDF金額は税込 → 税抜（÷1.1）を判定基準値に
        // 派遣社員: PDF金額は税抜 → 税込（×1.1）を判定基準値に
        // その他: そのまま
        // 社内外注: PDF金額は税込 → 税抜（÷1.1）を判定基準値に
        // 派遣社員: デフォルト税抜として税込（×1.1）を判定基準値に
        let criteriaValue = amount;
        let taxType = "taxExcluded"; // 派遣社員のPDF記載形式: taxExcluded=税抜記載, taxIncluded=税込記載
        if (match?.employmentType === "社内外注") {
          criteriaValue = Math.round(amount / 1.1);
        } else if (match?.employmentType === "派遣社員") {
          criteriaValue = Math.round(amount * 1.1); // 税抜記載→税込に変換
        }
        newRows.push({
          id: Math.random().toString(36).slice(2),
          extractedName,
          amount,
          name: match ? match.name : extractedName,
          employmentType: match ? match.employmentType : "",
          department: match ? match.department : "",
          criteriaValue,
          taxType, // 派遣社員のPDF記載形式
          matched: !!match,
          fileName: file.name,
        });
      });
      if (parsed.length === 0) {
        // 解析できなかった場合も行を追加（手入力できるように）
        newRows.push({
          id: Math.random().toString(36).slice(2),
          extractedName: "（抽出できませんでした）",
          amount: 0,
          name: "",
          employmentType: "",
          department: "",
          criteriaValue: 0,
          matched: false,
          fileName: file.name,
        });
      }
    }
    setRows(prev => [...prev, ...newRows]);
    setLoading(false);
    setStatus(`解析完了: ${newRows.length}件抽出`);
    e.target.value = "";
  };

  const updateRow = (id, field, value) => {
    setRows(prev => prev.map(r => {
      if (r.id !== id) return r;
      const updated = { ...r, [field]: value };
      // マスタ再照合
      if (field === "name") {
        const norm = normalizeSpaces(value);
        const match = masterData.find(m => normalizeSpaces(m.name) === norm);
        updated.matched = !!match;
        if (match) {
          updated.employmentType = match.employmentType;
          updated.department = match.department;
        }
      }
      // 判定基準値の自動計算
      if (field === "amount") {
        const amt = parseFloat(value) || 0;
        if (updated.employmentType === "社内外注") {
          updated.criteriaValue = Math.round(amt / 1.1); // 税込→税抜
        } else if (updated.employmentType === "派遣社員") {
          // taxType に応じて変換
          updated.criteriaValue = updated.taxType === "taxIncluded"
            ? amt  // 税込記載→そのまま
            : Math.round(amt * 1.1); // 税抜記載→税込に変換
        }
      }
      // 派遣社員のtaxType切り替え時に criteriaValue を再計算
      if (field === "taxType" && r.employmentType === "派遣社員") {
        const amt = parseFloat(r.amount) || 0;
        updated.criteriaValue = value === "taxIncluded"
          ? amt  // 税込記載→そのまま
          : Math.round(amt * 1.1); // 税抜記載→税込に変換
      }
      return updated;
    }));
  };

  const handleExport = () => {
    const data = rows.map(r => ({
      氏名: r.name,
      雇用形態: r.employmentType,
      判定基準値: r.criteriaValue,
      部門1: r.department,
      割合1: 100,
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "インポート用");
    XLSX.writeFile(wb, "インポート用データ.xlsx");
  };

  return (
    <div className="space-y-5">
      {/* STEP 1 */}
      <div className="border border-gray-200 rounded-xl p-4 bg-white shadow-sm">
        <div className="flex items-center gap-2 mb-3">
          <span className="w-6 h-6 rounded-full bg-indigo-600 text-white text-xs flex items-center justify-center font-bold">1</span>
          <span className="font-semibold text-gray-800">マスタ読み込み</span>
          {masterData.length > 0 && <span className="text-xs text-emerald-600 bg-emerald-50 px-2 py-0.5 rounded-full">✓ {masterData.length}件読込済</span>}
        </div>
        <p className="text-sm text-gray-500 mb-1">氏名・雇用形態・部門が記載されたExcelファイルをアップロードしてください。</p>
        <p className="text-xs text-amber-600 bg-amber-50 rounded px-2 py-1 mb-3">
          ⚠ 部門の列名は「部門」「部門1」「所属部門」「配賦先」「部署」「部門名」のいずれかに対応しています。
        </p>
        <button onClick={() => masterRef.current.click()}
          className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 text-sm rounded-lg transition-colors">
          📂 マスタExcelを選択
        </button>
        <input ref={masterRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={handleMasterImport} />
        {masterData.length > 0 && (
          <div className="mt-3 max-h-32 overflow-y-auto text-xs">
            <table className="w-full border-collapse">
              <thead><tr className="bg-gray-50"><th className="px-2 py-1 text-left border-b">氏名</th><th className="px-2 py-1 text-left border-b">雇用形態</th><th className="px-2 py-1 text-left border-b">部門</th></tr></thead>
              <tbody>{masterData.map((m, i) => <tr key={i} className="border-b border-gray-100"><td className="px-2 py-1">{m.name}</td><td className="px-2 py-1">{m.employmentType}</td><td className="px-2 py-1">{m.department}</td></tr>)}</tbody>
            </table>
          </div>
        )}
      </div>

      {/* STEP 2 */}
      <div className="border border-gray-200 rounded-xl p-4 bg-white shadow-sm">
        <div className="flex items-center gap-2 mb-3">
          <span className="w-6 h-6 rounded-full bg-indigo-600 text-white text-xs flex items-center justify-center font-bold">2</span>
          <span className="font-semibold text-gray-800">請求書PDF解析</span>
        </div>
        <p className="text-sm text-gray-500 mb-3">複数のPDFを一括アップロードして氏名と請求金額を自動抽出します。</p>
        <button onClick={() => pdfRef.current.click()} disabled={loading}
          className="px-4 py-2 bg-amber-500 hover:bg-amber-600 text-white text-sm rounded-lg transition-colors disabled:opacity-50">
          {loading ? "⏳ 解析中..." : "📄 PDFを選択（複数可）"}
        </button>
        <input ref={pdfRef} type="file" accept=".pdf" multiple className="hidden" onChange={handlePdfImport} />
        {status && <p className="mt-2 text-sm text-gray-500">{status}</p>}
        {debugText && (
          <div className="mt-3">
            <button onClick={() => setShowDebug(!showDebug)}
              className="text-xs text-indigo-600 hover:underline">
              {showDebug ? "▲ PDFテキスト抽出結果を隠す" : "▼ PDFテキスト抽出結果を確認する（デバッグ用）"}
            </button>
            {showDebug && (
              <textarea readOnly value={debugText}
                className="mt-2 w-full h-64 text-xs font-mono border border-gray-200 rounded p-2 bg-gray-50"
                onClick={e => e.target.select()}
              />
            )}
          </div>
        )}
      </div>

      {/* STEP 3 */}
      {rows.length > 0 && (
        <div className="border border-gray-200 rounded-xl p-4 bg-white shadow-sm">
          <div className="flex items-center justify-between mb-3">
            <div className="flex items-center gap-2">
              <span className="w-6 h-6 rounded-full bg-indigo-600 text-white text-xs flex items-center justify-center font-bold">3</span>
              <span className="font-semibold text-gray-800">確認・修正</span>
            </div>
            <div className="flex gap-2">
              <button onClick={() => setRows(prev => [...prev, {
                id: Math.random().toString(36).slice(2),
                extractedName: "（手入力）",
                amount: 0,
                name: "",
                employmentType: "派遣社員",
                department: "",
                criteriaValue: 0,
                taxType: "taxExcluded",
                matched: false,
                fileName: "手入力",
              }])}
                className="flex items-center gap-1.5 px-3 py-2 bg-indigo-100 text-indigo-700 text-sm rounded-lg hover:bg-indigo-200 transition-colors">
                ＋ 行追加
              </button>
              <button onClick={handleExport}
                className="flex items-center gap-1.5 px-3 py-2 bg-emerald-600 text-white text-sm rounded-lg hover:bg-emerald-700 transition-colors">
                📥 Excel出力
              </button>
            </div>
          </div>
          <p className="text-sm text-gray-500 mb-3">
            <span className="text-red-500 font-medium">赤背景</span>はマスタと不一致の行です。氏名を修正してください。
          </p>
          <div className="overflow-x-auto">
            <table className="w-full text-sm border-collapse">
              <thead>
                <tr className="bg-gray-50">
                  <th className="px-3 py-2 text-left border-b font-semibold text-gray-600">ファイル</th>
                  <th className="px-3 py-2 text-left border-b font-semibold text-gray-600">抽出氏名</th>
                  <th className="px-3 py-2 text-left border-b font-semibold text-gray-600">氏名（修正可）</th>
                  <th className="px-3 py-2 text-left border-b font-semibold text-gray-600">請求金額<br/><span className="text-xs font-normal text-gray-400">（派遣:税抜/税込を選択）</span></th>
                  <th className="px-3 py-2 text-left border-b font-semibold text-gray-600">雇用形態</th>
                  <th className="px-3 py-2 text-left border-b font-semibold text-gray-600">判定基準値<br/><span className="text-xs font-normal text-gray-400">社内外注:税抜 / 派遣:税込</span></th>
                  <th className="px-3 py-2 text-left border-b font-semibold text-gray-600">部門</th>
                  <th className="px-3 py-2 border-b" />
                </tr>
              </thead>
              <tbody>
                {rows.map(row => (
                  <tr key={row.id} className={row.matched ? "border-b border-gray-100" : "border-b border-red-200 bg-red-50"}>
                    <td className="px-3 py-2 text-xs text-gray-400 max-w-[100px] truncate">{row.fileName}</td>
                    <td className="px-3 py-2 text-gray-500 text-xs">{row.extractedName}</td>
                    <td className="px-3 py-2">
                      <input value={row.name} onChange={e => updateRow(row.id, "name", e.target.value)}
                        className={`border rounded px-2 py-1 text-sm w-28 focus:ring-2 outline-none ${row.matched ? "border-gray-200 focus:ring-indigo-300" : "border-red-300 bg-red-50 focus:ring-red-300"}`} />
                    </td>
                    <td className="px-3 py-2">
                      <input type="number" value={row.amount} onChange={e => updateRow(row.id, "amount", e.target.value)}
                        className="border border-gray-200 rounded px-2 py-1 text-sm w-28 focus:ring-2 focus:ring-indigo-300 outline-none" />
                    </td>
                    <td className="px-3 py-2">
                      <select value={row.employmentType} onChange={e => updateRow(row.id, "employmentType", e.target.value)}
                        className="border border-gray-200 rounded px-2 py-1 text-sm focus:ring-2 focus:ring-indigo-300 outline-none">
                        <option value="">選択</option>
                        {EMPLOYMENT_TYPES.map(t => <option key={t}>{t}</option>)}
                      </select>
                    </td>
                    <td className="px-3 py-2">
                      {row.employmentType === "派遣社員" ? (
                        <div className="space-y-1">
                          {/* 税抜/税込 切り替えボタン */}
                          <div className="flex rounded-lg overflow-hidden border border-gray-300 w-fit text-xs">
                            <button
                              onClick={() => updateRow(row.id, "taxType", "taxExcluded")}
                              className={`px-2 py-1 transition-colors ${(row.taxType || "taxExcluded") === "taxExcluded" ? "bg-indigo-600 text-white" : "bg-white text-gray-600 hover:bg-gray-50"}`}>
                              税抜記載
                            </button>
                            <button
                              onClick={() => updateRow(row.id, "taxType", "taxIncluded")}
                              className={`px-2 py-1 transition-colors ${row.taxType === "taxIncluded" ? "bg-indigo-600 text-white" : "bg-white text-gray-600 hover:bg-gray-50"}`}>
                              税込記載
                            </button>
                          </div>
                          <div className="text-indigo-700 font-medium text-xs">
                            判定基準値(税込): ¥{fmt(row.criteriaValue)}
                          </div>
                        </div>
                      ) : row.employmentType === "社内外注" ? (
                        <div>
                          <div className="text-indigo-700 font-medium">¥{fmt(row.criteriaValue)}</div>
                          <div className="text-xs text-gray-400">税抜換算</div>
                        </div>
                      ) : (
                        <input type="number" value={row.criteriaValue} onChange={e => updateRow(row.id, "criteriaValue", e.target.value)}
                          className="border border-gray-200 rounded px-2 py-1 text-sm w-24 focus:ring-2 focus:ring-indigo-300 outline-none" />
                      )}
                    </td>
                    <td className="px-3 py-2">
                      <input value={row.department} onChange={e => updateRow(row.id, "department", e.target.value)}
                        className="border border-gray-200 rounded px-2 py-1 text-sm w-24 focus:ring-2 focus:ring-indigo-300 outline-none" />
                    </td>
                    <td className="px-3 py-2">
                      <button onClick={() => setRows(prev => prev.filter(r => r.id !== row.id))} className="text-red-400 hover:text-red-600">×</button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}

// ============================================================
// MAIN APP
// ============================================================
export default function App() {
  const [tab, setTab] = useState("input");
  const [employees, setEmployees] = useState([]);

  const tabs = [
    { id: "input", label: "📝 データ入力" },
    { id: "dept", label: "📊 部門別集計" },
    { id: "individual", label: "👤 個人別明細" },
    { id: "prep", label: "🔧 インポート準備" },
  ];

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-indigo-50 font-sans">
      {/* Header */}
      <div className="bg-white border-b border-gray-200 shadow-sm">
        <div className="max-w-5xl mx-auto px-4 py-4">
          <h1 className="text-xl font-bold text-gray-800">間接費用配賦額 計算ツール</h1>
          <p className="text-xs text-gray-400 mt-0.5">従業員データに基づく部門別請求額の自動算出</p>
        </div>
      </div>

      {/* Tabs */}
      <div className="bg-white border-b border-gray-200">
        <div className="max-w-5xl mx-auto px-4">
          <div className="flex gap-0">
            {tabs.map(t => (
              <button key={t.id} onClick={() => setTab(t.id)}
                className={`px-5 py-3 text-sm font-medium border-b-2 transition-colors ${tab === t.id ? "border-indigo-600 text-indigo-600" : "border-transparent text-gray-500 hover:text-gray-700"}`}>
                {t.label}
              </button>
            ))}
          </div>
        </div>
      </div>

      {/* Content */}
      <div className="max-w-5xl mx-auto px-4 py-6">
        {tab === "input" && <DataInputTab employees={employees} setEmployees={setEmployees} />}
        {tab === "dept" && <DeptSummaryTab employees={employees} />}
        {tab === "individual" && <IndividualTab employees={employees} />}
        {tab === "prep" && <ImportPrepTab />}
      </div>
    </div>
  );
}
