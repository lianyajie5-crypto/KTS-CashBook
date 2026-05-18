import React, { useEffect, useMemo, useState } from "react";
import ExcelJS from "exceljs";

const CATEGORIES = ["员工工资","办公费用","车辆燃油","车辆维修","员工福利","劳保用品","低值易耗","办公维修","固定资产","差旅费用","业务招待","劳务支出","咨询服务","在建工程","通讯费用","水电气费","开办费用","租金费用","其他支出","借款","备用金"];
const DEFAULT_HANDLERS = ["连雅杰", "袁杰"];
const STORAGE_KEY = "pkr_cashbook_records_v3";
const CASH_KEY = "pkr_cashbook_initial_cash_v3";
const HANDLERS_KEY = "pkr_cashbook_handlers_v3";

function todayISO() { return new Date().toISOString().slice(0, 10); }
function money(value) { return Number(value || 0).toLocaleString("en-US"); }

async function downloadExcelJSWorkbook(filename, workbook) {
  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(url);
}
function copyRowStyle(sourceRow, targetRow) {
  targetRow.height = sourceRow.height;
  sourceRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    const targetCell = targetRow.getCell(colNumber);
    targetCell.style = JSON.parse(JSON.stringify(cell.style || {}));
    targetCell.numFmt = cell.numFmt;
  });
}
function setCell(sheet, address, value) { sheet.getCell(address).value = value; }

async function buildWorkbookFromEmbeddedTemplate(recordsForDay, dateFilter, totalExpense, invoicedAmount, pendingAmount) {
  const response = await fetch("/templates/payment_approval_template.xlsx");
  if (!response.ok) throw new Error("没有找到内置模板文件：public/templates/payment_approval_template.xlsx");
  const arrayBuffer = await response.arrayBuffer();
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  const voucherSheet = workbook.getWorksheet("付款审批单");
  const listSheet = workbook.getWorksheet("清单");
  if (!voucherSheet || !listSheet) throw new Error("模板中必须包含工作表：付款审批单、清单");

  setCell(voucherSheet, "J7", dateFilter);
  setCell(voucherSheet, "B9", "现金支出");
  setCell(voucherSheet, "C9", `${dateFilter} 项目现场现金支出，详见发票清单`);
  setCell(voucherSheet, "H9", Number(totalExpense || 0));
  setCell(voucherSheet, "J9", `已收票：${money(invoicedAmount)}；未收票：${money(pendingAmount)}`);
  setCell(voucherSheet, "H11", Number(totalExpense || 0));
  setCell(voucherSheet, "E12", "PKR/卢比");
  setCell(voucherSheet, "C15", "连雅杰");

  const templateStartRow = 4;
  const templateEndRow = 18;
  const templateCapacity = templateEndRow - templateStartRow + 1;
  if (recordsForDay.length > templateCapacity) {
    const extraRows = recordsForDay.length - templateCapacity;
    listSheet.spliceRows(templateEndRow + 1, 0, ...Array.from({ length: extraRows }, () => []));
    const styleSourceRow = listSheet.getRow(templateEndRow);
    for (let i = 0; i < extraRows; i += 1) copyRowStyle(styleSourceRow, listSheet.getRow(templateEndRow + 1 + i));
  }
  const totalRowNumber = templateEndRow + Math.max(0, recordsForDay.length - templateCapacity) + 1;
  for (let r = templateStartRow; r < totalRowNumber; r += 1) {
    for (let c = 2; c <= 8; c += 1) listSheet.getRow(r).getCell(c).value = "";
  }
  recordsForDay.forEach((item, index) => {
    const row = listSheet.getRow(templateStartRow + index);
    row.getCell(1).value = index + 1;
    row.getCell(2).value = item.date;
    row.getCell(3).value = item.category;
    row.getCell(4).value = item.note || item.category;
    row.getCell(5).value = "PKR/卢比";
    row.getCell(6).value = Number(item.amount || 0);
    row.getCell(7).value = item.handler;
    row.getCell(8).value = item.invoiced ? "已收票" : "未收票";
  });
  const visibleRows = Math.max(recordsForDay.length, templateCapacity);
  for (let i = recordsForDay.length; i < visibleRows; i += 1) {
    const row = listSheet.getRow(templateStartRow + i);
    row.getCell(1).value = i + 1;
    for (let c = 2; c <= 8; c += 1) row.getCell(c).value = "";
  }
  listSheet.getCell(`D${totalRowNumber}`).value = "合计";
  listSheet.getCell(`F${totalRowNumber}`).value = Number(totalExpense || 0);
  return workbook;
}

export default function PKRCashBookApp() {
  const [records, setRecords] = useState([]);
  const [initialCash, setInitialCash] = useState(500000);
  const [handlers, setHandlers] = useState(DEFAULT_HANDLERS);
  const [dateFilter, setDateFilter] = useState(todayISO());
  const [form, setForm] = useState({ date: todayISO(), amount: "", handler: "连雅杰", category: "车辆燃油", note: "", invoiced: false, receiptPhoto: "" });

  useEffect(() => {
    try {
      setRecords(JSON.parse(localStorage.getItem(STORAGE_KEY) || "[]"));
      setInitialCash(Number(localStorage.getItem(CASH_KEY) || 500000));
      const savedHandlers = JSON.parse(localStorage.getItem(HANDLERS_KEY) || "null");
      setHandlers(Array.isArray(savedHandlers) && savedHandlers.length ? savedHandlers : DEFAULT_HANDLERS);
    } catch { setRecords([]); setInitialCash(500000); setHandlers(DEFAULT_HANDLERS); }
  }, []);
  useEffect(() => localStorage.setItem(STORAGE_KEY, JSON.stringify(records)), [records]);
  useEffect(() => localStorage.setItem(CASH_KEY, String(initialCash)), [initialCash]);
  useEffect(() => localStorage.setItem(HANDLERS_KEY, JSON.stringify(handlers)), [handlers]);

  const filteredRecords = useMemo(() => records.filter(item => item.date === dateFilter), [records, dateFilter]);
  const totalExpense = filteredRecords.reduce((sum, item) => sum + Number(item.amount || 0), 0);
  const invoicedAmount = filteredRecords.filter(item => item.invoiced).reduce((sum, item) => sum + Number(item.amount || 0), 0);
  const pendingAmount = totalExpense - invoicedAmount;
  const allExpense = records.reduce((sum, item) => sum + Number(item.amount || 0), 0);
  const cashBalance = Number(initialCash || 0) - allExpense;
  const receiptRate = totalExpense > 0 ? Math.round((invoicedAmount / totalExpense) * 100) : 0;
  const handlerStats = filteredRecords.reduce((acc, item) => { acc[item.handler || "未填写"] = (acc[item.handler || "未填写"] || 0) + Number(item.amount || 0); return acc; }, {});
  const categoryStats = filteredRecords.reduce((acc, item) => { acc[item.category || "其他支出"] = (acc[item.category || "其他支出"] || 0) + Number(item.amount || 0); return acc; }, {});

  function updateForm(key, value) { setForm(prev => ({ ...prev, [key]: value })); }
  function handleReceiptPhoto(file) {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => updateForm("receiptPhoto", String(reader.result || ""));
    reader.readAsDataURL(file);
  }
  function saveRecord() {
    const amount = Number(form.amount);
    const handler = form.handler.trim();
    if (!amount || amount <= 0) return alert("请输入正确金额");
    if (!handler) return alert("请输入经办人");
    if (!handlers.includes(handler)) setHandlers(prev => [handler, ...prev]);
    const newRecord = { id: Date.now(), date: form.date || todayISO(), amount, handler, category: form.category, note: form.note, invoiced: form.invoiced || Boolean(form.receiptPhoto), receiptPhoto: form.receiptPhoto, createdAt: new Date().toISOString() };
    setRecords(prev => [newRecord, ...prev]);
    setDateFilter(newRecord.date);
    setForm(prev => ({ date: newRecord.date, amount: "", handler, category: prev.category, note: "", invoiced: false, receiptPhoto: "" }));
  }
  function markInvoiced(id) { setRecords(prev => prev.map(item => item.id === id ? { ...item, invoiced: true } : item)); }
  function deleteRecord(id) { if (confirm("确认删除这条支出记录？")) setRecords(prev => prev.filter(item => item.id !== id)); }
  async function exportDailyApprovalXLSX() {
    try {
      const wb = await buildWorkbookFromEmbeddedTemplate(filteredRecords, dateFilter, totalExpense, invoicedAmount, pendingAmount);
      await downloadExcelJSWorkbook(`付款审批单_发票清单_${dateFilter}.xlsx`, wb);
    } catch (error) { alert(error.message || "导出Excel失败，请检查模板文件。"); }
  }
  async function exportAllXLSX() {
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet("现金总台账");
    ws.addRow(["序号", "日期", "经办人", "类别", "金额", "支付方式", "币种", "票据状态", "是否有票据照片", "备注"]);
    records.forEach((item, index) => ws.addRow([index + 1, item.date, item.handler, item.category, Number(item.amount || 0), "现金", "PKR/卢比", item.invoiced ? "已收票" : "未收票", item.receiptPhoto ? "有" : "无", item.note || ""]));
    ws.columns = [{ width: 8 }, { width: 14 }, { width: 14 }, { width: 16 }, { width: 14 }, { width: 12 }, { width: 12 }, { width: 12 }, { width: 16 }, { width: 28 }];
    ws.getRow(1).font = { bold: true };
    await downloadExcelJSWorkbook(`PKR现金总台账_${todayISO()}.xlsx`, wb);
  }

  return <div className="min-h-screen bg-zinc-950 text-white p-4 max-w-md mx-auto select-none">
    <div className="flex items-center justify-between mb-5"><div><h1 className="text-2xl font-bold">PKR 现金记账</h1><p className="text-zinc-400 text-sm">纯本地离线版 · 数据保存在本机</p></div><div className="bg-yellow-500 text-black px-3 py-1 rounded-xl text-sm font-bold">PKR</div></div>
    <div className="bg-yellow-500 text-black rounded-3xl p-5 mb-5 shadow-2xl"><div className="text-sm mb-2 font-semibold">今日现金流出</div><div className="text-4xl font-black mb-2">{money(totalExpense)} PKR</div><div className="flex items-center justify-between text-sm"><span>现金笔数：{filteredRecords.length} 笔</span><span>收票率：{receiptRate}%</span></div></div>
    <div className="grid grid-cols-3 gap-3 mb-5"><div className="bg-zinc-900 rounded-2xl p-3 border border-zinc-800"><div className="text-zinc-400 text-xs mb-1">已收票</div><div className="font-bold text-lg text-green-400">{money(invoicedAmount)}</div></div><div className="bg-zinc-900 rounded-2xl p-3 border border-zinc-800"><div className="text-zinc-400 text-xs mb-1">未收票</div><div className="font-bold text-lg text-red-400">{money(pendingAmount)}</div></div><div className="bg-zinc-900 rounded-2xl p-3 border border-zinc-800"><div className="text-zinc-400 text-xs mb-1">余额</div><div className="font-bold text-lg text-yellow-400">{money(cashBalance)}</div></div></div>
    <div className="bg-zinc-900 border border-zinc-800 rounded-2xl p-4 mb-5"><div className="flex items-center justify-between mb-3"><div><div className="text-lg font-semibold">备用金余额</div><div className="text-xs text-zinc-500">初始备用金 - 全部现金支出</div></div><div className="text-right"><div className="text-green-400 font-bold text-xl">{money(cashBalance)}</div><div className="text-xs text-zinc-500">PKR</div></div></div><input type="number" className="w-full bg-zinc-800 rounded-xl p-3 outline-none text-lg" value={initialCash} onChange={e => setInitialCash(Number(e.target.value || 0))} placeholder="输入初始备用金" /></div>
    <div className="bg-zinc-900 border border-zinc-800 rounded-2xl p-4 mb-5"><div className="text-lg font-semibold mb-4">新增支出</div><div className="space-y-3"><input type="date" className="w-full bg-zinc-800 rounded-xl p-4 outline-none" value={form.date} onChange={e => updateForm("date", e.target.value)} /><input type="number" className="w-full bg-zinc-800 rounded-xl p-4 text-2xl font-bold outline-none" placeholder="输入金额 PKR" value={form.amount} onChange={e => updateForm("amount", e.target.value)} /><input list="handler-list" className="w-full bg-zinc-800 rounded-xl p-4 outline-none" placeholder="经办人" value={form.handler} onChange={e => updateForm("handler", e.target.value)} /><datalist id="handler-list">{handlers.map(item => <option key={item} value={item} />)}</datalist><select className="w-full bg-zinc-800 rounded-xl p-4 outline-none" value={form.category} onChange={e => updateForm("category", e.target.value)}>{CATEGORIES.map(item => <option key={item}>{item}</option>)}</select><input className="w-full bg-zinc-800 rounded-xl p-4 outline-none" placeholder="备注，例如：柴油 61L" value={form.note} onChange={e => updateForm("note", e.target.value)} /><label className="flex items-center gap-3 bg-zinc-800 rounded-xl p-4"><input type="checkbox" className="w-5 h-5" checked={form.invoiced} onChange={e => updateForm("invoiced", e.target.checked)} /><span>已收到票据</span></label><div className="bg-zinc-800 rounded-xl p-4"><div className="text-sm mb-2 text-zinc-300">拍照/上传票据</div><input type="file" accept="image/*" capture="environment" className="w-full text-sm" onChange={e => handleReceiptPhoto(e.target.files?.[0])} />{form.receiptPhoto && <img src={form.receiptPhoto} alt="票据预览" className="mt-3 rounded-xl max-h-48 w-full object-cover" />}</div><button className="w-full bg-green-600 py-4 rounded-xl text-lg font-bold active:scale-95 transition" onClick={saveRecord}>保存记录</button></div></div>
    <div className="bg-zinc-900 border border-zinc-800 rounded-2xl p-4 mb-5"><div className="text-lg font-semibold mb-3">导出审批表格</div><input type="date" className="w-full bg-zinc-800 rounded-xl p-3 outline-none mb-3" value={dateFilter} onChange={e => setDateFilter(e.target.value)} /><div className="grid grid-cols-2 gap-3"><button className="bg-blue-600 rounded-xl py-3 font-semibold" onClick={exportDailyApprovalXLSX}>导出当日审批表</button><button className="bg-zinc-700 rounded-xl py-3 font-semibold" onClick={exportAllXLSX}>导出总台账</button></div><div className="text-xs text-zinc-500 mt-3">使用内置Excel模板导出，保留付款审批单和清单格式。</div></div>
    <div className="mb-4 flex items-center justify-between"><div className="text-lg font-semibold">{dateFilter} 现金流水</div><div className="text-sm text-zinc-400">{filteredRecords.length} 笔</div></div>
    <div className="space-y-3 mb-6">{filteredRecords.length === 0 && <div className="text-zinc-500 text-center py-8">当天暂无记录</div>}{filteredRecords.map(item => <div key={item.id} className={`rounded-2xl p-4 border ${item.invoiced ? "bg-green-950 border-green-700" : "bg-red-950 border-red-700"}`}><div className="flex items-center justify-between mb-2"><div className="font-semibold text-lg">{item.handler}</div><div className="font-bold text-lg">{money(item.amount)} PKR</div></div><div className="text-sm mb-1">{item.category}</div><div className="text-sm text-zinc-300 mb-3">{item.note || "无备注"}</div>{item.receiptPhoto && <img src={item.receiptPhoto} alt="票据" className="mb-3 rounded-xl max-h-40 w-full object-cover" />}<div className="flex gap-2 justify-between items-center">{item.invoiced ? <span className="bg-green-600 text-white text-xs px-3 py-1 rounded-full">已收票</span> : <span className="bg-red-600 text-white text-xs px-3 py-1 rounded-full">未收票</span>}<div className="flex gap-2">{!item.invoiced && <button className="bg-yellow-500 text-black text-xs px-3 py-1 rounded-full" onClick={() => markInvoiced(item.id)}>标记收票</button>}<button className="bg-zinc-700 text-white text-xs px-3 py-1 rounded-full" onClick={() => deleteRecord(item.id)}>删除</button></div></div></div>)}</div>
    <div className="grid grid-cols-1 gap-4 pb-10"><div className="bg-zinc-900 border border-zinc-800 rounded-2xl p-4"><div className="text-lg font-semibold mb-3">经办人现金统计</div>{Object.entries(handlerStats).length === 0 && <div className="text-zinc-500 text-sm">暂无数据</div>}{Object.entries(handlerStats).map(([name, value]) => <div key={name} className="flex justify-between py-2 border-b border-zinc-800 last:border-0"><span>{name}</span><span className="font-bold">{money(value)} PKR</span></div>)}</div><div className="bg-zinc-900 border border-zinc-800 rounded-2xl p-4"><div className="text-lg font-semibold mb-3">分类现金统计</div>{Object.entries(categoryStats).length === 0 && <div className="text-zinc-500 text-sm">暂无数据</div>}{Object.entries(categoryStats).map(([name, value]) => <div key={name} className="flex justify-between py-2 border-b border-zinc-800 last:border-0"><span>{name}</span><span className="font-bold">{money(value)} PKR</span></div>)}</div></div>
  </div>;
}
