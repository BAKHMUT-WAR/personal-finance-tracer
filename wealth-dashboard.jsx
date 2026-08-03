import React, { useState, useEffect, useMemo } from 'react';
import {
  Plus, Pencil, Trash2, Copy, ChevronLeft, ChevronRight, BookOpen,
  LayoutDashboard, ArrowLeftRight, Landmark, TrendingUp, Target, ShieldCheck,
  Repeat, Download, Info, Sparkles
} from 'lucide-react';
import {
  ResponsiveContainer, AreaChart, Area, BarChart, Bar,
  PieChart, Pie, Cell, XAxis, YAxis, Tooltip, CartesianGrid
} from 'recharts';
import * as XLSX from 'xlsx';

/* ============================== CONSTANTS ============================== */

const CURRENCY = '₹';

const STORAGE_KEYS = {
  tx: 'wealth-transactions',
  nw: 'wealth-networth',
  inv: 'wealth-investments',
  goals: 'wealth-goals',
  protection: 'wealth-protection',
};

const TRANSACTION_TYPES = ['Income', 'Expense', 'Transfer', 'Investment', 'Loan', 'Tax', 'Asset Purchase'];

const EXPENSE_CATEGORIES = ['Housing', 'Food', 'Groceries', 'Utilities', 'Fuel', 'Transportation', 'Healthcare', 'Education', 'Insurance', 'Entertainment', 'Travel', 'Shopping', 'Charity', 'Subscriptions', 'Taxes', 'Miscellaneous'];
const INCOME_CATEGORIES = ['Salary', 'Business Income', 'Freelance', 'Consulting', 'Dividend', 'Interest', 'Rental Income', 'Bonus', 'Gifts', 'Pension', 'Refunds', 'Other Income'];
const TRANSFER_CATEGORIES = ['Between Own Accounts', 'Other'];
const INVESTMENT_TX_CATEGORIES = ['Mutual Fund', 'Stocks', 'PPF', 'EPF', 'NPS', 'Gold / SGB', 'Fixed Deposit', 'Other'];
const LOAN_TX_CATEGORIES = ['Home Loan', 'Education Loan', 'Vehicle Loan', 'Personal Loan', 'Other'];
const TAX_TX_CATEGORIES = ['Advance Tax', 'TDS', 'Self-Assessment Tax', 'Property Tax', 'Other'];
const ASSETPURCHASE_TX_CATEGORIES = ['Real Estate', 'Vehicle', 'Electronics', 'Other'];

function categoriesForType(type) {
  switch (type) {
    case 'Income': return INCOME_CATEGORIES;
    case 'Expense': return EXPENSE_CATEGORIES;
    case 'Transfer': return TRANSFER_CATEGORIES;
    case 'Investment': return INVESTMENT_TX_CATEGORIES;
    case 'Loan': return LOAN_TX_CATEGORIES;
    case 'Tax': return TAX_TX_CATEGORIES;
    case 'Asset Purchase': return ASSETPURCHASE_TX_CATEGORIES;
    default: return ['Other'];
  }
}

const ASSET_CATEGORIES = ['Savings Account', 'Current Account', 'Cash', 'Fixed Deposit', 'Recurring Deposit', 'Stocks', 'Mutual Funds', 'ETFs', 'Gold', 'Sovereign Gold Bonds', 'PPF', 'EPF', 'NPS', 'Bonds', 'REITs', 'InvITs', 'Real Estate', 'Vehicles', 'Business Ownership', 'Other'];
const LIQUID_ASSET_CATEGORIES = ['Savings Account', 'Current Account', 'Cash'];
const LIABILITY_CATEGORIES = ['Home Loan', 'Education Loan', 'Vehicle Loan', 'Personal Loan', 'Credit Card Debt', 'Business Debt', 'Other'];

const INVESTMENT_CATEGORIES = ['Direct Stocks', 'Mutual Funds', 'ETFs / Index Funds', 'Debt Funds', 'International Funds', 'Gold ETF / SGB', 'PPF', 'EPF', 'NPS', 'Bonds', 'REITs / InvITs'];

const GOAL_TYPES = ['Emergency Fund', 'Education', 'House Purchase', 'Vehicle', 'Marriage', 'Travel', 'Business', 'Retirement', 'Other'];
const INSURANCE_TYPES = ['Life', 'Health', 'Vehicle', 'Property'];

const PALETTE = ['#B6694C', '#51687D', '#8C6A8C', '#A89570', '#C99A3E', '#5B8270', '#A4566B', '#3E7C8C', '#7A8C5B', '#9C6644'];

/* ============================== DATE HELPERS ============================== */

const pad = (n) => String(n).padStart(2, '0');
const toISODate = (d) => `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
const parseISODate = (s) => { const [y, m, d] = s.split('-').map(Number); return new Date(y, m - 1, d); };
const todayISO = () => toISODate(new Date());
const uid = () => `e_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 8)}`;

function startOfDay(d) { const x = new Date(d); x.setHours(0, 0, 0, 0); return x; }
function endOfDay(d) { const x = new Date(d); x.setHours(23, 59, 59, 999); return x; }
function startOfWeek(d) { const x = new Date(d); const day = x.getDay(); const diff = day === 0 ? -6 : 1 - day; x.setDate(x.getDate() + diff); x.setHours(0, 0, 0, 0); return x; }
function endOfWeek(d) { const s = startOfWeek(d); const e = new Date(s); e.setDate(s.getDate() + 6); e.setHours(23, 59, 59, 999); return e; }
function startOfMonth(d) { return new Date(d.getFullYear(), d.getMonth(), 1); }
function endOfMonth(d) { return new Date(d.getFullYear(), d.getMonth() + 1, 0, 23, 59, 59, 999); }
function startOfQuarter(d) { const q = Math.floor(d.getMonth() / 3); return new Date(d.getFullYear(), q * 3, 1); }
function endOfQuarter(d) { const s = startOfQuarter(d); return new Date(s.getFullYear(), s.getMonth() + 3, 0, 23, 59, 59, 999); }
function startOfYear(d) { return new Date(d.getFullYear(), 0, 1); }
function endOfYear(d) { return new Date(d.getFullYear(), 11, 31, 23, 59, 59, 999); }
function startOfFY(d) { const y = d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1; return new Date(y, 3, 1); }
function endOfFY(d) { const s = startOfFY(d); return new Date(s.getFullYear() + 1, 2, 31, 23, 59, 59, 999); }

function getRange(view, anchor) {
  switch (view) {
    case 'daily': return { start: startOfDay(anchor), end: endOfDay(anchor) };
    case 'weekly': return { start: startOfWeek(anchor), end: endOfWeek(anchor) };
    case 'monthly': return { start: startOfMonth(anchor), end: endOfMonth(anchor) };
    case 'quarterly': return { start: startOfQuarter(anchor), end: endOfQuarter(anchor) };
    case 'yearly': return { start: startOfYear(anchor), end: endOfYear(anchor) };
    case 'fy': return { start: startOfFY(anchor), end: endOfFY(anchor) };
    default: return { start: startOfDay(anchor), end: endOfDay(anchor) };
  }
}

function navigateAnchor(view, anchor, dir) {
  const x = new Date(anchor);
  if (view === 'daily') { x.setDate(x.getDate() + dir); return x; }
  if (view === 'weekly') { x.setDate(x.getDate() + 7 * dir); return x; }
  if (view === 'monthly') return new Date(anchor.getFullYear(), anchor.getMonth() + dir, 1);
  if (view === 'quarterly') { x.setMonth(x.getMonth() + 3 * dir); return x; }
  if (view === 'yearly') { x.setFullYear(x.getFullYear() + dir); return x; }
  if (view === 'fy') { x.setFullYear(x.getFullYear() + dir); return x; }
  return x;
}

const monthShort = (d) => d.toLocaleDateString('en-IN', { month: 'short' });
const formatLongDate = (d) => d.toLocaleDateString('en-IN', { weekday: 'short', day: 'numeric', month: 'short', year: 'numeric' });
const formatShortDate = (s) => parseISODate(s).toLocaleDateString('en-IN', { day: 'numeric', month: 'short' });

function fyLabel(d) { const y = d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1; return `FY ${y}-${String((y + 1) % 100).padStart(2, '0')}`; }

function periodLabel(view, anchor) {
  if (view === 'daily') { const isToday = toISODate(anchor) === todayISO(); return (isToday ? 'Today · ' : '') + formatLongDate(anchor); }
  if (view === 'weekly') { const s = startOfWeek(anchor), e = endOfWeek(anchor); return s.getMonth() === e.getMonth() ? `${s.getDate()}–${e.getDate()} ${monthShort(e)} ${e.getFullYear()}` : `${s.getDate()} ${monthShort(s)} – ${e.getDate()} ${monthShort(e)} ${e.getFullYear()}`; }
  if (view === 'monthly') return anchor.toLocaleDateString('en-IN', { month: 'long', year: 'numeric' });
  if (view === 'quarterly') { const q = Math.floor(anchor.getMonth() / 3) + 1; return `Q${q} ${startOfQuarter(anchor).getFullYear()}`; }
  if (view === 'yearly') return String(anchor.getFullYear());
  if (view === 'fy') return fyLabel(anchor);
  return '';
}

/* ============================== FORMATTING ============================== */

const fmt = (n) => `${CURRENCY}${Math.abs(Number(n) || 0).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
const fmtSigned = (n) => `${Number(n) < 0 ? '–' : ''}${fmt(n)}`;
function fmtCompact(n) {
  const v = Number(n) || 0;
  const abs = Math.abs(v);
  const sign = v < 0 ? '–' : '';
  if (abs >= 1e7) return `${sign}${CURRENCY}${(abs / 1e7).toFixed(2)} Cr`;
  if (abs >= 1e5) return `${sign}${CURRENCY}${(abs / 1e5).toFixed(2)} L`;
  return fmtSigned(v);
}

function categoryColor(name) {
  let h = 0;
  for (let i = 0; i < name.length; i++) h = (h * 31 + name.charCodeAt(i)) >>> 0;
  return PALETTE[h % PALETTE.length];
}

/* ============================== FINANCIAL MATH ============================== */

function futureValueSeries(monthlyContribution, annualReturnPct, years, startingCorpus) {
  const r = (annualReturnPct || 0) / 100 / 12;
  const n = Math.max(0, Math.round((years || 0) * 12));
  const fvLump = (startingCorpus || 0) * Math.pow(1 + r, n);
  const fvSeries = r === 0 ? (monthlyContribution || 0) * n : (monthlyContribution || 0) * ((Math.pow(1 + r, n) - 1) / r) * (1 + r);
  return fvLump + fvSeries;
}

function yearsToTarget(corpus, monthlyContribution, annualReturnPct, target) {
  if (corpus >= target) return 0;
  if ((monthlyContribution || 0) <= 0 && (annualReturnPct || 0) <= 0) return null;
  let bal = corpus;
  const r = (annualReturnPct || 0) / 100 / 12;
  for (let year = 1; year <= 80; year++) {
    for (let m = 0; m < 12; m++) bal = bal * (1 + r) + (monthlyContribution || 0);
    if (bal >= target) return year;
  }
  return null;
}

function simpleCAGR(invested, current, purchaseDateISO) {
  if (!invested || invested <= 0 || !purchaseDateISO) return null;
  const days = (Date.now() - parseISODate(purchaseDateISO).getTime()) / 86400000;
  if (days < 30) return null;
  const years = days / 365;
  const ratio = current / invested;
  if (ratio <= 0) return -100;
  return (Math.pow(ratio, 1 / years) - 1) * 100;
}

function retirementProjection({ currentAge, retirementAge, currentCorpus, monthlyInvestment, expectedReturn, inflation }, annualExpenses) {
  const years = Math.max(0, (retirementAge || 0) - (currentAge || 0));
  const projectedCorpus = futureValueSeries(monthlyInvestment, expectedReturn, years, currentCorpus);
  const inflatedAnnualExpenses = (annualExpenses || 0) * Math.pow(1 + (inflation || 0) / 100, years);
  const targetCorpus = inflatedAnnualExpenses * 25;
  const readiness = targetCorpus > 0 ? Math.min(150, (projectedCorpus / targetCorpus) * 100) : 0;
  return { years, projectedCorpus, targetCorpus, readiness, inflatedAnnualExpenses };
}

function fireProjection({ investableNetWorth, annualExpenses, monthlyInvestment, expectedReturn, fiMultiplier }) {
  const mult = fiMultiplier || 25;
  const fiTarget = (annualExpenses || 0) * mult;
  const progress = fiTarget > 0 ? Math.min(150, (investableNetWorth / fiTarget) * 100) : 0;
  const yearsRemaining = investableNetWorth >= fiTarget ? 0 : yearsToTarget(investableNetWorth, monthlyInvestment, expectedReturn, fiTarget);
  let projectedDate = null;
  if (yearsRemaining !== null) { const d = new Date(); d.setFullYear(d.getFullYear() + yearsRemaining); projectedDate = d; }
  return { fiTarget, progress, yearsRemaining, projectedDate };
}

function coastFireNumber({ currentAge, retirementAge, expectedReturn, targetCorpus }) {
  const years = Math.max(0, (retirementAge || 0) - (currentAge || 0));
  return targetCorpus / Math.pow(1 + (expectedReturn || 0) / 100, years);
}

/* ============================== INDIAN TAX (FY 2026-27) ============================== */
/* Figures verified for FY 2026-27: new regime slabs unchanged from FY 2025-26 per Union
   Budget 2026 (no slab changes announced). This is an estimate for planning only — not
   tax advice; confirm specifics with a CA before filing. */

const NEW_REGIME_BOUNDARIES = [0, 400000, 800000, 1200000, 1600000, 2000000, 2400000, Infinity];
const NEW_REGIME_RATES = [0, 0.05, 0.10, 0.15, 0.20, 0.25, 0.30];

const OLD_REGIME_BANDS = {
  general: { boundaries: [0, 250000, 500000, 1000000, Infinity], rates: [0, 0.05, 0.20, 0.30] },
  senior: { boundaries: [0, 300000, 500000, 1000000, Infinity], rates: [0, 0.05, 0.20, 0.30] },
  superSenior: { boundaries: [0, 500000, 1000000, Infinity], rates: [0, 0.20, 0.30] },
};

function slabTax(taxable, boundaries, rates) {
  let tax = 0;
  for (let i = 1; i < boundaries.length; i++) {
    const lower = boundaries[i - 1], upper = boundaries[i];
    if (taxable > lower) tax += (Math.min(taxable, upper) - lower) * rates[i - 1];
    else break;
  }
  return tax;
}

function rebateWithMarginalRelief(taxable, rawTax, threshold, rebateCap) {
  if (taxable <= threshold) return Math.max(0, rawTax - Math.min(rebateCap, rawTax));
  return Math.min(rawTax, Math.max(0, taxable - threshold));
}

function computeSurcharge(taxable, tax, regime) {
  if (tax <= 0) return 0;
  let rate = 0;
  if (taxable > 50000000) rate = regime === 'old' ? 0.37 : 0.25;
  else if (taxable > 20000000) rate = 0.25;
  else if (taxable > 10000000) rate = 0.15;
  else if (taxable > 5000000) rate = 0.10;
  return tax * rate;
}

function computeNewRegimeTax(grossIncome, isSalaried) {
  const stdDeduction = isSalaried ? 75000 : 0;
  const taxable = Math.max(0, grossIncome - stdDeduction);
  const rawTax = slabTax(taxable, NEW_REGIME_BOUNDARIES, NEW_REGIME_RATES);
  const netTax = rebateWithMarginalRelief(taxable, rawTax, 1200000, 60000);
  const surcharge = computeSurcharge(taxable, netTax, 'new');
  const cess = (netTax + surcharge) * 0.04;
  return { taxable, rawTax, netTax, surcharge, cess, total: netTax + surcharge + cess };
}

function computeOldRegimeTax(grossIncome, isSalaried, ageBand, deductions) {
  const stdDeduction = isSalaried ? 50000 : 0;
  const cappedDeductions = Math.min(150000, deductions.section80C || 0) + Math.min(100000, deductions.section80D || 0) + Math.min(50000, deductions.nps80CCD1B || 0) + Math.min(200000, deductions.homeLoanInterest || 0) + Math.max(0, deductions.other || 0);
  const taxable = Math.max(0, grossIncome - stdDeduction - cappedDeductions);
  const band = OLD_REGIME_BANDS[ageBand] || OLD_REGIME_BANDS.general;
  const rawTax = slabTax(taxable, band.boundaries, band.rates);
  const netTax = rebateWithMarginalRelief(taxable, rawTax, 500000, 12500);
  const surcharge = computeSurcharge(taxable, netTax, 'old');
  const cess = (netTax + surcharge) * 0.04;
  return { taxable, rawTax, netTax, surcharge, cess, total: netTax + surcharge + cess, deductionsUsed: cappedDeductions };
}

/* ============================== SEED DATA ============================== */

function seedTransactions() {
  const today = new Date();
  const iso = (offset) => { const d = new Date(today); d.setDate(d.getDate() - offset); return toISODate(d); };
  return [
    { id: uid(), type: 'Income', date: iso(48), amount: 145000, category: 'Salary', account: 'HDFC Savings', note: 'Monthly salary', tags: '', recurring: true, recurrenceFrequency: 'monthly' },
    { id: uid(), type: 'Income', date: iso(18), amount: 145000, category: 'Salary', account: 'HDFC Savings', note: 'Monthly salary', tags: '', recurring: true, recurrenceFrequency: 'monthly' },
    { id: uid(), type: 'Expense', date: iso(45), amount: 28000, category: 'Housing', account: 'HDFC Savings', note: 'Rent', tags: 'fixed', recurring: true, recurrenceFrequency: 'monthly' },
    { id: uid(), type: 'Expense', date: iso(40), amount: 3200, category: 'Utilities', account: 'HDFC Savings', note: 'Electricity and water', tags: '', recurring: false, recurrenceFrequency: '' },
    { id: uid(), type: 'Expense', date: iso(35), amount: 6500, category: 'Groceries', account: 'HDFC Savings', note: 'Monthly groceries', tags: 'essential', recurring: false, recurrenceFrequency: '' },
    { id: uid(), type: 'Investment', date: iso(33), amount: 25000, category: 'Mutual Fund', account: 'HDFC Savings', note: 'SIP - Index Fund', tags: 'sip', recurring: true, recurrenceFrequency: 'monthly' },
    { id: uid(), type: 'Expense', date: iso(28), amount: 1800, category: 'Fuel', account: 'ICICI Credit Card', note: 'Petrol', tags: '', recurring: false, recurrenceFrequency: '' },
    { id: uid(), type: 'Expense', date: iso(22), amount: 2400, category: 'Entertainment', account: 'ICICI Credit Card', note: 'Movies and dining', tags: '', recurring: false, recurrenceFrequency: '' },
    { id: uid(), type: 'Income', date: iso(20), amount: 18000, category: 'Freelance', account: 'HDFC Savings', note: 'Design project', tags: '', recurring: false, recurrenceFrequency: '' },
    { id: uid(), type: 'Expense', date: iso(15), amount: 4500, category: 'Healthcare', account: 'HDFC Savings', note: 'Dental checkup', tags: '', recurring: false, recurrenceFrequency: '' },
    { id: uid(), type: 'Loan', date: iso(13), amount: 32000, category: 'Home Loan', account: 'HDFC Savings', note: 'EMI', tags: 'emi', recurring: true, recurrenceFrequency: 'monthly' },
    { id: uid(), type: 'Expense', date: iso(9), amount: 1200, category: 'Subscriptions', account: 'ICICI Credit Card', note: 'OTT and music', tags: '', recurring: true, recurrenceFrequency: 'monthly' },
    { id: uid(), type: 'Expense', date: iso(6), amount: 3100, category: 'Shopping', account: 'ICICI Credit Card', note: 'Clothes', tags: '', recurring: false, recurrenceFrequency: '' },
    { id: uid(), type: 'Tax', date: iso(4), amount: 15000, category: 'Advance Tax', account: 'HDFC Savings', note: 'Q1 advance tax', tags: '', recurring: false, recurrenceFrequency: '' },
    { id: uid(), type: 'Expense', date: iso(2), amount: 950, category: 'Groceries', account: 'HDFC Savings', note: 'Top-up groceries', tags: '', recurring: false, recurrenceFrequency: '' },
  ];
}

function seedNetWorth() {
  const today = new Date();
  const iso = (offset) => { const d = new Date(today); d.setDate(d.getDate() - offset); return toISODate(d); };
  const assets = [
    { id: uid(), category: 'Savings Account', name: 'HDFC Savings', value: 285000 },
    { id: uid(), category: 'Cash', name: 'Cash in hand', value: 12000 },
    { id: uid(), category: 'Fixed Deposit', name: 'SBI FD', value: 200000 },
    { id: uid(), category: 'Mutual Funds', name: 'Equity & Index SIPs', value: 480000 },
    { id: uid(), category: 'Stocks', name: 'Direct equity', value: 165000 },
    { id: uid(), category: 'Gold', name: 'Physical gold + SGB', value: 95000 },
    { id: uid(), category: 'PPF', name: 'PPF account', value: 310000 },
    { id: uid(), category: 'EPF', name: 'EPF balance', value: 420000 },
    { id: uid(), category: 'Real Estate', name: 'Family plot (share)', value: 1500000 },
  ];
  const liabilities = [
    { id: uid(), category: 'Home Loan', name: 'HDFC Home Loan', amount: 2400000, emi: 32000, interestRate: 8.6, remainingTenureMonths: 168 },
    { id: uid(), category: 'Credit Card Debt', name: 'ICICI Credit Card', amount: 18000, emi: 0, interestRate: 36, remainingTenureMonths: 1 },
  ];
  const totalAssets = assets.reduce((s, a) => s + a.value, 0);
  const totalLiabilities = liabilities.reduce((s, l) => s + l.amount, 0);
  const history = [60, 30, 0].map((off) => ({
    date: iso(off),
    netWorth: (totalAssets - totalLiabilities) - off * 900,
    totalAssets: totalAssets - off * 700,
    totalLiabilities: totalLiabilities - off * (-200),
  }));
  return { assets, liabilities, history };
}

function seedInvestments() {
  const today = new Date();
  const iso = (offset) => { const d = new Date(today); d.setDate(d.getDate() - offset); return toISODate(d); };
  return [
    { id: uid(), category: 'Mutual Funds', name: 'Nifty 50 Index Fund', investedAmount: 280000, currentValue: 342000, purchaseDate: iso(900), isSIP: true, monthlySIP: 15000 },
    { id: uid(), category: 'Mutual Funds', name: 'Flexicap Fund', investedAmount: 120000, currentValue: 138000, purchaseDate: iso(540), isSIP: true, monthlySIP: 10000 },
    { id: uid(), category: 'Direct Stocks', name: 'Direct equity portfolio', investedAmount: 150000, currentValue: 165000, purchaseDate: iso(620), isSIP: false, monthlySIP: 0 },
    { id: uid(), category: 'Gold ETF / SGB', name: 'Sovereign Gold Bonds', investedAmount: 80000, currentValue: 95000, purchaseDate: iso(700), isSIP: false, monthlySIP: 0 },
    { id: uid(), category: 'PPF', name: 'PPF account', investedAmount: 280000, currentValue: 310000, purchaseDate: iso(1500), isSIP: true, monthlySIP: 5000 },
    { id: uid(), category: 'EPF', name: 'EPF balance', investedAmount: 380000, currentValue: 420000, purchaseDate: iso(1800), isSIP: true, monthlySIP: 0 },
  ];
}

function seedGoals() {
  const today = new Date();
  const d1 = new Date(today); d1.setFullYear(d1.getFullYear() + 1);
  const d2 = new Date(today); d2.setFullYear(d2.getFullYear() + 5);
  const d3 = new Date(today); d3.setFullYear(d3.getFullYear() + 2);
  return {
    goals: [
      { id: uid(), name: 'Emergency Fund', type: 'Emergency Fund', targetAmount: 600000, currentAmount: 297000, deadline: toISODate(d1) },
      { id: uid(), name: "Child's Education", type: 'Education', targetAmount: 2500000, currentAmount: 320000, deadline: toISODate(d2) },
      { id: uid(), name: 'New Car', type: 'Vehicle', targetAmount: 900000, currentAmount: 180000, deadline: toISODate(d3) },
    ],
    retirement: { currentAge: 32, retirementAge: 58, currentCorpus: 1300000, monthlyInvestment: 30000, expectedReturn: 11, inflation: 6, annualExpenses: 600000 },
    fireSettings: { fiMultiplier: 25, expectedReturn: 11, monthlyInvestment: 30000 },
  };
}

function seedProtection() {
  const today = new Date();
  const iso = (offset) => { const d = new Date(today); d.setDate(d.getDate() + offset); return toISODate(d); };
  return {
    emergencyFundTargetMonths: 6,
    taxInputs: { grossAnnualIncome: 1500000, isSalaried: true, ageBand: 'general', section80C: 150000, section80D: 25000, nps80CCD1B: 0, homeLoanInterest: 0, otherDeductions: 0, tdsAlreadyPaid: 90000, advanceTaxPaid: 0 },
    insurance: [
      { id: uid(), type: 'Health', provider: 'Star Health', coverage: 1000000, premium: 18000, renewalDate: iso(140) },
      { id: uid(), type: 'Life', provider: 'HDFC Life Term', coverage: 10000000, premium: 14000, renewalDate: iso(260) },
      { id: uid(), type: 'Vehicle', provider: 'ICICI Lombard', coverage: 600000, premium: 7000, renewalDate: iso(40) },
    ],
    creditCards: [
      { id: uid(), name: 'ICICI Amazon Pay', limit: 200000, outstanding: 18000, dueDate: iso(8), billingDay: 18 },
      { id: uid(), name: 'HDFC Regalia', limit: 300000, outstanding: 0, dueDate: iso(22), billingDay: 28 },
    ],
  };
}

/* ============================== WEALTH SCORE ============================== */

function computeWealthScore({ savingsRatePct, emergencyMonths, emergencyTargetMonths, debtToIncomePct, investmentByCategory, insurance, annualIncome, netWorthHistory, fireProgressPct, retirementReadinessPct }) {
  const breakdown = [];

  const savingsScore = Math.max(0, Math.min(20, (savingsRatePct / 30) * 20));
  breakdown.push({ label: 'Savings rate', score: savingsScore, max: 20, hint: 'Aim for 20-30%+ of income saved each month.' });

  const efRatio = emergencyTargetMonths > 0 ? emergencyMonths / emergencyTargetMonths : 0;
  const efScore = Math.max(0, Math.min(15, efRatio * 15));
  breakdown.push({ label: 'Emergency fund', score: efScore, max: 15, hint: `Build toward ${emergencyTargetMonths} months of essential expenses in liquid savings.` });

  const debtScore = Math.max(0, Math.min(15, 15 - (debtToIncomePct / 50) * 15));
  breakdown.push({ label: 'Debt level', score: debtScore, max: 15, hint: 'Keep annual debt servicing well under 50% of income.' });

  const totalInvested = Object.values(investmentByCategory || {}).reduce((s, v) => s + v, 0);
  const numCategories = Object.keys(investmentByCategory || {}).length;
  const maxShare = totalInvested > 0 ? Math.max(...Object.values(investmentByCategory)) / totalInvested : 1;
  let diversificationScore = Math.min(10, numCategories * 2);
  if (maxShare > 0.7) diversificationScore = Math.max(0, diversificationScore - 5);
  breakdown.push({ label: 'Diversification', score: diversificationScore, max: 10, hint: 'Spread investments across more than one category and avoid concentration.' });

  const lifeCover = (insurance || []).filter((i) => i.type === 'Life').reduce((s, i) => s + Number(i.coverage || 0), 0);
  const hasHealth = (insurance || []).some((i) => i.type === 'Health');
  const lifeAdequacy = annualIncome > 0 ? Math.min(1, lifeCover / (annualIncome * 10)) : 0;
  const insuranceScore = (lifeAdequacy * 5) + (hasHealth ? 5 : 0);
  breakdown.push({ label: 'Insurance adequacy', score: insuranceScore, max: 10, hint: 'Aim for life cover near 10x annual income, plus a standalone health policy.' });

  let trendScore = 5;
  if (netWorthHistory && netWorthHistory.length >= 2) {
    const sorted = [...netWorthHistory].sort((a, b) => a.date.localeCompare(b.date));
    const delta = sorted[sorted.length - 1].netWorth - sorted[0].netWorth;
    trendScore = delta > 0 ? 10 : delta === 0 ? 5 : 0;
  }
  breakdown.push({ label: 'Net worth trend', score: trendScore, max: 10, hint: 'A consistently rising net worth over time is the clearest sign of progress.' });

  const fireScore = Math.max(0, Math.min(10, (fireProgressPct / 100) * 10));
  breakdown.push({ label: 'FI progress', score: fireScore, max: 10, hint: 'Grow your investable net worth toward 25x your annual expenses.' });

  const retScore = Math.max(0, Math.min(10, (retirementReadinessPct / 100) * 10));
  breakdown.push({ label: 'Retirement readiness', score: retScore, max: 10, hint: 'Increase your monthly investment or extend your investment horizon.' });

  const total = Math.round(breakdown.reduce((s, b) => s + b.score, 0));
  const tips = [...breakdown].sort((a, b) => (a.score / a.max) - (b.score / b.max)).slice(0, 2).map((b) => b.hint);
  return { total, breakdown, tips };
}

/* ============================== EXPORT HELPERS ============================== */

function downloadCSV(filename, rows) {
  if (!rows.length) return;
  const headers = Object.keys(rows[0]);
  const escape = (v) => `"${String(v ?? '').replace(/"/g, '""')}"`;
  const csv = [headers.join(','), ...rows.map((r) => headers.map((h) => escape(r[h])).join(','))].join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename;
  document.body.appendChild(a); a.click(); document.body.removeChild(a);
  URL.revokeObjectURL(url);
}

function downloadXLSX(filename, sheets) {
  const wb = XLSX.utils.book_new();
  Object.entries(sheets).forEach(([name, rows]) => {
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  XLSX.writeFile(wb, filename);
}

/* ============================== SMALL SHARED UI ============================== */

function NavTabs({ active, onChange, items }) {
  return (
    <div className="wd-navtabs" role="tablist" aria-label="Dashboard section">
      {items.map((it) => (
        <button key={it.key} role="tab" aria-selected={active === it.key} className={`wd-navtab ${active === it.key ? 'active' : ''}`} onClick={() => onChange(it.key)}>
          <it.icon size={15} />
          <span>{it.label}</span>
        </button>
      ))}
    </div>
  );
}

function PeriodSwitcher({ view, setView, anchor, setAnchor }) {
  const range = getRange(view, anchor);
  const todayRange = getRange(view, new Date());
  const isCurrent = toISODate(range.start) === toISODate(todayRange.start);
  return (
    <>
      <div className="wd-tabs" role="tablist" aria-label="Time period">
        {['daily', 'weekly', 'monthly', 'quarterly', 'yearly', 'fy'].map((v) => (
          <button key={v} role="tab" aria-selected={view === v} className={`wd-tab ${view === v ? 'active' : ''}`} onClick={() => setView(v)}>
            {v === 'fy' ? 'FY' : v[0].toUpperCase() + v.slice(1)}
          </button>
        ))}
      </div>
      <div className="wd-period-nav">
        <button className="wd-nav-btn" aria-label="Previous period" onClick={() => setAnchor((a) => navigateAnchor(view, a, -1))}><ChevronLeft size={17} /></button>
        <div className="wd-period-center">
          <span className="wd-period-label">{periodLabel(view, anchor)}</span>
          {!isCurrent && <button className="wd-jump-today" onClick={() => setAnchor(new Date())}>Jump to today</button>}
        </div>
        <button className="wd-nav-btn" aria-label="Next period" onClick={() => setAnchor((a) => navigateAnchor(view, a, 1))}><ChevronRight size={17} /></button>
      </div>
    </>
  );
}

function ConfirmInline({ label, onConfirm, onCancel }) {
  return (
    <div className="wd-confirm">
      <span className="wd-confirm-label">{label}</span>
      <button className="yes" onClick={onConfirm}>Confirm</button>
      <button onClick={onCancel}>Cancel</button>
    </div>
  );
}

/* ============================== MAIN COMPONENT ============================== */

export default function WealthDashboard() {
  const [transactions, setTransactions] = useState([]);
  const [netWorth, setNetWorth] = useState({ assets: [], liabilities: [], history: [] });
  const [investments, setInvestments] = useState([]);
  const [goalsFire, setGoalsFire] = useState({ goals: [], retirement: { currentAge: 30, retirementAge: 58, currentCorpus: 0, monthlyInvestment: 0, expectedReturn: 11, inflation: 6, annualExpenses: 0 }, fireSettings: { fiMultiplier: 25, expectedReturn: 11, monthlyInvestment: 0 } });
  const [protection, setProtection] = useState({ emergencyFundTargetMonths: 6, taxInputs: { grossAnnualIncome: 0, isSalaried: true, ageBand: 'general', section80C: 0, section80D: 0, nps80CCD1B: 0, homeLoanInterest: 0, otherDeductions: 0, tdsAlreadyPaid: 0, advanceTaxPaid: 0 }, insurance: [], creditCards: [] });

  const [loaded, setLoaded] = useState(false);
  const [storageNotice, setStorageNotice] = useState('');
  const [activeModule, setActiveModule] = useState('overview');
  const [view, setView] = useState('monthly');
  const [anchor, setAnchor] = useState(new Date());
  const [confirmReset, setConfirmReset] = useState(false);

  useEffect(() => {
    let alive = true;
    (async () => {
      const safeGet = async (key, seedFn) => {
        try {
          const res = await window.storage.get(key);
          return res && res.value ? JSON.parse(res.value) : seedFn();
        } catch (err) {
          return seedFn();
        }
      };
      const [tx, nw, inv, gf, pr] = await Promise.all([
        safeGet(STORAGE_KEYS.tx, seedTransactions),
        safeGet(STORAGE_KEYS.nw, seedNetWorth),
        safeGet(STORAGE_KEYS.inv, seedInvestments),
        safeGet(STORAGE_KEYS.goals, seedGoals),
        safeGet(STORAGE_KEYS.protection, seedProtection),
      ]);
      if (!alive) return;
      setTransactions(tx); setNetWorth(nw); setInvestments(inv); setGoalsFire(gf); setProtection(pr);
      setLoaded(true);
    })();
    return () => { alive = false; };
  }, []);

  function useSave(key, value) {
    useEffect(() => {
      if (!loaded) return;
      (async () => {
        try { await window.storage.set(key, JSON.stringify(value)); setStorageNotice(''); }
        catch (err) { setStorageNotice("Changes aren't saving right now — edits may be lost on refresh."); }
      })();
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, [value, loaded]);
  }
  useSave(STORAGE_KEYS.tx, transactions);
  useSave(STORAGE_KEYS.nw, netWorth);
  useSave(STORAGE_KEYS.inv, investments);
  useSave(STORAGE_KEYS.goals, goalsFire);
  useSave(STORAGE_KEYS.protection, protection);

  /* ---------- computed: cash flow for selected period ---------- */
  const range = useMemo(() => getRange(view, anchor), [view, anchor]);
  const periodTransactions = useMemo(() => transactions
    .filter((t) => { const d = parseISODate(t.date); return d >= range.start && d <= range.end; })
    .sort((a, b) => (b.date.localeCompare(a.date) || b.id.localeCompare(a.id))), [transactions, range]);

  const periodTotals = useMemo(() => {
    let income = 0, spending = 0, capital = 0, transfer = 0;
    periodTransactions.forEach((t) => {
      const amt = Number(t.amount) || 0;
      if (t.type === 'Income') income += amt;
      else if (t.type === 'Expense' || t.type === 'Tax') spending += amt;
      else if (t.type === 'Investment' || t.type === 'Loan' || t.type === 'Asset Purchase') capital += amt;
      else if (t.type === 'Transfer') transfer += amt;
    });
    return { income, spending, capital, transfer, net: income - spending };
  }, [periodTransactions]);

  const categoryTotals = useMemo(() => {
    const map = {};
    periodTransactions.forEach((t) => {
      if (t.type !== 'Expense' && t.type !== 'Tax') return;
      const key = t.category || 'Other';
      map[key] = (map[key] || 0) + (Number(t.amount) || 0);
    });
    return Object.entries(map).map(([name, total]) => ({ name, total })).sort((a, b) => b.total - a.total);
  }, [periodTransactions]);

  const capitalByType = useMemo(() => {
    const map = {};
    periodTransactions.forEach((t) => {
      if (t.type !== 'Investment' && t.type !== 'Loan' && t.type !== 'Asset Purchase') return;
      map[t.type] = (map[t.type] || 0) + (Number(t.amount) || 0);
    });
    return Object.entries(map).map(([name, total]) => ({ name, total }));
  }, [periodTransactions]);

  const monthRange = useMemo(() => getRange('monthly', new Date()), []);
  const monthTotals = useMemo(() => {
    let income = 0, spending = 0;
    transactions.forEach((t) => {
      const d = parseISODate(t.date);
      if (d < monthRange.start || d > monthRange.end) return;
      const amt = Number(t.amount) || 0;
      if (t.type === 'Income') income += amt;
      else if (t.type === 'Expense' || t.type === 'Tax') spending += amt;
    });
    return { income, spending, net: income - spending, savingsRate: income > 0 ? ((income - spending) / income) * 100 : 0 };
  }, [transactions, monthRange]);

  const trendSeries = useMemo(() => {
    const points = [];
    const count = 6;
    for (let i = count - 1; i >= 0; i--) {
      const a = navigateAnchor(view, anchor, -i);
      const r = getRange(view, a);
      let income = 0, spending = 0;
      transactions.forEach((t) => {
        const d = parseISODate(t.date);
        if (d < r.start || d > r.end) return;
        const amt = Number(t.amount) || 0;
        if (t.type === 'Income') income += amt;
        else if (t.type === 'Expense' || t.type === 'Tax') spending += amt;
      });
      points.push({ label: periodLabel(view, a).split(' · ').pop(), income, spending, net: income - spending });
    }
    return points;
  }, [transactions, view, anchor]);

  /* ---------- computed: net worth ---------- */
  const netWorthTotals = useMemo(() => {
    const totalAssets = netWorth.assets.reduce((s, a) => s + (Number(a.value) || 0), 0);
    const totalLiabilities = netWorth.liabilities.reduce((s, l) => s + (Number(l.amount) || 0), 0);
    const liquidCash = netWorth.assets.filter((a) => LIQUID_ASSET_CATEGORIES.includes(a.category)).reduce((s, a) => s + (Number(a.value) || 0), 0);
    const excludedAssetCats = ['Real Estate', 'Vehicles', 'Business Ownership'];
    const excludedLiabCats = ['Home Loan', 'Vehicle Loan'];
    const investableAssets = netWorth.assets.filter((a) => !excludedAssetCats.includes(a.category)).reduce((s, a) => s + (Number(a.value) || 0), 0);
    const investableLiabilities = netWorth.liabilities.filter((l) => !excludedLiabCats.includes(l.category)).reduce((s, l) => s + (Number(l.amount) || 0), 0);
    return { totalAssets, totalLiabilities, netWorthValue: totalAssets - totalLiabilities, liquidCash, investableNetWorth: investableAssets - investableLiabilities };
  }, [netWorth]);

  const assetAllocation = useMemo(() => {
    const map = {};
    netWorth.assets.forEach((a) => { map[a.category] = (map[a.category] || 0) + (Number(a.value) || 0); });
    return Object.entries(map).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
  }, [netWorth.assets]);

  /* ---------- computed: investments ---------- */
  const investmentTotals = useMemo(() => {
    const totalInvested = investments.reduce((s, i) => s + (Number(i.investedAmount) || 0), 0);
    const totalCurrentValue = investments.reduce((s, i) => s + (Number(i.currentValue) || 0), 0);
    const totalGain = totalCurrentValue - totalInvested;
    const gainPct = totalInvested > 0 ? (totalGain / totalInvested) * 100 : 0;
    const monthlySIPTotal = investments.filter((i) => i.isSIP).reduce((s, i) => s + (Number(i.monthlySIP) || 0), 0);
    const byCategory = {};
    investments.forEach((i) => { byCategory[i.category] = (byCategory[i.category] || 0) + (Number(i.currentValue) || 0); });
    return { totalInvested, totalCurrentValue, totalGain, gainPct, monthlySIPTotal, byCategory };
  }, [investments]);

  /* ---------- computed: goals / fire / retirement ---------- */
  const retirement = useMemo(() => retirementProjection(goalsFire.retirement, goalsFire.retirement.annualExpenses), [goalsFire.retirement]);
  const fire = useMemo(() => fireProjection({
    investableNetWorth: netWorthTotals.investableNetWorth,
    annualExpenses: goalsFire.retirement.annualExpenses,
    monthlyInvestment: goalsFire.fireSettings.monthlyInvestment,
    expectedReturn: goalsFire.fireSettings.expectedReturn,
    fiMultiplier: goalsFire.fireSettings.fiMultiplier,
  }), [netWorthTotals.investableNetWorth, goalsFire.retirement.annualExpenses, goalsFire.fireSettings]);

  /* ---------- computed: protection / safety ---------- */
  const essentialMonthlyExpense = useMemo(() => {
    if (goalsFire.retirement.annualExpenses > 0) return goalsFire.retirement.annualExpenses / 12;
    return monthTotals.spending || 1;
  }, [goalsFire.retirement.annualExpenses, monthTotals.spending]);
  const emergencyMonths = essentialMonthlyExpense > 0 ? netWorthTotals.liquidCash / essentialMonthlyExpense : 0;
  const annualDebtServicing = useMemo(() => netWorth.liabilities.reduce((s, l) => s + (Number(l.emi) || 0) * 12, 0), [netWorth.liabilities]);
  const annualIncomeEstimate = monthTotals.income > 0 ? monthTotals.income * 12 : transactions.filter((t) => t.type === 'Income').reduce((s, t) => s + Number(t.amount || 0), 0);
  const debtToIncomePct = annualIncomeEstimate > 0 ? (annualDebtServicing / annualIncomeEstimate) * 100 : 0;
  const debtToAssetPct = netWorthTotals.totalAssets > 0 ? (netWorthTotals.totalLiabilities / netWorthTotals.totalAssets) * 100 : 0;

  const wealthScore = useMemo(() => computeWealthScore({
    savingsRatePct: monthTotals.savingsRate,
    emergencyMonths,
    emergencyTargetMonths: protection.emergencyFundTargetMonths,
    debtToIncomePct,
    investmentByCategory: investmentTotals.byCategory,
    insurance: protection.insurance,
    annualIncome: annualIncomeEstimate,
    netWorthHistory: netWorth.history,
    fireProgressPct: fire.progress,
    retirementReadinessPct: retirement.readiness,
  }), [monthTotals.savingsRate, emergencyMonths, protection.emergencyFundTargetMonths, debtToIncomePct, investmentTotals.byCategory, protection.insurance, annualIncomeEstimate, netWorth.history, fire.progress, retirement.readiness]);

  /* ---------- net worth history snapshot ---------- */
  useEffect(() => {
    if (!loaded) return;
    const todayStr = todayISO();
    setNetWorth((prev) => {
      const filtered = prev.history.filter((h) => h.date !== todayStr);
      const snap = { date: todayStr, netWorth: netWorthTotals.netWorthValue, totalAssets: netWorthTotals.totalAssets, totalLiabilities: netWorthTotals.totalLiabilities };
      const prevLast = filtered[filtered.length - 1];
      if (prevLast && prevLast.netWorth === snap.netWorth && prevLast.totalAssets === snap.totalAssets) return prev;
      return { ...prev, history: [...filtered, snap].sort((a, b) => a.date.localeCompare(b.date)) };
    });
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [loaded, netWorthTotals.netWorthValue, netWorthTotals.totalAssets, netWorthTotals.totalLiabilities]);

  /* ---------- CRUD: transactions ---------- */
  const addTransaction = (payload) => setTransactions((prev) => [...prev, { id: uid(), ...payload }]);
  const updateTransaction = (id, payload) => setTransactions((prev) => prev.map((t) => (t.id === id ? { ...t, ...payload } : t)));
  const deleteTransaction = (id) => setTransactions((prev) => prev.filter((t) => t.id !== id));
  const duplicateTransaction = (t) => setTransactions((prev) => [...prev, { ...t, id: uid(), date: todayISO() }]);

  const moduleItems = [
    { key: 'overview', label: 'Overview', icon: LayoutDashboard },
    { key: 'transactions', label: 'Cash Flow', icon: ArrowLeftRight },
    { key: 'networth', label: 'Net Worth', icon: Landmark },
    { key: 'investments', label: 'Investments', icon: TrendingUp },
    { key: 'goals', label: 'Goals & FIRE', icon: Target },
    { key: 'tax', label: 'Tax & Safety', icon: ShieldCheck },
  ];

  if (!loaded) {
    return (
      <div className="wd-wrap">
        <style>{CSS}</style>
        <div className="wd-loading">Loading your wealth dashboard…</div>
      </div>
    );
  }

  return (
    <div className="wd-wrap">
      <style>{CSS}</style>
      <div className="wd-container">

        <div className="wd-header">
          <div className="wd-brand"><BookOpen size={14} /> Wealth OS</div>
          <span className="wd-balance-label">Net worth</span>
          <span className={`wd-balance-amount wd-mono ${netWorthTotals.netWorthValue < 0 ? 'negative' : ''}`}>{fmtCompact(netWorthTotals.netWorthValue)}</span>
          <div className="wd-balance-sub wd-mono">As of today · {netWorth.assets.length} assets · {netWorth.liabilities.length} liabilities</div>

          <div className="wd-metric-grid">
            <div className="wd-metric"><span className="m-label">Total assets</span><span className="m-value wd-mono">{fmtCompact(netWorthTotals.totalAssets)}</span></div>
            <div className="wd-metric"><span className="m-label">Total liabilities</span><span className="m-value wd-mono clay">{fmtCompact(netWorthTotals.totalLiabilities)}</span></div>
            <div className="wd-metric"><span className="m-label">Available cash</span><span className="m-value wd-mono">{fmtCompact(netWorthTotals.liquidCash)}</span></div>
            <div className="wd-metric"><span className="m-label">Monthly cash flow</span><span className={`m-value wd-mono ${monthTotals.net < 0 ? 'clay' : 'sage'}`}>{fmtCompact(monthTotals.net)}</span></div>
            <div className="wd-metric"><span className="m-label">Savings rate</span><span className="m-value wd-mono">{monthTotals.savingsRate.toFixed(1)}%</span></div>
            <div className="wd-metric"><span className="m-label">Portfolio value</span><span className="m-value wd-mono">{fmtCompact(investmentTotals.totalCurrentValue)}</span></div>
            <div className="wd-metric"><span className="m-label">Emergency fund</span><span className="m-value wd-mono">{emergencyMonths.toFixed(1)} mo</span></div>
            <div className="wd-metric"><span className="m-label">FI progress</span><span className="m-value wd-mono">{fire.progress.toFixed(0)}%</span></div>
            <div className="wd-metric"><span className="m-label">Retirement ready</span><span className="m-value wd-mono">{retirement.readiness.toFixed(0)}%</span></div>
            <div className="wd-metric"><span className="m-label">Wealth score</span><span className="m-value wd-mono slate">{wealthScore.total}/100</span></div>
          </div>
        </div>

        {storageNotice && <div className="wd-notice">{storageNotice}</div>}

        <NavTabs active={activeModule} onChange={setActiveModule} items={moduleItems} />

        {activeModule === 'overview' && (
          <OverviewModule
            view={view} setView={setView} anchor={anchor} setAnchor={setAnchor}
            periodTotals={periodTotals} categoryTotals={categoryTotals} capitalByType={capitalByType}
            trendSeries={trendSeries} wealthScore={wealthScore} assetAllocation={assetAllocation}
            netWorthHistory={netWorth.history}
          />
        )}

        {activeModule === 'transactions' && (
          <TransactionsModule
            view={view} setView={setView} anchor={anchor} setAnchor={setAnchor}
            periodTransactions={periodTransactions} periodTotals={periodTotals} categoryTotals={categoryTotals}
            capitalByType={capitalByType} allTransactions={transactions}
            onAdd={addTransaction} onUpdate={updateTransaction} onDelete={deleteTransaction} onDuplicate={duplicateTransaction}
          />
        )}

        {activeModule === 'networth' && (
          <NetWorthModule netWorth={netWorth} setNetWorth={setNetWorth} totals={netWorthTotals} allocation={assetAllocation} />
        )}

        {activeModule === 'investments' && (
          <InvestmentsModule investments={investments} setInvestments={setInvestments} totals={investmentTotals} />
        )}

        {activeModule === 'goals' && (
          <GoalsFireModule goalsFire={goalsFire} setGoalsFire={setGoalsFire} retirement={retirement} fire={fire} investableNetWorth={netWorthTotals.investableNetWorth} />
        )}

        {activeModule === 'tax' && (
          <TaxSafetyModule
            protection={protection} setProtection={setProtection}
            netWorth={netWorth} netWorthTotals={netWorthTotals}
            emergencyMonths={emergencyMonths} debtToIncomePct={debtToIncomePct} debtToAssetPct={debtToAssetPct}
            annualIncomeEstimate={annualIncomeEstimate}
          />
        )}

        <div className="wd-footer">
          <div className="wd-footer-links">
            <button className="wd-link-btn" onClick={() => downloadXLSX('wealth-export.xlsx', {
              Transactions: transactions.map((t) => ({ Date: t.date, Type: t.type, Category: t.category, Account: t.account, Amount: t.amount, Note: t.note, Tags: t.tags })),
              Assets: netWorth.assets.map((a) => ({ Category: a.category, Name: a.name, Value: a.value })),
              Liabilities: netWorth.liabilities.map((l) => ({ Category: l.category, Name: l.name, Amount: l.amount, EMI: l.emi, InterestRate: l.interestRate })),
              Investments: investments.map((i) => ({ Category: i.category, Name: i.name, Invested: i.investedAmount, CurrentValue: i.currentValue, PurchaseDate: i.purchaseDate })),
            })}><Download size={13} /> Export everything (.xlsx)</button>
            {confirmReset ? (
              <ConfirmInline label="Clear all data across every module?" onConfirm={() => {
                setTransactions([]); setNetWorth({ assets: [], liabilities: [], history: [] }); setInvestments([]);
                setGoalsFire({ goals: [], retirement: { currentAge: 30, retirementAge: 58, currentCorpus: 0, monthlyInvestment: 0, expectedReturn: 11, inflation: 6, annualExpenses: 0 }, fireSettings: { fiMultiplier: 25, expectedReturn: 11, monthlyInvestment: 0 } });
                setProtection({ emergencyFundTargetMonths: 6, taxInputs: { grossAnnualIncome: 0, isSalaried: true, ageBand: 'general', section80C: 0, section80D: 0, nps80CCD1B: 0, homeLoanInterest: 0, otherDeductions: 0, tdsAlreadyPaid: 0, advanceTaxPaid: 0 }, insurance: [], creditCards: [] });
                setConfirmReset(false);
              }} onCancel={() => setConfirmReset(false)} />
            ) : (
              <button className="wd-reset-link" onClick={() => setConfirmReset(true)}>Clear all data</button>
            )}
          </div>
          <p className="wd-disclaimer">Figures are estimates based on the data you enter — this isn't financial, tax, or investment advice. Verify tax and investment decisions with a qualified professional.</p>
        </div>

      </div>
    </div>
  );
}

/* ============================== CHART HELPERS ============================== */

function CurrencyTooltip({ active, payload, label }) {
  if (!active || !payload || !payload.length) return null;
  return (
    <div className="wd-chart-tooltip">
      <div className="tt-label">{label}</div>
      {payload.map((p, i) => (
        <div key={i} className="tt-row"><span className="tt-dot" style={{ background: p.color || p.fill }} />{p.name}: <strong className="wd-mono">{fmt(p.value)}</strong></div>
      ))}
    </div>
  );
}

function capAllocation(list, max = 7) {
  if (list.length <= max) return list;
  const top = list.slice(0, max - 1);
  const rest = list.slice(max - 1).reduce((s, x) => s + x.value, 0);
  return [...top, { name: 'Other', value: rest }];
}

function CategoryBars({ data, emptyText }) {
  if (!data.length) return <div className="wd-empty">{emptyText || 'Nothing to show for this period.'}</div>;
  const max = data[0].total;
  return (
    <>
      {data.map((c) => (
        <div className="wd-bar-row" key={c.name}>
          <span className="wd-bar-label" title={c.name}>{c.name}</span>
          <div className="wd-bar-track"><div className="wd-bar-fill" style={{ width: `${max ? (c.total / max) * 100 : 0}%`, background: categoryColor(c.name) }} /></div>
          <span className="wd-bar-amount wd-mono">{fmt(c.total)}</span>
        </div>
      ))}
    </>
  );
}

function WealthScorePanel({ wealthScore }) {
  return (
    <div className="wd-card">
      <div className="wd-score-head">
        <div>
          <h3 className="wd-section-title">Wealth score</h3>
          <p className="wd-muted-text">A snapshot across savings, safety net, debt, diversification, and progress toward your goals.</p>
        </div>
        <div className="wd-score-badge"><span className="wd-mono">{wealthScore.total}</span><span className="wd-score-max">/100</span></div>
      </div>
      {wealthScore.breakdown.map((b) => (
        <div className="wd-score-row" key={b.label}>
          <span className="wd-score-label">{b.label}</span>
          <div className="wd-score-track"><div className="wd-score-fill" style={{ width: `${(b.score / b.max) * 100}%` }} /></div>
          <span className="wd-score-value wd-mono">{Math.round(b.score)}/{b.max}</span>
        </div>
      ))}
      <div className="wd-score-tips">
        {wealthScore.tips.map((t, i) => (<div className="wd-tip" key={i}><Sparkles size={13} /> {t}</div>))}
      </div>
    </div>
  );
}

/* ============================== OVERVIEW MODULE ============================== */

function OverviewModule({ view, setView, anchor, setAnchor, periodTotals, categoryTotals, capitalByType, trendSeries, wealthScore, assetAllocation, netWorthHistory }) {
  const allocationData = useMemo(() => capAllocation(assetAllocation), [assetAllocation]);
  const hasHistory = netWorthHistory && netWorthHistory.length >= 2;

  return (
    <>
      <PeriodSwitcher view={view} setView={setView} anchor={anchor} setAnchor={setAnchor} />

      <div className="wd-summary-grid">
        <div className="wd-summary-card in"><span className="sc-label">Money in</span><span className="sc-value wd-mono">{fmt(periodTotals.income)}</span></div>
        <div className="wd-summary-card out"><span className="sc-label">Money out</span><span className="sc-value wd-mono">{fmt(periodTotals.spending)}</span></div>
        <div className="wd-summary-card net"><span className="sc-label">Net</span><span className="sc-value wd-mono">{fmtSigned(periodTotals.net)}</span></div>
      </div>

      {capitalByType.length > 0 && (
        <div className="wd-capital-note">
          <Info size={13} />
          <span>Also moved this period — kept separate from spending: {capitalByType.map((c) => `${c.name} ${fmt(c.total)}`).join(' · ')}</span>
        </div>
      )}

      <div className="wd-card">
        <h3 className="wd-section-title">Income vs. spending — last 6 {view === 'fy' ? 'years' : view}</h3>
        <div style={{ width: '100%', height: 230 }}>
          <ResponsiveContainer>
            <BarChart data={trendSeries} margin={{ top: 6, right: 4, left: -16, bottom: 0 }}>
              <CartesianGrid stroke="var(--line)" vertical={false} />
              <XAxis dataKey="label" tick={{ fontSize: 11, fill: 'var(--ink-soft)' }} axisLine={{ stroke: 'var(--line)' }} tickLine={false} />
              <YAxis tick={{ fontSize: 11, fill: 'var(--ink-soft)' }} axisLine={false} tickLine={false} width={48} tickFormatter={(v) => fmtCompact(v)} />
              <Tooltip content={<CurrencyTooltip />} />
              <Bar dataKey="income" name="Income" fill="var(--sage)" radius={[4, 4, 0, 0]} />
              <Bar dataKey="spending" name="Spending" fill="var(--clay)" radius={[4, 4, 0, 0]} />
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>

      <div className="wd-card">
        <h3 className="wd-section-title">Spending by category</h3>
        <CategoryBars data={categoryTotals} emptyText="No expenses logged for this period." />
      </div>

      <div className="wd-grid-2">
        <div className="wd-card">
          <h3 className="wd-section-title">Asset allocation</h3>
          {allocationData.length === 0 ? <div className="wd-empty">Add assets in Net Worth to see your allocation.</div> : (
            <div style={{ width: '100%', height: 220 }}>
              <ResponsiveContainer>
                <PieChart>
                  <Pie data={allocationData} dataKey="value" nameKey="name" innerRadius="55%" outerRadius="85%" paddingAngle={2}>
                    {allocationData.map((entry, i) => <Cell key={i} fill={categoryColor(entry.name)} stroke="var(--card)" strokeWidth={2} />)}
                  </Pie>
                  <Tooltip content={<CurrencyTooltip />} />
                </PieChart>
              </ResponsiveContainer>
            </div>
          )}
          <div className="wd-legend">
            {allocationData.map((a) => (
              <div className="wd-legend-item" key={a.name}><span className="dot" style={{ background: categoryColor(a.name) }} />{a.name}</div>
            ))}
          </div>
        </div>

        <div className="wd-card">
          <h3 className="wd-section-title">Net worth growth</h3>
          {!hasHistory ? <div className="wd-empty">Keep updating your Net Worth tab — a growth line builds up as snapshots accumulate over time.</div> : (
            <div style={{ width: '100%', height: 220 }}>
              <ResponsiveContainer>
                <AreaChart data={netWorthHistory} margin={{ top: 6, right: 4, left: -16, bottom: 0 }}>
                  <defs>
                    <linearGradient id="nwGradient" x1="0" y1="0" x2="0" y2="1">
                      <stop offset="0%" stopColor="var(--slate)" stopOpacity={0.35} />
                      <stop offset="100%" stopColor="var(--slate)" stopOpacity={0.02} />
                    </linearGradient>
                  </defs>
                  <CartesianGrid stroke="var(--line)" vertical={false} />
                  <XAxis dataKey="date" tickFormatter={(d) => formatShortDate(d)} tick={{ fontSize: 11, fill: 'var(--ink-soft)' }} axisLine={{ stroke: 'var(--line)' }} tickLine={false} />
                  <YAxis tick={{ fontSize: 11, fill: 'var(--ink-soft)' }} axisLine={false} tickLine={false} width={48} tickFormatter={(v) => fmtCompact(v)} />
                  <Tooltip content={<CurrencyTooltip />} labelFormatter={(d) => formatShortDate(d)} />
                  <Area type="monotone" dataKey="netWorth" name="Net worth" stroke="var(--slate)" fill="url(#nwGradient)" strokeWidth={2} />
                </AreaChart>
              </ResponsiveContainer>
            </div>
          )}
        </div>
      </div>

      <WealthScorePanel wealthScore={wealthScore} />
    </>
  );
}

/* ============================== TRANSACTIONS MODULE ============================== */

function txColor(type) {
  if (type === 'Income') return 'var(--sage)';
  if (type === 'Expense' || type === 'Tax') return 'var(--clay)';
  if (type === 'Transfer') return 'var(--slate)';
  return 'var(--amber)';
}
function amountPrefix(type) {
  if (type === 'Income') return '+';
  if (type === 'Expense' || type === 'Tax') return '–';
  return '→';
}
function emptyTxForm(type, date) {
  return { type, date, amount: '', category: '', account: '', note: '', tags: '', recurring: false, recurrenceFrequency: 'monthly' };
}

function TransactionsModule({ view, setView, anchor, setAnchor, periodTransactions, periodTotals, categoryTotals, capitalByType, allTransactions, onAdd, onUpdate, onDelete, onDuplicate }) {
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [formData, setFormData] = useState(emptyTxForm('Expense', todayISO()));
  const [formError, setFormError] = useState('');
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);
  const [search, setSearch] = useState('');
  const [typeFilter, setTypeFilter] = useState('all');

  const catOptions = useMemo(() => {
    const base = categoriesForType(formData.type);
    const used = allTransactions.filter((t) => t.type === formData.type).map((t) => t.category);
    return Array.from(new Set([...base, ...used]));
  }, [formData.type, allTransactions]);

  const accountOptions = useMemo(() => Array.from(new Set(allTransactions.map((t) => t.account).filter(Boolean))), [allTransactions]);

  const filtered = useMemo(() => {
    return periodTransactions.filter((t) => {
      if (typeFilter !== 'all' && t.type !== typeFilter) return false;
      if (!search.trim()) return true;
      const q = search.trim().toLowerCase();
      return [t.category, t.note, t.tags, t.account].some((f) => (f || '').toLowerCase().includes(q));
    });
  }, [periodTransactions, search, typeFilter]);

  function openAdd() {
    setEditingId(null); setFormError('');
    setFormData(emptyTxForm('Expense', view === 'daily' ? toISODate(anchor) : todayISO()));
    setShowForm(true);
  }
  function openEdit(t) {
    setEditingId(t.id); setFormError('');
    setFormData({ type: t.type, date: t.date, amount: String(t.amount), category: t.category, account: t.account || '', note: t.note || '', tags: t.tags || '', recurring: !!t.recurring, recurrenceFrequency: t.recurrenceFrequency || 'monthly' });
    setShowForm(true);
  }
  function closeForm() { setShowForm(false); setEditingId(null); setFormError(''); }
  function setTxType(type) { setFormData((f) => ({ ...f, type, category: '' })); }

  function submit(e) {
    e.preventDefault();
    if (!formData.date) { setFormError('Pick a date.'); return; }
    if (!formData.category.trim()) { setFormError('Add a category.'); return; }
    const amt = parseFloat(formData.amount);
    if (isNaN(amt) || amt <= 0) { setFormError('Enter an amount greater than 0.'); return; }
    const payload = { type: formData.type, date: formData.date, amount: amt, category: formData.category.trim(), account: formData.account.trim(), note: formData.note.trim(), tags: formData.tags.trim(), recurring: formData.recurring, recurrenceFrequency: formData.recurring ? formData.recurrenceFrequency : '' };
    if (editingId) onUpdate(editingId, payload); else onAdd(payload);
    closeForm();
  }

  return (
    <>
      <PeriodSwitcher view={view} setView={setView} anchor={anchor} setAnchor={setAnchor} />

      <div className="wd-summary-grid">
        <div className="wd-summary-card in"><span className="sc-label">Money in</span><span className="sc-value wd-mono">{fmt(periodTotals.income)}</span></div>
        <div className="wd-summary-card out"><span className="sc-label">Money out</span><span className="sc-value wd-mono">{fmt(periodTotals.spending)}</span></div>
        <div className="wd-summary-card net"><span className="sc-label">Net</span><span className="sc-value wd-mono">{fmtSigned(periodTotals.net)}</span></div>
      </div>

      {capitalByType.length > 0 && (
        <div className="wd-capital-note">
          <Info size={13} />
          <span>Excluded from spending above: {capitalByType.map((c) => `${c.name} ${fmt(c.total)}`).join(' · ')}</span>
        </div>
      )}

      <div className="wd-card">
        <h3 className="wd-section-title">Spending by category</h3>
        <CategoryBars data={categoryTotals} emptyText="No expenses logged for this period." />
      </div>

      <div className="wd-card">
        <div className="wd-section-header">
          <h3 className="wd-section-title">Transactions</h3>
          <div className="wd-header-actions">
            <button className="wd-btn-ghost" onClick={() => downloadCSV(`transactions-${view}-${toISODate(anchor)}.csv`, filtered.map((t) => ({ Date: t.date, Type: t.type, Category: t.category, Account: t.account, Amount: t.amount, Note: t.note, Tags: t.tags })))}>
              <Download size={13} /> CSV
            </button>
            <button className="wd-btn-primary" onClick={() => (showForm && !editingId ? closeForm() : openAdd())}><Plus size={15} /> Add transaction</button>
          </div>
        </div>

        <div className="wd-filter-row">
          <input className="wd-search" type="text" placeholder="Search category, note, tag, or account…" value={search} onChange={(e) => setSearch(e.target.value)} />
          <select className="wd-type-filter" value={typeFilter} onChange={(e) => setTypeFilter(e.target.value)}>
            <option value="all">All types</option>
            {TRANSACTION_TYPES.map((t) => <option value={t} key={t}>{t}</option>)}
          </select>
        </div>

        {showForm && (
          <form className="wd-form" onSubmit={submit}>
            <div className="wd-type-toggle wd-type-toggle-wrap">
              {TRANSACTION_TYPES.map((t) => (
                <button type="button" key={t} className={`wd-type-btn ${formData.type === t ? 'active' : ''}`} style={formData.type === t ? { background: `${txColor(t)}22`, color: txColor(t), borderColor: txColor(t) } : {}} onClick={() => setTxType(t)}>{t}</button>
              ))}
            </div>

            <div className="wd-row-2">
              <div className="wd-field"><label>Date</label><input type="date" value={formData.date} onChange={(e) => setFormData((f) => ({ ...f, date: e.target.value }))} /></div>
              <div className="wd-field"><label>Amount ({CURRENCY})</label><input className="amount-input" type="number" min="0.01" step="0.01" inputMode="decimal" placeholder="0.00" value={formData.amount} onChange={(e) => setFormData((f) => ({ ...f, amount: e.target.value }))} /></div>
            </div>

            <div className="wd-row-2">
              <div className="wd-field">
                <label>Category</label>
                <input list="wd-cat-options" placeholder="e.g. Groceries" value={formData.category} onChange={(e) => setFormData((f) => ({ ...f, category: e.target.value }))} />
                <datalist id="wd-cat-options">{catOptions.map((c) => <option value={c} key={c} />)}</datalist>
              </div>
              <div className="wd-field">
                <label>Account</label>
                <input list="wd-acc-options" placeholder="e.g. HDFC Savings" value={formData.account} onChange={(e) => setFormData((f) => ({ ...f, account: e.target.value }))} />
                <datalist id="wd-acc-options">{accountOptions.map((a) => <option value={a} key={a} />)}</datalist>
              </div>
            </div>

            <div className="wd-row-2">
              <div className="wd-field"><label>Note (optional)</label><input type="text" placeholder="e.g. Dinner with friends" value={formData.note} onChange={(e) => setFormData((f) => ({ ...f, note: e.target.value }))} /></div>
              <div className="wd-field"><label>Tags (optional, comma separated)</label><input type="text" placeholder="e.g. essential, family" value={formData.tags} onChange={(e) => setFormData((f) => ({ ...f, tags: e.target.value }))} /></div>
            </div>

            <div className="wd-recurring-row">
              <label className="wd-checkbox"><input type="checkbox" checked={formData.recurring} onChange={(e) => setFormData((f) => ({ ...f, recurring: e.target.checked }))} /> This repeats</label>
              {formData.recurring && (
                <select value={formData.recurrenceFrequency} onChange={(e) => setFormData((f) => ({ ...f, recurrenceFrequency: e.target.value }))}>
                  <option value="weekly">Weekly</option>
                  <option value="monthly">Monthly</option>
                  <option value="yearly">Yearly</option>
                </select>
              )}
            </div>

            {formError && <div className="wd-form-error">{formError}</div>}
            <div className="wd-form-actions">
              <button type="button" className="wd-btn-ghost" onClick={closeForm}>Cancel</button>
              <button type="submit" className="wd-btn-primary">{editingId ? 'Save changes' : 'Add transaction'}</button>
            </div>
          </form>
        )}

        {filtered.length === 0 ? (
          <div className="wd-empty">No transactions match this period and filter.</div>
        ) : (
          <div className="wd-ledger">
            {filtered.map((t) => (
              <div className="wd-tx-row" key={t.id}>
                <div className="wd-entry-spine" style={{ background: txColor(t.type) }} />
                <div className="wd-entry-date wd-mono">{formatShortDate(t.date)}</div>
                <div className="wd-entry-main">
                  <span className="wd-entry-category">{t.category} {t.recurring && <Repeat size={11} className="wd-recur-icon" />}</span>
                  <span className="wd-entry-sub">{[t.account, t.note].filter(Boolean).join(' · ')}</span>
                  {t.tags && <span className="wd-tag-row">{t.tags.split(',').map((tag) => tag.trim()).filter(Boolean).map((tag) => <span className="wd-tag" key={tag}>{tag}</span>)}</span>}
                </div>
                <span className="wd-entry-amount wd-mono" style={{ color: txColor(t.type) }}>{amountPrefix(t.type)}{fmt(t.amount)}</span>
                <div className="wd-entry-actions">
                  {confirmDeleteId === t.id ? (
                    <div className="wd-confirm"><button className="yes" onClick={() => { onDelete(t.id); setConfirmDeleteId(null); }}>Delete</button><button onClick={() => setConfirmDeleteId(null)}>Cancel</button></div>
                  ) : (
                    <>
                      <button className="wd-icon-btn" aria-label="Duplicate" onClick={() => onDuplicate(t)}><Copy size={14} /></button>
                      <button className="wd-icon-btn" aria-label="Edit" onClick={() => openEdit(t)}><Pencil size={14} /></button>
                      <button className="wd-icon-btn" aria-label="Delete" onClick={() => setConfirmDeleteId(t.id)}><Trash2 size={14} /></button>
                    </>
                  )}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </>
  );
}

/* ============================== NET WORTH MODULE ============================== */

function emptyAssetForm() { return { category: ASSET_CATEGORIES[0], name: '', value: '' }; }
function emptyLiabilityForm() { return { category: LIABILITY_CATEGORIES[0], name: '', amount: '', emi: '', interestRate: '', remainingTenureMonths: '' }; }

function NetWorthModule({ netWorth, setNetWorth, totals, allocation }) {
  const [showAssetForm, setShowAssetForm] = useState(false);
  const [editingAssetId, setEditingAssetId] = useState(null);
  const [assetForm, setAssetForm] = useState(emptyAssetForm());
  const [assetError, setAssetError] = useState('');
  const [confirmDeleteAsset, setConfirmDeleteAsset] = useState(null);

  const [showLiabForm, setShowLiabForm] = useState(false);
  const [editingLiabId, setEditingLiabId] = useState(null);
  const [liabForm, setLiabForm] = useState(emptyLiabilityForm());
  const [liabError, setLiabError] = useState('');
  const [confirmDeleteLiab, setConfirmDeleteLiab] = useState(null);

  const allocationData = useMemo(() => capAllocation(allocation), [allocation]);

  function openAddAsset() { setEditingAssetId(null); setAssetError(''); setAssetForm(emptyAssetForm()); setShowAssetForm(true); }
  function openEditAsset(a) { setEditingAssetId(a.id); setAssetError(''); setAssetForm({ category: a.category, name: a.name, value: String(a.value) }); setShowAssetForm(true); }
  function closeAssetForm() { setShowAssetForm(false); setEditingAssetId(null); setAssetError(''); }
  function submitAsset(e) {
    e.preventDefault();
    if (!assetForm.name.trim()) { setAssetError('Give this asset a name or label.'); return; }
    const val = parseFloat(assetForm.value);
    if (isNaN(val) || val < 0) { setAssetError('Enter a value of 0 or more.'); return; }
    const payload = { category: assetForm.category, name: assetForm.name.trim(), value: val };
    setNetWorth((prev) => ({
      ...prev,
      assets: editingAssetId ? prev.assets.map((a) => (a.id === editingAssetId ? { ...a, ...payload } : a)) : [...prev.assets, { id: uid(), ...payload }],
    }));
    closeAssetForm();
  }
  function deleteAsset(id) { setNetWorth((prev) => ({ ...prev, assets: prev.assets.filter((a) => a.id !== id) })); setConfirmDeleteAsset(null); }

  function openAddLiab() { setEditingLiabId(null); setLiabError(''); setLiabForm(emptyLiabilityForm()); setShowLiabForm(true); }
  function openEditLiab(l) { setEditingLiabId(l.id); setLiabError(''); setLiabForm({ category: l.category, name: l.name, amount: String(l.amount), emi: l.emi ? String(l.emi) : '', interestRate: l.interestRate ? String(l.interestRate) : '', remainingTenureMonths: l.remainingTenureMonths ? String(l.remainingTenureMonths) : '' }); setShowLiabForm(true); }
  function closeLiabForm() { setShowLiabForm(false); setEditingLiabId(null); setLiabError(''); }
  function submitLiab(e) {
    e.preventDefault();
    if (!liabForm.name.trim()) { setLiabError('Give this liability a name or label.'); return; }
    const amt = parseFloat(liabForm.amount);
    if (isNaN(amt) || amt < 0) { setLiabError('Enter an outstanding amount of 0 or more.'); return; }
    const payload = {
      category: liabForm.category, name: liabForm.name.trim(), amount: amt,
      emi: liabForm.emi ? parseFloat(liabForm.emi) || 0 : 0,
      interestRate: liabForm.interestRate ? parseFloat(liabForm.interestRate) || 0 : 0,
      remainingTenureMonths: liabForm.remainingTenureMonths ? parseInt(liabForm.remainingTenureMonths) || 0 : 0,
    };
    setNetWorth((prev) => ({
      ...prev,
      liabilities: editingLiabId ? prev.liabilities.map((l) => (l.id === editingLiabId ? { ...l, ...payload } : l)) : [...prev.liabilities, { id: uid(), ...payload }],
    }));
    closeLiabForm();
  }
  function deleteLiab(id) { setNetWorth((prev) => ({ ...prev, liabilities: prev.liabilities.filter((l) => l.id !== id) })); setConfirmDeleteLiab(null); }

  return (
    <>
      <div className="wd-summary-grid wd-summary-grid-4">
        <div className="wd-summary-card"><span className="sc-label">Total assets</span><span className="sc-value wd-mono">{fmtCompact(totals.totalAssets)}</span></div>
        <div className="wd-summary-card out"><span className="sc-label">Total liabilities</span><span className="sc-value wd-mono">{fmtCompact(totals.totalLiabilities)}</span></div>
        <div className="wd-summary-card net"><span className="sc-label">Net worth</span><span className="sc-value wd-mono">{fmtCompact(totals.netWorthValue)}</span></div>
        <div className="wd-summary-card in"><span className="sc-label">Liquid cash</span><span className="sc-value wd-mono">{fmtCompact(totals.liquidCash)}</span></div>
      </div>

      <div className="wd-card">
        <h3 className="wd-section-title">Allocation</h3>
        {allocationData.length === 0 ? <div className="wd-empty">Add an asset below to see your allocation.</div> : (
          <div style={{ width: '100%', height: 200 }}>
            <ResponsiveContainer>
              <PieChart>
                <Pie data={allocationData} dataKey="value" nameKey="name" innerRadius="55%" outerRadius="85%" paddingAngle={2}>
                  {allocationData.map((entry, i) => <Cell key={i} fill={categoryColor(entry.name)} stroke="var(--card)" strokeWidth={2} />)}
                </Pie>
                <Tooltip content={<CurrencyTooltip />} />
              </PieChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>

      <div className="wd-card">
        <div className="wd-section-header">
          <h3 className="wd-section-title">Assets</h3>
          <div className="wd-header-actions">
            <button className="wd-btn-ghost" onClick={() => downloadCSV('assets.csv', netWorth.assets.map((a) => ({ Category: a.category, Name: a.name, Value: a.value })))}><Download size={13} /> CSV</button>
            <button className="wd-btn-primary" onClick={() => (showAssetForm && !editingAssetId ? closeAssetForm() : openAddAsset())}><Plus size={15} /> Add asset</button>
          </div>
        </div>

        {showAssetForm && (
          <form className="wd-form" onSubmit={submitAsset}>
            <div className="wd-row-2">
              <div className="wd-field"><label>Category</label>
                <select value={assetForm.category} onChange={(e) => setAssetForm((f) => ({ ...f, category: e.target.value }))}>
                  {ASSET_CATEGORIES.map((c) => <option value={c} key={c}>{c}</option>)}
                </select>
              </div>
              <div className="wd-field"><label>Value ({CURRENCY})</label><input className="amount-input" type="number" min="0" step="0.01" value={assetForm.value} onChange={(e) => setAssetForm((f) => ({ ...f, value: e.target.value }))} /></div>
            </div>
            <div className="wd-field"><label>Name / label</label><input type="text" placeholder="e.g. HDFC Savings" value={assetForm.name} onChange={(e) => setAssetForm((f) => ({ ...f, name: e.target.value }))} /></div>
            {assetError && <div className="wd-form-error">{assetError}</div>}
            <div className="wd-form-actions"><button type="button" className="wd-btn-ghost" onClick={closeAssetForm}>Cancel</button><button type="submit" className="wd-btn-primary">{editingAssetId ? 'Save changes' : 'Add asset'}</button></div>
          </form>
        )}

        {netWorth.assets.length === 0 ? <div className="wd-empty">No assets added yet.</div> : (
          <div className="wd-ledger">
            {netWorth.assets.map((a) => (
              <div className="wd-simple-row" key={a.id}>
                <span className="wd-row-dot" style={{ background: categoryColor(a.category) }} />
                <div className="wd-entry-main"><span className="wd-entry-category">{a.name}</span><span className="wd-entry-sub">{a.category}</span></div>
                <span className="wd-entry-amount wd-mono">{fmt(a.value)}</span>
                <div className="wd-entry-actions">
                  {confirmDeleteAsset === a.id ? (
                    <div className="wd-confirm"><button className="yes" onClick={() => deleteAsset(a.id)}>Delete</button><button onClick={() => setConfirmDeleteAsset(null)}>Cancel</button></div>
                  ) : (<><button className="wd-icon-btn" onClick={() => openEditAsset(a)}><Pencil size={14} /></button><button className="wd-icon-btn" onClick={() => setConfirmDeleteAsset(a.id)}><Trash2 size={14} /></button></>)}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      <div className="wd-card">
        <div className="wd-section-header">
          <h3 className="wd-section-title">Liabilities</h3>
          <div className="wd-header-actions">
            <button className="wd-btn-ghost" onClick={() => downloadCSV('liabilities.csv', netWorth.liabilities.map((l) => ({ Category: l.category, Name: l.name, Amount: l.amount, EMI: l.emi, InterestRate: l.interestRate, RemainingTenureMonths: l.remainingTenureMonths })))}><Download size={13} /> CSV</button>
            <button className="wd-btn-primary" onClick={() => (showLiabForm && !editingLiabId ? closeLiabForm() : openAddLiab())}><Plus size={15} /> Add liability</button>
          </div>
        </div>

        {showLiabForm && (
          <form className="wd-form" onSubmit={submitLiab}>
            <div className="wd-row-2">
              <div className="wd-field"><label>Category</label>
                <select value={liabForm.category} onChange={(e) => setLiabForm((f) => ({ ...f, category: e.target.value }))}>
                  {LIABILITY_CATEGORIES.map((c) => <option value={c} key={c}>{c}</option>)}
                </select>
              </div>
              <div className="wd-field"><label>Outstanding amount ({CURRENCY})</label><input className="amount-input" type="number" min="0" step="0.01" value={liabForm.amount} onChange={(e) => setLiabForm((f) => ({ ...f, amount: e.target.value }))} /></div>
            </div>
            <div className="wd-field"><label>Name / label</label><input type="text" placeholder="e.g. HDFC Home Loan" value={liabForm.name} onChange={(e) => setLiabForm((f) => ({ ...f, name: e.target.value }))} /></div>
            <div className="wd-row-3">
              <div className="wd-field"><label>EMI (optional)</label><input className="amount-input" type="number" min="0" step="0.01" value={liabForm.emi} onChange={(e) => setLiabForm((f) => ({ ...f, emi: e.target.value }))} /></div>
              <div className="wd-field"><label>Interest rate % (optional)</label><input className="amount-input" type="number" min="0" step="0.01" value={liabForm.interestRate} onChange={(e) => setLiabForm((f) => ({ ...f, interestRate: e.target.value }))} /></div>
              <div className="wd-field"><label>Tenure left, months (optional)</label><input className="amount-input" type="number" min="0" step="1" value={liabForm.remainingTenureMonths} onChange={(e) => setLiabForm((f) => ({ ...f, remainingTenureMonths: e.target.value }))} /></div>
            </div>
            {liabError && <div className="wd-form-error">{liabError}</div>}
            <div className="wd-form-actions"><button type="button" className="wd-btn-ghost" onClick={closeLiabForm}>Cancel</button><button type="submit" className="wd-btn-primary">{editingLiabId ? 'Save changes' : 'Add liability'}</button></div>
          </form>
        )}

        {netWorth.liabilities.length === 0 ? <div className="wd-empty">No liabilities added — nicely done, or add one to track.</div> : (
          <div className="wd-ledger">
            {netWorth.liabilities.map((l) => (
              <div className="wd-simple-row" key={l.id}>
                <span className="wd-row-dot" style={{ background: 'var(--clay)' }} />
                <div className="wd-entry-main">
                  <span className="wd-entry-category">{l.name}</span>
                  <span className="wd-entry-sub">{l.category}{l.emi ? ` · EMI ${fmt(l.emi)}` : ''}{l.interestRate ? ` · ${l.interestRate}%` : ''}{l.remainingTenureMonths ? ` · ${l.remainingTenureMonths} mo left` : ''}</span>
                </div>
                <span className="wd-entry-amount wd-mono" style={{ color: 'var(--clay)' }}>{fmt(l.amount)}</span>
                <div className="wd-entry-actions">
                  {confirmDeleteLiab === l.id ? (
                    <div className="wd-confirm"><button className="yes" onClick={() => deleteLiab(l.id)}>Delete</button><button onClick={() => setConfirmDeleteLiab(null)}>Cancel</button></div>
                  ) : (<><button className="wd-icon-btn" onClick={() => openEditLiab(l)}><Pencil size={14} /></button><button className="wd-icon-btn" onClick={() => setConfirmDeleteLiab(l.id)}><Trash2 size={14} /></button></>)}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>
    </>
  );
}

/* ============================== INVESTMENTS MODULE ============================== */

function emptyInvestmentForm() { return { category: INVESTMENT_CATEGORIES[0], name: '', investedAmount: '', currentValue: '', purchaseDate: todayISO(), isSIP: false, monthlySIP: '' }; }

function InvestmentsModule({ investments, setInvestments, totals }) {
  const [showForm, setShowForm] = useState(false);
  const [editingId, setEditingId] = useState(null);
  const [formData, setFormData] = useState(emptyInvestmentForm());
  const [formError, setFormError] = useState('');
  const [confirmDeleteId, setConfirmDeleteId] = useState(null);

  const allocationData = useMemo(() => {
    const list = Object.entries(totals.byCategory).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
    return capAllocation(list);
  }, [totals.byCategory]);

  function openAdd() { setEditingId(null); setFormError(''); setFormData(emptyInvestmentForm()); setShowForm(true); }
  function openEdit(inv) {
    setEditingId(inv.id); setFormError('');
    setFormData({ category: inv.category, name: inv.name, investedAmount: String(inv.investedAmount), currentValue: String(inv.currentValue), purchaseDate: inv.purchaseDate, isSIP: !!inv.isSIP, monthlySIP: inv.monthlySIP ? String(inv.monthlySIP) : '' });
    setShowForm(true);
  }
  function closeForm() { setShowForm(false); setEditingId(null); setFormError(''); }

  function submit(e) {
    e.preventDefault();
    if (!formData.name.trim()) { setFormError('Give this holding a name.'); return; }
    const invested = parseFloat(formData.investedAmount);
    const current = parseFloat(formData.currentValue);
    if (isNaN(invested) || invested < 0) { setFormError('Enter the amount invested.'); return; }
    if (isNaN(current) || current < 0) { setFormError('Enter the current value.'); return; }
    if (!formData.purchaseDate) { setFormError('Pick a purchase or start date.'); return; }
    const payload = { category: formData.category, name: formData.name.trim(), investedAmount: invested, currentValue: current, purchaseDate: formData.purchaseDate, isSIP: formData.isSIP, monthlySIP: formData.isSIP ? (parseFloat(formData.monthlySIP) || 0) : 0 };
    if (editingId) setInvestments((prev) => prev.map((i) => (i.id === editingId ? { ...i, ...payload } : i)));
    else setInvestments((prev) => [...prev, { id: uid(), ...payload }]);
    closeForm();
  }
  function remove(id) { setInvestments((prev) => prev.filter((i) => i.id !== id)); setConfirmDeleteId(null); }

  return (
    <>
      <div className="wd-summary-grid wd-summary-grid-4">
        <div className="wd-summary-card"><span className="sc-label">Invested capital</span><span className="sc-value wd-mono">{fmtCompact(totals.totalInvested)}</span></div>
        <div className="wd-summary-card in"><span className="sc-label">Current value</span><span className="sc-value wd-mono">{fmtCompact(totals.totalCurrentValue)}</span></div>
        <div className={`wd-summary-card ${totals.totalGain < 0 ? 'out' : 'in'}`}><span className="sc-label">Gain / loss</span><span className="sc-value wd-mono">{fmtSigned(totals.totalGain)}</span></div>
        <div className="wd-summary-card net"><span className="sc-label">Return</span><span className="sc-value wd-mono">{totals.gainPct >= 0 ? '+' : ''}{totals.gainPct.toFixed(1)}%</span></div>
      </div>

      {totals.monthlySIPTotal > 0 && (
        <div className="wd-capital-note"><Info size={13} /><span>Active monthly SIP commitments total {fmt(totals.monthlySIPTotal)}.</span></div>
      )}

      <div className="wd-card">
        <h3 className="wd-section-title">Portfolio allocation</h3>
        {allocationData.length === 0 ? <div className="wd-empty">Add a holding to see your allocation.</div> : (
          <div style={{ width: '100%', height: 200 }}>
            <ResponsiveContainer>
              <PieChart>
                <Pie data={allocationData} dataKey="value" nameKey="name" innerRadius="55%" outerRadius="85%" paddingAngle={2}>
                  {allocationData.map((entry, i) => <Cell key={i} fill={categoryColor(entry.name)} stroke="var(--card)" strokeWidth={2} />)}
                </Pie>
                <Tooltip content={<CurrencyTooltip />} />
              </PieChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>

      <div className="wd-card">
        <div className="wd-section-header">
          <h3 className="wd-section-title">Holdings</h3>
          <div className="wd-header-actions">
            <button className="wd-btn-ghost" onClick={() => downloadCSV('investments.csv', investments.map((i) => ({ Category: i.category, Name: i.name, Invested: i.investedAmount, CurrentValue: i.currentValue, PurchaseDate: i.purchaseDate, SIP: i.isSIP, MonthlySIP: i.monthlySIP })))}><Download size={13} /> CSV</button>
            <button className="wd-btn-primary" onClick={() => (showForm && !editingId ? closeForm() : openAdd())}><Plus size={15} /> Add holding</button>
          </div>
        </div>

        {showForm && (
          <form className="wd-form" onSubmit={submit}>
            <div className="wd-row-2">
              <div className="wd-field"><label>Category</label>
                <select value={formData.category} onChange={(e) => setFormData((f) => ({ ...f, category: e.target.value }))}>
                  {INVESTMENT_CATEGORIES.map((c) => <option value={c} key={c}>{c}</option>)}
                </select>
              </div>
              <div className="wd-field"><label>Purchase / start date</label><input type="date" value={formData.purchaseDate} onChange={(e) => setFormData((f) => ({ ...f, purchaseDate: e.target.value }))} /></div>
            </div>
            <div className="wd-field"><label>Name / label</label><input type="text" placeholder="e.g. Nifty 50 Index Fund" value={formData.name} onChange={(e) => setFormData((f) => ({ ...f, name: e.target.value }))} /></div>
            <div className="wd-row-2">
              <div className="wd-field"><label>Invested amount ({CURRENCY})</label><input className="amount-input" type="number" min="0" step="0.01" value={formData.investedAmount} onChange={(e) => setFormData((f) => ({ ...f, investedAmount: e.target.value }))} /></div>
              <div className="wd-field"><label>Current value ({CURRENCY})</label><input className="amount-input" type="number" min="0" step="0.01" value={formData.currentValue} onChange={(e) => setFormData((f) => ({ ...f, currentValue: e.target.value }))} /></div>
            </div>
            <div className="wd-recurring-row">
              <label className="wd-checkbox"><input type="checkbox" checked={formData.isSIP} onChange={(e) => setFormData((f) => ({ ...f, isSIP: e.target.checked }))} /> Active SIP</label>
              {formData.isSIP && <input className="amount-input wd-sip-input" type="number" min="0" step="0.01" placeholder="Monthly SIP amount" value={formData.monthlySIP} onChange={(e) => setFormData((f) => ({ ...f, monthlySIP: e.target.value }))} />}
            </div>
            {formError && <div className="wd-form-error">{formError}</div>}
            <div className="wd-form-actions"><button type="button" className="wd-btn-ghost" onClick={closeForm}>Cancel</button><button type="submit" className="wd-btn-primary">{editingId ? 'Save changes' : 'Add holding'}</button></div>
          </form>
        )}

        {investments.length === 0 ? <div className="wd-empty">No holdings added yet.</div> : (
          <div className="wd-ledger">
            {investments.map((inv) => {
              const gain = inv.currentValue - inv.investedAmount;
              const cagr = simpleCAGR(inv.investedAmount, inv.currentValue, inv.purchaseDate);
              return (
                <div className="wd-simple-row" key={inv.id}>
                  <span className="wd-row-dot" style={{ background: categoryColor(inv.category) }} />
                  <div className="wd-entry-main">
                    <span className="wd-entry-category">{inv.name} {inv.isSIP && <Repeat size={11} className="wd-recur-icon" />}</span>
                    <span className="wd-entry-sub">{inv.category} · Invested {fmt(inv.investedAmount)}{cagr !== null ? ` · ${cagr >= 0 ? '+' : ''}${cagr.toFixed(1)}% CAGR` : ''}</span>
                  </div>
                  <div className="wd-inv-values">
                    <span className="wd-entry-amount wd-mono">{fmt(inv.currentValue)}</span>
                    <span className={`wd-inv-gain wd-mono ${gain < 0 ? 'clay' : 'sage'}`}>{gain >= 0 ? '+' : ''}{fmt(gain)}</span>
                  </div>
                  <div className="wd-entry-actions">
                    {confirmDeleteId === inv.id ? (
                      <div className="wd-confirm"><button className="yes" onClick={() => remove(inv.id)}>Delete</button><button onClick={() => setConfirmDeleteId(null)}>Cancel</button></div>
                    ) : (<><button className="wd-icon-btn" onClick={() => openEdit(inv)}><Pencil size={14} /></button><button className="wd-icon-btn" onClick={() => setConfirmDeleteId(inv.id)}><Trash2 size={14} /></button></>)}
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </>
  );
}

/* ============================== GOALS & FIRE MODULE ============================== */

function emptyGoalForm() { return { name: '', type: GOAL_TYPES[0], targetAmount: '', currentAmount: '', deadline: '' }; }

function monthsBetween(today, deadline) {
  const days = (deadline.getTime() - today.getTime()) / 86400000;
  return Math.max(1, Math.round(days / 30.44));
}

function GoalsFireModule({ goalsFire, setGoalsFire, retirement, fire, investableNetWorth }) {
  const [showGoalForm, setShowGoalForm] = useState(false);
  const [editingGoalId, setEditingGoalId] = useState(null);
  const [goalForm, setGoalForm] = useState(emptyGoalForm());
  const [goalError, setGoalError] = useState('');
  const [confirmDeleteGoal, setConfirmDeleteGoal] = useState(null);

  function openAddGoal() { setEditingGoalId(null); setGoalError(''); setGoalForm(emptyGoalForm()); setShowGoalForm(true); }
  function openEditGoal(g) { setEditingGoalId(g.id); setGoalError(''); setGoalForm({ name: g.name, type: g.type, targetAmount: String(g.targetAmount), currentAmount: String(g.currentAmount), deadline: g.deadline }); setShowGoalForm(true); }
  function closeGoalForm() { setShowGoalForm(false); setEditingGoalId(null); setGoalError(''); }

  function submitGoal(e) {
    e.preventDefault();
    if (!goalForm.name.trim()) { setGoalError('Give this goal a name.'); return; }
    const target = parseFloat(goalForm.targetAmount);
    const current = parseFloat(goalForm.currentAmount || '0');
    if (isNaN(target) || target <= 0) { setGoalError('Enter a target amount greater than 0.'); return; }
    if (isNaN(current) || current < 0) { setGoalError('Enter a current amount of 0 or more.'); return; }
    if (!goalForm.deadline) { setGoalError('Pick a target date.'); return; }
    const payload = { name: goalForm.name.trim(), type: goalForm.type, targetAmount: target, currentAmount: current, deadline: goalForm.deadline };
    setGoalsFire((prev) => ({
      ...prev,
      goals: editingGoalId ? prev.goals.map((g) => (g.id === editingGoalId ? { ...g, ...payload } : g)) : [...prev.goals, { id: uid(), ...payload }],
    }));
    closeGoalForm();
  }
  function deleteGoal(id) { setGoalsFire((prev) => ({ ...prev, goals: prev.goals.filter((g) => g.id !== id) })); setConfirmDeleteGoal(null); }

  function setRetirementField(field, value) { setGoalsFire((prev) => ({ ...prev, retirement: { ...prev.retirement, [field]: value } })); }
  function setFireField(field, value) { setGoalsFire((prev) => ({ ...prev, fireSettings: { ...prev.fireSettings, [field]: value } })); }

  const coastNumber = coastFireNumber({ currentAge: goalsFire.retirement.currentAge, retirementAge: goalsFire.retirement.retirementAge, expectedReturn: goalsFire.fireSettings.expectedReturn, targetCorpus: fire.fiTarget });
  const isCoastFI = investableNetWorth >= coastNumber;

  return (
    <>
      <div className="wd-card">
        <div className="wd-section-header">
          <h3 className="wd-section-title">Goals</h3>
          <button className="wd-btn-primary" onClick={() => (showGoalForm && !editingGoalId ? closeGoalForm() : openAddGoal())}><Plus size={15} /> Add goal</button>
        </div>

        {showGoalForm && (
          <form className="wd-form" onSubmit={submitGoal}>
            <div className="wd-row-2">
              <div className="wd-field"><label>Goal name</label><input type="text" placeholder="e.g. Emergency Fund" value={goalForm.name} onChange={(e) => setGoalForm((f) => ({ ...f, name: e.target.value }))} /></div>
              <div className="wd-field"><label>Type</label><select value={goalForm.type} onChange={(e) => setGoalForm((f) => ({ ...f, type: e.target.value }))}>{GOAL_TYPES.map((t) => <option value={t} key={t}>{t}</option>)}</select></div>
            </div>
            <div className="wd-row-2">
              <div className="wd-field"><label>Target amount ({CURRENCY})</label><input className="amount-input" type="number" min="0" step="0.01" value={goalForm.targetAmount} onChange={(e) => setGoalForm((f) => ({ ...f, targetAmount: e.target.value }))} /></div>
              <div className="wd-field"><label>Current amount ({CURRENCY})</label><input className="amount-input" type="number" min="0" step="0.01" value={goalForm.currentAmount} onChange={(e) => setGoalForm((f) => ({ ...f, currentAmount: e.target.value }))} /></div>
            </div>
            <div className="wd-field"><label>Target date</label><input type="date" value={goalForm.deadline} onChange={(e) => setGoalForm((f) => ({ ...f, deadline: e.target.value }))} /></div>
            {goalError && <div className="wd-form-error">{goalError}</div>}
            <div className="wd-form-actions"><button type="button" className="wd-btn-ghost" onClick={closeGoalForm}>Cancel</button><button type="submit" className="wd-btn-primary">{editingGoalId ? 'Save changes' : 'Add goal'}</button></div>
          </form>
        )}

        {goalsFire.goals.length === 0 ? <div className="wd-empty">No goals yet — add one to start tracking progress.</div> : (
          <div className="wd-goal-list">
            {goalsFire.goals.map((g) => {
              const pct = Math.min(100, (g.currentAmount / g.targetAmount) * 100);
              const months = monthsBetween(new Date(), parseISODate(g.deadline));
              const monthlyNeeded = Math.max(0, g.targetAmount - g.currentAmount) / months;
              const overdue = parseISODate(g.deadline) < new Date() && pct < 100;
              return (
                <div className="wd-goal-card" key={g.id}>
                  <div className="wd-goal-top">
                    <div>
                      <span className="wd-goal-name">{g.name}</span>
                      <span className="wd-goal-type">{g.type}</span>
                    </div>
                    <div className="wd-entry-actions">
                      {confirmDeleteGoal === g.id ? (
                        <div className="wd-confirm"><button className="yes" onClick={() => deleteGoal(g.id)}>Delete</button><button onClick={() => setConfirmDeleteGoal(null)}>Cancel</button></div>
                      ) : (<><button className="wd-icon-btn" onClick={() => openEditGoal(g)}><Pencil size={14} /></button><button className="wd-icon-btn" onClick={() => setConfirmDeleteGoal(g.id)}><Trash2 size={14} /></button></>)}
                    </div>
                  </div>
                  <div className="wd-bar-track wd-goal-track"><div className="wd-bar-fill" style={{ width: `${pct}%`, background: 'var(--sage)' }} /></div>
                  <div className="wd-goal-stats wd-mono">
                    <span>{fmt(g.currentAmount)} of {fmt(g.targetAmount)} · {pct.toFixed(0)}%</span>
                    <span className={overdue ? 'clay' : ''}>{overdue ? 'Past target date' : `${fmtCompact(monthlyNeeded)}/mo needed · ${months} mo left`}</span>
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      <div className="wd-card">
        <h3 className="wd-section-title">Retirement planner</h3>
        <p className="wd-muted-text">Projects your retirement corpus from your current savings, monthly investing, and expected returns — adjusted for inflation.</p>
        <div className="wd-row-3">
          <div className="wd-field"><label>Current age</label><input className="amount-input" type="number" min="0" value={goalsFire.retirement.currentAge} onChange={(e) => setRetirementField('currentAge', Number(e.target.value))} /></div>
          <div className="wd-field"><label>Retirement age</label><input className="amount-input" type="number" min="0" value={goalsFire.retirement.retirementAge} onChange={(e) => setRetirementField('retirementAge', Number(e.target.value))} /></div>
          <div className="wd-field"><label>Current corpus ({CURRENCY})</label><input className="amount-input" type="number" min="0" value={goalsFire.retirement.currentCorpus} onChange={(e) => setRetirementField('currentCorpus', Number(e.target.value))} /></div>
        </div>
        <div className="wd-row-3">
          <div className="wd-field"><label>Monthly investment ({CURRENCY})</label><input className="amount-input" type="number" min="0" value={goalsFire.retirement.monthlyInvestment} onChange={(e) => setRetirementField('monthlyInvestment', Number(e.target.value))} /></div>
          <div className="wd-field"><label>Expected return % p.a.</label><input className="amount-input" type="number" min="0" step="0.1" value={goalsFire.retirement.expectedReturn} onChange={(e) => setRetirementField('expectedReturn', Number(e.target.value))} /></div>
          <div className="wd-field"><label>Inflation % p.a.</label><input className="amount-input" type="number" min="0" step="0.1" value={goalsFire.retirement.inflation} onChange={(e) => setRetirementField('inflation', Number(e.target.value))} /></div>
        </div>
        <div className="wd-field"><label>Annual expenses today ({CURRENCY})</label><input className="amount-input" type="number" min="0" value={goalsFire.retirement.annualExpenses} onChange={(e) => setRetirementField('annualExpenses', Number(e.target.value))} /></div>

        <div className="wd-summary-grid wd-summary-grid-4 wd-mt">
          <div className="wd-summary-card"><span className="sc-label">Years to retirement</span><span className="sc-value wd-mono">{retirement.years}</span></div>
          <div className="wd-summary-card in"><span className="sc-label">Projected corpus</span><span className="sc-value wd-mono">{fmtCompact(retirement.projectedCorpus)}</span></div>
          <div className="wd-summary-card out"><span className="sc-label">Target corpus</span><span className="sc-value wd-mono">{fmtCompact(retirement.targetCorpus)}</span></div>
          <div className="wd-summary-card net"><span className="sc-label">Readiness</span><span className="sc-value wd-mono">{retirement.readiness.toFixed(0)}%</span></div>
        </div>
      </div>

      <div className="wd-card">
        <h3 className="wd-section-title">Financial independence (FIRE)</h3>
        <p className="wd-muted-text">FI number = annual expenses × multiplier (25x by default, the classic 4% rule). Progress compares this to your investable net worth — primary residence and vehicles excluded.</p>
        <div className="wd-row-3">
          <div className="wd-field"><label>FI multiplier (×)</label><input className="amount-input" type="number" min="1" step="0.5" value={goalsFire.fireSettings.fiMultiplier} onChange={(e) => setFireField('fiMultiplier', Number(e.target.value))} /></div>
          <div className="wd-field"><label>Expected return % p.a.</label><input className="amount-input" type="number" min="0" step="0.1" value={goalsFire.fireSettings.expectedReturn} onChange={(e) => setFireField('expectedReturn', Number(e.target.value))} /></div>
          <div className="wd-field"><label>Monthly investment ({CURRENCY})</label><input className="amount-input" type="number" min="0" value={goalsFire.fireSettings.monthlyInvestment} onChange={(e) => setFireField('monthlyInvestment', Number(e.target.value))} /></div>
        </div>
        <div className="wd-summary-grid wd-summary-grid-4 wd-mt">
          <div className="wd-summary-card"><span className="sc-label">FI target</span><span className="sc-value wd-mono">{fmtCompact(fire.fiTarget)}</span></div>
          <div className="wd-summary-card in"><span className="sc-label">Current progress</span><span className="sc-value wd-mono">{fire.progress.toFixed(0)}%</span></div>
          <div className="wd-summary-card"><span className="sc-label">Years remaining</span><span className="sc-value wd-mono">{fire.yearsRemaining === null ? '—' : fire.yearsRemaining}</span></div>
          <div className="wd-summary-card net"><span className="sc-label">Projected FI date</span><span className="sc-value wd-mono">{fire.projectedDate ? fire.projectedDate.toLocaleDateString('en-IN', { month: 'short', year: 'numeric' }) : '—'}</span></div>
        </div>
        <div className="wd-coast-note">
          <Info size={13} />
          <span>Coast FIRE number at today's age: <strong className="wd-mono">{fmt(coastNumber)}</strong> — {isCoastFI ? "you've already crossed it, so future contributions are optional to reach your FI target by retirement age." : "the amount you'd need invested today, growing untouched, to still hit your FI target by retirement age."}</span>
        </div>
      </div>
    </>
  );
}

/* ============================== TAX & SAFETY MODULE ============================== */

function emptyInsuranceForm() { return { type: INSURANCE_TYPES[0], provider: '', coverage: '', premium: '', renewalDate: '' }; }
function emptyCardForm() { return { name: '', limit: '', outstanding: '', dueDate: '', billingDay: '' }; }
function daysUntil(dateISO) { return Math.round((parseISODate(dateISO).getTime() - startOfDay(new Date()).getTime()) / 86400000); }

function TaxSafetyModule({ protection, setProtection, netWorth, netWorthTotals, emergencyMonths, debtToIncomePct, debtToAssetPct, annualIncomeEstimate }) {
  const ti = protection.taxInputs;
  function setTaxField(field, value) { setProtection((prev) => ({ ...prev, taxInputs: { ...prev.taxInputs, [field]: value } })); }

  const newRegime = useMemo(() => computeNewRegimeTax(ti.grossAnnualIncome, ti.isSalaried), [ti.grossAnnualIncome, ti.isSalaried]);
  const oldRegime = useMemo(() => computeOldRegimeTax(ti.grossAnnualIncome, ti.isSalaried, ti.ageBand, { section80C: ti.section80C, section80D: ti.section80D, nps80CCD1B: ti.nps80CCD1B, homeLoanInterest: ti.homeLoanInterest, other: ti.otherDeductions }), [ti]);
  const recommended = newRegime.total <= oldRegime.total ? 'new' : 'old';
  const chosen = recommended === 'new' ? newRegime : oldRegime;
  const taxSavingsByChoosing = Math.abs(newRegime.total - oldRegime.total);
  const effectiveRate = ti.grossAnnualIncome > 0 ? (chosen.total / ti.grossAnnualIncome) * 100 : 0;
  const alreadyPaid = (Number(ti.tdsAlreadyPaid) || 0) + (Number(ti.advanceTaxPaid) || 0);
  const remaining = chosen.total - alreadyPaid;

  /* ---------- insurance CRUD ---------- */
  const [showInsForm, setShowInsForm] = useState(false);
  const [editingInsId, setEditingInsId] = useState(null);
  const [insForm, setInsForm] = useState(emptyInsuranceForm());
  const [insError, setInsError] = useState('');
  const [confirmDeleteIns, setConfirmDeleteIns] = useState(null);

  function openAddIns() { setEditingInsId(null); setInsError(''); setInsForm(emptyInsuranceForm()); setShowInsForm(true); }
  function openEditIns(i) { setEditingInsId(i.id); setInsError(''); setInsForm({ type: i.type, provider: i.provider, coverage: String(i.coverage), premium: String(i.premium), renewalDate: i.renewalDate }); setShowInsForm(true); }
  function closeInsForm() { setShowInsForm(false); setEditingInsId(null); setInsError(''); }
  function submitIns(e) {
    e.preventDefault();
    if (!insForm.provider.trim()) { setInsError('Add a provider or policy name.'); return; }
    const coverage = parseFloat(insForm.coverage);
    if (isNaN(coverage) || coverage <= 0) { setInsError('Enter coverage greater than 0.'); return; }
    if (!insForm.renewalDate) { setInsError('Pick a renewal date.'); return; }
    const payload = { type: insForm.type, provider: insForm.provider.trim(), coverage, premium: parseFloat(insForm.premium) || 0, renewalDate: insForm.renewalDate };
    setProtection((prev) => ({ ...prev, insurance: editingInsId ? prev.insurance.map((i) => (i.id === editingInsId ? { ...i, ...payload } : i)) : [...prev.insurance, { id: uid(), ...payload }] }));
    closeInsForm();
  }
  function deleteIns(id) { setProtection((prev) => ({ ...prev, insurance: prev.insurance.filter((i) => i.id !== id) })); setConfirmDeleteIns(null); }

  /* ---------- credit card CRUD ---------- */
  const [showCardForm, setShowCardForm] = useState(false);
  const [editingCardId, setEditingCardId] = useState(null);
  const [cardForm, setCardForm] = useState(emptyCardForm());
  const [cardError, setCardError] = useState('');
  const [confirmDeleteCard, setConfirmDeleteCard] = useState(null);

  function openAddCard() { setEditingCardId(null); setCardError(''); setCardForm(emptyCardForm()); setShowCardForm(true); }
  function openEditCard(c) { setEditingCardId(c.id); setCardError(''); setCardForm({ name: c.name, limit: String(c.limit), outstanding: String(c.outstanding), dueDate: c.dueDate, billingDay: String(c.billingDay || '') }); setShowCardForm(true); }
  function closeCardForm() { setShowCardForm(false); setEditingCardId(null); setCardError(''); }
  function submitCard(e) {
    e.preventDefault();
    if (!cardForm.name.trim()) { setCardError('Name this card.'); return; }
    const limit = parseFloat(cardForm.limit);
    if (isNaN(limit) || limit <= 0) { setCardError('Enter a credit limit greater than 0.'); return; }
    if (!cardForm.dueDate) { setCardError('Pick the next due date.'); return; }
    const payload = { name: cardForm.name.trim(), limit, outstanding: parseFloat(cardForm.outstanding) || 0, dueDate: cardForm.dueDate, billingDay: parseInt(cardForm.billingDay) || 0 };
    setProtection((prev) => ({ ...prev, creditCards: editingCardId ? prev.creditCards.map((c) => (c.id === editingCardId ? { ...c, ...payload } : c)) : [...prev.creditCards, { id: uid(), ...payload }] }));
    closeCardForm();
  }
  function deleteCard(id) { setProtection((prev) => ({ ...prev, creditCards: prev.creditCards.filter((c) => c.id !== id) })); setConfirmDeleteCard(null); }

  const efRatio = Math.min(1, emergencyMonths / Math.max(1, protection.emergencyFundTargetMonths));

  return (
    <>
      <div className="wd-card">
        <h3 className="wd-section-title">Indian tax estimator — FY 2026-27</h3>
        <p className="wd-muted-text">Based on current FY 2026-27 slabs (no change announced in Budget 2026). This is a planning estimate, not tax advice — confirm specifics with a CA before filing.</p>

        <div className="wd-row-3">
          <div className="wd-field"><label>Gross annual income ({CURRENCY})</label><input className="amount-input" type="number" min="0" value={ti.grossAnnualIncome} onChange={(e) => setTaxField('grossAnnualIncome', Number(e.target.value))} /></div>
          <div className="wd-field"><label>Income type</label>
            <select value={ti.isSalaried ? 'salaried' : 'other'} onChange={(e) => setTaxField('isSalaried', e.target.value === 'salaried')}>
              <option value="salaried">Salaried / pension</option>
              <option value="other">Business / other</option>
            </select>
          </div>
          <div className="wd-field"><label>Age band (old regime)</label>
            <select value={ti.ageBand} onChange={(e) => setTaxField('ageBand', e.target.value)}>
              <option value="general">Below 60</option>
              <option value="senior">60 to 79 (senior)</option>
              <option value="superSenior">80+ (super senior)</option>
            </select>
          </div>
        </div>

        <p className="wd-muted-text wd-mt-sm">Old regime deductions (ignored under the new regime):</p>
        <div className="wd-row-3">
          <div className="wd-field"><label>Section 80C (cap ₹1.5L)</label><input className="amount-input" type="number" min="0" value={ti.section80C} onChange={(e) => setTaxField('section80C', Number(e.target.value))} /></div>
          <div className="wd-field"><label>Section 80D health cover</label><input className="amount-input" type="number" min="0" value={ti.section80D} onChange={(e) => setTaxField('section80D', Number(e.target.value))} /></div>
          <div className="wd-field"><label>NPS 80CCD(1B) (cap ₹50k)</label><input className="amount-input" type="number" min="0" value={ti.nps80CCD1B} onChange={(e) => setTaxField('nps80CCD1B', Number(e.target.value))} /></div>
        </div>
        <div className="wd-row-3">
          <div className="wd-field"><label>Home loan interest (cap ₹2L)</label><input className="amount-input" type="number" min="0" value={ti.homeLoanInterest} onChange={(e) => setTaxField('homeLoanInterest', Number(e.target.value))} /></div>
          <div className="wd-field"><label>TDS already deducted</label><input className="amount-input" type="number" min="0" value={ti.tdsAlreadyPaid} onChange={(e) => setTaxField('tdsAlreadyPaid', Number(e.target.value))} /></div>
          <div className="wd-field"><label>Advance tax paid</label><input className="amount-input" type="number" min="0" value={ti.advanceTaxPaid} onChange={(e) => setTaxField('advanceTaxPaid', Number(e.target.value))} /></div>
        </div>

        <div className="wd-tax-compare">
          <div className={`wd-tax-col ${recommended === 'new' ? 'recommended' : ''}`}>
            <span className="wd-tax-col-title">New regime {recommended === 'new' && <span className="wd-rec-badge">Recommended</span>}</span>
            <span className="wd-tax-amount wd-mono">{fmt(newRegime.total)}</span>
            <span className="wd-tax-sub">Taxable income {fmt(newRegime.taxable)}</span>
          </div>
          <div className={`wd-tax-col ${recommended === 'old' ? 'recommended' : ''}`}>
            <span className="wd-tax-col-title">Old regime {recommended === 'old' && <span className="wd-rec-badge">Recommended</span>}</span>
            <span className="wd-tax-amount wd-mono">{fmt(oldRegime.total)}</span>
            <span className="wd-tax-sub">Taxable income {fmt(oldRegime.taxable)}</span>
          </div>
        </div>

        <div className="wd-summary-grid wd-summary-grid-4 wd-mt">
          <div className="wd-summary-card"><span className="sc-label">Estimated tax</span><span className="sc-value wd-mono">{fmt(chosen.total)}</span></div>
          <div className="wd-summary-card net"><span className="sc-label">Effective rate</span><span className="sc-value wd-mono">{effectiveRate.toFixed(1)}%</span></div>
          <div className="wd-summary-card in"><span className="sc-label">Saved by best regime</span><span className="sc-value wd-mono">{fmt(taxSavingsByChoosing)}</span></div>
          <div className={`wd-summary-card ${remaining > 0 ? 'out' : 'in'}`}><span className="sc-label">{remaining > 0 ? 'Remaining liability' : 'Refund due'}</span><span className="sc-value wd-mono">{fmt(Math.abs(remaining))}</span></div>
        </div>
      </div>

      <div className="wd-card">
        <h3 className="wd-section-title">Emergency fund</h3>
        <div className="wd-row-2">
          <div className="wd-field"><label>Target (months of essential expenses)</label>
            <select value={protection.emergencyFundTargetMonths} onChange={(e) => setProtection((prev) => ({ ...prev, emergencyFundTargetMonths: Number(e.target.value) }))}>
              <option value={3}>3 months</option>
              <option value={6}>6 months</option>
              <option value={12}>12 months</option>
            </select>
          </div>
          <div className="wd-field"><label>Liquid cash available</label><div className="wd-static-value wd-mono">{fmt(netWorthTotals.liquidCash)}</div></div>
        </div>
        <div className="wd-bar-track wd-ef-track"><div className="wd-bar-fill" style={{ width: `${efRatio * 100}%`, background: efRatio >= 1 ? 'var(--sage)' : efRatio >= 0.5 ? 'var(--amber)' : 'var(--clay)' }} /></div>
        <div className="wd-goal-stats wd-mono"><span>{emergencyMonths.toFixed(1)} months covered</span><span>Target: {protection.emergencyFundTargetMonths} months</span></div>
      </div>

      <div className="wd-card">
        <h3 className="wd-section-title">Debt overview</h3>
        <div className="wd-summary-grid wd-summary-grid-4">
          <div className="wd-summary-card"><span className="sc-label">Debt-to-income</span><span className="sc-value wd-mono">{debtToIncomePct.toFixed(1)}%</span></div>
          <div className="wd-summary-card"><span className="sc-label">Debt-to-asset</span><span className="sc-value wd-mono">{debtToAssetPct.toFixed(1)}%</span></div>
          <div className="wd-summary-card out"><span className="sc-label">Total liabilities</span><span className="sc-value wd-mono">{fmtCompact(netWorthTotals.totalLiabilities)}</span></div>
          <div className="wd-summary-card"><span className="sc-label">Open loans</span><span className="sc-value wd-mono">{netWorth.liabilities.length}</span></div>
        </div>
        {netWorth.liabilities.length > 0 && (
          <div className="wd-ledger wd-mt">
            {netWorth.liabilities.map((l) => (
              <div className="wd-simple-row" key={l.id}>
                <span className="wd-row-dot" style={{ background: 'var(--clay)' }} />
                <div className="wd-entry-main"><span className="wd-entry-category">{l.name}</span><span className="wd-entry-sub">{l.category}{l.emi ? ` · EMI ${fmt(l.emi)}` : ''}{l.remainingTenureMonths ? ` · ${l.remainingTenureMonths} mo left` : ''}</span></div>
                <span className="wd-entry-amount wd-mono" style={{ color: 'var(--clay)' }}>{fmt(l.amount)}</span>
              </div>
            ))}
          </div>
        )}
        <p className="wd-muted-text wd-mt-sm">Manage or edit loans in the Net Worth tab — this view is read-only here.</p>
      </div>

      <div className="wd-card">
        <div className="wd-section-header">
          <h3 className="wd-section-title">Insurance</h3>
          <button className="wd-btn-primary" onClick={() => (showInsForm && !editingInsId ? closeInsForm() : openAddIns())}><Plus size={15} /> Add policy</button>
        </div>
        {showInsForm && (
          <form className="wd-form" onSubmit={submitIns}>
            <div className="wd-row-2">
              <div className="wd-field"><label>Type</label><select value={insForm.type} onChange={(e) => setInsForm((f) => ({ ...f, type: e.target.value }))}>{INSURANCE_TYPES.map((t) => <option value={t} key={t}>{t}</option>)}</select></div>
              <div className="wd-field"><label>Provider / policy</label><input type="text" placeholder="e.g. Star Health" value={insForm.provider} onChange={(e) => setInsForm((f) => ({ ...f, provider: e.target.value }))} /></div>
            </div>
            <div className="wd-row-3">
              <div className="wd-field"><label>Coverage ({CURRENCY})</label><input className="amount-input" type="number" min="0" value={insForm.coverage} onChange={(e) => setInsForm((f) => ({ ...f, coverage: e.target.value }))} /></div>
              <div className="wd-field"><label>Annual premium ({CURRENCY})</label><input className="amount-input" type="number" min="0" value={insForm.premium} onChange={(e) => setInsForm((f) => ({ ...f, premium: e.target.value }))} /></div>
              <div className="wd-field"><label>Renewal date</label><input type="date" value={insForm.renewalDate} onChange={(e) => setInsForm((f) => ({ ...f, renewalDate: e.target.value }))} /></div>
            </div>
            {insError && <div className="wd-form-error">{insError}</div>}
            <div className="wd-form-actions"><button type="button" className="wd-btn-ghost" onClick={closeInsForm}>Cancel</button><button type="submit" className="wd-btn-primary">{editingInsId ? 'Save changes' : 'Add policy'}</button></div>
          </form>
        )}
        {protection.insurance.length === 0 ? <div className="wd-empty">No policies added yet.</div> : (
          <div className="wd-ledger">
            {protection.insurance.map((i) => {
              const dleft = daysUntil(i.renewalDate);
              return (
                <div className="wd-simple-row" key={i.id}>
                  <span className="wd-row-dot" style={{ background: categoryColor(i.type) }} />
                  <div className="wd-entry-main">
                    <span className="wd-entry-category">{i.provider} <span className="wd-pill">{i.type}</span></span>
                    <span className={`wd-entry-sub ${dleft <= 30 ? 'clay' : ''}`}>Coverage {fmt(i.coverage)} · Premium {fmt(i.premium)}/yr · {dleft <= 30 ? `Renews in ${dleft}d` : `Renews ${formatShortDate(i.renewalDate)}`}</span>
                  </div>
                  <div className="wd-entry-actions">
                    {confirmDeleteIns === i.id ? (
                      <div className="wd-confirm"><button className="yes" onClick={() => deleteIns(i.id)}>Delete</button><button onClick={() => setConfirmDeleteIns(null)}>Cancel</button></div>
                    ) : (<><button className="wd-icon-btn" onClick={() => openEditIns(i)}><Pencil size={14} /></button><button className="wd-icon-btn" onClick={() => setConfirmDeleteIns(i.id)}><Trash2 size={14} /></button></>)}
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>

      <div className="wd-card">
        <div className="wd-section-header">
          <h3 className="wd-section-title">Credit cards</h3>
          <button className="wd-btn-primary" onClick={() => (showCardForm && !editingCardId ? closeCardForm() : openAddCard())}><Plus size={15} /> Add card</button>
        </div>
        {showCardForm && (
          <form className="wd-form" onSubmit={submitCard}>
            <div className="wd-field"><label>Card name</label><input type="text" placeholder="e.g. HDFC Regalia" value={cardForm.name} onChange={(e) => setCardForm((f) => ({ ...f, name: e.target.value }))} /></div>
            <div className="wd-row-3">
              <div className="wd-field"><label>Credit limit ({CURRENCY})</label><input className="amount-input" type="number" min="0" value={cardForm.limit} onChange={(e) => setCardForm((f) => ({ ...f, limit: e.target.value }))} /></div>
              <div className="wd-field"><label>Outstanding ({CURRENCY})</label><input className="amount-input" type="number" min="0" value={cardForm.outstanding} onChange={(e) => setCardForm((f) => ({ ...f, outstanding: e.target.value }))} /></div>
              <div className="wd-field"><label>Next due date</label><input type="date" value={cardForm.dueDate} onChange={(e) => setCardForm((f) => ({ ...f, dueDate: e.target.value }))} /></div>
            </div>
            {cardError && <div className="wd-form-error">{cardError}</div>}
            <div className="wd-form-actions"><button type="button" className="wd-btn-ghost" onClick={closeCardForm}>Cancel</button><button type="submit" className="wd-btn-primary">{editingCardId ? 'Save changes' : 'Add card'}</button></div>
          </form>
        )}
        {protection.creditCards.length === 0 ? <div className="wd-empty">No cards added yet.</div> : (
          <div className="wd-ledger">
            {protection.creditCards.map((c) => {
              const util = c.limit > 0 ? (c.outstanding / c.limit) * 100 : 0;
              const dleft = daysUntil(c.dueDate);
              return (
                <div className="wd-simple-row" key={c.id}>
                  <span className="wd-row-dot" style={{ background: util > 50 ? 'var(--clay)' : util > 30 ? 'var(--amber)' : 'var(--sage)' }} />
                  <div className="wd-entry-main">
                    <span className="wd-entry-category">{c.name}</span>
                    <span className={`wd-entry-sub ${dleft <= 7 ? 'clay' : ''}`}>{util.toFixed(0)}% utilized · Limit {fmt(c.limit)} · {dleft <= 7 ? `Due in ${dleft}d` : `Due ${formatShortDate(c.dueDate)}`}</span>
                  </div>
                  <span className="wd-entry-amount wd-mono">{fmt(c.outstanding)}</span>
                  <div className="wd-entry-actions">
                    {confirmDeleteCard === c.id ? (
                      <div className="wd-confirm"><button className="yes" onClick={() => deleteCard(c.id)}>Delete</button><button onClick={() => setConfirmDeleteCard(null)}>Cancel</button></div>
                    ) : (<><button className="wd-icon-btn" onClick={() => openEditCard(c)}><Pencil size={14} /></button><button className="wd-icon-btn" onClick={() => setConfirmDeleteCard(c.id)}><Trash2 size={14} /></button></>)}
                  </div>
                </div>
              );
            })}
          </div>
        )}
      </div>
    </>
  );
}

/* ============================== STYLES ============================== */

const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,500;9..144,600&family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Mono:wght@400;500;600&display=swap');

.wd-wrap {
  --paper: #F7F7F2;
  --ink: #232925;
  --ink-soft: #6B7570;
  --line: #DEE0D6;
  --card: #FFFFFF;
  --sage: #5B8270;
  --sage-soft: #E4EBE3;
  --clay: #B6694C;
  --clay-soft: #F2E3DA;
  --slate: #51687D;
  --slate-soft: #E4E9ED;
  --amber: #C99A3E;
  --font-display: 'Fraunces', Georgia, serif;
  --font-sans: 'IBM Plex Sans', -apple-system, BlinkMacSystemFont, sans-serif;
  --font-mono: 'IBM Plex Mono', 'SFMono-Regular', Menlo, monospace;
  background: var(--paper);
  color: var(--ink);
  font-family: var(--font-sans);
  min-height: 100%;
  padding: 20px 16px 40px;
  box-sizing: border-box;
}
.wd-wrap *, .wd-wrap *::before, .wd-wrap *::after { box-sizing: border-box; }
.wd-mono { font-family: var(--font-mono); font-variant-numeric: tabular-nums; }
.wd-container { max-width: 760px; margin: 0 auto; display: flex; flex-direction: column; gap: 16px; }
.wd-loading { display: flex; align-items: center; justify-content: center; min-height: 300px; color: var(--ink-soft); font-size: 14px; }
.clay { color: var(--clay) !important; }
.sage { color: var(--sage) !important; }
.slate { color: var(--slate) !important; }

/* header */
.wd-header { background: var(--card); border: 1px solid var(--line); border-radius: 14px; padding: 22px 22px 18px; text-align: center; }
.wd-brand { display: flex; align-items: center; justify-content: center; gap: 6px; font-size: 12.5px; letter-spacing: 0.14em; text-transform: uppercase; color: var(--ink-soft); margin-bottom: 12px; }
.wd-balance-label { display: block; font-size: 13px; color: var(--ink-soft); margin-bottom: 4px; }
.wd-balance-amount { display: block; font-family: var(--font-display); font-size: clamp(28px, 7vw, 38px); font-weight: 500; letter-spacing: -0.01em; line-height: 1.1; }
.wd-balance-amount.negative { color: var(--clay); }
.wd-balance-sub { display: block; margin-top: 6px; font-size: 12px; color: var(--ink-soft); }
.wd-metric-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(120px, 1fr)); gap: 8px; margin-top: 16px; }
.wd-metric { background: var(--paper); border: 1px solid var(--line); border-radius: 10px; padding: 9px 8px; text-align: center; }
.wd-metric .m-label { display: block; font-size: 10.5px; color: var(--ink-soft); text-transform: uppercase; letter-spacing: .04em; margin-bottom: 4px; }
.wd-metric .m-value { display: block; font-size: clamp(12px, 3.4vw, 14.5px); font-weight: 500; }

.wd-notice { background: var(--clay-soft); color: var(--clay); border: 1px solid var(--clay); border-radius: 10px; padding: 9px 12px; font-size: 12.5px; text-align: center; }

/* nav */
.wd-navtabs { display: flex; gap: 4px; background: var(--card); border: 1px solid var(--line); border-radius: 12px; padding: 4px; overflow-x: auto; }
.wd-navtab { flex: 1; display: flex; align-items: center; justify-content: center; gap: 5px; padding: 9px 6px; border-radius: 9px; border: none; background: transparent; font-size: 12.5px; font-weight: 500; color: var(--ink-soft); cursor: pointer; font-family: var(--font-sans); white-space: nowrap; transition: background .15s, color .15s; }
.wd-navtab.active { background: var(--slate-soft); color: var(--slate); }

/* period switcher */
.wd-tabs { display: flex; background: var(--card); border: 1px solid var(--line); border-radius: 12px; padding: 4px; gap: 4px; overflow-x: auto; }
.wd-tab { flex: 1; padding: 8px 0; text-align: center; border-radius: 9px; border: none; background: transparent; font-size: 13px; font-weight: 500; color: var(--ink-soft); cursor: pointer; font-family: var(--font-sans); white-space: nowrap; }
.wd-tab.active { background: var(--slate-soft); color: var(--slate); }
.wd-period-nav { display: flex; align-items: center; justify-content: space-between; gap: 8px; }
.wd-nav-btn { width: 34px; height: 34px; border-radius: 9px; border: 1px solid var(--line); background: var(--card); display: flex; align-items: center; justify-content: center; cursor: pointer; color: var(--ink); flex-shrink: 0; }
.wd-nav-btn:hover { background: var(--paper); }
.wd-period-center { flex: 1; text-align: center; display: flex; flex-direction: column; align-items: center; gap: 3px; min-width: 0; }
.wd-period-label { font-family: var(--font-display); font-size: 16px; }
.wd-jump-today { border: none; background: none; color: var(--slate); font-size: 12px; cursor: pointer; text-decoration: underline; padding: 0; font-family: var(--font-sans); }

/* summary cards */
.wd-summary-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 10px; }
.wd-summary-grid-4 { grid-template-columns: repeat(4, 1fr); }
.wd-summary-card { background: var(--card); border: 1px solid var(--line); border-radius: 12px; padding: 12px 8px; text-align: center; }
.wd-summary-card .sc-label { display: block; font-size: 10.5px; color: var(--ink-soft); margin-bottom: 6px; text-transform: uppercase; letter-spacing: .05em; }
.wd-summary-card .sc-value { display: block; font-size: clamp(12.5px, 3.6vw, 16px); font-weight: 500; }
.wd-summary-card.in .sc-value { color: var(--sage); }
.wd-summary-card.out .sc-value { color: var(--clay); }
.wd-summary-card.net .sc-value { color: var(--slate); }
.wd-mt { margin-top: 12px; }
.wd-mt-sm { margin-top: 6px; }

/* generic card */
.wd-card { background: var(--card); border: 1px solid var(--line); border-radius: 14px; padding: 18px 18px 20px; }
.wd-section-title { font-family: var(--font-display); font-size: 16px; margin: 0 0 10px; font-weight: 500; }
.wd-section-header { display: flex; align-items: center; justify-content: space-between; gap: 10px; margin-bottom: 6px; flex-wrap: wrap; }
.wd-section-header .wd-section-title { margin-bottom: 0; }
.wd-header-actions { display: flex; gap: 8px; flex-wrap: wrap; }
.wd-muted-text { font-size: 12.5px; color: var(--ink-soft); line-height: 1.5; margin: 0 0 12px; }
.wd-grid-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
.wd-capital-note { display: flex; align-items: flex-start; gap: 7px; background: var(--slate-soft); color: var(--slate); border-radius: 10px; padding: 9px 12px; font-size: 12px; line-height: 1.5; }
.wd-coast-note { display: flex; align-items: flex-start; gap: 7px; background: var(--paper); border: 1px solid var(--line); border-radius: 10px; padding: 10px 12px; font-size: 12.5px; line-height: 1.6; color: var(--ink-soft); margin-top: 12px; }

/* legend */
.wd-legend { display: flex; flex-wrap: wrap; gap: 8px 14px; margin-top: 10px; justify-content: center; }
.wd-legend-item { display: flex; align-items: center; gap: 5px; font-size: 11.5px; color: var(--ink-soft); }
.wd-legend-item .dot { width: 8px; height: 8px; border-radius: 50%; flex-shrink: 0; }

/* chart tooltip */
.wd-chart-tooltip { background: var(--card); border: 1px solid var(--line); border-radius: 8px; padding: 8px 10px; font-size: 12px; box-shadow: 0 4px 14px rgba(0,0,0,0.08); }
.tt-label { color: var(--ink-soft); margin-bottom: 4px; font-size: 11px; }
.tt-row { display: flex; align-items: center; gap: 6px; color: var(--ink); }
.tt-dot { width: 7px; height: 7px; border-radius: 50%; display: inline-block; }

/* wealth score */
.wd-score-head { display: flex; align-items: flex-start; justify-content: space-between; gap: 12px; margin-bottom: 14px; }
.wd-score-badge { display: flex; align-items: baseline; gap: 2px; font-family: var(--font-display); font-size: 28px; color: var(--slate); flex-shrink: 0; }
.wd-score-max { font-size: 14px; color: var(--ink-soft); }
.wd-score-row { display: grid; grid-template-columns: 110px 1fr 56px; align-items: center; gap: 10px; padding: 6px 0; }
.wd-score-label { font-size: 12.5px; color: var(--ink); }
.wd-score-track { background: var(--paper); border-radius: 6px; height: 8px; overflow: hidden; }
.wd-score-fill { height: 100%; background: var(--slate); border-radius: 6px; }
.wd-score-value { font-size: 12px; color: var(--ink-soft); text-align: right; }
.wd-score-tips { margin-top: 12px; display: flex; flex-direction: column; gap: 6px; }
.wd-tip { display: flex; align-items: flex-start; gap: 7px; font-size: 12.5px; color: var(--ink-soft); line-height: 1.5; }
.wd-tip svg { flex-shrink: 0; margin-top: 2px; color: var(--amber); }

/* bar rows (category) */
.wd-bar-row { display: grid; grid-template-columns: 92px 1fr 76px; align-items: center; gap: 10px; padding: 7px 0; }
.wd-bar-label { font-size: 13px; color: var(--ink); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.wd-bar-track { background: var(--paper); border-radius: 6px; height: 10px; overflow: hidden; }
.wd-bar-fill { height: 100%; border-radius: 6px; transition: width .3s ease; }
.wd-bar-amount { font-size: 12.5px; color: var(--ink-soft); text-align: right; }
.wd-ef-track { height: 12px; margin-bottom: 8px; }
.wd-goal-track { margin: 8px 0; }

.wd-empty { text-align: center; padding: 18px 10px; color: var(--ink-soft); font-size: 13.5px; }

/* buttons */
.wd-btn-primary { display: inline-flex; align-items: center; gap: 6px; background: var(--ink); color: var(--paper); border: none; border-radius: 9px; padding: 9px 14px; font-size: 13px; font-weight: 500; cursor: pointer; font-family: var(--font-sans); white-space: nowrap; }
.wd-btn-primary:hover { opacity: .88; }
.wd-btn-ghost { display: inline-flex; align-items: center; gap: 6px; background: transparent; border: 1px solid var(--line); border-radius: 9px; padding: 9px 14px; font-size: 13px; cursor: pointer; color: var(--ink); font-family: var(--font-sans); white-space: nowrap; }
.wd-btn-ghost:hover { background: var(--paper); }
.wd-link-btn { display: inline-flex; align-items: center; gap: 6px; background: none; border: none; color: var(--slate); font-size: 12.5px; cursor: pointer; text-decoration: underline; font-family: var(--font-sans); padding: 0; }

/* forms */
@keyframes wd-fade-in { from { opacity: 0; transform: translateY(-4px); } to { opacity: 1; transform: translateY(0); } }
.wd-form { animation: wd-fade-in .18s ease; margin: 12px 0 16px; padding: 16px; background: var(--paper); border: 1px solid var(--line); border-radius: 12px; display: flex; flex-direction: column; gap: 12px; }
.wd-type-toggle { display: flex; gap: 6px; }
.wd-type-toggle-wrap { flex-wrap: wrap; }
.wd-type-btn { flex: 1; min-width: 90px; padding: 8px 6px; text-align: center; border-radius: 8px; border: 1px solid var(--line); background: var(--card); cursor: pointer; font-size: 12.5px; font-weight: 500; color: var(--ink-soft); font-family: var(--font-sans); }
.wd-field { display: flex; flex-direction: column; gap: 5px; }
.wd-field label { font-size: 12px; color: var(--ink-soft); }
.wd-field input, .wd-field select { border: 1px solid var(--line); border-radius: 8px; padding: 9px 10px; font-size: 14px; font-family: var(--font-sans); background: var(--card); color: var(--ink); width: 100%; }
.wd-field input.amount-input { font-family: var(--font-mono); }
.wd-row-2 { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.wd-row-3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 10px; }
.wd-form-error { color: var(--clay); font-size: 12.5px; }
.wd-form-actions { display: flex; gap: 8px; justify-content: flex-end; }
.wd-filter-row { display: flex; gap: 8px; margin-bottom: 10px; flex-wrap: wrap; }
.wd-search { flex: 1; min-width: 160px; border: 1px solid var(--line); border-radius: 8px; padding: 9px 10px; font-size: 13.5px; font-family: var(--font-sans); background: var(--paper); color: var(--ink); }
.wd-type-filter { border: 1px solid var(--line); border-radius: 8px; padding: 9px 10px; font-size: 13px; font-family: var(--font-sans); background: var(--paper); color: var(--ink); }
.wd-checkbox { display: flex; align-items: center; gap: 6px; font-size: 13px; color: var(--ink); cursor: pointer; }
.wd-recurring-row { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
.wd-recurring-row select { border: 1px solid var(--line); border-radius: 8px; padding: 7px 9px; font-size: 13px; font-family: var(--font-sans); background: var(--card); }
.wd-sip-input { max-width: 180px; }
.wd-static-value { padding: 9px 10px; background: var(--card); border: 1px solid var(--line); border-radius: 8px; font-size: 14px; }

/* ledger rows */
.wd-ledger { display: flex; flex-direction: column; }
.wd-tx-row, .wd-simple-row { display: grid; grid-template-columns: 4px 52px 1fr auto auto; align-items: center; gap: 10px; padding: 11px 0; border-bottom: 1px solid var(--line); }
.wd-simple-row { grid-template-columns: 4px 1fr auto auto; }
.wd-tx-row:last-child, .wd-simple-row:last-child { border-bottom: none; }
.wd-entry-spine, .wd-row-dot { width: 4px; height: 28px; border-radius: 3px; }
.wd-row-dot { width: 8px; height: 8px; border-radius: 50%; align-self: center; }
.wd-entry-date { font-size: 12px; color: var(--ink-soft); }
.wd-entry-main { display: flex; flex-direction: column; gap: 2px; min-width: 0; }
.wd-entry-category { font-size: 14px; font-weight: 500; display: flex; align-items: center; gap: 5px; }
.wd-entry-sub { font-size: 12px; color: var(--ink-soft); overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.wd-recur-icon { color: var(--ink-soft); flex-shrink: 0; }
.wd-tag-row { display: flex; gap: 4px; flex-wrap: wrap; margin-top: 2px; }
.wd-tag { font-size: 10.5px; background: var(--slate-soft); color: var(--slate); padding: 1px 7px; border-radius: 20px; }
.wd-pill { font-size: 10.5px; background: var(--paper); border: 1px solid var(--line); color: var(--ink-soft); padding: 1px 7px; border-radius: 20px; margin-left: 4px; }
.wd-entry-amount { font-size: 14px; font-weight: 500; text-align: right; white-space: nowrap; }
.wd-inv-values { display: flex; flex-direction: column; align-items: flex-end; gap: 1px; }
.wd-inv-gain { font-size: 11.5px; }
.wd-entry-actions { display: flex; gap: 2px; }
.wd-icon-btn { width: 28px; height: 28px; border-radius: 7px; border: none; background: transparent; color: var(--ink-soft); display: flex; align-items: center; justify-content: center; cursor: pointer; }
.wd-icon-btn:hover { background: var(--paper); color: var(--ink); }
.wd-confirm { display: flex; align-items: center; gap: 6px; }
.wd-confirm-label { font-size: 12px; color: var(--ink-soft); margin-right: 4px; }
.wd-confirm button { font-size: 11.5px; padding: 4px 8px; border-radius: 6px; border: 1px solid var(--line); background: var(--card); cursor: pointer; font-family: var(--font-sans); color: var(--ink); }
.wd-confirm .yes { color: var(--clay); border-color: var(--clay); }

/* goals */
.wd-goal-list { display: flex; flex-direction: column; gap: 14px; }
.wd-goal-card { border: 1px solid var(--line); border-radius: 12px; padding: 13px 14px; }
.wd-goal-top { display: flex; align-items: flex-start; justify-content: space-between; gap: 10px; margin-bottom: 8px; }
.wd-goal-name { font-size: 14px; font-weight: 500; display: block; }
.wd-goal-type { font-size: 11.5px; color: var(--ink-soft); display: block; margin-top: 1px; }
.wd-goal-stats { display: flex; justify-content: space-between; font-size: 11.5px; color: var(--ink-soft); flex-wrap: wrap; gap: 4px; }

/* tax */
.wd-tax-compare { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; margin-top: 6px; }
.wd-tax-col { border: 1px solid var(--line); border-radius: 12px; padding: 14px; display: flex; flex-direction: column; gap: 4px; text-align: center; }
.wd-tax-col.recommended { border-color: var(--sage); background: var(--sage-soft); }
.wd-tax-col-title { font-size: 12.5px; color: var(--ink-soft); display: flex; align-items: center; justify-content: center; gap: 6px; flex-wrap: wrap; }
.wd-rec-badge { font-size: 10px; background: var(--sage); color: var(--card); padding: 2px 7px; border-radius: 20px; }
.wd-tax-amount { font-size: 19px; font-weight: 500; }
.wd-tax-sub { font-size: 11px; color: var(--ink-soft); }

/* footer */
.wd-footer { text-align: center; padding-top: 4px; display: flex; flex-direction: column; gap: 8px; }
.wd-footer-links { display: flex; align-items: center; justify-content: center; gap: 16px; flex-wrap: wrap; }
.wd-reset-link { background: none; border: none; color: var(--ink-soft); font-size: 12px; text-decoration: underline; cursor: pointer; font-family: var(--font-sans); }
.wd-disclaimer { font-size: 11px; color: var(--ink-soft); max-width: 520px; margin: 0 auto; line-height: 1.5; }

.wd-wrap *:focus-visible { outline: 2px solid var(--slate); outline-offset: 2px; }

@media (prefers-reduced-motion: reduce) {
  .wd-wrap *, .wd-wrap *::before, .wd-wrap *::after { animation: none !important; transition: none !important; }
}

@media (max-width: 560px) {
  .wd-grid-2 { grid-template-columns: 1fr; }
  .wd-row-3 { grid-template-columns: 1fr; }
  .wd-tax-compare { grid-template-columns: 1fr; }
}

@media (max-width: 380px) {
  .wd-wrap { padding: 14px 10px 32px; }
  .wd-header { padding: 16px 12px 12px; }
  .wd-card { padding: 14px 14px 16px; }
  .wd-tx-row, .wd-simple-row { grid-template-columns: 4px 44px 1fr auto auto; gap: 7px; }
  .wd-simple-row { grid-template-columns: 4px 1fr auto auto; }
  .wd-bar-row { grid-template-columns: 68px 1fr 64px; gap: 7px; }
  .wd-row-2 { grid-template-columns: 1fr; }
  .wd-summary-grid-4 { grid-template-columns: repeat(2, 1fr); }
  .wd-score-row { grid-template-columns: 88px 1fr 48px; }
}
`;
