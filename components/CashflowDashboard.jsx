'use client';
import React, { useState, useEffect, useMemo, useRef } from 'react';
import {
  Wallet,
  TrendingUp,
  TrendingDown,
  PiggyBank,
  CreditCard,
  Plus,
  RefreshCw,
  ArrowUpRight,
  ArrowDownRight,
  Sun,
  Moon,
  Home,
  List,
  LineChart,
  Landmark,
  Trash2,
  X,
  CheckCircle2,
  Coins,
  Briefcase,
  ShieldCheck,
  Sparkles,
  ChevronRight,
  Play,
  Pause,
  RotateCcw,
  AlertCircle,
  Globe,
  FileSpreadsheet,
  Image as ImageIcon,
  Cloud,
  Download,
  Copy,
  Check,
  Share2,
  Smartphone,
  LogIn,
  LogOut
} from 'lucide-react';
import {
  AreaChart,
  Area,
  XAxis,
  YAxis,
  Tooltip,
  ResponsiveContainer
} from 'recharts';
import { initializeApp, getApps, getApp } from 'firebase/app';
import {
  getAuth,
  signInAnonymously,
  onAuthStateChanged,
  GoogleAuthProvider,
  signInWithPopup,
  signOut
} from 'firebase/auth';
import { getFirestore, doc, setDoc, onSnapshot } from 'firebase/firestore';

// --- INISIALISASI FIREBASE CLOUD (MULTI-DEVICE SYNC NEXT.JS & VERCEL) ---
// Config web Firebase bersifat publik (keamanan ada di Firestore Rules), jadi aman dijadikan fallback.
const firebaseConfig = {
  apiKey: process.env.NEXT_PUBLIC_FIREBASE_API_KEY || 'AIzaSyD34TR6qrkTe_a42SXLE-6BlDom39iUpUc',
  authDomain: process.env.NEXT_PUBLIC_FIREBASE_AUTH_DOMAIN || 'live-pocket-rj.firebaseapp.com',
  projectId: process.env.NEXT_PUBLIC_FIREBASE_PROJECT_ID || 'live-pocket-rj',
  storageBucket: process.env.NEXT_PUBLIC_FIREBASE_STORAGE_BUCKET || 'live-pocket-rj.firebasestorage.app',
  messagingSenderId: process.env.NEXT_PUBLIC_FIREBASE_MESSAGING_SENDER_ID || '1094369218496',
  appId: process.env.NEXT_PUBLIC_FIREBASE_APP_ID || '1:1094369218496:web:e222f2c375fe0b801a4225',
  measurementId: process.env.NEXT_PUBLIC_FIREBASE_MEASUREMENT_ID || 'G-X52P6YEJL5'
};

const app = getApps().length > 0 ? getApp() : initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
const appId = process.env.NEXT_PUBLIC_APP_ID || 'dompetku-ririn-jeki';

// --- KODE GOOGLE APPS SCRIPT SIAP PAKAI UNTUK USER ---
const GOOGLE_APPS_SCRIPT_TEMPLATE = `function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var payload = JSON.parse(e.postData.contents);
    
    // 1. Tab Transaksi & Running Balance
    var txSheet = ss.getSheetByName('Transaksi') || ss.insertSheet('Transaksi');
    if (txSheet.getLastRow() === 0) {
      txSheet.appendRow(['ID', 'Tanggal', 'User', 'Tipe', 'Kategori', 'Judul', 'Nominal (IDR)', 'Running Balance (IDR)', 'Catatan']);
      txSheet.getRange('A1:I1').setFontWeight('bold').setBackground('#10B981').setFontColor('#FFFFFF');
    }
    if (payload.mode === 'single_tx' && payload.transaction) {
      var t = payload.transaction;
      txSheet.appendRow([t.id, t.date, t.user || 'Jeki', t.type, t.category, t.title, t.amount, t.runningBalance || '-', t.note || '-']);
    } else if (payload.mode === 'full_sync') {
      txSheet.clear();
      txSheet.appendRow(['ID', 'Tanggal', 'User', 'Tipe', 'Kategori', 'Judul', 'Nominal (IDR)', 'Running Balance (IDR)', 'Catatan']);
      txSheet.getRange('A1:I1').setFontWeight('bold').setBackground('#10B981').setFontColor('#FFFFFF');
      (payload.transactions || []).forEach(function(t) {
        txSheet.appendRow([t.id, t.date, t.user || 'Jeki', t.type, t.category, t.title, t.amount, t.runningBalance || '-', t.note || '-']);
      });
      
      // 2. Tab Portofolio Investasi Live
      var pfSheet = ss.getSheetByName('Portofolio') || ss.insertSheet('Portofolio');
      pfSheet.clear();
      pfSheet.appendRow(['Kode', 'Nama Aset', 'Kategori', 'Kepemilikan', 'Satuan', 'Harga Live (IDR)', 'Nilai Total (IDR)']);
      pfSheet.getRange('A1:G1').setFontWeight('bold').setBackground('#F59E0B').setFontColor('#000000');
      (payload.assets || []).forEach(function(a) {
        var totalVal = a.holdings * a.multiplier * a.price;
        pfSheet.appendRow([a.symbol, a.name, a.type, a.holdings, a.unitLabel, a.price, totalVal]);
      });

      // 3. Tab Hutang & Cicilan
      var dbSheet = ss.getSheetByName('Hutang') || ss.insertSheet('Hutang');
      dbSheet.clear();
      dbSheet.appendRow(['Nama Pinjaman', 'User', 'Kategori', 'Plafon Awal', 'Sisa Hutang', 'Cicilan/Bulan', 'Jatuh Tempo']);
      dbSheet.getRange('A1:G1').setFontWeight('bold').setBackground('#8B5CF6').setFontColor('#FFFFFF');
      (payload.debts || []).forEach(function(d) {
        dbSheet.appendRow([d.name, d.user || 'Bersama', d.category, d.totalAmount, d.remainingAmount, d.monthlyInstallment, d.dueDate]);
      });
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'success', timestamp: new Date() }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}`;

// --- KONSTANTA & DATA REAL PORTFOLIOMU (DALAM SATUAN IDR & RIBUAN) ---
const INITIAL_USER_OPENING = {
  Jeki: 2875000, // 2,875 (dalam ribu)
  Ririn: 20376410 // 20,376.410 (dalam ribu)
};

const INITIAL_MARKET_ASSETS = [
  // 1. KATEGORI EMAS LOGAM MULIA & PERHIASAN (DATA REAL)
  {
    id: 'gold-antam',
    symbol: 'ANTM',
    name: 'LM Antam (10g + 5g + 2g)',
    type: 'Emas',
    unitLabel: 'gram',
    multiplier: 1,
    price: 2495000,
    prevPrice: 2489000,
    volatility: 0.0022,
    holdings: 17.0, // 10.0 (19-Mar-26) + 5.0 (22-May-26) + 2.0 (11-Nov-26)
    avgBuyPrice: 42735000 / 17.0, // Total modal 28,980.0 + 13,755.0 + 0 = 42,735.0 rb
    color: '#F59E0B'
  },
  {
    id: 'gold-semar',
    symbol: 'SMR',
    name: 'LM Semar Nusantara',
    type: 'Emas',
    unitLabel: 'gram',
    multiplier: 1,
    price: 2236000,
    prevPrice: 2231000,
    volatility: 0.0023,
    holdings: 3.0, // 11-Jun-2025: 3.0 gr -> 5,628.0 rb
    avgBuyPrice: 5628000 / 3.0, // 1,876.0 rb/gr
    color: '#D97706'
  },
  {
    id: 'gold-galeri24',
    symbol: 'G24',
    name: 'LM Galeri 24 Pegadaian',
    type: 'Emas',
    unitLabel: 'gram',
    multiplier: 1,
    price: 2334000,
    prevPrice: 2328000,
    volatility: 0.0022,
    holdings: 0,
    avgBuyPrice: 2334000,
    color: '#B45309'
  },
  {
    id: 'gold-ubs',
    symbol: 'UBS',
    name: 'LM UBS Gold',
    type: 'Emas',
    unitLabel: 'gram',
    multiplier: 1,
    price: 2357000,
    prevPrice: 2352000,
    volatility: 0.0022,
    holdings: 0,
    avgBuyPrice: 2357000,
    color: '#EAB308'
  },
  {
    id: 'gold-cincin',
    symbol: 'PRH-CNC',
    name: 'Perhiasan - Cincin',
    type: 'Emas',
    unitLabel: 'gram',
    multiplier: 1,
    price: 1850000,
    prevPrice: 1846000,
    volatility: 0.0018,
    holdings: 2.0, // 11-Jun-2025: 2.0 gr -> 3,336.0 rb
    avgBuyPrice: 3336000 / 2.0, // 1,668.0 rb/gr
    color: '#F472B6'
  },
  {
    id: 'gold-anting',
    symbol: 'PRH-ANT',
    name: 'Perhiasan - Anting Jepit',
    type: 'Emas',
    unitLabel: 'gram',
    multiplier: 1,
    price: 1750000,
    prevPrice: 1747000,
    volatility: 0.0018,
    holdings: 1.6, // 11-Jun-2025: 1.6 gr -> 2,452.8 rb
    avgBuyPrice: 2452800 / 1.6, // 1,533.0 rb/gr
    color: '#EC4899'
  },
  {
    id: 'gold-gelang',
    symbol: 'PRH-GLG',
    name: 'Perhiasan - Gelang',
    type: 'Emas',
    unitLabel: 'gram',
    multiplier: 1,
    price: 1350000,
    prevPrice: 1348000,
    volatility: 0.0018,
    holdings: 1.7, // 11-Jun-2025: 1.7 gr -> 1,825.8 rb
    avgBuyPrice: 1825800 / 1.7, // 1,074.0 rb/gr
    color: '#DB2777'
  },
  {
    id: 'gold-kalung',
    symbol: 'PRH-KLG',
    name: 'Perhiasan - Kalung',
    type: 'Emas',
    unitLabel: 'gram',
    multiplier: 1,
    price: 1350000,
    prevPrice: 1347000,
    volatility: 0.0018,
    holdings: 4.8, // 11-Jun-2025: 4.8 gr -> 5,289.6 rb
    avgBuyPrice: 5289600 / 4.8, // 1,102.0 rb/gr
    color: '#BE185D'
  },

  // 2. KATEGORI DEPOSITO BANK (SEABANK & BCA)
  {
    id: 'dep-seabank',
    symbol: 'DEP-SEA',
    name: 'Deposito SeaBank (6.0% p.a)',
    type: 'Deposito',
    unitLabel: 'unit',
    multiplier: 1,
    price: 1000, // 1 unit = 1 ribu (Rp 1.000) agar mudah dibaca
    prevPrice: 1000,
    volatility: 0.00015,
    holdings: 25000, // 1-Oct-2026: 25,000.0 rb (Rp 25.000.000)
    avgBuyPrice: 1000,
    color: '#F97316'
  },
  {
    id: 'dep-bca',
    symbol: 'DEP-BCA',
    name: 'Deposito Berjangka BCA (3.5% p.a)',
    type: 'Deposito',
    unitLabel: 'unit',
    multiplier: 1,
    price: 1000,
    prevPrice: 1000,
    volatility: 0.0001,
    holdings: 0,
    avgBuyPrice: 1000,
    color: '#0284C7'
  },

  // 3. KATEGORI SAHAM & REKSADANA SAHAM (SUCORINVEST EQUITY FUND KELAS A)
  {
    id: 'stock-sef-a',
    symbol: 'SEF-A',
    name: 'Sucorinvest Equity Fund Kelas A',
    type: 'Saham',
    unitLabel: 'unit',
    multiplier: 1,
    price: 2510.19,
    prevPrice: 2507.45,
    volatility: 0.0035,
    holdings: 4749.2863, // 9-Jan-2026: modal real 15,900.0 rb (Rp 15.900.000)
    avgBuyPrice: 15900000 / 4749.2863, // ~3,347.8714
    color: '#2563EB'
  },
  {
    id: 'stock-bbca',
    symbol: 'BBCA',
    name: 'Bank Central Asia Tbk',
    type: 'Saham',
    unitLabel: 'lot',
    multiplier: 100,
    price: 10450,
    prevPrice: 10375,
    volatility: 0.0055,
    holdings: 0,
    avgBuyPrice: 10450,
    color: '#3B82F6'
  },

  // 4. KATEGORI VALAS (USD, JPY, SGD, AUD, MYR, VND)
  {
    id: 'fx-usd',
    symbol: 'USD',
    name: 'US Dollar Amerika',
    type: 'Valas',
    unitLabel: 'USD',
    multiplier: 1,
    price: 15850,
    prevPrice: 15820,
    volatility: 0.002,
    holdings: 0,
    avgBuyPrice: 15850,
    color: '#06B6D4'
  },
  {
    id: 'fx-jpy',
    symbol: 'JPY',
    name: 'Japanese Yen',
    type: 'Valas',
    unitLabel: 'JPY',
    multiplier: 1,
    price: 106.5,
    prevPrice: 106.1,
    volatility: 0.003,
    holdings: 0,
    avgBuyPrice: 106.5,
    color: '#0891B2'
  },
  {
    id: 'fx-sgd',
    symbol: 'SGD',
    name: 'Singapore Dollar',
    type: 'Valas',
    unitLabel: 'SGD',
    multiplier: 1,
    price: 11820,
    prevPrice: 11800,
    volatility: 0.0018,
    holdings: 0,
    avgBuyPrice: 11820,
    color: '#0284C7'
  },
  {
    id: 'fx-aud',
    symbol: 'AUD',
    name: 'Australian Dollar',
    type: 'Valas',
    unitLabel: 'AUD',
    multiplier: 1,
    price: 10450,
    prevPrice: 10480,
    volatility: 0.0028,
    holdings: 0,
    avgBuyPrice: 10450,
    color: '#0EA5E9'
  },
  {
    id: 'fx-myr',
    symbol: 'MYR',
    name: 'Malaysian Ringgit',
    type: 'Valas',
    unitLabel: 'MYR',
    multiplier: 1,
    price: 3620,
    prevPrice: 3610,
    volatility: 0.0022,
    holdings: 0,
    avgBuyPrice: 3620,
    color: '#14B8A6'
  },
  {
    id: 'fx-vnd',
    symbol: 'VND',
    name: 'Vietnamese Dong',
    type: 'Valas',
    unitLabel: 'VND',
    multiplier: 1,
    price: 0.63,
    prevPrice: 0.63,
    volatility: 0.004,
    holdings: 0,
    avgBuyPrice: 0.63,
    color: '#0D9488'
  },

  // 5. KATEGORI REKSADANA PASAR UANG (RDPU)
  {
    id: 'rdpu-sucor',
    symbol: 'SMMF',
    name: 'Sucorinvest Money Market Fund',
    type: 'Reksadana Pasar Uang',
    unitLabel: 'unit',
    multiplier: 1,
    price: 1842.55,
    prevPrice: 1842.1,
    volatility: 0.0003,
    holdings: 0,
    avgBuyPrice: 1842.55,
    color: '#10B981'
  }
];

const INITIAL_DEBTS = [];

// HISTORI DATA REAL ASET & SALDO AWAL RIRIN & JEKI
const INITIAL_TRANSACTIONS = [
  {
    id: 'tx-real-11',
    user: 'Ririn',
    type: 'Saving',
    category: 'Emas',
    title: 'Antam (2.0 gr - Bonus/Hadiah)',
    amount: 0,
    date: '2026-11-11',
    assetId: 'gold-antam',
    units: 2.0,
    note: '11-Nov-2026 • 2.0 gr (Nominal 0)'
  },
  {
    id: 'tx-real-10',
    user: 'Ririn',
    type: 'Saving',
    category: 'Deposito',
    title: 'Deposito SeaBank / BCA',
    amount: 25000000,
    date: '2026-10-01',
    assetId: 'dep-seabank',
    units: 25000,
    note: '1-Oct-2026 • Penempatan Deposito 25,000.0 rb'
  },
  {
    id: 'tx-real-09',
    user: 'Ririn',
    type: 'Saving',
    category: 'Emas',
    title: 'Antam (5.0 gr)',
    amount: 13755000,
    date: '2026-05-22',
    assetId: 'gold-antam',
    units: 5.0,
    note: '22-May-2026 • Beli 5.0 gr senilai 13,755.0 rb'
  },
  {
    id: 'tx-real-08',
    user: 'Ririn',
    type: 'Saving',
    category: 'Emas',
    title: 'Antam (10.0 gr)',
    amount: 28980000,
    date: '2026-03-19',
    assetId: 'gold-antam',
    units: 10.0,
    note: '19-Mar-2026 • Beli 10.0 gr senilai 28,980.0 rb'
  },
  {
    id: 'tx-real-07',
    user: 'Jeki',
    type: 'Saving',
    category: 'Saham',
    title: 'Saham Sucorinvest Equity Fund Kelas A',
    amount: 15900000,
    date: '2026-01-09',
    assetId: 'stock-sef-a',
    units: 4749.2863,
    note: '9-Jan-2026 • Modal 15,900.0 rb (4,749.2863 unit)'
  },
  {
    id: 'tx-real-06',
    user: 'Ririn',
    type: 'Saving',
    category: 'Emas',
    title: 'Perhiasan - Kalung (4.8 gr)',
    amount: 5289600,
    date: '2025-06-11',
    assetId: 'gold-kalung',
    units: 4.8,
    note: '11-Jun-2025 • Kalung 4.8 gr senilai 5,289.6 rb'
  },
  {
    id: 'tx-real-05',
    user: 'Ririn',
    type: 'Saving',
    category: 'Emas',
    title: 'Perhiasan - Gelang (1.7 gr)',
    amount: 1825800,
    date: '2025-06-11',
    assetId: 'gold-gelang',
    units: 1.7,
    note: '11-Jun-2025 • Gelang 1.7 gr senilai 1,825.8 rb'
  },
  {
    id: 'tx-real-04',
    user: 'Ririn',
    type: 'Saving',
    category: 'Emas',
    title: 'Perhiasan - Anting Jepit (1.6 gr)',
    amount: 2452800,
    date: '2025-06-11',
    assetId: 'gold-anting',
    units: 1.6,
    note: '11-Jun-2025 • Anting Jepit 1.6 gr senilai 2,452.8 rb'
  },
  {
    id: 'tx-real-03',
    user: 'Ririn',
    type: 'Saving',
    category: 'Emas',
    title: 'Perhiasan - Cincin (2.0 gr)',
    amount: 3336000,
    date: '2025-06-11',
    assetId: 'gold-cincin',
    units: 2.0,
    note: '11-Jun-2025 • Cincin 2.0 gr senilai 3,336.0 rb'
  },
  {
    id: 'tx-real-02',
    user: 'Ririn',
    type: 'Saving',
    category: 'Emas',
    title: 'Semar LM (3.0 gr)',
    amount: 5628000,
    date: '2025-06-11',
    assetId: 'gold-semar',
    units: 3.0,
    note: '11-Jun-2025 • LM Semar 3.0 gr senilai 5,628.0 rb'
  },
  {
    id: 'tx-saldo-ririn',
    user: 'Ririn',
    type: 'Income',
    category: 'Saldo Awal',
    title: 'Saldo Awal Kas Ririn',
    amount: 20376410,
    date: '2025-06-01',
    isOpening: true,
    note: 'Saldo kas aktif Ririn: 20,376.410 rb'
  },
  {
    id: 'tx-saldo-jeki',
    user: 'Jeki',
    type: 'Income',
    category: 'Saldo Awal',
    title: 'Saldo Awal Kas Jeki',
    amount: 2875000,
    date: '2025-06-01',
    isOpening: true,
    note: 'Saldo kas aktif Jeki: 2,875.000 rb'
  }
];

const CATEGORIES = {
  Income: ['Gaji', 'Bonus', 'Freelance', 'Dividen', 'Bunga Deposito', 'Saldo Awal', 'Lainnya'],
  Outcome: [
    'Makanan & Minuman',
    'Transportasi',
    'Tagihan & Utilitas',
    'Belanja',
    'Hiburan',
    'Kesehatan'
  ],
  Debt: ['Cicilan KPR/Kendaraan', 'Kartu Kredit', 'Pinjaman Pribadi'],
  Saving: [
    'Tabungan Tunai/Bank',
    'Deposito',
    'Emas',
    'Valas',
    'Saham',
    'Reksadana Pasar Uang'
  ]
};

const CASHFLOW_TREND_DATA = [
  { bulan: 'Jun', income: 23251410, outcome: 4200000 },
  { bulan: 'Jul', income: 18500000, outcome: 6100000 },
  { bulan: 'Agt', income: 19200000, outcome: 5800000 },
  { bulan: 'Sep', income: 20100000, outcome: 6400000 },
  { bulan: 'Okt', income: 25000000, outcome: 5900000 },
  { bulan: 'Nov', income: 23251410, outcome: 4850000 }
];


// --- HELPER FORMAT RIBUAN (CONTOH: 2,875 / 20,376.410) & RUPIAH PENUH ---
const formatRibu = (num, showUnit = true) => {
  if (isNaN(num)) return showUnit ? '0 rb' : '0';
  const inThousands = num / 1000;
  const hasDecimals = Math.abs(inThousands % 1) > 0.0001;
  const formatted = new Intl.NumberFormat('en-US', {
    minimumFractionDigits: hasDecimals ? 1 : 0,
    maximumFractionDigits: 3
  }).format(inThousands);
  return showUnit ? `${formatted} rb` : formatted;
};

const formatIDR = (num, compact = false, inThousandsMode = true) => {
  if (isNaN(num)) return inThousandsMode ? '0 rb' : 'Rp 0';
  if (inThousandsMode) {
    return formatRibu(num, true);
  }
  if (compact) {
    const abs = Math.abs(num);
    const sign = num < 0 ? '-' : '';
    if (abs >= 1_000_000_000) return `${sign}Rp ${(abs / 1_000_000_000).toFixed(2)} M`;
    if (abs >= 1_000_000) return `${sign}Rp ${(abs / 1_000_000).toFixed(1)} Jt`;
    if (abs >= 1_000) return `${sign}Rp ${(abs / 1_000).toFixed(0)} Rb`;
  }
  return new Intl.NumberFormat('id-ID', {
    style: 'currency',
    currency: 'IDR',
    minimumFractionDigits: 0,
    maximumFractionDigits: 0
  }).format(Math.round(num));
};

const formatUnitIDR = (num, inThousandsMode = false) => {
  if (isNaN(num)) return 'Rp 0';
  if (inThousandsMode && num >= 1000) {
    return formatRibu(num, true);
  }
  const hasDecimals = num < 10000 && num % 1 !== 0;
  return new Intl.NumberFormat('id-ID', {
    style: 'currency',
    currency: 'IDR',
    minimumFractionDigits: hasDecimals ? 2 : 0,
    maximumFractionDigits: num < 10 ? 3 : num < 10000 ? 4 : 2
  }).format(num);
};

const formatUnitsCount = (units) => {
  if (isNaN(units)) return '0';
  return new Intl.NumberFormat('en-US', {
    minimumFractionDigits: 0,
    maximumFractionDigits: units % 1 !== 0 ? 4 : 1
  }).format(units);
};

// --- HELPER EFEK MUTASI KAS PER TRANSAKSI (UNTUK RUNNING BALANCE) ---
const getCashDelta = (tx) => {
  if (tx.type === 'Income') return Number(tx.amount) || 0;
  if (tx.type === 'Outcome') return -(Number(tx.amount) || 0);
  if (tx.type === 'Debt') {
    return tx.debtAction === 'pay' ? -(Number(tx.amount) || 0) : Number(tx.amount) || 0;
  }
  if (tx.type === 'Saving') {
    return tx.category === 'Tabungan Tunai/Bank'
      ? Number(tx.amount) || 0
      : -(Number(tx.amount) || 0);
  }
  return 0;
};

// --- HELPER GAMBAR ROUNDED RECTANGLE DI CANVAS ---
function drawRoundRect(ctx, x, y, width, height, radius) {
  ctx.beginPath();
  ctx.moveTo(x + radius, y);
  ctx.lineTo(x + width - radius, y);
  ctx.quadraticCurveTo(x + width, y, x + width, y + radius);
  ctx.lineTo(x + width, y + height - radius);
  ctx.quadraticCurveTo(x + width, y + height, x + width - radius, y + height);
  ctx.lineTo(x + radius, y + height);
  ctx.quadraticCurveTo(x, y + height, x, y + height - radius);
  ctx.lineTo(x, y + radius);
  ctx.quadraticCurveTo(x, y, x + radius, y);
  ctx.closePath();
}

// --- KOMPONEN UTAMA ---
export default function App() {
  const [darkMode, setDarkMode] = useState(true);
  const [inThousandsMode, setInThousandsMode] = useState(true); // Default tampilan dalam ribuan (contoh: 2,875 / 20,376.410)
  const [activeTab, setActiveTab] = useState('dashboard');

  // --- STATE USER AKTIF (RIRIN vs JEKI vs SEMUA) ---
  const [activeUserScope, setActiveUserScope] = useState('Semua'); // 'Semua' | 'Ririn' | 'Jeki'
  const [txUserFilter, setTxUserFilter] = useState('Semua'); // 'Semua' | 'Ririn' | 'Jeki'

  // --- STATE AUTENTIKASI & CLOUD SYNC ---
  const [user, setUser] = useState(null);
  const [customVaultKey, setCustomVaultKey] = useState('');
  const [syncStatus, setSyncStatus] = useState('connecting'); // connecting | synced | saving | error
  const [lastCloudSync, setLastCloudSync] = useState(null);
  const isRemoteSnapshotRef = useRef(false);
  const hasLoadedInitialCloudRef = useRef(false);

  // --- STATE DATA KEUANGAN (SALDO AWAL REAL: JEKI 2,875 + RIRIN 20,376.410 = 23,251.410) ---
  const [transactions, setTransactions] = useState(INITIAL_TRANSACTIONS);
  const [assets, setAssets] = useState(INITIAL_MARKET_ASSETS);
  const [debts, setDebts] = useState(INITIAL_DEBTS);
  const [baseCashBalance, setBaseCashBalance] = useState(
    INITIAL_USER_OPENING.Jeki + INITIAL_USER_OPENING.Ririn
  );

  // --- STATE UI, MARKET & GOOGLE SHEETS ---
  const [isMarketLive, setIsMarketLive] = useState(true);
  const [lastMarketUpdate, setLastMarketUpdate] = useState(new Date());
  const [txFilter, setTxFilter] = useState('Semua');
  const [marketFilter, setMarketFilter] = useState('Semua');

  const [sheetsWebhookUrl, setSheetsWebhookUrl] = useState('');
  const [sheetsAutoSync, setSheetsAutoSync] = useState(true);
  const [sheetsSyncing, setSheetsSyncing] = useState(false);
  const [lastSheetsSyncTime, setLastSheetsSyncTime] = useState(null);

  const [isTxModalOpen, setIsTxModalOpen] = useState(false);
  const [tradeModal, setTradeModal] = useState({ open: false, asset: null, action: 'buy' });
  const [pngModal, setPngModal] = useState({ open: false, dataUrl: '' });
  const [toastMessage, setToastMessage] = useState(null);
  const [confirmResetModal, setConfirmResetModal] = useState(false);

  const showToast = (text, variant = 'success') => {
    setToastMessage({ text, variant });
    setTimeout(() => setToastMessage(null), 3200);
  };

  // --- AUTENTIKASI FIREBASE OTOMATIS ---
  useEffect(() => {
    const initAuth = async () => {
      try {
        if (!auth.currentUser) {
          await signInAnonymously(auth);
        }
      } catch (err) {
        console.error('Auth error:', err);
        setSyncStatus('error');
      }
    };
    initAuth();
    const unsubAuth = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser);
    });
    return () => unsubAuth();
  }, []);

  const getCloudDocRef = () => {
    if (!user) return null;
    const vaultClean = customVaultKey.trim().toLowerCase().replace(/[^a-z0-9-_]/g, '');
    if (vaultClean.length >= 4) {
      return doc(db, 'artifacts', appId, 'public', 'data', `vault_${vaultClean}`);
    }
    return doc(db, 'artifacts', appId, 'users', user.uid, 'finance', 'state');
  };

  // Fungsi simpan eksplisit ke Firestore Cloud
  const persistStateToCloud = async (overrideState = {}) => {
    if (!user) return;
    const docRef = getCloudDocRef();
    if (!docRef) return;

    try {
      setSyncStatus('saving');
      const payload = {
        transactions: overrideState.transactions ?? transactions,
        assets: (overrideState.assets ?? assets).map((a) => ({
          id: a.id,
          symbol: a.symbol,
          holdings: a.holdings,
          avgBuyPrice: a.avgBuyPrice
        })),
        debts: overrideState.debts ?? debts,
        baseCashBalance: overrideState.baseCashBalance ?? baseCashBalance,
        sheetsWebhookUrl: overrideState.sheetsWebhookUrl ?? sheetsWebhookUrl,
        sheetsAutoSync: overrideState.sheetsAutoSync ?? sheetsAutoSync,
        updatedAt: new Date().toISOString(),
        ownerEmail: user.email || 'anonymous'
      };
      await setDoc(docRef, payload, { merge: true });
      setLastCloudSync(new Date());
      setSyncStatus('synced');
    } catch (err) {
      console.error('Gagal menyimpan ke Cloud:', err);
      setSyncStatus('error');
    }
  };

  // --- REALTIME SNAPSHOT FIRESTORE CLOUD ---
  useEffect(() => {
    if (!user) return;
    const docRef = getCloudDocRef();
    if (!docRef) return;

    // eslint-disable-next-line react-hooks/set-state-in-effect
    setSyncStatus('connecting');
    const unsubscribe = onSnapshot(
      docRef,
      (snap) => {
        if (snap.exists()) {
          const data = snap.data();
          isRemoteSnapshotRef.current = true;
          if (Array.isArray(data.transactions)) setTransactions(data.transactions);
          if (Array.isArray(data.debts)) setDebts(data.debts);
          if (typeof data.baseCashBalance === 'number') setBaseCashBalance(data.baseCashBalance);
          if (typeof data.sheetsWebhookUrl === 'string') setSheetsWebhookUrl(data.sheetsWebhookUrl);
          if (typeof data.sheetsAutoSync === 'boolean') setSheetsAutoSync(data.sheetsAutoSync);

          if (Array.isArray(data.assets)) {
            const savedMap = new Map(data.assets.map((a) => [a.id, a]));
            setAssets(() =>
              INITIAL_MARKET_ASSETS.map((def) =>
                savedMap.has(def.id)
                  ? {
                      ...def,
                      holdings: savedMap.get(def.id).holdings ?? def.holdings,
                      avgBuyPrice: savedMap.get(def.id).avgBuyPrice ?? def.avgBuyPrice
                    }
                  : def
              )
            );
          }
          setLastCloudSync(new Date());
          setSyncStatus('synced');
          hasLoadedInitialCloudRef.current = true;
        } else {
          hasLoadedInitialCloudRef.current = true;
          persistStateToCloud({
            transactions,
            assets,
            debts,
            baseCashBalance,
            sheetsWebhookUrl,
            sheetsAutoSync
          });
        }
      },
      (err) => {
        console.error('Firestore snapshot error:', err);
        setSyncStatus('error');
      }
    );

    return () => unsubscribe();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [user, customVaultKey]);

  // --- LOGIN GOOGLE UNTUK SINKRONISASI LINTAS PERANGKAT ---
  const handleGoogleSignIn = async () => {
    try {
      const provider = new GoogleAuthProvider();
      await signInWithPopup(auth, provider);
      showToast('Berhasil terhubung dengan Akun Google!');
    } catch (err) {
      console.error('Google Sign-In error:', err);
      showToast('Popup Google diblokir/tidak tersedia. Gunakan Kode Sinkronisasi Cloud di bawah.', 'info');
    }
  };

  const handleSignOutToAnon = async () => {
    try {
      await signOut(auth);
      await signInAnonymously(auth);
      showToast('Keluar dari akun Google');
    } catch (err) {
      console.error('Sign out error:', err);
    }
  };

  // --- INTEGRASI GOOGLE SHEETS WEBHOOK ---
  const pushToGoogleSheets = async (mode = 'full_sync', singleTx = null, customUrl = null) => {
    const targetUrl = (customUrl ?? sheetsWebhookUrl).trim();
    if (!targetUrl) {
      if (mode === 'full_sync') {
        showToast('Masukkan URL Google Apps Script terlebih dahulu!', 'info');
      }
      return false;
    }

    try {
      setSheetsSyncing(true);
      const bodyPayload =
        mode === 'single_tx'
          ? { mode: 'single_tx', transaction: singleTx }
          : {
              mode: 'full_sync',
              transactions: transactionsWithRunningBalance,
              assets,
              debts,
              baseCashBalance
            };

      await fetch(targetUrl, {
        method: 'POST',
        mode: 'no-cors',
        headers: {
          'Content-Type': 'text/plain;charset=utf-8'
        },
        body: JSON.stringify(bodyPayload)
      });

      setLastSheetsSyncTime(new Date());
      if (mode === 'full_sync') {
        showToast('Data berhasil dikirim ke Google Sheets!');
      }
      return true;
    } catch (err) {
      console.error('Google Sheets sync error:', err);
      if (mode === 'full_sync') {
        showToast('Gagal menghubungi Google Sheets. Periksa URL Web App.', 'info');
      }
      return false;
    } finally {
      setSheetsSyncing(false);
    }
  };

  // --- HITUNG RUNNING BALANCE TOTAL (REKENING KORAN KRONOLOGIS) ---
  // (React Compiler Next 16 sudah memoize otomatis, useMemo manual tidak diperlukan)
  const transactionsWithRunningBalance = (() => {
    // Transaksi historis aset awal (tx-real-*) merupakan portofolio existing di luar saldo kas tunai aktif
    const deltas = transactions.map((tx) =>
      tx.id?.startsWith('tx-real-') ? 0 : getCashDelta(tx)
    );
    // balances[i] = saldo SESUDAH transaksi i (urutan terbaru -> terlama)
    const balances = deltas.reduce(
      (acc, delta) => {
        acc.push(acc[acc.length - 1] - delta);
        return acc;
      },
      [baseCashBalance]
    );
    return transactions.map((tx, i) => ({
      ...tx,
      user: tx.user || 'Ririn',
      cashDelta: deltas[i],
      runningBalance: balances[i]
    }));
  })();

  // --- EXPORT SPREADSHEET CSV (DENGAN KOLOM USER RIRIN/JEKI & RUNNING BALANCE) ---
  const handleDownloadCSV = () => {
    const headers = [
      'ID',
      'Tanggal',
      'User',
      'Tipe',
      'Kategori',
      'Judul',
      'Nominal (IDR)',
      'Mutasi Kas (IDR)',
      'Running Balance (IDR)',
      'Catatan'
    ];
    const rows = transactionsWithRunningBalance.map((t) => [
      t.id,
      t.date,
      t.user,
      t.type,
      `"${(t.category || '').replace(/"/g, '""')}"`,
      `"${(t.title || '').replace(/"/g, '""')}"`,
      t.amount,
      t.cashDelta,
      t.runningBalance,
      `"${(t.note || '').replace(/"/g, '""')}"`
    ]);

    const csvContent =
      '\uFEFF' + [headers.join(','), ...rows.map((r) => r.join(','))].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute(
      'download',
      `Laporan_DompetKu_Ririn_Jeki_${new Date().toISOString().split('T')[0]}.csv`
    );
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showToast('File CSV dengan Running Balance & User berhasil diunduh!');
  };

  // --- SIMULASI LIVE MARKET HARGA REAL-TIME ---
  const simulateMarketTick = () => {
    setAssets((prevAssets) =>
      prevAssets.map((asset) => {
        const direction =
          asset.type === 'Reksadana Pasar Uang' || asset.type === 'Deposito'
            ? Math.random() > 0.1
              ? 1
              : 0
            : Math.random() > 0.48
            ? 1
            : -1;

        const changeFactor = direction * Math.random() * asset.volatility;
        let newPrice = asset.price * (1 + changeFactor);

        if (asset.type === 'Emas') {
          newPrice = Math.round(newPrice / 500) * 500;
        } else if (asset.type === 'Deposito') {
          newPrice = Number(newPrice.toFixed(2));
        } else if (asset.type === 'Valas') {
          if (asset.symbol === 'VND') {
            newPrice = Math.max(0.5, Number(newPrice.toFixed(3)));
          } else if (asset.symbol === 'JPY') {
            newPrice = Math.max(50, Number(newPrice.toFixed(2)));
          } else {
            newPrice = Math.max(100, Math.round(newPrice / 5) * 5);
          }
        } else if (asset.type === 'Saham') {
          if (asset.unitLabel === 'unit') {
            newPrice = Math.max(100, Number(newPrice.toFixed(4)));
          } else {
            const tick = newPrice >= 5000 ? 25 : newPrice >= 2000 ? 10 : 5;
            newPrice = Math.max(50, Math.round(newPrice / tick) * tick);
          }
        } else {
          newPrice = Number(newPrice.toFixed(2));
        }

        return {
          ...asset,
          prevPrice: asset.price,
          price: newPrice
        };
      })
    );
    setLastMarketUpdate(new Date());
  };

  useEffect(() => {
    if (!isMarketLive) return;
    const interval = setInterval(simulateMarketTick, 3500);
    return () => clearInterval(interval);
  }, [isMarketLive]);

  // --- KALKULASI RINGKASAN KEUANGAN (GABUNGAN & PER USER RIRIN / JEKI) ---
  const summary = useMemo(() => {
    const scopedTxs =
      activeUserScope === 'Semua'
        ? transactions
        : transactions.filter((t) => (t.user || 'Ririn') === activeUserScope);

    const totalIncome = scopedTxs
      .filter((t) => t.type === 'Income')
      .reduce((acc, t) => acc + t.amount, 0);

    const totalOutcome = scopedTxs
      .filter((t) => t.type === 'Outcome')
      .reduce((acc, t) => acc + t.amount, 0);

    const totalDebtPayments = scopedTxs
      .filter((t) => t.type === 'Debt' && t.debtAction === 'pay')
      .reduce((acc, t) => acc + t.amount, 0);

    // Breakdown kontribusi saldo & arus kas Ririn vs Jeki
    const byUser = ['Ririn', 'Jeki'].reduce((acc, person) => {
      const pTxs = transactions.filter((t) => (t.user || 'Ririn') === person);
      const inc = pTxs
        .filter((t) => t.type === 'Income')
        .reduce((s, t) => s + t.amount, 0);
      const out = pTxs
        .filter((t) => t.type === 'Outcome' || (t.type === 'Debt' && t.debtAction === 'pay'))
        .reduce((s, t) => s + t.amount, 0);
      const sav = pTxs
        .filter((t) => t.type === 'Saving')
        .reduce((s, t) => s + t.amount, 0);
      const netFlow = pTxs.reduce(
        (s, t) => s + (t.id?.startsWith('tx-real-') ? 0 : getCashDelta(t)),
        0
      );
      acc[person] = { income: inc, outcome: out, saving: sav, netFlow };
      return acc;
    }, {});

    const portfolioStats = assets.reduce(
      (acc, asset) => {
        const currentValue = asset.holdings * asset.multiplier * asset.price;
        const investedCapital = asset.holdings * asset.multiplier * asset.avgBuyPrice;
        acc.totalValue += currentValue;
        acc.totalCost += investedCapital;
        acc.byType[asset.type] = (acc.byType[asset.type] || 0) + currentValue;
        return acc;
      },
      { totalValue: 0, totalCost: 0, byType: {} }
    );

    const unrealizedPnL = portfolioStats.totalValue - portfolioStats.totalCost;
    const unrealizedPnLPercent =
      portfolioStats.totalCost > 0 ? (unrealizedPnL / portfolioStats.totalCost) * 100 : 0;

    const totalSavingAndInvestment = baseCashBalance + portfolioStats.totalValue;
    const totalDebtRemaining = debts.reduce((acc, d) => acc + d.remainingAmount, 0);
    const totalDebtInitial = debts.reduce((acc, d) => acc + d.totalAmount, 0);

    const netWorth = totalSavingAndInvestment - totalDebtRemaining;

    const outcomeMap = {};
    scopedTxs
      .filter((t) => t.type === 'Outcome')
      .forEach((t) => {
        outcomeMap[t.category] = (outcomeMap[t.category] || 0) + t.amount;
      });

    const outcomeBreakdown = Object.entries(outcomeMap)
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value);

    return {
      totalIncome,
      totalOutcome: totalOutcome + totalDebtPayments,
      pureOutcome: totalOutcome,
      totalDebtPayments,
      totalSavingAndInvestment,
      baseCashBalance,
      portfolioValue: portfolioStats.totalValue,
      portfolioByType: portfolioStats.byType,
      unrealizedPnL,
      unrealizedPnLPercent,
      totalDebtRemaining,
      totalDebtInitial,
      netWorth,
      outcomeBreakdown,
      byUser
    };
  }, [transactions, assets, debts, baseCashBalance, activeUserScope]);

  // --- GENERATOR LAPORAN KEUANGAN KE PNG RESOLUSI TINGGI ---
  const handleGenerateReportPNG = () => {
    const scale = 2; // Retina 2x High-Res (1080 x 1680 px)
    const width = 540;
    const height = 840;
    const canvas = document.createElement('canvas');
    canvas.width = width * scale;
    canvas.height = height * scale;
    const ctx = canvas.getContext('2d');
    ctx.scale(scale, scale);

    // 1. Background Utama
    const bgGrad = ctx.createLinearGradient(0, 0, width, height);
    bgGrad.addColorStop(0, '#090D16');
    bgGrad.addColorStop(1, '#0F172A');
    ctx.fillStyle = bgGrad;
    ctx.fillRect(0, 0, width, height);

    // 2. Header Banner
    const headerGrad = ctx.createLinearGradient(24, 24, width - 24, 195);
    headerGrad.addColorStop(0, '#059669');
    headerGrad.addColorStop(0.6, '#0D9488');
    headerGrad.addColorStop(1, '#111827');
    drawRoundRect(ctx, 24, 24, width - 48, 172, 22);
    ctx.fillStyle = headerGrad;
    ctx.fill();

    ctx.fillStyle = '#A7F3D0';
    ctx.font = 'bold 12px sans-serif';
    ctx.fillText('CASHFLOW LIVE • LAPORAN KEUANGAN RIRIN & JEKI', 44, 52);

    ctx.fillStyle = '#FFFFFF';
    ctx.font = '11px sans-serif';
    const printDate = new Date().toLocaleString('id-ID', {
      day: '2-digit',
      month: 'short',
      year: 'numeric',
      hour: '2-digit',
      minute: '2-digit'
    });
    ctx.fillText(`Dicetak: ${printDate} WIB • Filter: ${activeUserScope}`, 44, 70);

    ctx.fillStyle = '#D1FAE5';
    ctx.font = '12px sans-serif';
    ctx.fillText('Total Kekayaan Bersih (Net Worth) & Running Balance Kas', 44, 102);

    ctx.fillStyle = '#FFFFFF';
    ctx.font = 'bold 28px sans-serif';
    ctx.fillText(formatIDR(summary.netWorth, false, inThousandsMode), 44, 136);

    ctx.fillStyle = '#FDE68A';
    ctx.font = 'bold 11.5px sans-serif';
    ctx.fillText(
      `Kas Aktif: ${formatIDR(summary.baseCashBalance, false, inThousandsMode)}  |  Portofolio: ${formatIDR(
        summary.portfolioValue,
        false,
        inThousandsMode
      )}`,
      44,
      166
    );

    // 3. Breakdown Kontribusi User (Ririn vs Jeki)
    drawRoundRect(ctx, 24, 208, 240, 66, 16);
    ctx.fillStyle = '#1E1B2E';
    ctx.fill();
    ctx.strokeStyle = '#F43F5E';
    ctx.lineWidth = 1;
    ctx.stroke();
    ctx.fillStyle = '#FDA4AF';
    ctx.font = 'bold 10.5px sans-serif';
    ctx.fillText('USER: RIRIN (Saldo Kas)', 38, 230);
    ctx.fillStyle = '#FFFFFF';
    ctx.font = 'bold 15px sans-serif';
    const ririnNet = summary.byUser?.Ririn?.netFlow || 0;
    ctx.fillText(`${formatIDR(ririnNet, false, inThousandsMode)}`, 38, 254);

    drawRoundRect(ctx, 276, 208, 240, 66, 16);
    ctx.fillStyle = '#17253B';
    ctx.fill();
    ctx.strokeStyle = '#38BDF8';
    ctx.lineWidth = 1;
    ctx.stroke();
    ctx.fillStyle = '#7DD3FC';
    ctx.font = 'bold 10.5px sans-serif';
    ctx.fillText('USER: JEKI (Saldo Kas)', 290, 230);
    ctx.fillStyle = '#FFFFFF';
    ctx.font = 'bold 15px sans-serif';
    const jekiNet = summary.byUser?.Jeki?.netFlow || 0;
    ctx.fillText(`${formatIDR(jekiNet, false, inThousandsMode)}`, 290, 254);

    // 4. Ringkasan 4 Pilar Cashflow
    const cards = [
      {
        label: 'PEMASUKAN / SALDO KAS',
        val: formatIDR(summary.totalIncome, false, inThousandsMode),
        color: '#34D399',
        x: 24,
        y: 286
      },
      {
        label: 'PENGELUARAN (OUTCOME)',
        val: formatIDR(summary.totalOutcome, false, inThousandsMode),
        color: '#FB7185',
        x: 276,
        y: 286
      },
      {
        label: 'TOTAL KAS & INVESTASI',
        val: formatIDR(summary.totalSavingAndInvestment, false, inThousandsMode),
        color: '#FBBF24',
        x: 24,
        y: 366
      },
      {
        label: 'SISA HUTANG (DEBT)',
        val: formatIDR(summary.totalDebtRemaining, false, inThousandsMode),
        color: '#C084FC',
        x: 276,
        y: 366
      }
    ];

    cards.forEach((c) => {
      drawRoundRect(ctx, c.x, c.y, 240, 68, 16);
      ctx.fillStyle = '#1E293B';
      ctx.fill();
      ctx.strokeStyle = '#334155';
      ctx.lineWidth = 1;
      ctx.stroke();

      ctx.fillStyle = '#94A3B8';
      ctx.font = 'bold 10px sans-serif';
      ctx.fillText(c.label, c.x + 16, c.y + 24);

      ctx.fillStyle = c.color;
      ctx.font = 'bold 17px sans-serif';
      ctx.fillText(c.val, c.x + 16, c.y + 50);
    });

    // 5. Rincian Komposisi Aset & Harga Emas + Deposito + Sucorinvest Equity Fund Kelas A
    drawRoundRect(ctx, 24, 448, width - 48, 224, 18);
    ctx.fillStyle = '#1E293B';
    ctx.fill();
    ctx.strokeStyle = '#334155';
    ctx.stroke();

    ctx.fillStyle = '#F8FAFC';
    ctx.font = 'bold 13.5px sans-serif';
    ctx.fillText('Portofolio Aset Real & Harga Pasar Live', 42, 474);

    const antamPrice = assets.find((a) => a.id === 'gold-antam')?.price || 2495000;
    const semarPrice = assets.find((a) => a.id === 'gold-semar')?.price || 2236000;
    const sefAsset = assets.find((a) => a.id === 'stock-sef-a');
    const sefVal = sefAsset ? sefAsset.holdings * sefAsset.price : 0;

    const assetRows = [
      {
        name: 'Emas LM (Antam 17g, Semar 3g) & Perhiasan (10.1g)',
        sub: `Antam ${formatIDR(antamPrice, true, inThousandsMode)} | Semar ${formatIDR(semarPrice, true, inThousandsMode)}`,
        val: summary.portfolioByType['Emas'] || 0,
        dot: '#F59E0B'
      },
      {
        name: 'Deposito Bank (SeaBank / BCA)',
        sub: 'Penempatan Deposito 25,000.0 rb + Bunga Berjalan',
        val: summary.portfolioByType['Deposito'] || 0,
        dot: '#F97316'
      },
      {
        name: 'Sucorinvest Equity Fund Kelas A (SEF-A)',
        sub: `${formatUnitsCount(sefAsset?.holdings || 0)} unit • Modal 15,900.0 rb`,
        val: sefVal,
        dot: '#3B82F6'
      },
      {
        name: 'Valuta Asing (USD, JPY, SGD, AUD, MYR, VND)',
        sub: `Kurs USD: ${formatIDR(assets.find((a) => a.id === 'fx-usd')?.price || 0, false, inThousandsMode)}`,
        val: summary.portfolioByType['Valas'] || 0,
        dot: '#06B6D4'
      },
      {
        name: 'Reksadana Pasar Uang (RDPU)',
        sub: 'Sucorinvest Money Market Fund',
        val: summary.portfolioByType['Reksadana Pasar Uang'] || 0,
        dot: '#34D399'
      }
    ];

    assetRows.forEach((row, idx) => {
      const rowY = 502 + idx * 33;
      ctx.beginPath();
      ctx.arc(48, rowY - 4, 4.5, 0, Math.PI * 2);
      ctx.fillStyle = row.dot;
      ctx.fill();

      ctx.fillStyle = '#F1F5F9';
      ctx.font = 'bold 11.5px sans-serif';
      ctx.fillText(row.name, 62, rowY);

      ctx.fillStyle = '#94A3B8';
      ctx.font = '10px sans-serif';
      ctx.fillText(row.sub, 62, rowY + 13);

      ctx.fillStyle = '#FFFFFF';
      ctx.font = 'bold 12px sans-serif';
      ctx.textAlign = 'right';
      ctx.fillText(formatIDR(row.val, false, inThousandsMode), width - 42, rowY + 4);
      ctx.textAlign = 'left';
    });

    // 6. Mutasi & Running Balance Terakhir
    drawRoundRect(ctx, 24, 686, width - 48, 108, 16);
    ctx.fillStyle = '#1E293B';
    ctx.fill();

    ctx.fillStyle = '#94A3B8';
    ctx.font = 'bold 10.5px sans-serif';
    ctx.fillText('MUTASI & RUNNING BALANCE TERBARU', 42, 708);

    transactionsWithRunningBalance.slice(0, 3).forEach((tx, idx) => {
      const ty = 730 + idx * 21;
      ctx.fillStyle = tx.user === 'Ririn' ? '#FDA4AF' : '#7DD3FC';
      ctx.font = 'bold 10.5px sans-serif';
      ctx.fillText(`[${tx.user}]`, 42, ty);

      ctx.fillStyle = '#E2E8F0';
      ctx.font = '11px sans-serif';
      ctx.fillText(`${tx.title.slice(0, 26)}`, 86, ty);

      ctx.textAlign = 'right';
      ctx.fillStyle = '#94A3B8';
      ctx.font = '10px sans-serif';
      ctx.fillText(`Bal: ${formatIDR(tx.runningBalance, true, inThousandsMode)}`, width - 42, ty);
      ctx.textAlign = 'left';
    });

    ctx.fillStyle = '#64748B';
    ctx.font = '10px sans-serif';
    ctx.textAlign = 'center';
    ctx.fillText(
      'Generated by Cashflow Live • Ririn & Jeki Cloud Wealth Dashboard',
      width / 2,
      818
    );
    ctx.textAlign = 'left';

    const dataUrl = canvas.toDataURL('image/png', 1.0);
    setPngModal({ open: true, dataUrl });
  };

  // --- HANDLER TAMBAH TRANSAKSI BARU ---
  const handleAddTransaction = async (newTx) => {
    const txUser = newTx.user || (activeUserScope !== 'Semua' ? activeUserScope : 'Jeki');
    const tempTx = {
      ...newTx,
      user: txUser,
      id: `tx-${Date.now()}`,
      date: newTx.date || new Date().toISOString().split('T')[0]
    };

    let nextCash = baseCashBalance;
    let nextDebts = debts;
    let nextAssets = assets;

    if (newTx.type === 'Income') {
      nextCash = baseCashBalance + newTx.amount;
      showToast(`[${txUser}] Pemasukan ${formatIDR(newTx.amount)} dicatat!`);
    } else if (newTx.type === 'Outcome') {
      nextCash = Math.max(0, baseCashBalance - newTx.amount);
      showToast(`[${txUser}] Pengeluaran ${formatIDR(newTx.amount)} dicatat!`);
    } else if (newTx.type === 'Debt') {
      if (newTx.debtAction === 'pay' && newTx.debtId) {
        nextDebts = debts.map((d) =>
          d.id === newTx.debtId
            ? { ...d, remainingAmount: Math.max(0, d.remainingAmount - newTx.amount) }
            : d
        );
        nextCash = Math.max(0, baseCashBalance - newTx.amount);
        showToast(`[${txUser}] Cicilan ${formatIDR(newTx.amount)} dibayar!`);
      } else {
        const newDebtObj = {
          id: `debt-${Date.now()}`,
          user: txUser,
          name: newTx.title,
          category: newTx.category,
          totalAmount: newTx.amount,
          remainingAmount: newTx.amount,
          monthlyInstallment: Math.round(newTx.amount / 12),
          dueDate: 'Tgl 20 setiap bulan'
        };
        nextDebts = [newDebtObj, ...debts];
        nextCash = baseCashBalance + newTx.amount;
        showToast(`[${txUser}] Hutang baru ${formatIDR(newTx.amount)} dicatat!`);
      }
    } else if (newTx.type === 'Saving') {
      if (newTx.category === 'Tabungan Tunai/Bank') {
        nextCash = baseCashBalance + newTx.amount;
        showToast(`[${txUser}] Tabungan tunai ${formatIDR(newTx.amount)} ditambahkan!`);
      } else if (newTx.assetId && newTx.units > 0) {
        nextAssets = assets.map((asset) => {
          if (asset.id !== newTx.assetId) return asset;
          const totalOldCost = asset.holdings * asset.multiplier * asset.avgBuyPrice;
          const newHoldings = asset.holdings + Number(newTx.units);
          const newAvgPrice =
            newHoldings > 0
              ? (totalOldCost + newTx.amount) / (newHoldings * asset.multiplier)
              : asset.price;
          return {
            ...asset,
            holdings: Number(newHoldings.toFixed(4)),
            avgBuyPrice: newAvgPrice
          };
        });
        nextCash = Math.max(0, baseCashBalance - newTx.amount);
        showToast(`[${txUser}] Investasi ${newTx.category} ${formatIDR(newTx.amount)} berhasil!`);
      }
    }

    const txWithRunning = {
      ...tempTx,
      runningBalance: nextCash
    };
    const nextTransactions = [txWithRunning, ...transactions];

    setTransactions(nextTransactions);
    setBaseCashBalance(nextCash);
    setDebts(nextDebts);
    setAssets(nextAssets);
    setIsTxModalOpen(false);

    await persistStateToCloud({
      transactions: nextTransactions,
      baseCashBalance: nextCash,
      debts: nextDebts,
      assets: nextAssets
    });

    if (sheetsAutoSync && sheetsWebhookUrl.trim()) {
      pushToGoogleSheets('single_tx', txWithRunning);
    }
  };

  // --- HANDLER BELI / JUAL LANGSUNG DI TAB INVESTASI ---
  const handleAssetTrade = async ({ asset, action, units, tradeUser }) => {
    const unitNum = Number(units);
    if (!unitNum || unitNum <= 0) return;
    const actorUser = tradeUser || (activeUserScope !== 'Semua' ? activeUserScope : 'Jeki');

    const totalNominal = Math.round(unitNum * asset.multiplier * asset.price);
    let nextAssets = assets;
    let nextCash = baseCashBalance;
    let newTxRecord = null;

    if (action === 'buy') {
      nextAssets = assets.map((a) => {
        if (a.id !== asset.id) return a;
        const oldCost = a.holdings * a.multiplier * a.avgBuyPrice;
        const updatedHoldings = a.holdings + unitNum;
        const updatedAvg = (oldCost + totalNominal) / (updatedHoldings * a.multiplier);
        return {
          ...a,
          holdings: Number(updatedHoldings.toFixed(4)),
          avgBuyPrice: updatedAvg
        };
      });
      nextCash = Math.max(0, baseCashBalance - totalNominal);
      newTxRecord = {
        id: `tx-${Date.now()}`,
        user: actorUser,
        type: 'Saving',
        category: asset.type,
        title: `Beli ${asset.symbol} (${formatUnitsCount(unitNum)} ${asset.unitLabel})`,
        amount: totalNominal,
        runningBalance: nextCash,
        date: new Date().toISOString().split('T')[0],
        assetId: asset.id,
        units: unitNum,
        note: `Harga eksekusi ${formatUnitIDR(asset.price)}`
      };
      showToast(`[${actorUser}] Beli ${formatUnitsCount(unitNum)} ${asset.unitLabel} ${asset.symbol}!`);
    } else {
      const sellUnits = Math.min(unitNum, asset.holdings);
      const sellNominal = Math.round(sellUnits * asset.multiplier * asset.price);

      nextAssets = assets.map((a) =>
        a.id === asset.id
          ? { ...a, holdings: Number(Math.max(0, a.holdings - sellUnits).toFixed(4)) }
          : a
      );
      nextCash = baseCashBalance + sellNominal;
      newTxRecord = {
        id: `tx-${Date.now()}`,
        user: actorUser,
        type: 'Income',
        category: 'Lainnya',
        title: `Jual Aset ${asset.symbol} (${formatUnitsCount(sellUnits)} ${asset.unitLabel})`,
        amount: sellNominal,
        runningBalance: nextCash,
        date: new Date().toISOString().split('T')[0],
        note: `Pencairan investasi ke saldo tunai`
      };
      showToast(`[${actorUser}] Jual ${formatUnitsCount(sellUnits)} ${asset.unitLabel} ${asset.symbol}!`);
    }

    const nextTransactions = [newTxRecord, ...transactions];
    setAssets(nextAssets);
    setBaseCashBalance(nextCash);
    setTransactions(nextTransactions);
    setTradeModal({ open: false, asset: null, action: 'buy' });

    await persistStateToCloud({
      transactions: nextTransactions,
      baseCashBalance: nextCash,
      assets: nextAssets
    });

    if (sheetsAutoSync && sheetsWebhookUrl.trim()) {
      pushToGoogleSheets('single_tx', newTxRecord);
    }
  };

  const handleDeleteTransaction = async (id) => {
    const nextTransactions = transactions.filter((t) => t.id !== id);
    setTransactions(nextTransactions);
    showToast('Transaksi dihapus & disinkronkan ke Cloud', 'info');
    await persistStateToCloud({ transactions: nextTransactions });
  };

  const handleResetDemoData = async () => {
    const initialCash = INITIAL_USER_OPENING.Jeki + INITIAL_USER_OPENING.Ririn;
    setTransactions(INITIAL_TRANSACTIONS);
    setAssets(INITIAL_MARKET_ASSETS);
    setDebts(INITIAL_DEBTS);
    setBaseCashBalance(initialCash);
    setConfirmResetModal(false);
    showToast('Data dikembalikan ke data awal Ririn & Jeki!');
    await persistStateToCloud({
      transactions: INITIAL_TRANSACTIONS,
      assets: INITIAL_MARKET_ASSETS,
      debts: INITIAL_DEBTS,
      baseCashBalance: initialCash
    });
  };

  const theme = {
    bgOuter: darkMode ? 'bg-slate-950 text-slate-100' : 'bg-slate-200 text-slate-900',
    appShell: darkMode
      ? 'bg-slate-900 border-slate-800 shadow-2xl shadow-black/60'
      : 'bg-slate-50 border-slate-300 shadow-2xl shadow-slate-400/40',
    card: darkMode
      ? 'bg-slate-800/80 border-slate-700/70'
      : 'bg-white border-slate-200/90 shadow-sm',
    cardSubtle: darkMode ? 'bg-slate-800/40' : 'bg-slate-100/80',
    textMuted: darkMode ? 'text-slate-400' : 'text-slate-500',
    border: darkMode ? 'border-slate-800' : 'border-slate-200',
    navBg: darkMode
      ? 'bg-slate-900/95 border-slate-800'
      : 'bg-white/95 border-slate-200'
  };

  return (
    <div
      className={`h-dvh w-full flex justify-center sm:items-center sm:py-4 overflow-hidden transition-colors duration-300 ${theme.bgOuter}`}
    >
      <div
        className={`relative w-full max-w-[460px] h-full sm:max-h-[900px] sm:rounded-[36px] sm:border flex flex-col overflow-hidden ${theme.appShell}`}
      >
        {/* TOAST NOTIFICATION */}
        {toastMessage && (
          <div className="absolute top-16 left-4 right-4 z-50 animate-bounce">
            <div className="bg-emerald-600 text-white px-4 py-3 rounded-2xl shadow-lg flex items-center gap-2.5 text-xs font-medium">
              <CheckCircle2 className="w-4 h-4 shrink-0" />
              <span>{toastMessage.text}</span>
            </div>
          </div>
        )}

        {/* TOP APP BAR */}
        <header
          className={`px-4 pt-4 pb-3 flex items-center justify-between border-b shrink-0 backdrop-blur-md z-20 ${theme.border}`}
        >
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-xl bg-gradient-to-br from-emerald-500 to-teal-600 flex items-center justify-center text-white shadow-md shadow-emerald-500/20">
              <Wallet className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-1.5">
                <h1 className="font-bold text-sm sm:text-base tracking-tight">Cashflow Live</h1>
                <button
                  onClick={() => setActiveTab('sync')}
                  title="Status Sinkronisasi Cloud"
                  className={`inline-flex items-center gap-1 px-1.5 py-0.5 rounded-full text-[9px] font-bold ${
                    syncStatus === 'synced'
                      ? 'bg-emerald-500/15 text-emerald-400'
                      : syncStatus === 'saving' || syncStatus === 'connecting'
                      ? 'bg-amber-500/15 text-amber-400'
                      : 'bg-rose-500/15 text-rose-400'
                  }`}
                >
                  <Cloud className="w-2.5 h-2.5" />
                  {syncStatus === 'synced'
                    ? 'CLOUD'
                    : syncStatus === 'saving'
                    ? 'SYNCING'
                    : 'OFFLINE'}
                </button>
              </div>
              <p className={`text-[10px] truncate max-w-[165px] ${theme.textMuted}`}>
                {customVaultKey.trim()
                  ? `Vault: ${customVaultKey.trim()}`
                  : user?.email
                  ? user.email
                  : 'Cloud Multi-Device Aktif'}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-1.5">
            {/* SAKELAR FORMAT RIBUAN (2,875 rb) vs RUPIAH PENUH */}
            <button
              onClick={() => setInThousandsMode(!inThousandsMode)}
              title="Ubah Tampilan Ribuan (2,875) / Rupiah Penuh"
              className={`px-2 py-2 rounded-xl border text-[10px] font-extrabold transition ${
                inThousandsMode
                  ? 'bg-amber-400/15 border-amber-400/40 text-amber-400'
                  : `${theme.card} ${theme.textMuted}`
              }`}
            >
              {inThousandsMode ? 'Rb' : 'Rp'}
            </button>

            {/* TOMBOL EXPORT LAPORAN PNG */}
            <button
              onClick={handleGenerateReportPNG}
              title="Export Laporan ke PNG"
              className="px-2.5 py-2 rounded-xl bg-emerald-500/15 hover:bg-emerald-500/25 text-emerald-400 border border-emerald-500/30 flex items-center gap-1 text-[11px] font-bold transition"
            >
              <ImageIcon className="w-3.5 h-3.5" />
              <span>PNG</span>
            </button>

            <button
              onClick={() => setDarkMode(!darkMode)}
              title="Ganti Tema"
              className={`p-2 rounded-xl border transition ${theme.card}`}
            >
              {darkMode ? (
                <Sun className="w-4 h-4 text-amber-400" />
              ) : (
                <Moon className="w-4 h-4 text-slate-700" />
              )}
            </button>

            <button
              onClick={() => setConfirmResetModal(true)}
              title="Reset Data Awal"
              className={`p-2 rounded-xl border transition ${theme.card}`}
            >
              <RotateCcw className={`w-4 h-4 ${theme.textMuted}`} />
            </button>
          </div>
        </header>

        {/* AREA KONTEN SCROLLABLE */}
        <main className="flex-1 min-h-0 overflow-y-auto overscroll-contain px-4 pt-4 pb-28 space-y-4 scroll-smooth">
          {activeTab === 'dashboard' && (
            <DashboardView
              summary={summary}
              assets={assets}
              transactions={transactionsWithRunningBalance}
              activeUserScope={activeUserScope}
              setActiveUserScope={setActiveUserScope}
              inThousandsMode={inThousandsMode}
              darkMode={darkMode}
              theme={theme}
              isMarketLive={isMarketLive}
              setIsMarketLive={setIsMarketLive}
              onNavigate={setActiveTab}
              onFilterMarket={(cat) => {
                setMarketFilter(cat);
                setActiveTab('market');
              }}
              onExportPNG={handleGenerateReportPNG}
            />
          )}

          {activeTab === 'transactions' && (
            <TransactionsView
              transactions={transactionsWithRunningBalance}
              baseCashBalance={baseCashBalance}
              txFilter={txFilter}
              setTxFilter={setTxFilter}
              txUserFilter={txUserFilter}
              setTxUserFilter={setTxUserFilter}
              inThousandsMode={inThousandsMode}
              onDelete={handleDeleteTransaction}
              onOpenTxModal={() => setIsTxModalOpen(true)}
              theme={theme}
            />
          )}

          {activeTab === 'market' && (
            <MarketView
              assets={assets}
              summary={summary}
              inThousandsMode={inThousandsMode}
              isMarketLive={isMarketLive}
              setIsMarketLive={setIsMarketLive}
              onManualRefresh={simulateMarketTick}
              lastMarketUpdate={lastMarketUpdate}
              marketFilter={marketFilter}
              setMarketFilter={setMarketFilter}
              onOpenTrade={(asset, action) =>
                setTradeModal({ open: true, asset, action })
              }
              theme={theme}
            />
          )}

          {activeTab === 'debt' && (
            <DebtView
              debts={debts}
              summary={summary}
              inThousandsMode={inThousandsMode}
              onPayDebt={(debt) => {
                handleAddTransaction({
                  user: debt.user || (activeUserScope !== 'Semua' ? activeUserScope : 'Jeki'),
                  type: 'Debt',
                  category: debt.category,
                  title: `Bayar ${debt.name}`,
                  amount: Math.min(debt.monthlyInstallment, debt.remainingAmount),
                  debtId: debt.id,
                  debtAction: 'pay',
                  note: 'Pembayaran cicilan cepat'
                });
              }}
              onOpenNewDebt={() => setIsTxModalOpen(true)}
              theme={theme}
            />
          )}

          {activeTab === 'sync' && (
            <SyncAndExportView
              user={user}
              customVaultKey={customVaultKey}
              setCustomVaultKey={setCustomVaultKey}
              syncStatus={syncStatus}
              lastCloudSync={lastCloudSync}
              onGoogleSignIn={handleGoogleSignIn}
              onSignOut={handleSignOutToAnon}
              sheetsWebhookUrl={sheetsWebhookUrl}
              setSheetsWebhookUrl={(val) => {
                setSheetsWebhookUrl(val);
                persistStateToCloud({ sheetsWebhookUrl: val });
              }}
              sheetsAutoSync={sheetsAutoSync}
              setSheetsAutoSync={(val) => {
                setSheetsAutoSync(val);
                persistStateToCloud({ sheetsAutoSync: val });
              }}
              sheetsSyncing={sheetsSyncing}
              lastSheetsSyncTime={lastSheetsSyncTime}
              onPushSheetsNow={() => pushToGoogleSheets('full_sync')}
              onDownloadCSV={handleDownloadCSV}
              onGeneratePNG={handleGenerateReportPNG}
              showToast={showToast}
              darkMode={darkMode}
              theme={theme}
            />
          )}
        </main>

        {/* TOMBOL FLOATING TAMBAH TRANSAKSI */}
        <button
          onClick={() => setIsTxModalOpen(true)}
          aria-label="Tambah Transaksi"
          className="absolute bottom-[calc(5rem+env(safe-area-inset-bottom))] right-4 z-30 flex items-center gap-2 bg-gradient-to-r from-emerald-500 to-teal-600 hover:from-emerald-400 hover:to-teal-500 text-white font-semibold text-xs px-4 py-3.5 rounded-full shadow-lg shadow-emerald-600/35 active:scale-95 transition-all"
        >
          <Plus className="w-5 h-5" />
          <span>Catat</span>
        </button>

        {/* BOTTOM NAVIGATION BAR (5 TAB) */}
        <nav
          className={`shrink-0 border-t px-1.5 pt-2 pb-[calc(0.5rem+env(safe-area-inset-bottom))] flex items-center justify-around backdrop-blur-lg z-20 ${theme.navBg}`}
        >
          <NavButton
            active={activeTab === 'dashboard'}
            onClick={() => setActiveTab('dashboard')}
            icon={<Home className="w-4 h-4" />}
            label="Beranda"
            darkMode={darkMode}
          />
          <NavButton
            active={activeTab === 'transactions'}
            onClick={() => setActiveTab('transactions')}
            icon={<List className="w-4 h-4" />}
            label="Mutasi"
            darkMode={darkMode}
          />
          <NavButton
            active={activeTab === 'market'}
            onClick={() => setActiveTab('market')}
            icon={<LineChart className="w-4 h-4" />}
            label="Market"
            badge="LIVE"
            darkMode={darkMode}
          />
          <NavButton
            active={activeTab === 'debt'}
            onClick={() => setActiveTab('debt')}
            icon={<Landmark className="w-4 h-4" />}
            label="Hutang"
            darkMode={darkMode}
          />
          <NavButton
            active={activeTab === 'sync'}
            onClick={() => setActiveTab('sync')}
            icon={<FileSpreadsheet className="w-4 h-4" />}
            label="Sync & PNG"
            darkMode={darkMode}
          />
        </nav>

        {/* MODAL PENCATATAN TRANSAKSI */}
        {isTxModalOpen && (
          <TransactionModal
            onClose={() => setIsTxModalOpen(false)}
            onSubmit={handleAddTransaction}
            defaultUser={activeUserScope !== 'Semua' ? activeUserScope : 'Ririn'}
            assets={assets}
            debts={debts}
            darkMode={darkMode}
          />
        )}

        {/* MODAL BELI / JUAL ASET INVESTASI */}
        {tradeModal.open && tradeModal.asset && (
          <AssetTradeModal
            asset={tradeModal.asset}
            initialAction={tradeModal.action}
            defaultUser={activeUserScope !== 'Semua' ? activeUserScope : 'Ririn'}
            baseCashBalance={baseCashBalance}
            onClose={() => setTradeModal({ open: false, asset: null, action: 'buy' })}
            onConfirm={handleAssetTrade}
            darkMode={darkMode}
          />
        )}

        {/* MODAL PRATINJAU & UNDUH LAPORAN PNG */}
        {pngModal.open && (
          <PngReportModal
            dataUrl={pngModal.dataUrl}
            onClose={() => setPngModal({ open: false, dataUrl: '' })}
            showToast={showToast}
            darkMode={darkMode}
          />
        )}

        {/* MODAL KONFIRMASI RESET DATA */}
        {confirmResetModal && (
          <div className="absolute inset-0 z-50 bg-black/70 backdrop-blur-sm flex items-center justify-center p-4">
            <div
              className={`w-full max-w-xs rounded-3xl p-5 border ${
                darkMode ? 'bg-slate-900 border-slate-700' : 'bg-white border-slate-200'
              }`}
            >
              <div className="w-11 h-11 rounded-2xl bg-amber-500/15 text-amber-400 flex items-center justify-center mb-3">
                <AlertCircle className="w-6 h-6" />
              </div>
              <h3 className="font-bold text-base mb-1">Reset Data Demo?</h3>
              <p className={`text-xs mb-5 leading-relaxed ${theme.textMuted}`}>
                Semua transaksi baru, perubahan saldo, dan portofolio Emas/Valas/Saham akan dikembalikan ke data awal dan disinkronkan ke Cloud.
              </p>
              <div className="flex gap-2.5">
                <button
                  onClick={() => setConfirmResetModal(false)}
                  className={`flex-1 py-2.5 rounded-xl text-xs font-semibold border ${theme.border}`}
                >
                  Batal
                </button>
                <button
                  onClick={handleResetDemoData}
                  className="flex-1 py-2.5 rounded-xl text-xs font-semibold bg-amber-500 text-slate-950 hover:bg-amber-400"
                >
                  Ya, Reset
                </button>
              </div>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

// --- KOMPONEN NAVIGASI BAWAH ---
function NavButton({ active, onClick, icon, label, badge, darkMode }) {
  return (
    <button
      onClick={onClick}
      className={`relative flex flex-col items-center justify-center py-1.5 px-2 rounded-2xl transition-all ${
        active
          ? 'text-emerald-400 bg-emerald-500/10 font-semibold'
          : darkMode
          ? 'text-slate-400 hover:text-slate-200'
          : 'text-slate-500 hover:text-slate-800'
      }`}
    >
      {badge && (
        <span className="absolute -top-0.5 right-0.5 px-1 text-[7px] font-bold rounded bg-amber-500 text-slate-950">
          {badge}
        </span>
      )}
      {icon}
      <span className="text-[10px] mt-1">{label}</span>
    </button>
  );
}

// --- VIEW 1: DASHBOARD RINGKASAN ---
function DashboardView({
  summary,
  assets,
  transactions,
  activeUserScope,
  setActiveUserScope,
  inThousandsMode,
  darkMode,
  theme,
  isMarketLive,
  setIsMarketLive,
  onNavigate,
  onFilterMarket,
  onExportPNG
}) {
  const sefAsset = assets.find((a) => a.id === 'stock-sef-a');
  const sefCurrentVal = sefAsset ? sefAsset.holdings * sefAsset.price : 0;
  const sefCostVal = sefAsset ? sefAsset.holdings * sefAsset.avgBuyPrice : 0;
  const sefPnL = sefCurrentVal - sefCostVal;
  const sefPnLPct = sefCostVal > 0 ? (sefPnL / sefCostVal) * 100 : 0;

  return (
    <div className="space-y-4">
      {/* SWITCHER USER: GABUNGAN vs RIRIN vs JEKI */}
      <div className={`p-1.5 rounded-2xl border flex items-center gap-1 ${theme.card}`}>
        {[
          { id: 'Semua', label: 'Gabungan (Ririn & Jeki)', dot: 'bg-emerald-400' },
          { id: 'Ririn', label: 'Ririn', dot: 'bg-rose-400' },
          { id: 'Jeki', label: 'Jeki', dot: 'bg-sky-400' }
        ].map((u) => (
          <button
            key={u.id}
            onClick={() => setActiveUserScope(u.id)}
            className={`flex-1 py-2 px-2 rounded-xl text-[11px] font-bold flex items-center justify-center gap-1.5 transition ${
              activeUserScope === u.id
                ? u.id === 'Ririn'
                  ? 'bg-rose-500 text-white shadow-sm'
                  : u.id === 'Jeki'
                  ? 'bg-sky-500 text-slate-950 shadow-sm'
                  : 'bg-emerald-500 text-slate-950 shadow-sm'
                : theme.textMuted
            }`}
          >
            <span className={`w-2 h-2 rounded-full ${activeUserScope === u.id ? 'bg-white' : u.dot}`} />
            <span className="truncate">{u.label}</span>
          </button>
        ))}
      </div>

      {/* KARTU UTAMA NET WORTH & RUNNING BALANCE */}
      <div className="relative overflow-hidden rounded-3xl p-5 bg-gradient-to-br from-emerald-600 via-teal-700 to-slate-900 text-white shadow-xl shadow-emerald-950/30">
        <div className="absolute -right-8 -top-8 w-36 h-36 rounded-full bg-white/10 blur-2xl" />
        <div className="absolute -left-8 -bottom-10 w-32 h-32 rounded-full bg-amber-400/15 blur-xl" />

        <div className="relative z-10">
          <div className="flex items-center justify-between mb-1">
            <span className="text-xs font-medium text-emerald-100/90 flex items-center gap-1.5">
              <ShieldCheck className="w-4 h-4 text-emerald-300" />
              Total Kekayaan Bersih {inThousandsMode ? '(Dalam Ribuan)' : '(Net Worth)'}
            </span>
            <button
              onClick={onExportPNG}
              className="text-[10px] px-2.5 py-0.5 rounded-full bg-black/30 hover:bg-black/45 text-emerald-100 border border-white/15 flex items-center gap-1 transition"
            >
              <ImageIcon className="w-3 h-3" /> Simpan PNG
            </button>
          </div>

          <div className="text-2xl sm:text-3xl font-extrabold tracking-tight my-1.5">
            {formatIDR(summary.netWorth, false, inThousandsMode)}
          </div>

          <div className="flex flex-wrap items-center gap-2 text-[11px] text-emerald-100/90 mb-3.5">
            <span className="inline-flex items-center gap-0.5 px-2 py-0.5 rounded-lg bg-white/15 font-semibold">
              {summary.unrealizedPnL >= 0 ? (
                <ArrowUpRight className="w-3.5 h-3.5 text-emerald-300" />
              ) : (
                <ArrowDownRight className="w-3.5 h-3.5 text-rose-300" />
              )}
              {summary.unrealizedPnL >= 0 ? '+' : ''}
              {formatIDR(summary.unrealizedPnL, true, inThousandsMode)} ({summary.unrealizedPnLPercent.toFixed(2)}%)
            </span>
            <span className="opacity-80">imbal hasil aset</span>
          </div>

          {/* RUNNING BALANCE TOTAL KAS & TOTAL PORTOFOLIO */}
          <div className="grid grid-cols-2 gap-2.5 pt-3 border-t border-white/15 text-xs">
            <div
              onClick={() => onNavigate('transactions')}
              className="cursor-pointer hover:opacity-90"
            >
              <p className="text-[10px] text-emerald-200/85 flex items-center gap-1">
                <span>Running Balance Kas</span>
                <ChevronRight className="w-3 h-3" />
              </p>
              <p className="font-extrabold text-sm sm:text-base">
                {formatIDR(summary.baseCashBalance, false, inThousandsMode)}
              </p>
            </div>
            <div
              onClick={() => onNavigate('market')}
              className="cursor-pointer hover:opacity-90"
            >
              <p className="text-[10px] text-amber-200/90 flex items-center gap-1">
                <span>Total Nilai Investasi</span>
                <ChevronRight className="w-3 h-3" />
              </p>
              <p className="font-extrabold text-sm sm:text-base text-amber-300">
                {formatIDR(summary.portfolioValue, false, inThousandsMode)}
              </p>
            </div>
          </div>

          {/* BREAKDOWN SALDO KAS RIRIN (20,376.410) VS JEKI (2,875) */}
          <div className="grid grid-cols-2 gap-2 mt-3 pt-2.5 border-t border-white/15 text-[11px]">
            <div className="px-2.5 py-1.5 rounded-xl bg-rose-500/20 border border-rose-300/25 flex items-center justify-between">
              <span className="font-bold text-rose-200 flex items-center gap-1">
                <span className="w-2 h-2 rounded-full bg-rose-400" /> Ririn
              </span>
              <span className="font-bold text-white">
                {formatIDR(summary.byUser?.Ririn?.netFlow || 0, false, inThousandsMode)}
              </span>
            </div>
            <div className="px-2.5 py-1.5 rounded-xl bg-sky-500/20 border border-sky-300/25 flex items-center justify-between">
              <span className="font-bold text-sky-200 flex items-center gap-1">
                <span className="w-2 h-2 rounded-full bg-sky-400" /> Jeki
              </span>
              <span className="font-bold text-white">
                {formatIDR(summary.byUser?.Jeki?.netFlow || 0, false, inThousandsMode)}
              </span>
            </div>
          </div>
        </div>
      </div>

      {/* 4 KARTU RINGKASAN CASHFLOW */}
      <div className="grid grid-cols-2 gap-3">
        <div className={`p-3.5 rounded-2xl border ${theme.card}`}>
          <div className="flex items-center justify-between mb-2">
            <span className="w-8 h-8 rounded-xl bg-emerald-500/15 text-emerald-400 flex items-center justify-center">
              <TrendingUp className="w-4 h-4" />
            </span>
            <span className="text-[10px] font-semibold text-emerald-400 bg-emerald-500/10 px-2 py-0.5 rounded-full">
              Income ({activeUserScope})
            </span>
          </div>
          <p className={`text-[11px] ${theme.textMuted}`}>Saldo & Pemasukan</p>
          <p className="font-bold text-sm sm:text-base mt-0.5 text-emerald-400">
            {formatIDR(summary.totalIncome, false, inThousandsMode)}
          </p>
        </div>

        <div className={`p-3.5 rounded-2xl border ${theme.card}`}>
          <div className="flex items-center justify-between mb-2">
            <span className="w-8 h-8 rounded-xl bg-rose-500/15 text-rose-400 flex items-center justify-center">
              <TrendingDown className="w-4 h-4" />
            </span>
            <span className="text-[10px] font-semibold text-rose-400 bg-rose-500/10 px-2 py-0.5 rounded-full">
              Outcome ({activeUserScope})
            </span>
          </div>
          <p className={`text-[11px] ${theme.textMuted}`}>Pengeluaran & Cicilan</p>
          <p className="font-bold text-sm sm:text-base mt-0.5 text-rose-400">
            {formatIDR(summary.totalOutcome, false, inThousandsMode)}
          </p>
        </div>

        <div
          onClick={() => onNavigate('market')}
          className={`p-3.5 rounded-2xl border cursor-pointer hover:border-amber-500/50 transition ${theme.card}`}
        >
          <div className="flex items-center justify-between mb-2">
            <span className="w-8 h-8 rounded-xl bg-amber-500/15 text-amber-400 flex items-center justify-center">
              <PiggyBank className="w-4 h-4" />
            </span>
            <span className="text-[10px] font-semibold text-amber-400 bg-amber-500/10 px-2 py-0.5 rounded-full">
              Saving
            </span>
          </div>
          <p className={`text-[11px] ${theme.textMuted}`}>Total Kas & Aset</p>
          <p className="font-bold text-sm sm:text-base mt-0.5 text-amber-400">
            {formatIDR(summary.totalSavingAndInvestment, false, inThousandsMode)}
          </p>
        </div>

        <div
          onClick={() => onNavigate('debt')}
          className={`p-3.5 rounded-2xl border cursor-pointer hover:border-purple-500/50 transition ${theme.card}`}
        >
          <div className="flex items-center justify-between mb-2">
            <span className="w-8 h-8 rounded-xl bg-purple-500/15 text-purple-400 flex items-center justify-center">
              <CreditCard className="w-4 h-4" />
            </span>
            <span className="text-[10px] font-semibold text-purple-400 bg-purple-500/10 px-2 py-0.5 rounded-full">
              Debt
            </span>
          </div>
          <p className={`text-[11px] ${theme.textMuted}`}>Sisa Kewajiban Hutang</p>
          <p className="font-bold text-sm sm:text-base mt-0.5 text-purple-400">
            {formatIDR(summary.totalDebtRemaining, false, inThousandsMode)}
          </p>
        </div>
      </div>

      {/* LIVE TICKER HARGA EMAS, DEPOSITO SEABANK/BCA & SUCORINVEST EQUITY FUND */}
      <div className={`p-4 rounded-3xl border ${theme.card}`}>
        <div className="flex items-center justify-between mb-3">
          <div>
            <h2 className="font-bold text-sm flex items-center gap-1.5">
              <Sparkles className="w-4 h-4 text-amber-400" />
              Portofolio Real & Harga Live
            </h2>
            <p className={`text-[11px] ${theme.textMuted}`}>
              Emas LM & Perhiasan, Deposito SeaBank & Sucorinvest
            </p>
          </div>
          <div className="flex items-center gap-1.5">
            <button
              onClick={() => setIsMarketLive(!isMarketLive)}
              className={`px-2.5 py-1 rounded-xl text-[11px] font-semibold flex items-center gap-1 ${
                isMarketLive
                  ? 'bg-emerald-500/15 text-emerald-400'
                  : 'bg-slate-500/20 text-slate-400'
              }`}
            >
              {isMarketLive ? <Pause className="w-3 h-3" /> : <Play className="w-3 h-3" />}
              {isMarketLive ? 'Live' : 'Paused'}
            </button>
            <button
              onClick={() => onNavigate('market')}
              className="text-xs text-emerald-400 font-semibold flex items-center"
            >
              Semua <ChevronRight className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>

        {/* HIGHLIGHT KARTU SUCORINVEST EQUITY FUND KELAS A (MODAL 15,900.0 rb) */}
        {sefAsset && (
          <div
            onClick={() => onFilterMarket('Saham')}
            className={`p-3 rounded-2xl border border-blue-500/30 mb-3 cursor-pointer ${theme.cardSubtle}`}
          >
            <div className="flex items-center justify-between mb-1">
              <span className="text-[11px] font-bold text-blue-400 flex items-center gap-1.5">
                <Briefcase className="w-3.5 h-3.5" />
                Sucorinvest Equity Fund Kelas A (SEF-A)
              </span>
              <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-blue-500/15 text-blue-400">
                NAV {formatUnitIDR(sefAsset.price)}
              </span>
            </div>
            <div className="flex items-center justify-between text-xs mt-1.5">
              <div>
                <p className={`text-[10px] ${theme.textMuted}`}>
                  {formatUnitsCount(sefAsset.holdings)} unit • Modal: {formatIDR(sefCostVal, false, inThousandsMode)}
                </p>
                <p className="font-extrabold text-sm mt-0.5">
                  {formatIDR(sefCurrentVal, false, inThousandsMode)}
                </p>
              </div>
              <div className="text-right">
                <p className={`text-[10px] ${theme.textMuted}`}>
                  Beli: {formatUnitIDR(sefAsset.avgBuyPrice)}
                </p>
                <p
                  className={`font-bold text-xs mt-0.5 ${
                    sefPnL >= 0 ? 'text-emerald-400' : 'text-rose-400'
                  }`}
                >
                  {sefPnL >= 0 ? '+' : ''}
                  {formatIDR(sefPnL, true, inThousandsMode)} ({sefPnLPct.toFixed(2)}%)
                </p>
              </div>
            </div>
          </div>
        )}

        {/* TICKER HORIZONTAL EMAS (ANTAM 17g, SEMAR 3g, PERHIASAN) & DEPOSITO (25,000.0 rb) */}
        <div className="flex gap-2 overflow-x-auto pb-1 no-scrollbar mb-3">
          {[
            assets.find((a) => a.id === 'gold-antam'),
            assets.find((a) => a.id === 'gold-semar'),
            assets.find((a) => a.id === 'dep-seabank'),
            assets.find((a) => a.id === 'gold-kalung'),
            assets.find((a) => a.id === 'gold-cincin'),
            assets.find((a) => a.id === 'gold-anting'),
            assets.find((a) => a.id === 'gold-gelang')
          ]
            .filter(Boolean)
            .map((asset) => {
              const diff = asset.price - asset.prevPrice;
              const isUp = diff >= 0;
              const valTotal = asset.holdings * asset.multiplier * asset.price;
              return (
                <div
                  key={asset.id}
                  onClick={() => onNavigate('market')}
                  className={`shrink-0 px-3 py-2 rounded-2xl border cursor-pointer min-w-[148px] ${theme.border} ${theme.cardSubtle}`}
                >
                  <div className="flex items-center justify-between text-[11px] font-bold">
                    <span className="truncate max-w-[110px]">
                      {asset.name.replace('(10g + 5g + 2g)', '').replace('(6.0% p.a)', '')}
                    </span>
                    <span className={isUp ? 'text-emerald-400' : 'text-rose-400'}>
                      {isUp ? '▲' : '▼'}
                    </span>
                  </div>
                  <div className="text-xs font-extrabold mt-0.5 text-amber-400">
                    {formatIDR(valTotal, false, inThousandsMode)}
                  </div>
                  <div className={`text-[10px] ${theme.textMuted}`}>
                    {asset.type === 'Deposito'
                      ? 'Deposito Aktif'
                      : `${formatUnitsCount(asset.holdings)} gr • ${formatUnitIDR(asset.price, inThousandsMode)}/gr`}
                  </div>
                </div>
              );
            })}
        </div>

        <div className="grid grid-cols-2 gap-2">
          <div
            onClick={() => onFilterMarket('Emas')}
            className={`p-2.5 rounded-2xl cursor-pointer hover:ring-1 hover:ring-amber-400/50 transition ${theme.cardSubtle}`}
          >
            <div className="flex items-center justify-between text-[10px] font-semibold text-amber-400 mb-1">
              <span className="flex items-center gap-1">
                <Coins className="w-3 h-3" /> Emas & Perhiasan (30.1g)
              </span>
            </div>
            <p className="font-bold text-xs sm:text-sm">
              {formatIDR(summary.portfolioByType['Emas'] || 0, false, inThousandsMode)}
            </p>
          </div>

          <div
            onClick={() => onFilterMarket('Deposito')}
            className={`p-2.5 rounded-2xl cursor-pointer hover:ring-1 hover:ring-orange-400/50 transition ${theme.cardSubtle}`}
          >
            <div className="flex items-center justify-between text-[10px] font-semibold text-orange-400 mb-1">
              <span className="flex items-center gap-1">
                <Landmark className="w-3 h-3" /> Deposito SeaBank/BCA
              </span>
            </div>
            <p className="font-bold text-xs sm:text-sm">
              {formatIDR(summary.portfolioByType['Deposito'] || 0, false, inThousandsMode)}
            </p>
          </div>

          <div
            onClick={() => onFilterMarket('Saham')}
            className={`p-2.5 rounded-2xl cursor-pointer hover:ring-1 hover:ring-blue-400/50 transition ${theme.cardSubtle}`}
          >
            <div className="flex items-center justify-between text-[10px] font-semibold text-blue-400 mb-1">
              <span className="flex items-center gap-1">
                <Briefcase className="w-3 h-3" /> Saham Sucorinvest
              </span>
            </div>
            <p className="font-bold text-xs sm:text-sm">
              {formatIDR(summary.portfolioByType['Saham'] || 0, false, inThousandsMode)}
            </p>
          </div>

          <div
            onClick={() => onFilterMarket('Valas')}
            className={`p-2.5 rounded-2xl cursor-pointer hover:ring-1 hover:ring-cyan-400/50 transition ${theme.cardSubtle}`}
          >
            <div className="flex items-center justify-between text-[10px] font-semibold text-cyan-400 mb-1">
              <span className="flex items-center gap-1">
                <Globe className="w-3 h-3" /> Valas & RDPU
              </span>
            </div>
            <p className="font-bold text-xs sm:text-sm">
              {formatIDR(
                (summary.portfolioByType['Valas'] || 0) +
                  (summary.portfolioByType['Reksadana Pasar Uang'] || 0),
                false,
                inThousandsMode
              )}
            </p>
          </div>
        </div>
      </div>

      {/* GRAFIK TREN INCOME VS OUTCOME */}
      <div className={`p-4 rounded-3xl border ${theme.card}`}>
        <div className="flex items-center justify-between mb-3">
          <div>
            <h2 className="font-bold text-sm">Tren Arus Kas (6 Bulan)</h2>
            <p className={`text-[11px] ${theme.textMuted}`}>
              Perbandingan Pemasukan vs Pengeluaran
            </p>
          </div>
          <div className="flex items-center gap-3 text-[10px]">
            <span className="flex items-center gap-1 text-emerald-400">
              <span className="w-2 h-2 rounded-full bg-emerald-400" /> Income
            </span>
            <span className="flex items-center gap-1 text-rose-400">
              <span className="w-2 h-2 rounded-full bg-rose-400" /> Outcome
            </span>
          </div>
        </div>

        <div className="h-44 w-full">
          <ResponsiveContainer width="100%" height="100%">
            <AreaChart data={CASHFLOW_TREND_DATA} margin={{ top: 5, right: 5, left: -22, bottom: 0 }}>
              <defs>
                <linearGradient id="incGrad" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#10B981" stopOpacity={0.35} />
                  <stop offset="95%" stopColor="#10B981" stopOpacity={0} />
                </linearGradient>
                <linearGradient id="outGrad" x1="0" y1="0" x2="0" y2="1">
                  <stop offset="5%" stopColor="#F43F5E" stopOpacity={0.3} />
                  <stop offset="95%" stopColor="#F43F5E" stopOpacity={0} />
                </linearGradient>
              </defs>
              <XAxis
                dataKey="bulan"
                tick={{ fontSize: 10, fill: darkMode ? '#94A3B8' : '#64748B' }}
                axisLine={false}
                tickLine={false}
              />
              <YAxis
                tickFormatter={(v) => `${v / 1000000}Jt`}
                tick={{ fontSize: 10, fill: darkMode ? '#94A3B8' : '#64748B' }}
                axisLine={false}
                tickLine={false}
              />
              <Tooltip
                formatter={(value) => formatIDR(value, false, inThousandsMode)}
                contentStyle={{
                  backgroundColor: darkMode ? '#0F172A' : '#FFFFFF',
                  borderColor: darkMode ? '#334155' : '#E2E8F0',
                  borderRadius: '12px',
                  fontSize: '11px'
                }}
              />
              <Area
                type="monotone"
                dataKey="income"
                name="Pemasukan"
                stroke="#10B981"
                strokeWidth={2.5}
                fillOpacity={1}
                fill="url(#incGrad)"
              />
              <Area
                type="monotone"
                dataKey="outcome"
                name="Pengeluaran"
                stroke="#F43F5E"
                strokeWidth={2.5}
                fillOpacity={1}
                fill="url(#outGrad)"
              />
            </AreaChart>
          </ResponsiveContainer>
        </div>
      </div>

      {/* TRANSAKSI TERBARU DENGAN RUNNING BALANCE & USER BADGE */}
      <div className={`p-4 rounded-3xl border ${theme.card}`}>
        <div className="flex items-center justify-between mb-3">
          <div>
            <h2 className="font-bold text-sm">Catatan Aset & Running Balance</h2>
            <p className={`text-[10px] ${theme.textMuted}`}>
              Menampilkan sumber user (Ririn / Jeki) & nominal {inThousandsMode ? 'dalam ribuan' : 'IDR'}
            </p>
          </div>
          <button
            onClick={() => onNavigate('transactions')}
            className="text-xs font-semibold text-emerald-400"
          >
            Lihat Semua
          </button>
        </div>
        <div className="space-y-2.5">
          {transactions
            .filter((tx) => activeUserScope === 'Semua' || tx.user === activeUserScope)
            .slice(0, 5)
            .map((tx) => (
              <TransactionItem
                key={tx.id}
                tx={tx}
                inThousandsMode={inThousandsMode}
                theme={theme}
                compact
              />
            ))}
        </div>
      </div>
    </div>
  );
}

// --- VIEW 2: RIWAYAT TRANSAKSI & RUNNING BALANCE REKENING KORAN ---
function TransactionsView({
  transactions,
  baseCashBalance,
  txFilter,
  setTxFilter,
  txUserFilter,
  setTxUserFilter,
  inThousandsMode,
  onDelete,
  onOpenTxModal,
  theme
}) {
  const filters = ['Semua', 'Income', 'Outcome', 'Saving', 'Debt'];
  const userFilters = ['Semua', 'Ririn', 'Jeki'];

  const filteredList = useMemo(() => {
    return transactions.filter((t) => {
      const matchType = txFilter === 'Semua' || t.type === txFilter;
      const matchUser = txUserFilter === 'Semua' || (t.user || 'Ririn') === txUserFilter;
      return matchType && matchUser;
    });
  }, [transactions, txFilter, txUserFilter]);

  return (
    <div className="space-y-4">
      {/* BANNER RUNNING BALANCE TOTAL */}
      <div className="rounded-3xl p-4 bg-gradient-to-r from-slate-800 via-slate-900 to-emerald-950/80 border border-emerald-500/25 text-white flex items-center justify-between">
        <div>
          <p className="text-[11px] text-emerald-300 font-semibold">
            Running Balance Total (Saldo Kas Aktif)
          </p>
          <p className="text-xl font-extrabold mt-0.5">
            {formatIDR(baseCashBalance, false, inThousandsMode)}
          </p>
          <p className="text-[10px] text-slate-400 mt-0.5">
            {filteredList.length} catatan ditampilkan • Format: {inThousandsMode ? 'Ribuan (rb)' : 'Rupiah (IDR)'}
          </p>
        </div>
        <button
          onClick={onOpenTxModal}
          className="px-3.5 py-2.5 rounded-2xl bg-emerald-500 hover:bg-emerald-400 text-slate-950 text-xs font-bold flex items-center gap-1 shrink-0"
        >
          <Plus className="w-4 h-4" /> Catat
        </button>
      </div>

      {/* FILTER USER (SEMUA / RIRIN / JEKI) */}
      <div className="flex items-center gap-1.5">
        {userFilters.map((u) => (
          <button
            key={u}
            onClick={() => setTxUserFilter(u)}
            className={`flex-1 py-2 rounded-xl text-xs font-bold border transition ${
              txUserFilter === u
                ? u === 'Ririn'
                  ? 'bg-rose-500 border-rose-500 text-white'
                  : u === 'Jeki'
                  ? 'bg-sky-500 border-sky-500 text-slate-950'
                  : 'bg-emerald-500 border-emerald-500 text-slate-950'
                : `${theme.card} ${theme.textMuted}`
            }`}
          >
            {u === 'Semua' ? 'Semua User' : `User: ${u}`}
          </button>
        ))}
      </div>

      {/* FILTER TIPE TRANSAKSI */}
      <div className="flex gap-1.5 overflow-x-auto pb-1 no-scrollbar">
        {filters.map((f) => (
          <button
            key={f}
            onClick={() => setTxFilter(f)}
            className={`px-3.5 py-1.5 rounded-full text-xs font-semibold transition shrink-0 ${
              txFilter === f
                ? 'bg-amber-400 text-slate-950 shadow-sm'
                : `${theme.card} border ${theme.textMuted}`
            }`}
          >
            {f}
          </button>
        ))}
      </div>

      <div className="space-y-2.5">
        {filteredList.length === 0 ? (
          <div className={`p-8 rounded-3xl border text-center ${theme.card}`}>
            <p className={`text-xs ${theme.textMuted}`}>
              Tidak ada transaksi untuk filter yang dipilih.
            </p>
          </div>
        ) : (
          filteredList.map((tx) => (
            <TransactionItem
              key={tx.id}
              tx={tx}
              inThousandsMode={inThousandsMode}
              theme={theme}
              onDelete={() => onDelete(tx.id)}
            />
          ))
        )}
      </div>
    </div>
  );
}

// --- KOMPONEN ITEM TRANSAKSI + BADGE USER (RIRIN/JEKI) + RUNNING BALANCE ---
function TransactionItem({ tx, inThousandsMode, theme, onDelete, compact = false }) {
  const badgeConfig = {
    Income: {
      icon: <ArrowUpRight className="w-4 h-4" />,
      bg: 'bg-emerald-500/15 text-emerald-400',
      amountColor: 'text-emerald-400',
      prefix: '+'
    },
    Outcome: {
      icon: <ArrowDownRight className="w-4 h-4" />,
      bg: 'bg-rose-500/15 text-rose-400',
      amountColor: 'text-rose-400',
      prefix: '-'
    },
    Saving: {
      icon: <PiggyBank className="w-4 h-4" />,
      bg: 'bg-amber-500/15 text-amber-400',
      amountColor: 'text-amber-400',
      prefix: tx.id?.startsWith('tx-real-') ? '' : tx.category === 'Tabungan Tunai/Bank' ? '+' : '-'
    },
    Debt: {
      icon: <CreditCard className="w-4 h-4" />,
      bg: 'bg-purple-500/15 text-purple-400',
      amountColor: 'text-purple-400',
      prefix: tx.debtAction === 'pay' ? '-' : '+'
    }
  };

  const cfg = badgeConfig[tx.type] || badgeConfig.Outcome;
  const actor = tx.user || 'Ririn';
  const isRirin = actor === 'Ririn';

  return (
    <div
      className={`p-3 rounded-2xl border flex items-center justify-between gap-2.5 ${theme.card}`}
    >
      <div className="flex items-center gap-3 min-w-0">
        <div
          className={`w-10 h-10 rounded-xl flex items-center justify-center shrink-0 ${cfg.bg}`}
        >
          {cfg.icon}
        </div>
        <div className="min-w-0">
          <div className="flex items-center gap-1.5 mb-0.5">
            <span
              className={`px-1.5 py-0.5 rounded-md text-[9px] font-extrabold uppercase tracking-wide ${
                isRirin
                  ? 'bg-rose-500/20 text-rose-400 border border-rose-500/30'
                  : 'bg-sky-500/20 text-sky-400 border border-sky-500/30'
              }`}
            >
              {actor}
            </span>
            <p className="font-semibold text-xs sm:text-sm truncate">{tx.title}</p>
          </div>
          <div className={`flex items-center gap-1.5 text-[10px] ${theme.textMuted}`}>
            <span className="font-medium">{tx.category}</span>
            <span>•</span>
            <span>{tx.date}</span>
          </div>
          {!compact && tx.note && (
            <p className={`text-[10px] mt-0.5 truncate ${theme.textMuted}`}>{tx.note}</p>
          )}
        </div>
      </div>

      <div className="flex items-center gap-2 shrink-0">
        <div className="text-right">
          <p className={`font-bold text-xs sm:text-sm ${cfg.amountColor}`}>
            {cfg.prefix}
            {formatIDR(tx.amount, false, inThousandsMode)}
          </p>
          {typeof tx.runningBalance === 'number' && (
            <p className={`text-[10px] font-semibold mt-0.5 ${theme.textMuted}`}>
              Bal: {formatIDR(tx.runningBalance, compact, inThousandsMode)}
            </p>
          )}
        </div>
        {onDelete && (
          <button
            onClick={onDelete}
            aria-label="Hapus transaksi"
            className="p-1.5 rounded-lg text-slate-500 hover:text-rose-400 hover:bg-rose-500/10 transition"
          >
            <Trash2 className="w-3.5 h-3.5" />
          </button>
        )}
      </div>
    </div>
  );
}

// --- VIEW 3: INVESTASI & LIVE MARKET (EMAS, DEPOSITO, VALAS, SAHAM, RDPU) ---
function MarketView({
  assets,
  summary,
  inThousandsMode,
  isMarketLive,
  setIsMarketLive,
  onManualRefresh,
  lastMarketUpdate,
  marketFilter,
  setMarketFilter,
  onOpenTrade,
  theme
}) {
  const categories = ['Semua', 'Emas', 'Deposito', 'Saham', 'Valas', 'Reksadana Pasar Uang'];

  const filteredAssets = useMemo(() => {
    if (marketFilter === 'Semua') return assets;
    return assets.filter((a) => a.type === marketFilter);
  }, [assets, marketFilter]);

  return (
    <div className="space-y-4">
      <div className="rounded-3xl p-5 bg-gradient-to-br from-slate-800 via-slate-900 to-amber-950/60 border border-amber-500/25 text-white shadow-lg">
        <div className="flex items-center justify-between mb-1">
          <span className="text-xs text-amber-300 font-semibold flex items-center gap-1.5">
            <Coins className="w-4 h-4" /> Total Portofolio Emas, Deposito & Saham
          </span>
          <div className="flex items-center gap-1.5">
            <button
              onClick={() => setIsMarketLive(!isMarketLive)}
              className={`px-2.5 py-1 rounded-full text-[10px] font-bold flex items-center gap-1 ${
                isMarketLive
                  ? 'bg-emerald-500 text-slate-950'
                  : 'bg-slate-700 text-slate-200'
              }`}
            >
              {isMarketLive ? <Pause className="w-3 h-3" /> : <Play className="w-3 h-3" />}
              {isMarketLive ? 'LIVE FEED' : 'PAUSED'}
            </button>
            <button
              onClick={onManualRefresh}
              title="Refresh Harga Sekarang"
              className="p-1.5 rounded-full bg-white/10 hover:bg-white/20 transition"
            >
              <RefreshCw className="w-3.5 h-3.5" />
            </button>
          </div>
        </div>

        <div className="text-2xl font-extrabold my-1">
          {formatIDR(summary.portfolioValue, false, inThousandsMode)}
        </div>

        <div className="flex items-center justify-between text-[11px] mt-2 pt-2.5 border-t border-white/10">
          <span className="text-slate-300">Unrealized Gain/Loss:</span>
          <span
            className={`font-bold ${
              summary.unrealizedPnL >= 0 ? 'text-emerald-400' : 'text-rose-400'
            }`}
          >
            {summary.unrealizedPnL >= 0 ? '+' : ''}
            {formatIDR(summary.unrealizedPnL, false, inThousandsMode)} ({summary.unrealizedPnLPercent.toFixed(2)}%)
          </span>
        </div>
        <p className="text-[10px] text-slate-400 mt-1.5">
          Update pasar terakhir: {lastMarketUpdate.toLocaleTimeString('id-ID')} WIB
        </p>
      </div>

      <div className="flex gap-1.5 overflow-x-auto pb-1 no-scrollbar">
        {categories.map((cat) => (
          <button
            key={cat}
            onClick={() => setMarketFilter(cat)}
            className={`px-3.5 py-1.5 rounded-full text-xs font-semibold shrink-0 transition ${
              marketFilter === cat
                ? 'bg-amber-400 text-slate-950 shadow-sm'
                : `${theme.card} border ${theme.textMuted}`
            }`}
          >
            {cat === 'Reksadana Pasar Uang' ? 'RDPU' : cat}
          </button>
        ))}
      </div>

      <div className="space-y-3">
        {filteredAssets.map((asset) => {
          const tickDiff = asset.price - asset.prevPrice;
          const tickPct = asset.prevPrice > 0 ? (tickDiff / asset.prevPrice) * 100 : 0;
          const isTickUp = tickDiff >= 0;

          const currentVal = asset.holdings * asset.multiplier * asset.price;
          const costVal = asset.holdings * asset.multiplier * asset.avgBuyPrice;
          const pnl = currentVal - costVal;
          const pnlPct = costVal > 0 ? (pnl / costVal) * 100 : 0;

          return (
            <div
              key={asset.id}
              className={`p-4 rounded-3xl border transition-all ${theme.card}`}
            >
              <div className="flex items-start justify-between gap-2 mb-2.5">
                <div>
                  <div className="flex items-center gap-2">
                    <span
                      className="px-2 py-0.5 rounded-lg text-[10px] font-bold text-white"
                      style={{ backgroundColor: asset.color }}
                    >
                      {asset.symbol}
                    </span>
                    <span className={`text-[10px] font-medium ${theme.textMuted}`}>
                      {asset.type}
                    </span>
                  </div>
                  <h3 className="font-bold text-sm mt-1">{asset.name}</h3>
                </div>

                <div className="text-right">
                  <div className="font-extrabold text-sm">
                    {formatUnitIDR(asset.price, inThousandsMode)}
                    <span className={`text-[10px] font-normal ml-0.5 ${theme.textMuted}`}>
                      /{asset.type === 'Saham' ? 'lbr' : asset.unitLabel}
                    </span>
                  </div>
                  <div
                    className={`inline-flex items-center gap-0.5 text-[10px] font-bold px-1.5 py-0.5 rounded-md ${
                      isTickUp
                        ? 'bg-emerald-500/15 text-emerald-400'
                        : 'bg-rose-500/15 text-rose-400'
                    }`}
                  >
                    {isTickUp ? '▲' : '▼'} {Math.abs(tickPct).toFixed(2)}%
                  </div>
                </div>
              </div>

              <div
                className={`grid grid-cols-2 gap-2 p-3 rounded-2xl text-xs mb-3 ${theme.cardSubtle}`}
              >
                <div>
                  <p className={`text-[10px] ${theme.textMuted}`}>
                    Kepemilikan ({asset.unitLabel})
                  </p>
                  <p className="font-bold">
                    {formatUnitsCount(asset.holdings)} {asset.unitLabel}
                    {asset.multiplier > 1 && (
                      <span className={`text-[10px] font-normal ml-1 ${theme.textMuted}`}>
                        ({formatUnitsCount(asset.holdings * asset.multiplier)} lbr)
                      </span>
                    )}
                  </p>
                  <p className={`text-[10px] mt-0.5 ${theme.textMuted}`}>
                    Modal: {formatIDR(costVal, false, inThousandsMode)}
                  </p>
                </div>

                <div className="text-right">
                  <p className={`text-[10px] ${theme.textMuted}`}>Nilai Saat Ini</p>
                  <p className="font-bold text-sm">
                    {formatIDR(currentVal, false, inThousandsMode)}
                  </p>
                  <p
                    className={`text-[10px] font-semibold mt-0.5 ${
                      pnl >= 0 ? 'text-emerald-400' : 'text-rose-400'
                    }`}
                  >
                    {pnl >= 0 ? '+' : ''}
                    {formatIDR(pnl, true, inThousandsMode)} ({pnlPct.toFixed(1)}%)
                  </p>
                </div>
              </div>

              <div className="flex gap-2">
                <button
                  onClick={() => onOpenTrade(asset, 'buy')}
                  className="flex-1 py-2 rounded-xl bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-bold text-xs transition"
                >
                  + Tambah {asset.symbol}
                </button>
                <button
                  onClick={() => onOpenTrade(asset, 'sell')}
                  disabled={asset.holdings <= 0}
                  className={`flex-1 py-2 rounded-xl border font-bold text-xs transition ${
                    asset.holdings > 0
                      ? 'border-rose-500/40 text-rose-400 hover:bg-rose-500/10'
                      : 'opacity-40 cursor-not-allowed border-slate-600 text-slate-500'
                  }`}
                >
                  Cairkan / Jual
                </button>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
}

// --- VIEW 4: MANAJEMEN HUTANG (DEBT) ---
function DebtView({ debts, summary, inThousandsMode, onPayDebt, onOpenNewDebt, theme }) {
  return (
    <div className="space-y-4">
      <div className="rounded-3xl p-5 bg-gradient-to-br from-purple-700 via-indigo-800 to-slate-900 text-white shadow-lg">
        <div className="flex items-center justify-between mb-1">
          <span className="text-xs text-purple-200 font-semibold">
            Total Sisa Kewajiban Hutang
          </span>
          <button
            onClick={onOpenNewDebt}
            className="px-2.5 py-1 rounded-xl bg-white/15 hover:bg-white/25 text-[11px] font-semibold"
          >
            + Catat Hutang
          </button>
        </div>
        <div className="text-2xl font-extrabold my-1">
          {formatIDR(summary.totalDebtRemaining, false, inThousandsMode)}
        </div>
        <p className="text-[11px] text-purple-200/80">
          {debts.length === 0
            ? 'Luar biasa! Saat ini Anda bebas dari kewajiban hutang.'
            : `Dari total plafon awal ${formatIDR(summary.totalDebtInitial, false, inThousandsMode)}`}
        </p>
      </div>

      {debts.length === 0 ? (
        <div className={`p-8 rounded-3xl border text-center ${theme.card}`}>
          <ShieldCheck className="w-10 h-10 text-emerald-400 mx-auto mb-2" />
          <p className="font-bold text-sm">Zero Debt (Bebas Hutang)</p>
          <p className={`text-xs mt-1 ${theme.textMuted}`}>
            Klik &quot;+ Catat Hutang&quot; di atas jika suatu saat ingin mencatat cicilan atau pinjaman.
          </p>
        </div>
      ) : (
        <div className="space-y-3">
          {debts.map((debt) => {
            const paid = debt.totalAmount - debt.remainingAmount;
            const pct =
              debt.totalAmount > 0 ? Math.round((paid / debt.totalAmount) * 100) : 100;
            const isPaidOff = debt.remainingAmount <= 0;

            return (
              <div key={debt.id} className={`p-4 rounded-3xl border ${theme.card}`}>
                <div className="flex items-start justify-between mb-2">
                  <div>
                    <span className="text-[10px] font-semibold px-2 py-0.5 rounded-full bg-purple-500/15 text-purple-400">
                      {debt.category}
                    </span>
                    <h3 className="font-bold text-sm mt-1.5">{debt.name}</h3>
                    <p className={`text-[11px] ${theme.textMuted}`}>
                      Jatuh tempo: {debt.dueDate}
                    </p>
                  </div>
                  <div className="text-right">
                    <p className={`text-[10px] ${theme.textMuted}`}>Sisa Hutang</p>
                    <p className="font-extrabold text-sm text-purple-400">
                      {formatIDR(debt.remainingAmount, false, inThousandsMode)}
                    </p>
                  </div>
                </div>

                <div className="my-3">
                  <div className="flex justify-between text-[11px] mb-1">
                    <span className={theme.textMuted}>
                      Terbayar: {formatIDR(paid, true, inThousandsMode)}
                    </span>
                    <span className="font-bold text-emerald-400">{pct}%</span>
                  </div>
                  <div className="w-full h-2 rounded-full bg-slate-700/40 overflow-hidden">
                    <div
                      className="h-full rounded-full bg-gradient-to-r from-purple-500 to-emerald-400 transition-all duration-500"
                      style={{ width: `${pct}%` }}
                    />
                  </div>
                </div>

                <div className="flex items-center justify-between pt-2 border-t border-slate-700/30">
                  <div className="text-xs">
                    <span className={theme.textMuted}>Cicilan/bln: </span>
                    <span className="font-bold">
                      {formatIDR(debt.monthlyInstallment, false, inThousandsMode)}
                    </span>
                  </div>

                  {isPaidOff ? (
                    <span className="px-3 py-1.5 rounded-xl bg-emerald-500/15 text-emerald-400 text-xs font-bold">
                      ✓ Lunas
                    </span>
                  ) : (
                    <button
                      onClick={() => onPayDebt(debt)}
                      className="px-3.5 py-2 rounded-xl bg-purple-500 hover:bg-purple-400 text-white font-bold text-xs transition active:scale-95"
                    >
                      Bayar Cicilan
                    </button>
                  )}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// --- VIEW 5: SYNC CLOUD, GOOGLE SHEETS & EXPORT PNG ---
function SyncAndExportView({
  user,
  customVaultKey,
  setCustomVaultKey,
  syncStatus,
  lastCloudSync,
  onGoogleSignIn,
  onSignOut,
  sheetsWebhookUrl,
  setSheetsWebhookUrl,
  sheetsAutoSync,
  setSheetsAutoSync,
  sheetsSyncing,
  lastSheetsSyncTime,
  onPushSheetsNow,
  onDownloadCSV,
  onGeneratePNG,
  showToast,
  darkMode,
  theme
}) {
  const [vaultDraft, setVaultDraft] = useState(customVaultKey);
  const [urlDraft, setUrlDraft] = useState(sheetsWebhookUrl);
  const [copiedScript, setCopiedScript] = useState(false);
  const [showGuide, setShowGuide] = useState(false);

  const handleApplyVaultKey = (e) => {
    e.preventDefault();
    setCustomVaultKey(vaultDraft.trim());
    showToast(
      vaultDraft.trim()
        ? `Terhubung ke Ruang Cloud "${vaultDraft.trim()}"!`
        : 'Kembali ke ruang akun pribadi'
    );
  };

  const handleSaveSheetsUrl = (e) => {
    e.preventDefault();
    setSheetsWebhookUrl(urlDraft.trim());
    showToast('URL Google Sheets Web App disimpan!');
  };

  const handleCopyScript = () => {
    navigator.clipboard.writeText(GOOGLE_APPS_SCRIPT_TEMPLATE);
    setCopiedScript(true);
    showToast('Kode Google Apps Script berhasil disalin!');
    setTimeout(() => setCopiedScript(false), 2500);
  };

  const isGoogleConnected = user && !user.isAnonymous && user.email;

  return (
    <div className="space-y-4">
      {/* 1. KARTU EXPORT LAPORAN KEUANGAN PNG & CSV */}
      <div className="rounded-3xl p-5 bg-gradient-to-br from-emerald-600 via-teal-700 to-slate-900 text-white shadow-lg">
        <div className="flex items-center gap-2 mb-1.5">
          <ImageIcon className="w-5 h-5 text-emerald-200" />
          <h2 className="font-bold text-base">Laporan Keuangan Siap Bagikan</h2>
        </div>
        <p className="text-xs text-emerald-100/85 mb-4 leading-relaxed">
          Cetak ringkasan Net Worth, Cashflow, dan Portofolio Live ke gambar PNG beresolusi tinggi atau unduh file Spreadsheet.
        </p>

        <div className="grid grid-cols-2 gap-2.5">
          <button
            onClick={onGeneratePNG}
            className="py-3 px-3 rounded-2xl bg-white text-slate-950 font-bold text-xs flex items-center justify-center gap-1.5 shadow-md active:scale-95 transition"
          >
            <ImageIcon className="w-4 h-4 text-emerald-600" />
            <span>Export Laporan PNG</span>
          </button>

          <button
            onClick={onDownloadCSV}
            className="py-3 px-3 rounded-2xl bg-black/30 hover:bg-black/45 text-white border border-white/20 font-bold text-xs flex items-center justify-center gap-1.5 active:scale-95 transition"
          >
            <Download className="w-4 h-4 text-amber-300" />
            <span>Download CSV</span>
          </button>
        </div>
      </div>

      {/* 2. SINKRONISASI CLOUD LINTAS BROWSER & DEVICE */}
      <div className={`p-4 rounded-3xl border space-y-3 ${theme.card}`}>
        <div className="flex items-start justify-between">
          <div>
            <h3 className="font-bold text-sm flex items-center gap-1.5">
              <Smartphone className="w-4 h-4 text-emerald-400" />
              Sinkronisasi Multi-Device & Browser
            </h3>
            <p className={`text-[11px] mt-0.5 ${theme.textMuted}`}>
              Akses dashboard dari HP/Laptop lain tanpa kehilangan data
            </p>
          </div>
          <span
            className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${
              syncStatus === 'synced'
                ? 'bg-emerald-500/15 text-emerald-400'
                : 'bg-amber-500/15 text-amber-400'
            }`}
          >
            {syncStatus === 'synced' ? '✓ Tersinkron' : 'Menghubungkan...'}
          </span>
        </div>

        {/* Opsi A: Login Google */}
        <div className={`p-3 rounded-2xl ${theme.cardSubtle}`}>
          {isGoogleConnected ? (
            <div className="flex items-center justify-between gap-2">
              <div className="min-w-0">
                <p className="text-[10px] text-emerald-400 font-bold">
                  TERHUBUNG DENGAN AKUN GOOGLE
                </p>
                <p className="text-xs font-semibold truncate">{user.email}</p>
              </div>
              <button
                onClick={onSignOut}
                className="px-2.5 py-1.5 rounded-xl bg-rose-500/15 text-rose-400 text-[11px] font-semibold flex items-center gap-1 shrink-0"
              >
                <LogOut className="w-3.5 h-3.5" /> Keluar
              </button>
            </div>
          ) : (
            <div className="space-y-2">
              <p className={`text-[11px] ${theme.textMuted}`}>
                Hubungkan akun Google Anda agar data otomatis sinkron saat login di perangkat lain:
              </p>
              <button
                onClick={onGoogleSignIn}
                className="w-full py-2.5 px-3 rounded-xl bg-white text-slate-900 font-bold text-xs flex items-center justify-center gap-2 shadow-sm hover:bg-slate-100 transition"
              >
                <LogIn className="w-4 h-4 text-blue-600" />
                <span>Hubungkan Akun Google</span>
              </button>
            </div>
          )}
        </div>

        {/* Opsi B: Kode Sinkronisasi Cloud (Sync Vault Key) */}
        <form onSubmit={handleApplyVaultKey} className="space-y-2 pt-1">
          <label className="block text-[11px] font-semibold">
            Kode Sinkronisasi Lintas Browser (Sync Key)
          </label>
          <p className={`text-[10px] leading-relaxed ${theme.textMuted}`}>
            Masukkan kode unik pribadi (min. 4 karakter), lalu masukkan kode yang sama saat membuka aplikasi di browser atau HP lain.
          </p>
          <div className="flex gap-2">
            <input
              type="text"
              value={vaultDraft}
              onChange={(e) => setVaultDraft(e.target.value)}
              placeholder="Contoh: dompet-jeki-2026"
              className={`flex-1 px-3 py-2 rounded-xl border text-xs font-semibold outline-none ${
                darkMode
                  ? 'bg-slate-900 border-slate-700'
                  : 'bg-slate-50 border-slate-300'
              }`}
            />
            <button
              type="submit"
              className="px-3.5 py-2 rounded-xl bg-emerald-500 text-slate-950 font-bold text-xs shrink-0"
            >
              Hubungkan
            </button>
          </div>
          {lastCloudSync && (
            <p className={`text-[10px] ${theme.textMuted}`}>
              Sinkronisasi cloud terakhir: {lastCloudSync.toLocaleTimeString('id-ID')} WIB
            </p>
          )}
        </form>
      </div>

      {/* 3. INTEGRASI & AUTO-SYNC GOOGLE SHEETS */}
      <div className={`p-4 rounded-3xl border space-y-3 ${theme.card}`}>
        <div className="flex items-center justify-between">
          <div>
            <h3 className="font-bold text-sm flex items-center gap-1.5">
              <FileSpreadsheet className="w-4 h-4 text-emerald-400" />
              Integrasi Google Sheets Real-Time
            </h3>
            <p className={`text-[11px] ${theme.textMuted}`}>
              Kirim transaksi, portofolio & hutang langsung ke Google Sheet
            </p>
          </div>
        </div>

        <form onSubmit={handleSaveSheetsUrl} className="space-y-2.5">
          <div>
            <label className="block text-[11px] font-semibold mb-1">
              URL Google Apps Script Web App
            </label>
            <input
              type="url"
              value={urlDraft}
              onChange={(e) => setUrlDraft(e.target.value)}
              placeholder="https://script.google.com/macros/s/.../exec"
              className={`w-full px-3 py-2.5 rounded-xl border text-xs outline-none ${
                darkMode
                  ? 'bg-slate-900 border-slate-700'
                  : 'bg-slate-50 border-slate-300'
              }`}
            />
          </div>

          <div className="flex items-center justify-between py-1">
            <span className="text-xs font-medium">
              Auto-Sync setiap transaksi baru
            </span>
            <button
              type="button"
              onClick={() => setSheetsAutoSync(!sheetsAutoSync)}
              className={`w-11 h-6 rounded-full transition p-0.5 ${
                sheetsAutoSync ? 'bg-emerald-500' : 'bg-slate-600'
              }`}
            >
              <div
                className={`w-5 h-5 rounded-full bg-white transition-transform ${
                  sheetsAutoSync ? 'translate-x-5' : 'translate-x-0'
                }`}
              />
            </button>
          </div>

          <div className="flex gap-2">
            <button
              type="submit"
              className={`flex-1 py-2.5 rounded-xl border font-bold text-xs ${theme.border}`}
            >
              Simpan URL
            </button>
            <button
              type="button"
              onClick={onPushSheetsNow}
              disabled={sheetsSyncing}
              className="flex-1 py-2.5 rounded-xl bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-bold text-xs flex items-center justify-center gap-1.5 disabled:opacity-50"
            >
              <RefreshCw className={`w-3.5 h-3.5 ${sheetsSyncing ? 'animate-spin' : ''}`} />
              <span>{sheetsSyncing ? 'Mengirim...' : 'Sync Sekarang'}</span>
            </button>
          </div>
          {lastSheetsSyncTime && (
            <p className="text-[10px] text-emerald-400">
              ✓ Terkirim ke Google Sheets pukul {lastSheetsSyncTime.toLocaleTimeString('id-ID')} WIB
            </p>
          )}
        </form>

        {/* PANDUAN & SALIN KODE APPS SCRIPT */}
        <div className={`pt-3 border-t ${theme.border}`}>
          <button
            type="button"
            onClick={() => setShowGuide(!showGuide)}
            className="w-full flex items-center justify-between text-xs font-semibold text-amber-400"
          >
            <span>Panduan Setup Google Sheets (1 Menit)</span>
            <ChevronRight
              className={`w-4 h-4 transition-transform ${showGuide ? 'rotate-90' : ''}`}
            />
          </button>

          {showGuide && (
            <div className="mt-3 space-y-2.5 text-[11px]">
              <ol className={`list-decimal pl-4 space-y-1.5 ${theme.textMuted}`}>
                <li>
                  Buka <strong>Google Sheets</strong> baru di browser Anda.
                </li>
                <li>
                  Klik menu <strong>Extensions (Ekstensi) &gt; Apps Script</strong>.
                </li>
                <li>
                  Hapus kode default, lalu klik tombol <strong>Salin Kode Script</strong> di bawah dan tempelkan ke editor Apps Script.
                </li>
                <li>
                  Klik <strong>Deploy &gt; New deployment &gt; Web app</strong>. Set <em>Who has access</em> ke <strong>Anyone (Siapa saja)</strong>, lalu salin URL <code>.../exec</code> ke kolom di atas.
                </li>
              </ol>

              <button
                type="button"
                onClick={handleCopyScript}
                className="w-full py-2.5 px-3 rounded-xl bg-amber-400 hover:bg-amber-300 text-slate-950 font-bold text-xs flex items-center justify-center gap-1.5 transition"
              >
                {copiedScript ? (
                  <>
                    <Check className="w-4 h-4" />
                    <span>Kode Apps Script Tersalin!</span>
                  </>
                ) : (
                  <>
                    <Copy className="w-4 h-4" />
                    <span>Salin Kode Google Apps Script</span>
                  </>
                )}
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

// --- MODAL PRATINJAU & UNDUH GAMBAR LAPORAN PNG ---
function PngReportModal({ dataUrl, onClose, showToast, darkMode }) {
  const handleDownload = () => {
    const link = document.createElement('a');
    link.href = dataUrl;
    link.download = `Laporan_Keuangan_DompetKu_${
      new Date().toISOString().split('T')[0]
    }.png`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    showToast('Gambar Laporan PNG berhasil diunduh!');
  };

  const handleShareImage = async () => {
    try {
      const res = await fetch(dataUrl);
      const blob = await res.blob();
      const file = new File([blob], 'Laporan_DompetKu.png', { type: 'image/png' });
      if (navigator.canShare && navigator.canShare({ files: [file] })) {
        await navigator.share({
          title: 'Laporan Keuangan DompetKu Live',
          text: 'Ringkasan Kekayaan Bersih & Cashflow Pribadi',
          files: [file]
        });
      } else {
        handleDownload();
      }
    } catch {
      handleDownload();
    }
  };

  return (
    <div className="absolute inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4">
      <div
        className={`w-full max-w-sm rounded-3xl p-4 border flex flex-col max-h-[90%] ${
          darkMode
            ? 'bg-slate-900 border-slate-700 text-slate-100'
            : 'bg-white border-slate-200 text-slate-900'
        }`}
      >
        <div className="flex items-center justify-between mb-3">
          <div>
            <h3 className="font-bold text-sm">Pratinjau Laporan PNG</h3>
            <p className="text-[10px] opacity-70">Resolusi Tinggi (1080 × 1540 px)</p>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-full bg-slate-500/15 hover:bg-slate-500/25"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        <div className="flex-1 overflow-y-auto rounded-2xl border border-slate-700/60 bg-slate-950 p-2 mb-3">
          {dataUrl && (
            // eslint-disable-next-line @next/next/no-img-element
            <img
              src={dataUrl}
              alt="Pratinjau Laporan Keuangan"
              className="w-full h-auto rounded-xl shadow-md"
            />
          )}
        </div>

        <div className="grid grid-cols-2 gap-2.5">
          <button
            onClick={handleDownload}
            className="py-3 rounded-2xl bg-emerald-500 hover:bg-emerald-400 text-slate-950 font-bold text-xs flex items-center justify-center gap-1.5 transition"
          >
            <Download className="w-4 h-4" />
            <span>Unduh PNG</span>
          </button>
          <button
            onClick={handleShareImage}
            className="py-3 rounded-2xl bg-blue-500 hover:bg-blue-400 text-white font-bold text-xs flex items-center justify-center gap-1.5 transition"
          >
            <Share2 className="w-4 h-4" />
            <span>Bagikan</span>
          </button>
        </div>
      </div>
    </div>
  );
}
// --- MODAL PENCATATAN TRANSAKSI BARU (DENGAN PILIHAN USER RIRIN / JEKI) ---
function TransactionModal({ onClose, onSubmit, defaultUser = 'Ririn', assets, debts, darkMode }) {
  const [txUser, setTxUser] = useState(defaultUser);
  const [type, setType] = useState('Outcome');
  const [category, setCategory] = useState(CATEGORIES.Outcome[0]);
  const [title, setTitle] = useState('');
  const [amount, setAmount] = useState('');
  const [date, setDate] = useState(new Date().toISOString().split('T')[0]);
  const [note, setNote] = useState('');

  const [debtAction, setDebtAction] = useState('pay');
  const [selectedDebtId, setSelectedDebtId] = useState(debts[0]?.id || '');

  const [selectedAssetId, setSelectedAssetId] = useState('');
  const [units, setUnits] = useState('');

  const handleTypeChange = (newType) => {
    setType(newType);
    const defaultCat = CATEGORIES[newType][0];
    setCategory(defaultCat);
    setSelectedAssetId('');
    setUnits('');
    setAmount('');
  };

  const handleCategoryChange = (newCat) => {
    setCategory(newCat);
    if (type === 'Saving' && newCat !== 'Tabungan Tunai/Bank') {
      const matchingAsset = assets.find((a) => a.type === newCat);
      if (matchingAsset) {
        setSelectedAssetId(matchingAsset.id);
        const defaultUnits = matchingAsset.symbol === 'VND' ? 1000000 : 1;
        setUnits(String(defaultUnits));
        setAmount(
          String(Math.round(defaultUnits * matchingAsset.multiplier * matchingAsset.price))
        );
      }
    } else {
      setSelectedAssetId('');
      setUnits('');
    }
  };

  const handleUnitsChange = (val) => {
    setUnits(val);
    const asset = assets.find((a) => a.id === selectedAssetId);
    if (asset && Number(val) > 0) {
      setAmount(String(Math.round(Number(val) * asset.multiplier * asset.price)));
    } else {
      setAmount('');
    }
  };

  const handleAmountChange = (val) => {
    setAmount(val);
    if (type === 'Saving' && category !== 'Tabungan Tunai/Bank' && selectedAssetId) {
      const asset = assets.find((a) => a.id === selectedAssetId);
      const numRp = Number(val);
      if (asset && numRp > 0) {
        const computedUnits = numRp / (asset.multiplier * asset.price);
        setUnits(String(Number(computedUnits.toFixed(4))));
      } else {
        setUnits('');
      }
    }
  };

  const handleAssetSelect = (assetId) => {
    setSelectedAssetId(assetId);
    const asset = assets.find((a) => a.id === assetId);
    if (asset) {
      if (Number(amount) > 0) {
        const computedUnits = Number(amount) / (asset.multiplier * asset.price);
        setUnits(String(Number(computedUnits.toFixed(4))));
      } else if (Number(units) > 0) {
        setAmount(String(Math.round(Number(units) * asset.multiplier * asset.price)));
      }
    }
  };

  const handleFormSubmit = (e) => {
    e.preventDefault();
    const numericAmount = Number(amount);
    if (!numericAmount || numericAmount <= 0) return;

    const chosenAsset = assets.find((a) => a.id === selectedAssetId);

    const finalTitle =
      title.trim() ||
      (type === 'Debt' && debtAction === 'pay'
        ? `Bayar ${debts.find((d) => d.id === selectedDebtId)?.name || 'Cicilan'}`
        : chosenAsset
        ? `Beli ${chosenAsset.symbol} (${formatUnitsCount(Number(units))} ${chosenAsset.unitLabel})`
        : `${category}`);

    onSubmit({
      user: txUser,
      type,
      category,
      title: finalTitle,
      amount: numericAmount,
      date,
      note,
      debtAction: type === 'Debt' ? debtAction : undefined,
      debtId: type === 'Debt' && debtAction === 'pay' ? selectedDebtId : undefined,
      assetId: selectedAssetId || undefined,
      units: units ? Number(units) : undefined
    });
  };

  const filteredAssetsByCat = assets.filter((a) => a.type === category);
  const currentSelectedAsset = assets.find((a) => a.id === selectedAssetId);

  return (
    <div className="absolute inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-end sm:items-center justify-center">
      <div
        className={`w-full max-h-[92%] overflow-y-auto rounded-t-3xl sm:rounded-3xl p-5 border ${
          darkMode
            ? 'bg-slate-900 border-slate-700 text-slate-100'
            : 'bg-white border-slate-200 text-slate-900'
        }`}
      >
        <div className="flex items-center justify-between mb-3.5">
          <h3 className="font-bold text-base">Catat Cashflow Baru</h3>
          <button
            onClick={onClose}
            className="p-1.5 rounded-full bg-slate-500/15 hover:bg-slate-500/25"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* PILIH SUMBER USER: RIRIN ATAU JEKI */}
        <div className="mb-3.5">
          <label className="block text-[11px] font-semibold mb-1.5 opacity-80">
            Sumber User / Pelaku Transaksi
          </label>
          <div className="grid grid-cols-2 gap-2">
            {['Ririn', 'Jeki'].map((person) => (
              <button
                key={person}
                type="button"
                onClick={() => setTxUser(person)}
                className={`py-2.5 rounded-xl text-xs font-extrabold border flex items-center justify-center gap-2 transition ${
                  txUser === person
                    ? person === 'Ririn'
                      ? 'bg-rose-500 border-rose-500 text-white shadow-md shadow-rose-500/25'
                      : 'bg-sky-500 border-sky-500 text-slate-950 shadow-md shadow-sky-500/25'
                    : 'border-slate-600/50 opacity-70'
                }`}
              >
                <span
                  className={`w-2.5 h-2.5 rounded-full ${
                    txUser === person
                      ? 'bg-white'
                      : person === 'Ririn'
                      ? 'bg-rose-400'
                      : 'bg-sky-400'
                  }`}
                />
                <span>{person}</span>
              </button>
            ))}
          </div>
        </div>

        <div className="grid grid-cols-4 gap-1.5 mb-4 p-1 rounded-2xl bg-slate-500/10">
          {['Income', 'Outcome', 'Saving', 'Debt'].map((t) => (
            <button
              key={t}
              type="button"
              onClick={() => handleTypeChange(t)}
              className={`py-2 rounded-xl text-xs font-bold transition ${
                type === t
                  ? t === 'Income'
                    ? 'bg-emerald-500 text-slate-950'
                    : t === 'Outcome'
                    ? 'bg-rose-500 text-white'
                    : t === 'Saving'
                    ? 'bg-amber-400 text-slate-950'
                    : 'bg-purple-500 text-white'
                  : 'opacity-70'
              }`}
            >
              {t}
            </button>
          ))}
        </div>

        <form onSubmit={handleFormSubmit} className="space-y-3.5 text-xs">
          <div>
            <label className="block font-semibold mb-1 opacity-80">
              {type === 'Saving' ? 'Tipe Tabungan / Investasi' : 'Kategori'}
            </label>
            <select
              value={category}
              onChange={(e) => handleCategoryChange(e.target.value)}
              className={`w-full px-3.5 py-2.5 rounded-xl border outline-none ${
                darkMode
                  ? 'bg-slate-800 border-slate-700'
                  : 'bg-slate-50 border-slate-300'
              }`}
            >
              {CATEGORIES[type].map((c) => (
                <option key={c} value={c}>
                  {c}
                </option>
              ))}
            </select>
          </div>

          {type === 'Saving' && category !== 'Tabungan Tunai/Bank' && (
            <div className="p-3.5 rounded-2xl bg-amber-500/10 border border-amber-500/30 space-y-2.5">
              <div>
                <label className="block font-semibold mb-1 text-amber-400">
                  Pilih Instrumen {category} (Harga / Kurs Live)
                </label>
                <select
                  value={selectedAssetId}
                  onChange={(e) => handleAssetSelect(e.target.value)}
                  className={`w-full px-3 py-2 rounded-xl border outline-none ${
                    darkMode
                      ? 'bg-slate-800 border-slate-700'
                      : 'bg-white border-slate-300'
                  }`}
                >
                  {filteredAssetsByCat.map((a) => (
                    <option key={a.id} value={a.id}>
                      {a.symbol} - {a.name} ({formatUnitIDR(a.price)}/
                      {a.type === 'Saham' ? 'lbr' : a.unitLabel})
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block font-semibold mb-1 text-amber-400">
                  Jumlah ({currentSelectedAsset?.unitLabel || 'unit'}) — Otomatis terhitung dari Rp
                </label>
                <input
                  type="number"
                  step="any"
                  min="0.0001"
                  value={units}
                  onChange={(e) => handleUnitsChange(e.target.value)}
                  placeholder="Contoh: 1.5"
                  className={`w-full px-3 py-2 rounded-xl border font-semibold outline-none ${
                    darkMode
                      ? 'bg-slate-800 border-slate-700'
                      : 'bg-white border-slate-300'
                  }`}
                />
              </div>
            </div>
          )}

          {type === 'Debt' && (
            <div className="p-3 rounded-2xl bg-purple-500/10 border border-purple-500/30 space-y-2.5">
              <div className="flex gap-2">
                <button
                  type="button"
                  onClick={() => setDebtAction('pay')}
                  className={`flex-1 py-1.5 rounded-lg font-semibold ${
                    debtAction === 'pay' ? 'bg-purple-500 text-white' : 'bg-slate-500/20'
                  }`}
                >
                  Bayar Cicilan
                </button>
                <button
                  type="button"
                  onClick={() => setDebtAction('new')}
                  className={`flex-1 py-1.5 rounded-lg font-semibold ${
                    debtAction === 'new' ? 'bg-purple-500 text-white' : 'bg-slate-500/20'
                  }`}
                >
                  Hutang Baru
                </button>
              </div>

              {debtAction === 'pay' && debts.length > 0 && (
                <div>
                  <label className="block font-semibold mb-1 text-purple-300">
                    Pilih Pinjaman yang Dibayar
                  </label>
                  <select
                    value={selectedDebtId}
                    onChange={(e) => setSelectedDebtId(e.target.value)}
                    className={`w-full px-3 py-2 rounded-xl border outline-none ${
                      darkMode
                        ? 'bg-slate-800 border-slate-700'
                        : 'bg-white border-slate-300'
                    }`}
                  >
                    {debts.map((d) => (
                      <option key={d.id} value={d.id}>
                        {d.name} (Sisa: {formatIDR(d.remainingAmount, true)})
                      </option>
                    ))}
                  </select>
                </div>
              )}
            </div>
          )}

          <div>
            <label className="block font-semibold mb-1 opacity-80">Nominal (Rp)</label>
            <input
              type="number"
              required
              min="1"
              value={amount}
              onChange={(e) => handleAmountChange(e.target.value)}
              placeholder="Contoh: 1500000"
              className={`w-full px-3.5 py-2.5 rounded-xl border font-bold text-sm outline-none ${
                darkMode
                  ? 'bg-slate-800 border-slate-700'
                  : 'bg-slate-50 border-slate-300'
              }`}
            />
          </div>

          <div>
            <label className="block font-semibold mb-1 opacity-80">
              Judul / Keterangan (Opsional)
            </label>
            <input
              type="text"
              value={title}
              onChange={(e) => setTitle(e.target.value)}
              placeholder="Otomatis terisi jika dikosongkan"
              className={`w-full px-3.5 py-2.5 rounded-xl border outline-none ${
                darkMode
                  ? 'bg-slate-800 border-slate-700'
                  : 'bg-slate-50 border-slate-300'
              }`}
            />
          </div>

          <div className="grid grid-cols-2 gap-2.5">
            <div>
              <label className="block font-semibold mb-1 opacity-80">Tanggal</label>
              <input
                type="date"
                value={date}
                onChange={(e) => setDate(e.target.value)}
                className={`w-full px-3 py-2 rounded-xl border outline-none ${
                  darkMode
                    ? 'bg-slate-800 border-slate-700'
                    : 'bg-slate-50 border-slate-300'
                }`}
              />
            </div>
            <div>
              <label className="block font-semibold mb-1 opacity-80">Catatan</label>
              <input
                type="text"
                value={note}
                onChange={(e) => setNote(e.target.value)}
                placeholder="Catatan..."
                className={`w-full px-3 py-2 rounded-xl border outline-none ${
                  darkMode
                    ? 'bg-slate-800 border-slate-700'
                    : 'bg-slate-50 border-slate-300'
                }`}
              />
            </div>
          </div>

          <button
            type="submit"
            className="w-full py-3.5 rounded-2xl bg-gradient-to-r from-emerald-500 to-teal-600 text-white font-bold text-sm shadow-lg shadow-emerald-600/25 active:scale-98 transition mt-2"
          >
            Simpan ({txUser}) & Sinkronkan
          </button>
        </form>
      </div>
    </div>
  );
}

// --- MODAL EKSEKUSI BELI / JUAL ASET ---
function AssetTradeModal({
  asset,
  initialAction,
  defaultUser = 'Ririn',
  baseCashBalance,
  onClose,
  onConfirm,
  darkMode
}) {
  const [tradeUser, setTradeUser] = useState(defaultUser);
  const [action, setAction] = useState(initialAction);
  const [units, setUnits] = useState(asset.symbol === 'VND' ? '1000000' : '1');

  const estimatedTotal = Math.round(
    (Number(units) || 0) * asset.multiplier * asset.price
  );

  return (
    <div className="absolute inset-0 z-50 bg-black/75 backdrop-blur-sm flex items-end sm:items-center justify-center p-4">
      <div
        className={`w-full max-w-sm rounded-3xl p-5 border ${
          darkMode
            ? 'bg-slate-900 border-slate-700 text-slate-100'
            : 'bg-white border-slate-200 text-slate-900'
        }`}
      >
        <div className="flex items-center justify-between mb-3">
          <div>
            <span className="text-[10px] font-bold px-2 py-0.5 rounded bg-amber-400 text-slate-950">
              {asset.symbol} • {asset.type}
            </span>
            <h3 className="font-bold text-base mt-1">{asset.name}</h3>
          </div>
          <button
            onClick={onClose}
            className="p-1.5 rounded-full bg-slate-500/15 hover:bg-slate-500/25"
          >
            <X className="w-4 h-4" />
          </button>
        </div>

        {/* PILIH SUMBER USER: RIRIN ATAU JEKI */}
        <div className="grid grid-cols-2 gap-1.5 mb-3">
          {['Ririn', 'Jeki'].map((person) => (
            <button
              key={person}
              type="button"
              onClick={() => setTradeUser(person)}
              className={`py-2 rounded-xl text-xs font-extrabold border transition ${
                tradeUser === person
                  ? person === 'Ririn'
                    ? 'bg-rose-500 border-rose-500 text-white'
                    : 'bg-sky-500 border-sky-500 text-slate-950'
                  : 'border-slate-600/50 opacity-70'
              }`}
            >
              User: {person}
            </button>
          ))}
        </div>

        <div className="grid grid-cols-2 gap-1.5 p-1 rounded-2xl bg-slate-500/10 mb-4">
          <button
            type="button"
            onClick={() => setAction('buy')}
            className={`py-2 rounded-xl text-xs font-bold ${
              action === 'buy' ? 'bg-emerald-500 text-slate-950' : 'opacity-70'
            }`}
          >
            Beli {asset.symbol}
          </button>
          <button
            type="button"
            onClick={() => setAction('sell')}
            className={`py-2 rounded-xl text-xs font-bold ${
              action === 'sell' ? 'bg-rose-500 text-white' : 'opacity-70'
            }`}
          >
            Jual {asset.symbol}
          </button>
        </div>

        <div className="space-y-3 text-xs">
          <div className="flex justify-between">
            <span className="opacity-70">Harga / Kurs Live:</span>
            <span className="font-bold text-amber-400">
              {formatUnitIDR(asset.price)} / {asset.type === 'Saham' ? 'lembar' : asset.unitLabel}
            </span>
          </div>

          <div className="flex justify-between">
            <span className="opacity-70">
              {action === 'buy' ? 'Saldo Kas Tersedia:' : 'Kepemilikan Anda:'}
            </span>
            <span className="font-semibold">
              {action === 'buy'
                ? formatIDR(baseCashBalance)
                : `${formatUnitsCount(asset.holdings)} ${asset.unitLabel}`}
            </span>
          </div>

          <div>
            <label className="block font-semibold mb-1 opacity-80">
              Jumlah ({asset.unitLabel})
            </label>
            <input
              type="number"
              step="any"
              min="0.0001"
              max={action === 'sell' ? asset.holdings : undefined}
              value={units}
              onChange={(e) => setUnits(e.target.value)}
              className={`w-full px-3.5 py-2.5 rounded-xl border font-bold text-sm outline-none ${
                darkMode
                  ? 'bg-slate-800 border-slate-700'
                  : 'bg-slate-50 border-slate-300'
              }`}
            />
          </div>

          <div className="p-3 rounded-2xl bg-slate-500/10 flex items-center justify-between">
            <span className="opacity-80">Estimasi Nilai (IDR):</span>
            <span className="font-extrabold text-sm text-emerald-400">
              {formatIDR(estimatedTotal)}
            </span>
          </div>

          <button
            type="button"
            onClick={() => onConfirm({ asset, action, units, tradeUser })}
            disabled={Number(units) <= 0 || (action === 'sell' && Number(units) > asset.holdings)}
            className={`w-full py-3 rounded-2xl font-bold text-sm transition ${
              action === 'buy'
                ? 'bg-emerald-500 hover:bg-emerald-400 text-slate-950'
                : 'bg-rose-500 hover:bg-rose-400 text-white'
            } disabled:opacity-40`}
          >
            Konfirmasi {action === 'buy' ? 'Pembelian' : 'Penjualan'} ({tradeUser})
          </button>
        </div>
      </div>
    </div>
  );
}