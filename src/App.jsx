// ============================================================
// Portal Administrasi SE2026 — BPS Kota Jakarta Timur
// Dependencies: npm install xlsx docxtemplater pizzip file-saver docx-preview
// ============================================================

import React, { useState, useRef, useCallback } from "react";
import { motion, AnimatePresence } from "framer-motion";
import {
  FileText, Users, ClipboardList, Car, Receipt, Briefcase,
  Map as MapIcon, ChevronRight, X, Printer, ArrowLeft, Check, Plus,
  Trash2, Menu, LayoutDashboard, Upload, Download, Filter,
  AlertCircle, CheckCircle, MapPin, LoaderCircle,
} from "lucide-react";
import * as XLSX from "xlsx";
import PizZip from "pizzip";
import JSZip from "jszip";
import Docxtemplater from "docxtemplater";
import { saveAs } from "file-saver";
import { renderAsync } from "docx-preview";

// ─── DATA ────────────────────────────────────────────────────────────────────

const DOC_TYPES = [
  { id: "daftar-hadir",  icon: <ClipboardList />, label: "Daftar Hadir", desc: "Fitur dikunci. Saat ini hanya Lampiran yang aktif.", color: "orange", disabled: true, lockedMessage: "Fitur Daftar Hadir dikunci. Saat ini hanya Lampiran yang aktif." },
  { id: "tanda-terima",  icon: <Briefcase />,     label: "Tanda Terima", desc: "Fitur dikunci. Saat ini hanya Lampiran yang aktif.", color: "amber", disabled: true, lockedMessage: "Fitur Tanda Terima dikunci. Saat ini hanya Lampiran yang aktif." },
  { id: "surat-pernyataan-kendaraan", icon: <Car />, label: "Super Kendis", desc: "Fitur dikunci. Saat ini hanya Lampiran yang aktif.", color: "orange", disabled: true, lockedMessage: "Fitur Super Kendis dikunci. Saat ini hanya Lampiran yang aktif." },
  { id: "pengeluaran-riil", icon: <Receipt />,    label: "DPR", desc: "Fitur dikunci. Saat ini hanya Lampiran yang aktif.", color: "amber", disabled: true, lockedMessage: "Fitur DPR dikunci. Saat ini hanya Lampiran yang aktif." },
  { id: "spj",           icon: <FileText />,      label: "SPJ", desc: "Fitur dikunci. Saat ini hanya Lampiran yang aktif.", color: "orange", disabled: true, lockedMessage: "Fitur SPJ dikunci. Saat ini hanya Lampiran yang aktif." },
  { id: "spd",           icon: <MapIcon />,           label: "SPD", desc: "Fitur dikunci. Saat ini hanya Lampiran yang aktif.", color: "amber", disabled: true, lockedMessage: "Fitur SPD dikunci. Saat ini hanya Lampiran yang aktif." },
  { id: "surat-tugas",   icon: <Users />,         label: "Surtug", desc: "Fitur dikunci. Saat ini hanya Lampiran yang aktif.", color: "orange", disabled: true, lockedMessage: "Fitur Surat Tugas dikunci. Saat ini hanya Lampiran yang aktif." },
  { id: "bapp",          icon: <FileText />,      label: "BAPP", desc: "BAPP PML/PPL", color: "amber" },
  { id: "bast",          icon: <FileText />,      label: "BAST", desc: "BAST PML/PPL", color: "amber" },
  { id: "surat-pernyataan-penyelesaian-lapangan", icon: <FileText />, label: "Surat Pernyataan Penyelesaian Lapangan", desc: "Khusus PML", color: "amber" },
  { id: "lampiran",      icon: <FileText />,      label: "Lampiran", desc: "Lampiran SPK PML/PPL", color: "amber" },
];

// ─── XLSX PARSER ─────────────────────────────────────────────────────────────

function normalizeRowHeaders(row) {
  const normalized = {};
  Object.entries(row).forEach(([k, v]) => {
    const key = String(k ?? "").trim().toLowerCase().replace(/\s+/g, " ");
    normalized[key] = String(v ?? "").trim();
  });
  return {
    no:           normalized["no"] ?? "",
    nama:         normalized["nama"] ?? normalized["nama lengkap"] ?? normalized["nama_lengkap"] ?? normalized["nama-lengkap"] ?? "",
    nik:          normalized["nik"] ?? normalized["nip"] ?? "",
    asal:         normalized["asal"] ?? "",
    wilTugas:     (normalized["wil. tugas"] ?? normalized["wil tugas"] ?? normalized["wil.tugas"] ?? normalized["wilayah tugas"] ?? "").toUpperCase(),
    jabatan:      (normalized["jabatan"] ?? normalized["posisi"] ?? "").toUpperCase(),
    pangkatGol:   normalized["pangkat/gol"] ?? normalized["pangkat gol"] ?? normalized["pangkatgol"] ?? "",
    kelas:        String(normalized["kelas"] ?? "").trim(),
    hotel:        (normalized["tc"] ?? normalized["hotel"] ?? "").toUpperCase(),
    gelombang:    String(normalized["gelombang"] ?? "").trim(),
    tc:           (normalized["tc"] ?? "").toUpperCase(),
    sobatId:      normalized["sobat id"] ?? normalized["sobatid"] ?? "",
    email:        normalized["email"] ?? "",
    jenisKelamin: normalized["jenis kelamin"] ?? normalized["jeniskelamin"] ?? "",
  };
}

function parseXlsxData(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  
  // Cari sheet yang berisi data administrasi secara otomatis
  const adminSheetInfo = findAdministrasiSheet(workbook);
  if (!adminSheetInfo) {
    console.warn("Sheet data administrasi tidak ditemukan. Mencoba sheet pertama...");
    // Fallback ke sheet pertama jika tidak ditemukan
    const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
    if (!firstSheet) throw new Error("File XLSX kosong");
    const raw = XLSX.utils.sheet_to_json(firstSheet, { defval: "", raw: false });
    return raw.map(normalizeRowHeaders).filter(r => r.nama !== "" || r.nik !== "" || r.sobatId !== "");
  }
  
  const raw = adminSheetInfo.data;
  return raw.map(normalizeRowHeaders).filter(r => r.nama !== "" || r.nik !== "" || r.sobatId !== "");
}

function normalizeBappRow(row) {
  const normalized = {};
  Object.entries(row).forEach(([k, v]) => {
    const key = String(k ?? "").trim().toLowerCase().replace(/\s+/g, " ");
    normalized[key] = String(v ?? "").trim();
  });

  const get = (...keys) => {
    for (const key of keys) {
      const value = normalized[key];
      if (value != null && String(value).trim() !== "") return String(value).trim();
    }
    return "";
  };

  return {
    no: get("no", "nomor"),
    nama: get("nama", "nama lengkap", "nama_lengkap", "nama petugas", "nama peserta", "nama_petugas"),
    jabatan: get("jabatan", "posisi", "jenis petugas", "role", "kategori"),
    jabatan_raw: get("jabatan", "posisi", "jenis petugas", "role", "kategori"),
    wilayah: get("wilayah", "wil tugas", "wil. tugas", "wilayah tugas", "kecamatan", "asal"),
    email: get("email", "email petugas", "email peserta", "mail"),
    kelas: get("kelas"),
    gelombang: get("gelombang"),
    tempat: get("tempat", "hotel", "tc"),
    telp: get("telp", "no hp", "nomor hp"),
    nama_pml: get("nama pml", "pengawas"),
    nama_ppl: get("nama ppl", "pencacah"),
    nik: get("nik", "nik petugas", "nik peserta"),
    sls_40: get("sls 40%", "sls 40", "sls_40", "sls40", "sls 40 persen"),
    nomor_spk: get("nomor spk", "nomor_spk", "nomor spk", "nomor kontrak", "nomor_kontrak", "spk"),
    nomor_kontrak: get("nomor spk", "nomor_spk", "nomor spk", "nomor kontrak", "nomor_kontrak", "spk"),
  };
}

function parseBappData(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames.find(
    (name) => String(name ?? "").trim().toLowerCase() === "pembayaran"
  );

  if (!sheetName) {
    console.warn("Sheet bernama 'Pembayaran' tidak ditemukan; BAPP tidak akan memuat data.");
    return [];
  }

  const sheet = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  return raw.map(normalizeBappRow).filter((r) => r.nama || r.email || r.jabatan);
}

function isBappRowForRole(row, role = "PML") {
  const jabatan = upperText(row?.jabatan || row?.jabatan_raw || "");
  if (role === "PML") return /PML|PENGAWAS/.test(jabatan);
  if (role === "PPL") return /PPL|PENCACAH/.test(jabatan);
  return true;
}

function extractNomorPrefix(value) {
  const text = String(value ?? "").trim();
  if (!text) return "";
  const match = text.match(/(\d+)/);
  return match ? match[1] : "";
}

function parseDateInput(dateStr) {
  if (!dateStr) return null;
  const parts = String(dateStr).split("-");
  if (parts.length === 3) {
    const [year, month, day] = parts.map(Number);
    if ([year, month, day].every((n) => Number.isFinite(n))) {
      return new Date(year, month - 1, day);
    }
  }
  const parsed = new Date(dateStr);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function numberToIndonesianWords(value) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < 1 || number > 31) return "";
  const words = [
    "", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan",
    "sepuluh", "sebelas", "dua belas", "tiga belas", "empat belas", "lima belas", "enam belas",
    "tujuh belas", "delapan belas", "sembilan belas", "dua puluh", "dua puluh satu",
    "dua puluh dua", "dua puluh tiga", "dua puluh empat", "dua puluh lima", "dua puluh enam",
    "dua puluh tujuh", "dua puluh delapan", "dua puluh sembilan", "tiga puluh", "tiga puluh satu"
  ];
  return words[number];
}

function getBappDateParts(dateStr) {
  const parsed = parseDateInput(dateStr);
  if (!parsed) return { hari_terbilang: "", tanggal_terbilang: "", tanggal: "", bulan: "", bulan_terbilang: "" };
  const monthNumber = String(parsed.getMonth() + 1);
  const monthNames = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
  ];
  return {
    hari_terbilang: parsed.toLocaleDateString("id-ID", { weekday: "long" }),
    tanggal_terbilang: numberToIndonesianWords(parsed.getDate()),
    tanggal: String(parsed.getDate()),
    bulan: monthNumber,
    bulan_terbilang: monthNames[parsed.getMonth()] || "",
  };
}

function getBappIdentityKey(row, role = "PML") {
  const email = cleanText(row?.email || "");
  if (email) return `EMAIL::${upperText(email)}`;
  const name = cleanText(row?.nama || "");
  const jabatan = cleanText(row?.jabatan_raw || row?.jabatan || role || "");
  return `NAME::${upperText(name)}::${upperText(jabatan)}`;
}

function dedupeBappRows(rows = []) {
  const seen = new Set();
  return (rows || []).filter((row) => {
    const key = getBappIdentityKey(row);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function normalizeLampiranRow(row) {
  const normalized = {};

  Object.entries(row).forEach(([k, v]) => {
    const key = String(k ?? "")
      .trim()
      .toLowerCase()
      .replace(/\s+/g, " ");

    normalized[key] = String(v ?? "").trim();
  });

  const get = (...keys) => {
    for (const key of keys) {
      const value = normalized[key];
      if (value != null && String(value).trim() !== "") return String(value).trim();
    }
    return "";
  };

  // 🔥 FIX: jangan andalkan exact-match nama header untuk nomor kontrak — variasi
  // penulisan header di Google Sheet (titik, spasi, urutan kata, dll) gampang
  // membuat exact-match get() gagal dan hasilnya jadi kosong/undefined.
  // Cari secara fuzzy: kolom apa pun yang namanya mengandung "kontrak" DAN
  // mengandung "pml" (atau "ppl") akan dipakai, berapa pun variasi penulisannya.
  const findKontrakFuzzy = (token) => {
    for (const [key, value] of Object.entries(normalized)) {
      if (key.includes("kontrak") && key.includes(token) && String(value ?? "").trim() !== "") {
        return String(value).trim();
      }
    }
    return "";
  };

  return {
    no: get("no"),

    // Struktur sheet Lampiran dari Google Sheet:
    // PENGAWAS = PML, PENCACAH = PPL.
    nama_pml: get("pengawas", "nama pml", "nama_pml", "pml"),
    nama_ppl: get("pencacah", "nama ppl", "nama_ppl", "ppl", "nama petugas lapangan sensus"),

    email_pengawas: get("email pengawas", "mail pengawas", "email pml"),
    email_pencacah: get("email pencacah", "mail pencacah", "email ppl"),

    kdprov: get("kdprov"),
    kdkab: get("kdkab"),
    kdkec: get("kdkec"),
    kddesa: get("kddesa"),
    kdsls: get("kdsls"),
    kdsubsls: get("kdsubsls"),
    kdsubslspanjang: get("kdsubsls_25_2", "kdsubsls_25", "kdsubsls panjang"),

    nmprov: get("nmprov"),
    nmkab: get("nmkab"),
    kecamatan: get("nmkec", "kecamatan", "kecamatan/distrik", "kecamatan / distrik").toUpperCase(),
    kelurahan: get("nmdesa", "kelurahan", "desa/kampung/nagari", "desa / kampung / nagari").toUpperCase(),
    sls: get("nmsls", "sls"),
    subsls: get("nmsubsls", "sub-sls", "sub sls"),

    // Kolom ini opsional. Kalau tidak ada, jumlah dihitung dari banyaknya baris.
    jumlah: get("jumlah", "jumlah sls/sub-sls", "jumlah sls/sub sls", "jumlah sls / sub-sls", "jumlah sls"),

    jabatan: get("jabatan").toUpperCase(),
    kelas: get("kelas"),
    gelombang: get("gelombang"),
    hotel: get("tc", "hotel", "tempat").toUpperCase(),
    tc: get("tc", "hotel", "tempat").toUpperCase(),

    // Nomor kontrak per jenis petugas. Dipakai untuk variabel {nomor_kontrak} di
    // template, sesuai jenis dokumen yang sedang digenerate (PML atau PPL).
    // Coba exact-match alias dulu, kalau tidak ketemu baru fallback ke fuzzy search.
    nomor_kontrak_pml:
      get("no kontrak pml", "nomor kontrak pml", "no_kontrak_pml", "kontrak pml", "no. kontrak pml") ||
      findKontrakFuzzy("pml"),
    nomor_kontrak_ppl:
      get("no kontrak ppl", "nomor kontrak ppl", "no_kontrak_ppl", "kontrak ppl", "no. kontrak ppl") ||
      findKontrakFuzzy("ppl"),
  };
}

// 🔥 FIX: Google Sheet sumber Lampiran biasanya pakai "merged cell" secara visual —
// nama Pengawas/Pencacah, kecamatan, kelurahan, dst hanya diisi SEKALI di baris pertama
// tiap blok, lalu baris-baris SLS berikutnya di bawahnya dikosongkan.
// XLSX.utils.sheet_to_json TIDAK menurunkan nilai merged cell, jadi baris-baris itu
// terbaca kosong dan akhirnya DIBUANG oleh generateLampiran() (karena identity-nya kosong).
// Ini sebabnya 1 PML yang sebenarnya membawahi 7-8 SLS/PPL hanya muncul 2-3 baris saja.
//
// Solusinya: forward-fill — isi sel kosong dengan nilai terakhir yang valid di kolom
// yang sama, KHUSUS untuk kolom yang memang lazim merged (nama petugas, lokasi, dll).
// Kolom kode SLS/Sub-SLS sengaja TIDAK di-forward-fill karena itu harus unik per baris.
const LAMPIRAN_FORWARD_FILL_KEYS = [
  "nama_pml", "nama_ppl",
  "email_pengawas", "email_pencacah",
  "kdprov", "kdkab", "kdkec", "kddesa",
  "nmprov", "nmkab", "kecamatan", "kelurahan",
  "jabatan", "kelas", "gelombang", "hotel", "tc",
  "nomor_kontrak_pml", "nomor_kontrak_ppl",
];

function forwardFillLampiranRows(rows) {
  const lastValue = {};
  return rows.map((row) => {
    const filled = { ...row };
    for (const key of LAMPIRAN_FORWARD_FILL_KEYS) {
      const value = cleanText(filled[key]);
      if (value) {
        lastValue[key] = value;
      } else if (lastValue[key]) {
        // Sel kosong karena merged cell -> turunkan nilai dari baris di atasnya
        filled[key] = lastValue[key];
      }
    }
    return filled;
  });
}

function parseLampiranXlsxData(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });

  // Cari tab bernama Lampiran secara toleran:
  // - tidak sensitif kapital
  // - mengabaikan spasi di awal/akhir
  // Ini mencegah kasus tab "Lampiran " terbaca di console tetapi dianggap kosong di frontend.
  const sheetName = workbook.SheetNames.find(
    (name) => String(name ?? "").trim().toLowerCase() === "lampiran"
  );

  console.log("Daftar sheet terbaca:", workbook.SheetNames);

  if (!sheetName) {
    console.warn("Sheet bernama 'Lampiran' tidak ditemukan. Sheet tersedia:", workbook.SheetNames);
    return [];
  }

  const sheet = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });

  console.log("Raw Lampiran:", raw);

  const normalizedRows = raw.map(normalizeLampiranRow);

  // 🔥 Diagnostik: tampilkan contoh hasil pembacaan No Kontrak PML/PPL dari 3 baris
  // pertama, supaya kalau masih kosong/undefined, gampang dicek header apa saja yang
  // terbaca dari sheet vs nilai yang berhasil diambil.
  if (normalizedRows.length > 0) {
    console.log(
      "Lampiran: header mentah baris pertama ->",
      raw.length ? Object.keys(raw[0]) : []
    );
    console.log(
      "Lampiran: contoh hasil No Kontrak PML/PPL (3 baris pertama) ->",
      normalizedRows.slice(0, 3).map((r) => ({
        nama_pml: r.nama_pml,
        nama_ppl: r.nama_ppl,
        nomor_kontrak_pml: r.nomor_kontrak_pml,
        nomor_kontrak_ppl: r.nomor_kontrak_ppl,
      }))
    );
  }

  // 🔥 FIX: Google Sheet biasanya punya banyak baris kosong tambahan di ekor sheet
  // (range default jauh lebih panjang dari data aslinya). Kalau forward-fill langsung
  // dijalankan ke SEMUA baris (termasuk baris kosong di ekor), baris-baris kosong itu
  // akan "ketarik" nilai PENCACAH/kecamatan/kelurahan terakhir yang valid, sehingga
  // jumlah baris yang diproses bisa membengkak jadi ribuan baris palsu milik PPL
  // terakhir. Ini yang menyebabkan error "Array buffer allocation failed" saat
  // generate Lampiran PPL.
  //
  // Solusinya: saring dulu baris yang BENAR-BENAR baris data, berdasarkan kolom yang
  // TIDAK PERNAH di-forward-fill (kdsls/kdkec/kddesa/sls/subsls) — kolom ini aman
  // dipakai sebagai penanda baris asli, karena nilainya selalu apa adanya dari sheet,
  // bukan hasil "tebakan" forward-fill.
  const candidateRows = normalizedRows.filter((r) =>
    cleanText(r.kdsls) || cleanText(r.kdkec) || cleanText(r.kddesa) || cleanText(r.sls) || cleanText(r.subsls)
  );

  if (normalizedRows.length !== candidateRows.length) {
    console.log(
      `Lampiran: membuang ${normalizedRows.length - candidateRows.length} baris kosong/bukan-data sebelum forward-fill (dari ${normalizedRows.length} baris mentah).`
    );
  }

  // 🔥 FIX: turunkan nilai dari merged cell SEBELUM difilter,
  // supaya baris dengan nama_pml/nama_ppl kosong tidak ikut terbuang.
  const filledRows = forwardFillLampiranRows(candidateRows);

  const parsed = filledRows
    .filter((r) => r.nama_pml || r.nama_ppl || r.kecamatan || r.kelurahan || r.sls || r.subsls);

  console.log("Lampiran parsed rows (setelah forward-fill):", parsed);

  return parsed;
}

function normalizeGoogleSheetUrl(url) {
  if (!url) return null;
  try {
    const parsed = new URL(url.trim());
    const sheetIdMatch = parsed.pathname.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!sheetIdMatch) return null;
    return {
      spreadsheetId: sheetIdMatch[1],
      exportUrl: `https://docs.google.com/spreadsheets/d/${sheetIdMatch[1]}/export?format=xlsx`,
    };
  } catch { return null; }
}

function buildGoogleSheetsApiMetadataUrl(spreadsheetId, apiKey) {
  return `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}?key=${encodeURIComponent(apiKey)}`;
}

function buildGoogleSheetsApiValuesUrl(spreadsheetId, apiKey, sheetName) {
  return `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}/values/${encodeURIComponent(sheetName)}?key=${encodeURIComponent(apiKey)}`;
}

function parseGoogleSheetApiRows(values = []) {
  const headers = (values[0] || []).map((value) => String(value ?? "").trim());
  return values.slice(1).map((row) => {
    const record = {};
    headers.forEach((header, index) => {
      record[header] = String(row?.[index] ?? "");
    });
    return record;
  });
}

function findAdministrasiSheetFromRows(sheetRows = []) {
  for (const sheet of sheetRows) {
    const rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    if (rows.length === 0) continue;
    const headers = Object.keys(rows[0]).map((h) => String(h ?? "").trim());
    if (isAdministrasiSheet(headers)) {
      return { sheet, data: rows, headers, sheetName: sheet.sheetName };
    }
  }
  return null;
}

function findLampiranSheetFromRows(sheetRows = []) {
  for (const sheet of sheetRows) {
    const rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    if (rows.length === 0) continue;
    const headers = Object.keys(rows[0]).map((h) => String(h ?? "").trim());
    const normalizedHeaders = headers.map((h) => String(h ?? "").trim().toLowerCase());
    const hasPml = normalizedHeaders.some((h) => ["nama pml", "nama_pml", "pengawas", "pml"].includes(h));
    const hasPpl = normalizedHeaders.some((h) => ["nama ppl", "nama_ppl", "pencacah", "ppl"].includes(h));
    const hasKecamatan = normalizedHeaders.some((h) => ["kecamatan", "kecamatan/distrik", "kecamatan / distrik", "nmkec"].includes(h));
    const hasSls = normalizedHeaders.some((h) => ["sls", "nmsls"].includes(h));
    if (hasPml && hasPpl && hasKecamatan && hasSls) {
      return { sheet, data: rows, headers, sheetName: sheet.sheetName };
    }
  }
  return null;
}

function findBappSheetFromRows(sheetRows = []) {
  for (const sheet of sheetRows) {
    const rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    if (rows.length === 0) continue;
    const headers = Object.keys(rows[0]).map((h) => String(h ?? "").trim());
    const normalizedHeaders = headers.map((h) => String(h ?? "").trim().toLowerCase());
    const hasName = normalizedHeaders.some((h) => ["nama", "nama lengkap", "nama_lengkap", "nama petugas", "nama_petugas"].includes(h));
    const hasEmail = normalizedHeaders.some((h) => ["email", "email petugas", "email peserta", "mail"].includes(h));
    const hasJabatan = normalizedHeaders.some((h) => ["jabatan", "posisi", "jenis petugas", "role", "kategori"].includes(h));
    const hasNomorSpk = normalizedHeaders.some((h) => ["nomor spk", "nomor_spk", "nomor kontrak", "nomor_kontrak", "spk"].includes(h));
    if (hasName && hasEmail && hasJabatan && hasNomorSpk) {
      return { sheet, data: rows, headers, sheetName: sheet.sheetName };
    }
  }
  return null;
}

function normalizeHeaderKeys(headers) {
  return headers.map((h) => String(h ?? "").trim().toLowerCase());
}

function isAdministrasiSheet(headers) {
  const keys = normalizeHeaderKeys(headers);
  const hasName = keys.some((h) => ["nama", "nama lengkap", "nama_lengkap", "nama-lengkap"].includes(h));
  const hasJabatan = keys.some((h) => ["jabatan", "posisi"].includes(h));
  const hasKelas = keys.some((h) => ["kelas"].includes(h));
  const hasGelombang = keys.some((h) => ["gelombang"].includes(h));
  // "tc" dapat berupa kolom terpisah atau bagian dari header lain (seperti "Hotel")
  const hasTc = keys.some((h) => ["tc", "hotel", "tempat"].includes(h));
  return hasName && hasJabatan && hasKelas && hasGelombang && hasTc;
}

// Fungsi untuk menemukan sheet yang berisi data administrasi
function findAdministrasiSheet(workbook) {
  for (const sheetName of workbook.SheetNames) {
    const sheet = workbook.Sheets[sheetName];
    const data = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
    if (data.length > 0) {
      const headers = Object.keys(data[0]).map((h) => String(h ?? "").trim());
      if (isAdministrasiSheet(headers)) {
        return { sheet, data, headers, sheetName };
      }
    }
  }
  return null;
}

async function loadGoogleSheet(source, apiKey = "") {
  if (source?.spreadsheetId && apiKey) {
    const metadataResponse = await fetch(buildGoogleSheetsApiMetadataUrl(source.spreadsheetId, apiKey));
    if (!metadataResponse.ok) throw new Error(`Google Sheets API error: ${metadataResponse.status}`);

    const metadata = await metadataResponse.json();
    const sheets = Array.isArray(metadata.sheets) ? metadata.sheets : [];
    const sheetRows = [];

    for (const sheet of sheets) {
      const sheetName = sheet?.properties?.title;
      if (!sheetName) continue;
      const valuesResponse = await fetch(buildGoogleSheetsApiValuesUrl(source.spreadsheetId, apiKey, sheetName));
      if (!valuesResponse.ok) continue;
      const valuesData = await valuesResponse.json();
      const rows = parseGoogleSheetApiRows(valuesData.values || []);
      sheetRows.push({ sheetName, rows });
    }

    const adminSheetInfo = findAdministrasiSheetFromRows(sheetRows);
    let data = [];
    let rawHeaders = [];

    if (adminSheetInfo) {
      const { data: raw, headers: foundHeaders, sheetName: foundSheetName } = adminSheetInfo;
      console.log(`✓ Sheet data administrasi ditemukan via API: "${foundSheetName}"`);
      data = raw.map(normalizeRowHeaders).filter((r) => r.nama !== "" || r.nik !== "" || r.sobatId !== "");
      rawHeaders = foundHeaders;
    } else {
      const availableSheets = sheetRows.map((sheet) => sheet.sheetName).join(", ");
      console.warn(`Sheet data administrasi tidak ditemukan via API. Sheet tersedia: ${availableSheets}`);
    }

    const lampiranSheetInfo = findLampiranSheetFromRows(sheetRows);
    const lampiran = lampiranSheetInfo ? lampiranSheetInfo.data.map(normalizeLampiranRow) : [];

    const bappSheetInfo = findBappSheetFromRows(sheetRows);
    const bappData = bappSheetInfo ? bappSheetInfo.data.map(normalizeBappRow) : [];

    console.log("Data petugas parsed rows (API):", data);
    console.log("Lampiran state candidate rows (API):", lampiran);
    console.log("BAPP parsed rows (API):", bappData);

    return { data, lampiran, bappData, rawHeaders };
  }

  const response = await fetch(source?.exportUrl || source);
  if (!response.ok) throw new Error(`HTTP ${response.status}`);

  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });

  const adminSheetInfo = findAdministrasiSheet(workbook);
  let data = [];
  let rawHeaders = [];

  if (adminSheetInfo) {
    const { data: raw, headers: foundHeaders, sheetName: foundSheetName } = adminSheetInfo;
    console.log(`✓ Sheet data administrasi ditemukan: "${foundSheetName}"`);
    data = raw.map(normalizeRowHeaders).filter((r) => r.nama !== "" || r.nik !== "" || r.sobatId !== "");
    rawHeaders = foundHeaders;
  } else {
    const availableSheets = workbook.SheetNames.map((name, idx) => `[${idx}] ${name}`).join(", ");
    console.warn(`Sheet data administrasi tidak ditemukan. Lanjutkan hanya dengan data lampiran. Sheet tersedia: ${availableSheets}`);
  }

  const lampiran = parseLampiranXlsxData(arrayBuffer);
  const bappData = parseBappData(arrayBuffer);

  console.log("Data petugas parsed rows:", data);
  console.log("Lampiran state candidate rows:", lampiran);
  console.log("BAPP parsed rows:", bappData);

  return { data, lampiran, bappData, rawHeaders };
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────

function cleanText(value) { return String(value ?? "").trim(); }
function upperText(value) { return cleanText(value).toUpperCase(); }

function uniqueSorted(values) {
  return [...new Set(values.map(cleanText).filter(Boolean))].sort((a, b) =>
    a.localeCompare(b, "id-ID", { numeric: true, sensitivity: "base" })
  );
}

function formatTanggalIndonesia(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (Number.isNaN(d.getTime())) return dateStr;
  return d.toLocaleDateString("id-ID", { day: "numeric", month: "long", year: "numeric" });
}

function formatTanggalLengkapIndonesia(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (Number.isNaN(d.getTime())) return dateStr;
  return d.toLocaleDateString("id-ID", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
}

function formatTanggalBulanIndonesia(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (Number.isNaN(d.getTime())) return dateStr;
  return d.toLocaleDateString("id-ID", { day: "numeric", month: "long" });
}

function formatHariIndonesia(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr);
  if (Number.isNaN(d.getTime())) return dateStr;
  return d.toLocaleDateString("id-ID", { weekday: "long" });
}

function calcDurationDays(startDate, endDate) {
  if (!startDate || !endDate) return "";
  const start = new Date(startDate);
  const end = new Date(endDate);
  if (Number.isNaN(start.getTime()) || Number.isNaN(end.getTime())) return "";
  const diff = Math.floor((end.getTime() - start.getTime()) / (1000 * 60 * 60 * 24)) + 1;
  return diff > 0 ? diff : 0;
}

function normalizeJamIndonesia(value, fallback = "") {
  if (!value) return fallback;
  return String(value).trim().replace(":", ".");
}

function formatRupiah(value) {
  if (value == null || value === "") return "";
  const number = Number(String(value).replace(/[^0-9-]/g, ""));
  if (Number.isNaN(number)) return String(value);
  return number.toLocaleString("id-ID");
}

function capitalizeWords(text) {
  return String(text || "").split(/\s+/).filter(Boolean)
    .map((w) => w[0].toUpperCase() + w.slice(1)).join(" ");
}

function spellTerbilang(value) {
  const units = ["", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan"];
  const teens = ["sepuluh", "sebelas", "dua belas", "tiga belas", "empat belas", "lima belas",
    "enam belas", "tujuh belas", "delapan belas", "sembilan belas"];
  const toWords = (n) => {
    if (n < 10) return units[n];
    if (n < 20) return teens[n - 10];
    if (n < 100) return `${toWords(Math.floor(n / 10))} puluh${n % 10 ? ` ${toWords(n % 10)}` : ""}`.trim();
    if (n < 200) return `seratus${n % 100 ? ` ${toWords(n % 100)}` : ""}`.trim();
    if (n < 1000) return `${toWords(Math.floor(n / 100))} ratus${n % 100 ? ` ${toWords(n % 100)}` : ""}`.trim();
    if (n < 2000) return `seribu${n % 1000 ? ` ${toWords(n % 1000)}` : ""}`.trim();
    if (n < 1000000) return `${toWords(Math.floor(n / 1000))} ribu${n % 1000 ? ` ${toWords(n % 1000)}` : ""}`.trim();
    if (n < 1000000000) return `${toWords(Math.floor(n / 1000000))} juta${n % 1000000 ? ` ${toWords(n % 1000000)}` : ""}`.trim();
    if (n < 1000000000000) return `${toWords(Math.floor(n / 1000000000))} miliar${n % 1000000000 ? ` ${toWords(n % 1000000000)}` : ""}`.trim();
    return String(n);
  };
  const number = Number(String(value).replace(/[^0-9]/g, ""));
  if (Number.isNaN(number) || number === 0) return capitalizeWords("nol");
  return capitalizeWords(toWords(number));
}

const PARTICIPANT_ROLE_ORDER = { INDA: 0, PANITIA: 1, PML: 2, PPL: 3 };
function pesertaRoleOrder(jabatan) { return PARTICIPANT_ROLE_ORDER[upperText(jabatan)] ?? 99; }
function sortPesertaByJabatanOrder(peserta = []) {
  return [...peserta].sort((a, b) => {
    const diff = pesertaRoleOrder(a.jabatan) - pesertaRoleOrder(b.jabatan);
    if (diff !== 0) return diff;
    return cleanText(a.nama).localeCompare(cleanText(b.nama), "id-ID", { sensitivity: "base" });
  });
}

const DAFTAR_HADIR_PESERTA_ROLE_ORDER = { "KEPALA BPS JAKARTA TIMUR": -1, INDA: 0, PANITIA: 1, PML: 2, PPL: 3 };
function pesertaRoleOrderDaftarHadir(jabatan) { return DAFTAR_HADIR_PESERTA_ROLE_ORDER[upperText(jabatan)] ?? 99; }
function sortDaftarHadirPeserta(peserta = []) {
  return [...peserta].sort((a, b) => {
    const diff = pesertaRoleOrderDaftarHadir(a.jabatan) - pesertaRoleOrderDaftarHadir(b.jabatan);
    if (diff !== 0) return diff;
    return cleanText(a.nama).localeCompare(cleanText(b.nama), "id-ID", { sensitivity: "base" });
  });
}

// ─── FILTER GROUPS ───────────────────────────────────────────────────────────

const DAFTAR_HADIR_GROUPS = {
  "pml-ppl":      { label: "PML & PPL",      roles: ["PML", "PPL"] },
  "panitia-inda": { label: "Panitia & Inda", roles: ["PANITIA", "INDA", "KEPALA BPS JAKARTA TIMUR"] },
};

function jabatanMasukGroup(jabatan, groupKey) {
  const group = DAFTAR_HADIR_GROUPS[groupKey];
  return group ? group.roles.includes(upperText(jabatan)) : false;
}

// ─── TEMPLATE URLS ───────────────────────────────────────────────────────────

const DAFTAR_HADIR_TEMPLATE_URL               = "/templates/1. Daftar Hadir Pelatihan SE2026.docx";
const TANDA_TERIMA_TEMPLATE_URL               = "/templates/2. Tanda Terima Perlengkapan SE2026.docx";
const TANDA_TERIMA_LAPANGAN_TEMPLATE_URL      = "/templates/2. Tanda Terima Perlengkapan SE2026 - Copy.docx";
const SURAT_PERNYATAAN_KENDARAAN_TEMPLATE_URL = "/templates/3. Super Kendis Pelatihan SE2026.docx";
const PENGELUARAN_RIIL_TEMPLATE_URL           = "/templates/4. DPR_Pelatihan SE 2026.docx";
const SPJ_TEMPLATE_URL                        = "/templates/5. SPJ Pelatihan_SE26.docx";
const SPD_TEMPLATE_URL                        = "/templates/6. SPD.docx";
const SPD_LAMPIRAN_TEMPLATE_URL               = "/templates/6. Lampiran SPD.docx";
const SURAT_TUGAS_TEMPLATE_URL                = "/templates/6. Surat Tugas.docx";
const BAPP_PML_TEMPLATE_URL                   = "/templates/BAPP PML.docx";
const BAPP_PPL_TEMPLATE_URL                   = "/templates/BAPP PPL.docx";
const SURAT_PERNYATAAN_PENYELESAIAN_LAPANGAN_TEMPLATE_URL = "/templates/Dasar Pembayaran.docx";
const LAMPIRAN_PML_TEMPLATE_URL              = "/templates/LAMPIRAN PML.docx";
const LAMPIRAN_PPL_TEMPLATE_URL              = "/templates/LAMPIRAN PPL.docx";
const BAST_PML_TEMPLATE_URL = "/templates/BAST PML.docx";
const BAST_PPL_TEMPLATE_URL = "/templates/BAST PPL.docx";

// ─── TEMPLATE DATA BUILDERS ───────────────────────────────────────────────────

function buildBappTemplateData(formValues, row = {}, role = "PML") {
  const tanggalSurat = cleanText(formValues?.tanggal_surat || "");
  const nama = cleanText(row?.nama || row?.nama_pml || row?.nama_ppl || "");
  const jabatan = cleanText(row?.jabatan_raw || row?.jabatan || "");
  const wilayah = cleanText(row?.wilayah || row?.tempat || row?.asal || "");
  const dateParts = getBappDateParts(tanggalSurat);
  const nomorKontrak = cleanText(row?.nomor_spk || row?.nomor_kontrak || formValues?.nomor_kontrak || "");
  return {
    tanggal_surat: tanggalSurat,
    tanggal_surat_fmt: formatTanggalIndonesia(tanggalSurat),
    tanggal_surat_lengkap: formatTanggalLengkapIndonesia(tanggalSurat),
    hari_terbilang: dateParts.hari_terbilang,
    hari: dateParts.hari_terbilang,
    tanggal_terbilang: dateParts.tanggal_terbilang,
    tanggal: dateParts.tanggal,
    bulan: dateParts.bulan,
    bulan_terbilang: dateParts.bulan_terbilang,
    nama,
    nama_peserta: nama,
    nama_petugas: nama,
    jabatan,
    jabatan_peserta: jabatan,
    jabatan_petugas: jabatan,
    role,
    jenis: role,
    jenis_dokumen: "BAPP",
    wilayah,
    wilayah_tugas: wilayah,
    tempat: cleanText(formValues?.tempat || ""),
    email: cleanText(row?.email || ""),
    kelas: cleanText(row?.kelas || ""),
    gelombang: cleanText(row?.gelombang || ""),
    nomor_surat: cleanText(formValues?.nomor_surat || ""),
    nomor_dokumen: cleanText(formValues?.nomor_surat || ""),
    nomor_prefix: extractNomorPrefix(nomorKontrak),
    nomor_kontrak: nomorKontrak,
    nik: cleanText(row?.nik || ""),
    sls_40: cleanText(row?.sls_40 || ""),
  };
}

async function createBappBlob(templateUrl, formValues, row, role) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template BAPP: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildBappTemplateData(formValues || {}, row || {}, role));
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

async function generateSingleBapp(templateUrl, formValues, row, role) {
  const blob = await createBappBlob(templateUrl, formValues || {}, row || {}, role);
  const safeName = sanitizeFileName(cleanText(row?.nama || `${role}-bapp`));
  saveAs(blob, `BAPP ${role} - ${safeName}.docx`);
}

const BAPP_ZIP_BATCH_SIZE = 150;

async function generateBapp(templateUrl, formValues, rows, role, onProgress) {
  if (!rows || rows.length === 0) throw new Error("Tidak ada data BAPP untuk role yang dipilih.");
  const uniqueRows = dedupeBappRows(rows);
  const totalBatches = Math.ceil(uniqueRows.length / BAPP_ZIP_BATCH_SIZE);

  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
    const batchRows = uniqueRows.slice(
      batchIndex * BAPP_ZIP_BATCH_SIZE,
      (batchIndex + 1) * BAPP_ZIP_BATCH_SIZE
    );

    const files = [];
    for (const row of batchRows) {
      const blob = await createBappBlob(templateUrl, formValues || {}, row || {}, role);
      files.push({
        name: `BAPP ${role} - ${sanitizeFileName(cleanText(row?.nama || "Tanpa Nama"))}.docx`,
        blob,
      });
    }

    if (typeof onProgress === "function") {
      onProgress({
        batchIndex: batchIndex + 1,
        totalBatches,
        totalRows: uniqueRows.length,
      });
    }

    if (files.length === 1 && totalBatches === 1) {
      saveAs(files[0].blob, files[0].name);
      continue;
    }

    const batchSuffix = totalBatches > 1 ? ` - Bagian ${batchIndex + 1} dari ${totalBatches}` : "";
    await downloadMultipleAsZip(files, `BAPP ${role} ${cleanText(formValues?.tanggal_surat || "SE2026")}${batchSuffix}.zip`);
  }
}

async function createSuratPernyataanPenyelesaianLapanganBlob(templateUrl, formValues, row) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template surat: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render({
    nomor_prefix: cleanText(row?.nomor_prefix || extractNomorPrefix(row?.nomor_spk || row?.nomor_kontrak || "")),
    nama_petugas: cleanText(row?.nama || row?.nama_pml || row?.nama_ppl || ""),
    nik: cleanText(row?.nik || ""),
    nomor_kontrak: cleanText(row?.nomor_spk || row?.nomor_kontrak || ""),
  });
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

async function generateSingleSuratPernyataanPenyelesaianLapangan(templateUrl, row, onProgress) {
  const blob = await createSuratPernyataanPenyelesaianLapanganBlob(templateUrl, {}, row);
  const safeName = sanitizeFileName(cleanText(row?.nama || "Tanpa Nama"));
  if (typeof onProgress === "function") {
    onProgress({ batchIndex: 1, totalBatches: 1, totalRows: 1 });
  }
  saveAs(blob, `Surat Pernyataan Penyelesaian Lapangan - ${safeName}.docx`);
}

async function generateSuratPernyataanPenyelesaianLapangan(templateUrl, rows, onProgress) {
  if (!rows || rows.length === 0) throw new Error("Tidak ada data untuk Surat Pernyataan Penyelesaian Lapangan.");
  const uniqueRows = dedupeBappRows(rows);
  const totalBatches = Math.ceil(uniqueRows.length / BAPP_ZIP_BATCH_SIZE);

  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
    const batchRows = uniqueRows.slice(batchIndex * BAPP_ZIP_BATCH_SIZE, (batchIndex + 1) * BAPP_ZIP_BATCH_SIZE);
    const files = [];
    for (const row of batchRows) {
      const blob = await createSuratPernyataanPenyelesaianLapanganBlob(templateUrl, {}, row);
      files.push({
        name: `Surat Pernyataan Penyelesaian Lapangan - ${sanitizeFileName(cleanText(row?.nama || "Tanpa Nama"))}.docx`,
        blob,
      });
    }

    if (typeof onProgress === "function") {
      onProgress({ batchIndex: batchIndex + 1, totalBatches, totalRows: uniqueRows.length });
    }

    const batchSuffix = totalBatches > 1 ? ` - Bagian ${batchIndex + 1} dari ${totalBatches}` : "";
    await downloadMultipleAsZip(files, `Surat Pernyataan Penyelesaian Lapangan${batchSuffix}.zip`);
  }
}

// DAFTAR HADIR
const DAFTAR_HADIR_KEPALA_BPS = {
  nama:      "Widiastuti",
  jabatan:   "Penanggung Jawab",
  wilTugas:  "BPS Kota Jakarta Timur",
};

function buildDaftarHadirTemplateData(formValues, peserta, namaInda, selectedFilterGroup = "") {
  const jamMulai        = normalizeJamIndonesia(formValues.jamMulai,   "07.30");
  const jamSelesai      = normalizeJamIndonesia(formValues.jamSelesai, "18.00");
  const tanggalFmt      = formatTanggalIndonesia(formValues.tanggal);
  const tanggalKegiatan = formatTanggalLengkapIndonesia(formValues.tanggal);
  const jamFmt          = `${jamMulai} - ${jamSelesai}`;
  const isPanitiaInda   = selectedFilterGroup === "panitia-inda";
  const isPmlPpl        = selectedFilterGroup === "pml-ppl";

  const hotelValue = cleanText(formValues.hotel || formValues.tempat || "").toLowerCase();
  const isHotelBwp = hotelValue.includes("bwp");
  const gelombangValue = cleanText(formValues.gelombang || "");

  let kepalaBpsEntry = {
    no:        1,
    nama:      "Widiastuti",
    jabatan:   "Penanggung Jawab",
    kecamatan: "BPS Kota Jakarta Timur",
    wil_tugas: "BPS Kota Jakarta Timur",
    wilTugas:  "BPS Kota Jakarta Timur",
  };

  if (gelombangValue === "4") {
    if (hotelValue.includes("bwp")) {
      kepalaBpsEntry = {
        no:        1,
        nama:      "Budi Utami",
        jabatan:   "Penanggung Jawab",
        kecamatan: "BPS Kota Jakarta Timur",
        wil_tugas: "BPS Kota Jakarta Timur",
        wilTugas:  "BPS Kota Jakarta Timur",
      };
    } else if (hotelValue.includes("stis")) {
      kepalaBpsEntry = {
        no:        1,
        nama:      "Widiastuti",
        jabatan:   "Penanggung Jawab",
        kecamatan: "BPS Kota Jakarta Timur",
        wil_tugas: "BPS Kota Jakarta Timur",
        wilTugas:  "BPS Kota Jakarta Timur",
      };
    }
  }

  const includeKepalaBps = (isHotelBwp || (hotelValue.includes("stis") && gelombangValue === "4")) && !isPmlPpl;

  return {
    tanggal_kegiatan: tanggalKegiatan,
    tanggal_aja:      tanggalFmt,
    hari_tanggal:     tanggalFmt,
    tanggal:          tanggalFmt,
    jam_mulai:        jamMulai,
    jam_selesai:      jamSelesai,
    jam:              jamFmt,
    jam_kegiatan:     jamFmt,
    tempat:           formValues.tempat || formValues.hotel || "",
    tempat_kegiatan:  formValues.tempat || formValues.hotel || "",
    gelombang:        formValues.gelombang || "",
    kelas:            isPanitiaInda ? "-" : (formValues.kelas || ""),
    nama_inda:        isPanitiaInda ? "Ir. Tristiati, MA" : (namaInda || ""),
    keterangan_ttd:   isPanitiaInda ? "Kepala Sub Bagian Umum" : (isPmlPpl ? "Instruktur Daerah" : ""),
    peserta: includeKepalaBps ? [
      kepalaBpsEntry,
      ...sortDaftarHadirPeserta(peserta || []).map((p, idx) => ({
        no:        idx + 2,
        nama:      p.nama     || "",
        jabatan:   p.jabatan  || "",
        kecamatan: p.wilTugas || "",
        wil_tugas: p.wilTugas || "",
        wilTugas:  p.wilTugas || "",
      })),
    ] : sortDaftarHadirPeserta(peserta || []).map((p, idx) => ({
      no:        idx + 1,
      nama:      p.nama     || "",
      jabatan:   p.jabatan  || "",
      kecamatan: p.wilTugas || "",
      wil_tugas: p.wilTugas || "",
      wilTugas:  p.wilTugas || "",
    })),
  };
}

async function createDaftarHadirBlob(templateUrl, formValues, peserta, namaInda, selectedFilterGroup = "") {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildDaftarHadirTemplateData(formValues || {}, peserta || [], namaInda || "", selectedFilterGroup));
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

async function generateDaftarHadir(templateUrl, formValues, peserta, namaInda, selectedFilterGroup = "") {
  const blob = await createDaftarHadirBlob(templateUrl, formValues, peserta, namaInda, selectedFilterGroup);
  const safeGelombang = formValues?.gelombang || "X";
  const safeKelas     = selectedFilterGroup === "panitia-inda" ? "-" : (formValues?.kelas || "X");
  saveAs(blob, `Daftar Hadir Gelombang ${safeGelombang} Kelas ${safeKelas}.docx`);
}

// TANDA TERIMA
function buildTandaTerimaTemplateData(formValues, peserta) {
  const tanggalFmt = formatTanggalIndonesia(formValues.tanggal);
  const tanggalKegiatan = formatTanggalLengkapIndonesia(formValues.tanggal);
  const filtered   = sortPesertaByJabatanOrder((peserta || []).filter(p => ["PML", "PPL"].includes(upperText(p.jabatan))));
  return {
    tanggal_kegiatan: tanggalKegiatan,
    tanggal_aja:      tanggalFmt,
    tanggal:          tanggalFmt,
    tempat:           formValues.tempat || formValues.hotel || "",
    gelombang:        formValues.gelombang || "",
    kelas:            formValues.kelas || "",
    peserta: filtered.map((p, idx) => ({
      no: idx + 1, nama: p.nama || "", jabatan: p.jabatan || "",
      kecamatan: p.wilTugas || "", wil_tugas: p.wilTugas || "", wilTugas: p.wilTugas || "",
    })),
  };
}

async function createTandaTerimaBlob(templateUrl, formValues, peserta) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildTandaTerimaTemplateData(formValues || {}, peserta || []));
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

async function generateTandaTerimaSmart(formValues, peserta, tandaTerimaType) {
  const hotelValue = cleanText(formValues?.tempat || formValues?.hotel || "").toLowerCase();

  if (tandaTerimaType === "mitra-umum") {
    const wilTugas = peserta?.[0]?.wilTugas ? cleanText(peserta[0].wilTugas) : "MITRA UMUM";
    const blob = await createTandaTerimaBlob(TANDA_TERIMA_LAPANGAN_TEMPLATE_URL, formValues || {}, peserta || []);
    saveAs(blob, `Tanda Terima Perlengkapan Lapangan - ${wilTugas}.docx`);
    return;
  }

  if (tandaTerimaType === "lapangan") {
    if (hotelValue === "stis") {
      const blob = await createTandaTerimaBlob(TANDA_TERIMA_LAPANGAN_TEMPLATE_URL, formValues || {}, peserta || []);
      saveAs(blob, `Tanda Terima Perlengkapan STIS Gelombang ${formValues?.gelombang || "X"}.docx`);
      return;
    }

    const groups = (peserta || []).reduce((acc, p) => {
      const key = cleanText(p.wilTugas) || "LAINNYA";
      (acc[key] = acc[key] || []).push(p);
      return acc;
    }, {});

    for (const [kec, list] of Object.entries(groups)) {
      const blob = await createTandaTerimaBlob(TANDA_TERIMA_LAPANGAN_TEMPLATE_URL, formValues || {}, list);
      const safeKec = kec || "LAINNYA";
      saveAs(blob, `Tanda Terima Perlengkapan ${safeKec} Gelombang ${formValues?.gelombang || "X"} Kelas ${formValues?.kelas || "X"}.docx`);
    }
    return;
  }

  const blob = await createTandaTerimaBlob(TANDA_TERIMA_TEMPLATE_URL, formValues || {}, peserta || []);
  saveAs(blob, `Tanda Terima Perlengkapan Gelombang ${formValues?.gelombang || "X"} Kelas ${formValues?.kelas || "X"}.docx`);
}

// SURAT PERNYATAAN KENDARAAN / SUPER KENDIS
function buildSuratPernyataanKendaraanTemplateData(formValues, peserta) {
  const tanggalFmt = formatTanggalIndonesia(formValues.tanggal_surat || formValues.tanggal);
  const sorted     = sortPesertaByJabatanOrder(peserta || []);
  const hotelValue = cleanText(formValues.tempat || formValues.hotel || "").toLowerCase();

  let nomor_surtug_val = formValues.nomor_surat || formValues.nomor || "";
  let tanggal_kegiatan_val = formatTanggalLengkapIndonesia(formValues.tanggal) || "";
  let tanggal_surtug_val = tanggalFmt;

  if (hotelValue.includes("bwp")) {
    nomor_surtug_val = "B-999.1/3172/SS.220/2026";
    tanggal_kegiatan_val = "1 Juni - 3 Juni 2026";
    tanggal_surtug_val = "29 Mei 2026";
  } else if (hotelValue.includes("park")) {
    nomor_surtug_val = "B-999.3/3172/SS.220/2026";
    tanggal_kegiatan_val = "1 Juni - 3 Juni 2026";
    tanggal_surtug_val = "29 Mei 2026";
  } else if (hotelValue.includes("harper")) {
    nomor_surtug_val = "B-999.2/3172/SS.220/2026";
    tanggal_kegiatan_val = "1 Juni - 3 Juni 2026";
    tanggal_surtug_val = "29 Mei 2026";
  }

  return {
    tanggal_kegiatan: tanggal_kegiatan_val,
    nomor_surtug:     nomor_surtug_val,
    tanggal_surtug:   tanggal_surtug_val,
    tanggal_aja:      tanggalFmt,
    tanggal_surat:    tanggalFmt,
    tanggal:          tanggalFmt,
    tempat:           formValues.tempat || formValues.hotel || "",
    gelombang:        formValues.gelombang || "",
    kelas:            formValues.kelas || "-",
    peserta: sorted.map((p, idx) => {
      const pangkatGolRaw = cleanText(p.pangkatGol);
      const pangkatGolValue = (!pangkatGolRaw || upperText(pangkatGolRaw) === "#N/A") ? "-" : pangkatGolRaw;
      return {
        no:         idx + 1,
        nama:       p.nama || "",
        nik:        p.nik || "",
        jabatan:    p.jabatan || "",
        pangkat:    pangkatGolValue,
        pangkatGol: pangkatGolValue,
        pangkat_gol:pangkatGolValue,
        kecamatan:  p.wilTugas || "",
        wil_tugas:  p.wilTugas || "",
        wilTugas:   p.wilTugas || "",
      };
    }),
  };
}

async function createSuratPernyataanKendaraanBlob(templateUrl, formValues, peserta) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildSuratPernyataanKendaraanTemplateData(formValues || {}, peserta || []));
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

async function generateSuratPernyataanKendaraan(templateUrl, formValues, peserta) {
  const blob = await createSuratPernyataanKendaraanBlob(templateUrl, formValues, peserta);
  saveAs(blob, `Surat Pernyataan Kendaraan ${formValues?.tempat || "SE2026"} Gelombang ${formValues?.gelombang || "X"}.docx`);
}

// PENGELUARAN RIIL
function buildPengeluaranRiilTemplateData(formValues, peserta = []) {
  const sorted    = sortPesertaByJabatanOrder(peserta || []);
  const biayaTotal = 510000 * sorted.length;
  return {
    tanggal_aja:      formatTanggalIndonesia(formValues.tanggal_surat || ""),
    tanggal_surat:    formatTanggalIndonesia(formValues.tanggal_surat || ""),
    biaya_total:      formatRupiah(biayaTotal),
    biaya_terbilang:  `${spellTerbilang(biayaTotal)} rupiah`,
    peserta: sorted.map((p, idx) => {
      const isPanitiaInda = ["PANITIA", "INDA"].includes(upperText(p.jabatan));
      const pangkatGolRaw = cleanText(p.pangkatGol);
      const pangkatGolValue = (!pangkatGolRaw || upperText(pangkatGolRaw) === "#N/A") ? "-" : pangkatGolRaw;
      return { no: idx + 1, nama: cleanText(p.nama), nik: cleanText(p.nik), jabatan: cleanText(p.jabatan), pangkat: isPanitiaInda ? pangkatGolValue : pangkatGolRaw };
    }),
  };
}

async function createPengeluaranRiilBlob(templateUrl, formValues, peserta) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildPengeluaranRiilTemplateData(formValues || {}, peserta || []));
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

async function generatePengeluaranRiil(templateUrl, formValues, peserta) {
  const blob = await createPengeluaranRiilBlob(templateUrl, formValues, peserta);
  saveAs(blob, `Daftar Pengeluaran Riil ${formValues.no || "SE2026"}.docx`);
}

// SPJ
function buildSpjTemplateData(formValues, peserta = []) {
  const sorted = sortPesertaByJabatanOrder(peserta || []);
  const n = sorted.length;
  const total = 510000 * n;
  const tanggalAwal = formatTanggalBulanIndonesia(formValues.tanggal_awal_kegiatan || "");
  const tanggalAkhir = formatTanggalBulanIndonesia(formValues.tanggal_akhir_kegiatan || "");
  return {
    tanggal_aja:       formatTanggalIndonesia(formValues.tanggal_pelunasan || ""),
    tanggal_pelunasan: formatTanggalIndonesia(formValues.tanggal_pelunasan || ""),
    tanggal_awal:      tanggalAwal,
    tanggal_akhir:     tanggalAkhir,
    tanggal_awal_kegiatan: tanggalAwal,
    tanggal_akhir_kegiatan: tanggalAkhir,
    tempat: formValues.tempat || formValues.hotel || "",
    gelombang: formValues.gelombang || "",
    kelas: formValues.kelas || "-",
    kelompok_peserta: formValues.kelompokPeserta || "",
    kelompokPeserta: formValues.kelompokPeserta || "",
    jumlah_ok: n * 3,
    jumlah_uang: formatRupiah(170000 * n),
    total_jumlah_kotor: formatRupiah(total),
    total_jumlah_bersih: formatRupiah(total),
    total_terbilang: `${spellTerbilang(total)} rupiah`,
    ttd_kiri: formValues.ttd_kiri || "",
    ttd_kanan: formValues.ttd_kanan || "",
    peserta: sorted.map((p, idx) => ({
      no: idx + 1,
      nama: p.nama || "",
      nik: p.nik || "",
      jabatan: p.jabatan || "",
      kecamatan: p.wilTugas || "",
      wil_tugas: p.wilTugas || "",
      wilTugas: p.wilTugas || "",
    })),
  };
}

async function createSpjBlob(templateUrl, formValues, peserta) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildSpjTemplateData(formValues || {}, peserta || []));
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

async function generateSpj(templateUrl, formValues, peserta) {
  const blob = await createSpjBlob(templateUrl, formValues || {}, peserta || []);
  const safeKelompok = formValues?.kelompokPeserta || "SPJ";
  const safeTempat = formValues?.tempat || "SE2026";
  const safeGelombang = formValues?.gelombang || "X";
  const safeKelas = formValues?.kelas || "-";
  saveAs(blob, `SPJ ${safeKelompok} ${safeTempat} Gelombang ${safeGelombang} Kelas ${safeKelas}.docx`);
}

// SPD
function buildSpdTemplateData(formValues, peserta = []) {
  const sorted       = sortPesertaByJabatanOrder(peserta || []);
  const tanggalAwal  = formatTanggalIndonesia(formValues.tanggal_awal_kegiatan || "");
  const tanggalAkhir = formatTanggalIndonesia(formValues.tanggal_akhir_kegiatan || "");
  const lamaHari     = calcDurationDays(formValues.tanggal_awal_kegiatan, formValues.tanggal_akhir_kegiatan);
  const tanggalSurat = formatTanggalIndonesia(formValues.tanggal_surat || formValues.tanggal || "");

  return {
    nomor_dokumen: formValues.nomor_dokumen || formValues.nomor || "",
    nomor: formValues.nomor_dokumen || formValues.nomor || "",
    tanggal_aja: tanggalSurat,
    tanggal_surat: tanggalSurat,
    tanggal: tanggalSurat,
    tempat: formValues.tempat || formValues.hotel || "",
    lokasi: formValues.lokasi || formValues.tempat || formValues.hotel || "",
    gelombang: formValues.gelombang || "",
    kelas: formValues.kelas || "-",
    tanggal_awal_kegiatan: tanggalAwal,
    tanggal_akhir_kegiatan: tanggalAkhir,
    lama_hari: lamaHari,
    lama: lamaHari,
    jumlah_peserta: sorted.length || 0,
    ttd_nama: formValues.namaKabps || "",
    ttd_nip: formValues.nipKabps || "",
    peserta: sorted.map((p, idx) => {
      const jabatan = cleanText(p.jabatan);
      const pangkatGolRaw = cleanText(p.pangkatGol);
      const isPanitia = upperText(jabatan) === "PANITIA";
      const pangkatValue = isPanitia && (!pangkatGolRaw || upperText(pangkatGolRaw) === "#N/A") ? "" : pangkatGolRaw;

      return {
        no: idx + 1,
        nama: cleanText(p.nama),
        nik: cleanText(p.nik),
        jabatan,
        pangkat: pangkatValue,
        pangkatGol: pangkatValue,
        pangkat_gol: pangkatValue,
        kecamatan: cleanText(p.wilTugas),
        wil_tugas: cleanText(p.wilTugas),
        wilTugas: cleanText(p.wilTugas),
        tanggal_awal_kegiatan: tanggalAwal,
        tanggal_akhir_kegiatan: tanggalAkhir,
        lama: lamaHari,
      };
    }),
  };
}

function makeSlsKey(row) {
  return `${cleanText(row.kdkec)}|${cleanText(row.kddesa)}|${cleanText(row.kdsls)}|${cleanText(row.kdsubsls)}`;
}

async function createSpdBlob(templateUrl, formValues, peserta) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildSpdTemplateData(formValues || {}, peserta || []));
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

async function generateSpd(mainUrl, attachmentUrl, formValues, peserta) {
  const [blobMain, blobLampiran] = await Promise.all([
    createSpdBlob(mainUrl, formValues || {}, peserta || []),
    createSpdBlob(attachmentUrl, formValues || {}, peserta || []),
  ]);

  const safeNomor = formValues?.nomor_dokumen || formValues?.nomor || "SE2026";
  const safeTempat = formValues?.tempat || "SE2026";
  const safeGelombang = formValues?.gelombang || "X";
  const safeTanggal = formValues?.tanggal_surat || formValues?.tanggal || "";

  saveAs(blobMain, `SPD ${safeNomor} ${safeTempat} Gelombang ${safeGelombang} ${safeTanggal}.docx`);
  saveAs(blobLampiran, `Lampiran SPD ${safeNomor} ${safeTempat} Gelombang ${safeGelombang} ${safeTanggal}.docx`);
}

// SURAT TUGAS
function buildSuratTugasTemplateData(formValues, peserta = []) {
  const sorted       = sortPesertaByJabatanOrder(peserta || []);
  const tanggalAwal  = formatTanggalIndonesia(formValues.tanggal_awal_kegiatan || "");
  const tanggalAkhir = formatTanggalIndonesia(formValues.tanggal_akhir_kegiatan || "");
  const lamaHari     = calcDurationDays(formValues.tanggal_awal_kegiatan, formValues.tanggal_akhir_kegiatan) || "";
  const tanggalSurat = formatTanggalIndonesia(formValues.tanggal_surat || "");

  return {
    nomor_surat:            formValues.nomor_surat || "",
    nomor:                  formValues.nomor_surat || "",
    tanggal_aja:            tanggalSurat,
    tanggal_surat:          tanggalSurat,
    tanggal:                tanggalSurat,
    tempat:                 formValues.tempat || formValues.hotel || "",
    lokasi:                 formValues.lokasi || formValues.tempat || formValues.hotel || "",
    gelombang:              formValues.gelombang || "",
    kelas:                  formValues.kelas || "-",
    tanggal_awal_kegiatan:  tanggalAwal,
    tanggal_akhir_kegiatan: tanggalAkhir,
    lama:                   lamaHari,
    lama_hari:              lamaHari,
    jumlah_peserta:         sorted.length || 0,
    peserta: sorted.map((p, idx) => ({
      no:          idx + 1,
      nama:        cleanText(p.nama),
      nik:         cleanText(p.nik),
      jabatan:     cleanText(p.jabatan),
      pangkat:     cleanText(p.pangkatGol),
      pangkatGol:  cleanText(p.pangkatGol),
      pangkat_gol: cleanText(p.pangkatGol),
      kecamatan:   cleanText(p.wilTugas),
      wil_tugas:   cleanText(p.wilTugas),
      wilTugas:    cleanText(p.wilTugas),
    })),
  };
}

async function createSuratTugasBlob(templateUrl, formValues, peserta) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildSuratTugasTemplateData(formValues || {}, peserta || []));
  return doc.getZip().generate({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
}

async function generateSuratTugas(templateUrl, formValues, peserta) {
  const blob = await createSuratTugasBlob(templateUrl, formValues || {}, peserta || []);
  const safeNomor = formValues?.nomor_surat || "SE2026";
  const safeTempat = formValues?.tempat || "SE2026";
  const safeGelombang = formValues?.gelombang || "X";
  const safeTanggal = formValues?.tanggal_surat || "";
  saveAs(blob, `Surat Tugas ${safeNomor} ${safeTempat} Gelombang ${safeGelombang} ${safeTanggal}.docx`);
}

// LAMPIRAN
function formatKodeNama(kode, nama) {
  const kodeText = cleanText(kode);
  const namaText = cleanText(nama).toUpperCase();
  if (kodeText && namaText) return `${kodeText} ${namaText}`;
  return namaText || kodeText;
}

function groupLampiranRows(lampiranRows = [], jenis = "PML") {
  const isPml = upperText(jenis) === "PML";

  const map = new Map();

  for (const r of lampiranRows || []) {
    const namaPml = cleanText(r.nama_pml);
    const namaPpl = cleanText(r.nama_ppl);

    const kec = cleanText(r.kdkec);
    const desa = cleanText(r.kddesa);

    // 🔥 FIX: untuk Lampiran PML, satu PML bisa membawahi banyak PPL yang
    // kebetulan bertugas di kecamatan/kelurahan yang SAMA. Kalau key grouping
    // hanya namaPml|kec|desa, semua PPL berbeda dengan kec/desa sama akan
    // ditumpuk jadi SATU baris saja (nama PPL pertama menang, PPL lain hilang).
    // Maka nama_ppl WAJIB ikut jadi bagian key supaya setiap PPL tetap dapat
    // baris sendiri-sendiri.
    const keyBase = isPml
      ? `${namaPml}|${namaPpl}|${kec}|${desa}`
      : `${kec}|${desa}`;

    if (!map.has(keyBase)) {
      map.set(keyBase, {
        nama_pml: namaPml,
        nama_ppl: namaPpl,
        kdkec: kec,
        kddesa: desa,
        kecamatan: cleanText(r.kecamatan).toUpperCase(),
        kelurahan: cleanText(r.kelurahan).toUpperCase(),

        // Nomor kontrak dibawa dari baris pertama tiap grup. Diasumsikan konsisten
        // untuk satu petugas yang sama (PML/PPL yang sama harus punya 1 nomor kontrak).
        nomor_kontrak_pml: cleanText(r.nomor_kontrak_pml),
        nomor_kontrak_ppl: cleanText(r.nomor_kontrak_ppl),

        // 🔥 FIX: pakai SET untuk UNIQUE SLS
        slsSet: new Set(),
      });
    }

    const item = map.get(keyBase);

    const slsKey = makeSlsKey(r);
    item.slsSet.add(slsKey);
  }

  return [...map.values()]
    .map(v => ({
      ...v,
      jumlah: v.slsSet.size, // 🔥 FIX: UNIQUE SLS/SUBSLS
    }))
    .sort((a, b) => {
      // PML: urutkan per nama PPL (alfabetis) supaya tidak acak sesuai urutan baris sheet.
      // PPL: urutkan per kecamatan lalu kelurahan.
      if (isPml) {
        return cleanText(a.nama_ppl).localeCompare(cleanText(b.nama_ppl), "id-ID", { sensitivity: "base" });
      }
      const kecDiff = cleanText(a.kecamatan).localeCompare(cleanText(b.kecamatan), "id-ID", { sensitivity: "base" });
      if (kecDiff !== 0) return kecDiff;
      return cleanText(a.kelurahan).localeCompare(cleanText(b.kelurahan), "id-ID", { sensitivity: "base" });
    });
}

function buildLampiranTemplateData(formValues, lampiranRows = [], jenis = "PML") {
  // Fungsi ini menerima rows yang SUDAH dipisah per orang oleh generateLampiran().
  // Jadi output DOCX akan berisi satu petugas saja:
  // - PML: satu file per PENGAWAS
  // - PPL: satu file per PENCACAH
  const isPml = upperText(jenis) === "PML";
  const namaPetugas = isPml
    ? cleanText(lampiranRows?.[0]?.nama_pml)
    : cleanText(lampiranRows?.[0]?.nama_ppl);

  const grouped = groupLampiranRows(lampiranRows || [], jenis);

  // 🔥 BARU: variabel "sekali saja" untuk ditaruh di ATAS surat (di luar blok
  // {#peserta}...{/peserta}), bukan per baris tabel.
  // - nomor_kontrak: nomor kontrak orang ini (sama untuk semua barisnya, ambil satu saja
  //   dari baris pertama).
  // - total_jumlah: total SEMUA SLS/Sub-SLS milik orang ini, digabung dari semua
  //   kecamatan/kelurahan yang dia kerjakan (bukan cuma satu baris).
  // - total_jumlah_40 / total_jumlah_60: pembagian 40%/60% dari total_jumlah, dengan
  //   total_jumlah_60 = total_jumlah - total_jumlah_40 (supaya jumlahnya pas, tidak ada
  //   selisih pembulatan).
  const nomorKontrakDokumen = cleanText(
    isPml ? lampiranRows?.[0]?.nomor_kontrak_pml : lampiranRows?.[0]?.nomor_kontrak_ppl
  );
  const totalJumlahDokumen = grouped.reduce((sum, r) => sum + (r.jumlah || 0), 0);
  const totalJumlah40Dokumen = Math.ceil(totalJumlahDokumen * 0.4);
  const totalJumlah60Dokumen = totalJumlahDokumen - totalJumlah40Dokumen;

  return {
    jenis_lampiran: jenis,
    nama_petugas: namaPetugas,
    nama_pml: isPml ? namaPetugas : cleanText(lampiranRows?.[0]?.nama_pml),
    nama_ppl: !isPml ? namaPetugas : cleanText(lampiranRows?.[0]?.nama_ppl),
    tempat: formValues.tempat || formValues.hotel || "",
    hotel: formValues.hotel || formValues.tempat || "",
    gelombang: formValues.gelombang || "",
    kelas: formValues.kelas || "",
    jumlah_baris: grouped.length,

    // Variabel "sekali saja" di atas surat:
    nomor_kontrak: nomorKontrakDokumen,
    total_jumlah: totalJumlahDokumen,
    total_jumlah_40: totalJumlah40Dokumen,
    total_jumlah_60: totalJumlah60Dokumen,

    peserta: grouped.map((r, idx) => {
      // 🔥 FIX BARU: variabel tambahan untuk template.
      // - nomor_kontrak: ambil dari "No Kontrak PML" kalau jenisnya PML, atau
      //   "No Kontrak PPL" kalau jenisnya PPL.
      // - jumlah_40 / jumlah_60: pembagian 40%/60% dari jumlah SLS/Sub-SLS baris ini.
      //   jumlah_60 dihitung sebagai (total - jumlah_40), BUKAN dibulatkan sendiri-sendiri,
      //   supaya jumlah_40 + jumlah_60 selalu pas balik ke jumlah total (tidak ada selisih
      //   pembulatan kalau dijumlahkan manual).
      const totalJumlah = r.jumlah || 0;
      const jumlah40 = Math.ceil(totalJumlah * 0.4);
      const jumlah60 = totalJumlah - jumlah40;
      const nomorKontrak = isPml ? (r.nomor_kontrak_pml || "") : (r.nomor_kontrak_ppl || "");

      return {
        no: idx + 1,

        // Kompatibel dengan template lama:
        // - {nama_petugas} untuk nama utama
        // - {nama_pml} untuk PENGAWAS
        // - {nama_ppl} untuk PENCACAH
        nama_petugas: namaPetugas,
        nama_pml: isPml ? namaPetugas : r.nama_pml || "",
        nama_ppl: !isPml ? namaPetugas : r.nama_ppl || "",

        kecamatan: formatKodeNama(r.kdkec, r.kecamatan),
        kelurahan: formatKodeNama(r.kddesa, r.kelurahan),
        sls: r.sls || "",
        subsls: r.subsls || "",
        jumlah: totalJumlah,

        nomor_kontrak: nomorKontrak,
        jumlah_40: jumlah40,
        jumlah_60: jumlah60,
      };
    }),
  };
}

async function createLampiranBlob(templateUrl, formValues, lampiranRows, jenis) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template lampiran: ${response.status} ${response.statusText}`);

  const arrayBuffer = await response.arrayBuffer();
  return createLampiranBlobFromTemplateBuffer(arrayBuffer, formValues, lampiranRows, jenis);
}

// 🔥 FIX: dipisah dari createLampiranBlob() supaya template HANYA di-fetch SEKALI
// (lewat generateLampiran()), bukan di-fetch ulang dari network untuk SETIAP petugas.
// Untuk ratusan/ribuan petugas, fetch berulang ini sangat lambat dan boros memori.
// PizZip dibuat baru dari arrayBuffer yang sama setiap kali dipanggil — ini sesuai
// rekomendasi docxtemplater untuk batch generation (instance Docxtemplater TIDAK
// boleh dipakai ulang untuk render() berkali-kali, tapi arrayBuffer template-nya aman
// dipakai ulang berkali-kali untuk membuat PizZip baru).
function createLampiranBlobFromTemplateBuffer(templateArrayBuffer, formValues, lampiranRows, jenis) {
  const zip = new PizZip(templateArrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });

  doc.render(buildLampiranTemplateData(formValues || {}, lampiranRows || [], jenis));

  return doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

function sanitizeFileName(value) {
  return cleanText(value || "Tanpa Nama")
    .replace(/[\/:*?"<>|]/g, "-")
    .replace(/\s+/g, " ")
    .trim();
}

async function downloadMultipleAsZip(files, zipName = "dokumen.zip") {
  const zip = new JSZip();

  for (const file of files) {
    // file: { name, blob }
    zip.file(file.name, file.blob);
  }

  const content = await zip.generateAsync({ type: "blob" });
  saveAs(content, zipName);
}



// 🔥 FIX: jumlah PPL/PML bisa mencapai ribuan. Kalau semua dokumen ditahan di memori
// lalu di-zip jadi SATU file raksasa sekaligus, browser bisa kehabisan memori
// ("Array buffer allocation failed"). Maka proses zip dipecah per-batch — setiap
// batch jadi satu file .zip terpisah, sehingga beban memori di setiap tahap kecil
// dan terkendali. Bisa diturunkan lagi (misal 50-100) kalau template lampiran-nya berat.
const LAMPIRAN_ZIP_BATCH_SIZE = 150;

async function generateLampiran(templateUrl, formValues, lampiranRows, jenis) {
  const sourceRows = lampiranRows || [];
  const isPml = upperText(jenis) === "PML";

  // 🔥 FIX: jangan kelompokkan HANYA berdasarkan teks nama. Kalau ada 2 petugas
  // berbeda dengan nama yang sama persis (nama kembar), pengelompokan by-nama-saja
  // akan keliru menyatukan data SLS milik 2 orang berbeda jadi 1 dokumen — salah
  // satu identitas aslinya akan "tertelan" oleh yang lain.
  //
  // Solusinya: kunci pengelompokan diutamakan pakai EMAIL (kolom Email Pengawas /
  // Email Pencacah), karena email jauh lebih unik per-orang dibanding nama. Nama
  // tetap dipakai untuk ditampilkan di dokumen & sebagai dasar nama file. Kalau email
  // kosong di data, fallback ke nama saja (risiko nama kembar tetap ada untuk kasus
  // ini, tapi sudah jauh lebih baik daripada selalu mengandalkan nama).
  const groups = new Map(); // identity -> { displayName, rows: [] }

  for (const row of sourceRows) {
    const namaPml = cleanText(row.nama_pml || row.pengawas || row.nama || "");
    const namaPpl = cleanText(row.nama_ppl || row.pencacah || row.nama || "");
    const emailPml = cleanText(row.email_pengawas || "");
    const emailPpl = cleanText(row.email_pencacah || "");

    const displayName = isPml ? namaPml : namaPpl;
    if (!displayName) continue;

    const emailKey = upperText(isPml ? emailPml : emailPpl);
    const identity = emailKey || `NAMA::${upperText(displayName)}`;

    if (!groups.has(identity)) groups.set(identity, { displayName, rows: [] });
    groups.get(identity).rows.push(row);
  }

  if (groups.size === 0) {
    throw new Error(`Tidak ada data ${jenis}`);
  }

  // 🔥 FIX: deteksi nama kembar (identity berbeda tapi displayName sama persis) supaya
  // nama file tidak saling menimpa. Kalau ketemu, file dibedakan jadi "Nama (1).docx",
  // "Nama (2).docx", dst sesuai urutan kemunculan di data.
  const nameOccurrences = new Map();
  for (const { displayName } of groups.values()) {
    const key = upperText(displayName);
    nameOccurrences.set(key, (nameOccurrences.get(key) || 0) + 1);
  }
  const nameRunningIndex = new Map();
  const buildFileBaseName = (displayName) => {
    const key = upperText(displayName);
    const total = nameOccurrences.get(key) || 1;
    if (total <= 1) return sanitizeFileName(displayName);
    const idx = (nameRunningIndex.get(key) || 0) + 1;
    nameRunningIndex.set(key, idx);
    console.warn(`Lampiran ${jenis}: nama kembar terdeteksi -> "${displayName}" (salinan ke-${idx} dari ${total}, dibedakan via email).`);
    return `${sanitizeFileName(displayName)} (${idx})`;
  };

  // 🔥 FIX: fetch template SEKALI saja di sini, lalu arrayBuffer-nya dipakai ulang
  // untuk membuat setiap dokumen. Sebelumnya template di-fetch dari network untuk
  // SETIAP petugas (bisa 1000+ kali fetch untuk file yang sama) — sangat lambat.
  const templateResponse = await fetch(templateUrl);
  if (!templateResponse.ok) {
    throw new Error(`Gagal memuat template lampiran: ${templateResponse.status} ${templateResponse.statusText}`);
  }
  const templateArrayBuffer = await templateResponse.arrayBuffer();

  const entries = [...groups.values()]; // [{ displayName, rows }, ...]
  const totalBatches = Math.ceil(entries.length / LAMPIRAN_ZIP_BATCH_SIZE);

  console.log(
    `Lampiran ${jenis}: ${entries.length} petugas ditemukan, dipecah jadi ${totalBatches} file zip ` +
    `(maks ${LAMPIRAN_ZIP_BATCH_SIZE} dokumen/zip) untuk menghindari kehabisan memori.`
  );

  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
    const batchEntries = entries.slice(
      batchIndex * LAMPIRAN_ZIP_BATCH_SIZE,
      (batchIndex + 1) * LAMPIRAN_ZIP_BATCH_SIZE
    );

    const zipFiles = [];
    for (const { displayName, rows } of batchEntries) {
      const blob = createLampiranBlobFromTemplateBuffer(templateArrayBuffer, formValues || {}, rows, jenis);
      zipFiles.push({
        name: `Lampiran ${jenis} - ${buildFileBaseName(displayName)}.docx`,
        blob,
      });
    }

    const batchSuffix = totalBatches > 1 ? ` - Bagian ${batchIndex + 1} dari ${totalBatches}` : "";
    await downloadMultipleAsZip(
      zipFiles,
      `Lampiran ${jenis} - ${formValues?.gelombang || "SE2026"}${batchSuffix}.zip`
    );
  }
}

// Generate single lampiran (one person) — fetch template once and render one doc
async function generateSingleLampiran(templateUrl, formValues, lampiranRows, jenis, displayName) {
  if (!lampiranRows || lampiranRows.length === 0) throw new Error("Tidak ada data untuk lampiran yang dipilih");
  const templateResponse = await fetch(templateUrl);
  if (!templateResponse.ok) throw new Error(`Gagal memuat template lampiran: ${templateResponse.status} ${templateResponse.statusText}`);
  const templateArrayBuffer = await templateResponse.arrayBuffer();
  const blob = createLampiranBlobFromTemplateBuffer(templateArrayBuffer, formValues || {}, lampiranRows, jenis);
  const safeName = sanitizeFileName(displayName || (jenis + "-lampiran"));
  saveAs(blob, `Lampiran ${jenis} - ${safeName}.docx`);
}
// ─── MAIN APP ─────────────────────────────────────────────────────────────────

export default function PortalAdministrasiSE2026() {
  const [view,        setView]        = useState("dashboard");
  const [selectedDoc, setSelectedDoc] = useState(null);
  const [formData,    setFormData]    = useState({});
  const [previewData, setPreviewData] = useState(null);
  const [sidebarOpen, setSidebarOpen] = useState(false);

  const [petugasData,        setPetugasData]        = useState([]);
  const [lampiranData,       setLampiranData]       = useState([]);
  const [bappData,           setBappData]           = useState([]);
  const [xlsxLoaded,         setXlsxLoaded]         = useState(false);
  const [xlsxFileName,       setXlsxFileName]       = useState("data-petugas.xlsx");
  const [googleSheetUrl,     setGoogleSheetUrl]     = useState("https://docs.google.com/spreadsheets/d/10jA_NOMNn5pBuy1OPrSdHstscRrUOUlEDElk-jOmXLQ/edit?gid=1095810027#gid=1095810027");
  const [googleSheetApiKey,  setGoogleSheetApiKey]  = useState("");
  const [googleSheetError,   setGoogleSheetError]   = useState(null);
  const [googleSheetLoading, setGoogleSheetLoading] = useState(false);

  // Lampiran preview controls: pilih jenis (PML/PPL) dan pilih identity (email/name) untuk generate satu-per-orang
  const [lampiranPreviewJenis, setLampiranPreviewJenis] = useState("PML");
  const [lampiranPreviewIdentity, setLampiranPreviewIdentity] = useState("__ALL__");

  const loadGoogleSheetData = async () => {
    const normalized = normalizeGoogleSheetUrl(googleSheetUrl);
    if (!normalized) { setGoogleSheetError("URL Google Sheets tidak valid."); return; }
    setGoogleSheetError(null);
    setGoogleSheetLoading(true);
    try {
      const { data, lampiran, bappData: loadedBappData } = await loadGoogleSheet(normalized, googleSheetApiKey);
      if (data.length === 0 && (!lampiran || lampiran.length === 0) && (!loadedBappData || loadedBappData.length === 0)) {
        setGoogleSheetError("Tidak ada data petugas, lampiran, maupun BAPP yang terbaca.");
        return;
      }

      setPetugasData(data || []);
      setLampiranData(lampiran || []);
      setBappData(loadedBappData || []);
      setXlsxLoaded(true);
      setXlsxFileName(`${data.length} petugas, ${lampiran?.length || 0} lampiran, ${loadedBappData?.length || 0} BAPP`);
    } catch (err) {
      setGoogleSheetError(`Gagal memuat Google Sheet: ${err.message}`);
    } finally {
      setGoogleSheetLoading(false);
    }
  };

  React.useEffect(() => {
    const loadLocal = async () => {
      try {
        const response = await fetch("/data/data-petugas.xlsx");
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const buffer = await response.arrayBuffer();
        const data = parseXlsxData(buffer);
        const lampiran = parseLampiranXlsxData(buffer);
        const bappRows = parseBappData(buffer);
        setPetugasData(data || []);
        setLampiranData(lampiran || []);
        setBappData(bappRows || []);
        setXlsxLoaded(Boolean(lampiran?.length || data?.length || bappRows?.length));
      } catch (err) {
        console.warn("Tidak dapat memuat data-petugas.xlsx:", err.message);
      }
    };
    googleSheetUrl ? loadGoogleSheetData() : loadLocal();
  }, []);

  const openForm = (docType) => {
    if (docType?.disabled) {
      alert(docType.lockedMessage || "Fitur ini dikunci. Saat ini hanya Lampiran yang aktif.");
      return;
    }
    setSelectedDoc(docType);
    setFormData({});
    setPreviewData(null);
    setView("form");
  };
  const handleBack = () => {
    if (view === "preview") { setView("form"); setPreviewData(null); }
    else { setView("dashboard"); setSelectedDoc(null); }
  };
  const handleXlsxUpload = useCallback((file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const buffer = e.target.result;
        const data = parseXlsxData(buffer);
        const lampiran = parseLampiranXlsxData(buffer);
        const bappRows = parseBappData(buffer);
        setPetugasData(data || []);
        setLampiranData(lampiran || []);
        setBappData(bappRows || []);
        setXlsxLoaded(Boolean(lampiran?.length || data?.length || bappRows?.length));
        setXlsxFileName(file.name);
      } catch (err) { alert("Gagal membaca file xlsx: " + err.message); }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  return (
    <div className="min-h-screen overflow-x-hidden bg-[#fff8f0] text-slate-950 selection:bg-orange-200 selection:text-orange-950">
      <div className="pointer-events-none fixed inset-0 z-0 opacity-70">
        <div className="absolute left-[-12rem] top-[-10rem] h-[28rem] w-[28rem] rounded-full bg-orange-300/40 blur-3xl" />
        <div className="absolute right-[-10rem] top-[12rem] h-[30rem] w-[30rem] rounded-full bg-amber-300/30 blur-3xl" />
        <div className="absolute bottom-[-12rem] left-1/2 h-[30rem] w-[30rem] -translate-x-1/2 rounded-full bg-orange-400/20 blur-3xl" />
      </div>

      <header className="fixed left-0 right-0 top-0 z-50 border-b border-orange-100/80 bg-white/75 backdrop-blur-2xl">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-5 py-3.5 lg:px-8">
          <div className="flex items-center gap-3">
            <button className="rounded-xl border border-orange-100 bg-white p-2 text-slate-700 shadow-sm lg:hidden" onClick={() => setSidebarOpen(!sidebarOpen)}>
              <Menu size={20} />
            </button>
            <div className="flex items-center gap-3">
              <div className="flex h-11 w-11 items-center justify-center rounded-2xl bg-orange-500 shadow-lg shadow-orange-500/25">
                <FileText size={22} className="text-white" />
              </div>
              <div>
                <p className="text-sm font-black leading-none tracking-tight text-slate-950">BPS Kota Jakarta Timur</p>
                <p className="mt-1 text-xs font-semibold text-orange-600">Portal Administrasi SE2026</p>
              </div>
            </div>
          </div>
          {xlsxLoaded && (
            <div className="hidden items-center gap-2 rounded-full border border-green-200 bg-green-50 px-3 py-1.5 lg:flex">
              <CheckCircle size={14} className="text-green-600" />
              <span className="text-xs font-bold text-green-700">
                Peserta {petugasData.length || 0} • SLS {lampiranData.length || 0} • Petugas {bappData.length || 0}
              </span>
            </div>
          )}
        </div>
      </header>

      <AnimatePresence>
        {sidebarOpen && (
          <>
            <motion.div initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}
              className="fixed inset-0 z-40 bg-black/30 backdrop-blur-sm lg:hidden" onClick={() => setSidebarOpen(false)} />
            <motion.div initial={{ x: -280 }} animate={{ x: 0 }} exit={{ x: -280 }}
              transition={{ type: "spring", damping: 28, stiffness: 300 }}
              className="fixed left-0 top-0 z-50 flex h-full w-72 flex-col border-r border-orange-100 bg-white shadow-2xl lg:hidden">
              <div className="flex items-center justify-between border-b border-orange-100 p-5">
                <p className="font-black text-slate-950">Menu Dokumen</p>
                <button onClick={() => setSidebarOpen(false)} className="rounded-xl bg-orange-50 p-2 text-orange-600"><X size={18} /></button>
              </div>
              <div className="flex-1 overflow-y-auto p-4">
                <button onClick={() => { setView("dashboard"); setSelectedDoc(null); setSidebarOpen(false); }}
                  className={`mb-2 flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-sm font-bold transition ${view === "dashboard" && !selectedDoc ? "bg-orange-500 text-white" : "text-slate-700 hover:bg-orange-50"}`}>
                  <LayoutDashboard size={18} /> Dashboard
                </button>
                {DOC_TYPES.map((d) => (
                  <button key={d.id} onClick={() => { if (d.disabled) { alert(d.lockedMessage || "Fitur ini dikunci."); return; } openForm(d); setSidebarOpen(false); }}
                    className={`mb-2 flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-sm font-bold transition ${selectedDoc?.id === d.id ? "bg-orange-500 text-white" : "text-slate-700 hover:bg-orange-50"} ${d.disabled ? "cursor-not-allowed opacity-70" : ""}`}>
                    {React.cloneElement(d.icon, { size: 18 })} {d.label}
                    {d.disabled && <span className="ml-auto rounded-full bg-amber-100 px-2 py-0.5 text-[10px] font-black uppercase tracking-wider text-amber-700">Kunci</span>}
                  </button>
                ))}
              </div>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      <main className="relative z-10 pt-20">
        <AnimatePresence mode="wait">
          {view === "dashboard" && (
            <motion.div key="dashboard" initial={{ opacity: 0, y: 18 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} transition={{ duration: 0.4 }} className="px-5 py-10 lg:px-8">
              <div className="mx-auto max-w-7xl">
                <div className="mb-10">
                  <p className="text-sm font-black uppercase tracking-[0.25em] text-orange-600">Sensus Ekonomi 2026</p>
                  <h1 className="mt-3 text-5xl font-black tracking-[-0.04em] text-slate-950 lg:text-6xl">
                    Portal Administrasi<br />
                    <span className="bg-gradient-to-r from-orange-600 via-orange-500 to-amber-500 bg-clip-text text-transparent">Pelatihan Petugas</span>
                  </h1>
                  <p className="mt-5 max-w-2xl text-base leading-8 text-slate-600">Terbitkan dokumen administrasi secara cepat dan terstandar.</p>
                </div>
                {!xlsxLoaded && (
                  <>
                    <XlsxUploadCard loaded={xlsxLoaded} fileName={xlsxFileName} petugasCount={petugasData.length} onUpload={handleXlsxUpload} />
                    <GoogleSheetCard url={googleSheetUrl} apiKey={googleSheetApiKey} onUrlChange={setGoogleSheetUrl} onApiKeyChange={setGoogleSheetApiKey} onLoad={loadGoogleSheetData} loading={googleSheetLoading} error={googleSheetError} />
                  </>
                )}
                {xlsxLoaded && (
                  <motion.div initial={{ opacity: 0, y: -10 }} animate={{ opacity: 1, y: 0 }}
                    className="mb-8 flex items-center gap-4 rounded-3xl border border-blue-200 bg-blue-50 px-6 py-4 shadow-sm">
                    <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-blue-500 shadow-lg shadow-blue-500/25">
                      <CheckCircle size={22} className="text-white" />
                    </div>
                    <div>
                      <p className="font-black text-blue-900">Data Siap Digunakan</p>
                      <p className="text-sm font-semibold text-blue-600">Peserta {petugasData.length || 0} • SLS {lampiranData.length || 0} • Petugas {bappData.length || 0}</p>
                    </div>
                  </motion.div>
                )}
                <div className="mb-10 grid grid-cols-3 gap-4">
                  <StatCard value="3" label="Jenis Dokumen Aktif" />
                  <StatCard value="SE2026" label="Kegiatan" />
                  <StatCard value={bappData.length || "—"} label="Petugas Aktif" highlight={bappData.length > 0} />
                </div>
                <div className="grid gap-5 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4">
                  {DOC_TYPES.map((doc, i) => <DocCard key={doc.id} doc={doc} index={i} onSelect={() => openForm(doc)} />)}
                </div>
              </div>
            </motion.div>
          )}

          {view === "form" && selectedDoc && (
            <motion.div key="form" initial={{ opacity: 0, y: 18 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} transition={{ duration: 0.4 }} className="px-5 py-10 lg:px-8">
              <div className="mx-auto max-w-3xl">
                <button onClick={handleBack} className="mb-6 inline-flex items-center gap-2 rounded-2xl border border-orange-100 bg-white/80 px-4 py-2 text-sm font-bold text-slate-700 shadow-sm backdrop-blur transition hover:border-orange-200 hover:bg-white">
                  <ArrowLeft size={16} /> Kembali
                </button>
                <div className="mb-8 flex items-center gap-4">
                  <div className="flex h-16 w-16 items-center justify-center rounded-3xl bg-orange-500 text-white shadow-xl shadow-orange-500/25">
                    {React.cloneElement(selectedDoc.icon, { size: 30 })}
                  </div>
                  <div>
                    <p className="text-sm font-black uppercase tracking-[0.2em] text-orange-600">Formulir Penerbitan</p>
                    <h2 className="text-3xl font-black tracking-tight text-slate-950">{selectedDoc.label}</h2>
                  </div>
                </div>
                <DocForm docType={selectedDoc} formData={formData} setFormData={setFormData} onPreview={(data) => { setPreviewData(data); setView("preview"); }} petugasData={petugasData} lampiranData={lampiranData} bappData={bappData} xlsxLoaded={xlsxLoaded} />
              </div>
            </motion.div>
          )}

          {view === "preview" && previewData && (
            <motion.div key="preview" initial={{ opacity: 0, y: 18 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }} transition={{ duration: 0.4 }} className="px-5 py-10 lg:px-8">
              <div className="mx-auto max-w-4xl">
                <div className="mb-6 flex items-center justify-between">
                  <button onClick={handleBack} className="inline-flex items-center gap-2 rounded-2xl border border-orange-100 bg-white/80 px-4 py-2 text-sm font-bold text-slate-700 shadow-sm backdrop-blur transition hover:border-orange-200 hover:bg-white">
                    <ArrowLeft size={16} /> Kembali ke Formulir
                  </button>
                  <div className="flex gap-3">
                    <button onClick={() => window.print()} className="inline-flex items-center gap-2 rounded-2xl border border-orange-200 bg-white px-5 py-2.5 text-sm font-black text-orange-600 shadow-sm transition hover:-translate-y-0.5 hover:bg-orange-50">
                      <Printer size={16} /> Cetak
                    </button>
                    {selectedDoc?.id === "lampiran" && (
                      <div className="flex items-center gap-3">
                        <div className="flex flex-col sm:flex-row sm:items-center gap-2">
                            <select value={lampiranPreviewJenis} onChange={(e) => { setLampiranPreviewJenis(e.target.value); setLampiranPreviewIdentity("__ALL__"); }}
                              className="w-full sm:w-40 rounded-2xl border border-orange-100 bg-white/80 px-3 py-2 text-sm font-semibold text-slate-800 shadow-sm outline-none">
                            <option value="PML">PML</option>
                            <option value="PPL">PPL</option>
                          </select>

                            <select value={lampiranPreviewIdentity} onChange={(e) => setLampiranPreviewIdentity(e.target.value)}
                              className="w-full sm:w-72 rounded-2xl border border-orange-100 bg-white/80 px-3 py-2 text-sm font-semibold text-slate-800 shadow-sm outline-none">
                            <option value="__ALL__">— Semua —</option>
                            {(() => {
                              const rows = previewData?.lampiranRows || [];
                              const isPml = upperText(lampiranPreviewJenis) === "PML";
                              const seen = new Set();
                              return rows.map((r) => {
                                const displayName = isPml ? cleanText(r.nama_pml) : cleanText(r.nama_ppl);
                                if (!displayName) return null;
                                const email = cleanText(isPml ? r.email_pengawas : r.email_pencacah) || "";
                                const emailKey = upperText(email);
                                const identity = emailKey || `NAMA::${upperText(displayName)}`;
                                if (seen.has(identity)) return null;
                                seen.add(identity);
                                return <option key={identity} value={identity}>{displayName}</option>;
                              });
                            })()}
                          </select>

                            <button onClick={async () => {
                            try {
                              if (!previewData?.lampiranRows) throw new Error("Data lampiran belum tersedia");
                              const rows = previewData.lampiranRows || [];
                              const isPml = upperText(lampiranPreviewJenis) === "PML";
                              if (lampiranPreviewIdentity === "__ALL__") {
                                // generate all for this jenis as zip
                                await generateLampiran(isPml ? LAMPIRAN_PML_TEMPLATE_URL : LAMPIRAN_PPL_TEMPLATE_URL, previewData.formValues, rows, lampiranPreviewJenis);
                                return;
                              }
                              const filtered = [];
                              for (const r of rows) {
                                const displayName = isPml ? cleanText(r.nama_pml) : cleanText(r.nama_ppl);
                                const email = cleanText(isPml ? r.email_pengawas : r.email_pencacah) || "";
                                const emailKey = upperText(email);
                                const identity = emailKey || `NAMA::${upperText(displayName)}`;
                                if (identity === lampiranPreviewIdentity) filtered.push(r);
                              }
                              if (filtered.length === 0) throw new Error("Tidak ada data untuk identity yang dipilih");
                              const displayName = (isPml ? filtered[0].nama_pml : filtered[0].nama_ppl) || "Tanpa Nama";
                              await generateSingleLampiran(isPml ? LAMPIRAN_PML_TEMPLATE_URL : LAMPIRAN_PPL_TEMPLATE_URL, previewData.formValues, filtered, lampiranPreviewJenis, displayName);
                            } catch (err) { alert(err.message || err); }
                            }} type="button" className="w-full sm:w-auto inline-flex items-center justify-center gap-2 rounded-2xl bg-orange-500 px-4 py-2 text-sm font-black text-white shadow transition hover:bg-orange-600">Generate Terpilih</button>
                          </div>

                        <div className="flex flex-col sm:flex-row sm:items-center gap-2">
                          <GenerateDocxButton label="Unduh Semua Lampiran PML" onGenerate={() => generateLampiran(LAMPIRAN_PML_TEMPLATE_URL, previewData.formValues, previewData.lampiranRows, "PML")} />
                          <GenerateDocxButton label="Unduh Semua Lampiran PPL" onGenerate={() => generateLampiran(LAMPIRAN_PPL_TEMPLATE_URL, previewData.formValues, previewData.lampiranRows, "PPL")} />
                        </div>
                      </div>
                    )}
                  </div>
                </div>
                <DocPreview docType={selectedDoc} data={previewData} />
              </div>
            </motion.div>
          )}
        </AnimatePresence>
      </main>

      <footer className="relative z-10 mt-10 border-t border-orange-100 bg-white/90 backdrop-blur-xl">
        <div className="mx-auto max-w-7xl px-5 py-6 lg:px-8">
          <div className="flex flex-col items-center justify-between gap-2 text-center text-xs font-semibold text-slate-500 sm:flex-row sm:text-left">
            <p>© 2026 BPS Kota Jakarta Timur — Portal Administrasi SE2026</p>
            <p>Sistem Penerbitan Dokumen Pelatihan Petugas</p>
          </div>
        </div>
      </footer>
    </div>
  );
}

// ─── XLSX UPLOAD CARD ─────────────────────────────────────────────────────────

function XlsxUploadCard({ loaded, fileName, petugasCount, onUpload }) {
  const fileRef = useRef(null);
  const [dragging, setDragging] = useState(false);
  const handleDrop = (e) => {
    e.preventDefault(); setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith(".xlsx")) onUpload(file);
  };
  return (
    <div className="mb-8">
      {loaded ? (
        <motion.div initial={{ opacity: 0, scale: 0.97 }} animate={{ opacity: 1, scale: 1 }}
          className="flex items-center justify-between rounded-3xl border border-green-200 bg-green-50 px-6 py-4 shadow-sm">
          <div className="flex items-center gap-4">
            <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-green-500 shadow-lg shadow-green-500/25"><CheckCircle size={22} className="text-white" /></div>
            <div>
              <p className="font-black text-green-800">Data Petugas Berhasil Dimuat</p>
              <p className="text-sm font-semibold text-green-600">{fileName} — {petugasCount} petugas</p>
            </div>
          </div>
          <button onClick={() => fileRef.current?.click()} className="rounded-2xl border border-green-200 bg-white px-4 py-2 text-sm font-bold text-green-700 transition hover:bg-green-50">Ganti File</button>
          <input ref={fileRef} type="file" accept=".xlsx" className="hidden" onChange={(e) => onUpload(e.target.files[0])} />
        </motion.div>
      ) : (
        <div onDragOver={(e) => { e.preventDefault(); setDragging(true); }} onDragLeave={() => setDragging(false)} onDrop={handleDrop} onClick={() => fileRef.current?.click()}
          className={`cursor-pointer rounded-3xl border-2 border-dashed p-8 text-center transition ${dragging ? "border-orange-400 bg-orange-50" : "border-orange-200 bg-white/60 hover:border-orange-300 hover:bg-orange-50/50"}`}>
          <div className="mx-auto mb-4 flex h-14 w-14 items-center justify-center rounded-2xl bg-orange-100"><Upload size={24} className="text-orange-600" /></div>
          <p className="font-black text-slate-800">Unggah Data Petugas (Opsional)</p>
          <p className="mt-1 text-sm font-semibold text-slate-500">Drag & drop file <span className="text-orange-600">data-petugas.xlsx</span> di sini untuk mengganti.</p>
          <p className="mt-2 text-xs text-slate-400">Kolom: Nama, NIK, Jabatan, Wil. Tugas, Pangkat/Gol, Kelas, Gelombang</p>
          <input ref={fileRef} type="file" accept=".xlsx" className="hidden" onChange={(e) => onUpload(e.target.files[0])} />
        </div>
      )}
    </div>
  );
}

function GoogleSheetCard({ url, apiKey, onUrlChange, onApiKeyChange, onLoad, loading, error }) {
  return (
    <motion.div initial={{ opacity: 0, y: -6 }} animate={{ opacity: 1, y: 0 }} className="mb-8 rounded-3xl border border-orange-200 bg-white p-6 shadow-sm">
      <div className="flex items-center gap-4">
        <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-orange-100"><MapPin size={22} className="text-orange-600" /></div>
        <div>
          <p className="font-black text-slate-900">Load dari Google Spreadsheet</p>
          <p className="text-sm text-slate-500">Gunakan sheet dengan kolom TC, NIK, Jabatan, Wil. Tugas, Pangkat/Gol, Kelas, Gelombang.</p>
        </div>
      </div>
      <div className="mt-4 grid gap-3 sm:grid-cols-[1fr_auto]">
        <input value={url} onChange={(e) => onUrlChange(e.target.value)} placeholder="https://docs.google.com/spreadsheets/d/ID_SHEET/edit#gid=0"
          className="w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm text-slate-800 outline-none transition focus:border-orange-300 focus:bg-white" />
        <button type="button" onClick={onLoad} disabled={loading || !url}
          className="rounded-2xl bg-orange-600 px-5 py-3 text-sm font-bold text-white transition enabled:hover:bg-orange-700 disabled:cursor-not-allowed disabled:bg-orange-200">
          {loading ? "Memuat..." : "Muat"}
        </button>
      </div>
      <div className="mt-3">
        <input value={apiKey || ""} onChange={(e) => onApiKeyChange(e.target.value)} placeholder="API Key Google Sheets (opsional)"
          className="w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm text-slate-800 outline-none transition focus:border-orange-300 focus:bg-white" />
      </div>
      {error && <p className="mt-3 text-sm text-red-600">{error}</p>}
      <p className="mt-3 text-xs text-slate-400">Jika API key diisi, data dibaca langsung dari Google Sheets API. Jika tidak, aplikasi memakai fallback export XLSX.</p>
    </motion.div>
  );
}

// ─── GENERATE DOCX BUTTON ────────────────────────────────────────────────────

function GenerateDocxButton({ onGenerate, label = "Unduh .docx" }) {
  const [loading, setLoading] = useState(false);
  const [error, setError]     = useState(null);
  const handleGenerate = async () => {
    setLoading(true); setError(null);
    try { await onGenerate(); } catch (err) { setError(err.message); } finally { setLoading(false); }
  };
  return (
    <div className="flex flex-col items-end gap-1">
      <button type="button" onClick={handleGenerate} disabled={loading}
        className="inline-flex items-center gap-2 rounded-2xl bg-orange-500 px-5 py-2.5 text-sm font-black text-white shadow-xl shadow-orange-500/20 transition hover:-translate-y-0.5 hover:bg-orange-600 disabled:opacity-60">
        <Download size={16} />
        {loading ? "Membuat..." : label}
      </button>
      {error && <p className="flex items-center gap-1 text-xs font-semibold text-red-500"><AlertCircle size={12} /> {error}</p>}
    </div>
  );
}

// ─── STAT + DOC CARDS ────────────────────────────────────────────────────────

function StatCard({ value, label, highlight }) {
  return (
    <div className={`rounded-3xl border p-5 shadow-sm backdrop-blur ${highlight ? "border-green-200 bg-green-50/80" : "border-orange-100 bg-white/75"}`}>
      <p className={`text-3xl font-black ${highlight ? "text-green-700" : "text-slate-950"}`}>{value}</p>
      <p className="mt-1 text-xs font-bold uppercase tracking-wider text-slate-500">{label}</p>
    </div>
  );
}

function DocCard({ doc, index, onSelect }) {
  return (
    <motion.button initial={{ opacity: 0, y: 24 }} animate={{ opacity: 1, y: 0 }} transition={{ duration: 0.5, delay: index * 0.07 }}
      onClick={doc.disabled ? undefined : onSelect}
      disabled={doc.disabled}
      className={`group relative overflow-hidden rounded-[2rem] border border-orange-100 bg-white/85 p-6 text-left shadow-lg shadow-orange-900/5 backdrop-blur transition ${doc.disabled ? "cursor-not-allowed opacity-70" : "hover:-translate-y-2 hover:shadow-2xl hover:shadow-orange-500/15"}`}>
      <div className="absolute right-0 top-0 h-24 w-24 rounded-bl-[4rem] bg-orange-50 transition group-hover:bg-orange-100" />
      <div className="relative flex h-14 w-14 items-center justify-center rounded-2xl bg-orange-500 text-white shadow-xl shadow-orange-500/25">
        {React.cloneElement(doc.icon, { size: 26 })}
      </div>
      <h3 className="relative mt-5 text-lg font-black tracking-tight text-slate-950">{doc.label}</h3>
      <p className="relative mt-2 text-sm leading-6 text-slate-500">{doc.desc}</p>
      <div className="relative mt-5 flex items-center gap-1 text-sm font-black text-orange-500">
        {doc.disabled ? "Terkunci" : "Buat Dokumen"} <ChevronRight size={16} className={`transition ${doc.disabled ? "" : "group-hover:translate-x-1"}`} />
      </div>
    </motion.button>
  );
}

// ─── PESERTA TABLE PREVIEW ────────────────────────────────────────────────────

function PesertaTablePreview({ peserta }) {
  if (!peserta || peserta.length === 0) return null;
  return (
    <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}>
      <div className="overflow-hidden rounded-2xl border border-orange-100 bg-white/80">
        <div className="grid grid-cols-[2rem_1fr_1fr_1fr] border-b border-orange-100 bg-orange-500 px-4 py-2 text-xs font-black uppercase text-white">
          <span>No</span><span>Nama</span><span>Jabatan</span><span>Wil. Tugas</span>
        </div>
        <div className="max-h-64 divide-y divide-orange-50 overflow-y-auto">
          {peserta.map((p, i) => (
            <div key={i} className="grid grid-cols-[2rem_1fr_1fr_1fr] px-4 py-2.5 text-xs font-semibold text-slate-700 hover:bg-orange-50/50">
              <span className="text-slate-400">{i + 1}</span>
              <span>{p.nama}</span>
              <span>{p.jabatan}</span>
              <span>{p.wilTugas}</span>
            </div>
          ))}
        </div>
      </div>
    </motion.div>
  );
}

// ─── FILTER PESERTA PANEL ─────────────────────────────────────────────────────

function FilterPesertaPanel({ xlsxLoaded, formData, setFormData, petugasData, mode = "grouped", selectedGroup = "", onFilterResult, prependRow = null }) {
  const inputCls = "w-full rounded-2xl border border-orange-100 bg-white/80 px-4 py-3 text-sm font-semibold text-slate-800 shadow-sm outline-none focus:border-orange-300 focus:ring-2 focus:ring-orange-100 transition";
  const labelCls = "mb-2 block text-xs font-black uppercase tracking-[0.2em] text-slate-500";

  const [filtered,        setFiltered]        = useState(false);
  const [filteredPeserta, setFilteredPeserta] = useState([]);
  const [namaInda,        setNamaInda]        = useState("");

  const showKelas = selectedGroup !== "panitia-inda";

  const selectedHotel     = cleanText(formData.hotel);
  const selectedKelas     = cleanText(formData.kelas);
  const selectedGelombang = cleanText(formData.gelombang);

  const hotelOptions = React.useMemo(() => uniqueSorted([...petugasData.map((p) => p.hotel), "STIS"]), [petugasData]);
  const gelombangOptions = React.useMemo(() => {
    if (cleanText(selectedHotel).toLowerCase() === "stis") return ["4"];
    return uniqueSorted(petugasData.filter((p) => !selectedHotel || cleanText(p.hotel) === selectedHotel).map((p) => p.gelombang));
  }, [petugasData, selectedHotel]);
  const kelasOptions = React.useMemo(() =>
    uniqueSorted(petugasData
      .filter((p) => !selectedHotel || cleanText(p.hotel) === selectedHotel)
      .filter((p) => !selectedGelombang || cleanText(p.gelombang) === selectedGelombang)
      .map((p) => p.kelas)),
    [petugasData, selectedHotel, selectedGelombang]
  );

  const resetFilterResult = () => {
    setFilteredPeserta([]); setNamaInda(""); setFiltered(false);
    onFilterResult([], "", "");
  };

  const runFilter = () => {
    if (!xlsxLoaded) { alert("Data petugas (.xlsx) belum berhasil dimuat."); return; }
    if (!selectedHotel || !selectedGelombang) { alert("Pilih Tempat dan Gelombang terlebih dahulu."); return; }
    if (showKelas && !selectedKelas) { alert("Pilih Kelas terlebih dahulu."); return; }

    const baseRows = petugasData.filter((p) =>
      cleanText(p.hotel)     === selectedHotel &&
      cleanText(p.gelombang) === selectedGelombang &&
      (showKelas ? cleanText(p.kelas) === selectedKelas : true)
    );

    const inda = baseRows.find((p) => upperText(p.jabatan) === "INDA");
    let hasil;
    if (mode === "all") {
      hasil = sortPesertaByJabatanOrder(baseRows.filter((p) => ["PML", "PPL"].includes(upperText(p.jabatan))));
    } else if (selectedGroup) {
      hasil = sortPesertaByJabatanOrder(baseRows.filter((p) => jabatanMasukGroup(p.jabatan, selectedGroup)));
    } else {
      hasil = sortPesertaByJabatanOrder(baseRows);
    }

    const displayPeserta = hasil.length > 0 && prependRow ? [prependRow, ...hasil] : hasil;
    setNamaInda(inda?.nama || "");
    setFilteredPeserta(displayPeserta);
    setFiltered(true);
    onFilterResult(hasil, inda?.nama || "", selectedGroup);
  };

  return (
    <div className="rounded-3xl border border-orange-100 bg-orange-50/60 p-5 space-y-4">
      <p className="flex items-center gap-2 text-xs font-black uppercase tracking-[0.2em] text-orange-700">
        <Filter size={14} /> Parameter Filter Peserta
      </p>

      <div>
        <label className={labelCls}>Tempat</label>
        <select className={inputCls} value={formData.hotel || ""}
          onChange={(e) => { const hotel = e.target.value; setFormData((prev) => ({ ...prev, hotel, tempat: hotel, kelas: "", gelombang: "" })); resetFilterResult(); }}
          disabled={!xlsxLoaded || hotelOptions.length === 0}>
          <option value="">Pilih tempat</option>
          {hotelOptions.map((h) => <option key={h} value={h}>{h}</option>)}
        </select>
        <p className="mt-1 text-xs font-semibold text-slate-400">Daftar tempat diambil dari kolom TC pada data XLSX.</p>
      </div>

      <div className={`grid gap-4 ${showKelas ? "sm:grid-cols-2" : ""}`}>
        <div>
          <label className={labelCls}>Gelombang</label>
          <select className={inputCls} value={formData.gelombang || ""}
            onChange={(e) => { setFormData((prev) => ({ ...prev, gelombang: e.target.value, kelas: "" })); resetFilterResult(); }}
            disabled={!selectedHotel}>
            <option value="">Pilih gelombang</option>
            {gelombangOptions.map((g) => <option key={g} value={g}>{g}</option>)}
          </select>
        </div>
        {showKelas && (
          <div>
            <label className={labelCls}>Kelas</label>
            <select className={inputCls} value={formData.kelas || ""}
              onChange={(e) => { setFormData((prev) => ({ ...prev, kelas: e.target.value })); resetFilterResult(); }}
              disabled={!selectedHotel || !selectedGelombang}>
              <option value="">Pilih kelas</option>
              {kelasOptions.map((k) => <option key={k} value={k}>{k}</option>)}
            </select>
          </div>
        )}
      </div>

      {!showKelas && (
        <div className="flex items-center gap-2 rounded-2xl border border-orange-100 bg-white/70 px-4 py-3 text-xs font-semibold text-slate-500">
          <AlertCircle size={14} className="shrink-0 text-orange-400" />
          Untuk Panitia &amp; Inda, kelas tidak diperlukan. Variabel kelas akan otomatis diisi "-" pada dokumen.
        </div>
      )}

      <button type="button" onClick={runFilter}
        className={`w-full inline-flex items-center justify-center gap-2 rounded-2xl px-5 py-3 text-sm font-black shadow-lg transition hover:-translate-y-0.5 ${filtered ? "bg-orange-600 text-white shadow-orange-500/25" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
        <Users size={16} /> Tampilkan Peserta
      </button>

      {filtered && (
        <div className="flex items-center gap-2 rounded-2xl border border-green-200 bg-green-50 px-4 py-3">
          <CheckCircle size={16} className="text-green-600" />
          <span className="text-sm font-bold text-green-700">
            {filteredPeserta.length} peserta ditemukan{selectedGroup && DAFTAR_HADIR_GROUPS[selectedGroup] ? ` — ${DAFTAR_HADIR_GROUPS[selectedGroup].label}` : ""}
          </span>
        </div>
      )}

      {filtered && filteredPeserta.length > 0 && <PesertaTablePreview peserta={filteredPeserta} />}
      {filtered && filteredPeserta.length === 0 && (
        <div className="rounded-2xl border border-dashed border-orange-200 bg-orange-50/50 p-6 text-center text-sm font-semibold text-slate-400">
          Tidak ada peserta untuk Tempat {formData.hotel}{showKelas ? `, Kelas ${formData.kelas}` : ""}, Gelombang {formData.gelombang}
        </div>
      )}
    </div>
  );
}

// ─── FILTER HOTEL PANEL ───────────────────────────────────────────────────────

function FilterPesertaHotelPanel({ xlsxLoaded, formData, setFormData, petugasData, onFilterResult }) {
  const inputCls = "w-full rounded-2xl border border-orange-100 bg-white/80 px-4 py-3 text-sm font-semibold text-slate-800 shadow-sm outline-none focus:border-orange-300 focus:ring-2 focus:ring-orange-100 transition";
  const labelCls = "mb-2 block text-xs font-black uppercase tracking-[0.2em] text-slate-500";
  const [filtered, setFiltered]               = useState(false);
  const [filteredPeserta, setFilteredPeserta] = useState([]);
  const selectedHotel = cleanText(formData.hotel);
  const hotelOptions  = React.useMemo(() => uniqueSorted([...petugasData.map((p) => p.hotel), "STIS"]), [petugasData]);

  const runFilter = () => {
    if (!xlsxLoaded) { alert("Data petugas belum dimuat."); return; }
    if (!selectedHotel) { alert("Pilih tempat terlebih dahulu."); return; }
    const hasil = sortPesertaByJabatanOrder(petugasData.filter((p) => cleanText(p.hotel) === selectedHotel));
    setFilteredPeserta(hasil); setFiltered(true); onFilterResult(hasil);
  };
  const resetFilter = () => { setFiltered(false); setFilteredPeserta([]); onFilterResult([]); };

  return (
    <div className="rounded-3xl border border-orange-100 bg-orange-50/60 p-5 space-y-4">
      <p className="flex items-center gap-2 text-xs font-black uppercase tracking-[0.2em] text-orange-700"><Filter size={14} /> Filter Peserta Berdasarkan Tempat</p>
      <div>
        <label className={labelCls}>Tempat</label>
        <select className={inputCls} value={formData.hotel || ""}
          onChange={(e) => { setFormData((prev) => ({ ...prev, hotel: e.target.value, tempat: e.target.value })); resetFilter(); }}
          disabled={!xlsxLoaded || hotelOptions.length === 0}>
          <option value="">Pilih tempat</option>
          {hotelOptions.map((h) => <option key={h} value={h}>{h}</option>)}
        </select>
      </div>
      <button type="button" onClick={runFilter}
        className={`w-full inline-flex items-center justify-center gap-2 rounded-2xl px-5 py-3 text-sm font-black shadow-lg transition hover:-translate-y-0.5 ${filtered ? "bg-orange-600 text-white shadow-orange-500/25" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
        <Users size={16} /> Tampilkan Peserta
      </button>
      {filtered && <div className="rounded-2xl border border-green-200 bg-green-50 px-4 py-3 text-sm font-bold text-green-700">{filteredPeserta.length} peserta ditemukan untuk tempat {selectedHotel}</div>}
      {filtered && filteredPeserta.length > 0 && <PesertaTablePreview peserta={filteredPeserta} />}
    </div>
  );
}

// ─── FILTER HOTEL + GELOMBANG PANEL ──────────────────────────────────────────

function FilterPesertaHotelGelombangPanel({ xlsxLoaded, formData, setFormData, petugasData, onFilterResult }) {
  const inputCls = "w-full rounded-2xl border border-orange-100 bg-white/80 px-4 py-3 text-sm font-semibold text-slate-800 shadow-sm outline-none focus:border-orange-300 focus:ring-2 focus:ring-orange-100 transition";
  const labelCls = "mb-2 block text-xs font-black uppercase tracking-[0.2em] text-slate-500";
  const [filtered, setFiltered]               = useState(false);
  const [filteredPeserta, setFilteredPeserta] = useState([]);
  const selectedHotel     = cleanText(formData.hotel);
  const selectedGelombang = cleanText(formData.gelombang);
  const hotelOptions      = React.useMemo(() => uniqueSorted([...petugasData.map((p) => p.hotel), "STIS"]), [petugasData]);
  const gelombangOptions  = React.useMemo(() => {
    if (cleanText(selectedHotel).toLowerCase() === "stis") return ["4"];
    return uniqueSorted(petugasData.filter((p) => !selectedHotel || cleanText(p.hotel) === selectedHotel).map((p) => p.gelombang));
  }, [petugasData, selectedHotel]);

  const runFilter = () => {
    if (!xlsxLoaded) { alert("Data petugas belum dimuat."); return; }
    if (!selectedHotel) { alert("Pilih tempat."); return; }
    if (!selectedGelombang) { alert("Pilih gelombang."); return; }
    const hasil = sortPesertaByJabatanOrder(petugasData.filter((p) => cleanText(p.hotel) === selectedHotel && cleanText(p.gelombang) === selectedGelombang));
    setFilteredPeserta(hasil); setFiltered(true); onFilterResult(hasil);
  };
  const resetFilter = () => { setFiltered(false); setFilteredPeserta([]); onFilterResult([]); };

  return (
    <div className="rounded-3xl border border-orange-100 bg-orange-50/60 p-5 space-y-4">
      <p className="flex items-center gap-2 text-xs font-black uppercase tracking-[0.2em] text-orange-700"><Filter size={14} /> Filter Peserta — Tempat &amp; Gelombang</p>
      <div>
        <label className={labelCls}>Tempat</label>
        <select className={inputCls} value={formData.hotel || ""}
          onChange={(e) => { setFormData((prev) => ({ ...prev, hotel: e.target.value, tempat: e.target.value, gelombang: "" })); resetFilter(); }}
          disabled={!xlsxLoaded || hotelOptions.length === 0}>
          <option value="">Pilih tempat</option>
          {hotelOptions.map((h) => <option key={h} value={h}>{h}</option>)}
        </select>
      </div>
      <div>
        <label className={labelCls}>Gelombang</label>
        <select className={inputCls} value={formData.gelombang || ""}
          onChange={(e) => { setFormData((prev) => ({ ...prev, gelombang: e.target.value })); resetFilter(); }}
          disabled={!selectedHotel}>
          <option value="">Pilih gelombang</option>
          {gelombangOptions.map((g) => <option key={g} value={g}>{g}</option>)}
        </select>
      </div>
      <button type="button" onClick={runFilter}
        className={`w-full inline-flex items-center justify-center gap-2 rounded-2xl px-5 py-3 text-sm font-black shadow-lg transition hover:-translate-y-0.5 ${filtered ? "bg-orange-600 text-white shadow-orange-500/25" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
        <Users size={16} /> Tampilkan Peserta
      </button>
      {filtered && <div className="rounded-2xl border border-green-200 bg-green-50 px-4 py-3 text-sm font-bold text-green-700">{filteredPeserta.length} peserta — {selectedHotel}, Gelombang {selectedGelombang}</div>}
      {filtered && filteredPeserta.length > 0 && <PesertaTablePreview peserta={filteredPeserta} />}
    </div>
  );
}

// ─── DOC FORM ─────────────────────────────────────────────────────────────────

function DocForm({ docType, formData, setFormData, onPreview, petugasData, lampiranData = [], bappData = [], xlsxLoaded }) {
  const update = (key, val) => setFormData((p) => ({ ...p, [key]: val }));

  const [daftarHadirPeserta,     setDaftarHadirPeserta]     = useState([]);
  const [daftarHadirNamaInda,    setDaftarHadirNamaInda]    = useState("");
  const [daftarHadirFiltered,    setDaftarHadirFiltered]    = useState(false);
  const [daftarHadirFilterGroup, setDaftarHadirFilterGroup] = useState("");

  const [tandaTerimaPeserta,      setTandaTerimaPeserta]      = useState([]);
  const [tandaTerimaFiltered,     setTandaTerimaFiltered]     = useState(false);
  const [tandaTerimaType,         setTandaTerimaType]         = useState("");
  const [tandaTerimaWilTugas,     setTandaTerimaWilTugas]     = useState("");
  const [tandaTerimaLokasi,       setTandaTerimaLokasi]       = useState("BPS Kota Jakarta Timur");
  const [suratPernyataanPeserta,  setSuratPernyataanPeserta]  = useState([]);
  const [suratPernyataanFiltered, setSuratPernyataanFiltered] = useState(false);
  const [suratPernyataanFilterGroup, setSuratPernyataanFilterGroup] = useState("");
  const [suratTugasPeserta,       setSuratTugasPeserta]       = useState([]);
  const [suratTugasFiltered,      setSuratTugasFiltered]      = useState(false);
  const [spjPeserta,              setSpjPeserta]              = useState([]);
  const [spjFiltered,             setSpjFiltered]             = useState(false);
  const [spjFilterGroup,          setSpjFilterGroup]          = useState("");
  const [spdPeserta,              setSpdPeserta]              = useState([]);
  const [spdFiltered,             setSpdFiltered]             = useState(false);
  const [pengeluaranPeserta,      setPengeluaranPeserta]      = useState([]);
  const [pengeluaranFiltered,     setPengeluaranFiltered]     = useState(false);
  const [pengeluaranFilterGroup,  setPengeluaranFilterGroup]  = useState("");
  const [bappRole,                setBappRole]                = useState("");
  const [bappManualSelect,        setBappManualSelect]        = useState("");
  const [bappGenerating,          setBappGenerating]          = useState(false);
  const [bappProgressText,        setBappProgressText]        = useState("");
  const [bastRole,                setBastRole]                = useState("");
  const [bastManualSelect,        setBastManualSelect]        = useState("");
  const [bastGenerating,          setBastGenerating]          = useState(false);
  const [bastProgressText,        setBastProgressText]        = useState("");
  const [suratPenyelesaianLapanganSelect, setSuratPenyelesaianLapanganSelect] = useState("");
  const [suratPenyelesaianLapanganGenerating, setSuratPenyelesaianLapanganGenerating] = useState(false);
  const [suratPenyelesaianLapanganProgressText, setSuratPenyelesaianLapanganProgressText] = useState("");

  // Lampiran form controls: combined manual select for PML + PPL
  const [lampiranManualSelect, setLampiranManualSelect] = useState("");

  React.useEffect(() => {
    if (docType.id === "daftar-hadir" && !formData.jamMulai && !formData.jamSelesai) {
      setFormData((prev) => ({ ...prev, jamMulai: "07.30", jamSelesai: "18.00" }));
    }
  }, [docType.id]);

  const inputCls = "w-full rounded-2xl border border-orange-100 bg-white/80 px-4 py-3 text-sm font-semibold text-slate-800 shadow-sm outline-none focus:border-orange-300 focus:ring-2 focus:ring-orange-100 transition";
  const labelCls = "mb-2 block text-xs font-black uppercase tracking-[0.2em] text-slate-500";

  const handleSubmit = (e) => {
    e.preventDefault();

    if (docType.id === "daftar-hadir") {
      if (!daftarHadirFilterGroup) {
        alert("Pilih kelompok peserta (PML & PPL atau Panitia & Inda) terlebih dahulu.");
        return;
      }
      if (!daftarHadirFiltered || daftarHadirPeserta.length === 0) {
        alert("Tampilkan peserta terlebih dahulu dengan mengklik tombol filter.");
        return;
      }
      const isPanitiaInda = daftarHadirFilterGroup === "panitia-inda";
      onPreview({
        formValues: {
          ...formData,
          tempat:          formData.tempat || formData.hotel || "",
          kelompokPeserta: DAFTAR_HADIR_GROUPS[daftarHadirFilterGroup]?.label || "",
          kelas:           isPanitiaInda ? "-" : (formData.kelas || ""),
        },
        peserta:             daftarHadirPeserta,
        namaInda:            daftarHadirNamaInda,
        selectedFilterGroup: daftarHadirFilterGroup,
      });
      return;
    }

    if (docType.id === "tanda-terima") {
      if (!tandaTerimaType) { alert("Pilih jenis tanda terima terlebih dahulu."); return; }
      if (!tandaTerimaFiltered || tandaTerimaPeserta.length === 0) { alert("Tampilkan peserta terlebih dahulu."); return; }
      onPreview({ formValues: { ...formData, tempat: tandaTerimaType === "mitra-umum" ? tandaTerimaLokasi : (formData.tempat || formData.hotel || "") }, peserta: tandaTerimaPeserta, tandaTerimaType: tandaTerimaType });
      return;
    }

    if (docType.id === "surat-pernyataan-kendaraan") {
      if (!suratPernyataanFilterGroup) {
        alert("Pilih kelompok peserta (PML & PPL atau Panitia & Inda) terlebih dahulu.");
        return;
      }
      if (!suratPernyataanFiltered || suratPernyataanPeserta.length === 0) { alert("Tampilkan peserta terlebih dahulu."); return; }
      const isPanitiaInda = suratPernyataanFilterGroup === "panitia-inda";
      onPreview({
        formValues: {
          tanggal_surat: formData.tanggal_surat || "",
          tempat:        formData.tempat || formData.hotel || "",
          hotel:         formData.hotel || "",
          gelombang:     formData.gelombang || "",
          kelompokPeserta: suratPernyataanFilterGroup === "panitia-inda" ? "Panitia & Inda" : "Petugas PML & PPL",
          kelas:         isPanitiaInda ? "-" : (formData.kelas || "-"),
        },
        peserta: suratPernyataanPeserta,
      });
      return;
    }

    if (docType.id === "pengeluaran-riil") {
      if (!pengeluaranFilterGroup) {
        alert("Pilih kelompok peserta (PML & PPL atau Panitia & Inda) terlebih dahulu.");
        return;
      }
      if (!pengeluaranFiltered || pengeluaranPeserta.length === 0) { alert("Tampilkan peserta terlebih dahulu."); return; }
      const isPanitiaInda = pengeluaranFilterGroup === "panitia-inda";
      onPreview({
        formValues: {
          tanggal_surat: formData.tanggal_surat || "",
          hotel:         formData.hotel || "",
          kelompokPeserta: pengeluaranFilterGroup === "panitia-inda" ? "Panitia & Inda" : "Petugas PML & PPL",
          kelas:         isPanitiaInda ? "-" : (formData.kelas || "-"),
        },
        peserta: pengeluaranPeserta,
      });
      return;
    }

    if (docType.id === "spj") {
      if (!spjFilterGroup) {
        alert("Pilih kelompok peserta SPJ (PML & PPL atau Panitia & Inda) terlebih dahulu.");
        return;
      }
      if (!spjFiltered || spjPeserta.length === 0) {
        alert("Tampilkan peserta terlebih dahulu.");
        return;
      }
      const isPanitiaInda = spjFilterGroup === "panitia-inda";
      onPreview({
        formValues: {
          ...formData,
          tanggal_pelunasan: formData.tanggal_pelunasan || "",
          tempat:            formData.tempat || formData.hotel || "",
          kelompokPeserta:   spjFilterGroup === "panitia-inda" ? "Panitia & Inda" : "Petugas PML & PPL",
          kelas:             isPanitiaInda ? "-" : (formData.kelas || ""),
          ttd_kiri:          formData.ttd_kiri || "",
          ttd_kanan:         formData.ttd_kanan || "",
        },
        peserta: spjPeserta,
        selectedFilterGroup: spjFilterGroup,
      });
      return;
    }

    if (docType.id === "spd") {
      if (!spdFiltered || spdPeserta.length === 0) {
        alert("Tampilkan peserta terlebih dahulu.");
        return;
      }

      onPreview({
        formValues: {
          nomor_dokumen:           formData.nomor_dokumen || formData.nomor || "",
          nomor:                   formData.nomor_dokumen || formData.nomor || "",
          tanggal_surat:           formData.tanggal_surat || "",
          tanggal:                 formData.tanggal_surat || "",
          tanggal_awal_kegiatan:   formData.tanggal_awal_kegiatan || "",
          tanggal_akhir_kegiatan:  formData.tanggal_akhir_kegiatan || "",
          tempat:                  formData.tempat || formData.hotel || "",
          hotel:                   formData.hotel || "",
          gelombang:               formData.gelombang || "",
          kelas:                   "-",
          lokasi:                  formData.tempat || formData.hotel || "",
          namaKabps:               formData.namaKabps || "",
          nipKabps:                formData.nipKabps || "",
        },
        peserta: spdPeserta,
      });
      return;
    }

    if (docType.id === "surat-tugas") {
      if (!suratTugasFiltered || suratTugasPeserta.length === 0) { alert("Tampilkan peserta terlebih dahulu."); return; }
      onPreview({
        formValues: {
          nomor_surat:            formData.nomor_surat || "",
          tanggal_surat:          formData.tanggal_surat || "",
          tanggal_awal_kegiatan:  formData.tanggal_awal_kegiatan || "",
          tanggal_akhir_kegiatan: formData.tanggal_akhir_kegiatan || "",
          tempat:                 formData.tempat || formData.hotel || "",
          hotel:                  formData.hotel || "",
          gelombang:              formData.gelombang || "",
          kelas:                  "-",
        },
        peserta: suratTugasPeserta,
      });
      return;
    }

    if (docType.id === "lampiran") {
      // Lampiran tidak memakai input Tempat/Gelombang/Kelas.
      // Data langsung diambil dari sheet bernama "Lampiran" dan tombol PML/PPL akan generate template masing-masing.
      onPreview({
        formValues: {},
        lampiranRows: lampiranData || [],
      });
      return;
    }

    if (docType.id === "bapp") {
      e.preventDefault();
      return;
    }

    if (docType.id === "bast") {
      e.preventDefault();
      return;
    }

    onPreview({ ...formData });
  };

  const renderFields = () => {
    switch (docType.id) {

      // ── DAFTAR HADIR ────────────────────────────────────────────────────────
      case "daftar-hadir":
        return (
          <>
            {!xlsxLoaded && (
              <div className="flex items-center gap-3 rounded-2xl border border-amber-200 bg-amber-50 px-4 py-3">
                <AlertCircle size={18} className="shrink-0 text-amber-500" />
                <p className="text-sm font-semibold text-amber-700">Data petugas belum terdeteksi. Unggah file dari dashboard.</p>
              </div>
            )}

            <div>
              <label className={labelCls}>Tanggal Kegiatan</label>
              <input type="date" className={inputCls} value={formData.tanggal || ""} onChange={(e) => update("tanggal", e.target.value)} />
            </div>

            <div className="grid gap-4 sm:grid-cols-2">
              <div>
                <label className={labelCls}>Jam Mulai</label>
                <input className={inputCls} placeholder="07.30" value={formData.jamMulai || "07.30"} onChange={(e) => update("jamMulai", normalizeJamIndonesia(e.target.value))} />
              </div>
              <div>
                <label className={labelCls}>Jam Selesai</label>
                <input className={inputCls} placeholder="18.00" value={formData.jamSelesai || "18.00"} onChange={(e) => update("jamSelesai", normalizeJamIndonesia(e.target.value))} />
              </div>
            </div>

            <div>
              <p className={labelCls}>Kelompok Peserta</p>
              <div className="flex gap-3">
                {Object.entries(DAFTAR_HADIR_GROUPS).map(([key, grp]) => (
                  <button key={key} type="button"
                    onClick={() => {
                      setDaftarHadirFilterGroup(key);
                      setDaftarHadirFiltered(false);
                      setDaftarHadirPeserta([]);
                      setDaftarHadirNamaInda("");
                      if (key === "panitia-inda") {
                        setFormData((prev) => ({ ...prev, kelas: "" }));
                      }
                    }}
                    className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${daftarHadirFilterGroup === key ? "bg-orange-600 text-white shadow-lg shadow-orange-500/20" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                    {key === "pml-ppl" ? <Users size={15} /> : <Briefcase size={15} />}
                    {grp.label}
                  </button>
                ))}
              </div>
              {!daftarHadirFilterGroup && (
                <p className="mt-2 text-xs font-semibold text-slate-400">Pilih kelompok untuk menampilkan filter peserta.</p>
              )}
            </div>

            <AnimatePresence mode="wait">
              {daftarHadirFilterGroup && (
                <motion.div key={daftarHadirFilterGroup} initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -8 }} transition={{ duration: 0.2 }}>
                  <FilterPesertaPanel
                    xlsxLoaded={xlsxLoaded}
                    formData={formData}
                    setFormData={setFormData}
                    petugasData={petugasData}
                    mode="grouped"
                    selectedGroup={daftarHadirFilterGroup}
                    prependRow={(() => {
                      const hotelLower = cleanText(formData.hotel || formData.tempat || "").toLowerCase();
                      const gelombangVal = cleanText(formData.gelombang || "");
                      if (daftarHadirFilterGroup === "panitia-inda" && gelombangVal === "4") {
                        if (hotelLower.includes("bwp")) return { nama: "Budi Utami", jabatan: "Penanggung Jawab", wilTugas: "BPS Kota Jakarta Timur" };
                        if (hotelLower.includes("stis")) return { nama: "Widiastuti", jabatan: "Kepala BPS Kota Jakarta Timur", wilTugas: "BPS Kota Jakarta Timur" };
                      }
                      return null;
                    })()}
                    onFilterResult={(peserta, namaInda) => {
                      setDaftarHadirPeserta(peserta);
                      setDaftarHadirNamaInda(namaInda);
                      setDaftarHadirFiltered(peserta.length > 0);
                    }}
                  />
                </motion.div>
              )}
            </AnimatePresence>

            {daftarHadirFiltered && daftarHadirFilterGroup && (
              <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}
                className="flex items-center justify-between rounded-2xl border border-orange-100 bg-white/80 px-4 py-3">
                <div>
                  <p className="text-xs font-black uppercase tracking-wider text-slate-500">Tanda Tangan</p>
                  <p className="mt-0.5 font-bold text-slate-800">
                    {daftarHadirFilterGroup === "panitia-inda" ? "Kepala Sub Bagian Umum" : "Instruktur Daerah"}
                  </p>
                </div>
                <div className="flex items-center gap-3">
                  <div className="text-right">
                    <p className="text-xs font-semibold text-slate-400">Kelas di dokumen</p>
                    <p className="font-black text-orange-600">
                      {daftarHadirFilterGroup === "panitia-inda" ? "-" : (formData.kelas || "—")}
                    </p>
                  </div>
                  <Check size={18} className="text-green-500" />
                </div>
              </motion.div>
            )}
          </>
        );

      // ── TANDA TERIMA ────────────────────────────────────────────────────────
      case "tanda-terima": {
        // ✅ FIX: wilTugasOptions hanya dari petugas non-STIS
        const wilTugasOptions = React.useMemo(
          () =>
            uniqueSorted(
              (petugasData || [])
                .filter((p) => cleanText(p.hotel).toUpperCase() !== "STIS")
                .map((p) => cleanText(p.wilTugas))
                .filter(Boolean)
            ),
          [petugasData]
        );

        return (
          <>
            <div>
              <p className={labelCls}>Jenis Tanda Terima</p>
              <div className="flex flex-wrap gap-3">
                <button type="button"
                  onClick={() => { setTandaTerimaType("pelatihan"); setTandaTerimaFiltered(false); setTandaTerimaPeserta([]); setTandaTerimaWilTugas(""); }}
                  className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${tandaTerimaType === "pelatihan" ? "bg-orange-600 text-white shadow-lg shadow-orange-500/20" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                  <Briefcase size={15} /> Tanda Terima Perlengkapan Pelatihan
                </button>
                <button type="button"
                  onClick={() => { setTandaTerimaType("lapangan"); setTandaTerimaFiltered(false); setTandaTerimaPeserta([]); setTandaTerimaWilTugas(""); }}
                  className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${tandaTerimaType === "lapangan" ? "bg-orange-600 text-white shadow-lg shadow-orange-500/20" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                  <Briefcase size={15} /> Tanda Terima Perlengkapan Lapangan
                </button>
                <button type="button"
                  onClick={() => { setTandaTerimaType("mitra-umum"); setTandaTerimaFiltered(false); setTandaTerimaPeserta([]); setTandaTerimaWilTugas(""); setTandaTerimaLokasi("BPS Kota Jakarta Timur"); }}
                  className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${tandaTerimaType === "mitra-umum" ? "bg-orange-600 text-white shadow-lg shadow-orange-500/20" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                  <Briefcase size={15} /> Tanda Terima Perlengkapan Lapangan - Mitra Umum
                </button>
              </div>
              {!tandaTerimaType && (
                <p className="mt-2 text-xs font-semibold text-slate-400">Pilih jenis tanda terima terlebih dahulu.</p>
              )}
            </div>
            {tandaTerimaType && (tandaTerimaType === "mitra-umum" ? (
              <>
                <div>
                  <label className={labelCls}>Tanggal Kegiatan</label>
                  <input type="date" className={inputCls} value={formData.tanggal || ""} onChange={(e) => update("tanggal", e.target.value)} />
                </div>
                <div>
                  <label className={labelCls}>Pilih Wilayah Tugas</label>
                  {/* ✅ FIX: dropdown wilTugas sudah exclude TC = STIS, filter data juga exclude STIS */}
                  <select className={inputCls} value={tandaTerimaWilTugas} onChange={(e) => {
                    const newWilTugas = e.target.value;
                    setTandaTerimaWilTugas(newWilTugas);
                    const filtered = newWilTugas
                      ? (petugasData || []).filter(
                          (p) =>
                            cleanText(p.wilTugas).toUpperCase() === newWilTugas.toUpperCase() &&
                            cleanText(p.hotel).toUpperCase() !== "STIS"
                        )
                      : [];
                    setTandaTerimaPeserta(filtered);
                    setTandaTerimaFiltered(filtered.length > 0);
                  }}>
                    <option value="">-- Pilih Wilayah Tugas --</option>
                    {wilTugasOptions.map((wil) => (
                      <option key={wil} value={wil}>{wil}</option>
                    ))}
                  </select>
                  {tandaTerimaWilTugas && (
                    <p className="mt-1 text-xs font-semibold text-slate-400">
                      {tandaTerimaPeserta.length} petugas non-STIS ditemukan untuk wilayah ini.
                    </p>
                  )}
                </div>
                <div>
                  <label className={labelCls}>Lokasi</label>
                  <input type="text" className={inputCls} value={tandaTerimaLokasi} onChange={(e) => setTandaTerimaLokasi(e.target.value)} />
                </div>
              </>
            ) : (
              <>
                <div>
                  <label className={labelCls}>Tanggal Kegiatan</label>
                  <input type="date" className={inputCls} value={formData.tanggal || ""} onChange={(e) => update("tanggal", e.target.value)} />
                </div>
                <FilterPesertaPanel xlsxLoaded={xlsxLoaded} formData={formData} setFormData={setFormData} petugasData={petugasData} mode="all" selectedGroup="" onFilterResult={(peserta) => { setTandaTerimaPeserta(peserta); setTandaTerimaFiltered(peserta.length > 0); }} />
              </>
            ))}
          </>
        );
      }

      // ── SURAT PERNYATAAN KENDARAAN / SUPER KENDIS ──────────────────────────
      case "surat-pernyataan-kendaraan":
        return (
          <>
            <div>
              <label className={labelCls}>Tanggal Surat</label>
              <input type="date" className={inputCls} value={formData.tanggal_surat || ""} onChange={(e) => update("tanggal_surat", e.target.value)} />
            </div>

            <div>
              <p className={labelCls}>Kelompok Peserta</p>
              <div className="flex gap-3">
                {Object.entries(DAFTAR_HADIR_GROUPS).map(([key, grp]) => (
                  <button key={key} type="button"
                    onClick={() => {
                      setSuratPernyataanFilterGroup(key);
                      setSuratPernyataanFiltered(false);
                      setSuratPernyataanPeserta([]);
                      if (key === "panitia-inda") setFormData((prev) => ({ ...prev, kelas: "" }));
                    }}
                    className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${suratPernyataanFilterGroup === key ? "bg-orange-600 text-white shadow-lg shadow-orange-500/20" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                    {key === "pml-ppl" ? <Users size={15} /> : <Briefcase size={15} />}
                    {key === "pml-ppl" ? "Petugas PML & PPL" : grp.label}
                  </button>
                ))}
              </div>
              {!suratPernyataanFilterGroup && (
                <p className="mt-2 text-xs font-semibold text-slate-400">Pilih kelompok peserta agar Super Kendis dapat dibagi menjadi PML/PPL atau Panitia/Inda.</p>
              )}
            </div>

            <AnimatePresence mode="wait">
              {suratPernyataanFilterGroup && (
                <motion.div key={suratPernyataanFilterGroup} initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -8 }} transition={{ duration: 0.2 }}>
                  <FilterPesertaPanel
                    xlsxLoaded={xlsxLoaded}
                    formData={formData}
                    setFormData={setFormData}
                    petugasData={petugasData}
                    mode="grouped"
                    selectedGroup={suratPernyataanFilterGroup}
                    onFilterResult={(peserta) => {
                      setSuratPernyataanPeserta(peserta);
                      setSuratPernyataanFiltered(peserta.length > 0);
                    }}
                  />
                </motion.div>
              )}
            </AnimatePresence>

            {suratPernyataanFiltered && suratPernyataanFilterGroup && (
              <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}
                className="flex items-center justify-between rounded-2xl border border-orange-100 bg-white/80 px-4 py-3">
                <div>
                  <p className="text-xs font-black uppercase tracking-wider text-slate-500">Kelompok Super Kendis</p>
                  <p className="mt-0.5 font-bold text-slate-800">
                    {suratPernyataanFilterGroup === "panitia-inda" ? "Panitia & Inda" : "Petugas PML & PPL"}
                  </p>
                </div>
                <div className="flex items-center gap-3">
                  <div className="text-right">
                    <p className="text-xs font-semibold text-slate-400">Kelas di dokumen</p>
                    <p className="font-black text-orange-600">
                      {suratPernyataanFilterGroup === "panitia-inda" ? "-" : (formData.kelas || "—")}
                    </p>
                  </div>
                  <Check size={18} className="text-green-500" />
                </div>
              </motion.div>
            )}
          </>
        );

      // ── PENGELUARAN RIIL ────────────────────────────────────────────────────
      case "pengeluaran-riil":
        return (
          <>
            <div>
              <label className={labelCls}>Tanggal Surat</label>
              <input type="date" className={inputCls} value={formData.tanggal_surat || ""} onChange={(e) => update("tanggal_surat", e.target.value)} />
            </div>
            <div>
              <p className={labelCls}>Kelompok Peserta</p>
              <div className="flex gap-3">
                {Object.entries(DAFTAR_HADIR_GROUPS).map(([key, grp]) => (
                  <button key={key} type="button"
                    onClick={() => {
                      setPengeluaranFilterGroup(key);
                      setPengeluaranFiltered(false);
                      setPengeluaranPeserta([]);
                      if (key === "panitia-inda") setFormData((prev) => ({ ...prev, kelas: "" }));
                    }}
                    className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${pengeluaranFilterGroup === key ? "bg-orange-600 text-white shadow-lg shadow-orange-500/20" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                    {key === "pml-ppl" ? <Users size={15} /> : <Briefcase size={15} />}
                    {key === "pml-ppl" ? "Petugas PML & PPL" : grp.label}
                  </button>
                ))}
              </div>
              {!pengeluaranFilterGroup && (
                <p className="mt-2 text-xs font-semibold text-slate-400">Pilih kelompok peserta agar DPR dapat dipisah menjadi PML/PPL atau Panitia/Inda.</p>
              )}
            </div>

            <AnimatePresence mode="wait">
              {pengeluaranFilterGroup && (
                <motion.div key={pengeluaranFilterGroup} initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -8 }} transition={{ duration: 0.2 }}>
                  <FilterPesertaPanel
                    xlsxLoaded={xlsxLoaded}
                    formData={formData}
                    setFormData={setFormData}
                    petugasData={petugasData}
                    mode="grouped"
                    selectedGroup={pengeluaranFilterGroup}
                    onFilterResult={(peserta) => {
                      setPengeluaranPeserta(peserta);
                      setPengeluaranFiltered(peserta.length > 0);
                    }}
                  />
                </motion.div>
              )}
            </AnimatePresence>

            {pengeluaranFiltered && pengeluaranFilterGroup && (
              <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}
                className="flex items-center justify-between rounded-2xl border border-orange-100 bg-white/80 px-4 py-3">
                <div>
                  <p className="text-xs font-black uppercase tracking-wider text-slate-500">Kelompok DPR</p>
                  <p className="mt-0.5 font-bold text-slate-800">
                    {pengeluaranFilterGroup === "panitia-inda" ? "Panitia & Inda" : "Petugas PML & PPL"}
                  </p>
                </div>
                <div className="flex items-center gap-3">
                  <div className="text-right">
                    <p className="text-xs font-semibold text-slate-400">Kelas di dokumen</p>
                    <p className="font-black text-orange-600">
                      {pengeluaranFilterGroup === "panitia-inda" ? "-" : (formData.kelas || "—")}
                    </p>
                  </div>
                  <Check size={18} className="text-green-500" />
                </div>
              </motion.div>
            )}
          </>
        );

      // ── SPJ ─────────────────────────────────────────────────────────────────
      case "spj":
        return (
          <>
            <div>
              <label className={labelCls}>Tanggal Pelunasan</label>
              <input type="date" className={inputCls} value={formData.tanggal_pelunasan || ""} onChange={(e) => update("tanggal_pelunasan", e.target.value)} />
            </div>

            <div className="grid gap-5 sm:grid-cols-2">
              <div>
                <label className={labelCls}>Tanggal Awal</label>
                <input type="date" className={inputCls} value={formData.tanggal_awal_kegiatan || ""} onChange={(e) => update("tanggal_awal_kegiatan", e.target.value)} />
              </div>
              <div>
                <label className={labelCls}>Tanggal Akhir</label>
                <input type="date" className={inputCls} value={formData.tanggal_akhir_kegiatan || ""} onChange={(e) => update("tanggal_akhir_kegiatan", e.target.value)} />
              </div>
            </div>

            <div>
              <p className={labelCls}>Kelompok Peserta SPJ</p>
              <div className="flex gap-3">
                {Object.entries(DAFTAR_HADIR_GROUPS).map(([key, grp]) => (
                  <button key={key} type="button"
                    onClick={() => {
                      setSpjFilterGroup(key);
                      setSpjFiltered(false);
                      setSpjPeserta([]);
                      if (key === "panitia-inda") {
                        setFormData((prev) => ({ ...prev, kelas: "" }));
                      }
                    }}
                    className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${spjFilterGroup === key ? "bg-orange-600 text-white shadow-lg shadow-orange-500/20" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                    {key === "pml-ppl" ? <Users size={15} /> : <Briefcase size={15} />}
                    {key === "pml-ppl" ? "Petugas PML & PPL" : grp.label}
                  </button>
                ))}
              </div>
              {!spjFilterGroup && (
                <p className="mt-2 text-xs font-semibold text-slate-400">Pilih kelompok agar SPJ terpisah antara Panitia & Inda dan Petugas PML & PPL.</p>
              )}
            </div>

            <AnimatePresence mode="wait">
              {spjFilterGroup && (
                <motion.div key={spjFilterGroup} initial={{ opacity: 0, y: 8 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -8 }} transition={{ duration: 0.2 }}>
                  <FilterPesertaPanel
                    xlsxLoaded={xlsxLoaded}
                    formData={formData}
                    setFormData={setFormData}
                    petugasData={petugasData}
                    mode="grouped"
                    selectedGroup={spjFilterGroup}
                    onFilterResult={(peserta) => {
                      setSpjPeserta(peserta);
                      setSpjFiltered(peserta.length > 0);
                    }}
                  />
                </motion.div>
              )}
            </AnimatePresence>

            {spjFiltered && spjFilterGroup && (
              <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}
                className="flex items-center justify-between rounded-2xl border border-orange-100 bg-white/80 px-4 py-3">
                <div>
                  <p className="text-xs font-black uppercase tracking-wider text-slate-500">Kelompok SPJ</p>
                  <p className="mt-0.5 font-bold text-slate-800">
                    {spjFilterGroup === "panitia-inda" ? "Panitia & Inda" : "Petugas PML & PPL"}
                  </p>
                </div>
                <div className="flex items-center gap-3">
                  <div className="text-right">
                    <p className="text-xs font-semibold text-slate-400">Kelas di dokumen</p>
                    <p className="font-black text-orange-600">
                      {spjFilterGroup === "panitia-inda" ? "-" : (formData.kelas || "—")}
                    </p>
                  </div>
                  <Check size={18} className="text-green-500" />
                </div>
              </motion.div>
            )}
          </>
        );

      // ── SPD ─────────────────────────────────────────────────────────────────
      case "spd":
        return (
          <>
            <div className="grid gap-5 sm:grid-cols-2">
              <div>
                <label className={labelCls}>Nomor Dokumen</label>
                <input
                  className={inputCls}
                  placeholder="Contoh: 100/BPS-3171/2026"
                  value={formData.nomor_dokumen || formData.nomor || ""}
                  onChange={(e) => {
                    update("nomor_dokumen", e.target.value);
                    update("nomor", e.target.value);
                  }}
                />
              </div>
              <div>
                <label className={labelCls}>Tanggal Surat</label>
                <input
                  type="date"
                  className={inputCls}
                  value={formData.tanggal_surat || ""}
                  onChange={(e) => update("tanggal_surat", e.target.value)}
                />
              </div>
            </div>

            <div className="grid gap-5 sm:grid-cols-2">
              <div>
                <label className={labelCls}>Tanggal Awal Kegiatan</label>
                <input
                  type="date"
                  className={inputCls}
                  value={formData.tanggal_awal_kegiatan || ""}
                  onChange={(e) => update("tanggal_awal_kegiatan", e.target.value)}
                />
              </div>
              <div>
                <label className={labelCls}>Tanggal Akhir Kegiatan</label>
                <input
                  type="date"
                  className={inputCls}
                  value={formData.tanggal_akhir_kegiatan || ""}
                  onChange={(e) => update("tanggal_akhir_kegiatan", e.target.value)}
                />
              </div>
            </div>

            <FilterPesertaHotelGelombangPanel
              xlsxLoaded={xlsxLoaded}
              formData={formData}
              setFormData={setFormData}
              petugasData={petugasData}
              onFilterResult={(peserta) => {
                setSpdPeserta(peserta);
                setSpdFiltered(peserta.length > 0);
              }}
            />
          </>
        );

      // ── SURAT TUGAS ─────────────────────────────────────────────────────────
      case "surat-tugas":
        return (
          <>
            <div className="grid gap-5 sm:grid-cols-2">
              <div><label className={labelCls}>Nomor Surat</label><input className={inputCls} placeholder="Contoh: 100/BPS-3171/2026" value={formData.nomor_surat || ""} onChange={(e) => update("nomor_surat", e.target.value)} /></div>
              <div><label className={labelCls}>Tanggal Surat</label><input type="date" className={inputCls} value={formData.tanggal_surat || ""} onChange={(e) => update("tanggal_surat", e.target.value)} /></div>
            </div>
            <div className="grid gap-5 sm:grid-cols-2">
              <div><label className={labelCls}>Tanggal Awal Penyelenggaraan</label><input type="date" className={inputCls} value={formData.tanggal_awal_kegiatan || ""} onChange={(e) => update("tanggal_awal_kegiatan", e.target.value)} /></div>
              <div><label className={labelCls}>Tanggal Akhir Penyelenggaraan</label><input type="date" className={inputCls} value={formData.tanggal_akhir_kegiatan || ""} onChange={(e) => update("tanggal_akhir_kegiatan", e.target.value)} /></div>
            </div>
            <FilterPesertaHotelGelombangPanel
              xlsxLoaded={xlsxLoaded}
              formData={formData}
              setFormData={setFormData}
              petugasData={petugasData}
              onFilterResult={(peserta) => {
                setSuratTugasPeserta(peserta);
                setSuratTugasFiltered(peserta.length > 0);
              }}
            />
          </>
        );

      // ── BAPP ────────────────────────────────────────────────────────────────
      case "bapp": {
        const filteredBappRows = React.useMemo(() => {
          const rows = Array.isArray(bappData) ? bappData : [];
          return rows.filter((row) => isBappRowForRole(row, bappRole));
        }, [bappData, bappRole]);

        const bappOptions = React.useMemo(() => {
          const seen = new Set();
          return filteredBappRows
            .map((row) => {
              const identity = getBappIdentityKey(row, bappRole);
              if (seen.has(identity)) return null;
              seen.add(identity);
              return {
                value: identity,
                label: `${cleanText(row.nama) || "Tanpa Nama"} — ${cleanText(row.jabatan_raw || row.jabatan || bappRole)}`,
                row,
              };
            })
            .filter(Boolean)
            .sort((a, b) => a.label.localeCompare(b.label, "id-ID", { sensitivity: "base" }));
        }, [filteredBappRows, bappRole]);

        return (
          <div className="space-y-5">
            <div className="rounded-3xl border border-orange-100 bg-orange-50/70 p-5">
              <p className="text-xs font-black uppercase tracking-[0.2em] text-orange-700">BAPP PML/PPL</p>
              <p className="mt-2 text-sm font-semibold leading-6 text-slate-600">
                Pilih role, tentukan tanggal surat, lalu unduh dokumen manual atau semua.
              </p>
              <p className="mt-3 text-xs font-bold text-slate-500">
                Data terbaca: {bappData.length} baris dari sheet Pembayaran
              </p>
            </div>

            <div>
              <p className={labelCls}>Pilih Role</p>
              <div className="flex gap-3">
                <button type="button" onClick={() => { setBappRole("PML"); setBappManualSelect(""); }} className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${bappRole === "PML" ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>PML</button>
                <button type="button" onClick={() => { setBappRole("PPL"); setBappManualSelect(""); }} className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${bappRole === "PPL" ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>PPL</button>
              </div>
            </div>

            {bappRole && (
              <>
                <div>
                  <label className={labelCls}>Tanggal Surat</label>
                  <input
                    type="date"
                    min="2026-07-15"
                    max="2026-07-31"
                    className={inputCls}
                    value={formData.tanggal_surat || ""}
                    onChange={(e) => update("tanggal_surat", e.target.value)}
                  />
                  <p className="mt-1 text-xs font-semibold text-slate-400">Rentang tanggal yang diizinkan: 15 Juli 2026 sampai 31 Juli 2026.</p>
                </div>

                <div>
                  <label className={labelCls}>Unduh Manual</label>
                  <select value={bappManualSelect} onChange={(e) => setBappManualSelect(e.target.value)} className={inputCls}>
                    <option value="">— Pilih Nama {bappRole} —</option>
                    {bappOptions.map((option) => (
                      <option key={option.value} value={option.value}>{option.label}</option>
                    ))}
                  </select>
                  <p className="mt-1 text-xs font-semibold text-slate-400">Daftar nama diurutkan berdasarkan abjad.</p>
                </div>

                {bappGenerating && (
                  <div className="flex items-center gap-3 rounded-2xl border border-orange-200 bg-orange-50 px-4 py-3 text-sm font-semibold text-orange-700">
                    <LoaderCircle size={18} className="animate-spin" />
                    <span>{bappProgressText || "Sedang menyiapkan file BAPP..."}</span>
                  </div>
                )}

                <div className="grid gap-4 sm:grid-cols-2">
                  <button type="button" onClick={async () => {
                    try {
                      if (!formData.tanggal_surat) throw new Error("Isi tanggal surat terlebih dahulu.");
                      const selectedDate = new Date(formData.tanggal_surat);
                      const minDate = new Date("2026-07-15T00:00:00");
                      const maxDate = new Date("2026-07-31T23:59:59");
                      if (selectedDate < minDate || selectedDate > maxDate) throw new Error("Tanggal surat hanya boleh 15 Juli 2026 sampai 31 Juli 2026.");
                      if (!bappManualSelect) throw new Error("Pilih nama terlebih dahulu.");
                      const chosenRow = bappOptions.find((option) => option.value === bappManualSelect)?.row;
                      if (!chosenRow) throw new Error("Data nama yang dipilih tidak ditemukan.");
                      setBappGenerating(true);
                      setBappProgressText("Membuat dokumen terpilih...");
                      await generateSingleBapp(bappRole === "PML" ? BAPP_PML_TEMPLATE_URL : BAPP_PPL_TEMPLATE_URL, formData, chosenRow, bappRole);
                    } catch (err) { alert(err.message || err); }
                    finally { setBappGenerating(false); setBappProgressText(""); }
                  }} disabled={!bappManualSelect || filteredBappRows.length === 0 || bappGenerating} className="inline-flex items-center justify-center rounded-2xl bg-orange-500 px-5 py-3 text-sm font-black text-white shadow transition hover:bg-orange-600 disabled:cursor-not-allowed disabled:bg-orange-200">
                    Download Terpilih
                  </button>

                  <button type="button" onClick={async () => {
                    try {
                      if (!formData.tanggal_surat) throw new Error("Isi tanggal surat terlebih dahulu.");
                      const selectedDate = new Date(formData.tanggal_surat);
                      const minDate = new Date("2026-07-15T00:00:00");
                      const maxDate = new Date("2026-07-31T23:59:59");
                      if (selectedDate < minDate || selectedDate > maxDate) throw new Error("Tanggal surat hanya boleh 15 Juli 2026 sampai 31 Juli 2026.");
                      setBappGenerating(true);
                      setBappProgressText("Mempersiapkan batch download...");
                      await generateBapp(bappRole === "PML" ? BAPP_PML_TEMPLATE_URL : BAPP_PPL_TEMPLATE_URL, formData, filteredBappRows, bappRole, ({ batchIndex, totalBatches }) => {
                        setBappProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`);
                      });
                    } catch (err) { alert(err.message || err); }
                    finally { setBappGenerating(false); setBappProgressText(""); }
                  }} disabled={filteredBappRows.length === 0 || bappGenerating} className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                    Download Semua
                  </button>
                </div>
              </>
            )}
          </div>
        );
      }

      // ── BAST ────────────────────────────────────────────────────────────────
      case "bast": {
        const rows = lampiranData || [];
        const isPml = bastRole === "PML";

        const nikLookup = React.useMemo(() => {
          const map = new Map();
          for (const p of petugasData || []) {
            const nama = upperText(p.nama);
            if (nama && !map.has(nama)) map.set(nama, cleanText(p.nik));
          }
          return map;
        }, [petugasData]);

        const bastOptions = React.useMemo(() => {
          if (!bastRole) return [];
          const map = new Map();
          for (const r of rows) {
            const nama = cleanText(isPml ? r.nama_pml : r.nama_ppl);
            if (!nama) continue;
            const email = cleanText(isPml ? r.email_pengawas : r.email_pencacah) || "";
            const identity = email ? upperText(email) : `NAMA::${upperText(nama)}`;
            if (!map.has(identity)) map.set(identity, { value: identity, label: nama, name: nama, rows: [] });
            map.get(identity).rows.push(r);
          }
          return [...map.values()].sort((a, b) => a.name.localeCompare(b.name, "id-ID", { sensitivity: "base" }));
        }, [rows, bastRole, isPml]);

        const filteredBastRows = React.useMemo(() => {
          if (!bastRole) return [];
          return rows.filter((r) => cleanText(isPml ? r.nama_pml : r.nama_ppl));
        }, [rows, bastRole, isPml]);

        return (
          <div className="space-y-5">
            <div className="rounded-3xl border border-orange-100 bg-orange-50/70 p-5">
              <p className="text-xs font-black uppercase tracking-[0.2em] text-orange-700">BAST PML/PPL</p>
              <p className="mt-2 text-sm font-semibold leading-6 text-slate-600">
                Pilih role, tentukan tanggal surat, lalu unduh dokumen manual atau semua.
              </p>
              <p className="mt-3 text-xs font-bold text-slate-500">
                Data terbaca: {rows.length} baris dari sheet Lampiran
              </p>
            </div>

            <div>
              <p className={labelCls}>Pilih Role</p>
              <div className="flex gap-3">
                <button type="button" onClick={() => { setBastRole("PML"); setBastManualSelect(""); }} className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${bastRole === "PML" ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>PML</button>
                <button type="button" onClick={() => { setBastRole("PPL"); setBastManualSelect(""); }} className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${bastRole === "PPL" ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>PPL</button>
              </div>
            </div>

            {bastRole && (
              <>
                <div>
                  <label className={labelCls}>Tanggal Surat</label>
                  <input
                    type="date"
                    min="2026-07-15"
                    max="2026-07-31"
                    className={inputCls}
                    value={formData.tanggal_surat || ""}
                    onChange={(e) => update("tanggal_surat", e.target.value)}
                  />
                  <p className="mt-1 text-xs font-semibold text-slate-400">Rentang tanggal yang diizinkan: 15 Juli 2026 sampai 31 Juli 2026.</p>
                </div>

                <div>
                  <label className={labelCls}>Unduh Manual</label>
                  <select value={bastManualSelect} onChange={(e) => setBastManualSelect(e.target.value)} className={inputCls}>
                    <option value="">— Pilih Nama {bastRole} —</option>
                    {bastOptions.map((option) => (
                      <option key={option.value} value={option.value}>{option.label}</option>
                    ))}
                  </select>
                  <p className="mt-1 text-xs font-semibold text-slate-400">Daftar nama diurutkan berdasarkan abjad.</p>
                </div>

                {bastGenerating && (
                  <div className="flex items-center gap-3 rounded-2xl border border-orange-200 bg-orange-50 px-4 py-3 text-sm font-semibold text-orange-700">
                    <LoaderCircle size={18} className="animate-spin" />
                    <span>{bastProgressText || "Sedang menyiapkan file BAST..."}</span>
                  </div>
                )}

                <div className="grid gap-4 sm:grid-cols-2">
                  <button type="button" onClick={async () => {
                    try {
                      if (!formData.tanggal_surat) throw new Error("Isi tanggal surat terlebih dahulu.");
                      const selectedDate = new Date(formData.tanggal_surat);
                      const minDate = new Date("2026-07-15T00:00:00");
                      const maxDate = new Date("2026-07-31T23:59:59");
                      if (selectedDate < minDate || selectedDate > maxDate) throw new Error("Tanggal surat hanya boleh 15 Juli 2026 sampai 31 Juli 2026.");
                      if (!bastManualSelect) throw new Error("Pilih nama terlebih dahulu.");
                      const chosen = bastOptions.find((option) => option.value === bastManualSelect);
                      if (!chosen) throw new Error("Data nama yang dipilih tidak ditemukan.");
                      setBastGenerating(true);
                      setBastProgressText("Membuat dokumen terpilih...");
                      await generateSingleBast(bastRole === "PML" ? BAST_PML_TEMPLATE_URL : BAST_PPL_TEMPLATE_URL, formData, chosen.rows, bastRole, nikLookup, chosen.name);
                    } catch (err) { alert(err.message || err); }
                    finally { setBastGenerating(false); setBastProgressText(""); }
                  }} disabled={!bastManualSelect || filteredBastRows.length === 0 || bastGenerating} className="inline-flex items-center justify-center rounded-2xl bg-orange-500 px-5 py-3 text-sm font-black text-white shadow transition hover:bg-orange-600 disabled:cursor-not-allowed disabled:bg-orange-200">
                    Download Terpilih
                  </button>

                  <button type="button" onClick={async () => {
                    try {
                      if (!formData.tanggal_surat) throw new Error("Isi tanggal surat terlebih dahulu.");
                      const selectedDate = new Date(formData.tanggal_surat);
                      const minDate = new Date("2026-07-15T00:00:00");
                      const maxDate = new Date("2026-07-31T23:59:59");
                      if (selectedDate < minDate || selectedDate > maxDate) throw new Error("Tanggal surat hanya boleh 15 Juli 2026 sampai 31 Juli 2026.");
                      setBastGenerating(true);
                      setBastProgressText("Mempersiapkan batch download...");
                      await generateBast(bastRole === "PML" ? BAST_PML_TEMPLATE_URL : BAST_PPL_TEMPLATE_URL, formData, filteredBastRows, bastRole, nikLookup, ({ batchIndex, totalBatches }) => {
                        setBastProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`);
                      });
                    } catch (err) { alert(err.message || err); }
                    finally { setBastGenerating(false); setBastProgressText(""); }
                  }} disabled={filteredBastRows.length === 0 || bastGenerating} className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                    Download Semua
                  </button>
                </div>
              </>
            )}
          </div>
        );
      }

      // ── SURAT PERNYATAAN PENYELESAIAN LAPANGAN ───────────────────────────
      case "surat-pernyataan-penyelesaian-lapangan": {
        const filteredSuratPenyelesaianLapanganRows = React.useMemo(() => {
          const rows = Array.isArray(bappData) ? bappData : [];
          return rows.filter((row) => isBappRowForRole(row, "PML"));
        }, [bappData]);

        const suratPenyelesaianLapanganOptions = React.useMemo(() => {
          const seen = new Set();
          return filteredSuratPenyelesaianLapanganRows
            .map((row) => {
              const identity = getBappIdentityKey(row, "PML");
              if (seen.has(identity)) return null;
              seen.add(identity);
              return {
                value: identity,
                label: `${cleanText(row.nama) || "Tanpa Nama"} — ${cleanText(row.jabatan_raw || row.jabatan || "PML")}`,
                row,
              };
            })
            .filter(Boolean)
            .sort((a, b) => a.label.localeCompare(b.label, "id-ID", { sensitivity: "base" }));
        }, [filteredSuratPenyelesaianLapanganRows]);

        return (
          <div className="space-y-5">
            <div className="rounded-3xl border border-orange-100 bg-orange-50/70 p-5">
              <p className="text-xs font-black uppercase tracking-[0.2em] text-orange-700">Surat Pernyataan Penyelesaian Lapangan</p>
              <p className="mt-2 text-sm font-semibold leading-6 text-slate-600">
                Pilih nama untuk unduh manual, atau unduh semua dalam batch.
              </p>
              <p className="mt-3 text-xs font-bold text-slate-500">
                Data terbaca: {bappData.length} baris dari sheet Pembayaran
              </p>
            </div>

            <div>
              <label className={labelCls}>Unduh Manual</label>
              <select value={suratPenyelesaianLapanganSelect} onChange={(e) => setSuratPenyelesaianLapanganSelect(e.target.value)} className={inputCls}>
                <option value="">— Pilih Nama PML —</option>
                {suratPenyelesaianLapanganOptions.map((option) => (
                  <option key={option.value} value={option.value}>{option.label}</option>
                ))}
              </select>
              <p className="mt-1 text-xs font-semibold text-slate-400">Daftar nama diurutkan berdasarkan abjad dan memakai email sebagai kunci unik.</p>
            </div>

            {suratPenyelesaianLapanganGenerating && (
              <div className="flex items-center gap-3 rounded-2xl border border-orange-200 bg-orange-50 px-4 py-3 text-sm font-semibold text-orange-700">
                <LoaderCircle size={18} className="animate-spin" />
                <span>{suratPenyelesaianLapanganProgressText || "Sedang menyiapkan dokumen..."}</span>
              </div>
            )}

            <div className="grid gap-4 sm:grid-cols-2">
              <button type="button" onClick={async () => {
                try {
                  if (!suratPenyelesaianLapanganSelect) throw new Error("Pilih nama terlebih dahulu.");
                  const chosenRow = suratPenyelesaianLapanganOptions.find((option) => option.value === suratPenyelesaianLapanganSelect)?.row;
                  if (!chosenRow) throw new Error("Data nama yang dipilih tidak ditemukan.");
                  setSuratPenyelesaianLapanganGenerating(true);
                  setSuratPenyelesaianLapanganProgressText("Membuat dokumen terpilih...");
                  await generateSingleSuratPernyataanPenyelesaianLapangan(SURAT_PERNYATAAN_PENYELESAIAN_LAPANGAN_TEMPLATE_URL, chosenRow);
                } catch (err) { alert(err.message || err); }
                finally { setSuratPenyelesaianLapanganGenerating(false); setSuratPenyelesaianLapanganProgressText(""); }
              }} disabled={!suratPenyelesaianLapanganSelect || filteredSuratPenyelesaianLapanganRows.length === 0 || suratPenyelesaianLapanganGenerating} className="inline-flex items-center justify-center rounded-2xl bg-orange-500 px-5 py-3 text-sm font-black text-white shadow transition hover:bg-orange-600 disabled:cursor-not-allowed disabled:bg-orange-200">
                Download Terpilih
              </button>

              <button type="button" onClick={async () => {
                try {
                  setSuratPenyelesaianLapanganGenerating(true);
                  setSuratPenyelesaianLapanganProgressText("Mempersiapkan batch download...");
                  await generateSuratPernyataanPenyelesaianLapangan(SURAT_PERNYATAAN_PENYELESAIAN_LAPANGAN_TEMPLATE_URL, filteredSuratPenyelesaianLapanganRows, ({ batchIndex, totalBatches }) => {
                    setSuratPenyelesaianLapanganProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`);
                  });
                } catch (err) { alert(err.message || err); }
                finally { setSuratPenyelesaianLapanganGenerating(false); setSuratPenyelesaianLapanganProgressText(""); }
              }} disabled={filteredSuratPenyelesaianLapanganRows.length === 0 || suratPenyelesaianLapanganGenerating} className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                Download Semua
              </button>
            </div>
          </div>
        );
      }

      // ── LAMPIRAN ────────────────────────────────────────────────────────────
      case "lampiran": {
        const rows = lampiranData || [];

        const generateLampiranPml = async () => {
          if (!xlsxLoaded) {
            throw new Error("Data XLSX/Google Sheet belum dimuat.");
          }
          if (rows.length === 0) {
            throw new Error("Sheet Lampiran belum terbaca atau kosong. Pastikan Google Sheet dibaca sebagai XLSX dan nama tab adalah Lampiran.");
          }
          await generateLampiran(LAMPIRAN_PML_TEMPLATE_URL, {}, rows, "PML");
        };

        const generateLampiranPpl = async () => {
          if (!xlsxLoaded) {
            throw new Error("Data XLSX/Google Sheet belum dimuat.");
          }
          if (rows.length === 0) {
            throw new Error("Sheet Lampiran belum terbaca atau kosong. Pastikan Google Sheet dibaca sebagai XLSX dan nama tab adalah Lampiran.");
          }
          await generateLampiran(LAMPIRAN_PPL_TEMPLATE_URL, {}, rows, "PPL");
        };

        return (
          <div className="space-y-5">
            <div className="rounded-3xl border border-orange-100 bg-orange-50/70 p-5">
              <p className="text-xs font-black uppercase tracking-[0.2em] text-orange-700">Generate Lampiran</p>
              <p className="mt-2 text-sm font-semibold leading-6 text-slate-600">
                Pilih jenis lampiran
              </p>
              <p className="mt-3 text-xs font-bold text-slate-500">
                Status data: {xlsxLoaded ? `${rows.length} baris Lampiran terbaca` : "data belum dimuat"}
              </p>
            </div>

            {/* Manual generate: combined dropdown PML + PPL */}
            <div className="rounded-2xl border border-orange-100 bg-white p-4">
              <p className="mb-2 text-xs font-black uppercase tracking-[0.2em] text-orange-700">Generate Manual (PML &amp; PPL)</p>
              <div className="flex flex-col sm:flex-row sm:items-center gap-2">
                <select value={lampiranManualSelect} onChange={(e) => setLampiranManualSelect(e.target.value)}
                  className="w-full sm:w-96 rounded-2xl border border-orange-100 bg-white/80 px-3 py-2 text-sm font-semibold text-slate-800 shadow-sm outline-none">
                  <option value="">— Pilih Petugas (PML / PPL) —</option>
                  {(() => {
                    const map = new Map();
                    for (const r of rows) {
                      // PML
                      const namePml = cleanText(r.nama_pml);
                      if (namePml) {
                        const email = cleanText(r.email_pengawas) || "";
                        const id = email ? upperText(email) : `NAMA::${upperText(namePml)}`;
                        const key = `PML::${id}`;
                        if (!map.has(key)) map.set(key, { value: key, label: `${namePml} (PML)` , name: namePml });
                      }
                      // PPL
                      const namePpl = cleanText(r.nama_ppl);
                      if (namePpl) {
                        const email = cleanText(r.email_pencacah) || "";
                        const id = email ? upperText(email) : `NAMA::${upperText(namePpl)}`;
                        const key = `PPL::${id}`;
                        if (!map.has(key)) map.set(key, { value: key, label: `${namePpl} (PPL)`, name: namePpl });
                      }
                    }
                    return Array.from(map.values()).sort((a, b) => a.name.localeCompare(b.name, 'id-ID', { sensitivity: 'base' })).map(o => (
                      <option key={o.value} value={o.value}>{o.label}</option>
                    ));
                  })()}
                </select>

                <button onClick={async () => {
                  try {
                    if (!lampiranManualSelect) { alert('Pilih petugas terlebih dahulu.'); return; }
                    if (!xlsxLoaded) throw new Error('Data belum dimuat');
                    const sepIndex = lampiranManualSelect.indexOf('::');
                    const role = lampiranManualSelect.slice(0, sepIndex);
                    const id = lampiranManualSelect.slice(sepIndex + 2);
                    const isPml = role === 'PML';
                    const filtered = rows.filter((r) => {
                      const name = isPml ? cleanText(r.nama_pml) : cleanText(r.nama_ppl);
                      const email = isPml ? cleanText(r.email_pengawas) || "" : cleanText(r.email_pencacah) || "";
                      const key = email ? upperText(email) : `NAMA::${upperText(name)}`;
                      return key === id;
                    });
                    if (filtered.length === 0) throw new Error('Tidak ada data untuk pilihan ini');
                    const displayName = isPml ? filtered[0].nama_pml : filtered[0].nama_ppl;
                    await generateSingleLampiran(isPml ? LAMPIRAN_PML_TEMPLATE_URL : LAMPIRAN_PPL_TEMPLATE_URL, {}, filtered, isPml ? 'PML' : 'PPL', displayName || 'Tanpa Nama');
                  } catch (err) { alert(err.message || err); }
                }} type="button" className="w-full sm:w-auto inline-flex items-center justify-center gap-2 rounded-2xl bg-orange-500 px-4 py-2 text-sm font-black text-white shadow transition hover:bg-orange-600">Generate Terpilih</button>
              </div>
            </div>

            <div className="grid gap-4 sm:grid-cols-2">
              <div className="rounded-3xl border border-orange-100 bg-white p-5 shadow-sm">
                <div className="mb-4 flex h-12 w-12 items-center justify-center rounded-2xl bg-orange-100 text-orange-600">
                  <Users size={22} />
                </div>
                <h3 className="text-lg font-black text-slate-900">Lampiran PML</h3>
                <p className="mt-1 text-sm font-semibold leading-6 text-slate-500">
                  Format lampiran dengan kolom No, Nama Petugas Lapangan Sensus, Kecamatan/Distrik, Desa/Kampung/Nagari, dan Jumlah SLS/Sub-SLS.
                </p>
                <div className="mt-5">
                  <GenerateDocxButton label="Generate Lampiran PML" onGenerate={generateLampiranPml} />
                </div>
              </div>

              <div className="rounded-3xl border border-orange-100 bg-white p-5 shadow-sm">
                <div className="mb-4 flex h-12 w-12 items-center justify-center rounded-2xl bg-orange-100 text-orange-600">
                  <FileText size={22} />
                </div>
                <h3 className="text-lg font-black text-slate-900">Lampiran PPL</h3>
                <p className="mt-1 text-sm font-semibold leading-6 text-slate-500">
                  Format lampiran dengan kolom No, Kecamatan/Distrik, Desa/Kampung/Nagari, dan Jumlah SLS/Sub-SLS.
                </p>
                <div className="mt-5">
                  <GenerateDocxButton label="Generate Lampiran PPL" onGenerate={generateLampiranPpl} />
                </div>
              </div>
            </div>

            <div className="rounded-2xl border border-orange-100 bg-white/75 p-4 text-sm font-semibold text-slate-600">
              <p className="font-bold text-slate-800">Statistik unik email</p>
              <p className="mt-1">PML unik (email): {new Set(rows.map(r => upperText(cleanText(r.email_pengawas))).filter(Boolean)).size}</p>
              <p className="mt-1">PPL unik (email): {new Set(rows.map(r => upperText(cleanText(r.email_pencacah))).filter(Boolean)).size}</p>
            </div>
          </div>
        );
      }

      default:
        return null;
    }
  };

  return (
    <form onSubmit={handleSubmit} className="space-y-5 rounded-[2.5rem] border border-orange-100 bg-white/80 p-6 shadow-xl shadow-orange-900/5 backdrop-blur md:p-10">
      {renderFields()}
      {docType.id !== "lampiran" && docType.id !== "bapp" && docType.id !== "bast" && docType.id !== "surat-pernyataan-penyelesaian-lapangan" && (
        <div className="border-t border-orange-100 pt-5">
          <button type="submit" className="group inline-flex items-center gap-2 rounded-2xl bg-orange-500 px-7 py-4 font-black text-white shadow-2xl shadow-orange-500/25 transition hover:-translate-y-1 hover:bg-orange-600">
            Pratinjau Dokumen <ChevronRight className="transition group-hover:translate-x-1" size={18} />
          </button>
        </div>
      )}
    </form>
  );
}

// BAST


// ─── DOCX PREVIEW COMPONENTS ─────────────────────────────────────────────────

const docxPreviewStyle = `
  .docx-wrapper { background: transparent !important; padding: 0 !important; }
  .docx-wrapper > section.docx { margin: 0 auto 24px auto !important; box-shadow: 0 20px 45px rgba(15,23,42,0.12) !important; }
  @media print {
    body * { visibility: hidden; }
    .docx-wrapper, .docx-wrapper * { visibility: visible; }
    .docx-wrapper { position: absolute; left: 0; top: 0; width: 100%; }
    .docx-wrapper > section.docx { box-shadow: none !important; margin: 0 !important; }
  }
`;

const RENDER_OPTS = { className: "docx-preview", inWrapper: true, ignoreWidth: false, ignoreHeight: false, ignoreFonts: false, breakPages: true, renderHeaders: true, renderFooters: true, useBase64URL: true };

function useSingleDocxPreview(createBlob, deps) {
  const containerRef = useRef(null);
  const [loading, setLoading] = useState(true);
  const [error, setError]     = useState("");
  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        setLoading(true); setError("");
        const blob = await createBlob();
        if (cancelled || !containerRef.current) return;
        containerRef.current.innerHTML = "";
        await renderAsync(blob, containerRef.current, null, RENDER_OPTS);
      } catch (err) {
        if (!cancelled) setError(err?.message || "Gagal memuat pratinjau DOCX.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, deps);
  return { containerRef, loading, error };
}

function DocxPreviewShell({ loading, error, templateName, children }) {
  return (
    <div className="rounded-[2rem] border border-orange-100 bg-white p-4 shadow-xl shadow-orange-900/5">
      <style>{docxPreviewStyle}</style>
      {loading && (
        <div className="rounded-2xl border border-orange-100 bg-orange-50/60 p-8 text-center">
          <p className="text-sm font-black text-orange-700">Memuat pratinjau dari template DOCX...</p>
          <p className="mt-1 text-xs font-semibold text-slate-500">Template: {templateName}</p>
        </div>
      )}
      {error && <div className="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm font-bold text-red-600">{error}</div>}
      {children}
    </div>
  );
}

function DaftarHadirDocxPreview({ formValues, peserta, namaInda, selectedFilterGroup }) {
  const { containerRef, loading, error } = useSingleDocxPreview(
    () => createDaftarHadirBlob(DAFTAR_HADIR_TEMPLATE_URL, formValues || {}, peserta || [], namaInda || "", selectedFilterGroup || ""),
    [formValues, peserta, namaInda, selectedFilterGroup]
  );
  return <DocxPreviewShell loading={loading} error={error} templateName="1. Daftar Hadir Pelatihan SE2026.docx"><div ref={containerRef} className="overflow-x-auto" /></DocxPreviewShell>;
}

function TandaTerimaDocxPreview({ formValues, peserta, tandaTerimaType }) {
  const templateUrl = (tandaTerimaType === "lapangan" || tandaTerimaType === "mitra-umum") ? TANDA_TERIMA_LAPANGAN_TEMPLATE_URL : TANDA_TERIMA_TEMPLATE_URL;
  const templateName = (tandaTerimaType === "lapangan" || tandaTerimaType === "mitra-umum") ? "2. Tanda Terima Perlengkapan SE2026 - Copy.docx" : "2. Tanda Terima Perlengkapan SE2026.docx";
  const { containerRef, loading, error } = useSingleDocxPreview(
    () => createTandaTerimaBlob(templateUrl, formValues || {}, peserta || []), [formValues, peserta, tandaTerimaType]
  );
  return <DocxPreviewShell loading={loading} error={error} templateName={templateName}><div ref={containerRef} className="overflow-x-auto" /></DocxPreviewShell>;
}

function SuratPernyataanKendaraanDocxPreview({ formValues, peserta }) {
  const { containerRef, loading, error } = useSingleDocxPreview(
    () => createSuratPernyataanKendaraanBlob(SURAT_PERNYATAAN_KENDARAAN_TEMPLATE_URL, formValues || {}, peserta || []), [formValues, peserta]
  );
  return <DocxPreviewShell loading={loading} error={error} templateName="3. Super Kendis Pelatihan SE2026.docx"><div ref={containerRef} className="overflow-x-auto" /></DocxPreviewShell>;
}

function PengeluaranRiilDocxPreview({ formValues, peserta }) {
  const { containerRef, loading, error } = useSingleDocxPreview(
    () => createPengeluaranRiilBlob(PENGELUARAN_RIIL_TEMPLATE_URL, formValues || {}, peserta || []), [formValues, peserta]
  );
  return <DocxPreviewShell loading={loading} error={error} templateName="4. DPR_Pelatihan SE 2026.docx"><div ref={containerRef} className="overflow-x-auto" /></DocxPreviewShell>;
}

function SpjDocxPreview({ formValues, peserta }) {
  const { containerRef, loading, error } = useSingleDocxPreview(
    () => createSpjBlob(SPJ_TEMPLATE_URL, formValues || {}, peserta || []), [formValues, peserta]
  );
  return <DocxPreviewShell loading={loading} error={error} templateName="5. SPJ Pelatihan_SE26.docx"><div ref={containerRef} className="overflow-x-auto" /></DocxPreviewShell>;
}

function SuratTugasDocxPreview({ formValues, peserta }) {
  const { containerRef, loading, error } = useSingleDocxPreview(
    () => createSuratTugasBlob(SURAT_TUGAS_TEMPLATE_URL, formValues || {}, peserta || []), [formValues, peserta]
  );
  return <DocxPreviewShell loading={loading} error={error} templateName="6. Surat Tugas.docx"><div ref={containerRef} className="overflow-x-auto" /></DocxPreviewShell>;
}

function SpdDocxPreview({ formValues, peserta }) {
  const mainRef     = useRef(null);
  const lampiranRef = useRef(null);
  const [loading, setLoading]             = useState(true);
  const [error, setError]                 = useState("");
  const [activePreview, setActivePreview] = useState("spd");

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        setLoading(true); setError("");
        const [blobMain, blobLampiran] = await Promise.all([
          createSpdBlob(SPD_TEMPLATE_URL, formValues || {}, peserta || []),
          createSpdBlob(SPD_LAMPIRAN_TEMPLATE_URL, formValues || {}, peserta || []),
        ]);
        if (cancelled) return;
        if (mainRef.current)     { mainRef.current.innerHTML     = ""; await renderAsync(blobMain,     mainRef.current,     null, RENDER_OPTS); }
        if (lampiranRef.current) { lampiranRef.current.innerHTML = ""; await renderAsync(blobLampiran, lampiranRef.current, null, RENDER_OPTS); }
      } catch (err) {
        if (!cancelled) setError(err?.message || "Gagal memuat pratinjau DOCX.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [formValues, peserta]);

  return (
    <div className="space-y-4 rounded-[2rem] border border-orange-100 bg-white p-4 shadow-xl shadow-orange-900/5">
      <style>{docxPreviewStyle}</style>
      {loading && <div className="rounded-2xl border border-orange-100 bg-orange-50/60 p-8 text-center"><p className="text-sm font-black text-orange-700">Memuat pratinjau SPD &amp; Lampiran...</p></div>}
      {error && <div className="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm font-bold text-red-600">{error}</div>}
      <div className="flex gap-3">
        {[{ key: "spd", label: "Preview SPD" }, { key: "lampiran", label: "Preview Lampiran SPD" }].map(({ key, label }) => (
          <button key={key} type="button" onClick={() => setActivePreview(key)}
            className={`rounded-2xl px-5 py-2.5 text-sm font-black transition ${activePreview === key ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
            {label}
          </button>
        ))}
      </div>
      <div className={activePreview === "spd" ? "block" : "hidden"}><div ref={mainRef} className="overflow-x-auto rounded-2xl border border-orange-100 bg-white" /></div>
      <div className={activePreview === "lampiran" ? "block" : "hidden"}><div ref={lampiranRef} className="overflow-x-auto rounded-2xl border border-orange-100 bg-white" /></div>
    </div>
  );
}

function LampiranDocxPreview({ formValues, lampiranRows }) {
  const pmlRef = useRef(null);
  const pplRef = useRef(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [activePreview, setActivePreview] = useState("pml");

  React.useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        setLoading(true); setError("");
        const [blobPml, blobPpl] = await Promise.all([
          createLampiranBlob(LAMPIRAN_PML_TEMPLATE_URL, formValues || {}, lampiranRows || [], "PML"),
          createLampiranBlob(LAMPIRAN_PPL_TEMPLATE_URL, formValues || {}, lampiranRows || [], "PPL"),
        ]);
        if (cancelled) return;
        if (pmlRef.current) { pmlRef.current.innerHTML = ""; await renderAsync(blobPml, pmlRef.current, null, RENDER_OPTS); }
        if (pplRef.current) { pplRef.current.innerHTML = ""; await renderAsync(blobPpl, pplRef.current, null, RENDER_OPTS); }
      } catch (err) {
        if (!cancelled) setError(err?.message || "Gagal memuat pratinjau Lampiran.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();
    return () => { cancelled = true; };
  }, [formValues, lampiranRows]);

  return (
    <div className="space-y-4 rounded-[2rem] border border-orange-100 bg-white p-4 shadow-xl shadow-orange-900/5">
      <style>{docxPreviewStyle}</style>
      {loading && <div className="rounded-2xl border border-orange-100 bg-orange-50/60 p-8 text-center"><p className="text-sm font-black text-orange-700">Memuat pratinjau Lampiran PML &amp; PPL...</p></div>}
      {error && <div className="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm font-bold text-red-600">{error}</div>}
      {/* Preview-only: removed manual controls (moved to form) */}
      <div className="flex gap-3">
        {[{ key: "pml", label: "Preview Lampiran PML" }, { key: "ppl", label: "Preview Lampiran PPL" }].map(({ key, label }) => (
          <button key={key} type="button" onClick={() => setActivePreview(key)}
            className={`rounded-2xl px-5 py-2.5 text-sm font-black transition ${activePreview === key ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
            {label}
          </button>
        ))}
      </div>
      <div className={activePreview === "pml" ? "block" : "hidden"}><div ref={pmlRef} className="overflow-x-auto rounded-2xl border border-orange-100 bg-white" /></div>
      <div className={activePreview === "ppl" ? "block" : "hidden"}><div ref={pplRef} className="overflow-x-auto rounded-2xl border border-orange-100 bg-white" /></div>
    </div>
  );
}

// ─── BAST ─────────────────────────────────────────────────
const BAST_NILAI_PERJANJIAN = {
  PML: { nilai: "Rp5.831.000,00", terbilang: "Lima juta delapan ratus tiga puluh satu ribu rupiah" },
  PPL: { nilai: "Rp5.534.000,00", terbilang: "Lima juta lima ratus tiga puluh empat ribu rupiah" },
};

// Nomor surat: pakai nomor kontrak petugas sebagai sumber nomor urut
// (pola yang sama dipakai buildBappTemplateData -> nomor_prefix).
function buildBastNomorSurat(nomorKontrak, jenis) {
  const prefix = extractNomorPrefix(nomorKontrak) || "...";
  const suffix = jenis === "PML"
    ? "BAST-I-SE2026/PML/3172/SS.340/2026"
    : "BAST-I-SE2026/PPL/3172/SS.330/2026";
  return `B-${prefix}/${suffix}`;
}

function buildBastTemplateData(formValues, personRows, jenis, nikLookup) {
  const isPml = upperText(jenis) === "PML";
  const grouped = groupLampiranRows(personRows || [], jenis);

  const namaPetugas = isPml
    ? cleanText(personRows?.[0]?.nama_pml)
    : cleanText(personRows?.[0]?.nama_ppl);

  const nomorKontrak = cleanText(
    isPml ? personRows?.[0]?.nomor_kontrak_pml : personRows?.[0]?.nomor_kontrak_ppl
  );

  const totalJumlah = grouped.reduce((sum, r) => sum + (r.jumlah || 0), 0);
  const dateParts = getBappDateParts(formValues?.tanggal_surat || "");
  const nilai = BAST_NILAI_PERJANJIAN[isPml ? "PML" : "PPL"];
  const nik = nikLookup?.get(upperText(namaPetugas)) || "";

  return {
    nomor_surat: buildBastNomorSurat(nomorKontrak, isPml ? "PML" : "PPL"),
    nomor_perjanjian: nomorKontrak || "...",
    hari: dateParts.hari_terbilang,
    tanggal_terbilang: dateParts.tanggal_terbilang,
    tanggal: dateParts.tanggal,
    bulan: dateParts.bulan,
    bulan_terbilang: dateParts.bulan_terbilang,
    nama: namaPetugas,
    nik,
    jumlah: totalJumlah,
    nilai_perjanjian: nilai.nilai,
    nilai_perjanjian_terbilang: nilai.terbilang,
    peserta: grouped.map((r, idx) => ({
      no: idx + 1,
      nama_petugas: isPml
          ? cleanText(r.nama_ppl)
          : namaPetugas,
      kecamatan: formatKodeNama(r.kdkec, r.kecamatan),
      kelurahan: formatKodeNama(r.kddesa, r.kelurahan),
      jumlah: r.jumlah || 0,
    })),
  };
}

// ── 5) CREATE / DOWNLOAD FUNCTIONS ───────────────────────────
// Mengikuti pola createLampiranBlobFromTemplateBuffer + generateLampiran,
// supaya fetch template hanya sekali lalu dipakai berulang untuk tiap orang.

function createBastBlobFromTemplateBuffer(templateArrayBuffer, formValues, personRows, jenis, nikLookup) {
  const zip = new PizZip(templateArrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildBastTemplateData(formValues || {}, personRows || [], jenis, nikLookup));
  return doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

async function generateSingleBast(templateUrl, formValues, personRows, jenis, nikLookup, displayName) {
  if (!personRows || personRows.length === 0) throw new Error("Tidak ada data untuk BAST yang dipilih");
  const templateResponse = await fetch(templateUrl);
  if (!templateResponse.ok) throw new Error(`Gagal memuat template BAST: ${templateResponse.status} ${templateResponse.statusText}`);
  const templateArrayBuffer = await templateResponse.arrayBuffer();
  const blob = createBastBlobFromTemplateBuffer(templateArrayBuffer, formValues || {}, personRows, jenis, nikLookup);
  const safeName = sanitizeFileName(displayName || `${jenis}-bast`);
  saveAs(blob, `BAST ${jenis} - ${safeName}.docx`);
}

const BAST_ZIP_BATCH_SIZE = 150;

async function generateBast(templateUrl, formValues, lampiranRows, jenis, nikLookup, onProgress) {
  const sourceRows = lampiranRows || [];
  const isPml = upperText(jenis) === "PML";

  // Kelompokkan baris per-orang, sama persis dengan logika di generateLampiran()
  // supaya email dipakai sebagai kunci utama (menghindari salah gabung nama kembar).
  const groups = new Map();
  for (const row of sourceRows) {
    const namaPml = cleanText(row.nama_pml || "");
    const namaPpl = cleanText(row.nama_ppl || "");
    const emailPml = cleanText(row.email_pengawas || "");
    const emailPpl = cleanText(row.email_pencacah || "");
    const displayName = isPml ? namaPml : namaPpl;
    if (!displayName) continue;
    const emailKey = upperText(isPml ? emailPml : emailPpl);
    const identity = emailKey || `NAMA::${upperText(displayName)}`;
    if (!groups.has(identity)) groups.set(identity, { displayName, rows: [] });
    groups.get(identity).rows.push(row);
  }

  if (groups.size === 0) throw new Error(`Tidak ada data ${jenis}`);

  const templateResponse = await fetch(templateUrl);
  if (!templateResponse.ok) throw new Error(`Gagal memuat template BAST: ${templateResponse.status} ${templateResponse.statusText}`);
  const templateArrayBuffer = await templateResponse.arrayBuffer();

  const entries = [...groups.values()];
  const totalBatches = Math.ceil(entries.length / BAST_ZIP_BATCH_SIZE);

  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
    const batchEntries = entries.slice(batchIndex * BAST_ZIP_BATCH_SIZE, (batchIndex + 1) * BAST_ZIP_BATCH_SIZE);
    const zipFiles = [];
    for (const { displayName, rows } of batchEntries) {
      const blob = createBastBlobFromTemplateBuffer(templateArrayBuffer, formValues || {}, rows, jenis, nikLookup);
      zipFiles.push({ name: `BAST ${jenis} - ${sanitizeFileName(displayName)}.docx`, blob });
    }
    if (typeof onProgress === "function") {
      onProgress({ batchIndex: batchIndex + 1, totalBatches, totalRows: entries.length });
    }
    const batchSuffix = totalBatches > 1 ? ` - Bagian ${batchIndex + 1} dari ${totalBatches}` : "";
    await downloadMultipleAsZip(zipFiles, `BAST ${jenis} ${formValues?.tanggal_surat || "SE2026"}${batchSuffix}.zip`);
  }
}

// ─── DOC PREVIEW (dispatcher) ─────────────────────────────────────────────────

function DocPreview({ docType, data }) {
  const renderDoc = () => {
    switch (docType.id) {
      case "daftar-hadir":            return <DaftarHadirDocxPreview formValues={data.formValues || {}} peserta={data.peserta || []} namaInda={data.namaInda || ""} selectedFilterGroup={data.selectedFilterGroup || ""} />;
      case "tanda-terima":            return <TandaTerimaDocxPreview formValues={data.formValues || {}} peserta={data.peserta || []} tandaTerimaType={data.tandaTerimaType || "pelatihan"} />;
      case "surat-pernyataan-kendaraan": return <SuratPernyataanKendaraanDocxPreview formValues={data.formValues || {}} peserta={data.peserta || []} />;
      case "pengeluaran-riil":        return <PengeluaranRiilDocxPreview formValues={data.formValues || {}} peserta={data.peserta || []} />;
      case "spj":                     return <SpjDocxPreview formValues={data.formValues || {}} peserta={data.peserta || []} />;
      case "spd":                     return <SpdDocxPreview formValues={data.formValues || data} peserta={data.peserta || []} />;
      case "surat-tugas":             return <SuratTugasDocxPreview formValues={data.formValues || {}} peserta={data.peserta || []} />;
      case "lampiran":                return <LampiranDocxPreview formValues={data.formValues || {}} lampiranRows={data.lampiranRows || []} />;
      default:                        return <p className="text-sm text-slate-500">Dokumen tidak dikenali.</p>;
    }
  };
  return (
    <div className="overflow-hidden rounded-[2.5rem] border border-orange-100 bg-white/80 shadow-xl shadow-orange-900/5 backdrop-blur">
      <div className="border-b border-orange-100 bg-orange-500 px-8 py-4">
        <div className="flex items-center gap-3">
          <div className="flex gap-1.5">
            <div className="h-3 w-3 rounded-full bg-white/40" /><div className="h-3 w-3 rounded-full bg-white/40" /><div className="h-3 w-3 rounded-full bg-white/40" />
          </div>
          <p className="text-sm font-black text-white">Pratinjau: {docType.label}</p>
        </div>
      </div>
      <div className="p-8 md:p-12 print:p-0">
        <div className="mx-auto w-full max-w-none font-serif text-slate-900">{renderDoc()}</div>
      </div>
    </div>
  );
}