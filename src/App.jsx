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
import ImageModule from "docxtemplater-image-module-free";
import { saveAs } from "file-saver";
import { renderAsync } from "docx-preview";

// ─── DATA ────────────────────────────────────────────────────────────────────

const DOC_TYPES = [
  { id: "daftar-hadir",  icon: <ClipboardList />, label: "Daftar Hadir", desc: "Fitur dikunci", color: "orange", disabled: true, lockedMessage: "Fitur Daftar Hadir dikunci" },
  { id: "tanda-terima",  icon: <Briefcase />,     label: "Tanda Terima", desc: "Fitur dikunci", color: "amber", disabled: true, lockedMessage: "Fitur Tanda Terima dikunci" },
  { id: "surat-pernyataan-kendaraan", icon: <Car />, label: "Super Kendis", desc: "Fitur dikunci", color: "orange", disabled: true, lockedMessage: "Fitur Super Kendis dikunci" },
  { id: "pengeluaran-riil", icon: <Receipt />,    label: "DPR", desc: "Fitur dikunci", color: "amber", disabled: true, lockedMessage: "Fitur DPR dikunci" },
  { id: "spj",           icon: <FileText />,      label: "SPJ", desc: "Fitur dibuka sementara", color: "orange", disabled: false, lockedMessage: "Fitur SPJ dikunci" },
  { id: "spd",           icon: <MapIcon />,           label: "SPD", desc: "Fitur dikunci", color: "amber", disabled: true, lockedMessage: "Fitur SPD dikunci" },
  { id: "surat-tugas",   icon: <Users />,         label: "Surtug", desc: "Fitur dikunci", color: "orange", disabled: true, lockedMessage: "Fitur Surat Tugas dikunci" },
  { id: "bapp",          icon: <FileText />,      label: "BAPP", desc: "BAPP PML/PPL", color: "amber" },
  { id: "bast",          icon: <FileText />,      label: "BAST", desc: "BAST PML/PPL", color: "amber" },
  { id: "surat-pernyataan-penyelesaian-lapangan", icon: <FileText />, label: "Surat Pernyataan Penyelesaian Lapangan", desc: "Khusus PML", color: "amber" },
  { id: "lampiran",      icon: <FileText />,      label: "Lampiran", desc: "Lampiran SPK PML/PPL", color: "amber" },
  { id: "gabungan-pembayaran", icon: <Receipt />, label: "Gabungan Administrasi Pembayaran", desc: "Generate satu berkas pembayaran lengkap per PML atau PPL", color: "amber" },
  {
  id: "surat-kepala",
  icon: <FileText />,
  label: "Surat Kepala",
  desc: "SPEPL Kepala BPS (PML & PPL)",
  color: "amber",
},
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
      const text = String(value ?? "").trim();
      // Sel formula Excel/Google Sheets seperti #N/A jangan dianggap sebagai data.
      if (text && !/^#(?:N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NUM!|NULL!)$/i.test(text)) return text;
    }
    return "";
  };

  const nama = get("nama", "nama lengkap", "nama_lengkap", "nama petugas", "nama peserta", "nama_petugas");
  const jabatan = get("jabatan", "posisi", "jenis petugas", "role", "kategori");
  const jabatanUpper = jabatan.toUpperCase();
  const isPml = /PML|PENGAWAS/.test(jabatanUpper);
  const isPpl = /PPL|PENCACAH/.test(jabatanUpper);
  const email = get("email", "email petugas", "email peserta", "mail");
  const emailPengawas = get("email pengawas", "email pml", "mail pengawas", "mail pml");
  const namaPengawas = get("nama pengawas", "nama pml", "pengawas");
  const namaPplKolom = get("nama ppl", "pencacah", "nama pencacah");
  const nomorSpk = get("nomor spk", "nomor_spk", "nomor kontrak", "nomor_kontrak", "spk");
  const slsOngoing = get(
    "sls (selesai + sedang dikerjakan)",
    "sls selesai + sedang dikerjakan",
    "sls (selesai dan sedang dikerjakan)",
    "sls selesai dan sedang dikerjakan",
    "sls ongoing",
    "sls_ongoing"
  );
  const prelistTotal = get("prelist total", "prelist_total", "target prelist", "target_prelist");
  const realisasiTotal = get(
    "realisasi total",
    "realisasi tot",
    "realisasi_total",
    "realisasi hasil pendataan",
    "realisasi_hasil",
    "realisasi",
    "realisasi_hasil_pendataan"
  );
  const persentasePrelist = get(
    "persentase prelist",
    "persentase pendataan",
    "persentase realisasi",
    "persentase",
    "persentase p",
    "persentase_p",
    "persentase_prelist"
  );

  return {
    no: get("no", "nomor"),
    nama,
    jabatan,
    jabatan_raw: jabatan,
    wilayah: get("wilayah", "wil tugas", "wil. tugas", "wilayah tugas", "kecamatan", "asal"),
    kecamatan: get("kecamatan", "wilayah", "wil tugas", "wil. tugas", "wilayah tugas", "asal"),
    email,
    username_sobat: get("username sobat", "username_sobat", "username", "sobat id", "sobatid"),
    sobat_id: get("sobat id", "sobatid", "username sobat", "username_sobat", "username"),
    kelas: get("kelas"),
    gelombang: get("gelombang"),
    tempat: get("tempat", "hotel", "tc"),
    telp: get("telp", "no hp", "nomor hp"),

    // Untuk baris PML, identitas PML adalah Nama Lengkap/Email baris itu sendiri.
    // Untuk baris PPL, nama dan email pengawas diambil dari kolom relasi pada sheet yang sama.
    nama_pengawas: namaPengawas || (isPml ? nama : ""),
    email_pengawas: emailPengawas || (isPml ? email : ""),
    nama_pml: namaPengawas || (isPml ? nama : ""),
    email_pml: emailPengawas || (isPml ? email : ""),
    nama_ppl: namaPplKolom || (isPpl ? nama : ""),
    email_ppl: isPpl ? email : get("email ppl", "email pencacah"),

    nik: get("nik", "nik petugas", "nik peserta"),
    nomor_spk: nomorSpk,
    nomor_kontrak: nomorSpk,

    sls_total: get("sls total", "sls_total"),
    sls_40: get("sls 40%", "sls 40", "sls_40", "sls40", "sls 40 persen"),
    sls_60: get("sls 60%", "sls 60", "sls_60", "sls60", "sls 60 persen"),
    sls_ongoing: slsOngoing,
    sls_selesai_sedang_dikerjakan: slsOngoing,
    persentase_sls: get("persentase sls", "persentase_sls", "% sls"),
    tanggal_screenshot: get("tanggal screenshot", "tanggal_screenshot"),

    prelist_total: prelistTotal,
    target_prelist: prelistTotal,
    realisasi_total: realisasiTotal,
    realisasi_hasil_pendataan: realisasiTotal,
    persentase_prelist: persentasePrelist,
    persentase_pendataan: persentasePrelist,
    flag: get("flag", "status flag"),
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

// ─── STATUS SLS PARSER ───────────────────────────────────────────────────────
// Sheet "Status SLS" dipakai khusus untuk menyaring tabel pada BERKAS
// PEMBAYARAN gabungan. Hanya SLS dengan status "Selesai" atau
// "Sedang Dikerjakan" yang akan dihitung dan ditampilkan.
function normalizeStatusSlsCode(value, width = 0) {
  const digits = String(value ?? "").replace(/[^0-9]/g, "");
  if (!digits) return "";
  if (!width) return digits;
  return digits.padStart(width, "0").slice(-width);
}

function normalizeStatusSlsLabel(value) {
  return String(value ?? "")
    .replace(/^\s*\[[^\]]*\]\s*/g, "")
    .replace(/^\s*[-:–—]+\s*/g, "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ");
}

function isAllowedStatusSls(value) {
  const status = normalizeStatusSlsLabel(value);
  return status === "SELESAI" || status === "SEDANG DIKERJAKAN";
}

function normalizeStatusSlsRow(row) {
  const normalized = {};
  Object.entries(row || {}).forEach(([key, value]) => {
    normalized[String(key ?? "").trim().toLowerCase().replace(/\s+/g, " ")] = String(value ?? "").trim();
  });

  const get = (...keys) => {
    for (const key of keys) {
      const value = normalized[key];
      if (value != null && String(value).trim() !== "") return String(value).trim();
    }
    return "";
  };

  const statusRaw = get("status", "status sls", "status_sls", "keterangan status", "keterangan");
  const kodeKecamatan = get("kode kecamatan", "kdkec", "kode_kecamatan", "kec");
  const kodeKelurahan = get("kode kelurahan", "kddesa", "kode_kelurahan", "desa");
  const kodeSls = get("kode sls", "kode_sls", "kdsls", "sls");

  return {
    nama_pml: get("nama pml", "nama_pml", "pengawas", "pml"),
    email_pml: get("email pml", "email_pml", "email pengawas", "mail pml", "mail pengawas"),
    nama_ppl: get("nama ppl", "nama_ppl", "pencacah", "ppl"),
    email_ppl: get("email ppl", "email_ppl", "email pencacah", "mail ppl", "mail pencacah"),

    kdkec: normalizeStatusSlsCode(kodeKecamatan, 3),
    kddesa: normalizeStatusSlsCode(kodeKelurahan, 3),
    kode_sls: normalizeStatusSlsCode(kodeSls, 6),

    jumlah_prelist: get("jumlah prelist", "prelist total", "prelist_total", "target prelist"),
    jumlah_realisasi: get("jumlah realisasi", "realisasi total", "realisasi_total", "realisasi"),
    persentase: get("persentase", "persentase pendataan", "persentase_pendataan"),
    status: normalizeStatusSlsLabel(statusRaw),
    status_raw: statusRaw,
    status_diperbolehkan: isAllowedStatusSls(statusRaw),
  };
}

function parseStatusSlsData(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames.find(
    (name) => String(name ?? "").trim().toLowerCase() === "status sls"
  );

  if (!sheetName) {
    console.warn("Sheet bernama 'Status SLS' tidak ditemukan; tabel gabungan tidak akan memakai filter status.");
    return [];
  }

  const sheet = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  return raw
    .map(normalizeStatusSlsRow)
    .filter((row) => row.nama_pml || row.nama_ppl || row.email_pml || row.email_ppl || row.kode_sls || row.status);
}

// ─── DATA PER SLS PARSER ─────────────────────────────────────────────────────
// Sheet "Data per SLS" memiliki tiga baris header bertingkat. Karena itu parser
// membaca nilai berdasarkan posisi kolom A:T, bukan berdasarkan nama header.
function cleanDataPerSlsCell(value) {
  const text = String(value ?? "").trim();
  if (!text || /^#(?:N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NUM!|NULL!)$/i.test(text)) return "";
  return text;
}

function normalizeDataPerSlsRowFromArray(row = [], fallbackNo = 0) {
  const get = (index) => cleanDataPerSlsCell(row?.[index]);
  const keterangan = get(19);

  return {
    no_sumber: get(0) || String(fallbackNo || ""),
    nama_pml: get(1),
    username_pml: get(2),
    nama_ppl: get(3),
    username_ppl: get(4),

    kdkec: normalizeStatusSlsCode(get(5), 3),
    kddesa: normalizeStatusSlsCode(get(6), 3),
    kode_sls: normalizeStatusSlsCode(get(7), 6),

    target_keluarga: get(8),
    target_usaha: get(9),
    target_jumlah: get(10),

    realisasi_dengan_tidak_ditemukan_keluarga: get(11),
    realisasi_dengan_tidak_ditemukan_usaha: get(12),
    realisasi_dengan_tidak_ditemukan_jumlah: get(13),
    persentase_dengan_tidak_ditemukan: get(14),

    realisasi_tanpa_tidak_ditemukan_keluarga: get(15),
    realisasi_tanpa_tidak_ditemukan_usaha: get(16),
    realisasi_tanpa_tidak_ditemukan_jumlah: get(17),
    persentase_tanpa_tidak_ditemukan: get(18),

    keterangan,
    status: normalizeStatusSlsLabel(keterangan),
  };
}

function parseDataPerSlsValues(values = []) {
  const rows = Array.isArray(values) ? values : [];

  // Cari baris data pertama. Baris nomor kolom seperti -1, -2, dan seterusnya
  // tidak dianggap data karena tidak memuat nama petugas.
  const firstDataIndex = rows.findIndex((row) => {
    const no = cleanDataPerSlsCell(row?.[0]);
    const namaPml = cleanDataPerSlsCell(row?.[1]);
    const namaPpl = cleanDataPerSlsCell(row?.[3]);
    const hasPersonName = /[A-Za-z]/.test(`${namaPml} ${namaPpl}`);
    const hasCode = [5, 6, 7].some((index) => /\d/.test(cleanDataPerSlsCell(row?.[index])));
    return /^\d+(?:[.,]0+)?$/.test(no) && hasPersonName && hasCode;
  });

  if (firstDataIndex < 0) return [];

  return rows
    .slice(firstDataIndex)
    .map((row, index) => normalizeDataPerSlsRowFromArray(row, index + 1))
    .filter((row) =>
      row.nama_pml || row.nama_ppl || row.username_pml || row.username_ppl ||
      row.kdkec || row.kddesa || row.kode_sls
    );
}

function parseDataPerSlsData(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames.find(
    (name) => String(name ?? "").trim().toLowerCase() === "data per sls"
  );

  if (!sheetName) {
    console.warn("Sheet bernama 'Data per SLS' tidak ditemukan; tabel beban kerja pada berkas pembayaran akan kosong.");
    return [];
  }

  const sheet = workbook.Sheets[sheetName];
  const values = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
    blankrows: false,
  });
  return parseDataPerSlsValues(values);
}

// ─── DATA PML PROGRESS PARSER ────────────────────────────────────────────────
// Sheet "Data PML Progress" juga punya header bertingkat seperti "Data per SLS",
// jadi dibaca berdasarkan posisi kolom (array), bukan nama header.
// Kolom L (index 11, 0-based) = "Realisasi Jumlah (Dengan Tidak Ditemukan) > Jumlah"
// itulah nilai yang dipakai sebagai realisasi PML untuk Surat Kepala.
function normalizeDataPmlProgressRowFromArray(row = [], fallbackNo = 0) {
  const get = (index) => cleanDataPerSlsCell(row?.[index]);
  return {
    no: get(0) || String(fallbackNo || ""),
    nama_pml: get(1),
    username_sobat_pml: get(2),
    beban_sls_1: get(3),
    beban_sls_2: get(4),
    persentase_sls: get(5),
    target_prelist_awal: get(6),
    kolom_h: get(7),
    jumlah_target: get(8),      // kolom I
    target_keluarga: get(9),    // kolom J
    target_usaha: get(10),      // kolom K
    realisasi_jumlah: get(11),  // kolom L <-- REALISASI YANG DIPAKAI
    persentase_realisasi: get(12), // kolom M
    keterangan: get(16),        // kolom Q
  };
}

function parseDataPmlProgressValues(values = []) {
  const rows = Array.isArray(values) ? values : [];

  // Baris data pertama = baris yang kolom No-nya angka DAN kolom Nama berisi huruf.
  const firstDataIndex = rows.findIndex((row) => {
    const no = cleanDataPerSlsCell(row?.[0]);
    const nama = cleanDataPerSlsCell(row?.[1]);
    const hasPersonName = /[A-Za-z]/.test(nama);
    return /^\d+(?:[.,]0+)?$/.test(no) && hasPersonName;
  });

  if (firstDataIndex < 0) return [];

  return rows
    .slice(firstDataIndex)
    .map((row, index) => normalizeDataPmlProgressRowFromArray(row, index + 1))
    .filter((row) => row.nama_pml);
}

function parseDataPmlProgressData(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames.find(
    (name) => String(name ?? "").trim().toLowerCase() === "data pml progress"
  );

  if (!sheetName) {
    console.warn("Sheet bernama 'Data PML Progress' tidak ditemukan; realisasi PML pada Surat Kepala akan pakai fallback lama.");
    return [];
  }

  const sheet = workbook.Sheets[sheetName];
  const values = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: "",
    raw: false,
    blankrows: false,
  });
  return parseDataPmlProgressValues(values);
}

// ─── APPROVE BY PML PARSER ──────────────────────────────────────────────────
// Untuk berkas PML, nilai realisasi/jumlah pemeriksaan bersumber dari kolom
// "Jumlah Approve PML" pada sheet "Approve by PML".
function normalizeApproveByPmlRow(row = {}) {
  const normalized = {};
  Object.entries(row || {}).forEach(([key, value]) => {
    const normalizedKey = String(key ?? "")
      .trim()
      .toLowerCase()
      .replace(/[_-]+/g, " ")
      .replace(/\s+/g, " ");
    normalized[normalizedKey] = cleanDataPerSlsCell(value);
  });

  const get = (...keys) => {
    for (const key of keys) {
      const value = normalized[String(key ?? "").trim().toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " ")];
      if (cleanDataPerSlsCell(value)) return cleanDataPerSlsCell(value);
    }
    return "";
  };

  const namaSls = get("nama sls", "sls", "nama_sls");
  const kodeDalamKurung = String(namaSls).match(/\[\s*(\d{1,6})\s*\]/);
  const kodeSlsLangsung = get("kode sls", "kode_sls", "kdsls");
  const kodeSls = normalizeStatusSlsCode(
    kodeSlsLangsung || (kodeDalamKurung ? kodeDalamKurung[1] : ""),
    6
  );

  return {
    nama_sls: namaSls,
    kode_sls: kodeSls,
    email_ppl: get("email ppl", "email pencacah", "email_ppl"),
    nama_ppl: get("nama ppl", "nama pencacah", "nama_ppl"),
    email_pml: get("email pml", "email pengawas", "email_pml"),
    nama_pml: get("nama pml", "nama pengawas", "nama_pml"),
    selesai: get("selesai", "status selesai", "status"),
    tanggal_screen: get("tanggal screen", "tanggal screenshot", "tanggal"),
    jumlah_submit: get("jumlah submit", "jumlah_submit"),
    jumlah_approve_pml: get(
      "jumlah approve pml",
      "jumlah approve by pml",
      "jumlah_approve_pml",
      "approve pml"
    ),
    submitted_by: get("submitted by", "submitted_by", "pengirim"),
    submitted_by_role: get("submitted by 1", "submitted by_1", "role submitted by", "submitted by role"),
    waktu_submit: get("waktu submit", "waktu_submit"),
    catatan: get("catatan", "keterangan"),
    // 🔥 BARU: kolom "Foto Bukti" bisa berisi lebih dari satu link (dipisah baris
    // baru/koma/titik koma). Disimpan sebagai array URL, siap dipakai loop gambar.
    foto_bukti: splitFotoBuktiUrls(get("foto bukti", "foto_bukti", "link foto", "foto")),
  };
}

// ─── FOTO BUKTI (GOOGLE DRIVE) ───────────────────────────────────────────────
// ─── FOTO BUKTI DARI SPREADSHEET TERPISAH (Database SLS) ────────────────────
// Link foto tidak ada di sheet "Approve by PML" yang dipakai utama, tapi ada di
// spreadsheet lain bernama "Database SLS [JANGAN DIUBAH]", tab "Submission-V2".
// 🔥 BARU: tab ini memisahkan foto jadi 2 kolom sendiri-sendiri, "Foto Bukti PPL"
// dan "Foto Bukti PML" (sebelumnya cuma satu kolom "Foto Bukti" gabungan).
// Data ini diambil terpisah lalu digabungkan ke approveByPmlRows berdasarkan
// Email PML + Email PPL.
const FOTO_BUKTI_SPREADSHEET_ID = "1U694SejnIYezDRgy6Ao_1Moik4ckW7iMBWJeOmgpkcI";
const FOTO_BUKTI_SPREADSHEET_EXPORT_URL = `https://docs.google.com/spreadsheets/d/${FOTO_BUKTI_SPREADSHEET_ID}/export?format=xlsx`;
const FOTO_BUKTI_SHEET_NAME = "Submission-V2";

// Cache supaya spreadsheet foto tidak di-fetch berulang kali dalam satu sesi.
let fotoBuktiDatabaseSlsCache = null;

function normalizeFotoBuktiRow(row = {}) {
  const normalized = {};
  Object.entries(row || {}).forEach(([key, value]) => {
    const normalizedKey = String(key ?? "").trim().toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " ");
    normalized[normalizedKey] = cleanDataPerSlsCell(value);
  });
  const get = (...keys) => {
    for (const key of keys) {
      const value = normalized[String(key ?? "").trim().toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " ")];
      if (cleanDataPerSlsCell(value)) return cleanDataPerSlsCell(value);
    }
    return "";
  };
  return {
    email_ppl: get("email ppl", "email_ppl"),
    email_pml: get("email pml", "email_pml"),
    // 🔥 BARU: dipisah per role. foto_bukti_ppl dari kolom "Foto Bukti PPL",
    // foto_bukti_pml dari kolom "Foto Bukti PML".
    foto_bukti_ppl: splitFotoBuktiUrls(get("foto bukti ppl", "foto_bukti_ppl")),
    foto_bukti_pml: splitFotoBuktiUrls(get("foto bukti pml", "foto_bukti_pml")),
  };
}

async function fetchFotoBuktiRowsFromDatabaseSls() {
  if (fotoBuktiDatabaseSlsCache) return fotoBuktiDatabaseSlsCache;

  const response = await fetch(FOTO_BUKTI_SPREADSHEET_EXPORT_URL);
  if (!response.ok) throw new Error(`Gagal memuat Database SLS: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const workbook = XLSX.read(arrayBuffer, { type: "array" });

  const sheetName = workbook.SheetNames.find(
    (name) => String(name ?? "").trim().toLowerCase() === FOTO_BUKTI_SHEET_NAME.toLowerCase()
  );
  if (!sheetName) {
    console.warn(`Sheet '${FOTO_BUKTI_SHEET_NAME}' tidak ditemukan di Database SLS. Sheet tersedia:`, workbook.SheetNames);
    fotoBuktiDatabaseSlsCache = [];
    return fotoBuktiDatabaseSlsCache;
  }

  const sheet = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  const rows = raw.map(normalizeFotoBuktiRow).filter((row) => row.email_pml || row.email_ppl);

  fotoBuktiDatabaseSlsCache = rows;
  return rows;
}

// Map: "EMAIL_PML::EMAIL_PPL" -> { fotoPml: Set, fotoPpl: Set } (satu pasangan
// PML+PPL bisa punya beberapa foto dari beberapa baris SLS berbeda).
// 🔥 BARU: foto PML dan foto PPL disimpan terpisah karena sumbernya sekarang
// dua kolom berbeda ("Foto Bukti PML" vs "Foto Bukti PPL").
function buildFotoBuktiMapByPmlPpl(fotoBuktiRows = []) {
  const map = new Map();
  for (const row of fotoBuktiRows || []) {
    const emailPml = upperText(row.email_pml);
    const emailPpl = upperText(row.email_ppl);
    if (!emailPml && !emailPpl) continue;
    const key = `${emailPml}::${emailPpl}`;
    if (!map.has(key)) map.set(key, { fotoPml: new Set(), fotoPpl: new Set() });
    const entry = map.get(key);
    for (const url of row.foto_bukti_pml || []) entry.fotoPml.add(url);
    for (const url of row.foto_bukti_ppl || []) entry.fotoPpl.add(url);
  }
  return map;
}

function mergeFotoBuktiIntoApproveByPmlRows(approveByPmlRows = [], fotoBuktiMap) {
  if (!fotoBuktiMap || fotoBuktiMap.size === 0) return approveByPmlRows;
  return (approveByPmlRows || []).map((row) => {
    const key = `${upperText(row.email_pml)}::${upperText(row.email_ppl)}`;
    const entry = fotoBuktiMap.get(key);
    if (!entry) return row;
    const mergedPml = new Set([...(row.foto_bukti_pml || []), ...entry.fotoPml]);
    const mergedPpl = new Set([...(row.foto_bukti_ppl || []), ...entry.fotoPpl]);
    return { ...row, foto_bukti_pml: [...mergedPml], foto_bukti_ppl: [...mergedPpl] };
  });
}

async function enrichApproveByPmlWithFotoBukti(approveByPmlRows = []) {
  try {
    const fotoBuktiRows = await fetchFotoBuktiRowsFromDatabaseSls();
    const fotoBuktiMap = buildFotoBuktiMapByPmlPpl(fotoBuktiRows);
    return mergeFotoBuktiIntoApproveByPmlRows(approveByPmlRows, fotoBuktiMap);
  } catch (err) {
    console.warn("Gagal memuat foto dari Database SLS:", err.message);
    return approveByPmlRows;
  }
}

function splitFotoBuktiUrls(value) {
  return String(value ?? "")
    .split(/[\n,;]+/)
    .map((item) => item.trim())
    .filter((item) => /^https?:\/\//i.test(item));
}

function chunkFotoBuktiIntoRows(fotoEntries = [], perRow = 3) {
  // Satu object = satu baris tabel Word dengan 3 kolom tetap.
  // Contoh 5 foto => baris 1: foto1-3, baris 2: foto4-5 + satu slot kosong.
  const entries = Array.isArray(fotoEntries)
    ? fotoEntries.filter(Boolean)
    : [];

  const rows = [];

  for (let i = 0; i < entries.length; i += perRow) {
    const slice = entries.slice(i, i + perRow);

    rows.push({
      foto1: getFotoUrlFromTagValue(slice[0]),
      foto2: getFotoUrlFromTagValue(slice[1]),
      foto3: getFotoUrlFromTagValue(slice[2]),
    });
  }

  return rows;
}

function extractGoogleDriveFileId(url) {
  const text = String(url ?? "");
  const patterns = [/\/d\/([a-zA-Z0-9_-]{15,})/, /[?&]id=([a-zA-Z0-9_-]{15,})/];
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match) return match[1];
  }
  return "";
}

// Beberapa bentuk URL Google Drive dicoba berurutan karena tidak semua endpoint
// selalu mengizinkan akses langsung (CORS) dari browser.
function buildGoogleDriveImageUrlCandidates(url) {
  const fileId = extractGoogleDriveFileId(url);
  const candidates = [];
  if (fileId) {
    candidates.push(`https://lh3.googleusercontent.com/d/${fileId}`);
    candidates.push(`https://drive.google.com/uc?export=view&id=${fileId}`);
    candidates.push(`https://drive.google.com/uc?export=download&id=${fileId}`);
  }
  if (url) candidates.push(url);
  return candidates;
}

// Cache supaya foto yang sama tidak diunduh berulang kali.
const fotoBuktiArrayBufferCache = new Map();

// PNG transparan 1x1 sebagai fallback, dipakai bila sebuah foto gagal diunduh
// (link rusak/tidak publik) supaya proses generate dokumen tidak gagal total.
const FALLBACK_FOTO_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
function base64ToArrayBuffer(base64) {
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes.buffer;
}

async function fetchFotoBuktiArrayBuffer(url) {
  const key = String(url ?? "").trim();
  if (!key) return null;
  if (fotoBuktiArrayBufferCache.has(key)) return fotoBuktiArrayBufferCache.get(key);

  let result = null;
  for (const candidateUrl of buildGoogleDriveImageUrlCandidates(key)) {
    try {
      const response = await fetch(candidateUrl);
      if (!response.ok) continue;
      const buffer = await response.arrayBuffer();
      if (buffer && buffer.byteLength > 0) {
        result = buffer;
        break;
      }
    } catch (err) {
      console.warn(`Gagal memuat foto bukti dari ${candidateUrl}:`, err);
    }
  }

  fotoBuktiArrayBufferCache.set(key, result);
  return result;
}


// ---------------------------------------------------------------------------
// FIX PPL: NORMALISASI TAG FOTO + PREFETCH SEBELUM DOCXTEMPLATER RENDER
// ---------------------------------------------------------------------------
// docxtemplater-image-module-free dapat bermasalah ketika getImage() async
// dipanggil di dalam loop gambar. Karena itu semua foto diunduh lebih dulu,
// disimpan di cache, kemudian image module membaca cache secara synchronous.
function getFotoUrlFromTagValue(tagValue) {
  if (!tagValue) return "";

  if (typeof tagValue === "string") {
    return tagValue.trim();
  }

  if (typeof tagValue === "object" && tagValue.url) {
    return String(tagValue.url).trim();
  }

  return "";
}

function collectFotoUrlsFromTemplateData(templateData = {}) {
  const urls = new Set();

  const addUrl = (value) => {
    const url = getFotoUrlFromTagValue(value);
    if (url && /^https?:\/\//i.test(url)) {
      urls.add(url);
    }
  };

  // Kompatibilitas dengan struktur foto lama.
  for (const item of templateData?.foto || []) {
    addUrl(item);
  }

  for (const item of templateData?.foto_bukti || []) {
    addUrl(item);
  }

  // Struktur grid baru: 3 foto per baris.
  for (const row of templateData?.foto_rows || []) {
    addUrl(row?.foto1);
    addUrl(row?.foto2);
    addUrl(row?.foto3);

    // Fallback untuk data lama yang masih memakai row.slot[].
    if (Array.isArray(row?.slot)) {
      for (const slot of row.slot) {
        addUrl(slot);
      }
    }
  }

  return [...urls];
}

async function prefetchFotoBuktiForTemplate(templateData = {}) {
  const urls = collectFotoUrlsFromTemplateData(templateData);

  console.log(`Prefetch ${urls.length} foto bukti sebelum render DOCX`);

  if (urls.length === 0) return;

  await Promise.all(
    urls.map(async (url) => {
      try {
        const buffer = await fetchFotoBuktiArrayBuffer(url);
        if (!buffer) {
          console.warn("Foto gagal dimuat, akan memakai fallback transparan:", url);
        }
      } catch (err) {
        console.warn("Prefetch foto gagal, akan memakai fallback transparan:", url, err);
      }
    })
  );
}

const DOCX_MIME_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

// 🔥 FIX: setelah render lewat renderAsync + modul gambar, doc.getZip().generate({type:"blob"})
// kadang menghasilkan objek yang gagal dikenali sebagai Blob asli oleh browser
// (URL.createObjectURL melempar "Overload resolution failed"). Solusinya, ambil hasilnya
// sebagai arraybuffer lalu bungkus manual pakai konstruktor Blob bawaan browser.
function zipToDocxBlob(zip) {
  const arrayBuffer = zip.generate({ type: "arraybuffer" });
  return new Blob([arrayBuffer], { type: DOCX_MIME_TYPE });
}

// Modul gambar docxtemplater — dibuat lewat fungsi (bukan instance tunggal)
// supaya aman dipakai berulang untuk banyak dokumen dalam satu batch download.
function createFotoBuktiImageModule() {
  return new ImageModule({
    centered: false,

    // PENTING: getImage synchronous.
    // Semua foto sudah di-prefetch sebelum doc.render().
    getImage: (tagValue) => {
      const url = getFotoUrlFromTagValue(tagValue);

      if (!url) {
        return base64ToArrayBuffer(FALLBACK_FOTO_BASE64);
      }

      const cached = fotoBuktiArrayBufferCache.get(url);

      if (cached instanceof ArrayBuffer && cached.byteLength > 0) {
        return cached;
      }

      if (ArrayBuffer.isView(cached) && cached.byteLength > 0) {
        return cached;
      }

      console.warn("Foto belum tersedia di cache, memakai fallback transparan:", url);
      return base64ToArrayBuffer(FALLBACK_FOTO_BASE64);
    },

    getSize: (img, tagValue) => {
      const url = getFotoUrlFromTagValue(tagValue);
      if (!url) return [1, 1];

      // Ukuran seragam agar tiga foto muat dan rapi dalam satu baris tabel Word.
      return [180, 135];
    },
  });
}

// PML => dicocokkan lewat Email/Nama PML (SELURUH foto PPL yang dia bawahi).
// PPL => dicocokkan lewat Email/Nama PPL (foto khusus miliknya sendiri).
function matchApproveByPmlRowForRole(approveRow, row, role) {
  const isPml = upperText(role) === "PML";
  const targetEmail = upperText(cleanText(row?.email || ""));
  const targetName = upperText(cleanText(row?.nama || ""));
  const approveEmail = upperText(cleanText(isPml ? approveRow?.email_pml : approveRow?.email_ppl));
  const approveName = upperText(cleanText(isPml ? approveRow?.nama_pml : approveRow?.nama_ppl));
  if (targetEmail && approveEmail) return targetEmail === approveEmail;
  if (targetName && approveName) return targetName === approveName;
  return false;
}

function filterApproveByPmlRowsForBappRow(approveByPmlRows = [], row = {}, role = "PML") {
  return (approveByPmlRows || []).filter((approveRow) => matchApproveByPmlRowForRole(approveRow, row, role));
}

// Kumpulkan URL unik. Template Word memakai {#foto_rows}{%foto1} | {%foto2} | {%foto3}{/foto_rows}.
// 🔥 BARU: role menentukan kolom foto yang dipakai — generate PML memakai
// foto_bukti_pml (foto dari SEMUA PPL di bawah PML tsb, karena approveRows yang
// dikirim ke sini sudah difilter per-PML sebelumnya), generate PPL memakai
// foto_bukti_ppl (foto milik PPL itu sendiri saja).
function collectFotoBuktiFromApproveRows(approveRows = [], role = "PML") {
  const isPml = upperText(role) === "PML";
  const seen = new Set();
  const entries = [];
  for (const approveRow of approveRows || []) {
    const urls = isPml ? approveRow?.foto_bukti_pml : approveRow?.foto_bukti_ppl;
    for (const url of urls || []) {
      const key = String(url ?? "").trim();
      if (!key || seen.has(key)) continue;
      seen.add(key);
      entries.push({ url: key });
    }
  }
  return entries;
}

function parseApproveByPmlData(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheetName = workbook.SheetNames.find(
    (name) => String(name ?? "").trim().toLowerCase() === "approve by pml"
  );

  if (!sheetName) {
    console.warn("Sheet bernama 'Approve by PML' tidak ditemukan; jumlah PML akan memakai fallback Data per SLS.");
    return [];
  }

  const sheet = workbook.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  return raw
    .map(normalizeApproveByPmlRow)
    .filter((row) =>
      row.nama_sls || row.kode_sls || row.nama_pml || row.email_pml ||
      row.nama_ppl || row.email_ppl || row.jumlah_approve_pml
    );
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

function findStatusSlsSheetFromRows(sheetRows = []) {
  // Prioritaskan nama tab persis agar tidak tertukar dengan sheet lain.
  const exact = (sheetRows || []).find(
    (sheet) => String(sheet?.sheetName ?? "").trim().toLowerCase() === "status sls"
  );
  if (exact && Array.isArray(exact.rows) && exact.rows.length > 0) {
    const headers = Object.keys(exact.rows[0]).map((h) => String(h ?? "").trim());
    return { sheet: exact, data: exact.rows, headers, sheetName: exact.sheetName };
  }

  // Fallback berdasarkan struktur header.
  for (const sheet of sheetRows || []) {
    const rows = Array.isArray(sheet.rows) ? sheet.rows : [];
    if (rows.length === 0) continue;
    const headers = Object.keys(rows[0]).map((h) => String(h ?? "").trim());
    const keys = headers.map((h) => h.toLowerCase());
    const hasPml = keys.some((h) => ["nama pml", "nama_pml", "pengawas"].includes(h));
    const hasPpl = keys.some((h) => ["nama ppl", "nama_ppl", "pencacah"].includes(h));
    const hasKodeSls = keys.some((h) => ["kode sls", "kode_sls", "kdsls"].includes(h));
    const hasStatus = keys.some((h) => ["status", "status sls", "status_sls"].includes(h));
    if (hasPml && hasPpl && hasKodeSls && hasStatus) {
      return { sheet, data: rows, headers, sheetName: sheet.sheetName };
    }
  }
  return null;
}

function findDataPerSlsSheetFromRows(sheetRows = []) {
  const exact = (sheetRows || []).find(
    (sheet) => String(sheet?.sheetName ?? "").trim().toLowerCase() === "data per sls"
  );
  if (exact && Array.isArray(exact.values) && exact.values.length > 0) return exact;

  // Fallback bila nama tab sedikit berubah, tetapi struktur header tetap sama.
  return (sheetRows || []).find((sheet) => {
    const headerText = (sheet?.values || [])
      .slice(0, 4)
      .flat()
      .map((value) => String(value ?? "").trim().toLowerCase())
      .join(" ");
    return headerText.includes("target prelist awal") &&
      headerText.includes("username sobat") &&
      headerText.includes("sls/sub-sls");
  }) || null;
}

function findDataPmlProgressSheetFromRows(sheetRows = []) {
  const exact = (sheetRows || []).find(
    (sheet) => String(sheet?.sheetName ?? "").trim().toLowerCase() === "data pml progress"
  );
  if (exact && Array.isArray(exact.values) && exact.values.length > 0) return exact;

  return (sheetRows || []).find((sheet) => {
    const headerText = (sheet?.values || [])
      .slice(0, 4)
      .flat()
      .map((value) => String(value ?? "").trim().toLowerCase())
      .join(" ");
    return headerText.includes("username sobat") &&
      headerText.includes("realisasi jumlah") &&
      headerText.includes("dengan tidak ditemukan");
  }) || null;
}

function findApproveByPmlSheetFromRows(sheetRows = []) {
  const exact = (sheetRows || []).find(
    (sheet) => String(sheet?.sheetName ?? "").trim().toLowerCase() === "approve by pml"
  );
  if (exact && Array.isArray(exact.rows) && exact.rows.length > 0) return exact;

  return (sheetRows || []).find((sheet) => {
    const rows = Array.isArray(sheet?.rows) ? sheet.rows : [];
    if (rows.length === 0) return false;
    const headers = Object.keys(rows[0] || {})
      .map((value) => String(value ?? "").trim().toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " "));
    const hasPml = headers.includes("nama pml") || headers.includes("email pml");
    const hasPpl = headers.includes("nama ppl") || headers.includes("email ppl");
    const hasApprove = headers.some((header) =>
      header === "jumlah approve pml" ||
      header === "jumlah approve by pml" ||
      header === "jumlah approve"
    );
    return hasPml && hasPpl && hasApprove;
  }) || null;
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
      sheetRows.push({ sheetName, rows, values: valuesData.values || [] });
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

    const statusSlsSheetInfo = findStatusSlsSheetFromRows(sheetRows);
    const statusSls = statusSlsSheetInfo
      ? statusSlsSheetInfo.data.map(normalizeStatusSlsRow).filter((row) => row.kode_sls || row.status || row.nama_ppl || row.email_ppl)
      : [];

    const dataPerSlsSheetInfo = findDataPerSlsSheetFromRows(sheetRows);
    const dataPerSls = dataPerSlsSheetInfo
      ? parseDataPerSlsValues(dataPerSlsSheetInfo.values || [])
      : [];

    const approveByPmlSheetInfo = findApproveByPmlSheetFromRows(sheetRows);
    const approveByPml = approveByPmlSheetInfo
      ? (approveByPmlSheetInfo.rows || []).map(normalizeApproveByPmlRow).filter((row) =>
          row.nama_sls || row.kode_sls || row.nama_pml || row.email_pml ||
          row.nama_ppl || row.email_ppl || row.jumlah_approve_pml
        )
      : [];

    const dataPmlProgressSheetInfo = findDataPmlProgressSheetFromRows(sheetRows);
    const dataPmlProgress = dataPmlProgressSheetInfo
      ? parseDataPmlProgressValues(dataPmlProgressSheetInfo.values || [])
      : [];

    console.log("Data PML Progress parsed rows (API):", dataPmlProgress);

    return { data, lampiran, bappData, statusSls, dataPerSls, approveByPml, dataPmlProgress, rawHeaders };
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
  const statusSls = parseStatusSlsData(arrayBuffer);
  const dataPerSls = parseDataPerSlsData(arrayBuffer);
  const approveByPml = parseApproveByPmlData(arrayBuffer);
  const dataPmlProgress = parseDataPmlProgressData(arrayBuffer);

  return { data, lampiran, bappData, statusSls, dataPerSls, approveByPml, dataPmlProgress, rawHeaders };
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────

function cleanText(value) { return String(value ?? "").trim(); }
function upperText(value) { return cleanText(value).toUpperCase(); }

// ─── SELECTION XLSX (fitur "Download Beberapa") ──────────────────────────────
// Memungkinkan pengguna mengunggah file Excel berisi kolom Nama dan/atau Email
// untuk memilih sekumpulan orang yang mau di-generate sekaligus (batch terpilih),
// tanpa harus memilih satu-per-satu lewat dropdown "Download Terpilih".
// Pencocokan diprioritaskan lewat Email (lebih unik), fallback ke Nama (uppercase).

function normalizeSelectionRow(row) {
  const normalized = {};
  Object.entries(row).forEach(([k, v]) => {
    const key = String(k ?? "").trim().toLowerCase().replace(/\s+/g, " ");
    normalized[key] = String(v ?? "").trim();
  });
  return {
    nama: normalized["nama"] || normalized["nama lengkap"] || normalized["nama_lengkap"] || normalized["nama-lengkap"] || "",
    email: normalized["email"] || normalized["mail"] || normalized["e-mail"] || "",
  };
}

function parseSelectionXlsx(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  if (!sheet) return [];
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
  return raw.map(normalizeSelectionRow).filter((r) => r.nama || r.email);
}

function buildSelectionKeySet(selectionRows = []) {
  const set = new Set();
  for (const r of selectionRows) {
    const email = cleanText(r.email);
    const nama = cleanText(r.nama);
    if (email) set.add(`EMAIL::${upperText(email)}`);
    if (nama) set.add(`NAME::${upperText(nama)}`);
  }
  return set;
}

function rowMatchesSelection(selectionKeySet, nama, email) {
  if (!selectionKeySet || selectionKeySet.size === 0) return false;
  const emailClean = cleanText(email);
  if (emailClean && selectionKeySet.has(`EMAIL::${upperText(emailClean)}`)) return true;
  const namaClean = cleanText(nama);
  if (namaClean && selectionKeySet.has(`NAME::${upperText(namaClean)}`)) return true;
  return false;
}

// Versi khusus email-only, dipakai HANYA di fitur Gabungan Administrasi Pembayaran.
function buildEmailOnlySelectionKeySet(selectionRows = []) {
  const set = new Set();
  for (const r of selectionRows) {
    const email = cleanText(r.email);
    if (email) set.add(upperText(email));
  }
  return set;
}

function rowMatchesSelectionByEmail(selectionKeySet, email) {
  if (!selectionKeySet || selectionKeySet.size === 0) return false;
  const emailClean = cleanText(email);
  if (!emailClean) return false;
  return selectionKeySet.has(upperText(emailClean));
}

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
const BERKAS_PEMBAYARAN_PML_TEMPLATE_URL = "/templates/BERKAS PEMBAYARAN PML.docx";
const BERKAS_PEMBAYARAN_PPL_TEMPLATE_URL = "/templates/BERKAS PEMBAYARAN PPL.docx";
const SURAT_KEPALA_TEMPLATE_URL = "/templates/SURAT PERNYATAAN KEPALA BPS.docx";


// ─── TEMPLATE DATA BUILDERS ───────────────────────────────────────────────────

function buildBappTemplateData(formValues, row = {}, role = "PML", approveByPmlRows = []) {
  const tanggalSurat = cleanText(formValues?.tanggal_surat || "");
  const nama = cleanText(row?.nama || row?.nama_pml || row?.nama_ppl || "");
  const jabatan = cleanText(row?.jabatan_raw || row?.jabatan || "");
  const wilayah = cleanText(row?.wilayah || row?.tempat || row?.asal || "");
  const dateParts = getBappDateParts(tanggalSurat);
  const nomorKontrak = cleanText(row?.nomor_spk || row?.nomor_kontrak || formValues?.nomor_kontrak || "");
  const fotoBukti = collectFotoBuktiFromApproveRows(approveByPmlRows, role);
  const fotoRows = chunkFotoBuktiIntoRows(fotoBukti, 3); // ganti 3 -> 2 kalau mau 2 foto/baris
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
    nomor_spk: nomorKontrak,
    nomor_kontrak: nomorKontrak,
    nik: cleanText(row?.nik || ""),
    nama_pml: cleanText(row?.nama_pml || row?.nama_pengawas || (role === "PML" ? nama : "")),
    nama_pengawas: cleanText(row?.nama_pengawas || row?.nama_pml || (role === "PML" ? nama : "")),
    email_pengawas: cleanText(row?.email_pengawas || row?.email_pml || (role === "PML" ? row?.email : "")),
    nama_ppl: cleanText(row?.nama_ppl || (role === "PPL" ? nama : "")),
    sls_total: cleanText(row?.sls_total || ""),
    sls_40: cleanText(row?.sls_40 || ""),
    sls_60: cleanText(row?.sls_60 || ""),
    sls_ongoing: cleanText(row?.sls_ongoing || row?.sls_selesai_sedang_dikerjakan || ""),
    sls_selesai_sedang_dikerjakan: cleanText(row?.sls_ongoing || row?.sls_selesai_sedang_dikerjakan || ""),
    persentase_sls: cleanText(row?.persentase_sls || ""),
    tanggal_screenshot: cleanText(row?.tanggal_screenshot || ""),
    target_prelist: cleanText(row?.target_prelist || row?.prelist_total || ""),
    prelist_total: cleanText(row?.prelist_total || row?.target_prelist || ""),
    realisasi_hasil_pendataan: cleanText(row?.realisasi_hasil_pendataan || row?.realisasi_total || ""),
    realisasi_total: cleanText(row?.realisasi_total || row?.realisasi_hasil_pendataan || ""),
    persentase_prelist: cleanText(row?.persentase_prelist || row?.persentase_pendataan || ""),
    persentase_pendataan: cleanText(row?.persentase_pendataan || row?.persentase_prelist || ""),
    flag: cleanText(row?.flag || ""),
    foto: fotoBukti,
    foto_bukti: fotoBukti,
    jumlah_foto: fotoBukti.length,
    jumlah_foto_bukti: fotoBukti.length,
    foto_rows: fotoRows,
  };
}

async function createBappBlob(templateUrl, formValues, row, role, approveByPmlRows = []) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template BAPP: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    modules: [createFotoBuktiImageModule()],
  });
  const fotoRows = filterApproveByPmlRowsForBappRow(approveByPmlRows, row || {}, role);
  // renderAsync wajib dipakai karena foto diambil lewat fetch() (asinkron).
  await doc.renderAsync(buildBappTemplateData(formValues || {}, row || {}, role, fotoRows));
  return zipToDocxBlob(doc.getZip());
}

async function generateSingleBapp(templateUrl, formValues, row, role, approveByPmlRows = []) {
  const blob = await createBappBlob(templateUrl, formValues || {}, row || {}, role, approveByPmlRows);
  const safeName = sanitizeFileName(cleanText(row?.nama || `${role}-bapp`));
  saveAs(blob, `BAPP ${role} - ${safeName}.docx`);
}

const BAPP_ZIP_BATCH_SIZE = 150;

async function generateBapp(templateUrl, formValues, rows, role, onProgress, approveByPmlRows = []) {
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
      const blob = await createBappBlob(templateUrl, formValues || {}, row || {}, role, approveByPmlRows);
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
      ? `${namaPml}|${namaPpl}`
      : `${kec}|${desa}`;

    if (!map.has(keyBase)) {
      map.set(keyBase, {
        nama_pml: namaPml,
        nama_ppl: namaPpl,
        email_pengawas: cleanText(r.email_pengawas),
        email_pencacah: cleanText(r.email_pencacah),
        kdprov: cleanText(r.kdprov),
        kdkab: cleanText(r.kdkab),
        kdkec: kec,
        kddesa: desa,
        nmprov: cleanText(r.nmprov).toUpperCase(),
        nmkab: cleanText(r.nmkab).toUpperCase(),
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
  const [statusSlsData,      setStatusSlsData]      = useState([]);
  const [dataPerSlsData,     setDataPerSlsData]     = useState([]);
  const [approveByPmlData,   setApproveByPmlData]   = useState([]);
  const [xlsxLoaded,         setXlsxLoaded]         = useState(false);
  const [xlsxFileName,       setXlsxFileName]       = useState("data-petugas.xlsx");
  const [googleSheetUrl,     setGoogleSheetUrl]     = useState("https://docs.google.com/spreadsheets/d/10jA_NOMNn5pBuy1OPrSdHstscRrUOUlEDElk-jOmXLQ/edit?gid=1095810027#gid=1095810027");
  const [googleSheetApiKey,  setGoogleSheetApiKey]  = useState("");
  const [googleSheetError,   setGoogleSheetError]   = useState(null);
  const [googleSheetLoading, setGoogleSheetLoading] = useState(false);
  const [dataPmlProgressData, setDataPmlProgressData] = useState([]);
  // Lampiran preview controls: pilih jenis (PML/PPL) dan pilih identity (email/name) untuk generate satu-per-orang
  const [lampiranPreviewJenis, setLampiranPreviewJenis] = useState("PML");
  const [lampiranPreviewIdentity, setLampiranPreviewIdentity] = useState("__ALL__");

  const loadGoogleSheetData = async () => {
    const normalized = normalizeGoogleSheetUrl(googleSheetUrl);
    if (!normalized) { setGoogleSheetError("URL Google Sheets tidak valid."); return; }
    setGoogleSheetError(null);
    setGoogleSheetLoading(true);
    try {
      const {
        data,
        lampiran,
        bappData: loadedBappData,
        statusSls: loadedStatusSls,
        dataPerSls: loadedDataPerSls,
        approveByPml: loadedApproveByPml,
        dataPmlProgress: loadedDataPmlProgress,   // ⬅️ baru
      } = await loadGoogleSheet(normalized, googleSheetApiKey);
      const enrichedApproveByPml = await enrichApproveByPmlWithFotoBukti(loadedApproveByPml || []);
      if (
        data.length === 0 &&
        (!lampiran || lampiran.length === 0) &&
        (!loadedBappData || loadedBappData.length === 0) &&
        (!loadedStatusSls || loadedStatusSls.length === 0) &&
        (!loadedDataPerSls || loadedDataPerSls.length === 0) &&
        (!loadedApproveByPml || loadedApproveByPml.length === 0)
      ) {
        setGoogleSheetError("Tidak ada data petugas, Lampiran, Pembayaran, Status SLS, Data per SLS, maupun Approve by PML yang terbaca.");
        return;
      }

      setPetugasData(data || []);
      setLampiranData(lampiran || []);
      setBappData(loadedBappData || []);
      setStatusSlsData(loadedStatusSls || []);
      setDataPerSlsData(loadedDataPerSls || []);
      setApproveByPmlData(enrichedApproveByPml || []);
      setXlsxLoaded(true);
      setXlsxFileName(`${data.length} petugas, ${lampiran?.length || 0} lampiran, ${loadedBappData?.length || 0} pembayaran, ${loadedStatusSls?.length || 0} status SLS, ${loadedDataPerSls?.length || 0} data per SLS, ${loadedApproveByPml?.length || 0} approve PML`);
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
        const statusSlsRows = parseStatusSlsData(buffer);
        const dataPerSlsRows = parseDataPerSlsData(buffer);
        let approveByPmlRows = parseApproveByPmlData(buffer);
        approveByPmlRows = await enrichApproveByPmlWithFotoBukti(approveByPmlRows);
        const dataPmlProgressRows = parseDataPmlProgressData(buffer);
        setDataPmlProgressData(dataPmlProgressRows || []);
        setPetugasData(data || []);
        setLampiranData(lampiran || []);
        setBappData(bappRows || []);
        setStatusSlsData(statusSlsRows || []);
        setDataPerSlsData(dataPerSlsRows || []);
        setApproveByPmlData(approveByPmlRows || []);
        setXlsxLoaded(Boolean(
          lampiran?.length || data?.length || bappRows?.length ||
          statusSlsRows?.length || dataPerSlsRows?.length || approveByPmlRows?.length
        ));
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
    reader.onload = async (e) => {
      try {
        const buffer = e.target.result;
        const data = parseXlsxData(buffer);
        const lampiran = parseLampiranXlsxData(buffer);
        const bappRows = parseBappData(buffer);
        const statusSlsRows = parseStatusSlsData(buffer);
        const dataPerSlsRows = parseDataPerSlsData(buffer);
        let approveByPmlRows = parseApproveByPmlData(buffer);
        approveByPmlRows = await enrichApproveByPmlWithFotoBukti(approveByPmlRows);
        setPetugasData(data || []);
        setLampiranData(lampiran || []);
        setBappData(bappRows || []);
        setStatusSlsData(statusSlsRows || []);
        setDataPerSlsData(dataPerSlsRows || []);
        setApproveByPmlData(approveByPmlRows || []);
        setXlsxLoaded(Boolean(
          lampiran?.length || data?.length || bappRows?.length ||
          statusSlsRows?.length || dataPerSlsRows?.length || approveByPmlRows?.length
        ));
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
                <DocForm docType={selectedDoc} formData={formData} setFormData={setFormData} onPreview={(data) => { setPreviewData(data); setView("preview"); }} petugasData={petugasData} lampiranData={lampiranData} bappData={bappData} statusSlsData={statusSlsData} dataPerSlsData={dataPerSlsData} approveByPmlData={approveByPmlData} dataPmlProgressData={dataPmlProgressData} xlsxLoaded={xlsxLoaded} />
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
          <p className="font-black text-slate-900">Sedang Load Data, Harap Tunggu!</p>
        </div>
      </div>
      {/* <div className="mt-4 grid gap-3 sm:grid-cols-[1fr_auto]">
        <input value={url} onChange={(e) => onUrlChange(e.target.value)} placeholder="https://docs.google.com/spreadsheets/d/ID_SHEET/edit#gid=0"
          className="w-full rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm text-slate-800 outline-none transition focus:border-orange-300 focus:bg-white" />
        <button type="button" onClick={onLoad} disabled={loading || !url}
          className="rounded-2xl bg-orange-600 px-5 py-3 text-sm font-bold text-white transition enabled:hover:bg-orange-700 disabled:cursor-not-allowed disabled:bg-orange-200">
          {loading ? "Memuat..." : "Muat"}
        </button>
      </div> */}
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

// ─── UNGGAH EXCEL NAMA/EMAIL (fitur "Download Beberapa") ────────────────────
const SELECTION_TEMPLATE_URL = "/templates/Template Download Beberapa (Nama-Email).xlsx";

async function downloadSelectionTemplate() {
  const response = await fetch(SELECTION_TEMPLATE_URL);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const blob = await response.blob();
  saveAs(blob, "Template Download Beberapa (Nama-Email).xlsx");
}

function SelectionUploadPanel({ selectionRows, onSelectionLoaded, onClear, hint }) {
  const fileRef = useRef(null);
  const [error, setError] = useState("");
  const [templateDownloading, setTemplateDownloading] = useState(false);

  const handleFile = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const rows = parseSelectionXlsx(e.target.result);
        if (rows.length === 0) throw new Error("Tidak ada baris Nama/Email yang terbaca dari file.");
        setError("");
        onSelectionLoaded(rows);
      } catch (err) {
        setError(err.message || "Gagal membaca file Excel.");
      } finally {
        if (fileRef.current) fileRef.current.value = "";
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const handleDownloadTemplate = async () => {
    setTemplateDownloading(true);
    try {
      await downloadSelectionTemplate();
    } catch (err) {
      setError(err.message || "Gagal mengunduh template.");
    } finally {
      setTemplateDownloading(false);
    }
  };

  return (
    <div className="rounded-2xl border border-dashed border-orange-200 bg-orange-50/50 p-4 space-y-2">
      <div className="flex flex-wrap items-center gap-2">
        <button type="button" onClick={() => fileRef.current?.click()}
          className="inline-flex items-center gap-2 rounded-2xl border border-orange-200 bg-white px-4 py-2 text-sm font-bold text-orange-700 shadow-sm transition hover:bg-orange-50">
          <Upload size={14} /> Unggah Excel Nama/Email
        </button>
        <input ref={fileRef} type="file" accept=".xlsx" className="hidden" onChange={(e) => handleFile(e.target.files[0])} />
        {selectionRows.length > 0 && (
          <span className="inline-flex items-center gap-2 rounded-full border border-green-200 bg-green-50 px-3 py-1 text-xs font-bold text-green-700">
            <CheckCircle size={12} /> {selectionRows.length} baris dimuat
            <button type="button" onClick={onClear} className="ml-1 text-green-700 hover:text-green-900"><X size={12} /></button>
          </span>
        )}
      </div>
      {error && <p className="flex items-center gap-1 text-xs font-semibold text-red-500"><AlertCircle size={12} /> {error}</p>}
      <p className="text-xs font-semibold text-slate-400">
        {hint || "Kolom yang dibaca: Nama dan/atau Email. Baris yang tidak cocok dengan data akan diabaikan."}{" "}
        <button
          type="button"
          onClick={handleDownloadTemplate}
          disabled={templateDownloading}
          className="font-bold text-orange-600 underline underline-offset-2 hover:text-orange-700 disabled:opacity-50"
        >
          {templateDownloading ? "Mengunduh..." : "disini"}
        </button>
      </p>
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

function DocForm({ docType, formData, setFormData, onPreview, petugasData, lampiranData = [], bappData = [], statusSlsData = [], dataPerSlsData = [], dataPmlProgressData = [], approveByPmlData = [], xlsxLoaded }) {
  const update = (key, val) => setFormData((p) => ({ ...p, [key]: val }));

const [gabunganRole, setGabunganRole] = useState(""); // "PML" | "PPL"
const [gabunganManualSelect, setGabunganManualSelect] = useState("");
const [gabunganGenerating, setGabunganGenerating] = useState(false);
const [gabunganProgressText, setGabunganProgressText] = useState("");
const [gabunganSelectionRows, setGabunganSelectionRows] = useState([]);

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
  const [bappSelectionRows,       setBappSelectionRows]       = useState([]);
  const [bastRole,                setBastRole]                = useState("");
  const [bastManualSelect,        setBastManualSelect]        = useState("");
  const [bastGenerating,          setBastGenerating]          = useState(false);
  const [bastProgressText,        setBastProgressText]        = useState("");
  const [bastSelectionRows,       setBastSelectionRows]       = useState([]);
  const [suratKepalaSelectionRows, setSuratKepalaSelectionRows] = useState([]);
  const [suratKepalaGenerating, setSuratKepalaGenerating] = useState(false);
  const [suratKepalaProgressText, setSuratKepalaProgressText] = useState("");
  const [suratPenyelesaianLapanganSelect, setSuratPenyelesaianLapanganSelect] = useState("");
  const [suratPenyelesaianLapanganGenerating, setSuratPenyelesaianLapanganGenerating] = useState(false);
  const [suratPenyelesaianLapanganProgressText, setSuratPenyelesaianLapanganProgressText] = useState("");
  const [suratPenyelesaianLapanganSelectionRows, setSuratPenyelesaianLapanganSelectionRows] = useState([]);

  // Lampiran form controls: combined manual select for PML + PPL
  const [lampiranManualSelect, setLampiranManualSelect] = useState("");
  const [lampiranSelectionRows, setLampiranSelectionRows] = useState([]);
  const [lampiranBeberapaGenerating, setLampiranBeberapaGenerating] = useState(false);
  const [lampiranBeberapaProgressText, setLampiranBeberapaProgressText] = useState("");

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

    if (docType.id === "gabungan-pembayaran") {
      e.preventDefault();
      return;
    }

    if (docType.id === "surat-kepala") {
      e.preventDefault();
      return;
    }

    onPreview({ ...formData });
  };

  const renderFields = () => {
    switch (docType.id) {

      // ── ADMINISTRASI PEMBAYARAN LENGKAP PER PML/PPL ──────────────────────
      case "gabungan-pembayaran": {
        const nikLookup = React.useMemo(() => {
          const map = new Map();
          for (const p of petugasData || []) {
            const nama = upperText(p.nama);
            if (nama && !map.has(nama)) map.set(nama, cleanText(p.nik));
          }
          return map;
        }, [petugasData]);

        const berkasRecords = React.useMemo(() => {
          if (!gabunganRole) return [];
          return buildBerkasPembayaranRecords(
            bappData || [],
            lampiranData || [],
            gabunganRole,
            statusSlsData || [],
            dataPerSlsData || [],
            approveByPmlData || []
          );
        }, [bappData, lampiranData, gabunganRole, statusSlsData, dataPerSlsData, approveByPmlData]);

        const berkasOptions = React.useMemo(() => {
          return berkasRecords.map((record) => ({
            value: record.identity,
            label: `${record.displayName || "Tanpa Nama"}${record.email ? ` — ${record.email}` : ""}`,
            record,
          }));
        }, [berkasRecords]);

        const validateTanggal = () => {
          if (!formData.tanggal_surat) throw new Error("Isi tanggal surat terlebih dahulu.");
          const selectedDate = new Date(formData.tanggal_surat);
          const minDate = new Date("2026-07-15T00:00:00");
          const maxDate = new Date("2026-08-31T23:59:59");
          if (selectedDate < minDate || selectedDate > maxDate) {
            throw new Error("Tanggal surat hanya boleh 15 Juli 2026 sampai 31 Agustus 2026.");
          }
        };

        const getTemplateUrl = () => gabunganRole === "PML"
          ? BERKAS_PEMBAYARAN_PML_TEMPLATE_URL
          : BERKAS_PEMBAYARAN_PPL_TEMPLATE_URL;

        return (
          <div className="space-y-5">
            <div className="rounded-3xl border border-orange-100 bg-orange-50/70 p-5">
              <p className="text-xs font-black uppercase tracking-[0.2em] text-orange-700">Gabungan Administrasi Pembayaran</p>
              <p className="mt-2 text-sm font-semibold leading-6 text-slate-600">
                Pilih PML atau PPL, lalu generate satu berkas pembayaran lengkap per orang. Pilihan tersedia untuk satu nama, beberapa nama melalui Excel, atau semua nama.
              </p>
              <p className="mt-3 text-xs font-bold text-slate-500">
                Template: {gabunganRole === "PML" ? "BERKAS PEMBAYARAN PML.docx" : gabunganRole === "PPL" ? "BERKAS PEMBAYARAN PPL.docx" : "pilih role terlebih dahulu"}
              </p>
            </div>

            <div>
              <p className={labelCls}>Pilih Role</p>
              <div className="flex gap-3">
                <button type="button" onClick={() => {
                  setGabunganRole("PML");
                  setGabunganManualSelect("");
                  setGabunganSelectionRows([]);
                }} className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${gabunganRole === "PML" ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                  PML
                </button>
                <button type="button" onClick={() => {
                  setGabunganRole("PPL");
                  setGabunganManualSelect("");
                  setGabunganSelectionRows([]);
                }} className={`inline-flex items-center gap-2 rounded-2xl px-5 py-2.5 text-sm font-black transition ${gabunganRole === "PPL" ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}>
                  PPL
                </button>
              </div>
            </div>

            {gabunganRole && (
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
                  <p className="mt-1 text-xs font-semibold text-slate-400">Rentang tanggal yang diizinkan: 15–31 Juli 2026.</p>
                </div>

                <div>
                  <label className={labelCls}>Pilih Sendiri</label>
                  <select value={gabunganManualSelect} onChange={(e) => setGabunganManualSelect(e.target.value)} className={inputCls}>
                    <option value="">— Pilih Nama {gabunganRole} —</option>
                    {berkasOptions.map((option) => (
                      <option key={option.value} value={option.value}>{option.label}</option>
                    ))}
                  </select>
                  <p className="mt-1 text-xs font-semibold text-slate-400">
                    {berkasRecords.length} orang terdeteksi. Filter Lampiran memakai {statusSlsData.length} baris Status SLS. Tabel beban kerja memakai {dataPerSlsData.length} baris Data per SLS. Jumlah PML memakai {approveByPmlData.length} baris Approve by PML.
                  </p>
                </div>

                <SelectionUploadPanel
                  selectionRows={gabunganSelectionRows}
                  onSelectionLoaded={setGabunganSelectionRows}
                  onClear={() => setGabunganSelectionRows([])}
                  hint={`Unggah Excel berisi kolom Nama dan/atau Email untuk memilih beberapa ${gabunganRole}.`}
                />

                {gabunganGenerating && (
                  <div className="flex items-center gap-3 rounded-2xl border border-orange-200 bg-orange-50 px-4 py-3 text-sm font-semibold text-orange-700">
                    <LoaderCircle size={18} className="animate-spin" />
                    <span>{gabunganProgressText || `Sedang membuat berkas pembayaran ${gabunganRole}...`}</span>
                  </div>
                )}

                <div className="grid gap-4 sm:grid-cols-3">
                  <button type="button" onClick={async () => {
                    try {
                      validateTanggal();
                      if (!gabunganManualSelect) throw new Error("Pilih nama terlebih dahulu.");
                      const chosen = berkasOptions.find((option) => option.value === gabunganManualSelect)?.record;
                      if (!chosen) throw new Error("Data nama yang dipilih tidak ditemukan.");
                      setGabunganGenerating(true);
                      setGabunganProgressText("Membuat satu berkas pembayaran...");
                      await generateSingleBerkasPembayaran(
                        getTemplateUrl(), formData, chosen, gabunganRole, nikLookup
                      );
                    } catch (err) { alert(err.message || err); }
                    finally { setGabunganGenerating(false); setGabunganProgressText(""); }
                  }} disabled={!gabunganManualSelect || berkasRecords.length === 0 || gabunganGenerating}
                    className="inline-flex items-center justify-center rounded-2xl bg-orange-500 px-5 py-3 text-sm font-black text-white shadow transition hover:bg-orange-600 disabled:cursor-not-allowed disabled:bg-orange-200">
                    Generate Terpilih
                  </button>

                  <button type="button" onClick={async () => {
                    try {
                      validateTanggal();
                      if (gabunganSelectionRows.length === 0) throw new Error("Unggah file Excel Email terlebih dahulu.");
                      const keySet = buildEmailOnlySelectionKeySet(gabunganSelectionRows);
                      const matchedRecords = berkasRecords.filter((record) =>
                        rowMatchesSelectionByEmail(keySet, record.email)
                      );
                      if (matchedRecords.length === 0) throw new Error("Tidak ada data yang cocok dengan file Excel.");
                      setGabunganGenerating(true);
                      setGabunganProgressText(`Menyiapkan ${matchedRecords.length} berkas pembayaran...`);
                      await generateBerkasPembayaran(
                        getTemplateUrl(), formData, matchedRecords, gabunganRole, nikLookup,
                        ({ batchIndex, totalBatches }) => setGabunganProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`)
                      );
                    } catch (err) { alert(err.message || err); }
                    finally { setGabunganGenerating(false); setGabunganProgressText(""); }
                  }} disabled={berkasRecords.length === 0 || gabunganGenerating || gabunganSelectionRows.length === 0}
                    className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                    Generate Beberapa
                  </button>

                  <button type="button" onClick={async () => {
                    try {
                      validateTanggal();
                      if (berkasRecords.length === 0) throw new Error(`Tidak ada data ${gabunganRole}.`);
                      setGabunganGenerating(true);
                      setGabunganProgressText("Mempersiapkan semua berkas pembayaran...");
                      await generateBerkasPembayaran(
                        getTemplateUrl(), formData, berkasRecords, gabunganRole, nikLookup,
                        ({ batchIndex, totalBatches }) => setGabunganProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`)
                      );
                    } catch (err) { alert(err.message || err); }
                    finally { setGabunganGenerating(false); setGabunganProgressText(""); }
                  }} disabled={berkasRecords.length === 0 || gabunganGenerating}
                    className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                    Generate Semua
                  </button>
                </div>
              </>
            )}
          </div>
        );
      }

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

                <SelectionUploadPanel
                  selectionRows={bappSelectionRows}
                  onSelectionLoaded={setBappSelectionRows}
                  onClear={() => setBappSelectionRows([])}
                  hint="Pastikan file dan isian file benar. Download template"
                />

                {bappGenerating && (
                  <div className="flex items-center gap-3 rounded-2xl border border-orange-200 bg-orange-50 px-4 py-3 text-sm font-semibold text-orange-700">
                    <LoaderCircle size={18} className="animate-spin" />
                    <span>{bappProgressText || "Sedang menyiapkan file BAPP..."}</span>
                  </div>
                )}

                <div className="grid gap-4 sm:grid-cols-3">
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
                      await generateSingleBapp(bappRole === "PML" ? BAPP_PML_TEMPLATE_URL : BAPP_PPL_TEMPLATE_URL, formData, chosenRow, bappRole, approveByPmlData);
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
                      if (bappSelectionRows.length === 0) throw new Error("Unggah file Excel Nama/Email terlebih dahulu.");
                      const keySet = buildSelectionKeySet(bappSelectionRows);
                      const matchedRows = filteredBappRows.filter((row) => rowMatchesSelection(keySet, row.nama, row.email));
                      if (matchedRows.length === 0) throw new Error("Tidak ada data yang cocok dengan file yang diunggah.");
                      setBappGenerating(true);
                      setBappProgressText(`Menyiapkan ${matchedRows.length} dokumen terpilih...`);
                      await generateBapp(bappRole === "PML" ? BAPP_PML_TEMPLATE_URL : BAPP_PPL_TEMPLATE_URL, formData, matchedRows, bappRole, ({ batchIndex, totalBatches }) => {
                        setBappProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`);
                      }, approveByPmlData);
                    } catch (err) { alert(err.message || err); }
                    finally { setBappGenerating(false); setBappProgressText(""); }
                  }} disabled={filteredBappRows.length === 0 || bappGenerating || bappSelectionRows.length === 0} className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                    Download Beberapa
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
                      await generateBapp(bappRole === "PML" ? BAPP_PML_TEMPLATE_URL : BAPP_PPL_TEMPLATE_URL, formData, matchedRows, bappRole, ({ batchIndex, totalBatches }) => {
                        setBappProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`);
                      }, approveByPmlData);
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

                <SelectionUploadPanel
                  selectionRows={bastSelectionRows}
                  onSelectionLoaded={setBastSelectionRows}
                  onClear={() => setBastSelectionRows([])}
                  hint={`Kolom Nama dan/atau Email dicocokkan dengan nama ${isPml ? "Pengawas & Email Pengawas" : "Pencacah & Email Pencacah"}.`}
                />

                {bastGenerating && (
                  <div className="flex items-center gap-3 rounded-2xl border border-orange-200 bg-orange-50 px-4 py-3 text-sm font-semibold text-orange-700">
                    <LoaderCircle size={18} className="animate-spin" />
                    <span>{bastProgressText || "Sedang menyiapkan file BAST..."}</span>
                  </div>
                )}

                <div className="grid gap-4 sm:grid-cols-3">
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
                      if (bastSelectionRows.length === 0) throw new Error("Unggah file Excel Nama/Email terlebih dahulu.");
                      const keySet = buildSelectionKeySet(bastSelectionRows);
                      const matchedRows = filteredBastRows.filter((row) => rowMatchesSelection(keySet, isPml ? row.nama_pml : row.nama_ppl, isPml ? row.email_pengawas : row.email_pencacah));
                      if (matchedRows.length === 0) throw new Error("Tidak ada data yang cocok dengan file yang diunggah.");
                      setBastGenerating(true);
                      setBastProgressText(`Menyiapkan dokumen terpilih dari file...`);
                      await generateBast(bastRole === "PML" ? BAST_PML_TEMPLATE_URL : BAST_PPL_TEMPLATE_URL, formData, matchedRows, bastRole, nikLookup, ({ batchIndex, totalBatches }) => {
                        setBastProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`);
                      });
                    } catch (err) { alert(err.message || err); }
                    finally { setBastGenerating(false); setBastProgressText(""); }
                  }} disabled={filteredBastRows.length === 0 || bastGenerating || bastSelectionRows.length === 0} className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                    Download Beberapa
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

            <SelectionUploadPanel
              selectionRows={suratPenyelesaianLapanganSelectionRows}
              onSelectionLoaded={setSuratPenyelesaianLapanganSelectionRows}
              onClear={() => setSuratPenyelesaianLapanganSelectionRows([])}
              hint="Kolom Nama dan/atau Email dicocokkan dengan data PML pada sheet Pembayaran."
            />

            {suratPenyelesaianLapanganGenerating && (
              <div className="flex items-center gap-3 rounded-2xl border border-orange-200 bg-orange-50 px-4 py-3 text-sm font-semibold text-orange-700">
                <LoaderCircle size={18} className="animate-spin" />
                <span>{suratPenyelesaianLapanganProgressText || "Sedang menyiapkan dokumen..."}</span>
              </div>
            )}

            <div className="grid gap-4 sm:grid-cols-3">
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
                  if (suratPenyelesaianLapanganSelectionRows.length === 0) throw new Error("Unggah file Excel Nama/Email terlebih dahulu.");
                  const keySet = buildSelectionKeySet(suratPenyelesaianLapanganSelectionRows);
                  const matchedRows = filteredSuratPenyelesaianLapanganRows.filter((row) => rowMatchesSelection(keySet, row.nama, row.email));
                  if (matchedRows.length === 0) throw new Error("Tidak ada data yang cocok dengan file yang diunggah.");
                  setSuratPenyelesaianLapanganGenerating(true);
                  setSuratPenyelesaianLapanganProgressText(`Menyiapkan ${matchedRows.length} dokumen terpilih...`);
                  await generateSuratPernyataanPenyelesaianLapangan(SURAT_PERNYATAAN_PENYELESAIAN_LAPANGAN_TEMPLATE_URL, matchedRows, ({ batchIndex, totalBatches }) => {
                    setSuratPenyelesaianLapanganProgressText(`Membuat batch ${batchIndex} dari ${totalBatches}...`);
                  });
                } catch (err) { alert(err.message || err); }
                finally { setSuratPenyelesaianLapanganGenerating(false); setSuratPenyelesaianLapanganProgressText(""); }
              }} disabled={filteredSuratPenyelesaianLapanganRows.length === 0 || suratPenyelesaianLapanganGenerating || suratPenyelesaianLapanganSelectionRows.length === 0} className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                Download Beberapa
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

            {/* Download Beberapa: unggah Excel Nama/Email lalu generate untuk PML dan/atau PPL sekaligus */}
            <div className="rounded-2xl border border-orange-100 bg-white p-4 space-y-3">
              <p className="text-xs font-black uppercase tracking-[0.2em] text-orange-700">Download Beberapa (dari file Excel Nama/Email)</p>
              <SelectionUploadPanel
                selectionRows={lampiranSelectionRows}
                onSelectionLoaded={setLampiranSelectionRows}
                onClear={() => setLampiranSelectionRows([])}
                hint="Silahkan upload file excel sesuai template. Download template disini"
              />

              {lampiranBeberapaGenerating && (
                <div className="flex items-center gap-3 rounded-2xl border border-orange-200 bg-orange-50 px-4 py-3 text-sm font-semibold text-orange-700">
                  <LoaderCircle size={18} className="animate-spin" />
                  <span>{lampiranBeberapaProgressText || "Sedang menyiapkan dokumen..."}</span>
                </div>
              )}

              <div className="grid gap-4 sm:grid-cols-2">
                <button type="button" disabled={lampiranSelectionRows.length === 0 || lampiranBeberapaGenerating} onClick={async () => {
                  try {
                    if (!xlsxLoaded) throw new Error("Data XLSX/Google Sheet belum dimuat.");
                    if (lampiranSelectionRows.length === 0) throw new Error("Unggah file Excel Nama/Email terlebih dahulu.");
                    const keySet = buildSelectionKeySet(lampiranSelectionRows);
                    const matchedRows = rows.filter((r) => rowMatchesSelection(keySet, r.nama_pml, r.email_pengawas));
                    if (matchedRows.length === 0) throw new Error("Tidak ada PML yang cocok dengan file yang diunggah.");
                    setLampiranBeberapaGenerating(true);
                    setLampiranBeberapaProgressText("Menyiapkan Lampiran PML terpilih...");
                    await generateLampiran(LAMPIRAN_PML_TEMPLATE_URL, {}, matchedRows, "PML");
                  } catch (err) { alert(err.message || err); }
                  finally { setLampiranBeberapaGenerating(false); setLampiranBeberapaProgressText(""); }
                }} className="inline-flex items-center justify-center rounded-2xl bg-orange-500 px-5 py-3 text-sm font-black text-white shadow transition hover:bg-orange-600 disabled:cursor-not-allowed disabled:bg-orange-200">
                  Download Beberapa PML
                </button>

                <button type="button" disabled={lampiranSelectionRows.length === 0 || lampiranBeberapaGenerating} onClick={async () => {
                  try {
                    if (!xlsxLoaded) throw new Error("Data XLSX/Google Sheet belum dimuat.");
                    if (lampiranSelectionRows.length === 0) throw new Error("Unggah file Excel Nama/Email terlebih dahulu.");
                    const keySet = buildSelectionKeySet(lampiranSelectionRows);
                    const matchedRows = rows.filter((r) => rowMatchesSelection(keySet, r.nama_ppl, r.email_pencacah));
                    if (matchedRows.length === 0) throw new Error("Tidak ada PPL yang cocok dengan file yang diunggah.");
                    setLampiranBeberapaGenerating(true);
                    setLampiranBeberapaProgressText("Menyiapkan Lampiran PPL terpilih...");
                    await generateLampiran(LAMPIRAN_PPL_TEMPLATE_URL, {}, matchedRows, "PPL");
                  } catch (err) { alert(err.message || err); }
                  finally { setLampiranBeberapaGenerating(false); setLampiranBeberapaProgressText(""); }
                }} className="inline-flex items-center justify-center rounded-2xl border border-orange-200 bg-white px-5 py-3 text-sm font-black text-orange-700 shadow transition hover:bg-orange-50 disabled:cursor-not-allowed disabled:opacity-60">
                  Download Beberapa PPL
                </button>
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

      case "surat-kepala": {
        return (
          <div className="space-y-5">
            <div className="rounded-3xl border border-orange-100 bg-orange-50/70 p-5">
              <p className="text-xs font-black uppercase tracking-[0.2em] text-orange-700">
                Surat Pernyataan Evaluasi Pelaksanaan Lapangan (Kepala BPS)
              </p>
              <p className="mt-2 text-sm font-semibold leading-6 text-slate-600">
                Unggah Excel berisi kolom Nama dan/atau Email untuk mengisi tabel lampiran.
                Satu file akan berisi semua nama yang cocok: PML tampil lebih dulu (A-Z),
                lalu PPL (A-Z). Target Prelist &amp; Jabatan diambil dari sheet Pembayaran;
                Realisasi PML dari kolom Realisasi Total, Realisasi PPL dijumlah dari
                sheet Data per SLS.
              </p>
            </div>
      
            <SelectionUploadPanel
              selectionRows={suratKepalaSelectionRows}
              onSelectionLoaded={setSuratKepalaSelectionRows}
              onClear={() => setSuratKepalaSelectionRows([])}
              hint="Kolom Nama dan/atau Email dicocokkan ke sheet Pembayaran."
            />
      
            {suratKepalaGenerating && (
              <div className="flex items-center gap-3 rounded-2xl border border-orange-200 bg-orange-50 px-4 py-3 text-sm font-semibold text-orange-700">
                <LoaderCircle size={18} className="animate-spin" />
                <span>{suratKepalaProgressText || "Sedang membuat surat..."}</span>
              </div>
            )}
      
            <button
              type="button"
              onClick={async () => {
                try {
                  if (suratKepalaSelectionRows.length === 0) {
                    throw new Error("Unggah file Excel Nama/Email terlebih dahulu.");
                  }
                  const { rows, skipped } = buildSuratKepalaRows(
                    suratKepalaSelectionRows,
                    bappData,
                    dataPerSlsData,
                    lampiranData,
                    statusSlsData,
                    approveByPmlData,
                    dataPmlProgressData
                  );
                  if (rows.length === 0) {
                    throw new Error("Tidak ada baris yang cocok untuk dimasukkan ke lampiran.");
                  }
                  setSuratKepalaGenerating(true);
                  setSuratKepalaProgressText(`Menyiapkan surat untuk ${rows.length} petugas...`);
                  await generateSuratKepala(SURAT_KEPALA_TEMPLATE_URL, rows);
                  if (skipped.length > 0) {
                    alert(
                      `Surat berhasil dibuat. ${skipped.length} baris dilewati:\n\n` +
                      skipped.map((s) => `- ${s.nama || s.email}: ${s.alasan}`).join("\n")
                    );
                  }
                } catch (err) {
                  alert(err.message || err);
                } finally {
                  setSuratKepalaGenerating(false);
                  setSuratKepalaProgressText("");
                }
              }}
              disabled={suratKepalaSelectionRows.length === 0 || suratKepalaGenerating}
              className="inline-flex items-center justify-center gap-2 rounded-2xl bg-orange-500 px-5 py-3 text-sm font-black text-white shadow transition hover:bg-orange-600 disabled:cursor-not-allowed disabled:bg-orange-200"
            >
              <Download size={16} /> Generate Surat Kepala
            </button>
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
      {docType.id !== "lampiran" && docType.id !== "bapp" && docType.id !== "bast" && docType.id !== "surat-pernyataan-penyelesaian-lapangan" && docType.id !== "gabungan-pembayaran" && (
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

// ─── BERKAS PEMBAYARAN PML/PPL ───────────────────────────────────────────────
// Satu record menyatukan identitas dari sheet Pembayaran dengan semua baris
// wilayah/SLS milik orang yang sama dari sheet Lampiran.
function getBerkasIdentity(nama, email) {
  const emailClean = cleanText(email);
  if (emailClean) return `EMAIL::${upperText(emailClean)}`;
  return `NAME::${upperText(nama)}`;
}

function buildBerkasPembayaranRecords(
  bappRows = [],
  lampiranRows = [],
  role = "PML",
  statusSlsRows = [],
  dataPerSlsRows = [],
  approveByPmlRows = []
) {
  const isPml = upperText(role) === "PML";
  const records = [];
  const byEmail = new Map();
  const byName = new Map();
  const ambiguousNames = new Set();

  const registerName = (record) => {
    const nameKey = upperText(record.displayName);
    if (!nameKey) return;
    if (byName.has(nameKey) && byName.get(nameKey) !== record) {
      ambiguousNames.add(nameKey);
      byName.delete(nameKey);
      return;
    }
    if (!ambiguousNames.has(nameKey)) byName.set(nameKey, record);
  };

  const ensureRecord = (displayName, email, bappRow = null) => {
    const identity = getBerkasIdentity(displayName, email);
    let record = records.find((item) => item.identity === identity);
    if (record) {
      if (!record.bappRow && bappRow) record.bappRow = bappRow;
      return record;
    }

    record = {
      identity,
      displayName: displayName || email || "Tanpa Nama",
      email,
      bappRow,
      lampiranRows: [],
      // Baris sheet Pembayaran yang terkait dengan record ini. Untuk PML,
      // isinya adalah baris-baris PPL di bawah pengawas tersebut. Untuk PPL,
      // isinya adalah baris Pembayaran milik PPL itu sendiri.
      pembayaranRows: [],
      // Baris dari sheet Status SLS milik PML/PPL ini. Dipakai untuk menyaring
      // baris Lampiran sebelum tabel gabungan dibentuk.
      statusSlsRows: [],
      // Baris sheet Data per SLS untuk tabel beban kerja pada halaman terakhir.
      dataPerSlsRows: [],
      // Baris sheet Approve by PML. Khusus dokumen PML, nilai ini menjadi sumber
      // realisasi/jumlah pemeriksaan, baik per SLS, per PPL, maupun total PML.
      approveByPmlRows: [],
      usernameSobat: cleanText(bappRow?.username_sobat || bappRow?.sobat_id || ""),
    };
    records.push(record);
    if (email) byEmail.set(upperText(email), record);
    registerName(record);
    return record;
  };

  // Record utama tetap dibuat berdasarkan role yang dipilih.
  const roleBappRows = (bappRows || []).filter((row) => isBappRowForRole(row, role));
  for (const row of roleBappRows) {
    const displayName = cleanText(row?.nama);
    const email = cleanText(row?.email);
    if (!displayName && !email) continue;
    ensureRecord(displayName, email, row);
  }

  // Data Lampiran masih dipertahankan untuk bagian BAST/lampiran dokumen gabungan.
  for (const row of lampiranRows || []) {
    const displayName = cleanText(isPml ? row?.nama_pml : row?.nama_ppl);
    const email = cleanText(isPml ? row?.email_pengawas : row?.email_pencacah);
    if (!displayName && !email) continue;

    const emailKey = upperText(email);
    const nameKey = upperText(displayName);
    let record = emailKey ? byEmail.get(emailKey) : null;
    if (!record && nameKey && !ambiguousNames.has(nameKey)) record = byName.get(nameKey);
    if (!record) record = ensureRecord(displayName, email, null);

    record.lampiranRows.push(row);
    if (!record.email && email) record.email = email;
    if ((!record.displayName || record.displayName === "Tanpa Nama") && displayName) record.displayName = displayName;
  }

  // Hubungkan metrik progres dari sheet Pembayaran ke record.
  //
  // FIX v5: jangan bergantung hanya pada kolom Email Pengawas/Nama Pengawas di
  // sheet Pembayaran. Pada banyak file, kolom relasi itu kosong atau berisi #N/A,
  // sehingga pembayaranRows menjadi kosong dan tag {prelist_total}, {realisasi},
  // {persentase}, serta {average_persentase} ikut kosong.
  //
  // Sumber relasi yang paling stabil adalah daftar PPL yang SUDAH menempel pada
  // masing-masing PML di sheet Lampiran. Jadi setiap PPL di Lampiran dicari langsung
  // ke sheet Pembayaran lewat Email Pencacah, lalu fallback ke Nama Pencacah.
  const pplPaymentRows = (bappRows || []).filter((row) => isBappRowForRole(row, "PPL"));
  const paymentByPplEmail = new Map();
  const paymentByPplName = new Map();
  const ambiguousPaymentNames = new Set();

  for (const row of pplPaymentRows) {
    const pplEmail = upperText(row?.email_ppl || row?.email || "");
    const pplName = upperText(row?.nama_ppl || row?.nama || "");
    if (pplEmail && !paymentByPplEmail.has(pplEmail)) paymentByPplEmail.set(pplEmail, row);
    if (pplName) {
      if (paymentByPplName.has(pplName) && paymentByPplName.get(pplName) !== row) {
        ambiguousPaymentNames.add(pplName);
        paymentByPplName.delete(pplName);
      } else if (!ambiguousPaymentNames.has(pplName)) {
        paymentByPplName.set(pplName, row);
      }
    }
  }

  const addPaymentRow = (record, row) => {
    if (!record || !row) return;
    const nama = cleanText(row?.nama_ppl || row?.nama);
    const email = cleanText(row?.email_ppl || row?.email);
    if (!nama && !email) return;
    const paymentIdentity = getBerkasIdentity(nama, email);
    const alreadyAdded = record.pembayaranRows.some((item) =>
      getBerkasIdentity(
        cleanText(item?.nama_ppl || item?.nama),
        cleanText(item?.email_ppl || item?.email)
      ) === paymentIdentity
    );
    if (!alreadyAdded) record.pembayaranRows.push(row);
  };

  // Jalur utama: Lampiran -> identitas PPL -> baris Pembayaran.
  for (const record of records) {
    if (!isPml) {
      addPaymentRow(record, record?.bappRow);
      continue;
    }

    for (const lampiranRow of record.lampiranRows || []) {
      const pplEmail = upperText(lampiranRow?.email_pencacah || lampiranRow?.email_ppl || "");
      const pplName = upperText(lampiranRow?.nama_ppl || "");
      let paymentRow = pplEmail ? paymentByPplEmail.get(pplEmail) : null;
      if (!paymentRow && pplName && !ambiguousPaymentNames.has(pplName)) {
        paymentRow = paymentByPplName.get(pplName) || null;
      }
      addPaymentRow(record, paymentRow);
    }
  }

  // Jalur cadangan: tetap manfaatkan kolom relasi Pengawas bila tersedia.
  // Ini berguna untuk PPL yang ada di Pembayaran tetapi belum tercantum di Lampiran.
  for (const row of pplPaymentRows) {
    const ownerName = cleanText(isPml ? (row?.nama_pengawas || row?.nama_pml) : row?.nama);
    const ownerEmail = cleanText(isPml ? (row?.email_pengawas || row?.email_pml) : row?.email);
    const emailKey = upperText(ownerEmail);
    const nameKey = upperText(ownerName);

    let record = emailKey ? byEmail.get(emailKey) : null;
    if (!record && nameKey && !ambiguousNames.has(nameKey)) record = byName.get(nameKey);
    if (!record) continue;
    addPaymentRow(record, row);
  }

  // Kaitkan sheet Status SLS ke pemilik dokumen. Untuk role PML, pemiliknya
  // ditentukan dari Nama/Email PML; untuk role PPL dari Nama/Email PPL.
  for (const statusRow of statusSlsRows || []) {
    const ownerName = cleanText(isPml ? statusRow?.nama_pml : statusRow?.nama_ppl);
    const ownerEmail = cleanText(isPml ? statusRow?.email_pml : statusRow?.email_ppl);
    const emailKey = upperText(ownerEmail);
    const nameKey = upperText(ownerName);

    let record = emailKey ? byEmail.get(emailKey) : null;
    if (!record && nameKey && !ambiguousNames.has(nameKey)) record = byName.get(nameKey);

    // Fallback paling kuat untuk PML: cari record yang Lampirannya memuat PPL
    // yang sama dengan baris Status SLS. Ini menolong ketika Nama/Email PML pada
    // sheet Status SLS kosong atau penulisannya berbeda.
    if (!record) {
      const statusPplEmail = upperText(statusRow?.email_ppl || "");
      const statusPplName = upperText(statusRow?.nama_ppl || "");
      record = records.find((candidate) =>
        (candidate?.lampiranRows || []).some((lampiranRow) => {
          const lampiranPplEmail = upperText(lampiranRow?.email_pencacah || lampiranRow?.email_ppl || "");
          const lampiranPplName = upperText(lampiranRow?.nama_ppl || "");
          return (statusPplEmail && lampiranPplEmail === statusPplEmail) ||
            (statusPplName && lampiranPplName === statusPplName);
        })
      ) || null;
    }

    if (!record) continue;
    record.statusSlsRows.push(statusRow);
  }

  // Kaitkan sheet Data per SLS ke dokumen PML/PPL. Nama menjadi kunci utama.
  // Username Sobat dipakai untuk membantu pencocokan bila nama tidak unik.
  const normalizePersonKey = (value) => upperText(value).replace(/\s+/g, " ");
  const dataPerSlsRowKey = (row) => [
    cleanText(row?.no_sumber),
    normalizePersonKey(row?.nama_pml),
    normalizePersonKey(row?.username_pml),
    normalizePersonKey(row?.nama_ppl),
    normalizePersonKey(row?.username_ppl),
    cleanText(row?.kdkec),
    cleanText(row?.kddesa),
    cleanText(row?.kode_sls),
  ].join("|");

  for (const dataRow of dataPerSlsRows || []) {
    const ownerName = cleanText(isPml ? dataRow?.nama_pml : dataRow?.nama_ppl);
    const ownerUsername = cleanText(isPml ? dataRow?.username_pml : dataRow?.username_ppl);
    const nameKey = normalizePersonKey(ownerName);
    const usernameKey = normalizePersonKey(ownerUsername);

    let record = null;

    if (usernameKey) {
      const usernameMatches = records.filter((candidate) => {
        const candidateUsername = normalizePersonKey(
          candidate?.usernameSobat ||
          candidate?.bappRow?.username_sobat ||
          candidate?.bappRow?.sobat_id ||
          ""
        );
        const candidateEmail = normalizePersonKey(candidate?.email || candidate?.bappRow?.email || "");
        const candidateEmailLocal = candidateEmail.includes("@") ? candidateEmail.split("@")[0] : candidateEmail;
        return candidateUsername === usernameKey ||
          candidateEmail === usernameKey ||
          candidateEmailLocal === usernameKey;
      });
      if (usernameMatches.length === 1) record = usernameMatches[0];
    }

    if (!record && nameKey) {
      const nameMatches = records.filter(
        (candidate) => normalizePersonKey(candidate?.displayName || "") === nameKey
      );
      if (nameMatches.length === 1) record = nameMatches[0];
    }

    if (!record && nameKey && !ambiguousNames.has(nameKey)) {
      record = byName.get(nameKey) || null;
    }

    // Data per SLS tetap dapat membentuk pilihan dokumen ketika orang tersebut
    // belum muncul pada sheet Pembayaran atau Lampiran.
    if (!record && ownerName) record = ensureRecord(ownerName, "", null);
    if (!record) continue;

    if (!record.usernameSobat && ownerUsername) record.usernameSobat = ownerUsername;
    const rowKey = dataPerSlsRowKey(dataRow);
    const alreadyAdded = record.dataPerSlsRows.some((item) => dataPerSlsRowKey(item) === rowKey);
    if (!alreadyAdded) record.dataPerSlsRows.push(dataRow);
  }

  // Kaitkan sheet Approve by PML ke pemilik dokumen. Untuk role PML, relasi
  // utama memakai Email PML lalu Nama PML. Untuk role PPL data ini tidak mengubah
  // hasil, tetapi tetap dapat ditempel untuk kebutuhan diagnostik.
  const approveRowKey = (row) => [
    upperText(row?.email_pml || ""),
    upperText(row?.nama_pml || ""),
    upperText(row?.email_ppl || ""),
    upperText(row?.nama_ppl || ""),
    normalizeStatusSlsCode(row?.kode_sls || "", 6),
    cleanText(row?.jumlah_approve_pml || ""),
    cleanText(row?.waktu_submit || ""),
  ].join("|");

  for (const approveRow of approveByPmlRows || []) {
    const ownerName = cleanText(isPml ? approveRow?.nama_pml : approveRow?.nama_ppl);
    const ownerEmail = cleanText(isPml ? approveRow?.email_pml : approveRow?.email_ppl);
    const emailKey = upperText(ownerEmail);
    const nameKey = upperText(ownerName);

    let record = emailKey ? byEmail.get(emailKey) : null;
    if (!record && nameKey && !ambiguousNames.has(nameKey)) record = byName.get(nameKey);

    if (!record && isPml) {
      const approvePplEmail = upperText(approveRow?.email_ppl || "");
      const approvePplName = upperText(approveRow?.nama_ppl || "");
      record = records.find((candidate) =>
        (candidate?.lampiranRows || []).some((lampiranRow) => {
          const lampiranPplEmail = upperText(lampiranRow?.email_pencacah || lampiranRow?.email_ppl || "");
          const lampiranPplName = upperText(lampiranRow?.nama_ppl || "");
          return (approvePplEmail && approvePplEmail === lampiranPplEmail) ||
            (approvePplName && approvePplName === lampiranPplName);
        })
      ) || null;
    }

    if (!record && ownerName) record = ensureRecord(ownerName, ownerEmail, null);
    if (!record) continue;

    const rowKey = approveRowKey(approveRow);
    const alreadyAdded = record.approveByPmlRows.some((item) => approveRowKey(item) === rowKey);
    if (!alreadyAdded) record.approveByPmlRows.push(approveRow);
  }

  return records.sort((a, b) =>
    cleanText(a.displayName).localeCompare(cleanText(b.displayName), "id-ID", { sensitivity: "base" })
  );
}

function pembayaranPersonIdentity(row = {}) {
  return getBerkasIdentity(
    cleanText(row?.nama_ppl || row?.nama),
    cleanText(row?.email_ppl || row?.email)
  );
}

function findPembayaranRowForLampiran(groupedLampiranRow = {}, pembayaranRows = []) {
  const emailPpl = upperText(groupedLampiranRow?.email_pencacah || groupedLampiranRow?.email_ppl || "");
  const namaPpl = upperText(groupedLampiranRow?.nama_ppl || "");

  if (emailPpl) {
    const byEmail = (pembayaranRows || []).find((row) =>
      upperText(row?.email_ppl || row?.email || "") === emailPpl
    );
    if (byEmail) return byEmail;
  }

  if (namaPpl) {
    return (pembayaranRows || []).find((row) =>
      upperText(row?.nama_ppl || row?.nama || "") === namaPpl
    ) || null;
  }

  return null;
}

function parsePercentageNumber(value) {
  let text = cleanText(value).replace(/%/g, "").replace(/\s+/g, "");
  if (!text || /^#(?:N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NUM!|NULL!)$/i.test(text)) return null;

  // Toleran terhadap format Indonesia (47,69) maupun format spreadsheet (47.69).
  if (text.includes(",") && text.includes(".")) {
    if (text.lastIndexOf(",") > text.lastIndexOf(".")) {
      text = text.replace(/\./g, "").replace(",", ".");
    } else {
      text = text.replace(/,/g, "");
    }
  } else if (text.includes(",")) {
    text = text.replace(",", ".");
  }

  const parsed = Number(text.replace(/[^0-9.+-]/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}

function formatPercentageNumber(value, maximumFractionDigits = 2) {
  if (!Number.isFinite(value)) return "";
  return value.toLocaleString("id-ID", {
    minimumFractionDigits: maximumFractionDigits,
    maximumFractionDigits,
  });
}

function calculateAveragePersentase(pembayaranRows = []) {
  const uniqueRows = new Map();
  for (const row of pembayaranRows || []) {
    const key = pembayaranPersonIdentity(row);
    if (!uniqueRows.has(key)) uniqueRows.set(key, row);
  }

  const values = [...uniqueRows.values()]
    .map((row) => parsePercentageNumber(row?.persentase_pendataan || row?.persentase_prelist))
    .filter((value) => value != null);

  if (values.length === 0) return { raw: null, formatted: "", count: 0 };
  const raw = values.reduce((sum, value) => sum + value, 0) / values.length;
  return { raw, formatted: formatPercentageNumber(raw, 2), count: values.length };
}

function parseDataPerSlsNumber(value) {
  let text = cleanText(value).replace(/%/g, "").replace(/\s+/g, "");
  if (!text || /^#(?:N\/A|VALUE!|REF!|DIV\/0!|NAME\?|NUM!|NULL!)$/i.test(text)) return null;

  if (text.includes(",") && text.includes(".")) {
    if (text.lastIndexOf(",") > text.lastIndexOf(".")) {
      text = text.replace(/\./g, "").replace(",", ".");
    } else {
      text = text.replace(/,/g, "");
    }
  } else if (text.includes(",")) {
    text = text.replace(",", ".");
  } else if (/^[+-]?\d{1,3}(?:\.\d{3})+$/.test(text)) {
    // Untuk angka hitungan, titik berulang tiga digit dianggap pemisah ribuan.
    text = text.replace(/\./g, "");
  }

  const parsed = Number(text.replace(/[^0-9.+-]/g, ""));
  return Number.isFinite(parsed) ? parsed : null;
}

function formatDataPerSlsAggregate(value) {
  if (!Number.isFinite(value)) return "";
  if (Math.abs(value - Math.round(value)) < 1e-9) return String(Math.round(value));
  return value.toLocaleString("id-ID", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function normalizeApprovePersonKey(value) {
  return upperText(value)
    .replace(/^@/, "")
    .replace(/\s+/g, " ");
}

function approveEmailLocalPart(value) {
  const text = normalizeApprovePersonKey(value);
  return text.includes("@") ? text.split("@")[0] : text;
}

function sumJumlahApprovePml(rows = []) {
  let total = 0;
  let count = 0;
  for (const row of rows || []) {
    const value = parseDataPerSlsNumber(row?.jumlah_approve_pml);
    if (value == null) continue;
    total += value;
    count += 1;
  }
  return {
    raw: count > 0 ? total : null,
    formatted: count > 0 ? formatDataPerSlsAggregate(total) : "",
    count,
  };
}

function approveRowMatchesPpl(approveRow = {}, target = {}) {
  const approveEmail = normalizeApprovePersonKey(approveRow?.email_ppl || "");
  const approveEmailLocal = approveEmailLocalPart(approveRow?.email_ppl || "");
  const approveName = normalizeApprovePersonKey(approveRow?.nama_ppl || "");

  const targetEmail = normalizeApprovePersonKey(
    target?.email_ppl || target?.email_pencacah || target?.email || ""
  );
  const targetEmailLocal = approveEmailLocalPart(
    target?.email_ppl || target?.email_pencacah || target?.email || ""
  );
  const targetUsername = normalizeApprovePersonKey(
    target?.username_ppl || target?.username_sobat_ppl || ""
  );
  const targetName = normalizeApprovePersonKey(target?.nama_ppl || target?.nama || "");

  return Boolean(
    (approveEmail && targetEmail && approveEmail === targetEmail) ||
    (approveEmailLocal && targetEmailLocal && approveEmailLocal === targetEmailLocal) ||
    (approveEmailLocal && targetUsername && approveEmailLocal === targetUsername) ||
    (approveName && targetName && approveName === targetName)
  );
}

function findApproveRowsForWorkloadRow(workloadRow = {}, approveRows = []) {
  const kodeSls = normalizeStatusSlsCode(workloadRow?.kode_sls || "", 6);
  const byCode = (approveRows || []).filter((row) =>
    normalizeStatusSlsCode(row?.kode_sls || "", 6) === kodeSls
  );
  if (byCode.length === 0) return [];

  const byPpl = byCode.filter((row) => approveRowMatchesPpl(row, workloadRow));
  if (byPpl.length > 0) return byPpl;

  const workloadPmlName = normalizeApprovePersonKey(workloadRow?.nama_pml || "");
  const workloadPmlUsername = normalizeApprovePersonKey(
    workloadRow?.username_pml || workloadRow?.username_sobat_pml || ""
  );
  const byPml = byCode.filter((row) => {
    const approvePmlName = normalizeApprovePersonKey(row?.nama_pml || "");
    const approvePmlEmailLocal = approveEmailLocalPart(row?.email_pml || "");
    return (workloadPmlName && approvePmlName === workloadPmlName) ||
      (workloadPmlUsername && approvePmlEmailLocal === workloadPmlUsername);
  });
  if (byPml.length > 0) return byPml;

  return byCode.length === 1 ? byCode : [];
}

function applyApproveByPmlToWorkload(workload = {}, approveRows = [], role = "PML") {
  if (upperText(role) !== "PML" || !Array.isArray(approveRows) || approveRows.length === 0) {
    return workload;
  }

  const rows = (workload?.rows || []).map((row) => {
    const matchingApproveRows = findApproveRowsForWorkloadRow(row, approveRows);
    const approved = sumJumlahApprovePml(matchingApproveRows);
    const approvedRaw = approved.raw == null ? 0 : approved.raw;
    const approvedFormatted = approved.raw == null ? "0" : approved.formatted;

    const target = parseDataPerSlsNumber(row?.target_jumlah);
    const percentageRaw = target && target > 0 ? (approvedRaw / target) * 100 : null;
    const percentage = percentageRaw == null ? "" : formatPercentageNumber(percentageRaw, 2);

    return {
      ...row,
      realisasi_jumlah: approvedFormatted,
      realisasi_dengan_tidak_ditemukan_jumlah: approvedFormatted,
      jumlah_approve_pml: approvedFormatted,
      jumlah_approve_pml_raw: approvedRaw,
      jumlah_baris_approve_pml: approved.count,
      approve_pml_ditemukan: approved.raw != null,
      persentase: percentage,
      persentase_dengan_tidak_ditemukan: percentage,
      persentase_raw: percentageRaw == null ? "" : percentageRaw,
      sumber_realisasi: "Approve by PML",
    };
  });

  const targetKeluargaRaw = rows.reduce((sum, row) => sum + (parseDataPerSlsNumber(row?.target_keluarga) || 0), 0);
  const targetUsahaRaw = rows.reduce((sum, row) => sum + (parseDataPerSlsNumber(row?.target_usaha) || 0), 0);
  const targetJumlahRaw = rows.reduce((sum, row) => sum + (parseDataPerSlsNumber(row?.target_jumlah) || 0), 0);
  const realisasiKeluargaRaw = rows.reduce((sum, row) => sum + (parseDataPerSlsNumber(row?.realisasi_keluarga) || 0), 0);
  const realisasiUsahaRaw = rows.reduce((sum, row) => sum + (parseDataPerSlsNumber(row?.realisasi_usaha) || 0), 0);
  const realisasiJumlahDariBarisRaw = rows.reduce(
    (sum, row) => sum + (parseDataPerSlsNumber(row?.realisasi_jumlah) || 0),
    0
  );
  // Total PML harus sama persis dengan penjumlahan seluruh kolom
  // "Jumlah Approve PML" pada record PML, walaupun ada kode SLS yang tidak cocok.
  const directApproveTotal = sumJumlahApprovePml(approveRows);
  const realisasiJumlahRaw = directApproveTotal.raw == null
    ? realisasiJumlahDariBarisRaw
    : directApproveTotal.raw;
  const percentageRaw = targetJumlahRaw > 0 ? (realisasiJumlahRaw / targetJumlahRaw) * 100 : null;

  const total = {
    ...(workload?.total || {}),
    target_keluarga: rows.length ? formatDataPerSlsAggregate(targetKeluargaRaw) : "",
    target_usaha: rows.length ? formatDataPerSlsAggregate(targetUsahaRaw) : "",
    target_jumlah: rows.length ? formatDataPerSlsAggregate(targetJumlahRaw) : "",
    realisasi_keluarga: rows.length ? formatDataPerSlsAggregate(realisasiKeluargaRaw) : "",
    realisasi_usaha: rows.length ? formatDataPerSlsAggregate(realisasiUsahaRaw) : "",
    realisasi_jumlah: rows.length ? formatDataPerSlsAggregate(realisasiJumlahRaw) : "",
    jumlah_approve_pml: rows.length ? formatDataPerSlsAggregate(realisasiJumlahRaw) : "",
    persentase: percentageRaw == null ? "" : formatPercentageNumber(percentageRaw, 2),
    persentase_raw: percentageRaw == null ? "" : percentageRaw,
    // Keterangan PML sengaja dikosongkan agar tidak ditulis pada pemberkasan.
    keterangan: "",
    sumber_realisasi: "Approve by PML",
  };

  return { rows, total };
}

function getDataPerSlsValueOrSum(primaryValue, firstValue, secondValue) {
  const primary = cleanText(primaryValue);
  if (primary) return primary;
  const first = parseDataPerSlsNumber(firstValue);
  const second = parseDataPerSlsNumber(secondValue);
  if (first == null && second == null) return "";
  return formatDataPerSlsAggregate((first || 0) + (second || 0));
}

function buildDataPerSlsWorkloadRows(sourceRows = [], role = "PML") {
  const sortedSource = [...(sourceRows || [])].sort((a, b) => {
    const noA = parseDataPerSlsNumber(a?.no_sumber);
    const noB = parseDataPerSlsNumber(b?.no_sumber);
    if (noA != null && noB != null && noA !== noB) return noA - noB;

    const kecDiff = cleanText(a?.kdkec).localeCompare(cleanText(b?.kdkec), "id-ID", { numeric: true });
    if (kecDiff !== 0) return kecDiff;
    const desaDiff = cleanText(a?.kddesa).localeCompare(cleanText(b?.kddesa), "id-ID", { numeric: true });
    if (desaDiff !== 0) return desaDiff;
    return cleanText(a?.kode_sls).localeCompare(cleanText(b?.kode_sls), "id-ID", { numeric: true });
  });

  const rows = sortedSource.map((row, index) => {
    const targetJumlah = getDataPerSlsValueOrSum(
      row?.target_jumlah,
      row?.target_keluarga,
      row?.target_usaha
    );
    const realisasiJumlah = getDataPerSlsValueOrSum(
      row?.realisasi_dengan_tidak_ditemukan_jumlah,
      row?.realisasi_dengan_tidak_ditemukan_keluarga,
      row?.realisasi_dengan_tidak_ditemukan_usaha
    );

    const targetNumber = parseDataPerSlsNumber(targetJumlah);
    const realisasiNumber = parseDataPerSlsNumber(realisasiJumlah);
    const percentageSource = cleanText(row?.persentase_dengan_tidak_ditemukan || "");
    const percentageNumber = parsePercentageNumber(percentageSource);
    const computedPercentage = percentageNumber != null
      ? percentageNumber
      : targetNumber && realisasiNumber != null
        ? (realisasiNumber / targetNumber) * 100
        : null;
    const percentageFormatted = computedPercentage == null
      ? percentageSource
      : formatPercentageNumber(computedPercentage, 2);

    return {
      no: index + 1,
      no_sumber: cleanText(row?.no_sumber || ""),
      nama_pml: cleanText(row?.nama_pml || ""),
      username_pml: cleanText(row?.username_pml || ""),
      username_sobat_pml: cleanText(row?.username_pml || ""),
      nama_ppl: cleanText(row?.nama_ppl || ""),
      username_ppl: cleanText(row?.username_ppl || ""),
      username_sobat_ppl: cleanText(row?.username_ppl || ""),

      kdkec: normalizeStatusSlsCode(row?.kdkec || "", 3),
      kode_kecamatan: normalizeStatusSlsCode(row?.kdkec || "", 3),
      kddesa: normalizeStatusSlsCode(row?.kddesa || "", 3),
      kode_kelurahan: normalizeStatusSlsCode(row?.kddesa || "", 3),
      kode_sls: normalizeStatusSlsCode(row?.kode_sls || "", 6),

      target_keluarga: cleanText(row?.target_keluarga || ""),
      target_usaha: cleanText(row?.target_usaha || ""),
      target_jumlah: targetJumlah,
      target_prelist_keluarga: cleanText(row?.target_keluarga || ""),
      target_prelist_usaha: cleanText(row?.target_usaha || ""),
      target_prelist_jumlah: targetJumlah,

      realisasi_keluarga: cleanText(row?.realisasi_dengan_tidak_ditemukan_keluarga || ""),
      realisasi_usaha: cleanText(row?.realisasi_dengan_tidak_ditemukan_usaha || ""),
      realisasi_jumlah: realisasiJumlah,
      realisasi_dengan_tidak_ditemukan_keluarga: cleanText(row?.realisasi_dengan_tidak_ditemukan_keluarga || ""),
      realisasi_dengan_tidak_ditemukan_usaha: cleanText(row?.realisasi_dengan_tidak_ditemukan_usaha || ""),
      realisasi_dengan_tidak_ditemukan_jumlah: realisasiJumlah,

      persentase: percentageFormatted,
      persentase_dengan_tidak_ditemukan: percentageFormatted,
      persentase_raw: percentageSource,

      realisasi_tanpa_tidak_ditemukan_keluarga: cleanText(row?.realisasi_tanpa_tidak_ditemukan_keluarga || ""),
      realisasi_tanpa_tidak_ditemukan_usaha: cleanText(row?.realisasi_tanpa_tidak_ditemukan_usaha || ""),
      realisasi_tanpa_tidak_ditemukan_jumlah: cleanText(row?.realisasi_tanpa_tidak_ditemukan_jumlah || ""),
      persentase_tanpa_tidak_ditemukan: cleanText(row?.persentase_tanpa_tidak_ditemukan || ""),

      keterangan: cleanText(row?.keterangan || ""),
      status: cleanText(row?.status || ""),
      role,
    };
  });

  const sum = (key) => rows.reduce((total, row) => {
    const value = parseDataPerSlsNumber(row?.[key]);
    return total + (value == null ? 0 : value);
  }, 0);

  const targetKeluargaRaw = sum("target_keluarga");
  const targetUsahaRaw = sum("target_usaha");
  const targetJumlahRaw = sum("target_jumlah");
  const realisasiKeluargaRaw = sum("realisasi_keluarga");
  const realisasiUsahaRaw = sum("realisasi_usaha");
  const realisasiJumlahRaw = sum("realisasi_jumlah");
  const percentageRaw = targetJumlahRaw > 0
    ? (realisasiJumlahRaw / targetJumlahRaw) * 100
    : null;

  const total = {
    target_keluarga: rows.length ? formatDataPerSlsAggregate(targetKeluargaRaw) : "",
    target_usaha: rows.length ? formatDataPerSlsAggregate(targetUsahaRaw) : "",
    target_jumlah: rows.length ? formatDataPerSlsAggregate(targetJumlahRaw) : "",
    realisasi_keluarga: rows.length ? formatDataPerSlsAggregate(realisasiKeluargaRaw) : "",
    realisasi_usaha: rows.length ? formatDataPerSlsAggregate(realisasiUsahaRaw) : "",
    realisasi_jumlah: rows.length ? formatDataPerSlsAggregate(realisasiJumlahRaw) : "",
    persentase: percentageRaw == null ? "" : formatPercentageNumber(percentageRaw, 2),
    persentase_raw: percentageRaw == null ? "" : percentageRaw,
    keterangan: percentageRaw == null
      ? ""
      : percentageRaw >= 40
        ? "Bisa Dibayar karena lebih dari 40%"
        : "Belum Bisa Dibayar karena kurang dari 40%",
  };

  return { rows, total };
}

function buildLampiranStatusSlsCode(row = {}) {
  // Sheet Lampiran menyimpan kode SLS sebagai kdsls (4 digit) + kdsubsls (2 digit).
  // Jangan memakai kdsubslspanjang karena kolom itu dapat memuat kode wilayah panjang
  // dan enam digit terakhirnya tidak selalu identik dengan Kode SLS pada sheet Status SLS.
  const direct = normalizeStatusSlsCode(row?.kode_sls || "", 6);
  if (direct) return direct;

  const kdsls = normalizeStatusSlsCode(row?.kdsls || "", 4);
  const kdsubsls = normalizeStatusSlsCode(row?.kdsubsls || "", 2);
  return kdsls ? `${kdsls}${kdsubsls || "00"}` : "";
}

function statusSlsRowsHaveUsableFilter(statusSlsRows = []) {
  return (statusSlsRows || []).some((row) =>
    cleanText(row?.kode_sls) || cleanText(row?.status) || cleanText(row?.status_raw)
  );
}

function findStatusSlsRowForLampiran(lampiranRow = {}, statusSlsRows = []) {
  const kodeKec = normalizeStatusSlsCode(lampiranRow?.kdkec || "", 3);
  const kodeDesa = normalizeStatusSlsCode(lampiranRow?.kddesa || "", 3);
  const kodeSls = buildLampiranStatusSlsCode(lampiranRow);
  const emailPpl = upperText(lampiranRow?.email_pencacah || lampiranRow?.email_ppl || "");
  const namaPpl = upperText(lampiranRow?.nama_ppl || "");
  const emailPml = upperText(lampiranRow?.email_pengawas || lampiranRow?.email_pml || "");
  const namaPml = upperText(lampiranRow?.nama_pml || "");

  const sameCodes = (row) => {
    const rowKec = normalizeStatusSlsCode(row?.kdkec || "", 3);
    const rowDesa = normalizeStatusSlsCode(row?.kddesa || "", 3);
    const rowSls = normalizeStatusSlsCode(row?.kode_sls || "", 6);
    return (!kodeKec || rowKec === kodeKec) &&
      (!kodeDesa || rowDesa === kodeDesa) &&
      (!kodeSls || rowSls === kodeSls);
  };

  const candidates = (statusSlsRows || []).filter(sameCodes);
  if (candidates.length === 0) return null;

  if (emailPpl) {
    const match = candidates.find((row) => upperText(row?.email_ppl || "") === emailPpl);
    if (match) return match;
  }
  if (namaPpl) {
    const match = candidates.find((row) => upperText(row?.nama_ppl || "") === namaPpl);
    if (match) return match;
  }
  if (emailPml) {
    const match = candidates.find((row) => upperText(row?.email_pml || "") === emailPml);
    if (match) return match;
  }
  if (namaPml) {
    const match = candidates.find((row) => upperText(row?.nama_pml || "") === namaPml);
    if (match) return match;
  }

  // Fallback berdasarkan kode hanya boleh dipakai bila hasilnya tunggal, agar SLS
  // dengan kode sama milik petugas berbeda tidak salah ditempelkan.
  return candidates.length === 1 ? candidates[0] : null;
}

function filterLampiranRowsByStatusSls(lampiranRows = [], statusSlsRows = []) {
  if (!statusSlsRowsHaveUsableFilter(statusSlsRows)) return [...(lampiranRows || [])];

  return (lampiranRows || [])
    .map((row) => {
      const statusRow = findStatusSlsRowForLampiran(row, statusSlsRows);
      if (!statusRow || !isAllowedStatusSls(statusRow?.status || statusRow?.status_raw)) return null;
      return {
        ...row,
        status_sls: normalizeStatusSlsLabel(statusRow?.status || statusRow?.status_raw),
        status_sls_raw: cleanText(statusRow?.status_raw || statusRow?.status || ""),
        status_jumlah_prelist: cleanText(statusRow?.jumlah_prelist || ""),
        status_jumlah_realisasi: cleanText(statusRow?.jumlah_realisasi || ""),
        status_persentase: cleanText(statusRow?.persentase || ""),
      };
    })
    .filter(Boolean);
}

function buildBerkasLampiranTableRows(lampiranRows = [], role = "PML", pembayaranRows = [], statusSlsRows = []) {
  if (!Array.isArray(lampiranRows) || lampiranRows.length === 0) return [];

  // Bila sheet Status SLS tersedia, hanya baris dengan status Selesai atau
  // Sedang Dikerjakan yang masuk ke pengelompokan dan perhitungan {jumlah}.
  const filteredLampiranRows = filterLampiranRowsByStatusSls(lampiranRows, statusSlsRows);
  const grouped = groupLampiranRows(filteredLampiranRows, role);

  return grouped.map((row, index) => {
    const jumlah = Number(row?.jumlah || 0);
    const jumlah40 = Math.ceil(jumlah * 0.4);
    const jumlah60 = jumlah - jumlah40;
    const namaPml = cleanText(row?.nama_pml || "");
    const namaPpl = cleanText(row?.nama_ppl || "");

    // Setiap baris tabel PML mewakili PPL, sehingga nomor generik pada baris
    // diarahkan ke kontrak PPL. Nomor kontrak PML tetap tersedia lewat tag khusus.
    const nomorKontrakPml = cleanText(row?.nomor_kontrak_pml || "");
    const nomorKontrakPpl = cleanText(row?.nomor_kontrak_ppl || "");
    const nomorKontrakBaris = nomorKontrakPpl;

    const pembayaranRow = findPembayaranRowForLampiran(row, pembayaranRows);
    const prelistTotal = cleanText(pembayaranRow?.prelist_total || pembayaranRow?.target_prelist || "");
    const realisasi = cleanText(pembayaranRow?.realisasi_total || pembayaranRow?.realisasi_hasil_pendataan || "");
    const persentaseRaw = cleanText(pembayaranRow?.persentase_pendataan || pembayaranRow?.persentase_prelist || "");
    const persentaseNumber = parsePercentageNumber(persentaseRaw);
    // Semua tag tampilan persentase memakai dua angka di belakang koma.
    // Nilai asli tetap tersedia melalui {persentase_raw}.
    const persentaseFormat = persentaseNumber == null
      ? persentaseRaw
      : formatPercentageNumber(persentaseNumber, 2);
    const persentase = persentaseFormat;

    return {
      no: index + 1,
      nama: namaPpl,
      nama_petugas: namaPpl,
      nama_pml: namaPml,
      nama_pengawas: namaPml,
      nama_ppl: namaPpl,
      email_pengawas: cleanText(row?.email_pengawas || ""),
      email_pml: cleanText(row?.email_pengawas || ""),
      email_pencacah: cleanText(row?.email_pencacah || ""),
      email_ppl: cleanText(row?.email_pencacah || ""),
      email: cleanText(row?.email_pencacah || ""),

      kdprov: cleanText(row?.kdprov || ""),
      kdkab: cleanText(row?.kdkab || ""),
      kdkec: cleanText(row?.kdkec || ""),
      kddesa: cleanText(row?.kddesa || ""),
      nmprov: cleanText(row?.nmprov || ""),
      nmkab: cleanText(row?.nmkab || ""),
      kecamatan: formatKodeNama(row?.kdkec, row?.kecamatan),
      kelurahan: formatKodeNama(row?.kddesa, row?.kelurahan),

      jumlah,
      total_jumlah: jumlah,
      jumlah_40: jumlah40,
      jumlah_60: jumlah60,
      sls_total: jumlah,
      sls_40: jumlah40,
      sls_60: jumlah60,

      nomor_spk: nomorKontrakBaris,
      nomor_kontrak: nomorKontrakBaris,
      nomor_kontrak_pml: nomorKontrakPml,
      nomor_kontrak_ppl: nomorKontrakPpl,

      // Struktur/identitas tabel tetap dari sheet Lampiran. Tiga metrik berikut
      // diambil dari baris PPL yang cocok pada sheet Pembayaran.
      prelist_total: prelistTotal,
      target_prelist: prelistTotal,
      realisasi,
      realisasi_total: realisasi,
      realisasi_hasil_pendataan: realisasi,
      persentase,
      persentase_pendataan: persentase,
      persentase_prelist: persentase,
      persentase_raw: persentaseRaw,
      persentase_format: persentaseFormat,

      // Ringkasan status SLS yang lolos filter untuk baris kelompok ini.
      // Karena satu baris tabel dapat berisi beberapa SLS, status digabung unik.
      status_sls: (() => {
        const statuses = filteredLampiranRows
          .filter((item) =>
            upperText(item?.nama_ppl || "") === upperText(row?.nama_ppl || "") &&
            normalizeStatusSlsCode(item?.kdkec || "", 3) === normalizeStatusSlsCode(row?.kdkec || "", 3) &&
            normalizeStatusSlsCode(item?.kddesa || "", 3) === normalizeStatusSlsCode(row?.kddesa || "", 3)
          )
          .map((item) => normalizeStatusSlsLabel(item?.status_sls || ""))
          .filter(Boolean);
        return [...new Set(statuses)].join(" & ");
      })(),

      // Alias lama tetap disediakan agar template sebelumnya tidak rusak.
      sls_ongoing: "",
      sls_selesai_sedang_dikerjakan: "",
      persentase_sls: "",
      tanggal_screenshot: "",
      flag: "",
    };
  });
}

function buildPembayaranTableRow(row = {}, index = 0) {
  const jabatan = upperText(row?.jabatan || row?.jabatan_raw || "");
  const namaPpl = cleanText(row?.nama_ppl || (/PPL|PENCACAH/.test(jabatan) ? row?.nama : ""));
  const namaPml = cleanText(row?.nama_pengawas || row?.nama_pml || (/PML|PENGAWAS/.test(jabatan) ? row?.nama : ""));
  const targetPrelist = cleanText(row?.target_prelist || row?.prelist_total || "");
  const realisasiHasil = cleanText(row?.realisasi_hasil_pendataan || row?.realisasi_total || "");
  const slsOngoing = cleanText(row?.sls_ongoing || row?.sls_selesai_sedang_dikerjakan || "");
  const persentasePrelist = cleanText(row?.persentase_prelist || row?.persentase_pendataan || "");
  const nomorSpk = cleanText(row?.nomor_spk || row?.nomor_kontrak || "");

  return {
    no: index + 1,
    nama: namaPpl || cleanText(row?.nama || ""),
    nama_petugas: namaPpl || cleanText(row?.nama || ""),
    nama_ppl: namaPpl,
    email_ppl: cleanText(row?.email_ppl || row?.email || ""),
    email: cleanText(row?.email || row?.email_ppl || ""),
    nik: cleanText(row?.nik || ""),
    jabatan: cleanText(row?.jabatan || row?.jabatan_raw || ""),

    nama_pml: namaPml,
    nama_pengawas: namaPml,
    email_pengawas: cleanText(row?.email_pengawas || row?.email_pml || ""),

    nomor_spk: nomorSpk,
    nomor_kontrak: nomorSpk,
    kecamatan: cleanText(row?.kecamatan || row?.wilayah || ""),
    sls_total: cleanText(row?.sls_total || ""),
    sls_40: cleanText(row?.sls_40 || ""),
    sls_60: cleanText(row?.sls_60 || ""),
    sls_ongoing: slsOngoing,
    sls_selesai_sedang_dikerjakan: slsOngoing,
    persentase_sls: cleanText(row?.persentase_sls || ""),
    tanggal_screenshot: cleanText(row?.tanggal_screenshot || ""),

    target_prelist: targetPrelist,
    prelist_total: targetPrelist,
    realisasi_hasil_pendataan: realisasiHasil,
    realisasi_total: realisasiHasil,
    persentase_prelist: persentasePrelist,
    persentase_pendataan: persentasePrelist,
    flag: cleanText(row?.flag || ""),
  };
}

function prefixTemplateData(prefix, data = {}) {
  const prefixed = {};
  for (const [key, value] of Object.entries(data || {})) {
    prefixed[`${prefix}_${key}`] = value;
  }
  return prefixed;
}

function buildBerkasPembayaranTemplateData(formValues, record, role, nikLookup) {
  const isPml = upperText(role) === "PML";
  const lampiranRows = Array.isArray(record?.lampiranRows) ? record.lampiranRows : [];
  const pembayaranRows = Array.isArray(record?.pembayaranRows) ? record.pembayaranRows : [];
  const statusSlsRows = Array.isArray(record?.statusSlsRows) ? record.statusSlsRows : [];
  const dataPerSlsRows = Array.isArray(record?.dataPerSlsRows) ? record.dataPerSlsRows : [];
  const approveByPmlRows = Array.isArray(record?.approveByPmlRows) ? record.approveByPmlRows : [];
   // 🔍 LOG SEMENTARA
  console.log("Isi foto_bukti per baris:", approveByPmlRows.map(r => ({
    nama_ppl: r.nama_ppl,
    foto_bukti_pml: r.foto_bukti_pml,
    foto_bukti_ppl: r.foto_bukti_ppl,
  })));
  const firstLampiran = lampiranRows[0] || {};
  const displayName = cleanText(
    record?.displayName || record?.bappRow?.nama || (isPml ? firstLampiran?.nama_pml : firstLampiran?.nama_ppl)
  );
  const email = cleanText(
    record?.email || record?.bappRow?.email || (isPml ? firstLampiran?.email_pengawas : firstLampiran?.email_pencacah)
  );
  const nomorKontrakLampiran = cleanText(
    isPml ? firstLampiran?.nomor_kontrak_pml : firstLampiran?.nomor_kontrak_ppl
  );
  // Isi tabel mengikuti sheet Lampiran, memakai pengelompokan yang sama dengan
  // dokumen Lampiran/BAST terpisah. Metrik target dan realisasi kemudian
  // diperkaya dari sheet Data per SLS, khusus kolom "Dengan Tidak Ditemukan".
  const lampiranTableRowsBase = buildBerkasLampiranTableRows(
    lampiranRows,
    role,
    pembayaranRows,
    statusSlsRows
  );
  const bebanKerjaBase = buildDataPerSlsWorkloadRows(dataPerSlsRows, role);
  const bebanKerja = applyApproveByPmlToWorkload(
    bebanKerjaBase,
    approveByPmlRows,
    role
  );

  const normalizeMatchKey = (value) => upperText(value).replace(/^@/, "").replace(/\s+/g, " ");
  const emailLocalPart = (value) => {
    const text = normalizeMatchKey(value);
    return text.includes("@") ? text.split("@")[0] : text;
  };

  // Untuk dokumen PML, setiap baris tabel peserta mewakili satu PPL. Target,
  // realisasi, dan persentase dijumlahkan dari seluruh SLS milik PPL tersebut.
  // Untuk dokumen PPL, seluruh baris Data per SLS pada record adalah milik PPL itu.
  const lampiranTableRows = lampiranTableRowsBase.map((tableRow) => {
    const pplName = normalizeMatchKey(tableRow?.nama_ppl || tableRow?.nama || "");
    const pplEmailLocal = emailLocalPart(tableRow?.email_ppl || tableRow?.email_pencacah || tableRow?.email || "");

    const matchedSourceRows = isPml
      ? dataPerSlsRows.filter((sourceRow) => {
          const sourceName = normalizeMatchKey(sourceRow?.nama_ppl || "");
          const sourceUsername = normalizeMatchKey(sourceRow?.username_ppl || "");
          return (pplName && sourceName === pplName) ||
            (pplEmailLocal && sourceUsername === pplEmailLocal);
        })
      : dataPerSlsRows;

    const matchedApproveRows = isPml
      ? approveByPmlRows.filter((approveRow) => approveRowMatchesPpl(approveRow, tableRow))
      : [];

    if (matchedSourceRows.length === 0 && matchedApproveRows.length === 0) return tableRow;

    const pplWorkloadBase = buildDataPerSlsWorkloadRows(matchedSourceRows, "PPL");
    const pplWorkload = isPml
      ? applyApproveByPmlToWorkload(pplWorkloadBase, matchedApproveRows, "PML")
      : pplWorkloadBase;
    const total = pplWorkload.total || {};
    const approveTotal = sumJumlahApprovePml(matchedApproveRows);
    const targetForPercentage = parseDataPerSlsNumber(total.target_jumlah || tableRow?.jumlah_pre || tableRow?.prelist_total);
    const useApproveSource = isPml && matchedApproveRows.length > 0;
    const approveValueRaw = useApproveSource
      ? (approveTotal.raw == null ? 0 : approveTotal.raw)
      : null;
    const approvePercentageRaw = useApproveSource && targetForPercentage
      ? (approveValueRaw / targetForPercentage) * 100
      : null;
    const approvePercentage = useApproveSource
      ? (approvePercentageRaw == null ? "" : formatPercentageNumber(approvePercentageRaw, 2))
      : total.persentase;
    const rowRealisasi = useApproveSource
      ? formatDataPerSlsAggregate(approveValueRaw)
      : total.realisasi_jumlah;

    return {
      ...tableRow,
      pre_keluarga: total.target_keluarga,
      pre_usaha: total.target_usaha,
      jumlah_pre: total.target_jumlah,
      target_keluarga: total.target_keluarga,
      target_usaha: total.target_usaha,
      target_jumlah: total.target_jumlah,
      prelist_total: total.target_jumlah,
      target_prelist: total.target_jumlah,

      realisasi_keluarga: total.realisasi_keluarga,
      realisasi_usaha: total.realisasi_usaha,
      jumlah_realisasi: rowRealisasi,
      realisasi: rowRealisasi,
      realisasi_total: rowRealisasi,
      realisasi_hasil_pendataan: rowRealisasi,
      jumlah_approve_pml: useApproveSource ? formatDataPerSlsAggregate(approveValueRaw) : "",

      persentase: approvePercentage,
      pesentase: approvePercentage,
      persentase_pendataan: approvePercentage,
      persentase_prelist: approvePercentage,
      persentase_raw: approvePercentageRaw == null ? total.persentase_raw : approvePercentageRaw,
      persentase_format: approvePercentage,
      // Pemberkasan PML dan PPL tidak menampilkan keterangan/status pembayaran.
      status: "",
      keterangan: "",
    };
  });

  // Hitung rata-rata dari baris yang benar-benar tampil pada tabel. Bila Data per
  // SLS tersedia, nilai yang dirata-ratakan sudah memakai realisasi "Dengan Tidak
  // Ditemukan", bukan lagi nilai lama dari sheet Pembayaran.
  const averagePersentase = calculateAveragePersentase(lampiranTableRows);

  const bappRow = record?.bappRow || {
    nama: displayName,
    email,
    jabatan: role,
    jabatan_raw: role,
    nik: nikLookup?.get(upperText(displayName)) || "",
    nomor_spk: nomorKontrakLampiran,
    nomor_kontrak: nomorKontrakLampiran,
  };

  const syntheticLampiranRow = {
    ...firstLampiran,
    nama_pml: isPml ? displayName : cleanText(firstLampiran?.nama_pml),
    nama_ppl: isPml ? cleanText(firstLampiran?.nama_ppl) : displayName,
    email_pengawas: isPml ? email : cleanText(firstLampiran?.email_pengawas),
    email_pencacah: isPml ? cleanText(firstLampiran?.email_pencacah) : email,
    nomor_kontrak_pml: isPml ? cleanText(bappRow?.nomor_spk || bappRow?.nomor_kontrak || nomorKontrakLampiran) : cleanText(firstLampiran?.nomor_kontrak_pml),
    nomor_kontrak_ppl: isPml ? cleanText(firstLampiran?.nomor_kontrak_ppl) : cleanText(bappRow?.nomor_spk || bappRow?.nomor_kontrak || nomorKontrakLampiran),
  };

  const rowsForBast = lampiranRows.length > 0 ? lampiranRows : [syntheticLampiranRow];
  console.log("Record approveByPmlRows:", record?.approveByPmlRows?.length, record?.approveByPmlRows);
  console.log("Record displayName/email:", record?.displayName, record?.email);
  
  const bappRaw = buildBappTemplateData(formValues || {}, bappRow, role, approveByPmlRows);
  const bastRaw = buildBastTemplateData(formValues || {}, rowsForBast, role, nikLookup);
  const nik = cleanText(bappRaw.nik || bastRaw.nik || nikLookup?.get(upperText(displayName)) || "");
  const nomorKontrak = cleanText(bappRaw.nomor_kontrak || bastRaw.nomor_perjanjian || nomorKontrakLampiran);
  const bapp = {
    ...bappRaw,
    nama: displayName,
    nama_peserta: displayName,
    nama_petugas: displayName,
    email,
    nik,
    nomor_kontrak: nomorKontrak,
    nomor_prefix: extractNomorPrefix(nomorKontrak),
  };
  const bast = {
    ...bastRaw,
    nama: displayName,
    nik,
    nomor_perjanjian: bastRaw.nomor_perjanjian || nomorKontrak,
  };
  const suratPernyataan = {
    nomor_prefix: extractNomorPrefix(nomorKontrak),
    nama_petugas: displayName,
    nama: displayName,
    nik,
    nomor_spk: nomorKontrak,
    nomor_kontrak: nomorKontrak,
  };
  const pembayaranRingkasan = buildPembayaranTableRow(bappRow, 0);
  const firstLampiranTableRow = lampiranTableRows[0] || {};
  const hasDataPerSls = bebanKerja.rows.length > 0;
  const dataPerSlsSummary = bebanKerja.total || {};
  const summaryTargetKeluarga = hasDataPerSls ? dataPerSlsSummary.target_keluarga : "";
  const summaryTargetUsaha = hasDataPerSls ? dataPerSlsSummary.target_usaha : "";
  const summaryTargetJumlah = hasDataPerSls
    ? dataPerSlsSummary.target_jumlah
    : pembayaranRingkasan.prelist_total;
  const summaryRealisasiKeluarga = hasDataPerSls ? dataPerSlsSummary.realisasi_keluarga : "";
  const summaryRealisasiUsaha = hasDataPerSls ? dataPerSlsSummary.realisasi_usaha : "";
  const summaryRealisasiJumlah = hasDataPerSls
    ? dataPerSlsSummary.realisasi_jumlah
    : pembayaranRingkasan.realisasi_total;
  const summaryPersentase = hasDataPerSls
    ? dataPerSlsSummary.persentase
    : pembayaranRingkasan.persentase_pendataan;

  // Variabel tanpa prefix dipertahankan untuk kompatibilitas dengan template lama.
  // Variabel berprefix memberi ruang bagi template gabungan saat ada nama tag yang
  // bentrok, misalnya {bapp_nomor_surat} dan {bast_nomor_surat}.
  return {
    ...bapp,
    ...bast,
    ...suratPernyataan,
    ...prefixTemplateData("bapp", bapp),
    ...prefixTemplateData("bast", bast),
    ...prefixTemplateData("surat_pernyataan", suratPernyataan),
    role,
    jenis: role,
    jenis_petugas: role,
    nama: displayName,
    nama_petugas: displayName,
    nama_pml: isPml ? displayName : cleanText(pembayaranRingkasan.nama_pml || firstLampiran?.nama_pml || ""),
    nama_pengawas: isPml ? displayName : cleanText(pembayaranRingkasan.nama_pengawas || firstLampiran?.nama_pml || ""),
    email_pengawas: isPml ? email : cleanText(pembayaranRingkasan.email_pengawas || firstLampiran?.email_pengawas || ""),
    nama_ppl: isPml ? cleanText(firstLampiranTableRow.nama_ppl || firstLampiran?.nama_ppl || "") : displayName,
    email,
    nik,
    nomor_spk: cleanText(pembayaranRingkasan.nomor_spk || nomorKontrak),
    nomor_kontrak: nomorKontrak,
    nomor_perjanjian: bast.nomor_perjanjian || nomorKontrak,
    nomor_surat: bast.nomor_surat || bapp.nomor_surat || "",
    nomor_prefix: extractNomorPrefix(nomorKontrak),

    // Ringkasan PML/PPL langsung dari kolom sheet Pembayaran.
    sls_total: pembayaranRingkasan.sls_total,
    sls_40: pembayaranRingkasan.sls_40,
    sls_60: pembayaranRingkasan.sls_60,
    sls_ongoing: pembayaranRingkasan.sls_ongoing,
    sls_selesai_sedang_dikerjakan: pembayaranRingkasan.sls_ongoing,
    persentase_sls: pembayaranRingkasan.persentase_sls,
    tanggal_screenshot: pembayaranRingkasan.tanggal_screenshot,
    // Ringkasan target/realisasi memakai agregat sheet Data per SLS bila tersedia.
    // Realisasi yang dipakai adalah kolom "Dengan Tidak Ditemukan".
    pre_keluarga: summaryTargetKeluarga,
    pre_usaha: summaryTargetUsaha,
    jumlah_pre: summaryTargetJumlah,
    target_keluarga: summaryTargetKeluarga,
    target_usaha: summaryTargetUsaha,
    target_jumlah: summaryTargetJumlah,
    target_prelist: summaryTargetJumlah,
    prelist_total: summaryTargetJumlah,

    realisasi_keluarga: summaryRealisasiKeluarga,
    realisasi_usaha: summaryRealisasiUsaha,
    jumlah_realisasi: summaryRealisasiJumlah,
    realisasi: summaryRealisasiJumlah,
    realisasi_hasil_pendataan: summaryRealisasiJumlah,
    realisasi_total: summaryRealisasiJumlah,

    persentase: summaryPersentase,
    pesentase: summaryPersentase,
    persentase_prelist: summaryPersentase,
    persentase_pendataan: summaryPersentase,
    // Keterangan/status tidak ditulis pada pemberkasan PML maupun PPL.
    status: "",
    keterangan: "",
    flag: pembayaranRingkasan.flag,
    kecamatan: pembayaranRingkasan.kecamatan,

    // Rata-rata aritmetika Persentase Pendataan seluruh PPL yang masuk ke tabel.
    // Setiap PPL dihitung sekali walaupun memiliki beberapa baris wilayah di Lampiran.
    average_persentase: averagePersentase.formatted,
    rata_rata_persentase: averagePersentase.formatted,
    average_persentase_raw: averagePersentase.raw == null ? "" : averagePersentase.raw,
    jumlah_persentase_dihitung: averagePersentase.count,

    // Diagnostik/filter Status SLS.
    status_filter_aktif: statusSlsRowsHaveUsableFilter(statusSlsRows),
    jumlah_status_sls_sumber: statusSlsRows.length,
    jumlah_baris_lampiran_sebelum_filter: lampiranRows.length,
    jumlah_baris_tabel_setelah_filter: lampiranTableRows.length,

    // Data tabel beban kerja pada halaman terakhir, langsung dari sheet Data per SLS.
    jumlah_baris_data_per_sls: bebanKerja.rows.length,
    beban_kerja_rows: bebanKerja.rows,
    data_per_sls_rows: bebanKerja.rows,
    rincian_data_per_sls: bebanKerja.rows,
    beban_kerja_total: bebanKerja.total,
    total_target_keluarga: bebanKerja.total.target_keluarga,
    total_target_usaha: bebanKerja.total.target_usaha,
    total_target_jumlah: bebanKerja.total.target_jumlah,
    total_realisasi_keluarga: bebanKerja.total.realisasi_keluarga,
    total_realisasi_usaha: bebanKerja.total.realisasi_usaha,
    total_realisasi_jumlah: bebanKerja.total.realisasi_jumlah,
    jumlah_approve_pml: isPml ? bebanKerja.total.realisasi_jumlah : "",
    total_jumlah_approve_pml: isPml ? bebanKerja.total.realisasi_jumlah : "",
    total_persentase_data_per_sls: bebanKerja.total.persentase,
    keterangan_pembayaran_data_per_sls: "",
    sumber_realisasi_pml: isPml && approveByPmlRows.length > 0 ? "Approve by PML" : "Data per SLS",
    jumlah_baris_approve_by_pml: approveByPmlRows.length,

    // Seluruh alias loop tabel memakai struktur sheet Lampiran. Kolom progres pada
    // tiap peserta diperkaya dari sheet Pembayaran melalui Email Pencacah/Nama PPL.
    peserta: lampiranTableRows,
    pembayaran: lampiranTableRows,
    pembayaran_rows: lampiranTableRows,
    rincian_pembayaran: lampiranTableRows,
    tabel_ppl: lampiranTableRows,
    ppl_rows: lampiranTableRows,
    lampiran_peserta: lampiranTableRows,
    bapp_peserta: bapp.peserta || [],
    bast_peserta: bast.peserta || [],
    tampil_surat_pernyataan: isPml,
    is_pml: isPml,
    is_ppl: !isPml,
  };
}

const WORDPROCESSING_ML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const XML_NS = "http://www.w3.org/XML/1998/namespace";

function wordLocalName(node) {
  return String(node?.localName || node?.nodeName || "").split(":").pop();
}

function getDirectWordChildren(node, localName) {
  return Array.from(node?.childNodes || []).filter(
    (child) => child?.nodeType === 1 && wordLocalName(child) === localName
  );
}

function getWordNodeText(node) {
  const textNodes = node?.getElementsByTagNameNS
    ? Array.from(node.getElementsByTagNameNS(WORDPROCESSING_ML_NS, "t"))
    : Array.from(node?.getElementsByTagName?.("w:t") || []);
  return textNodes.map((item) => item.textContent || "").join("");
}

function createWordElement(xmlDoc, localName) {
  return xmlDoc.createElementNS(WORDPROCESSING_ML_NS, `w:${localName}`);
}

function setWordCellText(xmlDoc, cell, value) {
  if (!cell) return;

  const text = String(value ?? "");
  const paragraphs = Array.from(
    cell.getElementsByTagNameNS(WORDPROCESSING_ML_NS, "p")
  );
  let paragraph = paragraphs[0];

  if (!paragraph) {
    paragraph = createWordElement(xmlDoc, "p");
    cell.appendChild(paragraph);
  }

  const firstRun = Array.from(
    paragraph.getElementsByTagNameNS(WORDPROCESSING_ML_NS, "r")
  )[0];
  const firstRunProperties = firstRun
    ? Array.from(firstRun.childNodes || []).find(
        (child) => child?.nodeType === 1 && wordLocalName(child) === "rPr"
      )
    : null;

  // Pertahankan properti paragraf, lalu ganti isi sel dengan satu run baru.
  for (const child of Array.from(paragraph.childNodes || [])) {
    if (!(child?.nodeType === 1 && wordLocalName(child) === "pPr")) {
      paragraph.removeChild(child);
    }
  }

  // Hapus paragraf tambahan dari sel template agar tinggi baris tidak membengkak.
  for (const extraParagraph of paragraphs.slice(1)) {
    extraParagraph.parentNode?.removeChild(extraParagraph);
  }

  const run = createWordElement(xmlDoc, "r");
  if (firstRunProperties) run.appendChild(firstRunProperties.cloneNode(true));

  const textNode = createWordElement(xmlDoc, "t");
  if (/^\s|\s$/.test(text)) textNode.setAttributeNS(XML_NS, "xml:space", "preserve");
  textNode.textContent = text;
  run.appendChild(textNode);
  paragraph.appendChild(run);
}

function setWordRowValues(xmlDoc, rowNode, values = []) {
  const cells = getDirectWordChildren(rowNode, "tc");
  values.forEach((value, index) => {
    if (cells[index]) setWordCellText(xmlDoc, cells[index], value);
  });
}

function fillBebanKerjaTableInDocxZip(zip, workloadRows = [], workloadTotal = {}) {
  const documentFile = zip?.file?.("word/document.xml");
  if (!documentFile || typeof DOMParser === "undefined" || typeof XMLSerializer === "undefined") {
    console.warn("Tabel beban kerja tidak dapat diisi karena document.xml atau XML DOM tidak tersedia.");
    return;
  }

  const xmlText = documentFile.asText();
  const xmlDoc = new DOMParser().parseFromString(xmlText, "application/xml");
  const parserErrors = xmlDoc.getElementsByTagName("parsererror");
  if (parserErrors.length > 0) {
    console.warn("Gagal membaca document.xml untuk mengisi tabel beban kerja.");
    return;
  }

  const tables = Array.from(
    xmlDoc.getElementsByTagNameNS(WORDPROCESSING_ML_NS, "tbl")
  );
  const targetTable = tables.find((table) => {
    const tableText = getWordNodeText(table).replace(/\s+/g, " ").toUpperCase();
    return tableText.includes("TARGET PRELIST AWAL") &&
      tableText.includes("USERNAME SOBAT") &&
      tableText.includes("SLS/SUB-SLS");
  });

  if (!targetTable) {
    console.warn("Tabel Beban Kerja pada halaman terakhir template tidak ditemukan.");
    return;
  }

  const tableRows = getDirectWordChildren(targetTable, "tr");
  if (tableRows.length < 5) {
    console.warn("Struktur tabel Beban Kerja tidak sesuai template.");
    return;
  }

  // Tiga baris pertama adalah header. Baris dengan teks "Jumlah" adalah footer.
  // Template PPL memiliki 16 kolom detail, sedangkan template PML yang dikirim
  // memiliki 14 kolom detail. Keduanya ditangani tanpa mengubah format Word.
  const headerRowCount = 3;
  let totalRow = tableRows.find((row, index) => {
    if (index < headerRowCount) return false;
    const cells = getDirectWordChildren(row, "tc");
    const firstCellText = getWordNodeText(cells[0] || "")
      .replace(/\s+/g, " ")
      .trim()
      .toUpperCase();
    return firstCellText === "JUMLAH";
  });
  if (!totalRow) totalRow = tableRows[tableRows.length - 1];

  const totalRowIndex = tableRows.indexOf(totalRow);
  const candidateDetailRows = tableRows.slice(headerRowCount, totalRowIndex);
  const templateRow = candidateDetailRows
    .map((row) => ({ row, count: getDirectWordChildren(row, "tc").length }))
    .sort((a, b) => b.count - a.count)[0]?.row;

  if (!templateRow) {
    console.warn("Baris template kosong pada tabel Beban Kerja tidak ditemukan.");
    return;
  }

  const detailColumnCount = getDirectWordChildren(templateRow, "tc").length;
  if (![14, 16].includes(detailColumnCount)) {
    console.warn(`Jumlah kolom tabel Beban Kerja tidak didukung: ${detailColumnCount}.`);
    return;
  }

  const templateClone = templateRow.cloneNode(true);
  const minimumTemplateRows = Math.max(totalRowIndex - headerRowCount, 1);
  for (const row of candidateDetailRows) {
    targetTable.removeChild(row);
  }

  const buildDetailValues = (row) => {
    const common = [
      row.no,
      row.nama_pml,
      row.username_pml,
      row.nama_ppl,
      row.username_ppl,
      row.kdkec,
      row.kddesa,
      row.kode_sls,
      row.target_keluarga,
      row.target_usaha,
      row.target_jumlah,
    ];

    // PPL: Realisasi Keluarga, Usaha, Jumlah.
    if (detailColumnCount === 16) {
      return [
        ...common,
        row.realisasi_keluarga,
        row.realisasi_usaha,
        row.realisasi_jumlah,
        row.persentase,
        "", // Kolom Keterangan PPL dikosongkan.
      ];
    }

    // PML: template hanya menyediakan satu kolom Realisasi, sehingga dipakai
    // nilai Jumlah dari blok "Dengan Tidak Ditemukan".
    return [
      ...common,
      row.realisasi_jumlah,
      row.persentase,
      "", // Kolom Keterangan PML dikosongkan.
    ];
  };

  // Pertahankan minimal jumlah baris bawaan template agar posisi footer stabil.
  const rowCountToInsert = Math.max(workloadRows.length, minimumTemplateRows);
  for (let rowIndex = 0; rowIndex < rowCountToInsert; rowIndex++) {
    const row = workloadRows[rowIndex] || null;
    const clonedRow = templateClone.cloneNode(true);
    const values = row
      ? buildDetailValues(row)
      : Array(detailColumnCount).fill("");

    setWordRowValues(xmlDoc, clonedRow, values);
    targetTable.insertBefore(clonedRow, totalRow);
  }

  const totalCellCount = getDirectWordChildren(totalRow, "tc").length;
  const hasRows = workloadRows.length > 0;
  const valueOrBlank = (value) => hasRows ? value : "";

  // Footer PPL memiliki 9 sel langsung: satu sel gabungan + 8 nilai.
  // Footer PML memiliki 7 sel langsung: satu sel gabungan + 6 nilai.
  const totalValues = detailColumnCount === 16
    ? [
        "Jumlah",
        valueOrBlank(workloadTotal?.target_keluarga),
        valueOrBlank(workloadTotal?.target_usaha),
        valueOrBlank(workloadTotal?.target_jumlah),
        valueOrBlank(workloadTotal?.realisasi_keluarga),
        valueOrBlank(workloadTotal?.realisasi_usaha),
        valueOrBlank(workloadTotal?.realisasi_jumlah),
        valueOrBlank(workloadTotal?.persentase),
        "", // Footer Keterangan PPL dikosongkan.
      ]
    : [
        "Jumlah",
        valueOrBlank(workloadTotal?.target_keluarga),
        valueOrBlank(workloadTotal?.target_usaha),
        valueOrBlank(workloadTotal?.target_jumlah),
        valueOrBlank(workloadTotal?.realisasi_jumlah),
        valueOrBlank(workloadTotal?.persentase),
        "", // Footer Keterangan PML dikosongkan.
      ];

  setWordRowValues(xmlDoc, totalRow, totalValues.slice(0, totalCellCount));

  const serialized = new XMLSerializer().serializeToString(xmlDoc);
  zip.file("word/document.xml", serialized);
}

async function createBerkasPembayaranBlobFromTemplateBuffer(
  templateArrayBuffer,
  formValues,
  record,
  role,
  nikLookup
) {
  // Bangun data sebelum membuat Docxtemplater agar URL foto bisa di-prefetch.
  const templateData = buildBerkasPembayaranTemplateData(
    formValues || {},
    record || {},
    role,
    nikLookup
  );

  console.log("Generate berkas pembayaran:", {
    role,
    nama: templateData?.nama_petugas,
    jumlahFoto: templateData?.jumlah_foto,
    jumlahBarisFoto: templateData?.foto_rows?.length || 0,
  });

  // FIX PPL:
  // Download seluruh foto lebih dulu. Setelah tahap ini ImageModule tidak perlu
  // mengembalikan Promise saat Docxtemplater memproses loop gambar.
  await prefetchFotoBuktiForTemplate(templateData);

  const zip = new PizZip(templateArrayBuffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
    modules: [createFotoBuktiImageModule()],
  });

  // Jangan gunakan await doc.renderAsync(templateData) untuk image module gratis
  // pada loop foto. Foto sudah tersedia di cache sehingga render sinkron aman.
  doc.render(templateData);

  console.log("Jumlah foto:", templateData.jumlah_foto, templateData.foto_bukti);

  // Template tetap dipakai apa adanya. Kode hanya mengganti baris kosong pada tabel
  // Beban Kerja halaman terakhir dengan data dari sheet Data per SLS.
  fillBebanKerjaTableInDocxZip(
    doc.getZip(),
    templateData.beban_kerja_rows || [],
    templateData.beban_kerja_total || {}
  );

  return zipToDocxBlob(doc.getZip());
}

async function generateSingleBerkasPembayaran(templateUrl, formValues, record, role, nikLookup) {
  if (!record) throw new Error("Data berkas pembayaran tidak ditemukan.");
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template berkas pembayaran: ${response.status} ${response.statusText}`);
  const templateArrayBuffer = await response.arrayBuffer();
  const blob = await createBerkasPembayaranBlobFromTemplateBuffer(templateArrayBuffer, formValues, record, role, nikLookup); // ⬅️ tambah await
  saveAs(blob, `BERKAS PEMBAYARAN ${role} - ${sanitizeFileName(record.displayName)}.docx`);
}

const BERKAS_PEMBAYARAN_ZIP_BATCH_SIZE = 100;

async function generateBerkasPembayaran(templateUrl, formValues, records, role, nikLookup, onProgress) {
  const entries = Array.isArray(records) ? records.filter(Boolean) : [];
  if (entries.length === 0) throw new Error(`Tidak ada data berkas pembayaran ${role}.`);

  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template berkas pembayaran: ${response.status} ${response.statusText}`);
  const templateArrayBuffer = await response.arrayBuffer();
  const totalBatches = Math.ceil(entries.length / BERKAS_PEMBAYARAN_ZIP_BATCH_SIZE);

  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
    const batchEntries = entries.slice(
      batchIndex * BERKAS_PEMBAYARAN_ZIP_BATCH_SIZE,
      (batchIndex + 1) * BERKAS_PEMBAYARAN_ZIP_BATCH_SIZE
    );
    const files = [];
    const fileNameCounts = new Map();

    for (const record of batchEntries) {
      const blob = await createBerkasPembayaranBlobFromTemplateBuffer(   // ⬅️ tambah await
        templateArrayBuffer, formValues || {}, record, role, nikLookup
      );
      const baseName = sanitizeFileName(record.displayName);
      const count = (fileNameCounts.get(upperText(baseName)) || 0) + 1;
      fileNameCounts.set(upperText(baseName), count);
      const duplicateSuffix = count > 1 ? ` (${count})` : "";
      files.push({
        name: `BERKAS PEMBAYARAN ${role} - ${baseName}${duplicateSuffix}.docx`,
        blob,
      });
    }

    if (typeof onProgress === "function") {
      onProgress({ batchIndex: batchIndex + 1, totalBatches, totalRows: entries.length });
    }

    if (files.length === 1 && totalBatches === 1) {
      saveAs(files[0].blob, files[0].name);
      continue;
    }

    const batchSuffix = totalBatches > 1 ? ` - Bagian ${batchIndex + 1} dari ${totalBatches}` : "";
    await downloadMultipleAsZip(
      files,
      `BERKAS PEMBAYARAN ${role} ${cleanText(formValues?.tanggal_surat || "SE2026")}${batchSuffix}.zip`
    );
  }
}

function normalizePmlProgressKey(value) {
  return upperText(value).replace(/^@/, "").replace(/\s+/g, " ");
}

function findDataPmlProgressRow(dataPmlProgressData = [], nama = "", email = "") {
  const namaKey = normalizePmlProgressKey(nama);
  const emailLocal = normalizePmlProgressKey(email).includes("@")
    ? normalizePmlProgressKey(email).split("@")[0]
    : normalizePmlProgressKey(email);

  // Prioritas: cocokkan lewat Username Sobat (biasanya = bagian depan email), lalu Nama.
  let row = null;
  if (emailLocal) {
    row = dataPmlProgressData.find(
      (r) => normalizePmlProgressKey(r.username_sobat_pml) === emailLocal
    );
  }
  if (!row && namaKey) {
    row = dataPmlProgressData.find(
      (r) => normalizePmlProgressKey(r.nama_pml) === namaKey
    );
  }
  return row || null;
}

function buildSuratKepalaRows(
  selectionRows = [],
  bappData = [],
  dataPerSlsData = [],
  lampiranData = [],
  statusSlsData = [],
  approveByPmlData = [],
  dataPmlProgressData = []   // ⬅️ parameter baru: sheet "Data PML Progress"
) {
  const keySet = buildSelectionKeySet(selectionRows);
  const rows = [];
  const skipped = [];

  const pmlRecords = buildBerkasPembayaranRecords(
    bappData || [],
    lampiranData || [],
    "PML",
    statusSlsData || [],
    dataPerSlsData || [],
    approveByPmlData || []
  );

  for (const sel of selectionRows) {
    const bappRow = (bappData || []).find((r) => rowMatchesSelection(keySet, r.nama, r.email) &&
      (upperText(r.jabatan_raw || r.jabatan) === "PML" || upperText(r.jabatan_raw || r.jabatan) === "PPL") &&
      (cleanText(sel.email) ? upperText(cleanText(sel.email)) === upperText(cleanText(r.email)) : upperText(cleanText(sel.nama)) === upperText(cleanText(r.nama)))
    );

    if (!bappRow) {
      skipped.push({ nama: sel.nama, email: sel.email, alasan: "Tidak ditemukan di sheet Pembayaran (atau jabatan bukan PML/PPL)" });
      continue;
    }

    const jabatan = upperText(bappRow.jabatan_raw || bappRow.jabatan);
    let target = parseDataPerSlsNumber(bappRow.prelist_total) || 0;
    let realisasi = 0;

    if (jabatan === "PML") {
      // 🔥 DISAMAKAN dengan fitur Gabungan Administrasi Pembayaran:
      // realisasi PML diambil dari total kolom "Jumlah Approve PML" pada sheet
      // "Approve by PML" (via applyApproveByPmlToWorkload), BUKAN lagi dari sheet
      // "Data PML Progress". Target tetap dijumlahkan dari sheet "Data per SLS".
      const identity = getBerkasIdentity(cleanText(bappRow.nama), cleanText(bappRow.email));
      const record = pmlRecords.find((r) => r.identity === identity) ||
        pmlRecords.find((r) => upperText(r.displayName) === upperText(cleanText(bappRow.nama)));

      const dataPerSlsRows = record?.dataPerSlsRows || [];
      const approveByPmlRows = record?.approveByPmlRows || [];
      const bebanKerjaBase = buildDataPerSlsWorkloadRows(dataPerSlsRows, "PML");
      const bebanKerja = applyApproveByPmlToWorkload(bebanKerjaBase, approveByPmlRows, "PML");
      const hasApproveData = approveByPmlRows.length > 0;
      const hasDataPerSls = bebanKerja.rows.length > 0;

      if (hasApproveData || hasDataPerSls) {
        // total.target_jumlah = jumlah target dari Data per SLS
        // total.realisasi_jumlah = jumlah "Jumlah Approve PML" dari Approve by PML
        // (lihat applyApproveByPmlToWorkload -> directApproveTotal), sama persis
        // dengan yang dipakai pada Gabungan Administrasi Pembayaran PML.
        target = parseDataPerSlsNumber(bebanKerja.total.target_jumlah) || target;
        realisasi = parseDataPerSlsNumber(bebanKerja.total.realisasi_jumlah) || 0;
      } else {
        // Fallback terakhir kalau PML ini sama sekali tidak punya baris di
        // Data per SLS maupun Approve by PML.
        const progressRow = findDataPmlProgressRow(dataPmlProgressData, bappRow.nama, bappRow.email);
        if (progressRow) {
          realisasi = parseDataPerSlsNumber(progressRow.realisasi_jumlah) || 0;
          const progressTarget = parseDataPerSlsNumber(progressRow.jumlah_target);
          if (progressTarget) target = progressTarget;
        } else {
          realisasi = parseDataPerSlsNumber(bappRow.realisasi_total) || 0;
        }
      }
    } else {
      // PPL: tetap seperti sebelumnya, jumlahkan semua baris Data per SLS miliknya.
      const namaKey = upperText(bappRow.nama);
      const emailLocal = upperText(cleanText(bappRow.email)).split("@")[0];
      const matchedSlsRows = (dataPerSlsData || []).filter((r) => {
        const rowName = upperText(r.nama_ppl || "");
        const rowUsername = upperText(r.username_ppl || "");
        return (namaKey && rowName === namaKey) || (emailLocal && rowUsername === emailLocal);
      });
      if (matchedSlsRows.length === 0) {
        skipped.push({ nama: bappRow.nama, email: bappRow.email, alasan: "PPL tidak ditemukan di sheet Data per SLS" });
        continue;
      }
      realisasi = matchedSlsRows.reduce(
        (sum, r) => sum + (parseDataPerSlsNumber(r.realisasi_dengan_tidak_ditemukan_jumlah) || 0),
        0
      );
    }

    rows.push({
      nama: cleanText(bappRow.nama),
      jabatan,
      target,
      realisasi,
      persentase: target ? (realisasi / target) * 100 : 0,
    });
  }

  const pml = rows.filter((r) => r.jabatan === "PML")
    .sort((a, b) => a.nama.localeCompare(b.nama, "id-ID", { sensitivity: "base" }));
  const ppl = rows.filter((r) => r.jabatan === "PPL")
    .sort((a, b) => a.nama.localeCompare(b.nama, "id-ID", { sensitivity: "base" }));

  return { rows: [...pml, ...ppl], skipped };
}
 
function buildSuratKepalaTemplateData(rows = []) {
  const totalTarget = rows.reduce((s, r) => s + r.target, 0);
  const totalRealisasi = rows.reduce((s, r) => s + r.realisasi, 0);
  const totalPersentase = totalTarget ? (totalRealisasi / totalTarget) * 100 : 0;

  return {
    peserta: rows.map((r, idx) => ({
      no: idx + 1,                         // ✅ tag {no}
      nama: r.nama,
      nama_petugas: r.nama,                // ✅ tag {nama_petugas}
      jabatan: r.jabatan,
      target_prelist: formatRupiah(r.target),
      realisasi: formatRupiah(r.realisasi),
      persentase: `${formatPercentageNumber(r.persentase, 2)}%`,
    })),
    total_target_prelist: formatRupiah(totalTarget),
    total_realisasi: formatRupiah(totalRealisasi),
    total_persentase: `${formatPercentageNumber(totalPersentase, 2)}%`,
    average_persentase: `${formatPercentageNumber(totalPersentase, 2)}%`,  // ✅ tag {average_persentase}
  };
}
 
async function generateSuratKepala(templateUrl, rows) {
  if (!rows || rows.length === 0) throw new Error("Tidak ada data untuk Surat Kepala.");
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template Surat Kepala: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildSuratKepalaTemplateData(rows));
  const blob = doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
  saveAs(blob, `Surat Pernyataan Kepala BPS - ${rows.length} Petugas.docx`);
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