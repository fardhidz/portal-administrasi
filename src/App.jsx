// ============================================================
// Portal Administrasi SE2026 — BPS Kota Jakarta Timur
// Dependencies: npm install xlsx docxtemplater pizzip file-saver docx-preview
//
// Templates di public/templates/
//
// Data petugas .xlsx: public/data/data-petugas.xlsx
// ============================================================

import React, { useState, useRef, useCallback } from "react";
import { motion, AnimatePresence } from "framer-motion";
import {
  FileText, Users, ClipboardList, Car, Receipt, Briefcase,
  Map, ChevronRight, X, Printer, ArrowLeft, Check, Plus,
  Trash2, Menu, LayoutDashboard, Upload, Download, Filter,
  AlertCircle, CheckCircle, Clock, MapPin, Layers,
} from "lucide-react";
import * as XLSX from "xlsx";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import { saveAs } from "file-saver";
import { renderAsync } from "docx-preview";

// ─── DATA ────────────────────────────────────────────────────────────────────

const DOC_TYPES = [
  { id: "daftar-hadir",  icon: <ClipboardList />, label: "Daftar Hadir",                    desc: "Absensi kehadiran peserta pelatihan",          color: "orange" },
  { id: "tanda-terima",  icon: <Briefcase />,     label: "Tanda Terima",                    desc: "Bukti serah terima perlengkapan petugas",      color: "amber"  },
  { id: "surat-pernyataan-kendaraan", icon: <Car />, label: "Super Kendis",                    desc: "Surat pernyataan tidak menggunakan kendaraan dinas", color: "orange" },
  { id: "pengeluaran-riil", icon: <Receipt />,    label: "DPR",                               desc: "Rincian pengeluaran operasional petugas",      color: "amber"  },
  { id: "spj",           icon: <FileText />,      label: "SPJ",                               desc: "SPJ",            color: "orange" },
  { id: "spd",           icon: <Map />,           label: "SPD",                               desc: "Surat Perjalanan Dinas",                         color: "amber"  },
  { id: "surat-tugas",   icon: <Users />,         label: "Surtug",                            desc: "Surat tugas pelaksanaan kegiatan",             color: "orange" },
];

// ─── XLSX PARSER ─────────────────────────────────────────────────────────────

function parseXlsxData(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(sheet, { defval: "" });

  return raw.map((row) => {
    const normalized = {};
    Object.entries(row).forEach(([k, v]) => {
      normalized[k.trim()] = typeof v === "string" ? v.trim() : v;
    });
    return {
      no:         normalized["No"]          ?? normalized["NO"]          ?? "",
      nama:       normalized["Nama"]        ?? normalized["NAMA"]        ?? "",
      nik:        normalized["NIK"]         ?? normalized["Nik"]         ?? "",
      asal:       normalized["Asal"]        ?? normalized["ASAL"]        ?? "",
      wilTugas:   normalized["Wil. Tugas"]  ?? normalized["Wil.Tugas"]   ?? normalized["Wilayah Tugas"] ?? "",
      jabatan:    normalized["Jabatan"]     ?? normalized["JABATAN"]     ?? "",
      pangkatGol: normalized["Pangkat/Gol"] ?? normalized["Pangkat/gol"] ?? normalized["PangkatGol"] ?? "",
      kelas:      String(normalized["Kelas"]      ?? normalized["KELAS"]      ?? "").trim(),
      hotel:      normalized["Hotel"]       ?? normalized["HOTEL"]       ?? "",
      gelombang:  String(normalized["Gelombang"]  ?? normalized["GELOMBANG"]  ?? "").trim(),
    };
  }).filter(r => r.nama !== "");
}

// ─── HELPERS ─────────────────────────────────────────────────────────────────

function cleanText(value) {
  return String(value ?? "").trim();
}

function upperText(value) {
  return cleanText(value).toUpperCase();
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
  return d.toLocaleDateString("id-ID", {
    weekday: "long", day: "numeric", month: "long", year: "numeric",
  });
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

// ─── FILTER GROUPS ───────────────────────────────────────────────────────────

const DAFTAR_HADIR_GROUPS = {
  "pml-ppl":     { label: "PML & PPL",      roles: ["PML", "PPL"] },
  "panitia-inda": { label: "Panitia & Inda", roles: ["PANITIA", "INDA"] },
};

function jabatanMasukGroup(jabatan, groupKey) {
  const group = DAFTAR_HADIR_GROUPS[groupKey];
  if (!group) return false;
  return group.roles.includes(upperText(jabatan));
}

// ─── TEMPLATE URLS ───────────────────────────────────────────────────────────

const DAFTAR_HADIR_TEMPLATE_URL   = "/templates/1. Daftar Hadir Pelatihan SE2026.docx";
const TANDA_TERIMA_TEMPLATE_URL   = "/templates/2. Tanda Terima Perlengkapan SE2026.docx";
const SURAT_PERNYATAAN_KENDARAAN_TEMPLATE_URL = "/templates/3. Super Kendis Pelatihan SE2026.docx";
const PENGELUARAN_RIIL_TEMPLATE_URL = "/templates/4. DPR_Pelatihan SE 2026.docx";
const SPJ_TEMPLATE_URL = "/templates/5. SPJ Pelatihan_SE26.docx";
const SPD_TEMPLATE_URL = "/templates/6. SPD.docx";
const SPD_LAMPIRAN_TEMPLATE_URL = "/templates/6. Lampiran SPD.docx";
const SURAT_TUGAS_TEMPLATE_URL = "/templates/6. Surat Tugas.docx";

function formatRupiah(value) {
  if (value == null || value === "") return "";
  const number = Number(String(value).replace(/[^0-9-]/g, ""));
  if (Number.isNaN(number)) return String(value);
  return number.toLocaleString("id-ID");
}

function spellTerbilang(value) {
  const units = ["", "satu", "dua", "tiga", "empat", "lima", "enam", "tujuh", "delapan", "sembilan"];
  const teens = ["sepuluh", "sebelas", "dua belas", "tiga belas", "empat belas", "lima belas", "enam belas", "tujuh belas", "delapan belas", "sembilan belas"];

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
  if (Number.isNaN(number) || number === 0) return "nol";
  return toWords(number);
}

function buildSpjTemplateData(formValues, peserta = []) {
  const jumlahOk = peserta.length;
  const jumlahUang = 170000;
  const totalJumlahKotor = jumlahOk * jumlahUang;
  const totalJumlahBersih = totalJumlahKotor;

  return {
    tanggal_pelunasan: formatTanggalIndonesia(formValues.tanggal_pelunasan),
    tempat: formValues.tempat || formValues.hotel || "",
    jumlah_ok: jumlahOk || 0,
    jumlah_uang: formatRupiah(jumlahUang),
    total_jumlah_kotor: formatRupiah(totalJumlahKotor),
    total_jumlah_bersih: formatRupiah(totalJumlahBersih),
    total_terbilang: `${spellTerbilang(totalJumlahBersih)} rupiah`,
    ttd_kiri: formValues.ttd_kiri || "",
    ttd_kanan: formValues.ttd_kanan || "",
    peserta: peserta.map((p, idx) => ({ no: idx + 1, nama: p.nama || "" })),
  };
}

async function createSpjBlob(templateUrl, formValues, peserta) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildSpjTemplateData(formValues || {}, peserta || []));
  return doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

async function generateSpj(templateUrl, formValues, peserta) {
  const blob = await createSpjBlob(templateUrl, formValues || {}, peserta || []);
  saveAs(blob, `SPJ ${formValues.tempat || "SE2026"} ${formValues.tanggal_pelunasan || ""}.docx`);
}

function buildSpdTemplateData(formValues, peserta = []) {
  const tanggalAwal = formatTanggalIndonesia(formValues.tanggal_awal_kegiatan);
  const tanggalAkhir = formatTanggalIndonesia(formValues.tanggal_akhir_kegiatan);
  const lamaHari = calcDurationDays(formValues.tanggal_awal_kegiatan, formValues.tanggal_akhir_kegiatan);

  return {
    nomor_dokumen: formValues.nomor || "",
    tanggal_surat: formatTanggalIndonesia(formValues.tanggal),
    kegiatan: formValues.kegiatan || "",
    lokasi: formValues.lokasi || formValues.tempat || formValues.hotel || "",
    tempat: formValues.tempat || formValues.hotel || "",
    tanggal_awal_kegiatan: tanggalAwal,
    tanggal_akhir_kegiatan: tanggalAkhir,
    lama_hari: lamaHari,
    jumlah_peserta: peserta.length || 0,
    ttd_nama: formValues.namaKabps || "",
    ttd_nip: formValues.nipKabps || "",
    peserta: peserta.map((p, idx) => ({
      no: idx + 1,
      nama: cleanText(p.nama),
      nik: cleanText(p.nik),
      jabatan: cleanText(p.jabatan),
      tanggal_awal_kegiatan: tanggalAwal,
      tanggal_akhir_kegiatan: tanggalAkhir,
      lama: lamaHari,
    })),
  };
}

async function createSpdBlob(templateUrl, formValues, peserta) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildSpdTemplateData(formValues || {}, peserta || []));
  return doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

async function generateSpd(mainTemplateUrl, attachmentTemplateUrl, formValues, peserta) {
  const blobMain = await createSpdBlob(mainTemplateUrl, formValues || {}, peserta || []);
  const blobAttachment = await createSpdBlob(attachmentTemplateUrl, formValues || {}, peserta || []);
  saveAs(blobMain, `SPD ${formValues.tempat || "SE2026"} ${formValues.tanggal || ""}.docx`);
  saveAs(blobAttachment, `Lampiran SPD ${formValues.tempat || "SE2026"} ${formValues.tanggal || ""}.docx`);
}

function buildSuratTugasTemplateData(formValues, peserta = []) {
  const tanggalSurat = formatTanggalIndonesia(formValues.tanggal_surat);
  const tanggalAwal = formatTanggalIndonesia(formValues.tanggal_awal_kegiatan);
  const tanggalAkhir = formatTanggalIndonesia(formValues.tanggal_akhir_kegiatan);
  const lama = calcDurationDays(formValues.tanggal_awal_kegiatan, formValues.tanggal_akhir_kegiatan);

  return {
    nomor_surat: formValues.nomor_surat || "",
    tanggal_surat: tanggalSurat,
    tanggal_awal_kegiatan: tanggalAwal,
    tanggal_akhir_kegiatan: tanggalAkhir,
    lama: lama || "",
    peserta: (peserta || []).map((p, idx) => ({
      no: idx + 1,
      nama: cleanText(p.nama),
      nik: cleanText(p.nik),
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
  return doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

async function generateSuratTugas(templateUrl, formValues, peserta) {
  const blob = await createSuratTugasBlob(templateUrl, formValues || {}, peserta || []);
  saveAs(blob, `Surat Tugas ${formValues.nomor_surat || formValues.tempat || "SE2026"}.docx`);
}

// ─── DAFTAR HADIR — DOCX GENERATOR ──────────────────────────────────────────

function buildDaftarHadirTemplateData(formValues, peserta, namaInda, selectedFilterGroup = "") {
  const jamMulai    = normalizeJamIndonesia(formValues.jamMulai,    "07.00");
  const jamSelesai  = normalizeJamIndonesia(formValues.jamSelesai,  "17.00");
  const tanggalFormatted = formatTanggalIndonesia(formValues.tanggal);
  const jamFormatted     = `${jamMulai} - ${jamSelesai}`;

  const isPanitiaInda     = selectedFilterGroup === "panitia-inda";
  const isPmlPpl          = selectedFilterGroup === "pml-ppl";
  const effectiveNamaInda = isPanitiaInda
    ? "Ir. Tristiati, MA"
    : (namaInda || "");
  const keteranganTtd     = isPanitiaInda
    ? "Kepala Sub Bagian Umum"
    : (isPmlPpl ? "Instruktur Daerah" : "");

  return {
    tanggal_kegiatan: tanggalFormatted,
    hari_tanggal:     tanggalFormatted,
    tanggal:          tanggalFormatted,
    jam_mulai:        jamMulai,
    jam_selesai:      jamSelesai,
    jam:              jamFormatted,
    jam_kegiatan:     jamFormatted,
    tempat:           formValues.tempat || formValues.hotel || "",
    tempat_kegiatan:  formValues.tempat || formValues.hotel || "",
    gelombang:        formValues.gelombang || "",
    kelas:            formValues.kelas     || "",
    nama_inda:        effectiveNamaInda,
    keterangan_ttd:   keteranganTtd,
    peserta: (peserta || []).map((p, idx) => ({
      no:          idx + 1,
      nama:        p.nama     || "",
      jabatan:     p.jabatan  || "",
      kecamatan:   p.wilTugas || "",
      wil_tugas:   p.wilTugas || "",
      wilTugas:    p.wilTugas || "",
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
  return doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

async function generateDaftarHadir(templateUrl, formValues, peserta, namaInda, selectedFilterGroup = "") {
  const blob = await createDaftarHadirBlob(templateUrl, formValues, peserta, namaInda, selectedFilterGroup);
  const safeGelombang = formValues?.gelombang || "X";
  const safeKelas     = formValues?.kelas     || "X";
  saveAs(blob, `Daftar Hadir Gelombang ${safeGelombang} Kelas ${safeKelas}.docx`);
}

// ─── TANDA TERIMA — DOCX GENERATOR ──────────────────────────────────────────
function buildTandaTerimaTemplateData(formValues, peserta) {
  const tanggalFormatted = formatTanggalIndonesia(formValues.tanggal);

  // Hanya PML dan PPL
  const filteredPeserta = (peserta || []).filter(
    (p) => ["PML", "PPL"].includes(upperText(p.jabatan))
  );

  return {
    tanggal_kegiatan: tanggalFormatted,
    tanggal:          tanggalFormatted,
    tempat:           formValues.tempat || formValues.hotel || "",
    gelombang:        formValues.gelombang || "",
    kelas:            formValues.kelas     || "",

    peserta: filteredPeserta.map((p, idx) => ({
      no:        idx + 1,
      nama:      p.nama     || "",
      jabatan:   p.jabatan  || "",
      kecamatan: p.wilTugas || "",
      wil_tugas: p.wilTugas || "",
      wilTugas:  p.wilTugas || "",
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
  return doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

async function generateTandaTerima(templateUrl, formValues, peserta) {
  const blob = await createTandaTerimaBlob(templateUrl, formValues, peserta);
  const safeGelombang = formValues?.gelombang || "X";
  const safeKelas     = formValues?.kelas     || "X";
  saveAs(blob, `Tanda Terima Perlengkapan Gelombang ${safeGelombang} Kelas ${safeKelas}.docx`);
}

function buildSuratPernyataanKendaraanTemplateData(formValues, peserta) {
  const tanggalFormatted = formatTanggalIndonesia(formValues.tanggal);

  const filteredPeserta = (peserta || []).filter((p) =>
    ["PML", "PPL"].includes(upperText(p.jabatan))
  );

  return {
    tanggal_surat: tanggalFormatted,
    tanggal:       tanggalFormatted,
    tempat:        formValues.tempat || formValues.hotel || "",
    gelombang:     formValues.gelombang || "",
    kelas:         formValues.kelas || "",
    peserta:       filteredPeserta.map((p, idx) => ({
      no:        idx + 1,
      nama:      p.nama     || "",
      jabatan:   p.jabatan  || "",
      kecamatan: p.wilTugas || "",
      wil_tugas: p.wilTugas || "",
      wilTugas:  p.wilTugas || "",
    })),
  };
}

async function createSuratPernyataanKendaraanBlob(templateUrl, formValues, peserta) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildSuratPernyataanKendaraanTemplateData(formValues || {}, peserta || []));
  return doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

async function generateSuratPernyataanKendaraan(templateUrl, formValues, peserta) {
  const blob = await createSuratPernyataanKendaraanBlob(templateUrl, formValues, peserta);
  const safeGelombang = formValues?.gelombang || "X";
  const safeKelas     = formValues?.kelas     || "X";
  saveAs(blob, `Surat Pernyataan Kendaraan Gelombang ${safeGelombang} Kelas ${safeKelas}.docx`);
}

function buildPengeluaranRiilTemplateData(formValues, peserta = []) {
  return {
    tanggal_surat: formatTanggalIndonesia(formValues.tanggal_surat),
    peserta: peserta.map((p, idx) => ({
      no:      idx + 1,
      nama:    cleanText(p.nama),
      nik:     cleanText(p.nik),
      jabatan: cleanText(p.jabatan),
    })),
  };
}

async function createPengeluaranRiilBlob(templateUrl, formValues, peserta) {
  const response = await fetch(templateUrl);
  if (!response.ok) throw new Error(`Gagal memuat template: ${response.status} ${response.statusText}`);
  const arrayBuffer = await response.arrayBuffer();
  const zip = new PizZip(arrayBuffer);
  const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
  doc.render(buildPengeluaranRiilTemplateData(formValues || {}, peserta || []));
  return doc.getZip().generate({
    type: "blob",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  });
}

async function generatePengeluaranRiil(templateUrl, formValues, peserta) {
  const blob = await createPengeluaranRiilBlob(templateUrl, formValues, peserta);
  saveAs(blob, `Daftar Pengeluaran Riil ${formValues.no || "SE2026"}.docx`);
}

// ─── MAIN APP ─────────────────────────────────────────────────────────────────

export default function PortalAdministrasiSE2026() {
  const [view,         setView]         = useState("dashboard");
  const [selectedDoc,  setSelectedDoc]  = useState(null);
  const [formData,     setFormData]     = useState({});
  const [previewData,  setPreviewData]  = useState(null);
  const [sidebarOpen,  setSidebarOpen]  = useState(false);

  const [petugasData,  setPetugasData]  = useState([]);
  const [xlsxLoaded,   setXlsxLoaded]  = useState(false);
  const [xlsxFileName, setXlsxFileName] = useState("data-petugas.xlsx");

  React.useEffect(() => {
    const loadXlsxData = async () => {
      try {
        const response = await fetch("/data/data-petugas.xlsx");
        if (!response.ok) throw new Error(`HTTP ${response.status}`);
        const arrayBuffer = await response.arrayBuffer();
        const data = parseXlsxData(arrayBuffer);
        setPetugasData(data);
        setXlsxLoaded(true);
      } catch (err) {
        console.warn("Tidak dapat memuat data-petugas.xlsx: " + err.message);
      }
    };
    loadXlsxData();
  }, []);

  const openForm = (docType) => {
    setSelectedDoc(docType);
    setFormData({});
    setPreviewData(null);
    setView("form");
  };

  const handlePreview = (data) => {
    setPreviewData(data);
    setView("preview");
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
        const data = parseXlsxData(e.target.result);
        setPetugasData(data);
        setXlsxLoaded(true);
        setXlsxFileName(file.name);
      } catch (err) {
        alert("Gagal membaca file xlsx: " + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }, []);

  return (
    <div className="min-h-screen overflow-x-hidden bg-[#fff8f0] text-slate-950 selection:bg-orange-200 selection:text-orange-950">
      {/* Background blobs */}
      <div className="pointer-events-none fixed inset-0 z-0 opacity-70">
        <div className="absolute left-[-12rem] top-[-10rem] h-[28rem] w-[28rem] rounded-full bg-orange-300/40 blur-3xl" />
        <div className="absolute right-[-10rem] top-[12rem] h-[30rem] w-[30rem] rounded-full bg-amber-300/30 blur-3xl" />
        <div className="absolute bottom-[-12rem] left-1/2 h-[30rem] w-[30rem] -translate-x-1/2 rounded-full bg-orange-400/20 blur-3xl" />
      </div>

      {/* Header */}
      <header className="fixed left-0 right-0 top-0 z-50 border-b border-orange-100/80 bg-white/75 backdrop-blur-2xl">
        <div className="mx-auto flex max-w-7xl items-center justify-between px-5 py-3.5 lg:px-8">
          <div className="flex items-center gap-3">
            <button
              className="rounded-xl border border-orange-100 bg-white p-2 text-slate-700 shadow-sm lg:hidden"
              onClick={() => setSidebarOpen(!sidebarOpen)}
            >
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

          <nav className="hidden items-center gap-1 rounded-full border border-orange-100 bg-white/80 p-1 shadow-sm lg:flex">
            <button
              onClick={() => { setView("dashboard"); setSelectedDoc(null); }}
              className={`rounded-full px-4 py-2 text-sm font-bold transition ${view === "dashboard" ? "bg-orange-500 text-white shadow" : "text-slate-600 hover:bg-orange-50 hover:text-orange-600"}`}
            >
              Dashboard
            </button>
            {DOC_TYPES.map((d) => (
              <button
                key={d.id}
                onClick={() => openForm(d)}
                className={`rounded-full px-3 py-2 text-xs font-bold transition ${selectedDoc?.id === d.id ? "bg-orange-500 text-white shadow" : "text-slate-600 hover:bg-orange-50 hover:text-orange-600"}`}
              >
                {d.label}
              </button>
            ))}
          </nav>

          {xlsxLoaded && (
            <div className="hidden items-center gap-2 rounded-full border border-green-200 bg-green-50 px-3 py-1.5 lg:flex">
              <CheckCircle size={14} className="text-green-600" />
              <span className="text-xs font-bold text-green-700">{xlsxFileName}</span>
            </div>
          )}
        </div>
      </header>

      {/* Mobile Sidebar */}
      <AnimatePresence>
        {sidebarOpen && (
          <>
            <motion.div
              initial={{ opacity: 0 }} animate={{ opacity: 1 }} exit={{ opacity: 0 }}
              className="fixed inset-0 z-40 bg-black/30 backdrop-blur-sm lg:hidden"
              onClick={() => setSidebarOpen(false)}
            />
            <motion.div
              initial={{ x: -280 }} animate={{ x: 0 }} exit={{ x: -280 }}
              transition={{ type: "spring", damping: 28, stiffness: 300 }}
              className="fixed left-0 top-0 z-50 flex h-full w-72 flex-col border-r border-orange-100 bg-white shadow-2xl lg:hidden"
            >
              <div className="flex items-center justify-between border-b border-orange-100 p-5">
                <p className="font-black text-slate-950">Menu Dokumen</p>
                <button onClick={() => setSidebarOpen(false)} className="rounded-xl bg-orange-50 p-2 text-orange-600">
                  <X size={18} />
                </button>
              </div>
              <div className="flex-1 overflow-y-auto p-4">
                <button
                  onClick={() => { setView("dashboard"); setSelectedDoc(null); setSidebarOpen(false); }}
                  className={`mb-2 flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-sm font-bold transition ${view === "dashboard" && !selectedDoc ? "bg-orange-500 text-white" : "text-slate-700 hover:bg-orange-50"}`}
                >
                  <LayoutDashboard size={18} /> Dashboard
                </button>
                {DOC_TYPES.map((d) => (
                  <button
                    key={d.id}
                    onClick={() => { openForm(d); setSidebarOpen(false); }}
                    className={`mb-2 flex w-full items-center gap-3 rounded-2xl px-4 py-3 text-sm font-bold transition ${selectedDoc?.id === d.id ? "bg-orange-500 text-white" : "text-slate-700 hover:bg-orange-50"}`}
                  >
                    {React.cloneElement(d.icon, { size: 18 })} {d.label}
                  </button>
                ))}
              </div>
            </motion.div>
          </>
        )}
      </AnimatePresence>

      {/* Main Content */}
      <main className="relative z-10 pt-20">
        <AnimatePresence mode="wait">

          {/* ── DASHBOARD ── */}
          {view === "dashboard" && (
            <motion.div
              key="dashboard"
              initial={{ opacity: 0, y: 18 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }}
              transition={{ duration: 0.4 }}
              className="px-5 py-10 lg:px-8"
            >
              <div className="mx-auto max-w-7xl">
                <div className="mb-10">
                  <p className="text-sm font-black uppercase tracking-[0.25em] text-orange-600">Sensus Ekonomi 2026</p>
                  <h1 className="mt-3 text-5xl font-black tracking-[-0.04em] text-slate-950 lg:text-6xl">
                    Portal Administrasi<br />
                    <span className="bg-gradient-to-r from-orange-600 via-orange-500 to-amber-500 bg-clip-text text-transparent">
                      Pelatihan Petugas
                    </span>
                  </h1>
                  <p className="mt-5 max-w-2xl text-base leading-8 text-slate-600">
                    Terbitkan dokumen administrasi secara cepat dan terstandar.
                  </p>
                </div>

                {!xlsxLoaded && (
                  <XlsxUploadCard
                    loaded={xlsxLoaded}
                    fileName={xlsxFileName}
                    petugasCount={petugasData.length}
                    onUpload={handleXlsxUpload}
                  />
                )}

                {xlsxLoaded && (
                  <motion.div
                    initial={{ opacity: 0, y: -10 }}
                    animate={{ opacity: 1, y: 0 }}
                    className="mb-8 flex items-center gap-4 rounded-3xl border border-blue-200 bg-blue-50 px-6 py-4 shadow-sm"
                  >
                    <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-blue-500 shadow-lg shadow-blue-500/25">
                      <CheckCircle size={22} className="text-white" />
                    </div>
                    <div>
                      <p className="font-black text-blue-900">✓ Data Petugas Sudah Terbaca</p>
                      <p className="text-sm font-semibold text-blue-600">{petugasData.length} petugas siap digunakan untuk filter gelombang dan kelas</p>
                    </div>
                  </motion.div>
                )}

                <div className="mb-10 grid grid-cols-3 gap-4">
                  <StatCard value="7"      label="Jenis Dokumen" />
                  <StatCard value="SE2026" label="Kegiatan" />
                  <StatCard value={petugasData.length || "—"} label="Data Petugas" highlight={petugasData.length > 0} />
                </div>

                <div className="grid gap-5 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4">
                  {DOC_TYPES.map((doc, i) => (
                    <DocCard key={doc.id} doc={doc} index={i} onSelect={() => openForm(doc)} />
                  ))}
                </div>
              </div>
            </motion.div>
          )}

          {/* ── FORM ── */}
          {view === "form" && selectedDoc && (
            <motion.div
              key="form"
              initial={{ opacity: 0, y: 18 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }}
              transition={{ duration: 0.4 }}
              className="px-5 py-10 lg:px-8"
            >
              <div className="mx-auto max-w-3xl">
                <button
                  onClick={handleBack}
                  className="mb-6 inline-flex items-center gap-2 rounded-2xl border border-orange-100 bg-white/80 px-4 py-2 text-sm font-bold text-slate-700 shadow-sm backdrop-blur transition hover:border-orange-200 hover:bg-white"
                >
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
                <DocForm
                  docType={selectedDoc}
                  formData={formData}
                  setFormData={setFormData}
                  onPreview={handlePreview}
                  petugasData={petugasData}
                  xlsxLoaded={xlsxLoaded}
                />
              </div>
            </motion.div>
          )}

          {/* ── PREVIEW ── */}
          {view === "preview" && previewData && (
            <motion.div
              key="preview"
              initial={{ opacity: 0, y: 18 }} animate={{ opacity: 1, y: 0 }} exit={{ opacity: 0, y: -10 }}
              transition={{ duration: 0.4 }}
              className="px-5 py-10 lg:px-8"
            >
              <div className="mx-auto max-w-4xl">
                <div className="mb-6 flex items-center justify-between">
                  <button
                    onClick={handleBack}
                    className="inline-flex items-center gap-2 rounded-2xl border border-orange-100 bg-white/80 px-4 py-2 text-sm font-bold text-slate-700 shadow-sm backdrop-blur transition hover:border-orange-200 hover:bg-white"
                  >
                    <ArrowLeft size={16} /> Kembali ke Formulir
                  </button>
                  <div className="flex gap-3">
                    <button
                      onClick={() => window.print()}
                      className="inline-flex items-center gap-2 rounded-2xl border border-orange-200 bg-white px-5 py-2.5 text-sm font-black text-orange-600 shadow-sm transition hover:-translate-y-0.5 hover:bg-orange-50"
                    >
                      <Printer size={16} /> Cetak
                    </button>

                    {/* Tombol unduh .docx per jenis dokumen */}
                    {selectedDoc?.id === "daftar-hadir" && (
                      <GenerateDocxButton
                        onGenerate={() =>
                          generateDaftarHadir(
                            DAFTAR_HADIR_TEMPLATE_URL,
                            previewData.formValues,
                            previewData.peserta,
                            previewData.namaInda,
                            previewData.selectedFilterGroup,
                          )
                        }
                      />
                    )}
                    {selectedDoc?.id === "tanda-terima" && (
                      <GenerateDocxButton
                        onGenerate={() =>
                          generateTandaTerima(
                            TANDA_TERIMA_TEMPLATE_URL,
                            previewData.formValues,
                            previewData.peserta,
                          )
                        }
                      />
                    )}
                    {selectedDoc?.id === "surat-pernyataan-kendaraan" && (
                      <GenerateDocxButton
                        onGenerate={() =>
                          generateSuratPernyataanKendaraan(
                            SURAT_PERNYATAAN_KENDARAAN_TEMPLATE_URL,
                            previewData.formValues,
                            previewData.peserta,
                          )
                        }
                      />
                    )}
                    {selectedDoc?.id === "pengeluaran-riil" && (
                      <GenerateDocxButton
                        onGenerate={() =>
                          generatePengeluaranRiil(
                            PENGELUARAN_RIIL_TEMPLATE_URL,
                            previewData.formValues,
                            previewData.peserta,
                          )
                        }
                      />
                    )}
                    {selectedDoc?.id === "spj" && (
                      <GenerateDocxButton
                        onGenerate={() =>
                          generateSpj(
                            SPJ_TEMPLATE_URL,
                            previewData.formValues,
                            previewData.peserta,
                          )
                        }
                      />
                    )}
                    {selectedDoc?.id === "spd" && (
                      <GenerateDocxButton
                        onGenerate={() =>
                          generateSpd(
                            SPD_TEMPLATE_URL,
                            SPD_LAMPIRAN_TEMPLATE_URL,
                            previewData.formValues,
                            previewData.peserta,
                          )
                        }
                      />
                    )}
                    {selectedDoc?.id === "surat-tugas" && (
                      <GenerateDocxButton
                        onGenerate={() =>
                          generateSuratTugas(
                            SURAT_TUGAS_TEMPLATE_URL,
                            previewData.formValues,
                            previewData.peserta,
                          )
                        }
                      />
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
    e.preventDefault();
    setDragging(false);
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith(".xlsx")) onUpload(file);
  };

  return (
    <div className="mb-8">
      {loaded ? (
        <motion.div
          initial={{ opacity: 0, scale: 0.97 }}
          animate={{ opacity: 1, scale: 1 }}
          className="flex items-center justify-between rounded-3xl border border-green-200 bg-green-50 px-6 py-4 shadow-sm"
        >
          <div className="flex items-center gap-4">
            <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-green-500 shadow-lg shadow-green-500/25">
              <CheckCircle size={22} className="text-white" />
            </div>
            <div>
              <p className="font-black text-green-800">Data Petugas Berhasil Dimuat</p>
              <p className="text-sm font-semibold text-green-600">{fileName} — {petugasCount} petugas</p>
            </div>
          </div>
          <button
            onClick={() => fileRef.current?.click()}
            className="rounded-2xl border border-green-200 bg-white px-4 py-2 text-sm font-bold text-green-700 transition hover:bg-green-50"
          >
            Ganti File
          </button>
          <input ref={fileRef} type="file" accept=".xlsx" className="hidden" onChange={(e) => onUpload(e.target.files[0])} />
        </motion.div>
      ) : (
        <div
          onDragOver={(e) => { e.preventDefault(); setDragging(true); }}
          onDragLeave={() => setDragging(false)}
          onDrop={handleDrop}
          onClick={() => fileRef.current?.click()}
          className={`cursor-pointer rounded-3xl border-2 border-dashed p-8 text-center transition ${dragging ? "border-orange-400 bg-orange-50" : "border-orange-200 bg-white/60 hover:border-orange-300 hover:bg-orange-50/50"}`}
        >
          <div className="mx-auto mb-4 flex h-14 w-14 items-center justify-center rounded-2xl bg-orange-100">
            <Upload size={24} className="text-orange-600" />
          </div>
          <p className="font-black text-slate-800">Unggah Data Petugas (Opsional)</p>
          <p className="mt-1 text-sm font-semibold text-slate-500">
            Data <span className="text-orange-600">data-petugas.xlsx</span> sudah dimuat dari folder publik. Drag & drop di sini untuk mengganti.
          </p>
          <p className="mt-2 text-xs text-slate-400">Kolom: Nama, NIK, Jabatan, Wil. Tugas, Pangkat/Gol, Kelas, Gelombang</p>
          <input ref={fileRef} type="file" accept=".xlsx" className="hidden" onChange={(e) => onUpload(e.target.files[0])} />
        </div>
      )}
    </div>
  );
}

// ─── GENERATE DOCX BUTTON ────────────────────────────────────────────────────
// Generik: terima fungsi onGenerate (async) sebagai prop

function GenerateDocxButton({ onGenerate }) {
  const [loading, setLoading] = useState(false);
  const [error,   setError]   = useState(null);

  const handleGenerate = async () => {
    setLoading(true);
    setError(null);
    try {
      await onGenerate();
    } catch (err) {
      setError(err.message);
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="flex flex-col items-end gap-1">
      <button
        onClick={handleGenerate}
        disabled={loading}
        className="inline-flex items-center gap-2 rounded-2xl bg-orange-500 px-5 py-2.5 text-sm font-black text-white shadow-xl shadow-orange-500/20 transition hover:-translate-y-0.5 hover:bg-orange-600 disabled:opacity-60"
      >
        <Download size={16} />
        {loading ? "Membuat..." : "Unduh .docx"}
      </button>
      {error && (
        <p className="flex items-center gap-1 text-xs font-semibold text-red-500">
          <AlertCircle size={12} /> {error}
        </p>
      )}
    </div>
  );
}

// ─── SUB-COMPONENTS ───────────────────────────────────────────────────────────

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
    <motion.button
      initial={{ opacity: 0, y: 24 }}
      animate={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.5, delay: index * 0.07 }}
      onClick={onSelect}
      className="group relative overflow-hidden rounded-[2rem] border border-orange-100 bg-white/85 p-6 text-left shadow-lg shadow-orange-900/5 backdrop-blur transition hover:-translate-y-2 hover:shadow-2xl hover:shadow-orange-500/15"
    >
      <div className="absolute right-0 top-0 h-24 w-24 rounded-bl-[4rem] bg-orange-50 transition group-hover:bg-orange-100" />
      <div className="relative flex h-14 w-14 items-center justify-center rounded-2xl bg-orange-500 text-white shadow-xl shadow-orange-500/25">
        {React.cloneElement(doc.icon, { size: 26 })}
      </div>
      <h3 className="relative mt-5 text-lg font-black tracking-tight text-slate-950">{doc.label}</h3>
      <p className="relative mt-2 text-sm leading-6 text-slate-500">{doc.desc}</p>
      <div className="relative mt-5 flex items-center gap-1 text-sm font-black text-orange-500">
        Buat Dokumen <ChevronRight size={16} className="transition group-hover:translate-x-1" />
      </div>
    </motion.button>
  );
}

// ─── FILTER PESERTA PANEL (reusable) ─────────────────────────────────────────
// Dipakai oleh Daftar Hadir dan Tanda Terima

function FilterPesertaPanel({
  xlsxLoaded,
  formData,
  setFormData,
  petugasData,
  // mode: "grouped" (daftar hadir pakai tombol PML/PPL vs Panitia)
  //       "all"     (tanda terima ambil semua jabatan sesuai filter hotel/kelas/gelombang)
  mode = "grouped",
  onFilterResult,   // callback(peserta[], namaInda, selectedFilterGroup)
}) {
  const inputCls = "w-full rounded-2xl border border-orange-100 bg-white/80 px-4 py-3 text-sm font-semibold text-slate-800 shadow-sm outline-none focus:border-orange-300 focus:ring-2 focus:ring-orange-100 transition";
  const labelCls = "mb-2 block text-xs font-black uppercase tracking-[0.2em] text-slate-500";

  const update = (key, val) => setFormData((p) => ({ ...p, [key]: val }));

  const [filtered,             setFiltered]             = useState(false);
  const [filteredPeserta,      setFilteredPeserta]      = useState([]);
  const [namaInda,             setNamaInda]             = useState("");
  const [selectedFilterGroup,  setSelectedFilterGroup]  = useState("");

  const selectedHotel     = cleanText(formData.hotel);
  const selectedKelas     = cleanText(formData.kelas);
  const selectedGelombang = cleanText(formData.gelombang);

  const hotelOptions = React.useMemo(
    () => uniqueSorted(petugasData.map((p) => p.hotel)),
    [petugasData]
  );

  const gelombangOptions = React.useMemo(() =>
    uniqueSorted(
      petugasData
        .filter((p) => !selectedHotel || cleanText(p.hotel) === selectedHotel)
        .map((p) => p.gelombang)
    ),
    [petugasData, selectedHotel]
  );

  const kelasOptions = React.useMemo(() =>
    uniqueSorted(
      petugasData
        .filter((p) => !selectedHotel || cleanText(p.hotel) === selectedHotel)
        .filter((p) => !selectedGelombang || cleanText(p.gelombang) === selectedGelombang)
        .map((p) => p.kelas)
    ),
    [petugasData, selectedHotel, selectedGelombang]
  );

  const resetFilterResult = () => {
    setFilteredPeserta([]);
    setNamaInda("");
    setFiltered(false);
    setSelectedFilterGroup("");
    onFilterResult([], "", "");
  };

  const runFilter = (groupKey) => {
    if (!xlsxLoaded) { alert("Data petugas (.xlsx) belum berhasil dimuat."); return; }
    if (!selectedHotel || !selectedKelas || !selectedGelombang) {
      alert("Pilih Hotel, Kelas, dan Gelombang terlebih dahulu.");
      return;
    }

    const baseRows = petugasData.filter((p) =>
      cleanText(p.hotel)     === selectedHotel &&
      cleanText(p.kelas)     === selectedKelas &&
      cleanText(p.gelombang) === selectedGelombang
    );

    const inda = baseRows.find((p) => upperText(p.jabatan) === "INDA");

    let hasil;
    if (mode === "all") {
      // Tanda Terima: hanya ambil peserta PML dan PPL
      hasil = baseRows.filter((p) => ["PML", "PPL"].includes(upperText(p.jabatan)));
    } else {
      hasil = baseRows.filter((p) => jabatanMasukGroup(p.jabatan, groupKey));
    }

    setNamaInda(inda?.nama || "");
    setFilteredPeserta(hasil);
    setSelectedFilterGroup(groupKey);
    setFiltered(true);
    onFilterResult(hasil, inda?.nama || "", groupKey);
  };

  return (
    <div className="rounded-3xl border border-orange-100 bg-orange-50/60 p-5 space-y-4">
      <p className="text-xs font-black uppercase tracking-[0.2em] text-orange-700 flex items-center gap-2">
        <Filter size={14} /> Parameter Filter Peserta
      </p>

      {/* Hotel */}
      <div>
        <label className={labelCls}>Tempat</label>
        <select
          className={inputCls}
          value={formData.hotel || ""}
          onChange={(e) => {
            const hotel = e.target.value;
            setFormData((prev) => ({ ...prev, hotel, tempat: hotel, kelas: "", gelombang: "" }));
            resetFilterResult();
          }}
          disabled={!xlsxLoaded || hotelOptions.length === 0}
        >
          <option value="">Pilih hotel</option>
          {hotelOptions.map((hotel) => (
            <option key={hotel} value={hotel}>{hotel}</option>
          ))}
        </select>
        <p className="mt-1 text-xs font-semibold text-slate-400">Daftar hotel diambil otomatis dari kolom Hotel pada data XLSX.</p>
      </div>

      {/* Gelombang + Kelas */}
      <div className="grid gap-4 sm:grid-cols-2">
        <div>
          <label className={labelCls}>Gelombang</label>
          <select
            className={inputCls}
            value={formData.gelombang || ""}
            onChange={(e) => {
              setFormData((prev) => ({ ...prev, gelombang: e.target.value, kelas: "" }));
              resetFilterResult();
            }}
            disabled={!selectedHotel}
          >
            <option value="">Pilih gelombang</option>
            {gelombangOptions.map((g) => <option key={g} value={g}>{g}</option>)}
          </select>
        </div>
        <div>
          <label className={labelCls}>Kelas</label>
          <select
            className={inputCls}
            value={formData.kelas || ""}
            onChange={(e) => {
              setFormData((prev) => ({ ...prev, kelas: e.target.value }));
              resetFilterResult();
            }}
            disabled={!selectedHotel || !selectedGelombang}
          >
            <option value="">Pilih kelas</option>
            {kelasOptions.map((k) => <option key={k} value={k}>{k}</option>)}
          </select>
        </div>
      </div>

      {/* Tombol Filter */}
      {mode === "grouped" ? (
        <div className="grid gap-3 pt-1 sm:grid-cols-2">
          {Object.entries(DAFTAR_HADIR_GROUPS).map(([key, grp]) => (
            <button
              key={key}
              type="button"
              onClick={() => runFilter(key)}
              className={`inline-flex items-center justify-center gap-2 rounded-2xl px-5 py-3 text-sm font-black shadow-lg transition hover:-translate-y-0.5 ${
                selectedFilterGroup === key
                  ? "bg-orange-600 text-white shadow-orange-500/25"
                  : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"
              }`}
            >
              {key === "pml-ppl" ? <Users size={16} /> : <Briefcase size={16} />}
              Khusus {grp.label}
            </button>
          ))}
        </div>
      ) : (
        <button
          type="button"
          onClick={() => runFilter("all")}
          className={`w-full inline-flex items-center justify-center gap-2 rounded-2xl px-5 py-3 text-sm font-black shadow-lg transition hover:-translate-y-0.5 ${
            filtered
              ? "bg-orange-600 text-white shadow-orange-500/25"
              : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"
          }`}
        >
          <Users size={16} /> Tampilkan PML & PPL
        </button>
      )}

      {/* Status hasil */}
      {filtered && (
        <div className="flex flex-wrap items-center gap-2 rounded-2xl border border-green-200 bg-green-50 px-4 py-3">
          <CheckCircle size={16} className="text-green-600" />
          <span className="text-sm font-bold text-green-700">
            {filteredPeserta.length} peserta ditemukan
            {selectedFilterGroup && selectedFilterGroup !== "all"
              ? ` untuk ${DAFTAR_HADIR_GROUPS[selectedFilterGroup]?.label}`
              : ""}
          </span>
        </div>
      )}

      {/* Tabel preview peserta */}
      {filtered && filteredPeserta.length > 0 && (
        <motion.div initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}>
          <div className="overflow-hidden rounded-2xl border border-orange-100 bg-white/80">
            <div className="grid grid-cols-[2rem_1fr_1fr_1fr] gap-0 border-b border-orange-100 bg-orange-500 px-4 py-2 text-xs font-black uppercase text-white">
              <span>No</span><span>Nama</span><span>Jabatan</span><span>Wil. Tugas</span>
            </div>
            <div className="divide-y divide-orange-50 max-h-64 overflow-y-auto">
              {filteredPeserta.map((p, i) => (
                <div key={i} className="grid grid-cols-[2rem_1fr_1fr_1fr] gap-0 px-4 py-2.5 text-xs font-semibold text-slate-700 hover:bg-orange-50/50">
                  <span className="text-slate-400">{i + 1}</span>
                  <span>{p.nama}</span>
                  <span>{p.jabatan}</span>
                  <span>{p.wilTugas}</span>
                </div>
              ))}
            </div>
          </div>
        </motion.div>
      )}

      {filtered && filteredPeserta.length === 0 && (
        <div className="rounded-2xl border border-dashed border-orange-200 bg-orange-50/50 p-6 text-center text-sm font-semibold text-slate-400">
          Tidak ada peserta untuk Hotel {formData.hotel}, Kelas {formData.kelas}, Gelombang {formData.gelombang}
        </div>
      )}
    </div>
  );
}

function FilterPesertaHotelPanel({ xlsxLoaded, formData, setFormData, petugasData, onFilterResult }) {
  const inputCls = "w-full rounded-2xl border border-orange-100 bg-white/80 px-4 py-3 text-sm font-semibold text-slate-800 shadow-sm outline-none focus:border-orange-300 focus:ring-2 focus:ring-orange-100 transition";
  const labelCls = "mb-2 block text-xs font-black uppercase tracking-[0.2em] text-slate-500";

  const [filtered, setFiltered] = useState(false);
  const [filteredPeserta, setFilteredPeserta] = useState([]);

  const selectedHotel = cleanText(formData.hotel);
  const hotelOptions = React.useMemo(
    () => uniqueSorted(petugasData.map((p) => p.hotel)),
    [petugasData]
  );

  const runFilter = () => {
    if (!xlsxLoaded) {
      alert("Data petugas (.xlsx) belum berhasil dimuat.");
      return;
    }
    if (!selectedHotel) {
      alert("Pilih tempat terlebih dahulu.");
      return;
    }

    const hasil = petugasData.filter((p) => cleanText(p.hotel) === selectedHotel);
    setFilteredPeserta(hasil);
    setFiltered(true);
    onFilterResult(hasil);
  };

  const resetFilter = () => {
    setFiltered(false);
    setFilteredPeserta([]);
    onFilterResult([]);
  };

  return (
    <div className="rounded-3xl border border-orange-100 bg-orange-50/60 p-5 space-y-4">
      <p className="text-xs font-black uppercase tracking-[0.2em] text-orange-700 flex items-center gap-2">
        <Filter size={14} /> Filter Peserta Berdasarkan Tempat
      </p>
      <div>
        <label className={labelCls}>Tempat</label>
        <select
          className={inputCls}
          value={formData.hotel || ""}
          onChange={(e) => {
            const hotel = e.target.value;
            setFormData((prev) => ({ ...prev, hotel, tempat: hotel }));
            resetFilter();
          }}
          disabled={!xlsxLoaded || hotelOptions.length === 0}
        >
          <option value="">Pilih tempat</option>
          {hotelOptions.map((hotel) => (
            <option key={hotel} value={hotel}>{hotel}</option>
          ))}
        </select>
      </div>
      <button
        type="button"
        onClick={runFilter}
        className={`w-full inline-flex items-center justify-center gap-2 rounded-2xl px-5 py-3 text-sm font-black shadow-lg transition hover:-translate-y-0.5 ${
          filtered
            ? "bg-orange-600 text-white shadow-orange-500/25"
            : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"
        }`}
      >
        <Users size={16} /> Tampilkan Peserta
      </button>
      {filtered && (
        <div className="rounded-2xl border border-green-200 bg-green-50 px-4 py-3 text-sm font-bold text-green-700">
          {filteredPeserta.length} peserta ditemukan untuk tempat {selectedHotel}
        </div>
      )}
      {filtered && filteredPeserta.length > 0 && (
        <div className="overflow-hidden rounded-2xl border border-orange-100 bg-white/80">
          <div className="grid grid-cols-[2rem_1fr_1fr_1fr] gap-0 border-b border-orange-100 bg-orange-500 px-4 py-2 text-xs font-black uppercase text-white">
            <span>No</span><span>Nama</span><span>Jabatan</span><span>Wil. Tugas</span>
          </div>
          <div className="divide-y divide-orange-50 max-h-64 overflow-y-auto">
            {filteredPeserta.map((p, i) => (
              <div key={i} className="grid grid-cols-[2rem_1fr_1fr_1fr] gap-0 px-4 py-2.5 text-xs font-semibold text-slate-700 hover:bg-orange-50/50">
                <span className="text-slate-400">{i + 1}</span>
                <span>{p.nama}</span>
                <span>{p.jabatan}</span>
                <span>{p.wilTugas}</span>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

// ─── DOC FORM ─────────────────────────────────────────────────────────────────

function DocForm({ docType, formData, setFormData, onPreview, petugasData, xlsxLoaded }) {
  const update = (key, val) => setFormData((p) => ({ ...p, [key]: val }));

  // State untuk daftar-hadir
  const [daftarHadirPeserta,      setDaftarHadirPeserta]      = useState([]);
  const [daftarHadirNamaInda,     setDaftarHadirNamaInda]     = useState("");
  const [daftarHadirFiltered,     setDaftarHadirFiltered]     = useState(false);
  const [daftarHadirFilterGroup,  setDaftarHadirFilterGroup]  = useState("");

  // State untuk tanda-terima
  const [tandaTerimaPeserta,  setTandaTerimaPeserta]  = useState([]);
  const [tandaTerimaFiltered, setTandaTerimaFiltered] = useState(false);

  // State untuk surat pernyataan kendaraan
  const [suratPernyataanPeserta,  setSuratPernyataanPeserta]  = useState([]);
  const [suratPernyataanFiltered, setSuratPernyataanFiltered] = useState(false);

  // State untuk surat tugas
  const [suratTugasPeserta, setSuratTugasPeserta] = useState([]);
  const [suratTugasFiltered, setSuratTugasFiltered] = useState(false);

  // State untuk SPJ
  const [spjPeserta, setSpjPeserta] = useState([]);
  const [spjFiltered, setSpjFiltered] = useState(false);

  // State untuk SPD
  const [spdPeserta, setSpdPeserta] = useState([]);
  const [spdFiltered, setSpdFiltered] = useState(false);

  // State form lainnya
  const [rows,             setRows]             = useState([]);
  const [perlengkapanRows, setPerlengkapanRows] = useState([]);
  const [pengeluaranRows,  setPengeluaranRows]  = useState([]);
  const [pengeluaranPeserta, setPengeluaranPeserta] = useState([]);
  const [pengeluaranFiltered, setPengeluaranFiltered] = useState(false);

  const addRow           = () => setRows((r) => [...r, { nama: "", jabatan: "", hadir: true }]);
  const updateRow        = (i, k, v) => setRows((r) => r.map((row, idx) => idx === i ? { ...row, [k]: v } : row));
  const removeRow        = (i) => setRows((r) => r.filter((_, idx) => idx !== i));
  const addPerlengkapan  = () => setPerlengkapanRows((r) => [...r, { nama: "", satuan: "buah", jumlah: 1 }]);
  const updatePerlengkapan = (i, k, v) => setPerlengkapanRows((r) => r.map((row, idx) => idx === i ? { ...row, [k]: v } : row));
  const removePerlengkapan = (i) => setPerlengkapanRows((r) => r.filter((_, idx) => idx !== i));
  const addPengeluaran   = () => setPengeluaranRows((r) => [...r, { uraian: "", jumlah: "" }]);
  const updatePengeluaran = (i, k, v) => setPengeluaranRows((r) => r.map((row, idx) => idx === i ? { ...row, [k]: v } : row));
  const removePengeluaran = (i) => setPengeluaranRows((r) => r.filter((_, idx) => idx !== i));

  // Set default jam
  React.useEffect(() => {
    if (docType.id === "daftar-hadir" && !formData.jamMulai && !formData.jamSelesai) {
      setFormData(prev => ({ ...prev, jamMulai: "07.00", jamSelesai: "17.00" }));
    }
  }, [docType.id]);

  const inputCls = "w-full rounded-2xl border border-orange-100 bg-white/80 px-4 py-3 text-sm font-semibold text-slate-800 shadow-sm outline-none focus:border-orange-300 focus:ring-2 focus:ring-orange-100 transition";
  const labelCls = "mb-2 block text-xs font-black uppercase tracking-[0.2em] text-slate-500";

  const CommonFields = () => (
    <>
      <div className="grid gap-5 sm:grid-cols-2">
        <div>
          <label className={labelCls}>Nomor Dokumen</label>
          <input className={inputCls} placeholder="Contoh: 100/BPS-3171/2026" value={formData.nomor || ""} onChange={(e) => update("nomor", e.target.value)} />
        </div>
        <div>
          <label className={labelCls}>Tanggal</label>
          <input type="date" className={inputCls} value={formData.tanggal || ""} onChange={(e) => update("tanggal", e.target.value)} />
        </div>
      </div>
      <div>
        <label className={labelCls}>Kegiatan</label>
        <input className={inputCls} placeholder="Pelatihan Petugas Sensus Ekonomi 2026" value={formData.kegiatan || ""} onChange={(e) => update("kegiatan", e.target.value)} />
      </div>
      <div className="grid gap-5 sm:grid-cols-2">
        <div>
          <label className={labelCls}>Lokasi Kegiatan</label>
          <input className={inputCls} placeholder="Kantor BPS Kota Jakarta Timur" value={formData.lokasi || ""} onChange={(e) => update("lokasi", e.target.value)} />
        </div>
        <div>
          <label className={labelCls}>Tanggal Kegiatan</label>
          <input className={inputCls} placeholder="5-7 Juni 2026" value={formData.tanggalKegiatan || ""} onChange={(e) => update("tanggalKegiatan", e.target.value)} />
        </div>
      </div>
    </>
  );

  const PenandatanganFields = () => (
    <div className="grid gap-5 sm:grid-cols-2">
      <div>
        <label className={labelCls}>Nama Kepala BPS</label>
        <input className={inputCls} placeholder="Nama lengkap" value={formData.namaKabps || ""} onChange={(e) => update("namaKabps", e.target.value)} />
      </div>
      <div>
        <label className={labelCls}>NIP Kepala BPS</label>
        <input className={inputCls} placeholder="NIP" value={formData.nipKabps || ""} onChange={(e) => update("nipKabps", e.target.value)} />
      </div>
    </div>
  );

  const PesertaTable = () => (
    <div>
      <div className="mb-3 flex items-center justify-between">
        <label className={labelCls + " mb-0"}>Daftar Peserta</label>
        <button type="button" onClick={addRow} className="inline-flex items-center gap-1.5 rounded-xl bg-orange-500 px-3 py-1.5 text-xs font-black text-white shadow transition hover:bg-orange-600">
          <Plus size={14} /> Tambah Baris
        </button>
      </div>
      {rows.length === 0 ? (
        <div className="rounded-2xl border border-dashed border-orange-200 bg-orange-50/50 p-8 text-center text-sm font-semibold text-slate-400">
          Belum ada peserta. Klik "Tambah Baris" untuk menambahkan.
        </div>
      ) : (
        <div className="space-y-2">
          {rows.map((row, i) => (
            <div key={i} className="grid items-center gap-2 rounded-2xl border border-orange-100 bg-white/80 p-3 sm:grid-cols-[1fr_1fr_auto_auto]">
              <input className="rounded-xl border border-orange-100 bg-white px-3 py-2 text-xs font-semibold text-slate-800 outline-none focus:border-orange-300" placeholder="Nama" value={row.nama} onChange={(e) => updateRow(i, "nama", e.target.value)} />
              <input className="rounded-xl border border-orange-100 bg-white px-3 py-2 text-xs font-semibold text-slate-800 outline-none focus:border-orange-300" placeholder="Jabatan" value={row.jabatan} onChange={(e) => updateRow(i, "jabatan", e.target.value)} />
              <select className="rounded-xl border border-orange-100 bg-white px-3 py-2 text-xs font-semibold text-slate-800 outline-none focus:border-orange-300" value={row.hadir ? "hadir" : "tidak"} onChange={(e) => updateRow(i, "hadir", e.target.value === "hadir")}>
                <option value="hadir">Hadir</option>
                <option value="tidak">Tidak Hadir</option>
              </select>
              <button type="button" onClick={() => removeRow(i)} className="rounded-xl bg-red-50 p-2 text-red-400 transition hover:bg-red-100"><Trash2 size={14} /></button>
            </div>
          ))}
        </div>
      )}
    </div>
  );

  const handleSubmit = (e) => {
    e.preventDefault();

    if (docType.id === "daftar-hadir") {
      if (!daftarHadirFiltered) {
        alert("Filter peserta dulu dengan memilih tombol PML & PPL atau Panitia & Inda.");
        return;
      }
      const finalFormValues = {
        ...formData,
        tempat: formData.tempat || formData.hotel || "",
        kelompokPeserta: DAFTAR_HADIR_GROUPS[daftarHadirFilterGroup]?.label || "",
      };
      onPreview({
        formValues:          finalFormValues,
        peserta:             daftarHadirPeserta,
        namaInda:            daftarHadirNamaInda,
        selectedFilterGroup: daftarHadirFilterGroup,
        kelompokPeserta:     DAFTAR_HADIR_GROUPS[daftarHadirFilterGroup]?.label || "",
        rows:                daftarHadirPeserta.map(p => ({ nama: p.nama, jabatan: p.jabatan, hadir: true })),
        perlengkapanRows:    [],
        pengeluaranRows:     [],
        nomor:               formData.nomor    || "",
        tanggal:             formData.tanggal  || "",
        kegiatan:            formData.kegiatan || "Pelatihan Petugas Sensus Ekonomi 2026",
        lokasi:              finalFormValues.tempat,
        tanggalKegiatan:     formData.tanggal  || "",
        namaKabps:           daftarHadirNamaInda,
        nipKabps:            "",
      });
      return;
    }

    if (docType.id === "surat-pernyataan-kendaraan") {
      if (!suratPernyataanFiltered || suratPernyataanPeserta.length === 0) {
        alert("Tampilkan peserta terlebih dahulu dengan mengklik tombol filter.");
        return;
      }
      const finalFormValues = {
        ...formData,
        tempat: formData.tempat || formData.hotel || "",
      };
      onPreview({
        formValues: finalFormValues,
        peserta:    suratPernyataanPeserta,
        rows:       suratPernyataanPeserta.map(p => ({ nama: p.nama, jabatan: p.jabatan })),
        nomor:      formData.nomor    || "",
        tanggal:    formData.tanggal  || "",
        kegiatan:   "Pelatihan Petugas Sensus Ekonomi 2026",
        lokasi:     finalFormValues.tempat,
        gelombang:  formData.gelombang || "",
        kelas:      formData.kelas     || "",
      });
      return;
    }

    if (docType.id === "pengeluaran-riil") {
      if (!pengeluaranFiltered || pengeluaranPeserta.length === 0) {
        alert("Tampilkan peserta terlebih dahulu dengan mengklik tombol filter tempat.");
        return;
      }
      onPreview({
        formValues: {
          tanggal_surat: formData.tanggal_surat || "",
          hotel:         formData.hotel || "",
        },
        peserta: pengeluaranPeserta,
      });
      return;
    }

    if (docType.id === "spj") {
      if (!spjFiltered || spjPeserta.length === 0) {
        alert("Tampilkan peserta terlebih dahulu dengan mengklik tombol filter tempat.");
        return;
      }
      const finalFormValues = {
        tanggal_pelunasan: formData.tanggal_pelunasan || "",
        tempat:            formData.tempat || formData.hotel || "",
        ttd_kiri:          formData.ttd_kiri || "",
        ttd_kanan:         formData.ttd_kanan || "",
      };
      onPreview({
        formValues: finalFormValues,
        peserta:    spjPeserta,
      });
      return;
    }

    if (docType.id === "spd") {
      if (!spdFiltered || spdPeserta.length === 0) {
        alert("Tampilkan peserta terlebih dahulu dengan mengklik tombol filter tempat.");
        return;
      }
      const pesertaRows = spdPeserta.map((p) => ({
        ...p,
        tanggal_awal_kegiatan: formData.tanggal_awal_kegiatan || "",
        tanggal_akhir_kegiatan: formData.tanggal_akhir_kegiatan || "",
        lama: calcDurationDays(formData.tanggal_awal_kegiatan || "", formData.tanggal_akhir_kegiatan || ""),
      }));
      const finalFormValues = {
        nomor:                 formData.nomor || "",
        tanggal:               formData.tanggal || "",
        kegiatan:              formData.kegiatan || "",
        lokasi:                formData.lokasi || formData.tempat || formData.hotel || "",
        tempat:                formData.tempat || formData.hotel || "",
        tanggal_awal_kegiatan: formData.tanggal_awal_kegiatan || "",
        tanggal_akhir_kegiatan: formData.tanggal_akhir_kegiatan || "",
        namaKabps:             formData.namaKabps || "",
        nipKabps:              formData.nipKabps || "",
      };
      onPreview({
        formValues: finalFormValues,
        peserta: pesertaRows,
      });
      return;
    }

    if (docType.id === "surat-tugas") {
      if (!suratTugasFiltered || suratTugasPeserta.length === 0) {
        alert("Tampilkan peserta terlebih dahulu dengan mengklik tombol filter tempat.");
        return;
      }
      const finalFormValues = {
        nomor_surat:            formData.nomor_surat || "",
        tanggal_surat:          formData.tanggal_surat || "",
        tanggal_awal_kegiatan:  formData.tanggal_awal_kegiatan || "",
        tanggal_akhir_kegiatan: formData.tanggal_akhir_kegiatan || "",
      };
      onPreview({
        formValues: finalFormValues,
        peserta: suratTugasPeserta,
      });
      return;
    }

    if (docType.id === "tanda-terima") {
      if (!tandaTerimaFiltered || tandaTerimaPeserta.length === 0) {
        alert("Tampilkan peserta terlebih dahulu dengan mengklik tombol filter.");
        return;
      }
      const finalFormValues = {
        ...formData,
        tempat: formData.tempat || formData.hotel || "",
      };
      onPreview({
        formValues:       finalFormValues,
        peserta:          tandaTerimaPeserta,
        rows:             tandaTerimaPeserta.map(p => ({ nama: p.nama, jabatan: p.jabatan })),
        perlengkapanRows: [],
        pengeluaranRows:  [],
        nomor:            formData.nomor    || "",
        tanggal:          formData.tanggal  || "",
        kegiatan:         "Pelatihan Petugas Sensus Ekonomi 2026",
        lokasi:           finalFormValues.tempat,
      });
      return;
    }

    onPreview({ ...formData, rows, perlengkapanRows, pengeluaranRows });
  };

  const renderFields = () => {
    switch (docType.id) {

      // ── DAFTAR HADIR ───────────────────────────────────────────────────────
      case "daftar-hadir":
        return (
          <>
            {!xlsxLoaded && (
              <div className="flex items-center gap-3 rounded-2xl border border-amber-200 bg-amber-50 px-4 py-3">
                <AlertCircle size={18} className="shrink-0 text-amber-500" />
                <p className="text-sm font-semibold text-amber-700">
                  Data petugas <strong>data-petugas.xlsx</strong> dimuat otomatis dari folder publik.
                </p>
              </div>
            )}

            {/* Tanggal & Jam */}
            <div>
              <label className={labelCls}>Tanggal Kegiatan</label>
              <input
                type="date"
                className={inputCls}
                value={formData.tanggal || ""}
                onChange={(e) => update("tanggal", e.target.value)}
              />
            </div>
            <div className="grid gap-4 sm:grid-cols-2">
              <div>
                <label className={labelCls}>Jam Mulai</label>
                <input
                  className={inputCls}
                  placeholder="07.00"
                  value={formData.jamMulai || "07.00"}
                  onChange={(e) => update("jamMulai", normalizeJamIndonesia(e.target.value))}
                />
              </div>
              <div>
                <label className={labelCls}>Jam Selesai</label>
                <input
                  className={inputCls}
                  placeholder="17.00"
                  value={formData.jamSelesai || "17.00"}
                  onChange={(e) => update("jamSelesai", normalizeJamIndonesia(e.target.value))}
                />
              </div>
            </div>

            {/* Filter Panel */}
            <FilterPesertaPanel
              xlsxLoaded={xlsxLoaded}
              formData={formData}
              setFormData={setFormData}
              petugasData={petugasData}
              mode="grouped"
              onFilterResult={(peserta, namaInda, groupKey) => {
                setDaftarHadirPeserta(peserta);
                setDaftarHadirNamaInda(namaInda);
                setDaftarHadirFilterGroup(groupKey);
                setDaftarHadirFiltered(peserta.length > 0 || groupKey !== "");
              }}
            />

            {/* Keterangan TTD */}
            {daftarHadirFiltered && daftarHadirFilterGroup && (
              <motion.div
                initial={{ opacity: 0, y: 10 }} animate={{ opacity: 1, y: 0 }}
                className="flex items-center justify-between rounded-2xl border border-orange-100 bg-white/80 px-4 py-3"
              >
                <div>
                  <p className="text-xs font-black uppercase tracking-wider text-slate-500">Keterangan Tanda Tangan</p>
                  <p className="mt-0.5 font-bold text-slate-800">
                    {daftarHadirFilterGroup === "panitia-inda" ? "Kepala Sub Bagian Umum" : "Instruktur Daerah"}
                  </p>
                </div>
                <Check size={18} className="text-green-500" />
              </motion.div>
            )}
          </>
        );

      // ── TANDA TERIMA PERLENGKAPAN ──────────────────────────────────────────
      case "tanda-terima":
        return (
          <>
            {!xlsxLoaded && (
              <div className="flex items-center gap-3 rounded-2xl border border-amber-200 bg-amber-50 px-4 py-3">
                <AlertCircle size={18} className="shrink-0 text-amber-500" />
                <p className="text-sm font-semibold text-amber-700">
                  Data petugas <strong>data-petugas.xlsx</strong> dimuat otomatis dari folder publik.
                </p>
              </div>
            )}

            {/* Tanggal */}
            <div>
              <label className={labelCls}>Tanggal Kegiatan</label>
              <input
                type="date"
                className={inputCls}
                value={formData.tanggal || ""}
                onChange={(e) => update("tanggal", e.target.value)}
              />
            </div>

            {/* Filter Panel — mode "all" */}
            <FilterPesertaPanel
              xlsxLoaded={xlsxLoaded}
              formData={formData}
              setFormData={setFormData}
              petugasData={petugasData}
              mode="all"
              onFilterResult={(peserta, _namaInda, _groupKey) => {
                setTandaTerimaPeserta(peserta);
                setTandaTerimaFiltered(true);
              }}
            />
          </>
        );

      // ── SURAT PERNYATAAN KENDARAAN ──────────────────────────────────────────
      case "surat-pernyataan-kendaraan":
        return (
          <>
            <div>
              <label className={labelCls}>Tanggal Surat</label>
              <input
                type="date"
                className={inputCls}
                value={formData.tanggal || ""}
                onChange={(e) => update("tanggal", e.target.value)}
              />
            </div>
            <FilterPesertaPanel
              xlsxLoaded={xlsxLoaded}
              formData={formData}
              setFormData={setFormData}
              petugasData={petugasData}
              mode="all"
              onFilterResult={(peserta, _namaInda, _groupKey) => {
                setSuratPernyataanPeserta(peserta);
                setSuratPernyataanFiltered(true);
              }}
            />
          </>
        );

      // ── PENGELUARAN RIIL ────────────────────────────────────────────────────
      case "pengeluaran-riil":
        return (
          <>
            <div>
              <label className={labelCls}>Tanggal Surat</label>
              <input
                type="date"
                className={inputCls}
                value={formData.tanggal_surat || ""}
                onChange={(e) => update("tanggal_surat", e.target.value)}
              />
            </div>
            <FilterPesertaHotelPanel
              xlsxLoaded={xlsxLoaded}
              formData={formData}
              setFormData={setFormData}
              petugasData={petugasData}
              onFilterResult={(peserta) => {
                setPengeluaranPeserta(peserta);
                setPengeluaranFiltered(peserta.length > 0);
              }}
            />
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
            {/* <div className="rounded-2xl border border-orange-100 bg-orange-50 p-4 text-sm text-slate-700">
              <p className="font-bold">Catatan SPJ</p>
              <p className="mt-2">Filter hotel digunakan untuk memilih peserta. Tidak perlu mengisi tempat/hotel secara terpisah.</p>
            </div> */}
            <FilterPesertaHotelPanel
              xlsxLoaded={xlsxLoaded}
              formData={formData}
              setFormData={setFormData}
              petugasData={petugasData}
              onFilterResult={(peserta) => {
                setSpjPeserta(peserta);
                setSpjFiltered(peserta.length > 0);
              }}
            />
          </>
        );

      // ── SPD ─────────────────────────────────────────────────────────────────
      case "spd":
        return (
          <>
            <div className="grid gap-5 sm:grid-cols-2">
              <div>
                <label className={labelCls}>Nomor Dokumen</label>
                <input className={inputCls} placeholder="Contoh: 100/BPS-3171/2026" value={formData.nomor || ""} onChange={(e) => update("nomor", e.target.value)} />
              </div>
              <div>
                <label className={labelCls}>Tanggal Surat</label>
                <input type="date" className={inputCls} value={formData.tanggal || ""} onChange={(e) => update("tanggal", e.target.value)} />
              </div>
            </div>
            <div className="grid gap-5 sm:grid-cols-2">
              <div>
                <label className={labelCls}>Tanggal Awal Kegiatan</label>
                <input type="date" className={inputCls} value={formData.tanggal_awal_kegiatan || ""} onChange={(e) => update("tanggal_awal_kegiatan", e.target.value)} />
              </div>
              <div>
                <label className={labelCls}>Tanggal Akhir Kegiatan</label>
                <input type="date" className={inputCls} value={formData.tanggal_akhir_kegiatan || ""} onChange={(e) => update("tanggal_akhir_kegiatan", e.target.value)} />
              </div>
            </div>
            {/* <div className="rounded-2xl border border-orange-100 bg-orange-50 p-4 text-sm text-slate-700">
              <p className="font-bold">Catatan SPD</p>
              <p className="mt-2">Gunakan filter tempat untuk pilih peserta yang akan masuk ke Lampiran SPD.</p>
            </div> */}
            <FilterPesertaHotelPanel
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
              <div>
                <label className={labelCls}>Nomor Surat</label>
                <input className={inputCls} placeholder="Contoh: 100/BPS-3171/2026" value={formData.nomor_surat || ""} onChange={(e) => update("nomor_surat", e.target.value)} />
              </div>
              <div>
                <label className={labelCls}>Tanggal Surat</label>
                <input type="date" className={inputCls} value={formData.tanggal_surat || ""} onChange={(e) => update("tanggal_surat", e.target.value)} />
              </div>
            </div>
            <div className="grid gap-5 sm:grid-cols-2">
              <div>
                <label className={labelCls}>Tanggal Awal Penyelenggaraan</label>
                <input type="date" className={inputCls} value={formData.tanggal_awal_kegiatan || ""} onChange={(e) => update("tanggal_awal_kegiatan", e.target.value)} />
              </div>
              <div>
                <label className={labelCls}>Tanggal Akhir Penyelenggaraan</label>
                <input type="date" className={inputCls} value={formData.tanggal_akhir_kegiatan || ""} onChange={(e) => update("tanggal_akhir_kegiatan", e.target.value)} />
              </div>
            </div>
            <FilterPesertaHotelPanel
              xlsxLoaded={xlsxLoaded}
              formData={formData}
              setFormData={setFormData}
              petugasData={petugasData}
              onFilterResult={(peserta) => {
                setSuratTugasPeserta(peserta);
                setSuratTugasFiltered(peserta.length > 0);
              }}
            />
            {suratTugasFiltered && suratTugasPeserta.length > 0 && (
              <div className="rounded-3xl border border-green-200 bg-green-50 p-4 text-sm text-slate-700">
                <p className="font-black">Peserta Surat Tugas</p>
                <p className="mt-2">{suratTugasPeserta.length} peserta terpilih dari tempat "{formData.hotel || formData.tempat || '-'}".</p>
              </div>
            )}
          </>
        );

      default:
        return <CommonFields />;
    }
  };

  return (
    <form onSubmit={handleSubmit} className="space-y-5 rounded-[2.5rem] border border-orange-100 bg-white/80 p-6 shadow-xl shadow-orange-900/5 backdrop-blur md:p-10">
      {renderFields()}
      <div className="border-t border-orange-100 pt-5">
        <button
          type="submit"
          className="group inline-flex items-center gap-2 rounded-2xl bg-orange-500 px-7 py-4 font-black text-white shadow-2xl shadow-orange-500/25 transition hover:-translate-y-1 hover:bg-orange-600"
        >
          Pratinjau Dokumen
          <ChevronRight className="transition group-hover:translate-x-1" size={18} />
        </button>
      </div>
    </form>
  );
}

// ─── DAFTAR HADIR — DOCX PREVIEW ─────────────────────────────────────────────

function DaftarHadirDocxPreview({ formValues, peserta, namaInda, selectedFilterGroup }) {
  const containerRef = useRef(null);
  const [loading, setLoading] = useState(true);
  const [error,   setError]   = useState("");

  React.useEffect(() => {
    let cancelled = false;
    const renderPreview = async () => {
      try {
        setLoading(true);
        setError("");
        const blob = await createDaftarHadirBlob(
          DAFTAR_HADIR_TEMPLATE_URL,
          formValues || {},
          peserta    || [],
          namaInda   || "",
          selectedFilterGroup || ""
        );
        if (cancelled || !containerRef.current) return;
        containerRef.current.innerHTML = "";
        await renderAsync(blob, containerRef.current, null, {
          className: "docx-preview", inWrapper: true,
          ignoreWidth: false, ignoreHeight: false, ignoreFonts: false,
          breakPages: true, renderHeaders: true, renderFooters: true, useBase64URL: true,
        });
      } catch (err) {
        if (!cancelled) setError(err?.message || "Gagal memuat pratinjau DOCX.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    };
    renderPreview();
    return () => { cancelled = true; };
  }, [formValues, peserta, namaInda, selectedFilterGroup]);

  return (
    <div className="rounded-[2rem] border border-orange-100 bg-white p-4 shadow-xl shadow-orange-900/5">
      <style>{`
        .docx-wrapper { background: transparent !important; padding: 0 !important; }
        .docx-wrapper > section.docx { margin: 0 auto 24px auto !important; box-shadow: 0 20px 45px rgba(15,23,42,0.12) !important; }
        @media print {
          body * { visibility: hidden; }
          .docx-wrapper, .docx-wrapper * { visibility: visible; }
          .docx-wrapper { position: absolute; left: 0; top: 0; width: 100%; }
          .docx-wrapper > section.docx { box-shadow: none !important; margin: 0 !important; }
        }
      `}</style>
      {loading && (
        <div className="rounded-2xl border border-orange-100 bg-orange-50/60 p-8 text-center">
          <p className="text-sm font-black text-orange-700">Memuat pratinjau dari template DOCX...</p>
          <p className="mt-1 text-xs font-semibold text-slate-500">Template: 1. Daftar Hadir Pelatihan SE2026.docx</p>
        </div>
      )}
      {error && (
        <div className="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm font-bold text-red-600">{error}</div>
      )}
      <div ref={containerRef} className="overflow-x-auto" />
    </div>
  );
}

// ─── TANDA TERIMA — DOCX PREVIEW ─────────────────────────────────────────────

function TandaTerimaDocxPreview({ formValues, peserta }) {
  const containerRef = useRef(null);
  const [loading, setLoading] = useState(true);
  const [error,   setError]   = useState("");

  React.useEffect(() => {
    let cancelled = false;
    const renderPreview = async () => {
      try {
        setLoading(true);
        setError("");
        const blob = await createTandaTerimaBlob(
          TANDA_TERIMA_TEMPLATE_URL,
          formValues || {},
          peserta    || []
        );
        if (cancelled || !containerRef.current) return;
        containerRef.current.innerHTML = "";
        await renderAsync(blob, containerRef.current, null, {
          className: "docx-preview", inWrapper: true,
          ignoreWidth: false, ignoreHeight: false, ignoreFonts: false,
          breakPages: true, renderHeaders: true, renderFooters: true, useBase64URL: true,
        });
      } catch (err) {
        if (!cancelled) setError(err?.message || "Gagal memuat pratinjau DOCX.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    };
    renderPreview();
    return () => { cancelled = true; };
  }, [formValues, peserta]);

  return (
    <div className="rounded-[2rem] border border-orange-100 bg-white p-4 shadow-xl shadow-orange-900/5">
      <style>{`
        .docx-wrapper { background: transparent !important; padding: 0 !important; }
        .docx-wrapper > section.docx { margin: 0 auto 24px auto !important; box-shadow: 0 20px 45px rgba(15,23,42,0.12) !important; }
        @media print {
          body * { visibility: hidden; }
          .docx-wrapper, .docx-wrapper * { visibility: visible; }
          .docx-wrapper { position: absolute; left: 0; top: 0; width: 100%; }
          .docx-wrapper > section.docx { box-shadow: none !important; margin: 0 !important; }
        }
      `}</style>
      {loading && (
        <div className="rounded-2xl border border-orange-100 bg-orange-50/60 p-8 text-center">
          <p className="text-sm font-black text-orange-700">Memuat pratinjau dari template DOCX...</p>
          <p className="mt-1 text-xs font-semibold text-slate-500">Template: 2. Tanda Terima Perlengkapan SE2026.docx</p>
        </div>
      )}
      {error && (
        <div className="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm font-bold text-red-600">{error}</div>
      )}
      <div ref={containerRef} className="overflow-x-auto" />
    </div>
  );
}

function SuratPernyataanKendaraanDocxPreview({ formValues, peserta }) {
  const containerRef = useRef(null);
  const [loading, setLoading] = useState(true);
  const [error,   setError]   = useState("");

  React.useEffect(() => {
    let cancelled = false;
    const renderPreview = async () => {
      try {
        setLoading(true);
        setError("");
        const blob = await createSuratPernyataanKendaraanBlob(
          SURAT_PERNYATAAN_KENDARAAN_TEMPLATE_URL,
          formValues || {},
          peserta    || []
        );
        if (cancelled || !containerRef.current) return;
        containerRef.current.innerHTML = "";
        await renderAsync(blob, containerRef.current, null, {
          className: "docx-preview",
          inWrapper: true,
          ignoreWidth: false,
          ignoreHeight: false,
          ignoreFonts: false,
          breakPages: true,
          renderHeaders: true,
          renderFooters: true,
          useBase64URL: true,
        });
      } catch (err) {
        if (!cancelled) setError(err?.message || "Gagal memuat pratinjau DOCX.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    };
    renderPreview();
    return () => { cancelled = true; };
  }, [formValues, peserta]);

  return (
    <div className="rounded-[2rem] border border-orange-100 bg-white p-4 shadow-xl shadow-orange-900/5">
      <style>{`
        .docx-wrapper { background: transparent !important; padding: 0 !important; }
        .docx-wrapper > section.docx { margin: 0 auto 24px auto !important; box-shadow: 0 20px 45px rgba(15,23,42,0.12) !important; }
        @media print {
          body * { visibility: hidden; }
          .docx-wrapper, .docx-wrapper * { visibility: visible; }
          .docx-wrapper { position: absolute; left: 0; top: 0; width: 100%; }
          .docx-wrapper > section.docx { box-shadow: none !important; margin: 0 !important; }
        }
      `}</style>
      {loading && (
        <div className="rounded-2xl border border-orange-100 bg-orange-50/60 p-8 text-center">
          <p className="text-sm font-black text-orange-700">Memuat pratinjau dari template DOCX...</p>
          <p className="mt-1 text-xs font-semibold text-slate-500">Template: 3. Super Kendis Pelatihan SE2026.docx</p>
        </div>
      )}
      {error && (
        <div className="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm font-bold text-red-600">{error}</div>
      )}
      <div ref={containerRef} className="overflow-x-auto" />
    </div>
  );
}

function PengeluaranRiilDocxPreview({ formValues, peserta }) {
  const containerRef = useRef(null);
  const [loading, setLoading] = useState(true);
  const [error,   setError]   = useState("");

  React.useEffect(() => {
    let cancelled = false;
    const renderPreview = async () => {
      try {
        setLoading(true);
        setError("");
        const blob = await createPengeluaranRiilBlob(
          PENGELUARAN_RIIL_TEMPLATE_URL,
          formValues || {},
          peserta    || []
        );
        if (cancelled || !containerRef.current) return;
        containerRef.current.innerHTML = "";
        await renderAsync(blob, containerRef.current, null, {
          className: "docx-preview",
          inWrapper: true,
          ignoreWidth: false,
          ignoreHeight: false,
          ignoreFonts: false,
          breakPages: true,
          renderHeaders: true,
          renderFooters: true,
          useBase64URL: true,
        });
      } catch (err) {
        if (!cancelled) setError(err?.message || "Gagal memuat pratinjau DOCX.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    };
    renderPreview();
    return () => { cancelled = true; };
  }, [formValues, peserta]);

  return (
    <div className="rounded-[2rem] border border-orange-100 bg-white p-4 shadow-xl shadow-orange-900/5">
      <style>{`
        .docx-wrapper { background: transparent !important; padding: 0 !important; }
        .docx-wrapper > section.docx { margin: 0 auto 24px auto !important; box-shadow: 0 20px 45px rgba(15,23,42,0.12) !important; }
        @media print {
          body * { visibility: hidden; }
          .docx-wrapper, .docx-wrapper * { visibility: visible; }
          .docx-wrapper { position: absolute; left: 0; top: 0; width: 100%; }
          .docx-wrapper > section.docx { box-shadow: none !important; margin: 0 !important; }
        }
      `}</style>
      {loading && (
        <div className="rounded-2xl border border-orange-100 bg-orange-50/60 p-8 text-center">
          <p className="text-sm font-black text-orange-700">Memuat pratinjau dari template DOCX...</p>
          <p className="mt-1 text-xs font-semibold text-slate-500">Template: 4. DPR_Pelatihan SE 2026.docx</p>
        </div>
      )}
      {error && (
        <div className="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm font-bold text-red-600">{error}</div>
      )}
      <div ref={containerRef} className="overflow-x-auto" />
    </div>
  );
}

function SpjDocxPreview({ formValues, peserta }) {
  const containerRef = useRef(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  React.useEffect(() => {
    let cancelled = false;
    const renderPreview = async () => {
      try {
        setLoading(true);
        setError("");
        const blob = await createSpjBlob(
          SPJ_TEMPLATE_URL,
          formValues || {},
          peserta || []
        );
        if (cancelled || !containerRef.current) return;
        containerRef.current.innerHTML = "";
        await renderAsync(blob, containerRef.current, null, {
          className: "docx-preview",
          inWrapper: true,
          ignoreWidth: false,
          ignoreHeight: false,
          ignoreFonts: false,
          breakPages: true,
          renderHeaders: true,
          renderFooters: true,
          useBase64URL: true,
        });
      } catch (err) {
        if (!cancelled) setError(err?.message || "Gagal memuat pratinjau DOCX.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    };
    renderPreview();
    return () => { cancelled = true; };
  }, [formValues, peserta]);

  return (
    <div className="rounded-[2rem] border border-orange-100 bg-white p-4 shadow-xl shadow-orange-900/5">
      <style>{`
        .docx-wrapper { background: transparent !important; padding: 0 !important; }
        .docx-wrapper > section.docx { margin: 0 auto 24px auto !important; box-shadow: 0 20px 45px rgba(15,23,42,0.12) !important; }
        @media print {
          body * { visibility: hidden; }
          .docx-wrapper, .docx-wrapper * { visibility: visible; }
          .docx-wrapper { position: absolute; left: 0; top: 0; width: 100%; }
          .docx-wrapper > section.docx { box-shadow: none !important; margin: 0 !important; }
        }
      `}</style>
      {loading && (
        <div className="rounded-2xl border border-orange-100 bg-orange-50/60 p-8 text-center">
          <p className="text-sm font-black text-orange-700">Memuat pratinjau dari template DOCX...</p>
          <p className="mt-1 text-xs font-semibold text-slate-500">Template: 5. SPJ Pelatihan_SE26.docx</p>
        </div>
      )}
      {error && (
        <div className="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm font-bold text-red-600">{error}</div>
      )}
      <div ref={containerRef} className="overflow-x-auto" />
    </div>
  );
}

function SpdDocxPreview({ formValues, peserta }) {
  const mainRef = useRef(null);
  const lampiranRef = useRef(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");
  const [activePreview, setActivePreview] = useState("spd");

  React.useEffect(() => {
    let cancelled = false;
    const renderPreview = async () => {
      try {
        setLoading(true);
        setError("");

        const mainBlob = await createSpdBlob(
          SPD_TEMPLATE_URL,
          formValues || {},
          peserta || []
        );
        const lampiranBlob = await createSpdBlob(
          SPD_LAMPIRAN_TEMPLATE_URL,
          formValues || {},
          peserta || []
        );

        if (cancelled) return;
        if (mainRef.current) {
          mainRef.current.innerHTML = "";
          await renderAsync(mainBlob, mainRef.current, null, {
            className: "docx-preview",
            inWrapper: true,
            ignoreWidth: false,
            ignoreHeight: false,
            ignoreFonts: false,
            breakPages: true,
            renderHeaders: true,
            renderFooters: true,
            useBase64URL: true,
          });
        }
        if (lampiranRef.current) {
          lampiranRef.current.innerHTML = "";
          await renderAsync(lampiranBlob, lampiranRef.current, null, {
            className: "docx-preview",
            inWrapper: true,
            ignoreWidth: false,
            ignoreHeight: false,
            ignoreFonts: false,
            breakPages: true,
            renderHeaders: true,
            renderFooters: true,
            useBase64URL: true,
          });
        }
      } catch (err) {
        if (!cancelled) setError(err?.message || "Gagal memuat pratinjau DOCX.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    };
    renderPreview();
    return () => { cancelled = true; };
  }, [formValues, peserta]);

  return (
    <div className="space-y-8 rounded-[2rem] border border-orange-100 bg-white p-4 shadow-xl shadow-orange-900/5">
      <style>{`
        .docx-wrapper { background: transparent !important; padding: 0 !important; }
        .docx-wrapper > section.docx { margin: 0 auto 24px auto !important; box-shadow: 0 20px 45px rgba(15,23,42,0.12) !important; }
        @media print {
          body * { visibility: hidden; }
          .docx-wrapper, .docx-wrapper * { visibility: visible; }
          .docx-wrapper { position: absolute; left: 0; top: 0; width: 100%; }
          .docx-wrapper > section.docx { box-shadow: none !important; margin: 0 !important; }
        }
      `}</style>
      {loading && (
        <div className="rounded-2xl border border-orange-100 bg-orange-50/60 p-8 text-center">
          <p className="text-sm font-black text-orange-700">Memuat pratinjau dari template DOCX...</p>
          <p className="mt-1 text-xs font-semibold text-slate-500">Template: 6. SPD.docx / 6. Lampiran SPD.docx</p>
        </div>
      )}
      {error && (
        <div className="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm font-bold text-red-600">{error}</div>
      )}
      <div className="flex flex-wrap gap-3">
        <button
          type="button"
          onClick={() => setActivePreview("spd")}
          className={`inline-flex items-center justify-center rounded-2xl px-5 py-3 text-sm font-black transition ${activePreview === "spd" ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}
        >
          Preview SPD
        </button>
        <button
          type="button"
          onClick={() => setActivePreview("lampiran")}
          className={`inline-flex items-center justify-center rounded-2xl px-5 py-3 text-sm font-black transition ${activePreview === "lampiran" ? "bg-orange-600 text-white" : "bg-white text-orange-700 border border-orange-200 hover:bg-orange-50"}`}
        >
          Preview Lampiran SPD
        </button>
      </div>
      <div className="mt-6">
        <div className={activePreview === "spd" ? "block" : "hidden"}>
          <div className="mb-4 rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm font-black uppercase tracking-[0.2em] text-slate-700">SPD</div>
          <div ref={mainRef} className="overflow-x-auto rounded-2xl border border-orange-100 bg-white" />
        </div>
        <div className={activePreview === "lampiran" ? "block" : "hidden"}>
          <div className="mb-4 rounded-2xl border border-slate-200 bg-slate-50 px-4 py-3 text-sm font-black uppercase tracking-[0.2em] text-slate-700">Lampiran SPD</div>
          <div ref={lampiranRef} className="overflow-x-auto rounded-2xl border border-orange-100 bg-white" />
        </div>
      </div>
    </div>
  );
}

function SuratTugasDocxPreview({ formValues, peserta }) {
  const containerRef = useRef(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  React.useEffect(() => {
    let cancelled = false;
    const renderPreview = async () => {
      try {
        setLoading(true);
        setError("");
        const blob = await createSuratTugasBlob(
          SURAT_TUGAS_TEMPLATE_URL,
          formValues || {},
          peserta || []
        );
        if (cancelled || !containerRef.current) return;
        containerRef.current.innerHTML = "";
        await renderAsync(blob, containerRef.current, null, {
          className: "docx-preview",
          inWrapper: true,
          ignoreWidth: false,
          ignoreHeight: false,
          ignoreFonts: false,
          breakPages: true,
          renderHeaders: true,
          renderFooters: true,
          useBase64URL: true,
        });
      } catch (err) {
        if (!cancelled) setError(err?.message || "Gagal memuat pratinjau DOCX.");
      } finally {
        if (!cancelled) setLoading(false);
      }
    };
    renderPreview();
    return () => { cancelled = true; };
  }, [formValues, peserta]);

  return (
    <div className="rounded-[2rem] border border-orange-100 bg-white p-4 shadow-xl shadow-orange-900/5">
      <style>{`
        .docx-wrapper { background: transparent !important; padding: 0 !important; }
        .docx-wrapper > section.docx { margin: 0 auto 24px auto !important; box-shadow: 0 20px 45px rgba(15,23,42,0.12) !important; }
        @media print {
          body * { visibility: hidden; }
          .docx-wrapper, .docx-wrapper * { visibility: visible; }
          .docx-wrapper { position: absolute; left: 0; top: 0; width: 100%; }
          .docx-wrapper > section.docx { box-shadow: none !important; margin: 0 !important; }
        }
      `}</style>
      {loading && (
        <div className="rounded-2xl border border-orange-100 bg-orange-50/60 p-8 text-center">
          <p className="text-sm font-black text-orange-700">Memuat pratinjau dari template DOCX...</p>
          <p className="mt-1 text-xs font-semibold text-slate-500">Template: 6. Surat Tugas.docx</p>
        </div>
      )}
      {error && (
        <div className="rounded-2xl border border-red-200 bg-red-50 p-4 text-sm font-bold text-red-600">{error}</div>
      )}
      <div ref={containerRef} className="overflow-x-auto" />
    </div>
  );
}

// ─── DOC PREVIEW ──────────────────────────────────────────────────────────────

function DocPreview({ docType, data }) {
  const fmtDate = (d) => {
    if (!d) return "_______________";
    const date = new Date(d);
    return date.toLocaleDateString("id-ID", { day: "numeric", month: "long", year: "numeric" });
  };

  const Header = () => (
    <div className="mb-6 flex items-start gap-4 border-b-2 border-slate-800 pb-4">
      <div className="h-20 w-20 shrink-0 rounded-xl bg-slate-200 flex items-center justify-center text-xs font-bold text-slate-500">LOGO BPS</div>
      <div className="flex-1 text-center">
        <p className="text-sm font-bold uppercase">Badan Pusat Statistik</p>
        <p className="text-base font-black uppercase">Kota Jakarta Timur</p>
        <p className="text-xs text-slate-600">Jl. Pemuda No.67, Jakarta Timur 13220 | Telp: (021) 4750252</p>
        <p className="text-xs text-slate-600">Email: bps3175@bps.go.id | Website: jaktimkota.bps.go.id</p>
      </div>
      <div className="h-20 w-20 shrink-0 rounded-xl bg-slate-200 flex items-center justify-center text-xs font-bold text-slate-500">LOGO SE</div>
    </div>
  );

  const TtdField = ({ nama, nip, jabatan = "Kepala BPS Kota Jakarta Timur," }) => (
    <div className="mt-10 flex justify-end">
      <div className="text-center text-sm">
        <p>Jakarta, {fmtDate(data.tanggal || data.formValues?.tanggal)}</p>
        <p className="font-bold">{jabatan}</p>
        <div className="my-14" />
        <p className="font-black underline">{nama || "______________________________"}</p>
        <p>NIP. {nip || "_______________________"}</p>
      </div>
    </div>
  );

  const renderDoc = () => {
    switch (docType.id) {

      case "daftar-hadir":
        return (
          <DaftarHadirDocxPreview
            formValues={data.formValues || {}}
            peserta={data.peserta || []}
            namaInda={data.namaInda || ""}
            selectedFilterGroup={data.selectedFilterGroup || ""}
          />
        );

      // ── TANDA TERIMA — render via DOCX template ────────────────────────────
      case "tanda-terima":
        return (
          <TandaTerimaDocxPreview
            formValues={data.formValues || {}}
            peserta={data.peserta || []}
          />
        );

      case "surat-pernyataan-kendaraan":
        return (
          <SuratPernyataanKendaraanDocxPreview
            formValues={data.formValues || {}}
            peserta={data.peserta || []}
          />
        );

      case "pengeluaran-riil":
        return (
          <PengeluaranRiilDocxPreview
            formValues={data.formValues || {}}
            peserta={data.peserta || []}
          />
        );

      case "spj":
        return (
          <SpjDocxPreview
            formValues={data.formValues || {}}
            peserta={data.peserta || []}
          />
        );

      case "spd":
        return (
          <SpdDocxPreview
            formValues={data.formValues || data}
            peserta={data.peserta || []}
          />
        );

      case "surat-tugas":
        return (
          <SuratTugasDocxPreview
            formValues={data.formValues || {}}
            peserta={data.peserta || []}
          />
        );

      default:
        return <p>Dokumen tidak dikenali.</p>;
    }
  };

  return (
    <div className="overflow-hidden rounded-[2.5rem] border border-orange-100 bg-white/80 shadow-xl shadow-orange-900/5 backdrop-blur">
      <div className="border-b border-orange-100 bg-orange-500 px-8 py-4">
        <div className="flex items-center gap-3">
          <div className="flex gap-1.5">
            <div className="h-3 w-3 rounded-full bg-white/40" />
            <div className="h-3 w-3 rounded-full bg-white/40" />
            <div className="h-3 w-3 rounded-full bg-white/40" />
          </div>
          <p className="text-sm font-black text-white">Pratinjau: {docType.label}</p>
        </div>
      </div>
      <div className="p-8 md:p-12 print:p-0">
        <div className="mx-auto w-full max-w-none font-serif text-slate-900">
          {renderDoc()}
        </div>
      </div>
    </div>
  );
}