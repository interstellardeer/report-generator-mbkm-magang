/**
 * generate_laporan_mbkm.js
 * Generates a DOCX template for Laporan Pelaksanaan Kegiatan MBKM
 * (Program MSIB / P3NK - Magang Mandiri)
 * Program Studi Rekayasa Perangkat Lunak, Kampus UPI di Cibiru
 *
 * Usage: node generate_laporan_mbkm.js
 * Output: Laporan_MBKM_Template.docx
 *
 * FILL-IN VARIABLES — edit the CONFIG section before running.
 */

const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  AlignmentType,
  HeadingLevel,
  LevelFormat,
  BorderStyle,
  WidthType,
  ShadingType,
  VerticalAlign,
  PageBreak,
  PageNumber,
  NumberFormat,
  Footer,
} = require("docx");
const fs = require("fs");

// ─────────────────────────────────────────────────────────────────────────────
// CONFIG — Replace all placeholder values before running
// ─────────────────────────────────────────────────────────────────────────────
const CONFIG = {
  // Student info
  nama_mahasiswa: "Nama Lengkap Mahasiswa",
  nim: "XXXXXXXXX",
  tahun: "2026",

  // Supervisor (Dosen Pembimbing)
  nama_dosen_pembimbing: "Nama Dosen, S.T., M.T.",
  nip_dosen: "XXXXXXXXXXXXXXXX",

  // Company supervisor (Penyelia)
  nama_penyelia: "Nama Penyelia",
  nip_penyelia: "XXXXXX",

  // Program head (Ketua Prodi)
  nama_kaprodi: "Nama Kaprodi, S.T., M.Kom.",
  nip_kaprodi: "XXXXXXXXXXXXXXXX",

  // Internship details
  nama_institusi_mitra: "Nama Perusahaan / Institusi",
  unit_penempatan: "Nama Divisi / Unit / Tim",
  periode_pelaksanaan: "DD Bulan YYYY – DD Bulan YYYY",
  durasi_kegiatan: "X bulan (setara dengan XXX jam kegiatan MBKM)",
  hari_jam_kerja: "Senin – Jumat, 08.00 – 17.00 WIB",
  lokasi_magang: "Nama Kota, Nama Gedung / Kantor",
  alamat_lokasi: "Jl. Nama Jalan No.XX, Kelurahan, Kecamatan, Kota, Kode Pos",
  internship_role: "Nama Role / Jabatan Selama Magang",

  // Chapter 1
  latar_belakang: "[Isi latar belakang: uraikan fenomena/masalah yang diamati, relevansi dengan bidang rekayasa perangkat lunak, alasan memilih institusi mitra ini, dan kaitannya dengan program MBKM. Minimal 3–4 paragraf.]",
  manfaat_tujuan: "[Uraikan manfaat dan tujuan kegiatan MBKM secara spesifik: manfaat bagi mahasiswa, institusi mitra, dan program studi; serta tujuan konkret yang ingin dicapai selama magang.]",
  ruang_lingkup: "[Jelaskan cakupan kegiatan: jenis proyek yang dikerjakan, teknologi yang terlibat, divisi yang dimasuki, dan batasan tugas selama magang.]",

  // Chapter 2
  gambaran_umum_institusi: "[Deskripsikan profil perusahaan/institusi: sejarah singkat, visi-misi, skala bisnis, produk atau layanan utama, dan posisinya dalam industri.]",
  bidang_kerja: "[Jelaskan bidang kerja atau usaha institusi mitra secara spesifik, termasuk divisi/unit tempat magang dilaksanakan dan fungsinya dalam struktur organisasi.]",
  peran_mahasiswa: "[Uraikan peran dan tanggung jawab selama magang: jabatan, deskripsi tugas harian, proyek yang diikuti, dan kontribusi teknis yang diberikan.]",

  // Chapter 3 — work plan rows: [tahapan, deskripsi_aktivitas]
  rencana_kegiatan_rows: [
    ["Orientasi dan pembekalan", "Pengenalan lingkungan kerja, SOP tim, alur proyek, dan perkenalan dengan anggota tim."],
    ["Studi dan analisis kebutuhan", "Mempelajari konteks proyek, kebutuhan pengguna, dan mendiskusikan skenario pengembangan."],
    ["Desain UI/UX", "Merancang antarmuka dan alur interaksi menggunakan tools desain (Figma, dll.)."],
    ["Implementasi teknis", "Mengembangkan fitur sesuai rencana menggunakan teknologi yang telah ditentukan."],
    ["Pengujian dan debugging", "Melakukan uji coba internal, mencatat bug, dan memperbaiki kesalahan fungsional."],
    ["Evaluasi dan dokumentasi", "Mengevaluasi hasil kerja mingguan, menyusun laporan teknis dan logbook kegiatan."],
  ],

  // Chapter 3 — implementation stage rows: [no, tahapan, deskripsi]
  implementasi_rows: [
    ["1", "Eksplorasi Kebutuhan", "Mempelajari konteks dan skenario penggunaan bersama tim."],
    ["2", "Perancangan Antarmuka", "Merancang UI dan alur interaksi pengguna."],
    ["3", "Pengembangan Aplikasi", "Mengimplementasikan fitur dan mengintegrasikan komponen."],
    ["4", "Pengujian & Debugging", "Uji coba internal dan perbaikan bug."],
    ["5", "Evaluasi & Dokumentasi", "Evaluasi output dan penyusunan laporan teknis."],
  ],

  // Chapter 3 — technology rows: [kategori, teknologi, fungsi]
  teknologi_rows: [
    ["[Kategori 1]", "[Nama Teknologi]", "[Fungsi/peran dalam proyek]"],
    ["[Kategori 2]", "[Nama Teknologi]", "[Fungsi/peran dalam proyek]"],
    ["[Kategori 3]", "[Nama Teknologi]", "[Fungsi/peran dalam proyek]"],
  ],

  // Chapter 3 — work results rows: [deskripsi_karya, status]
  hasil_karya_rows: [
    ["[Deskripsi karya/proyek 1]", "Selesai"],
    ["[Deskripsi karya/proyek 2]", "On Progress"],
  ],

  // Chapter 3 — thesis title proposals: [no, judul, deskripsi]
  judul_ta_rows: [
    ["1", "[Judul Tugas Akhir Usulan 1]", "[Deskripsi singkat topik dan kaitannya dengan kegiatan magang]"],
    ["2", "[Judul Tugas Akhir Usulan 2]", "[Deskripsi singkat topik]"],
    ["3", "[Judul Tugas Akhir Usulan 3]", "[Deskripsi singkat topik]"],
  ],

  // Chapter 4
  kesimpulan: "[Tuliskan ringkasan pencapaian kegiatan magang: kompetensi yang dikembangkan, proyek yang diselesaikan, pengalaman industri yang diperoleh, dan kaitannya dengan kurikulum program studi.]",
  saran_mahasiswa: "[Saran bagi mahasiswa berikutnya yang akan menjalani magang serupa.]",
  saran_institusi: "[Saran konstruktif bagi institusi mitra terkait pengelolaan program magang.]",
  saran_prodi: "[Saran bagi program studi terkait monitoring, pembekalan, dan dukungan akademik selama MBKM.]",
};

// ─────────────────────────────────────────────────────────────────────────────
// Helpers
// ─────────────────────────────────────────────────────────────────────────────

// A4 dimensions (DXA). Margins: left 4cm, others 3cm (as per pedoman)
const PAGE = {
  size: { width: 11906, height: 16838 },
  margin: {
    top: 1701,    // 3 cm
    right: 1701,  // 3 cm
    bottom: 1701, // 3 cm
    left: 2268,   // 4 cm
  },
};

// Content width (A4 minus margins)
const CONTENT_W = PAGE.size.width - PAGE.margin.left - PAGE.margin.right; // 7937

const CELL_BORDER = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const CELL_BORDERS = {
  top: CELL_BORDER,
  bottom: CELL_BORDER,
  left: CELL_BORDER,
  right: CELL_BORDER,
};
const CELL_MARGINS = { top: 80, bottom: 80, left: 120, right: 120 };
const HEADER_SHADING = { fill: "D9D9D9", type: ShadingType.CLEAR };

function para(text, opts = {}) {
  const align = opts.align || AlignmentType.JUSTIFIED;
  const needsIndent = !opts.noIndent && (align === AlignmentType.JUSTIFIED || align === AlignmentType.LEFT);
  return new Paragraph({
    alignment: align,
    indent: needsIndent ? { firstLine: 720 } : undefined,
    spacing: { before: opts.spaceBefore ?? 120, after: opts.spaceAfter ?? 120, line: opts.line ?? 360 },
    children: [
      new TextRun({
        text,
        font: "Times New Roman",
        size: opts.size ?? 24,
        bold: opts.bold ?? false,
        italics: opts.italic ?? false,
        color: opts.color ?? "000000",
      }),
    ],
    ...(opts.numbering ? { numbering: opts.numbering } : {}),
  });
}

function heading1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 360, after: 180 },
    children: [new TextRun({ text, font: "Times New Roman", size: 24, bold: true })],
  });
}

function heading2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 120 },
    children: [new TextRun({ text, font: "Times New Roman", size: 24, bold: true })],
  });
}

function centered(text, opts = {}) {
  return para(text, { ...opts, align: AlignmentType.CENTER });
}

function pageBreak() {
  return new Paragraph({
    children: [new PageBreak()],
    spacing: { before: 0, after: 0 },
  });
}

function spacer(size = 120) {
  return new Paragraph({ children: [new TextRun("")], spacing: { before: 0, after: size } });
}

function infoRow(label, value) {
  return new TableRow({
    children: [
      new TableCell({
        width: { size: 2200, type: WidthType.DXA },
        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
        margins: CELL_MARGINS,
        children: [new Paragraph({ children: [new TextRun({ text: label, font: "Times New Roman", size: 24 })] })],
      }),
      new TableCell({
        width: { size: 200, type: WidthType.DXA },
        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
        margins: CELL_MARGINS,
        children: [new Paragraph({ children: [new TextRun({ text: ":", font: "Times New Roman", size: 24 })] })],
      }),
      new TableCell({
        width: { size: CONTENT_W - 2400, type: WidthType.DXA },
        borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
        margins: CELL_MARGINS,
        children: [new Paragraph({ children: [new TextRun({ text: value, font: "Times New Roman", size: 24 })] })],
      }),
    ],
  });
}

function infoTable(rows) {
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [2200, 200, CONTENT_W - 2400],
    borders: {
      top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
      left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
      insideH: { style: BorderStyle.NONE }, insideV: { style: BorderStyle.NONE },
    },
    rows: rows.map(([label, value]) => infoRow(label, value)),
  });
}

function dataTable(headers, rows, colWidths) {
  const totalW = colWidths.reduce((a, b) => a + b, 0);
  return new Table({
    width: { size: totalW, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        children: headers.map((h, i) =>
          new TableCell({
            width: { size: colWidths[i], type: WidthType.DXA },
            borders: CELL_BORDERS,
            shading: HEADER_SHADING,
            margins: CELL_MARGINS,
            children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: h, font: "Times New Roman", size: 24, bold: true })] })],
          })
        ),
      }),
      ...rows.map((row) =>
        new TableRow({
          children: row.map((cell, i) =>
            new TableCell({
              width: { size: colWidths[i], type: WidthType.DXA },
              borders: CELL_BORDERS,
              margins: CELL_MARGINS,
              children: [new Paragraph({ children: [new TextRun({ text: cell, font: "Times New Roman", size: 24 })] })],
            })
          ),
        })
      ),
    ],
  });
}

// ─────────────────────────────────────────────────────────────────────────────
// Document sections
// ─────────────────────────────────────────────────────────────────────────────

// COVER PAGE
function buildCover() {
  return [
    spacer(1440),
    centered("Laporan Pelaksanaan Kegiatan MBKM", { bold: true, size: 24 }),
    centered("MBKM Program MSIB / P3NK (Magang Mandiri)", { bold: true, size: 24 }),
    spacer(480),
    centered("Diajukan sebagai salah satu syarat Kegiatan MBKM", { size: 20 }),
    centered("pada Program Studi Sastra Mesin", { size: 20 }),
    spacer(1440),
    centered("[LOGO KAMPUS]", { size: 20, italic: true, color: "888888" }),
    spacer(1440),
    centered("Oleh.", { size: 24 }),
    centered(CONFIG.nama_mahasiswa, { size: 24 }),
    centered(CONFIG.nim, { size: 24 }),
    spacer(1440),
    centered("FAKULTAS ILMU KOMPUTER", { bold: true, size: 28 }),
    centered("TEKNIK INFORMATIKA", { bold: true, size: 28 }),
    centered("UNIVERSITAS INDONESIA", { bold: true, size: 28 }),
    centered(CONFIG.tahun, { bold: true, size: 28 }),
    pageBreak(),
  ];
}

// LEMBAR PENGESAHAN
function buildPengesahan() {
  return [
    spacer(360),
    centered("Laporan Pelaksanaan Kegiatan MBKM", { bold: true, size: 28 }),
    centered("MBKM Program MSIB / P3NK (Magang Mandiri)", { bold: true, size: 28 }),
    spacer(360),
    centered("Lembar Pengesahan", { bold: false, size: 24 }),
    spacer(360),
    centered("Diajukan sebagai salah satu syarat kegiatan MBKM", { size: 20 }),
    centered("pada Program Studi Rekayasa Perangkat Lunak", { size: 20 }),
    spacer(720),
    new Table({
      width: { size: CONTENT_W, type: WidthType.DXA },
      columnWidths: [Math.floor(CONTENT_W / 2), Math.floor(CONTENT_W / 2)],
      borders: {
        top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
        left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
        insideH: { style: BorderStyle.NONE }, insideV: { style: BorderStyle.NONE },
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              width: { size: Math.floor(CONTENT_W / 2), type: WidthType.DXA },
              borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
              children: [
                new Paragraph({ children: [new TextRun({ text: "Dosen Pembimbing,", font: "Times New Roman", size: 24 })] }),
                spacer(960),
                new Paragraph({ children: [new TextRun({ text: CONFIG.nama_dosen_pembimbing, font: "Times New Roman", size: 24, underline: {} })] }),
                new Paragraph({ children: [new TextRun({ text: "NIP. " + CONFIG.nip_dosen, font: "Times New Roman", size: 24 })] }),
              ],
            }),
            new TableCell({
              width: { size: Math.floor(CONTENT_W / 2), type: WidthType.DXA },
              borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
              children: [
                new Paragraph({ children: [new TextRun({ text: "Penyelia,", font: "Times New Roman", size: 24 })] }),
                spacer(960),
                new Paragraph({ children: [new TextRun({ text: CONFIG.nama_penyelia, font: "Times New Roman", size: 24, underline: {} })] }),
                new Paragraph({ children: [new TextRun({ text: "NIP. " + CONFIG.nip_penyelia, font: "Times New Roman", size: 24 })] }),
              ],
            }),
          ],
        }),
      ],
    }),
    spacer(720),
    centered("Mengetahui,", { size: 24 }),
    centered("Ketua Program Studi Rekayasa Perangkat Lunak,", { size: 24 }),
    spacer(960),
    centered(CONFIG.nama_kaprodi, { size: 24 }),
    centered("NIP. " + CONFIG.nip_kaprodi, { size: 24 }),
    pageBreak(),
  ];
}

// KATA PENGANTAR
function buildKataPengantar() {
  return [
    heading1("KATA PENGANTAR"),
    para(
      "Puji dan syukur penulis panjatkan ke hadirat Tuhan Yang Maha Esa atas limpahan rahmat dan karunia-Nya sehingga laporan akhir magang ini dapat diselesaikan dengan baik dan tepat waktu. Laporan ini disusun sebagai bentuk pertanggungjawaban atas pelaksanaan kegiatan magang yang merupakan salah satu syarat akademik dalam menyelesaikan program studi di Rekayasa Perangkat Lunak, Universitas Pendidikan Indonesia Kampus Cibiru."
    ),
    para(
      `Selama proses magang yang dilaksanakan di ${CONFIG.nama_institusi_mitra}, banyak pengalaman berharga yang diperoleh—tidak hanya memperluas wawasan dan keterampilan teknis, tetapi juga meningkatkan pemahaman terhadap praktik kerja profesional secara langsung.`
    ),
    para("Ucapan terima kasih yang sebesar-besarnya disampaikan kepada:"),
    new Paragraph({
      numbering: { reference: "numbers", level: 0 },
      spacing: { before: 60, after: 60 },
      children: [new TextRun({ text: `Bapak/Ibu ${CONFIG.nama_dosen_pembimbing} selaku dosen pembimbing yang telah memberikan arahan dan bimbingan selama proses penyusunan laporan ini;`, font: "Times New Roman", size: 24 })],
    }),
    new Paragraph({
      numbering: { reference: "numbers", level: 0 },
      spacing: { before: 60, after: 60 },
      children: [new TextRun({ text: `Bapak/Ibu ${CONFIG.nama_penyelia} selaku pembimbing di instansi tempat magang yang telah membimbing dan memberikan kesempatan untuk belajar secara langsung di lingkungan kerja;`, font: "Times New Roman", size: 24 })],
    }),
    new Paragraph({
      numbering: { reference: "numbers", level: 0 },
      spacing: { before: 60, after: 60 },
      children: [new TextRun({ text: `Seluruh staf dan karyawan di ${CONFIG.nama_institusi_mitra} atas bantuan, kerja sama, serta pengalaman yang telah dibagikan;`, font: "Times New Roman", size: 24 })],
    }),
    new Paragraph({
      numbering: { reference: "numbers", level: 0 },
      spacing: { before: 60, after: 180 },
      children: [new TextRun({ text: "Serta keluarga dan teman-teman yang telah memberikan dukungan dan semangat selama pelaksanaan magang hingga penyusunan laporan ini.", font: "Times New Roman", size: 24 })],
    }),
    para(
      "Laporan ini masih jauh dari sempurna. Oleh karena itu, kritik dan saran yang membangun sangat diharapkan demi penyempurnaan laporan ini di masa mendatang. Semoga laporan ini dapat memberikan manfaat bagi semua pihak yang membacanya."
    ),
    spacer(360),
    para(`Bandung, ............... ${CONFIG.tahun}`, { align: AlignmentType.RIGHT }),
    spacer(720),
    para(CONFIG.nama_mahasiswa, { align: AlignmentType.RIGHT }),
    para(`NIM. ${CONFIG.nim}`, { align: AlignmentType.RIGHT }),
    pageBreak(),
  ];
}

// DAFTAR ISI (manual — placeholder)
function buildDaftarIsi() {
  return [
    heading1("DAFTAR ISI"),
    para("LEMBAR PENGESAHAN ...............................................................  i", { spaceBefore: 60, spaceAfter: 40 }),
    para("KATA PENGANTAR .................................................................  ii", { spaceBefore: 40, spaceAfter: 40 }),
    para("DAFTAR ISI .....................................................................  iii", { spaceBefore: 40, spaceAfter: 40 }),
    para("DAFTAR TABEL ..................................................................  iv", { spaceBefore: 40, spaceAfter: 40 }),
    para("DAFTAR GAMBAR .................................................................  v", { spaceBefore: 40, spaceAfter: 40 }),
    para("DAFTAR LAMPIRAN ...............................................................  vi", { spaceBefore: 40, spaceAfter: 60 }),
    para("BAB I PENDAHULUAN .............................................................  1", { spaceBefore: 60, spaceAfter: 40 }),
    para("   1.1 Latar Belakang .........................................................  1", { spaceBefore: 40, spaceAfter: 40 }),
    para("   1.2 Manfaat & Tujuan .......................................................  1", { spaceBefore: 40, spaceAfter: 40 }),
    para("   1.3 Waktu dan Tempat Pelaksanaan ...........................................  2", { spaceBefore: 40, spaceAfter: 40 }),
    para("   1.4 Ruang Lingkup Kegiatan MBKM ............................................  2", { spaceBefore: 40, spaceAfter: 60 }),
    para("BAB II PROFIL INSTITUSI MITRA .................................................  3", { spaceBefore: 60, spaceAfter: 40 }),
    para("   2.1 Gambaran Umum Institusi Mitra ..........................................  3", { spaceBefore: 40, spaceAfter: 40 }),
    para("   2.2 Bidang Kerja ...........................................................  3", { spaceBefore: 40, spaceAfter: 40 }),
    para("   2.3 Peran Mahasiswa dalam Mitra ............................................  3", { spaceBefore: 40, spaceAfter: 40 }),
    para("   2.4 Jadwal Kegiatan ........................................................  4", { spaceBefore: 40, spaceAfter: 60 }),
    para("BAB III PELAKSANAAN KEGIATAN ..................................................  5", { spaceBefore: 60, spaceAfter: 40 }),
    para("   3.1 Rencana Kegiatan .......................................................  5", { spaceBefore: 40, spaceAfter: 40 }),
    para("   3.2 Implementasi Kegiatan ..................................................  5", { spaceBefore: 40, spaceAfter: 40 }),
    para("   3.3 Teknologi dan Metode yang Diterapkan ...................................  6", { spaceBefore: 40, spaceAfter: 40 }),
    para("   3.4 Hasil Karya ............................................................  7", { spaceBefore: 40, spaceAfter: 40 }),
    para("   3.5 List untuk Judul Tugas Akhir ...........................................  8", { spaceBefore: 40, spaceAfter: 60 }),
    para("BAB IV KESIMPULAN DAN SARAN ...................................................  9", { spaceBefore: 60, spaceAfter: 40 }),
    para("   4.1 Kesimpulan .............................................................  9", { spaceBefore: 40, spaceAfter: 40 }),
    para("   4.2 Saran .................................................................  9", { spaceBefore: 40, spaceAfter: 40 }),
    para("LAMPIRAN ......................................................................  11", { spaceBefore: 60, spaceAfter: 40 }),
    pageBreak(),
  ];
}

// BAB I
function buildBab1() {
  return [
    heading1("BAB I PENDAHULUAN"),
    heading2("1.1 Latar Belakang"),
    para(CONFIG.latar_belakang),
    spacer(120),
    heading2("1.2 Manfaat & Tujuan"),
    para("Manfaat yang diharapkan dari kegiatan ini antara lain:"),
    new Paragraph({ numbering: { reference: "latin-lower", level: 0 }, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Mahasiswa memperoleh pengalaman kerja langsung dalam industri digital berbasis proyek nyata.", font: "Times New Roman", size: 24 })] }),
    new Paragraph({ numbering: { reference: "latin-lower", level: 0 }, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Meningkatkan kompetensi teknis sesuai dengan bidang yang dikerjakan selama magang.", font: "Times New Roman", size: 24 })] }),
    new Paragraph({ numbering: { reference: "latin-lower", level: 0 }, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Mengembangkan soft skill seperti komunikasi, kerja tim, problem solving, dan manajemen waktu.", font: "Times New Roman", size: 24 })] }),
    new Paragraph({ numbering: { reference: "latin-lower", level: 0 }, spacing: { before: 60, after: 120 }, children: [new TextRun({ text: "[Tambahkan manfaat lain yang relevan dengan konteks magang Anda]", font: "Times New Roman", size: 24, color: "888888" })] }),
    para("Tujuan dari pelaksanaan kegiatan ini adalah:"),
    new Paragraph({ numbering: { reference: "latin-lower2", level: 0 }, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Meningkatkan penguasaan kompetensi mahasiswa dalam rekayasa perangkat lunak.", font: "Times New Roman", size: 24 })] }),
    new Paragraph({ numbering: { reference: "latin-lower2", level: 0 }, spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "Mewujudkan pembelajaran berbasis pengalaman (experiential learning) dalam skema MBKM.", font: "Times New Roman", size: 24 })] }),
    new Paragraph({ numbering: { reference: "latin-lower2", level: 0 }, spacing: { before: 60, after: 120 }, children: [new TextRun({ text: "[Tambahkan tujuan lain yang relevan]", font: "Times New Roman", size: 24, color: "888888" })] }),
    spacer(120),
    heading2("1.3 Waktu dan Tempat Pelaksanaan"),
    para("Kegiatan magang dilaksanakan dengan rincian sebagai berikut:"),
    spacer(60),
    infoTable([
      ["Periode Pelaksanaan", CONFIG.periode_pelaksanaan],
      ["Durasi Kegiatan", CONFIG.durasi_kegiatan],
      ["Hari dan Jam Kerja", CONFIG.hari_jam_kerja],
      ["Lokasi Magang", CONFIG.lokasi_magang],
      ["Alamat Lokasi", CONFIG.alamat_lokasi],
      ["Unit/Tim Penempatan", CONFIG.unit_penempatan],
    ]),
    spacer(120),
    heading2("1.4 Ruang Lingkup Kegiatan MBKM"),
    para(CONFIG.ruang_lingkup),
    pageBreak(),
  ];
}

// BAB II
function buildBab2() {
  const jadwalRows = [
    ["Senin s.d. Jumat", "08.00 – 17.00 WIB", "Pengembangan proyek, diskusi tim, dan evaluasi"],
    ["[Kelas/Sesi Khusus]", "[Waktu]", "[Deskripsi kegiatan tambahan seperti kelas soft skill/hard skill]"],
  ];

  return [
    heading1("BAB II PROFIL INSTITUSI MITRA"),
    heading2("2.1 Gambaran Umum Institusi Mitra"),
    para(CONFIG.gambaran_umum_institusi),
    spacer(120),
    heading2("2.2 Bidang Kerja"),
    para(CONFIG.bidang_kerja),
    spacer(120),
    heading2("2.3 Peran Mahasiswa dalam Mitra"),
    para(CONFIG.peran_mahasiswa),
    spacer(120),
    heading2("2.4 Jadwal Kegiatan"),
    para(`Kegiatan magang dilaksanakan selama ${CONFIG.durasi_kegiatan}, dengan jadwal kegiatan sebagai berikut:`),
    spacer(60),
    dataTable(["Hari", "Waktu", "Kegiatan"], jadwalRows, [2000, 2000, CONTENT_W - 4000]),
    pageBreak(),
  ];
}

// BAB III
function buildBab3() {
  const metodeRows = CONFIG.teknologi_rows;

  return [
    heading1("BAB III PELAKSANAAN KEGIATAN"),
    heading2("3.1 Rencana Kegiatan"),
    para(`Selama mengikuti program magang di ${CONFIG.nama_institusi_mitra}, tanggung jawab utama berada pada ${CONFIG.internship_role}. Adapun rencana kegiatan yang dilakukan selama program magang dirangkum dalam tabel berikut:`),
    spacer(60),
    dataTable(["Tahapan Kegiatan", "Deskripsi Aktivitas"], CONFIG.rencana_kegiatan_rows, [Math.floor(CONTENT_W * 0.35), Math.floor(CONTENT_W * 0.65)]),
    spacer(120),
    heading2("3.2 Implementasi Kegiatan"),
    para(`Implementasi kegiatan magang dilaksanakan secara bertahap sesuai rencana yang telah disusun. Proses kerja dijalani secara penuh waktu di ${CONFIG.lokasi_magang}, dengan keterlibatan langsung dalam siklus pengembangan yang berlangsung secara iteratif menggunakan pendekatan agile, dengan evaluasi progres dilakukan setiap minggu.`),
    spacer(60),
    dataTable(["No", "Tahapan Implementasi", "Deskripsi Aktivitas"], CONFIG.implementasi_rows, [500, Math.floor(CONTENT_W * 0.35), Math.floor(CONTENT_W * 0.65) - 500]),
    spacer(120),
    heading2("3.3 Teknologi dan Metode yang Diterapkan Selama Kegiatan"),
    para("Selama pelaksanaan kegiatan magang, berbagai teknologi dan metode pengembangan diterapkan dalam rangka mendukung proses pembuatan produk/layanan. Teknologi utama yang diterapkan dalam proyek antara lain:"),
    spacer(60),
    dataTable(["Kategori", "Teknologi/Tools", "Fungsi"], metodeRows, [Math.floor(CONTENT_W * 0.25), Math.floor(CONTENT_W * 0.35), Math.floor(CONTENT_W * 0.40)]),
    spacer(120),
    para("[Uraikan metode pengembangan yang digunakan, misalnya Design Thinking, Agile/Scrum, atau metodologi lain. Jelaskan tiap tahapannya dan bagaimana diterapkan dalam proyek selama magang.]", { color: "888888", italic: true }),
    spacer(120),
    heading2("3.4 Hasil Karya"),
    para(`Selama masa magang di ${CONFIG.nama_institusi_mitra}, berikut merupakan ringkasan pencapaian dan kontribusi yang dihasilkan:`),
    spacer(60),
    dataTable(["Deskripsi Karya", "Status"], CONFIG.hasil_karya_rows, [Math.floor(CONTENT_W * 0.75), Math.floor(CONTENT_W * 0.25)]),
    spacer(120),
    para("[Tambahkan deskripsi naratif mengenai setiap hasil karya: teknologi yang digunakan, tantangan yang dihadapi, dan dampaknya bagi institusi mitra atau pengguna akhir.]", { color: "888888", italic: true }),
    spacer(120),
    heading2("3.5 List untuk Judul Tugas Akhir"),
    para("Berdasarkan pengalaman dan kegiatan yang telah dilakukan selama magang, berikut adalah beberapa usulan judul Tugas Akhir yang relevan untuk dikembangkan lebih lanjut:"),
    spacer(60),
    dataTable(["No", "Judul Tugas Akhir", "Deskripsi Singkat"], CONFIG.judul_ta_rows, [500, Math.floor(CONTENT_W * 0.45), Math.floor(CONTENT_W * 0.55) - 500]),
    pageBreak(),
  ];
}

// BAB IV
function buildBab4() {
  return [
    heading1("BAB IV KESIMPULAN DAN SARAN"),
    heading2("4.1 Kesimpulan"),
    para(CONFIG.kesimpulan),
    spacer(120),
    heading2("4.2 Saran"),
    para("Saran berikut disusun berdasarkan refleksi atas pelaksanaan kegiatan magang, mencakup perspektif dari beberapa pihak yang terlibat."),
    spacer(60),
    new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "a. Bagi Mahasiswa", font: "Times New Roman", size: 24, bold: true })] }),
    para(CONFIG.saran_mahasiswa),
    new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "b. Bagi Institusi Mitra", font: "Times New Roman", size: 24, bold: true })] }),
    para(CONFIG.saran_institusi),
    new Paragraph({ spacing: { before: 60, after: 60 }, children: [new TextRun({ text: "c. Bagi Program Studi", font: "Times New Roman", size: 24, bold: true })] }),
    para(CONFIG.saran_prodi),
    pageBreak(),
  ];
}

// LAMPIRAN placeholder
function buildLampiran() {
  return [
    heading1("LAMPIRAN"),
    para("Lampiran 1. Laporan Weekly / Biweekly", { bold: true }),
    para("[Tempel laporan weekly di sini atau lampirkan sebagai file terpisah]", { color: "888888", italic: true }),
    spacer(360),
    para("Lampiran 2. Screenshot Hasil Pekerjaan", { bold: true }),
    para("[Lampirkan screenshot hasil kerja: tangkapan layar IDE, aplikasi, deployment, dsb.]", { color: "888888", italic: true }),
    spacer(360),
    para("Lampiran 3. Bukti Keikutsertaan dalam Proyek Perusahaan", { bold: true }),
    para("[Lampirkan bukti keterlibatan: foto rapat, screenshot komunikasi tim, dsb.]", { color: "888888", italic: true }),
    spacer(360),
    para("Lampiran 4. Form Penilaian Penyelia", { bold: true }),
    para("[Lampirkan form penilaian penyelia yang telah ditandatangani]", { color: "888888", italic: true }),
    spacer(360),
    para("Lampiran 5. Sertifikat Magang (jika ada)", { bold: true }),
    para("[Lampirkan sertifikat keikutsertaan atau kompetensi yang diperoleh selama magang]", { color: "888888", italic: true }),
  ];
}

// ─────────────────────────────────────────────────────────────────────────────
// Assemble & write
// ─────────────────────────────────────────────────────────────────────────────

const children = [
  ...buildCover(),
  ...buildPengesahan(),
  ...buildKataPengantar(),
  ...buildDaftarIsi(),
  ...buildBab1(),
  ...buildBab2(),
  ...buildBab3(),
  ...buildBab4(),
  ...buildLampiran(),
];

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
      },
      {
        reference: "latin-lower",
        levels: [{ level: 0, format: LevelFormat.LOWER_LETTER, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
      },
      {
        reference: "latin-lower2",
        levels: [{ level: 0, format: LevelFormat.LOWER_LETTER, text: "%1.", alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }],
      },
    ],
  },
  styles: {
    default: {
      document: { run: { font: "Times New Roman", size: 24 } },
    },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 24, bold: true, font: "Times New Roman", color: "000000" },
        paragraph: { spacing: { before: 360, after: 180 }, outlineLevel: 0 },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { size: 24, bold: true, font: "Times New Roman", color: "000000" },
        paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 1 },
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: PAGE.size,
          margin: PAGE.margin,
        },
        pageNumberStart: 1,
        pageNumberFormatType: NumberFormat.DECIMAL,
      },
      footers: {
        default: new Footer({
          children: [
            new Paragraph({
              alignment: AlignmentType.RIGHT,
              spacing: { before: 0, after: 0 },
              children: [
                new TextRun({
                  font: "Times New Roman",
                  size: 24,
                  children: [PageNumber.CURRENT],
                }),
              ],
            }),
          ],
        }),
      },
      children,
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  const outPath = "Laporan_MBKM_Template.docx";
  fs.writeFileSync(outPath, buffer);
  console.log(`✓ Generated: ${outPath}`);
}).catch((err) => {
  console.error("Error generating DOCX:", err);
  process.exit(1);
});
