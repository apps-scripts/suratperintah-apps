<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  
  <!-- CSS LIBRARIES -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/lottie-web/5.9.6/lottie.min.js"></script>
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
  <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/select2-bootstrap-5-theme@1.3.0/dist/select2-bootstrap-5-theme.min.css" />
  
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/flatpickr/dist/flatpickr.min.css">
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.0/font/bootstrap-icons.css">
  <link rel="stylesheet" href="https://cdn.datatables.net/1.13.4/css/dataTables.bootstrap5.min.css">
  
  <!-- Virtual Select CSS & JS -->
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/virtual-select-plugin@1.0.39/dist/virtual-select.min.css">
  <script src="https://cdn.jsdelivr.net/npm/virtual-select-plugin@1.0.39/dist/virtual-select.min.js"></script>

  <!-- LIBRARIES UNTUK EXCEL & PDF -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/jspdf-autotable/3.5.28/jspdf.plugin.autotable.min.js"></script>

  <style>
    /* --- SEARCH BAR MODERN SEPERTI GAMBAR --- */
    .custom-search-container { display: flex; align-items: center; background: #ffffff; border-radius: 50px; padding: 4px 6px 4px 18px; border: 1px solid #e2e8f0; box-shadow: 0 2px 10px rgba(0,0,0,0.03); transition: all 0.3s ease; width: 350px; }
    .custom-search-container:focus-within { border-color: #38a169; box-shadow: 0 0 0 0.25rem rgba(56, 161, 105, 0.2); }
    .custom-search-input { flex-grow: 1; border: none !important; background: transparent !important; padding: 8px 0; outline: none !important; box-shadow: none !important; color: #4a5568; font-size: 0.9rem; }
    .custom-search-input::placeholder { color: #a0aec0; }
    .custom-search-btn { background-color: #4caf50; color: white; border: none; border-radius: 50%; width: 38px; height: 38px; display: flex; justify-content: center; align-items: center; cursor: pointer; transition: background 0.3s, transform 0.2s; flex-shrink: 0; }
    .custom-search-btn:hover { background-color: #388e3c; transform: scale(1.05); }

    /* --- THEME TANGGAL MODERN (PURPLE) --- */
/* --- COMPACT MODERN DATEPICKER (UKURAN DISESUAIKAN) --- */
.flatpickr-calendar {
  width: 280px !important; /* Diperkecil dari 320px agar pas dengan kolom */
  border-radius: 16px !important;
  padding: 8px !important;
  box-shadow: 0 10px 25px rgba(0,0,0,0.1) !important;
  border: 1px solid #e2e8f0 !important;
  font-size: 0.85rem !important; /* Font lebih kecil agar tidak sesak */
}

/* Container hari */
.flatpickr-days, 
.dayContainer {
  width: 260px !important;
  min-width: 260px !important;
  max-width: 260px !important;
}

/* Header: Bulan & Tahun */
.flatpickr-current-month {
  padding: 0 !important;
  font-size: 0.95rem !important;
}

.flatpickr-months .flatpickr-prev-month, 
.flatpickr-months .flatpickr-next-month {
  height: 30px !important;
  width: 30px !important;
  padding: 5px !important;
  top: 10px !important;
}

/* Ukuran Kotak Tanggal (Kunci agar tidak mepet) */
.flatpickr-day {
  max-width: 34px !important; /* Diperkecil dari 38px */
  height: 34px !important;
  line-height: 34px !important;
  margin: 1px !important; /* Jarak antar angka */
  border-radius: 8px !important;
}

/* Tanggal Terpilih (Bulat Sempurna) */
.flatpickr-day.selected {
  background: #5850eb !important;
  border-color: #5850eb !important;
  border-radius: 50% !important;
  font-weight: 700;
}

/* Label Hari (Sen, Sel, Rab...) */
span.flatpickr-weekday {
  font-size: 0.75rem !important;
  font-weight: 600 !important;
  color: #94a3b8 !important;
}

/* Tombol Set Date (Lebih ramping) */
.flatpickr-confirm {
  background: #5850eb !important;
  padding: 8px 0 !important; /* Padding lebih tipis */
  font-size: 0.85rem !important;
  border-radius: 10px !important;
  margin: 10px 5px 5px 5px !important;
  box-shadow: 0 4px 10px rgba(88, 80, 235, 0.2);
}

/* Warna Hari Libur */
.flatpickr-day.tanggal-merah {
  color: #ef4444 !important;
}

    /* --- MODERN VIRTUAL SELECT CUSTOMIZATION --- */
    .vscomp-wrapper .vscomp-checkbox-icon { border-radius: 6px !important; border: 2px solid #dcdcdc !important; width: 20px !important; height: 20px !important; overflow: hidden; transition: all 0.2s ease-in-out !important; box-sizing: border-box !important; }
    .vscomp-wrapper .vscomp-option.selected .vscomp-checkbox-icon, .vscomp-wrapper .vscomp-toggle-all-container.selected .vscomp-checkbox-icon { background-color: #333333 !important; border-color: #333333 !important; border-radius: 6px !important; }
    .vscomp-wrapper .vscomp-checkbox-icon::after { content: ''; position: absolute; border-color: #ffffff !important; border-style: solid !important; border-width: 0 2.5px 2.5px 0 !important; width: 5px !important; height: 10px !important; transform: rotate(45deg) translate(-50%, -50%); left: 45% !important; top: 5% !important; display: none; }
    .vscomp-wrapper .vscomp-option.selected .vscomp-checkbox-icon::after, .vscomp-wrapper .vscomp-toggle-all-container.selected .vscomp-checkbox-icon::after { display: block !important; }
    .vscomp-wrapper :focus { outline: none !important; }

    /* --- GLOBAL & SIDEBAR --- */
    body { font-family: Arial, Helvetica, sans-serif !important; background-color: #f4f6f9; }
    .sidebar { background: white; height: 100vh; overflow-y: auto; box-shadow: 2px 0 10px rgba(0,0,0,0.05); padding: 20px; }
    .form-label { font-weight: 600; font-size: 0.9rem; color: #495057; }
    .row-item { background: #f8f9fa; padding: 10px; border-radius: 6px; margin-bottom: 8px; border: 1px solid #e9ecef; }
    
    /* --- LOTTIE LOADING CSS ADJUSTED --- */
    #loading { 
      position: fixed; top: 0; left: 0; width: 100%; height: 100%; 
      background-color: #ffffff; z-index: 9999; 
      display: flex; justify-content: center; align-items: center; 
      font-family: Arial, Helvetica, sans-serif; 
    }
    .loader-content { 
      text-align: center; 
      width: 100%; 
      max-width: 350px; 
      display: flex;
      flex-direction: column;
      align-items: center;
    }
    
    #lottie-container {
      width: 180px;
      height: 180px;
      margin-bottom: 20px; /* Jarak antara pesawat dengan angka */
    }

    .percentage { 
      color: #1e293b; 
      font-size: 2rem; 
      font-weight: 800; 
      margin-bottom: 25px; /* MENAMBAH JARAK ANTARA ANGKA DAN PROGRESS BAR */
      line-height: 1;
      font-family: Arial, sans-serif;
    }

    .progress-container { 
      width: 100%; 
      padding: 0 40px; 
    }
    .progress-track { 
      width: 100%; 
      height: 10px; 
      background: #f1f5f9 !important; /* Warna dasar track (abu muda) */
      border-radius: 20px; 
      overflow: hidden; 
    }
    /* KUNCI: Mematikan paksa pesawat tambahan di ujung progress bar */
    .progress-fill::after, 
    .progress-fill::before {
        display: none !important;
        content: none !important;
    }

    /* Pastikan progress-fill bersih dari gambar latar pesawat */
    .progress-fill { 
      width: 0%; 
      height: 100%; 
      background-color: #5850eb !important; /* WARNA UNGU INDIGO BALIK LAGI */
      border-radius: 20px; 
      transition: width 0.4s ease; 
      display: block !important;
      background-image: none !important; /* Hapus pesawat di ujung garis */
    }

    .status-text { 
      color: #94a3b8; 
      font-size: 0.9rem; 
      letter-spacing: 1px; 
      margin-top: 25px;
      font-weight: 700;
      display: flex;
      align-items: center;
      gap: 3px;
      animation: textPulse 2s infinite ease-in-out; /* Efek teks berdenyut */
    }

     /* Animasi Titik Melompat */
    .loading-dots span {
      display: inline-block;
      font-size: 1.5rem;
      line-height: 0;
      color: #5850eb; /* Warna Indigo sesuai tema */
      animation: dotJump 1.4s infinite;
    }

    .loading-dots span:nth-child(2) { animation-delay: 0.2s; }
    .loading-dots span:nth-child(3) { animation-delay: 0.4s; }

    /* Keyframes Berdenyut */
    @keyframes textPulse {
      0%, 100% { opacity: 0.7; transform: scale(1); }
      50% { opacity: 1; transform: scale(1.03); color: #6366f1; }
    }

    /* Keyframes Titik Melompat */
    @keyframes dotJump {
      0%, 80%, 100% { transform: translateY(0); }
      40% { transform: translateY(-6px); }
    }
    
   /* --- MODERN DASHBOARD TABLE CSS (Sesuai Gambar) --- */
    :root {
      --primary-indigo: #5850eb;
      --success-emerald: #10b981;
      --danger-rose: #ef4444;
      --border-color: #e2e8f0;
    }

    .modal-content { border-radius: 12px; border: none; overflow: hidden; }
    .modal-header { border-bottom: 1px solid #f1f5f9; padding: 20px 24px; }
    
    /* --- TABEL IDENTIK DENGAN GAMBAR (DASHBOARD STYLE) --- */
    .table-container-modern {
      background: white;
      border-radius: 12px;
      border: 1px solid #e2e8f0;
      overflow: hidden;
      box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05);
    }

    .table-modern {
      width: 100% !important;
      border-collapse: separate;
      border-spacing: 0;
    }

    /* Header Tabel */
    .table-modern thead th {
      background-color: #f8fafc; /* Abu-abu sangat muda */
      color: #94a3b8; /* Teks abu-abu khas dashboard */
      font-weight: 700;
      font-size: 11px;
      text-transform: uppercase;
      letter-spacing: 0.05em;
      padding: 16px 20px;
      border-bottom: 1px solid #e2e8f0;
    }

    /* Baris/Body Tabel */
    .table-modern tbody td {
      /* 12px adalah jarak atas-bawah, 20px adalah jarak kiri-kanan */
      padding: 12px 20px !important; 
      font-size: 13px;
      color: #475569;
      vertical-align: middle;
      border-bottom: 1px solid #f1f5f9;
    }

    /* Nomor Registrasi (Warna Indigo Bold) */
    .td-reg-number {
      color: #5850eb;
      font-weight: 700;
    }

    /* Badge PIC (Kapsul Abu) */
    .badge-pic-capsule {
      background-color: #f8fafc;
      color: #1e293b;
      padding: 4px 10px;
      border-radius: 6px;
      font-weight: 700;
      font-size: 11px;
      border: 1px solid #e2e8f0;
    }

    /* Badge Status (Pill Hijau dengan Centang) */
    .status-pill-success {
      background-color: #ecfdf5;
      color: #059669;
      padding: 6px 14px;
      border-radius: 50px;
      font-weight: 600;
      font-size: 12px;
      display: inline-flex;
      align-items: center;
      gap: 6px;
    }

    /* Tombol Aksi Bulat Minimalis */
    .btn-action-circle {
      width: 32px;
      height: 32px;
      border-radius: 50%;
      display: inline-flex;
      align-items: center;
      justify-content: center;
      background: white;
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      border: 1.5px solid #e2e8f0; /* Border tipis saat normal */
      padding: 0;
      cursor: pointer;
    }

    .btn-action-circle i {
      font-size: 14px;
      transition: color 0.3s ease;
    }

    .btn-action-edit { color: #3b82f6; border-color: #3b82f6; }
    .btn-action-edit:hover { background: #eff6ff; }
    
    .btn-action-delete { color: #ef4444; border-color: #ef4444; }
    .btn-action-delete:hover { background: #fef2f2; }

    /* Mencegah penumpukan teks */
    .uraian-text {
      min-width: 250px;
      line-height: 1.6;
      text-align: left;
    }
    .btn-action-circle:hover {
      transform: translateY(-3px);
      box-shadow: 0 4px 8px rgba(0, 0, 0, 0.15);
      color: white !important; /* Semua ikon jadi putih saat bg berwarna */
    }

    /* Warna Edit (Orange/Amber) */
    .btn-edit {
      color: #f59e0b;
      border-color: #f59e0b;
    }
    .btn-edit:hover {
      background-color: #f59e0b !important;
    }

    /* Warna Cetak (Green/Emerald) */
    .btn-print {
      color: #10b981;
      border-color: #10b981;
    }
    .btn-print:hover {
      background-color: #10b981 !important;
    }

    /* Warna Hapus (Red/Rose) */
    .btn-delete {
      color: #ef4444;
      border-color: #ef4444;
    }
    .btn-delete:hover {
      background-color: #ef4444 !important;
    }
    .btn-action-circle.delete:hover { border-color: var(--danger-rose); color: var(--danger-rose); }

    /* Custom DataTables Pagination & Footer */
    .dataTables_wrapper .row:last-child {
      background: #f8fafc;
      padding: 15px 20px;
      border-top: 1px solid var(--border-color);
      margin: 0;
      align-items: center;
    }

    .dataTables_wrapper .dataTables_scrollBody,
    table.dataTable.no-footer {
      border-bottom: none !important;
    }
    .dataTables_info { font-size: 0.8rem; color: #64748b !important; padding: 0 !important; }
    .pagination .page-link {
      border: 1px solid #e2e8f0;
      margin: 0 2px;
      border-radius: 6px !important;
      font-size: 0.8rem;
      color: #475569;
      padding: 6px 12px;
    }
    .pagination .active .page-link {
      background-color: var(--primary-indigo) !important;
      border-color: var(--primary-indigo) !important;
      color: white !important;
    }
    
    /* entries dropdown (Tampilkan: 5) */
    .dataTables_length label { font-size: 0.85rem; color: #475569; display: flex; align-items: center; gap: 8px; }
    .dataTables_length select {
      border: 1px solid var(--border-color);
      border-radius: 8px;
      padding: 4px 30px 4px 10px;
      color: var(--primary-indigo);
      font-weight: 600;
    }

    /* --- FORM CUSTOM STYLES (HIJAU) --- */
    .form-control, .form-select { border: 1px solid #198754 !important; }
    .form-control:focus, .form-select:focus { border-color: #146c43 !important; box-shadow: 0 0 0 0.25rem rgba(25, 135, 84, 0.25) !important; }
    .select2-container--bootstrap-5 .select2-selection { border-color: #198754 !important; }
    .select2-container--bootstrap-5.select2-container--focus .select2-selection,
    .select2-container--bootstrap-5.select2-container--open .select2-selection { border-color: #146c43 !important; box-shadow: 0 0 0 0.25rem rgba(25, 135, 84, 0.25) !important; }
    .input-group-text { border-color: #198754 !important; background-color: #e8f5e9 !important; color: #198754 !important; }
    #selectPegawaiVirtual { border: 1px solid #198754 !important; border-radius: 6px; max-width: 100%; }
    .vscomp-toggle-button { padding: 8px 12px; border-radius: 6px; }
    .vscomp-option.selected .vscomp-checkbox-icon::after { background-color: #198754 !important; border-color: #198754 !important; }
    .vscomp-dropbox-container { border-radius: 8px; box-shadow: 0 10px 25px rgba(0,0,0,0.1); }

    /* --- ESTETIKA MODAL FORM (CLEAN & ELEGANT) --- */
    .custom-swal-modal { border-radius: 24px !important; box-shadow: 0 20px 40px rgba(0,0,0,0.15) !important; overflow: hidden !important; /* Mengunci semua efek visual di dalam kotak modal */
      transform: translateZ(0); padding: 30px 20px !important; font-family: Arial, Helvetica, sans-serif !important; }
    .custom-swal-title { font-weight: 700 !important; font-size: 1.5rem !important; color: #1e293b !important; text-align: left !important; margin-bottom: 5px !important; padding: 0 10px !important; }
    .custom-swal-content { text-align: left !important; margin: 0 !important; }
    .form-floating > label { color: #64748b; font-size: 0.9rem; }
    .form-floating > .form-control, .form-floating > .form-select { height: calc(3.5rem + 2px); line-height: 1.25; border-radius: 12px !important; background-color: #f8fafc; border: 1px solid #e2e8f0 !important; }
    .form-floating > .form-control:focus, .form-floating > .form-select:focus { background-color: #ffffff; border-color: #10b981 !important; box-shadow: 0 0 0 4px rgba(16, 185, 129, 0.1) !important; }
    
    /* Nav Tabs Custom untuk Swal */
    .custom-tabs { border-bottom: 2px solid #e2e8f0; margin-bottom: 20px; display: flex; gap: 20px; padding: 0 10px; }
    /* --- MODERN GRADIENT BUTTONS --- */
     .btn-gradient {
      padding: 12px 15px !important; /* Tombol sedikit lebih tinggi */
      font-size: 0.85rem !important;
    }

    .btn-gradient:hover {
      transform: translateY(-3px); /* Efek melayang saat hover */
      box-shadow: 0 8px 20px rgba(0, 0, 0, 0.15) !important;
      filter: brightness(1.1);
    }

    .btn-gradient:active {
      transform: translateY(-1px);
    }

    /* Warna Gradasi Biru (Untuk Master Data) */
    .btn-blue-grad {
      background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%) !important;
    }

    /* Warna Gradasi Hijau (Untuk Riwayat SP) */
    .btn-green-grad {
      background: linear-gradient(135deg, #0ba360 0%, #3cba92 100%) !important;
    }

    /* Warna Gradasi Indigo/Ungu (Untuk SP Baru) */
    .btn-purple-grad {
      background: linear-gradient(135deg, #667eea 0%, #764ba2 100%) !important;
    }

    /* Penyesuaian ukuran icon */
    .btn-gradient i {
      font-size: 1.1rem;
    }
    .tab-content { display: none; }
    .tab-content.active { display: block; animation: fadeIn 0.3s ease-in-out; }
    @keyframes fadeIn { from { opacity: 0; transform: translateY(5px); } to { opacity: 1; transform: translateY(0); } }

    /* Upload Area Styling */
    .file-drop-area { border: 2px dashed #cbd5e0; border-radius: 16px; padding: 40px 20px; text-align: center; background-color: #f8fafc; cursor: pointer; transition: all 0.3s ease; }
    .file-drop-area:hover { border-color: #10b981; background-color: #f0fdf4; }
    .file-drop-area i { font-size: 3rem; color: #94a3b8; margin-bottom: 10px; transition: color 0.3s; }
    .file-drop-area:hover i { color: #10b981; }
    .file-msg { display: block; font-size: 0.95rem; color: #475569; font-weight: 500; }
    .file-submsg { display: block; font-size: 0.8rem; color: #94a3b8; margin-top: 5px; }
    .file-input {
      position: absolute;
      left: 0;
      top: 0;
      height: 100%;
      width: 100%;
      opacity: 0;
      cursor: pointer;
      z-index: 10; /* Berada di atas teks instruksi upload */
    }
    .file-drop-wrapper {
      position: relative;
      margin-bottom: 20px;
      overflow: hidden; /* Mencegah input file meluber keluar */
      border-radius: 16px;
    }

    /* KUNCI: Memberikan z-index lebih tinggi pada tombol template agar tidak tertutup */
    .btn-template-excel {
      position: relative;
      z-index: 20 !important; /* HARUS lebih tinggi dari .file-input */
      border: 1px solid #10b981 !important;
      color: #10b981 !important;
      background: white !important;
    }
    
    .btn-template-excel:hover {
      background: #f0fdf4 !important;
    }

    .nav-control-panel .d-flex {
      gap: 12px !important; /* Memperlebar jarak antar tombol */
    }

    /* --- PDF VIEWER STYLES --- */
    .bg-dark-viewer { background-color: #323639; }
    .pdf-toolbar { background: rgba(45, 45, 45, 0.85); backdrop-filter: blur(10px); border-bottom: 1px solid rgba(255, 255, 255, 0.1); padding: 12px 25px; display: flex; justify-content: space-between; align-items: center; z-index: 100; position: sticky; top: 0; }
    .pdf-viewport { flex-grow: 1; overflow-y: auto; padding: 40px 20px; scroll-behavior: smooth; }
    .pdf-canvas-stack canvas { background-color: white; margin: 0 auto 30px auto; display: block; box-shadow: 0 10px 30px rgba(0,0,0,0.5); border-radius: 4px; max-width: 100%; }
     .empty-preview-state {
      display: flex;
      flex-direction: column;
      justify-content: center;
      align-items: center;
      height: 100%;
      text-align: center;
      padding: 20px;
    }
     .icon-circle {
      width: 120px;
      height: 120px;
      background: rgba(255, 255, 255, 0.03);
      border-radius: 50%;
      display: flex;
      justify-content: center;
      align-items: center;
      position: relative;
      border: 2px dashed rgba(88, 80, 235, 0.3); /* Indigo Dashed */
      transition: all 0.5s ease;
      animation: softPulse 3s infinite ease-in-out;
    }
    .icon-circle i {
      font-size: 3.5rem;
      background: linear-gradient(135deg, #ff4d4d, #f9cb28); /* PDF Red-Orange Gradient */
      -webkit-background-clip: text;
      -webkit-text-fill-color: transparent;
      filter: drop-shadow(0 0 15px rgba(255, 77, 77, 0.4));
    }

    /* Dekorasi Lingkaran Cahaya di Belakang */
    .icon-circle::after {
      content: '';
      position: absolute;
      width: 140px;
      height: 140px;
      border-radius: 50%;
      background: radial-gradient(circle, rgba(88, 80, 235, 0.15) 0%, transparent 70%);
      z-index: -1;
    }

    /* Animasi Berdenyut Halus */
    @keyframes softPulse {
      0% { transform: scale(1); border-color: rgba(88, 80, 235, 0.3); }
      50% { transform: scale(1.05); border-color: rgba(16, 185, 129, 0.6); }
      100% { transform: scale(1); border-color: rgba(88, 80, 235, 0.3); }
    }
    .preview-title {
      color: #ffffff;
      font-weight: 700;
      letter-spacing: 1px;
      margin-top: 20px;
      text-shadow: 0 2px 10px rgba(0,0,0,0.3);
    }

    .preview-text {
      color: #94a3b8;
      max-width: 320px;
      line-height: 1.6;
    }

    .highlight-text {
      color: #818cf8; /* Light Indigo */
      font-weight: 600;
    }
    .dot-success { height: 10px; width: 10px; background-color: #28a745; border-radius: 50%; display: inline-block; margin-right: 8px; box-shadow: 0 0 10px #28a745; animation: pulse 2s infinite; }
    @keyframes pulse { 0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(40, 167, 69, 0.7); } 70% { transform: scale(1); box-shadow: 0 0 0 10px rgba(40, 167, 69, 0); } 100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(40, 167, 69, 0); } }
    .pdf-viewport::-webkit-scrollbar { width: 8px; }
    .pdf-viewport::-webkit-scrollbar-track { background: transparent; }
    .pdf-viewport::-webkit-scrollbar-thumb { background: rgba(255,255,255,0.2); border-radius: 10px; }
    .pdf-viewport::-webkit-scrollbar-thumb:hover { background: rgba(255,255,255,0.3); }
    .pdf-paper-effect { background-color: white; margin: 0 auto 40px auto; display: block; box-shadow: 0 15px 35px rgba(0,0,0,0.4), 0 5px 15px rgba(0,0,0,0.2); border-radius: 2px; transition: transform 0.3s ease; }
    .pdf-paper-effect:hover { transform: translateY(-5px); }
    #pdfViewerContainer { background-color: #323639; padding-top: 50px; scroll-behavior: smooth; }
    #pdfViewerContainer::-webkit-scrollbar { width: 10px; }
    #pdfViewerContainer::-webkit-scrollbar-track { background: #2b2b2b; }
    #pdfViewerContainer::-webkit-scrollbar-thumb { background: #555; border-radius: 5px; }
    #pdfViewerContainer::-webkit-scrollbar-thumb:hover { background: #777; }

    /* Memastikan tabel tidak melakukan wrap teks sama sekali */

    .dataTables_scroll {
      border-radius: 12px !important;
      border: 1px solid #e2e8f0 !important;
      overflow: hidden !important; /* Memotong sudut tajam tabel */
      background-color: white;
    }

    .dataTables_scrollBody {
      border-bottom-left-radius: 12px !important;
      border-bottom-right-radius: 12px !important;
    }
    

    /* Memberikan sedikit scrollbar yang cantik jika tabel terlalu lebar */
    .dataTables_scrollBody::-webkit-scrollbar {
      height: 8px;
    }
    .dataTables_scrollBody::-webkit-scrollbar-track {
      background: #f1f5f9;
    }
    .dataTables_scrollBody::-webkit-scrollbar-thumb {
      background: #cbd5e1;
      border-radius: 10px;
    }

    /* Style Kotak Pegawai Sesuai Gambar */
    /* Style Kotak Pegawai - Ikon Hijau */
    .badge-pegawai-modern {
      display: inline-flex;
      align-items: center;
      background-color: #ffffff;
      border: 1px solid #64748b; /* Warna outline box */
      border-radius: 6px;
      padding: 5px 12px;
      color: #334155; /* Warna teks nama */
      font-weight: 700;
      font-size: 10.5px;
      text-transform: uppercase;
      white-space: nowrap;
      box-shadow: 0 1px 2px rgba(0,0,0,0.05);
      margin-bottom: 2px;
    }

    .badge-pegawai-modern i {
      margin-right: 10px;
      color: #10b981; /* MENGUBAH IKON MENJADI HIJAU */
      font-size: 14px;
      border-right: 1px solid #e2e8f0; /* Garis pembatas vertikal */
      padding-right: 8px;
      display: flex;
      align-items: center;
    }

    .pegawai-stack {
      display: flex;
      flex-direction: column;
      align-items: flex-start;
      gap: 4px;
    }

    /* Paksa tabel tidak wrap text */
    

    /* Warna Tombol SP Baru (Indigo Modern) */
    .btn-sp-baru {
      background-color: #5850eb !important;
      color: white !important;
      border: none !important;
      font-weight: 600 !important;
      transition: all 0.3s ease;
    }

    .btn-sp-baru:hover {
      background-color: #4338ca !important;
      transform: translateY(-2px);
      box-shadow: 0 4px 12px rgba(88, 80, 235, 0.3);
    }

    /* Class khusus agar teks bisa melipat (wrap) */
    .wrap-text {
      white-space: normal !important; /* Izinkan teks turun ke bawah */
      word-break: break-word !important;
      min-width: 300px !important; /* Beri lebar minimum agar proporsional */
      line-height: 1.6 !important;
      color: #475569 !important;
      text-align: justify;
    }

    /* Pastikan kolom lain tetap satu baris, tapi kolom uraian diizinkan melipat */
    .table-modern td {
      vertical-align: middle !important;
    }

    /* Target khusus untuk kolom Maksud / Uraian */
    .kolom-uraian {
      white-space: normal !important; /* WAJIB: Mengizinkan teks turun ke bawah */
      word-break: break-word !important;
      min-width: 350px !important; /* Memberi ruang agar teks punya tempat untuk melipat */
      max-width: 450px !important;
      line-height: 1.5 !important;
      text-align: justify;
    }

    /* Kolom selain uraian dipaksa tetap satu baris agar rapi */
    .kolom-nowrap {
      white-space: nowrap !important;
    }

    /* Tetap paksa kolom lain agar tidak wrap (kecuali maksud/uraian) */
    .table-modern td:not(.wrap-text), 
    .table-modern th:not(.wrap-text) {
      white-space: nowrap !important;
    }

    /* --- FIXED COLUMN WRAPPING CSS --- */
    #tableRiwayat {
      table-layout: fixed !important;
      width: 100% !important;
    }

    /* Tabel lainnya (Pegawai, Kegiatan, Pejabat) biarkan auto agar kolom menyesuaikan isi */
    #tablePegawaiMaster, 
    #tableKegiatanMaster, 
    #tablePejabatMaster {
      table-layout: auto !important;
      width: 100% !important;
      min-width: 900px; /* KUNCI: Beri lebar minimal agar scroll muncul */
    }

    .table-responsive::-webkit-scrollbar {
      height: 8px;
    }
    .table-responsive::-webkit-scrollbar-track {
      background: #f1f5f9;
    }
    .table-responsive::-webkit-scrollbar-thumb {
      background: #cbd5e1;
      border-radius: 10px;
    }
    .table-responsive::-webkit-scrollbar-thumb:hover {
      background: #5850eb;
    }

    /* Gaya sel umum */
    #tableRiwayat td {
      vertical-align: middle !important;
      overflow: hidden; /* Mencegah teks keluar dari batas sel */
    }

    /* Container Khusus Uraian agar melipat dengan rapi */
    .uraian-wrapper {
      white-space: normal !important; /* Izinkan teks melipat */
      word-wrap: break-word !important;
      overflow-wrap: break-word !important;
      line-height: 1.5 !important;
      text-align: justify;
      display: block;
      width: 350px; /* Kunci lebar kolom uraian di sini */
      font-size: 0.85rem;
      color: #475569;
    }

    /* Kolom lainnya dipaksa satu baris agar rapi */
    .td-nowrap {
      white-space: nowrap !important;
      text-overflow: ellipsis;
      overflow: hidden;
    }

    /* Warna merah untuk hari Minggu dan Sabtu di kalender */
.flatpickr-day.tanggal-merah {
  color: #dc3545 !important; /* Warna merah */
  font-weight: bold;
}

/* Warna saat tanggal merah dipilih */
.flatpickr-day.tanggal-merah.selected {
  background: #5850eb !important;
  color: white !important;
  border-color: #5850eb !important;
}

/* --- TOMBOL UTAMA PREMIUM (MAIN ACTION) --- */
      .btn-main-action {
      background: linear-gradient(135deg, #6366f1 0%, #a855f7 100%) !important;
      color: white !important;
      border: none !important;
      padding: 14px 20px !important;
      border-radius: 14px !important;
      font-weight: 700 !important;
      font-size: 0.95rem !important;
      position: relative;
      overflow: hidden;
      box-shadow: 0 4px 10px rgba(99, 102, 241, 0.3) !important;
      
      /* KUNCI POSISI TENGAH */
      display: flex !important;
      justify-content: center !important; /* Menarik konten ke tengah horizontal */
      align-items: center !important;     /* Menarik konten ke tengah vertikal */
      text-align: center !important;
      gap: 10px; /* Jarak antara ikon dan teks */
      
      transition: all 0.3s ease !important;
    }

    /* Warna Abu-abu Modern (Slate Gray Gradient) */
    .btn-batal-grad {
      background: linear-gradient(135deg, #94a3b8 0%, #64748b 100%) !important;
      color: white !important;
      padding: 14px 20px !important; 
      border-radius: 14px !important;
      font-weight: 700 !important;
      border: none !important;
      align-items: center;
      justify-content: center;
      gap: 10px;
      transition: all 0.3s ease !important;
      box-shadow: 0 4px 15px rgba(0, 0, 0, 0.1) !important;
      
      /* KUNCI: Hapus !important di sini agar bisa dikendalikan JavaScript */
      display: none; 
    }

      .btn-batal-grad.active-mode {
      display: flex !important;
    }

     .btn-batal-grad.show-mode {
      display: flex !important;
    }

    /* Efek Hover untuk tombol abu-abu */
    .btn-batal-grad:hover {
      transform: translateY(-3px);
      box-shadow: 0 8px 20px rgba(0, 0, 0, 0.15) !important;
      filter: brightness(1.1);
    }

    /* Class untuk menampilkan tombol sebagai flex saat mode edit */
    .btn-batal-grad.show-flex {
      display: flex !important;
    }

    .btn-main-action {
      background: linear-gradient(135deg, #6366f1 0%, #a855f7 100%) !important;
      color: white !important;
    }

    /* Pastikan tombol Submit secara default adalah flex agar terlihat */
    #btnSubmitSP {
      display: flex !important;
    }

    /* Tombol Batal biarkan mengikuti kontrol d-none/flex dari JS */
    #btnBatalEdit {
      display: none; /* Default sembunyi */
    }
    
    /* Class tambahan saat mode edit aktif agar tombol Batal muncul sebagai flex */
    #btnBatalEdit.show-flex {
      display: flex !important;
    }

    /* Efek Hover Seragam */
    .btn-main-action:hover, .btn-batal-grad:hover {
      transform: translateY(-2px) translateZ(0);
      box-shadow: 0 6px 15px rgba(99, 102, 241, 0.4) !important;
      filter: brightness(1.1);
    }

    /* Efek Kilatan Cahaya (Shine Effect) */
    .btn-main-action::before {
      content: '';
      position: absolute;
      top: 0;
      left: -100%;
      width: 100%;
      height: 100%;
      background: linear-gradient(
        120deg,
        transparent,
        rgba(255, 255, 255, 0.3),
        transparent
      );
      transition: all 0.6s;
    }

    .btn-main-action:hover::before {
      left: 100%;
    }

    .btn-main-action:hover {
      transform: scale(1.02) translateY(-4px); /* Sedikit membesar dan naik */
      box-shadow: 0 15px 30px -5px rgba(99, 102, 241, 0.5) !important;
      filter: brightness(1.1);
    }

    .btn-main-action:active {
      transform: scale(0.98);
    }

    /* Penyesuaian Icon agar lebih dinamis */
     .btn-main-action i {
      margin: 0 !important;
      line-height: 0;
    }

    .btn-main-action:hover i {
      transform: rotate(15deg) scale(1.2);
    }

    /* --- MODERN PILL TABS --- */
    .modern-tabs-container {
      background-color: #f1f5f9;
      padding: 6px;
      border-radius: 14px;
      display: flex;
      gap: 5px;
      margin-bottom: 25px;
      border: 1px solid #e2e8f0;
    }

    .modern-tab-btn {
      flex: 1;
      border: none;
      background: transparent;
      padding: 12px 10px;
      font-weight: 700;
      font-size: 0.85rem;
      color: #64748b;
      border-radius: 10px;
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      display: flex;
      align-items: center;
      justify-content: center;
      gap: 8px;
      cursor: pointer;
    }

    .modern-tab-btn i {
      font-size: 1.1rem;
    }

    /* Tab Aktif dengan Gradasi Indigo */
    .modern-tab-btn.active {
      background: white;
      color: #5850eb;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.08);
      transform: scale(1.02);
    }

    /* Efek khusus untuk tab Excel saat aktif */
    .modern-tab-btn.active.excel-active {
      color: #10b981; /* Warna hijau untuk excel */
    }

    /* Style input form dalam modal agar lebih clean */
    .custom-swal-content .form-floating > .form-control {
      border: 1px solid #e2e8f0 !important;
      background-color: #ffffff;
    }

    /* --- CUSTOM CANCEL BUTTON STYLE --- */
    .btn-cancel-modern {
      background-color: #f1f5f9 !important;
      color: #64748b !important;
      border: 1px solid #e2e8f0 !important;
      padding: 12px 30px !important;
      border-radius: 14px !important; /* Menyamakan dengan tombol Simpan */
      font-weight: 700 !important;
      transition: all 0.3s ease !important;
      font-size: 0.95rem !important;
    }

    .btn-cancel-modern:hover {
      background-color: #e2e8f0 !important;
      color: #475569 !important;
      transform: translateY(-2px);
    }

    /* --- MODERN GRADIENT EXPORT BUTTONS --- */
    .btn-action-grad {
      border: none !important;
      color: white !important;
      font-weight: 700 !important;
      padding: 10px 20px !important;
      border-radius: 12px !important;
      display: flex;
      align-items: center;
      gap: 8px;
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
      box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1) !important;
      font-size: 0.85rem !important;
    }

    .btn-action-grad:hover {
      transform: translateY(-3px);
      box-shadow: 0 8px 18px rgba(0, 0, 0, 0.15) !important;
      filter: brightness(1.1);
    }

    .btn-action-grad:active {
      transform: translateY(-1px);
    }

    /* Tambah Pegawai (Indigo Blue) */
    .btn-add-pegawai {
      background: linear-gradient(135deg, #3b82f6 0%, #2563eb 100%) !important;
    }

    /* Excel (Emerald Green) */
    .btn-excel-grad {
      background: linear-gradient(135deg, #10b981 0%, #059669 100%) !important;
    }

    /* PDF (Crimson Red) */
    .btn-pdf-grad {
      background: linear-gradient(135deg, #f43f5e 0%, #e11d48 100%) !important;
    }

    /* Styling Icon di dalam tombol export */
    .btn-action-grad i {
      font-size: 1rem;
    }

    /* --- FORCE ALL TEXT TO ARIAL --- */
    html, body, table, th, td, input, select, textarea, button, div, span, .badge {
      font-family: Arial, Helvetica, sans-serif !important;
    }

    /* --- STYLE NOMOR SP BOLD & UNGU --- */
    .td-primary-bold {
      color: #5850eb !important; /* Warna Ungu Indigo */
      font-weight: 800 !important; /* Tebal maksimal */
      font-size: 14px !important;
      letter-spacing: 0.3px;
    }

    /* Mengatasi class 'font-monospace' yang sering digunakan pada NIP & Kode Rekening */
    .font-monospace, 
    .td-primary-bold, 
    .badge-pic, 
    .badge-pegawai-modern,
    .uraian-wrapper,
    .dataTables_info,
    .pagination .page-link {
      font-family: Arial, Helvetica, sans-serif !important;
      letter-spacing: normal !important; /* Menghilangkan jarak renggang khas font kode */
    }

    /* Memastikan Header Tabel juga Arial */
    .table-modern thead th {
      white-space: nowrap !important;
      vertical-align: bottom !important;
      background-color: #f8fafc !important;
    }

    /* Memastikan Persentase Loading juga Arial */
    .percentage {
      font-family: Arial, sans-serif !important;
    }

    /* --- MODERN CIRCLE CLOSE BUTTON --- */
    /* --- MODERN CIRCLE CLOSE BUTTON (COMPACT SIZE) --- */
    .btn-close-circle {
      width: 32px; /* Diperkecil dari 38px */
      height: 32px; /* Diperkecil dari 38px */
      background-color: #f1f5f9; /* Warna dasar abu muda */
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      border: none;
      color: #64748b;
      transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
      cursor: pointer;
      outline: none !important;
      padding: 0;
    }

    .btn-close-circle i {
      font-size: 0.9rem; /* Ukuran ikon disesuaikan */
      line-height: 0;
      transition: transform 0.4s ease;
    }

    /* Efek Animasi Hover (Sesuai Gambar: Menjadi Merah) */
    .btn-close-circle:hover {
      background-color: #ef4444; /* Warna merah solid */
      color: white;
      transform: scale(1.1);
      box-shadow: 0 4px 12px rgba(239, 68, 68, 0.3); /* Shadow merah halus */
    }

    /* Efek Ikon Berputar saat Hover */
    .btn-close-circle:hover i {
      transform: rotate(90deg);
    }

    .btn-close-circle:active {
      transform: scale(0.95);
    }

    /* --- CSS UNTUK SUDUT TABEL ROUNDED --- */
    
    /* 1. Atur container tabelnya */
    .table-responsive, 
    .dataTables_scroll {
      border-radius: 12px !important;
      border: 1px solid #e2e8f0 !important;
      overflow: hidden !important; /* KUNCI: Memotong sudut agar melengkung */
      background-color: white;
      margin-bottom: 10px;
    }

    /* 2. Atur tabel di dalamnya */
    .table-modern {
      border: none !important;
      border-collapse: separate !important;
      border-spacing: 0 !important;
    }

     .table-modern thead th,
    .dataTables_scrollHead {
      background-color: #f8fafc !important;
      border-bottom: 1px solid #e2e8f0 !important;
    }

    /* 3. Menghilangkan border ganda pada sel pertama dan terakhir agar tidak bertumpuk dengan rounded */
    .table-modern thead th:first-child {
      border-top-left-radius: 12px;
    }
    .table-modern thead th:last-child {
      border-top-right-radius: 12px;
    }
    
    .table-modern tbody tr:last-child td {
      border-bottom: none !important;
    }
    /* 4. Pastikan bagian bawah tabel juga mengikuti rounded jika tidak ada footer */
    .table-modern tbody tr:last-child td:first-child {
      border-bottom-left-radius: 12px;
    }
    .table-modern tbody tr:last-child td:last-child {
      border-bottom-right-radius: 12px;
    }

    /* Menghilangkan garis tepi bawaan bootstrap yang kaku */
    .table-bordered {
      border: none !important;
    }

    /* Menghapus garis default datatables yang mengganggu sudut rounded */
    table.dataTable.no-footer {
      border-bottom: none !important;
    }
    
    .dataTables_scrollHead {
      background-color: #f8fafc !important;
      border-bottom: 1px solid #e2e8f0 !important;
    }

    /* --- FULLSCREEN MODAL TWEAKS --- */
    .modal-fullscreen .modal-content {
      background-color: #f8fafc; /* Background luar tabel sedikit abu-abu */
    }

    .modal-fullscreen .modal-body {
      padding: 2rem !important; /* Beri ruang lega di sekitar tabel */
    }

    /* Pastikan header modal fullscreen tetap putih bersih */
    .modal-fullscreen .modal-header {
      background-color: white;
      border-bottom: 1px solid #e2e8f0 !important;
      padding: 1.5rem 2rem !important;
    }

    /* --- COMPACT FORM STYLE (Sesuai Gambar) --- */
    .compact-form-group {
      margin-bottom: 15px;
      text-align: left;
    }
    
    .compact-label {
      font-size: 12px;
      font-weight: 700;
      color: #374151; /* Warna teks gelap minimalis */
      margin-bottom: 5px;
      display: block;
      text-transform: none; /* Mengikuti gambar, tidak harus uppercase */
    }

    .compact-input {
      width: 100%;
      padding: 10px 14px;
      border-radius: 10px;
      border: 1px solid #d1d5db !important;
      background-color: #ffffff !important;
      font-size: 14px !important;
      color: #1f2937;
      transition: all 0.2s ease;
    }

    .compact-input:focus {
      border-color: #5850eb !important;
      box-shadow: 0 0 0 3px rgba(88, 80, 235, 0.1) !important;
      outline: none;
    }

    /* Penyesuaian TextArea agar tidak terlalu besar */
    textarea.compact-input {
      min-height: 80px;
      resize: vertical;
    }

    /* --- STICKY SIDEBAR HEADER --- */
    .sidebar {
      padding: 0 !important; /* Kita pindahkan padding ke dalam konten */
      display: flex;
      flex-direction: column;
    }

    body.swal2-shown {
      padding-right: 0px !important;
      overflow-y: auto !important; /* Tetap biarkan scrollbar ada agar layout tidak melompat */
    }

   .sidebar-header-fixed {
      position: sticky;
      top: 0;
      background: white;
      z-index: 100;
      /* Menambah padding luar agar kartu tidak menempel ke pinggir layar */
      padding: 30px 25px 15px 25px; 
      border-bottom: 1px solid #f1f5f9;
    }

    .sidebar-scrollable-content {
      padding: 20px;
      overflow-y: auto;
      flex-grow: 1;
    }

    /* Penyesuaian scrollbar sidebar agar cantik */
    .sidebar-scrollable-content::-webkit-scrollbar {
      width: 6px;
    }
    .sidebar-scrollable-content::-webkit-scrollbar-thumb {
      background: #e2e8f0;
      border-radius: 10px;
    }

    /* --- STYLE TOMBOL SAAT DISABLED --- */
    .btn-nav-disabled {
      opacity: 0.5 !important;
      pointer-events: none !important;
      filter: grayscale(1) !important;
      cursor: not-allowed !important;
      box-shadow: none !important;
      transform: none !important;
    }

    #pdfViewerContainer {
      background-color: #323639;
      padding-top: 20px !important; /* Perkecil padding atas */
      scroll-behavior: smooth;
      display: block; /* Pastikan bukan flex jika PDF sudah muncul */
    }

    /* Pastikan placeholder tidak mengambil ruang saat disembunyikan */
    #placeholderInit {
      display: flex !important; /* Pastikan terlihat di awal */
      flex-direction: column;
      justify-content: center;
      align-items: center;
      height: 100%;
      width: 100%;
    }

     #placeholderInit.d-none {
      display: none !important;
    }

    /* --- BOX GROUPING NAVIGASI --- */
     .nav-control-panel {
      background-color: #ffffff;
      border: 1px solid #e2e8f0;
      border-radius: 20px; /* Lebih bulat */
      /* Menambah padding dalam agar tombol tidak mepet ke dinding kartu */
      padding: 24px; 
      margin-top: 10px;
      margin-bottom: 25px;
      box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.05);
    }


     .nav-group-label {
      font-size: 11px;
      font-weight: 800;
      text-transform: uppercase;
      color: #94a3b8;
      /* Menambah jarak bawah label agar tidak mepet ke tombol */
      margin-bottom: 20px; 
      letter-spacing: 1.5px;
      display: flex;
      align-items: center;
      gap: 10px;
    }

    .nav-group-label::after {
      content: "";
      height: 1px;
      flex-grow: 1;
      background-color: #f1f5f9;
    }

    /* --- FIX OFFSET & LOCK SCREEN HEIGHT --- */
    html, body {
      height: 100% !important;
      margin: 0 !important;
      padding: 0 !important;
      overflow: hidden !important; /* Mencegah seluruh halaman bergeser ke bawah */
    }

    .container-fluid {
      height: 100vh !important;
      overflow: hidden !important;
    }

    .row {
      height: 100% !important;
      margin: 0 !important; /* Menghapus gap samping bawaan bootstrap */
    }

    /* Memastikan sidebar dan preview mengambil tinggi sisa layar dengan pas */
    .sidebar {
      height: 100vh !important;
      max-height: 100vh !important;
      display: flex;
      flex-direction: column;
    }

    /* Container kanan (PDF Viewer) */
    .bg-dark-viewer {
      height: 100vh !important;
      max-height: 100vh !important;
      overflow: hidden !important;
      display: flex;
      flex-direction: column;
    }

    #pdfViewerContainer {
      flex-grow: 1; /* Mengambil sisa tinggi layar */
      overflow-y: auto !important; /* Hanya ini yang boleh scroll */
      background-color: #323639;
    }

    /* --- PERBAIKAN SPACING FORM --- */
    
    /* Memberikan jarak antar kelompok input agar tidak berdempetan */
    .sidebar-scrollable-content .mb-3 {
      margin-top: 20px !important;
      margin-bottom: 20px !important;
    }

    /* Khusus untuk label Dasar Surat Perintah */
    .form-label {
      margin-bottom: 12px !important;
      display: flex;
      align-items: center;
    }

    /* Memberikan sedikit napas pada tombol Tambah Baris */
    .btn-xs {
      box-shadow: 0 2px 5px rgba(59, 130, 246, 0.2);
      transition: all 0.2s;
    }

    .btn-xs:hover {
      transform: translateY(-1px);
      filter: brightness(1.1);
    }
  </style>
</head>
<body>

<!-- LOADING OVERLAY MODERN PAPER PLANE -->
<!-- LOADING OVERLAY LOTTIE PAPER PLANE - UPDATED -->
<div id="loading">
  <div class="loader-content">
    <!-- Hanya satu wadah ini untuk pesawat melayang -->
    <div id="lottie-container" style="width: 180px; height: 180px; margin-bottom: 20px;"></div>
    
    <div id="loaderPercentage" class="percentage">0%</div>
    
    <div class="progress-container">
      <div class="progress-track">
        <div id="progressBarFill" class="progress-fill"></div>
      </div>
    </div>
    
    <div class="status-text">
      Sedang memuat data. mohon tunggu
      <div class="loading-dots"><span>.</span><span>.</span><span>.</span></div>
    </div>
  </div>
</div>

<div class="container-fluid">
  <div class="row">
    
    <!-- LEFT COLUMN: FORM INPUT -->
    <div class="col-md-5 col-lg-4 sidebar">
      
      <!-- BAGIAN TETAP (FIXED HEADER) -->
      <!-- BAGIAN TETAP (FIXED HEADER) -->
      <div class="sidebar-header-fixed">
        <h4 class="mb-4 text-primary fw-bold" style="font-family: Arial, Helvetica, sans-serif;">FORM SURAT PERINTAH</h4>
        
        <!-- BOX GROUPING / CONTROL PANEL -->
        <div class="nav-control-panel">
          <div class="nav-group-label">Menu Navigasi & Master Data</div>
          
          <!-- TOMBOL MASTER DATA (Row 1) -->
          <div class="d-flex gap-3 mb-3">
            <button id="navBtnPegawai" class="btn btn-sm btn-gradient btn-blue-grad w-100" data-bs-toggle="modal" data-bs-target="#modalPegawai">
              <i class="bi bi-person-badge"></i> Pegawai
            </button>
            <button id="navBtnKegiatan" class="btn btn-sm btn-gradient btn-blue-grad w-100" data-bs-toggle="modal" data-bs-target="#modalKegiatan">
              <i class="bi bi-briefcase"></i> Sub Kegiatan
            </button>
            <button id="navBtnPejabat" class="btn btn-sm btn-gradient btn-blue-grad w-100" data-bs-toggle="modal" data-bs-target="#modalPejabat">
              <i class="bi bi-award"></i> Pejabat
            </button>
          </div>

          <!-- TOMBOL RIWAYAT & RESET (Row 2) -->
          <div class="d-flex gap-3">
            <button id="navBtnRiwayat" class="btn btn-sm btn-gradient btn-green-grad w-100" data-bs-toggle="modal" data-bs-target="#modalRiwayat">
              <i class="bi bi-clock-history"></i> Riwayat SP
            </button>
            <button id="navBtnBaru" type="button" class="btn btn-sm btn-gradient btn-purple-grad w-100" onclick="resetFormSP()">
              <i class="bi bi-file-earmark-plus"></i> SP Baru
            </button>
          </div>
        </div>
      </div>

      <!-- BAGIAN YANG BISA DI-SCROLL (FORMULIR) -->
      <div class="sidebar-scrollable-content">
        <input type="hidden" id="idRiwayat" value="">

        <form id="spForm">
          <!-- NOMOR & TANGGAL SP -->
          <div class="row g-2 mb-2">
            <div class="col-md-6">
              <label class="form-label">Nomor SP</label>
              <input type="text" class="form-control" name="nomorSP" placeholder="Contoh: 800/..." required>
            </div>
            <div class="col-md-6">
              <label class="form-label">Tanggal Pembuatan SP</label>
              <input type="date" class="form-control" name="tanggalSP" required>
            </div>
          </div>

          <!-- TANGGAL BERANGKAT & KEMBALI -->
          <div class="row g-2 mb-3">
            <div class="col-md-6">
              <label class="form-label">Tanggal Berangkat</label>
              <input type="date" class="form-control" id="tglBerangkat" required>
            </div>
            <div class="col-md-6">
              <label class="form-label">Tanggal Kembali</label>
              <input type="date" class="form-control" id="tglKembali" required>
            </div>
          </div>

          <!-- DASAR HUKUM -->
          <!-- TAMBAHKAN mt-4 DI SINI -->
            <div class="mb-3 mt-4"> 
              <div class="form-label d-flex justify-content-between align-items-center">
                <span style="font-weight: 700; color: #475569;">Dasar Surat Perintah</span>
                <button type="button" class="btn btn-xs btn-primary py-1 px-2" onclick="addDasarHukum()" style="font-size: 10px; border-radius: 8px;">
                  <i class="bi bi-plus-lg"></i> Tambah Baris
                </button>
              </div>
              <div id="containerDasarHukum" class="mt-2"></div>
            </div>

          <!-- PEJABAT -->
          <div class="mb-3">
            <label class="form-label">Pejabat Penanda Tangan</label>
            <select class="form-select select2-pejabat" id="selectPejabat" style="width:100%" required>
              <option value="">Pilih Pejabat...</option>
            </select>
            <input type="hidden" name="pejabat" id="inputPejabatNama">
            <input type="hidden" name="pejabatNip" id="inputPejabatNip">
            <input type="hidden" name="pejabatJab" id="inputPejabatJab">
          </div>

          <!-- SUB KEGIATAN -->
          <div class="mb-3 p-3 bg-light rounded border">
            <label class="form-label">Sub Kegiatan</label>
            <select class="form-select select2-kegiatan" id="selectKegiatan" required style="width:100%">
              <option value="">Pilih Kegiatan...</option>
            </select>
            <input type="hidden" name="kodeKegiatan" id="kodeKegiatan">
            <input type="hidden" name="namaKegiatan" id="namaKegiatan">
          </div>

          <!-- PILIH PEGAWAI -->
          <div class="mb-3">
            <div class="form-label text-primary fw-bold mb-2">Pilih Pegawai Yang Diperintahkan</div>
            <div id="selectPegawaiVirtual" class="w-100"></div>
            <div id="containerPegawai" class="mt-3"></div>
          </div>

          <!-- MAKSUD / URAIAN -->
          <div class="mb-3">
            <label class="form-label d-flex justify-content-between">
              <span>Maksud/Uraian Perjalanan</span>
            </label>
            <div class="bg-white p-2 border rounded shadow-sm">
              <textarea class="form-control border-0 bg-light" id="inputUraian" rows="4" placeholder="Inputkan maksud/tujuan melakukan perjalanan dinas"></textarea>
            </div>
          </div>

          <!-- CONTAINER TOMBOL AKSI UTAMA -->
          <div class="mt-3 mb-5" id="actionButtonContainer">
            <button type="button" onclick="generateDocument()" class="btn btn-main-action w-100" id="btnSubmitSP">
              <i class="bi bi-file-earmark-check-fill" id="iconSubmit"></i> 
              <span id="textSubmit">Buat Surat Perintah</span>
            </button>

            <button type="button" onclick="cancelEditMode()" class="btn btn-batal-grad w-100 mt-2 d-none" id="btnBatalEdit">
              <i class="bi bi-x-circle"></i> <span>Batal Edit</span>
            </button>
          </div>
        </form>
      </div> <!-- End sidebar-scrollable-content -->

    </div>

    <!-- RIGHT COLUMN: PDF PREVIEW MODERN -->
    <div class="col-md-7 col-lg-8 d-flex flex-column p-0 bg-dark-viewer" style="height: 100vh;">
      <div id="actionArea" class="d-none pdf-toolbar animate__animated animate__fadeInDown">
        <div class="d-flex align-items-center"><div class="status-indicator me-3"><span class="dot-success"></span><small class="text-white fw-semibold">PDF SIAP DICETAK</small></div></div>
        <div class="d-flex gap-2">
          <button onclick="printDocument()" class="btn btn-sm btn-outline-light rounded-pill px-3"><i class="bi bi-printer me-1"></i> Print</button>
          <a id="btnDownload" href="#" download="Surat_Perintah.pdf" class="btn btn-sm btn-primary rounded-pill px-3 shadow-sm"><i class="bi bi-cloud-arrow-down-fill me-1"></i> Download File</a>
        </div>
      </div>
      <div id="pdfViewerContainer" class="pdf-viewport">
        <div id="placeholderInit" class="empty-preview-state">
  <div class="icon-circle mb-4">
    <!-- Menggunakan ikon PDF yang lebih detail -->
    <i class="bi bi-file-earmark-pdf-fill"></i>
  </div>
  <h5 class="preview-title">Pratinjau Surat Perintah</h5>
  <p class="preview-text small">
    Isi data di form kiri lalu klik tombol 
    <span class="highlight-text">Buat Surat Perintah</span> 
    untuk menampilkan dokumen secara instan di sini.
  </p>
</div>
        <div id="canvasWrapper" class="pdf-canvas-stack"></div>
      </div>
    </div>
  </div>
</div>

<!-- MODAL RIWAYAT SP -->
<div class="modal fade" id="modalRiwayat" tabindex="-1">
  <!-- GANTI modal-xl MENJADI modal-fullscreen -->
  <div class="modal-dialog modal-fullscreen"> 
    <div class="modal-content border-0">
      
      <!-- HEADER -->
      <div class="modal-header border-0 pb-0 mt-2 px-4 d-flex justify-content-between align-items-center">
        <div>
          <h5 class="modal-title fw-bold" style="font-size: 1.5rem;">Riwayat Surat Perintah</h5>
          <p class="text-muted mb-0">Manajemen data riwayat perjalanan dinas seluruh pegawai.</p>
        </div>
        <button type="button" class="btn-close-circle" data-bs-dismiss="modal">
          <i class="bi bi-x-lg"></i>
        </button>
      </div>

      <div class="modal-body px-4 pb-4">
        <!-- Area Pencarian & Tombol -->
        <div class="d-flex justify-content-between align-items-center mb-4 mt-2">
           <div class="custom-search-container">
            <input type="text" id="customSearchBoxRiwayat" class="custom-search-input" placeholder="Cari Riwayat...">
            <button class="custom-search-btn" type="button"><i class="bi bi-search"></i></button>
          </div>
        </div>

        <!-- TABEL -->
        <div class="table-responsive">
          <table id="tableRiwayat" class="table table-modern w-100 text-nowrap">
            <thead>
              <tr>
                <th class="text-center">NO</th>
                <th>NOMOR & TANGGAL SP</th>
                <th>BERANGKAT</th>
                <th>KEMBALI</th>
                <th>PEGAWAI</th>
                <th>MAKSUD / URAIAN</th>
                <th>KEGIATAN</th>
                <th>PEJABAT</th>
                <th class="text-center">AKSI</th>
              </tr>
            </thead>
            <tbody>
              <!-- Data dimuat melalui fungsi renderTableRiwayat di JavaScript -->
            </tbody>
          </table>
        </div>
      </div>

    </div>
  </div>
</div>

<!-- MODAL KELOLA PEGAWAI -->
<div class="modal fade" id="modalPegawai" tabindex="-1">
  <div class="modal-dialog modal-xl modal-dialog-centered">
    <div class="modal-content border-0 shadow-lg">
      
      <!-- MODAL HEADER DENGAN TOMBOL CLOSE BULAT -->
      <div class="modal-header border-0 pb-0 mt-3 px-4 d-flex justify-content-between align-items-center">
        <div>
          <h5 class="modal-title fw-bold">Data Pegawai</h5>
          <p class="text-muted small mb-0">Kelola daftar pegawai, NIP, pangkat, dan jabatan.</p>
        </div>
        
        <!-- TOMBOL CLOSE BULAT MODERN -->
        <button type="button" class="btn-close-circle" data-bs-dismiss="modal" title="Tutup">
          <i class="bi bi-x-lg"></i>
        </button>
      </div>

      <div class="modal-body px-4 pb-4">
        <div class="d-flex justify-content-between align-items-center mb-3 mt-3">
          
          <div class="d-flex gap-2">
            <!-- Tombol Tambah Pegawai (Blue Gradient) -->
            <button class="btn-action-grad btn-add-pegawai shadow-sm" onclick="showAddPegawaiForm()">
              <i class="bi bi-plus-lg"></i> Tambah Pegawai
            </button>
            
            <!-- Tombol Export Excel (Green Gradient) -->
            <button class="btn-action-grad btn-excel-grad shadow-sm" onclick="exportPegawaiExcel()" title="Export Excel">
              <i class="bi bi-file-earmark-excel"></i> Excel
            </button>
            
            <!-- Tombol Export PDF (Red Gradient) -->
            <button class="btn-action-grad btn-pdf-grad shadow-sm" onclick="exportPegawaiPDF()" title="Export PDF">
              <i class="bi bi-file-earmark-pdf"></i> PDF
            </button>
          </div>

          <!-- Search Bar Modern -->
          <div class="custom-search-container">
            <input type="text" id="customSearchBox" class="custom-search-input" placeholder="Cari Nama / NIP...">
            <button class="custom-search-btn" type="button"><i class="bi bi-search"></i></button>
          </div>
        </div>

        <!-- Tabel Pegawai Modern -->
        <div class="table-responsive">
          <table id="tablePegawaiMaster" class="table table-hover table-modern w-100" style="width:100%">
            <thead>
              <tr>
                <th>NAMA PEGAWAI</th>
                <th>NIP</th>
                <th>GOLONGAN</th>
                <th>JABATAN</th>
                <th class="text-center" style="width: 100px;">AKSI</th> <!-- Pastikan ini ada -->
              </tr>
            </thead>
            <tbody>
              <!-- Data dimuat via JavaScript -->
            </tbody>
          </table>
        </div>
      </div>

    </div>
  </div>
</div>

<!-- MODAL KELOLA KEGIATAN -->
<div class="modal fade" id="modalKegiatan" tabindex="-1">
  <div class="modal-dialog modal-xl modal-dialog-centered">
    <div class="modal-content border-0 shadow-lg">
      
      <!-- HEADER DENGAN TOMBOL CLOSE BULAT -->
      <div class="modal-header border-0 pb-0 mt-3 px-4 d-flex justify-content-between align-items-center">
        <div>
          <h5 class="modal-title fw-bold">Data Sub Kegiatan</h5>
          <p class="text-muted small mb-0">Kelola kode rekening dan nama sub kegiatan.</p>
        </div>
        
        <!-- TOMBOL CLOSE BULAT MODERN -->
        <button type="button" class="btn-close-circle" data-bs-dismiss="modal" title="Tutup">
          <i class="bi bi-x-lg"></i>
        </button>
      </div>

      <div class="modal-body px-4 pb-4">
        <div class="d-flex justify-content-between align-items-center mb-3 mt-3">
          <!-- Tombol Tambah (Blue Gradient) -->
          <button class="btn-action-grad btn-blue-grad shadow-sm" onclick="showAddKegiatanForm()">
            <i class="bi bi-plus-lg"></i> Tambah Kegiatan
          </button>

          <!-- Search Bar Modern -->
          <div class="custom-search-container">
            <input type="text" id="customSearchBoxKegiatan" class="custom-search-input" placeholder="Cari Kode / Nama...">
            <button class="custom-search-btn" type="button"><i class="bi bi-search"></i></button>
          </div>
        </div>

        <div class="table-responsive">
          <table id="tableKegiatanMaster" class="table table-hover table-modern w-100" style="width:100%">
            <thead>
              <tr>
                <th width="25%">KODE REKENING</th>
                <th>NAMA SUB KEGIATAN</th>
                <th class="text-center" width="15%">AKSI</th>
              </tr>
            </thead>
            <tbody>
              <!-- Data dimuat via JavaScript -->
            </tbody>
          </table>
        </div>
      </div>

    </div>
  </div>
</div>

<!-- MODAL KELOLA PEJABAT -->
<div class="modal fade" id="modalPejabat" tabindex="-1">
  <div class="modal-dialog modal-xl modal-dialog-centered">
    <div class="modal-content border-0 shadow-lg">
      
      <!-- MODAL HEADER DENGAN TOMBOL CLOSE BULAT -->
      <div class="modal-header border-0 pb-0 mt-3 px-4 d-flex justify-content-between align-items-center">
        <div>
          <h5 class="modal-title fw-bold">Data Pejabat</h5>
          <p class="text-muted small mb-0">Kelola Pejabat Penanda Tangan</p>
        </div>
        
        <!-- TOMBOL CLOSE BULAT MODERN -->
        <button type="button" class="btn-close-circle" data-bs-dismiss="modal" title="Tutup">
          <i class="bi bi-x-lg"></i>
        </button>
      </div>

      <div class="modal-body px-4 pb-4">
        <div class="d-flex justify-content-between align-items-center mb-3 mt-3">
          <!-- Tombol Tambah Pejabat (Blue Gradient) -->
          <button class="btn-action-grad btn-blue-grad shadow-sm" onclick="showAddPejabatForm()">
            <i class="bi bi-plus-lg"></i> Tambah Pejabat
          </button>

          <!-- Search Bar Modern -->
          <div class="custom-search-container">
            <input type="text" id="customSearchBoxPejabat" class="custom-search-input" placeholder="Cari Pejabat...">
            <button class="custom-search-btn" type="button"><i class="bi bi-search"></i></button>
          </div>
        </div>

        <div class="table-responsive">
          <table id="tablePejabatMaster" class="table table-hover table-modern w-100" style="width:100%">
            <thead>
              <tr>
                <th>NAMA</th>
                <th>PANGKAT</th>
                <th>JABATAN</th>
                <th class="text-center">AKSI</th>
              </tr>
            </thead>
            <tbody>
              <!-- Data dimuat via JavaScript -->
            </tbody>
          </table>
        </div>
      </div>

    </div>
  </div>
</div>

<!-- JS LIBRARIES -->
<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/sweetalert2@11"></script>
<script src="https://cdn.datatables.net/1.13.4/js/jquery.dataTables.min.js"></script>
<script src="https://cdn.datatables.net/1.13.4/js/dataTables.bootstrap5.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/flatpickr"></script>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/flatpickr/dist/plugins/confirmDate/confirmDate.css">
<script src="https://cdn.jsdelivr.net/npm/flatpickr/dist/plugins/confirmDate/confirmDate.js"></script>
<script src="https://npmcdn.com/flatpickr/dist/l10n/id.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.min.js"></script>

<script>
  // ==========================================
  // 1. SETUP & INISIALISASI
  // ==========================================
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.4.120/pdf.worker.min.js';

// Gunakan var untuk semua instance tabel agar tidak bentrok
  var dataTableInstance, dataTableKegiatanInstance, dataTablePejabatInstance, dataTableRiwayatInstance;
  var dbPegawai = [], dbKegiatan = [], dbPejabat = [], dbRiwayat = [];
  var selectedPegawaiOrder = [];
  var currentPdfBlobUrl = null;
  var loadingInterval;
  
  

$(document).ready(function() {
  initLottie();
    startSmartLoading();
    setupEventListeners();
    initUIComponents();

    // Panggil data dari server
    loadData();
    toggleEditModeUI(false); 
  $('#placeholderInit').show().removeClass('d-none');
  $('#actionArea').addClass('d-none');

$('.modal').on('shown.bs.modal', function () {
  $($.fn.dataTable.tables(true)).DataTable().columns.adjust();
    if ($.fn.DataTable.isDataTable('#tableRiwayat')) {
        $('#tableRiwayat').DataTable().columns.adjust().draw();
    }
    if ($.fn.DataTable.isDataTable('#tablePegawaiMaster')) {
        $('#tablePegawaiMaster').DataTable().columns.adjust().draw();
    }
    if ($.fn.DataTable.isDataTable('#tableKegiatanMaster')) {
        $('#tableKegiatanMaster').DataTable().columns.adjust().draw();
    }
    if ($.fn.DataTable.isDataTable('#tablePejabatMaster')) {
        $('#tablePejabatMaster').DataTable().columns.adjust().draw();
    }
}); 

// --- OTOMATIS CLEAR PENCARIAN SAAT MODAL DITUTUP ---
  $('.modal').on('hidden.bs.modal', function () {
    // 1. Bersihkan nilai teks di semua kotak pencarian
    $('#customSearchBox, #customSearchBoxKegiatan, #customSearchBoxPejabat, #customSearchBoxRiwayat').val('');

    // 2. Kembalikan tampilan tabel ke kondisi awal (tanpa filter)
    if (dataTableInstance) dataTableInstance.search('').draw();
    if (dataTableKegiatanInstance) dataTableKegiatanInstance.search('').draw();
    if (dataTablePejabatInstance) dataTablePejabatInstance.search('').draw();
    if (dataTableRiwayatInstance) dataTableRiwayatInstance.search('').draw();
  });
    // --- KONFIGURASI FLATPICKR ---
const configFlatpickr = {
  altInput: true,
  altFormat: "d F Y", // 'j' untuk tanggal tanpa nol, 'F' untuk nama bulan lengkap, 'Y' untuk tahun
  dateFormat: "Y-m-d",
  locale: "id", // Pastikan locale Indonesia aktif agar bulan menjadi 'Juni' bukan 'June'
  
  // Plugin ConfirmDate (Tombol Set Date)
  plugins: [
    new confirmDatePlugin({
      confirmText: "Set Date",
      showAlways: true,
      theme: "light"
    })
  ],
  
  onDayCreate: function(dObj, dStr, fp, dayElem) {
    if (dayElem.dateObj.getDay() === 0) { 
      dayElem.classList.add('tanggal-merah'); 
    }
  }
};

// Pasang ke semua input tanggal
flatpickr("input[name='tanggalSP']", configFlatpickr);
flatpickr("#tglBerangkat", configFlatpickr);
flatpickr("#tglKembali", configFlatpickr);

// Inisialisasi masing-masing kalender
    const fpSP = flatpickr("input[name='tanggalSP']", configFlatpickr);
    const fpBerangkat = flatpickr("#tglBerangkat", configFlatpickr);
    const fpKembali = flatpickr("#tglKembali", configFlatpickr);

    
    // KUNCI KALENDER: Fungsi Sinkronisasi Disable / Abu-abu Tanggal
    if (fpSP && fpBerangkat && fpKembali) {
        
        // --- TAMBAHAN BARU: FUNGSI DISABLE/ENABLE TANGGAL KEMBALI ---
        function setKembaliState(aktif) {
            // Disable input asli & input virtual dari Flatpickr
            fpKembali._input.disabled = !aktif;
            if (fpKembali.altInput) {
                fpKembali.altInput.disabled = !aktif;
            }
            // Jika dimatikan, kosongkan isinya
            if (!aktif) {
                fpKembali.clear();
            }
        }

        // 1. Matikan (disable) Tanggal Kembali saat halaman pertama kali dimuat
        setKembaliState(false);

        // 2. Jika SP diubah -> Berangkat tidak boleh kurang dari SP
        fpSP.config.onChange.push(function(selectedDates) {
            if (selectedDates.length > 0) {
                fpBerangkat.set('minDate', selectedDates[0]);
            } else {
                fpBerangkat.set('minDate', null);
            }
        });

        // 3. Jika Berangkat diubah -> SP maks adalah Berangkat, Kembali min adalah Berangkat
        fpBerangkat.config.onChange.push(function(selectedDates) {
            if (selectedDates.length > 0) {
                fpSP.set('maxDate', selectedDates[0]);
                fpKembali.set('minDate', selectedDates[0]);
                
                // ENABLE: Aktifkan Tanggal Kembali karena Berangkat sudah diisi
                setKembaliState(true);
            } else {
                fpSP.set('maxDate', null);
                fpKembali.set('minDate', null);
                
                // DISABLE: Matikan lagi Tanggal Kembali karena Berangkat dihapus/kosong
                setKembaliState(false);
            }
        });

        // 4. Jika Kembali diubah -> Berangkat maks adalah Kembali
        fpKembali.config.onChange.push(function(selectedDates) {
            if (selectedDates.length > 0) {
                fpBerangkat.set('maxDate', selectedDates[0]);
            } else {
                fpBerangkat.set('maxDate', null);
            }
        });
    }
    // --- AKHIR PERBAIKAN FLATPICKR ---

    $('.select2-kegiatan').select2({ theme: 'bootstrap-5' });
    $('.select2-pejabat').select2({ theme: 'bootstrap-5' });
  
  });

  function initLottie() {
  const container = document.getElementById('lottie-container');
  if (!container) return;
  container.innerHTML = ''; 
  lottie.loadAnimation({
    container: container,
    renderer: 'svg',
    loop: true,
    autoplay: true,
    path: 'https://assets4.lottiefiles.com/packages/lf20_x62chJ.json'
  });
}

function updateLoadingUI(val) {
  let displayVal = Math.floor(val);
  $('#progressBarFill').css('width', displayVal + '%');
  $('#loaderPercentage').text(displayVal + '%');
}

  function setupEventListeners() {
    $('#selectKegiatan').on('change', function() {
      const idx = $(this).val();
      if(idx !== "" && dbKegiatan[idx]) { $('#kodeKegiatan').val(dbKegiatan[idx][0]); $('#namaKegiatan').val(dbKegiatan[idx][1]); }
    });
    $('#selectPejabat').on('change', function() {
      const idx = $(this).val();
      if(idx !== "" && dbPejabat[idx]) {
        const p = dbPejabat[idx]; $('#inputPejabatNama').val(p[0]); $('#inputPejabatNip').val(p[1]); $('#inputPejabatJab').val(p[2]);
      } else { $('#inputPejabatNama').val(''); }
    });
    $('#spForm').on('submit', function(e) { e.preventDefault(); generateDocument(); });

    $('#modalRiwayat').on('shown.bs.modal', function () {
    if ($.fn.DataTable.isDataTable('#tableRiwayat')) {
        $('#tableRiwayat').DataTable().columns.adjust().draw();
    }
    });
  }

  // Helper Format Tanggal Indonesia
  // Helper Format Tanggal Indonesia (Tetap menyertakan angka 0 di depan tanggal)
  function formatTanggalIndo(dateStr) {
    if(!dateStr) return "";
    const bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    const parts = dateStr.split("-");
    if(parts.length !== 3) return dateStr;
    return `${parts[2].padStart(2, '0')} ${bulan[parseInt(parts[1]) - 1]} ${parts[0]}`;
  }

  // ==========================================
  // 2. EXPORT DATA PEGAWAI (EXCEL AUTO-FIT & PDF)
  // ==========================================
  function exportPegawaiExcel() {
  // 1. Validasi jika data kosong
  if (dbPegawai.length === 0) {
    Swal.fire({
      title: 'Info',
      text: 'Tidak ada data Pegawai yang bisa diekspor.',
      icon: 'info',
      confirmButtonColor: '#64748b'
    });
    return;
  }

  try {
    // 2. Persiapan Data (Menyesuaikan Header Template)
    const dataToExport = dbPegawai.map((row, index) => ({
      "No": index + 1,
      "Nama Lengkap": row[0] || "",
      "NIP": row[1] || "",
      "Golongan": row[2] || "",
      "Jabatan": row[3] || ""
    }));

    // 3. Membuat Worksheet
    const ws = XLSX.utils.json_to_sheet(dataToExport);

    // 4. Logika Auto-Fit Kolom (Agar Excel Rapi)
    const wscols = [
      { wch: 5 },  // No
      { wch: 35 }, // Nama Lengkap
      { wch: 25 }, // NIP
      { wch: 15 }, // Golongan
      { wch: 45 }  // Jabatan
    ];
    ws['!cols'] = wscols;

    // 5. Membuat Workbook
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Database Pegawai");

    // 6. Proses Download
    XLSX.writeFile(wb, "Data_Pegawai_Export.xlsx");

    // 7. POPUP SUKSES DENGAN JUMLAH DATA
    const totalData = dbPegawai.length;
    Swal.fire({
      title: 'Ekspor Berhasil!',
      html: `Sebanyak <b>${totalData}</b> data pegawai telah berhasil diekspor ke dalam file Excel.`,
      icon: 'success',
      confirmButtonColor: '#10b981', // Warna hijau emerald
      confirmButtonText: 'Selesai'
    });

  } catch (error) {
    console.error("Gagal ekspor excel:", error);
    Swal.fire("Gagal Ekspor", "Terjadi kesalahan saat mengolah data Excel.", "error");
  }
}

  function exportPegawaiPDF() {
  // 1. Validasi jika data kosong
  if (dbPegawai.length === 0) {
    Swal.fire({
      title: 'Data Kosong',
      text: 'Tidak ada data Pegawai untuk diekspor ke PDF.',
      icon: 'info',
      confirmButtonColor: '#64748b'
    });
    return;
  }

  try {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('p', 'pt', 'a4'); 
    const totalData = dbPegawai.length;

    // Judul Dokumen
    doc.setFont("arial", "bold");
    doc.setFontSize(16);
    doc.text("DATA PEGAWAI", 40, 45);
    
    // Sub Judul (Tanggal Cetak)
    doc.setFont("arial", "normal");
    doc.setFontSize(10);
    doc.text("Dicetak pada: " + new Date().toLocaleString('id-ID'), 40, 65);

    // Header Tabel
    const headers = [["No", "Nama Pegawai", "NIP", "Golongan", "Jabatan"]];
    
    // Isi Tabel
    const data = dbPegawai.map((row, index) => [ 
      index + 1, 
      row[0] || "", 
      row[1] || "", 
      row[2] || "", 
      row[3] || "" 
    ]);

    // Generate Tabel menggunakan AutoTable
    doc.autoTable({
      startY: 80,
      head: headers,
      body: data,
      theme: 'grid',
      styles: { fontSize: 9, font: 'helvetica', cellPadding: 5 },
      headStyles: { 
        fillColor: [225, 29, 72], // Warna Merah Crimson (Sesuai tema tombol PDF)
        textColor: [255, 255, 255],
        fontStyle: 'bold',
        halign: 'center'
      },
      columnStyles: {
        0: { cellWidth: 30, halign: 'center' },
        1: { cellWidth: 140 },
        2: { cellWidth: 100 },
        3: { cellWidth: 70, halign: 'center' },
        4: { cellWidth: 'auto' }
      }
    });

    // 2. PROSES SIMPAN / DOWNLOAD
    doc.save("Database_Pegawai_Export.pdf");

    // 3. POPUP SUKSES DENGAN JUMLAH DATA
    Swal.fire({
      title: 'PDF Berhasil Dibuat!',
      html: `Dokumen berisi <b>${totalData}</b> data pegawai telah berhasil diunduh.`,
      icon: 'success',
      confirmButtonColor: '#e11d48', // Warna merah sesuai tema PDF
      confirmButtonText: 'Selesai'
    });

  } catch (error) {
    console.error("Gagal ekspor PDF:", error);
    Swal.fire("Gagal Ekspor", "Terjadi kesalahan saat membuat file PDF: " + error.message, "error");
  }
}

  // ==========================================
  // 3. LOADING SCREEN & DATA FETCHING
  // ==========================================
  function startSmartLoading() {
  if (loadingInterval) clearInterval(loadingInterval);
  let progress = 0;
  $('#loading').show().css('opacity', '1');
  loadingInterval = setInterval(() => {
    if (progress < 40) progress += 8;
    else if (progress < 85) progress += 1;
    else if (progress < 95) progress += 0.2;
    updateLoadingUI(progress);
  }, 150);
}

  function finishSmartLoading() {
  if (loadingInterval) clearInterval(loadingInterval);
  updateLoadingUI(100);
  setTimeout(() => { $('#loading').fadeOut(600); }, 500);
}

  function loadData() {
  google.script.run.withSuccessHandler(function(data) {
    try {
      dbPegawai = data.pegawai || [];
      dbKegiatan = data.kegiatan || [];
      dbPejabat = data.pejabat || [];
      dbRiwayat = data.riwayat || [];

      renderTablePegawai(dbPegawai);
      renderTableKegiatan(dbKegiatan);
      renderTablePejabat(dbPejabat);
      renderTableRiwayat(dbRiwayat);
      
      initMultiSelectPegawai();
      populateDropdown('#selectKegiatan', dbKegiatan, (item) => `${item[0]} - ${item[1]}`);
      populateDropdown('#selectPejabat', dbPejabat, (item) => `${item[0]}`);
      
      loadDefaultDasarHukum();
      finishSmartLoading();
    } catch (e) {
      console.error(e);
      finishSmartLoading();
    }
  }).withFailureHandler(function(err) {
    clearInterval(loadingInterval);
    alert("Koneksi gagal: " + err);
  }).getDataMaster();
}

  function initUIComponents() {
  const cfg = { altInput: true, altFormat: "j F Y", dateFormat: "Y-m-d", locale: "id", plugins:[new confirmDatePlugin({confirmText: "Set Date", showAlways:true, theme:"light"})], onDayCreate: (d,s,f,el) => { if(el.dateObj.getDay()===0) el.classList.add('tanggal-merah'); } };
  flatpickr("input[name='tanggalSP']", cfg); flatpickr("#tglBerangkat", cfg); flatpickr("#tglKembali", cfg);
  $('.select2-kegiatan, .select2-pejabat').select2({ theme: 'bootstrap-5' });
}

  function populateDropdown(selector, data, textCallback) {
    const $sel = $(selector); $sel.empty().append('<option value="">Pilih...</option>');
    data.forEach((item, index) => { $sel.append(`<option value="${index}">${textCallback(item)}</option>`); });
  }

  // ==========================================
  // 4. RENDER TABEL
  // ==========================================
  function renderTableRiwayat(data) {
    if ($.fn.DataTable.isDataTable('#tableRiwayat')) $('#tableRiwayat').DataTable().destroy();
    const tbody = $('#tableRiwayat tbody'); tbody.empty(); 

    if (data && data.length > 0) {
      let html = ""; const reversedData = [...data].reverse();
      reversedData.forEach((row, revIndex) => {
        if(!row) return; 
        const index = data.length - 1 - revIndex; 
        const id = row[0] || ""; 
        const noSP = row[1] || ""; 
        const tgl = row[2] || ""; 
        const pjb = row[3] || ""; 
        const keg = row[4] || "";

        let namaPegawaiHtml = "-";
        let tglBerangkat = "-";
        let tglKembali = "-";
        let uraian = "-";

        try {
          const formData = JSON.parse(row[5]);
          
          // 1. STYLE PEGAWAI (Box, Ikon Hijau, Uppercase)
          if (formData.listPegawai && formData.listPegawai.length > 0) {
            namaPegawaiHtml = `<div class="pegawai-stack">` + 
              formData.listPegawai.map(p => 
                `<div class="badge-pegawai-modern">
                   <i class="bi bi-person"></i> ${p.nama}
                 </div>`
              ).join('') + `</div>`;
          }

          if (formData.tanggalBerangkat) tglBerangkat = formatTanggalIndo(formData.tanggalBerangkat);
          if (formData.tanggalKembali) tglKembali = formatTanggalIndo(formData.tanggalKembali);
          
          uraian = formData.uraianRaw || (formData.uraian ? formData.uraian.join(' ') : "-");
        } catch(e) { }

        // KONSTRUKSI 9 KOLOM (Tanpa Kolom Status)
        html += `<tr>
            <td class="text-center text-muted small">${revIndex + 1}</td> <!-- 0. NO -->
            
            <td>
              <!-- BAGIAN NOMOR SP BOLD & UNGU -->
              <div class="td-primary-bold">${noSP}</div>
              <div class="small text-muted mt-1" style="font-size: 11px;">
                <i class="bi bi-calendar-event me-1"></i>${tgl}
              </div>
            </td> 

            <td class="small td-nowrap"><i class="bi bi-box-arrow-up-right me-1 text-primary"></i>${tglBerangkat}</td> <!-- 2. BERANGKAT -->
            <td class="small td-nowrap"><i class="bi bi-box-arrow-in-down-left me-1 text-danger"></i>${tglKembali}</td> <!-- 3. KEMBALI -->
            
            <td>${namaPegawaiHtml}</td> <!-- 4. PEGAWAI (BOX STYLE) -->
            
            <td><div class="uraian-wrapper">${uraian}</div></td> <!-- 5. MAKSUD (WRAP TEXT) -->
            
            <td><div class="td-nowrap small text-secondary" style="width: 220px; overflow:hidden; text-overflow:ellipsis;">${keg}</div></td> <!-- 6. KEGIATAN -->
            <td class="td-nowrap small text-secondary">${pjb}</td> <!-- 7. PEJABAT -->
            
            <td class="text-center">
              <div class="d-flex justify-content-center gap-2">
                  <button class="btn-action-circle btn-edit" onclick="editRiwayat(${index})" title="Edit">
                      <i class="bi bi-pencil-fill"></i>
                  </button>
                  <button class="btn-action-circle btn-print" onclick="cetakRiwayat(${index})" title="Cetak">
                      <i class="bi bi-printer-fill"></i>
                  </button>
                  <button class="btn-action-circle btn-delete" onclick="hapusRiwayat('${id}')" title="Hapus">
                      <i class="bi bi-trash-fill"></i>
                  </button>
              </div>
          </td>
        </tr>`;
      });
      tbody.html(html);
    }

    setTimeout(() => {
      dataTableRiwayatInstance = $('#tableRiwayat').DataTable({ 
        pageLength: 5,
        lengthMenu: [5, 10, 25, 50, 100], 
        scrollX: true, 
        autoWidth: false,
        ordering: false, 
        // Layout Footer Modern (Background Light)
        dom: "<'row'<'col-sm-12'tr>><'row bg-light py-3 border-top mt-3 align-items-center'<'col-sm-4'l><'col-sm-4 text-center text-muted'i><'col-sm-4'p>>", 
        language: { 
          lengthMenu: "Tampilkan: _MENU_", 
          paginate: { 
            next: '<i class="bi bi-chevron-right"></i>', 
            previous: '<i class="bi bi-chevron-left"></i>' 
          }, 
          info: "Menampilkan _START_ - _END_ dari _TOTAL_ data",
          emptyTable: "Belum ada riwayat Surat Perintah."
        },
        columnDefs: [
          { targets: 0, width: "40px", className: "text-center" },
          { targets: [1, 2, 3, 7], width: "160px" },
          { targets: 4, width: "350px" },
          { targets: 5, width: "350px" }, // Lebar Kolom Maksud agar Wrap bekerja
          { targets: 6, width: "220px" }, 
          { targets: 8, width: "130px", className: "text-center" }
        ]
      });

      $('#customSearchBoxRiwayat').off('keyup').on('keyup', function() { 
        dataTableRiwayatInstance.search(this.value).draw(); 
      });
    }, 100);
}

  function renderTablePegawai(data) {
    if ($.fn.DataTable.isDataTable('#tablePegawaiMaster')) $('#tablePegawaiMaster').DataTable().destroy();
    const tbody = $('#tablePegawaiMaster tbody'); tbody.empty();
    
    if (data && data.length > 0) {
      let html = "";
      data.forEach((row, index) => { 
        if(!row) return; 
        html += `<tr>
            <td class="td-primary-bold" style="white-space: nowrap;">${row[0]||""}</td>
            <td class="text-muted small">${row[1]||""}</td>
            <td><span class="badge-pic-capsule">${row[2]||""}</span></td>
            <td class="text-secondary small">${row[3]||""}</td>
            <td class="text-center">
                <div class="d-flex justify-content-center gap-2">
                    <button class="btn-action-circle btn-edit" onclick="showEditPegawaiForm(${index})" title="Edit"><i class="bi bi-pencil-fill"></i></button>
                    <button class="btn-action-circle btn-delete" onclick="hapusPegawai(${index}, '${row[0]||""}')" title="Hapus"><i class="bi bi-trash-fill"></i></button>
                </div>
            </td>
        </tr>`; 
      });
      tbody.html(html);
    }

    setTimeout(() => {
      // Masukkan ke variabel global (tanpa 'let' atau 'var' lagi)
      dataTableInstance = $('#tablePegawaiMaster').DataTable({ 
        pageLength: 5,
        lengthMenu: [5, 10, 25],
        scrollX: true,
        autoWidth: false,
        dom: "<'row'<'col-sm-12'tr>><'row bg-light py-2 border-top mt-2 align-items-center'<'col-sm-4'l><'col-sm-4 text-center text-muted'i><'col-sm-4'p>>",
        language: { lengthMenu: "Tampilkan: _MENU_", info: "_TOTAL_ Pegawai", paginate: { next: '>', previous: '<' } },
        columnDefs: [{ targets: 0, width: "300px" }, { targets: 4, width: "100px" }]
      });

      // Hubungkan Pencarian
      $('#customSearchBox').off('input').on('input', function() {
        dataTableInstance.search(this.value).draw();
      });
    }, 150);
}

  function renderTableKegiatan(data) {
    if ($.fn.DataTable.isDataTable('#tableKegiatanMaster')) $('#tableKegiatanMaster').DataTable().destroy();
    const tbody = $('#tableKegiatanMaster tbody'); tbody.empty();
    
    if (data && data.length > 0) {
      let html = "";
      data.forEach((row, index) => { 
        if(!row) return; 
        html += `<tr>
            <td class="td-primary-bold">${row[0]}</td>
            <td class="text-dark small">${row[1]}</td>
            <td class="text-center">
                <div class="d-flex justify-content-center gap-2">
                    <button class="btn-action-circle btn-edit" onclick="showEditKegiatanForm(${index})"><i class="bi bi-pencil-fill"></i></button>
                    <button class="btn-action-circle btn-delete" onclick="hapusKegiatan(${index})" title="Hapus"><i class="bi bi-trash-fill"></i></button>
                </div>
            </td>
        </tr>`; 
      });
      tbody.html(html);
    }

    setTimeout(() => {
      dataTableKegiatanInstance = $('#tableKegiatanMaster').DataTable({ 
        pageLength: 5,
        lengthMenu: [5, 10, 25],
        scrollX: true,
        autoWidth: false,
        dom: "<'row'<'col-sm-12'tr>><'row bg-light py-2 border-top mt-2 align-items-center'<'col-sm-6'i><'col-sm-6'p>>",
        language: { lengthMenu: "Tampilkan: _MENU_", info: "_TOTAL_ Kegiatan" }
      });

      // Hubungkan Pencarian
      $('#customSearchBoxKegiatan').off('input').on('input', function() {
        dataTableKegiatanInstance.search(this.value).draw();
      });
    }, 150);
}

function renderTablePejabat(data) {
    if ($.fn.DataTable.isDataTable('#tablePejabatMaster')) $('#tablePejabatMaster').DataTable().destroy();
    const tbody = $('#tablePejabatMaster tbody'); tbody.empty();
    
    if (data && data.length > 0) {
      let html = "";
      data.forEach((row, index) => { 
        if(!row) return; 
        html += `<tr>
            <td class="td-primary-bold">${row[0]}</td>
            <td class="text-muted small">${row[1]}</td>
            <td class="text-secondary small">${row[2]}</td>
            <td class="text-center">
                <div class="d-flex justify-content-center gap-2">
                    <button class="btn-action-circle btn-edit" onclick="showEditPejabatForm(${index})"><i class="bi bi-pencil-fill"></i></button>
                    <button class="btn-action-circle btn-delete" onclick="hapusPejabat(${index}, '${row[0]}')" title="Hapus"><i class="bi bi-trash-fill"></i></button>
                </div>
            </td>
        </tr>`; 
      });
      tbody.html(html);
    }

    setTimeout(() => {
      dataTablePejabatInstance = $('#tablePejabatMaster').DataTable({ 
        pageLength: 5,
        lengthMenu: [5, 10, 25],
        scrollX: true,
        autoWidth: false,
        dom: "<'row'<'col-sm-12'tr>><'row bg-light py-2 border-top mt-2 align-items-center'<'col-sm-6'i><'col-sm-6'p>>",
        language: { lengthMenu: "Tampilkan: _MENU_", info: "_TOTAL_ Pejabat" }
      });

      // Hubungkan Pencarian
      $('#customSearchBoxPejabat').off('input').on('input', function() {
        dataTablePejabatInstance.search(this.value).draw();
      });
    }, 150);
}

  // ==========================================
  // 5. FUNGSI VIRTUAL SELECT 
  // ==========================================
  function initMultiSelectPegawai() {
    const options = dbPegawai.map((p, i) => { return { label: `${p[0]} (NIP. ${p[1]})`, value: i.toString() }; });
    VirtualSelect.init({ ele: '#selectPegawaiVirtual', options: options, multiple: true, search: true, placeholder: 'Pilih Pegawai...', searchPlaceholderText: 'Ketik nama atau NIP...', showSelectedOptionsFirst: true, hasOptionDescription: true, maxWidth: '100%', });
    document.querySelector('#selectPegawaiVirtual').addEventListener('change', function() { selectedPegawaiOrder = this.value; renderDetailPegawai(selectedPegawaiOrder); });
  }

  function renderDetailPegawai(indices) {
    const container = $('#containerPegawai'); container.empty(); if (!indices || indices.length === 0) return;
    indices.forEach(idx => { const p = dbPegawai[idx]; container.append(`<div class="row-item bg-white border p-3 mb-2 rounded shadow-sm border-start border-success border-4 animate__animated animate__fadeIn"><div class="d-flex justify-content-between align-items-center mb-2"><div class="fw-bold text-dark"><i class="bi bi-person-check-fill me-2 text-success"></i>${p[0]}</div><button type="button" class="btn btn-sm btn-outline-danger border-0 py-0" onclick="removePegawaiFromSelect(${idx})"><i class="bi bi-x-circle-fill"></i> Hapus</button></div><div class="row g-2"><div class="col-md-5"><label class="small text-muted mb-0" style="font-size: 0.7rem;">NIP</label><input class="form-control form-control-sm bg-light border-0" value="${p[1]}" readonly></div><div class="col-md-2"><label class="small text-muted mb-0" style="font-size: 0.7rem;">GOL</label><input class="form-control form-control-sm bg-light border-0 text-center" value="${p[2]}" readonly></div><div class="col-md-5"><label class="small text-muted mb-0" style="font-size: 0.7rem;">JABATAN</label><input class="form-control form-control-sm bg-light border-0" value="${p[3]}" readonly></div></div></div>`); });
  }

  function removePegawaiFromSelect(idx) {
    const vs = document.querySelector('#selectPegawaiVirtual'); let currentVals = vs.value;
    currentVals = currentVals.filter(v => v != idx.toString()); vs.setValue(currentVals);
    selectedPegawaiOrder = currentVals; renderDetailPegawai(selectedPegawaiOrder);
  }

  function addDasarHukum(isiText = "") {
    $('#containerDasarHukum').append(`<div class="row-item bg-white border p-2 mb-2 rounded shadow-sm"><textarea class="form-control form-control-sm mb-1 border-0 bg-light" name="dasarHukum[]" rows="3" placeholder="Ketik dasar surat...">${isiText}</textarea><div class="text-end"><span class="badge bg-danger cursor-pointer" style="cursor:pointer" onclick="$(this).closest('.row-item').remove()">Hapus</span></div></div>`);
  }

  // ==========================================
  // 6. RIWAYAT (EDIT, RESET, CETAK, HAPUS)
  // ==========================================
  function editRiwayat(index) {
    const row = dbRiwayat[index];
    if(!row) return;
    const formData = JSON.parse(row[5]);

    $('#idRiwayat').val(formData.idRiwayat);
    $('[name="nomorSP"]').val(formData.nomorSP);
    
    // Set date dengan flag "true" agar trigger rules onChange kalender
    const fpSP = document.querySelector("input[name='tanggalSP']")._flatpickr;
    if(fpSP && formData.tanggalSP) fpSP.setDate(formData.tanggalSP, true);

    const fpBerangkat = document.querySelector("#tglBerangkat")._flatpickr;
    if(fpBerangkat && formData.tanggalBerangkat) fpBerangkat.setDate(formData.tanggalBerangkat, true);

    const fpKembali = document.querySelector("#tglKembali")._flatpickr;
    if(fpKembali && formData.tanggalKembali) fpKembali.setDate(formData.tanggalKembali, true);

    let pIdx = ""; dbPejabat.forEach((p, i) => { if(p[1] == formData.pejabatData.nip && p[0] == formData.pejabat) pIdx = i; });
    $('#selectPejabat').val(pIdx).trigger('change');

    let kIdx = ""; dbKegiatan.forEach((k, i) => { if(k[0] == formData.kodeKegiatan) kIdx = i; });
    $('#selectKegiatan').val(kIdx).trigger('change');

    const textUraian = formData.uraianRaw ? formData.uraianRaw : (formData.uraian ? formData.uraian.join('\n') : "");
    $('#inputUraian').val(textUraian);

    $('#containerDasarHukum').empty();
    if(formData.dasarHukum && formData.dasarHukum.length > 0) {
      formData.dasarHukum.forEach(dh => addDasarHukum(dh));
    } else { addDasarHukum(); }

    let selectedPegIndices =[];
    if(formData.listPegawai && formData.listPegawai.length > 0) {
       formData.listPegawai.forEach(peg => {
          dbPegawai.forEach((p, i) => { if(p[1] == peg.nip) selectedPegIndices.push(i.toString()); });
       });
    }
    
    const vs = document.querySelector('#selectPegawaiVirtual');
    if(vs) vs.setValue(selectedPegIndices);
    selectedPegawaiOrder = selectedPegIndices;
    renderDetailPegawai(selectedPegawaiOrder);

    $('#modalRiwayat').modal('hide');
    Swal.mixin({toast: true, position: 'top-end', showConfirmButton: false, timer: 3000}).fire({ icon: 'info', title: 'Data dimuat ke formulir' });

    toggleEditModeUI(true); 
    window.scrollTo({ top: 0, behavior: 'smooth' });
  }

  // Fungsi untuk memuat Dasar Hukum Standar (Sesuai Gambar)
  function loadDefaultDasarHukum() {
    $('#containerDasarHukum').empty();
    const defaultDasar = [
      "Peraturan Menteri Keuangan Republik Indonesia Nomor 119 Tahun 2023 tentang Perubahan Atas Peraturan Menteri Keuangan Nomor 113/PMK.05/2012 tentang Perjalanan Dinas bagi Pejabat Negara, Pegawai Negeri Sipil, dan Pegawai Tidak Tetap.",
      "Peraturan Gubernur Jawa Barat Nomor 140 Tahun 2022 tentang Tugas Pokok, Fungsi, Rincian Tugas Unit dan Tata Kerja Badan Kepegawaian Daerah Pemerintah Provinsi Jawa Barat."
    ];
    defaultDasar.forEach(teks => addDasarHukum(teks));
  }

  function resetFormSP() {
    Swal.fire({
      title: 'Buat SP Baru?', 
      text: 'Formulir akan dikosongkan', 
      icon: 'question',
      showCancelButton: true, 
      confirmButtonColor: '#5850eb',
      confirmButtonText: 'Ya, Kosongkan', 
      cancelButtonText: 'Batal'
    }).then((r) => {
      if (r.isConfirmed) {
         $('#idRiwayat').val('');
         $('[name="nomorSP"]').val('');
         
         // Clear Kalender
         document.querySelector("input[name='tanggalSP']")._flatpickr.clear();
         document.querySelector("#tglBerangkat")._flatpickr.clear();
         document.querySelector("#tglKembali")._flatpickr.clear();
         
         // Reset Dropdown
         $('#selectPejabat').val('').trigger('change');
         $('#selectKegiatan').val('').trigger('change');
         $('#inputUraian').val('');
         
         // MUAT ULANG DASAR HUKUM STANDAR (Bukan Kosong)
         loadDefaultDasarHukum();
         
         // Reset Pegawai
         const vs = document.querySelector('#selectPegawaiVirtual');
         if(vs) vs.reset(); 
         $('#containerPegawai').empty();
         selectedPegawaiOrder = [];

         Swal.mixin({toast: true, position: 'top-end', showConfirmButton: false, timer: 2000}).fire({
           icon: 'success', title: 'Formulir telah direset'
         });
      }
    });
  }

  function hapusRiwayat(id) {
  if (!id) return;

  Swal.fire({
    title: 'Hapus Riwayat?',
    text: "Data yang dihapus tidak dapat dikembalikan!",
    icon: 'warning',
    showCancelButton: true,
    confirmButtonColor: '#ef4444',
    cancelButtonColor: '#64748b',
    confirmButtonText: 'Ya, Hapus!',
    cancelButtonText: 'Batal'
  }).then((result) => {
    if (result.isConfirmed) {
      startSmartLoading(); // Jalankan animasi pesawat
      
      google.script.run
        .withSuccessHandler(function() {
          // Muat ulang data riwayat secara diam-diam
          google.script.run.withSuccessHandler(function(newData) {
            dbRiwayat = newData || [];
            renderTableRiwayat(dbRiwayat); // Gambar ulang tabel
            finishSmartLoading();
            Swal.fire('Terhapus!', 'Data riwayat telah berhasil dihapus.', 'success');
          }).getRiwayatData();
        })
        .withFailureHandler(function(err) {
          finishSmartLoading();
          Swal.fire("Gagal Hapus", err.message, "error");
        })
        .hapusRiwayatById(id);
    }
  });
}

  function cetakRiwayat(index) {
  const row = dbRiwayat[index]; 
  if(!row) return; 
  
  const formData = JSON.parse(row[5]); // Ambil data JSON yang tersimpan
  
  $('#modalRiwayat').modal('hide'); // Tutup modal riwayat
  startSmartLoading(); // Jalankan animasi pesawat

  google.script.run
    .withSuccessHandler(function(response) {
      finishSmartLoading();
      renderPDFViewer(response.base64); // Tampilkan di pratinjau
      Swal.fire({ toast: true, position: 'top-end', icon: 'success', title: 'File Siap Dicetak', showConfirmButton: false, timer: 3000 });
    })
    .withFailureHandler(function(err) {
      finishSmartLoading();
      Swal.fire("Gagal Cetak", " cek kembali jaringan anda", "warning");
    })
    .generatePDF(formData);
}

  function reloadRiwayatSilently() { google.script.run.withSuccessHandler(function(data){ dbRiwayat = data ||[]; renderTableRiwayat(dbRiwayat); }).getRiwayatData(); }

  // ==========================================
  // 7. GENERATE PDF 
  // ==========================================
  function generateDocument() {
  const noSP = $('[name="nomorSP"]').val();
  const tglSP = $('input[name="tanggalSP"]').val();
  const tglBerangkat = $('#tglBerangkat').val();
  const tglKembali = $('#tglKembali').val();
  const pejabat = $('#inputPejabatNama').val();
  const pejabatJab = $('#inputPejabatJab').val();
  const keg = $('#kodeKegiatan').val();
  const namaKegiatan = $('#namaKegiatan').val();
  const rawUraian = $('#inputUraian').val() ? $('#inputUraian').val().trim() : "";

  if (!noSP || !tglSP || !pejabat) return Swal.fire('Data Belum Lengkap', 'Mohon isi Nomor SP, Tanggal Pembuatan, dan pilih Pejabat.', 'warning');
  if (!selectedPegawaiOrder || selectedPegawaiOrder.length === 0) return Swal.fire('Pegawai Belum Dipilih', 'Mohon pilih minimal satu pegawai.', 'warning');

  // Jalankan animasi pesawat
  startSmartLoading();

  // Logika Header
  let ttdHeader1 = "", ttdHeader2 = "", ttdHeader3 = "", tteInfo = "";
  const namaPejabatUpper = pejabat.toUpperCase();
  const jabatanUpper = pejabatJab.toUpperCase();

  if (namaPejabatUpper.includes("DEDI SUPANDI")) {
    ttdHeader1 = "KEPALA BADAN KEPEGAWAIAN DAERAH";
    ttdHeader2 = "PROVINSI JAWA BARAT,";
    ttdHeader3 = ""; 
    tteInfo = "KEPALA BADAN KEPEGAWAIAN DAERAH\nPROVINSI JAWA BARAT";
  } else {
    ttdHeader1 = "a.n. KEPALA BADAN KEPEGAWAIAN DAERAH";
    ttdHeader2 = "PROVINSI JAWA BARAT";
    ttdHeader3 = jabatanUpper + ","; 
    tteInfo = jabatanUpper + " BADAN KEPEGAWAIAN DAERAH\nPROVINSI JAWA BARAT";
  }

  let stringTanggal = (tglBerangkat === tglKembali) 
    ? ` pada tanggal ${formatTanggalIndo(tglBerangkat)}` 
    : ` pada tanggal ${formatTanggalIndo(tglBerangkat)} s/d ${formatTanggalIndo(tglKembali)}`;

  const formData = {
    idRiwayat: $('#idRiwayat').val(),
    nomorSP: noSP,
    tanggalSP: tglSP,
    tanggalBerangkat: tglBerangkat,
    tanggalKembali: tglKembali,
    ttdHeader1: ttdHeader1,
    ttdHeader2: ttdHeader2,
    ttdHeader3: ttdHeader3,
    tteInfo: tteInfo,
    pejabat: pejabat,
    pejabatData: { nip: $('#inputPejabatNip').val(), jab: pejabatJab },
    kodeKegiatan: keg,
    namaKegiatan: namaKegiatan,
    dasarHukum: Array.from(document.getElementsByName('dasarHukum[]')).map(el => el.value),
    uraian: [rawUraian + stringTanggal],
    uraianRaw: rawUraian,
    listPegawai: selectedPegawaiOrder.map(idx => ({ nama: dbPegawai[idx][0], nip: dbPegawai[idx][1], gol: dbPegawai[idx][2], jab: dbPegawai[idx][3] }))
  };

  google.script.run
  .withSuccessHandler(function(response) {
  finishSmartLoading(); 
  
  setTimeout(() => {
    if (response && response.base64) {
      // KUNCI: Panggil executeClearForm TERLEBIH DAHULU
      executeClearForm(); 
      
      // BARU KEMUDIAN tampilkan PDF hasil generate
      renderPDFViewer(response.base64); 
      
      reloadRiwayatSilently();
      
      Swal.fire({
        toast: true,
        position: 'top-end',
        icon: 'success',
        title: 'Berhasil Memperbarui Data',
        showConfirmButton: false,
        timer: 3000
      });
    }
  }, 300);
})
  .withFailureHandler(function(err) {
    finishSmartLoading();
    Swal.fire("Gagal", "Terjadi kesalahan di server: " + err.message, "error");
  })
  .generatePDF(formData);
}

  function renderPDFViewer(base64) {
  if (!base64) return;

  // 1. Sembunyikan Placeholder secara paksa (Gunakan .addClass d-none)
  $('#placeholderInit').addClass('d-none').hide(); 
  
  // 2. Tampilkan Toolbar Download & Print
  $('#actionArea').removeClass('d-none').addClass('d-flex animate__animated animate__fadeInDown');
  
  const wrapper = document.getElementById('canvasWrapper');
  wrapper.innerHTML = ''; // Bersihkan sisa pratinjau

  try {
    const binaryString = window.atob(base64);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }

    const loadingTask = pdfjsLib.getDocument({data: bytes});
    loadingTask.promise.then(function(pdf) {
      for (let i = 1; i <= pdf.numPages; i++) {
        pdf.getPage(i).then(function(page) {
          const viewport = page.getViewport({scale: 1.3});
          const canvas = document.createElement('canvas');
          canvas.className = 'pdf-paper-effect shadow mb-4';
          const context = canvas.getContext('2d');
          canvas.height = viewport.height;
          canvas.width = viewport.width;
          wrapper.appendChild(canvas);
          page.render({canvasContext: context, viewport: viewport});
        });
      }
    });

    const blob = new Blob([bytes], { type: 'application/pdf' });
    if (currentPdfBlobUrl) URL.revokeObjectURL(currentPdfBlobUrl);
    currentPdfBlobUrl = URL.createObjectURL(blob);
    $('#btnDownload').attr('href', currentPdfBlobUrl);

  } catch (err) {
    console.error("Gagal render PDF:", err);
  }
}

  function printDocument() {
  if (!currentPdfBlobUrl) {
    return Swal.fire('Info', 'Silakan buat dokumen terlebih dahulu sebelum mencetak.', 'info');
  }

  // Membuat iframe tersembunyi untuk proses cetak
  let printFrame = document.getElementById('printFrame');
  if (printFrame) {
    document.body.removeChild(printFrame);
  }

  printFrame = document.createElement('iframe');
  printFrame.id = 'printFrame';
  printFrame.style.position = 'fixed';
  printFrame.style.right = '0';
  printFrame.style.bottom = '0';
  printFrame.style.width = '0';
  printFrame.style.height = '0';
  printFrame.style.border = '0';
  printFrame.src = currentPdfBlobUrl;

  document.body.appendChild(printFrame);

  // Jalankan perintah cetak setelah iframe memuat PDF
  printFrame.onload = function() {
    setTimeout(function() {
      try {
        printFrame.contentWindow.focus();
        printFrame.contentWindow.print();
      } catch (e) {
        console.error("Gagal mencetak lewat iframe, membuka tab baru...");
        // Fallback: Jika gagal lewat iframe, buka di tab baru
        window.open(currentPdfBlobUrl, '_blank');
      }
    }, 500); // Beri jeda 500ms agar PDF benar-benar ter-render
  };
}

  // ==========================================
  // 8. REDESIGN ELEGANT MODALS & EXCEL UPLOAD
  // ==========================================

  // FUNGSI UNTUK MENGUNDUH TEMPLATE EXCEL
function downloadTemplateExcel() {
  try {
    // 1. Definisikan Header (Harus sesuai dengan nama kolom di logika Import Anda)
    const dataTemplate = [
      ["Nama Lengkap", "NIP", "Golongan", "Jabatan"], // Header
      ["Ahmad Subagja, S.T.", "198001012005011001", "IV/a", "Kepala Bidang Kepegawaian"], // Baris Contoh 1
      ["Siti Aminah, S.Pd.", "199202022015032005", "III/c", "Analis Sumber Daya Manusia"]  // Baris Contoh 2
    ];

    // 2. Buat Worksheet menggunakan library XLSX
    const ws = XLSX.utils.aoa_to_sheet(dataTemplate);

    // 3. Atur Lebar Kolom agar terlihat profesional (Auto-fit)
    const wscols = [
      { wch: 35 }, // Kolom Nama
      { wch: 25 }, // Kolom NIP
      { wch: 15 }, // Kolom Golongan
      { wch: 40 }  // Kolom Jabatan
    ];
    ws['!cols'] = wscols;

    // 4. Buat Workbook dan masukkan Worksheet tadi
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Template Pegawai");

    // 5. Proses Unduh
    XLSX.writeFile(wb, "Template_Upload_Pegawai.xlsx");

    // Feedback sukses
    Swal.mixin({
      toast: true,
      position: 'top-end',
      showConfirmButton: false,
      timer: 3000
    }).fire({
      icon: 'success',
      title: 'Template berhasil diunduh'
    });

  } catch (error) {
    console.error("Gagal mengunduh template:", error);
    Swal.fire("Gagal Unduh", "Terjadi kesalahan teknis: " + error.message, "error");
  }
}

  function switchSwalTab(id) {
  $('.modern-tab-btn').removeClass('active');
  $(`[onclick="switchSwalTab('${id}')"]`).addClass('active');
  $('.tab-content').removeClass('active');
  $('#'+id).addClass('active');
}


  function updateFileName(inp) { $('#file-name-display').text(inp.files[0] ? inp.files[0].name : "Klik/Seret Excel"); }

   // ==========================================
// FUNGSI KELOLA PEGAWAI (TAMBAH & EDIT)
// ==========================================
async function showAddPegawaiForm() {
  $('#modalPegawai').modal('hide');
  
  const { value: formResult } = await Swal.fire({
    customClass: { 
      popup: 'custom-swal-modal', 
      confirmButton: 'btn-main-action px-5 py-2', 
      cancelButton: 'btn-cancel-modern ms-2' 
    },
    title: '<div class="custom-swal-title"><i class="bi bi-person-plus-fill text-primary me-2"></i> Tambah Pegawai</div>',
    html: `
      <div class="custom-swal-content mt-3" style="font-family: Arial, sans-serif;">
        <div class="modern-tabs-container">
          <button class="modern-tab-btn active" onclick="switchSwalTab('tab-manual')"><i class="bi bi-keyboard"></i> Manual</button>
          <button class="modern-tab-btn" onclick="switchSwalTab('tab-excel')"><i class="bi bi-file-earmark-excel"></i> Excel</button>
        </div>
        <div id="tab-manual" class="tab-content active px-2">
          <div class="form-floating mb-3"><input type="text" class="form-control" id="add-nama" placeholder="Nama"><label>Nama Lengkap & Gelar</label></div>
          <div class="form-floating mb-3"><input type="text" class="form-control" id="add-nip" placeholder="NIP"><label>NIP</label></div>
          <div class="row g-3">
            <div class="col-4"><div class="form-floating"><input type="text" class="form-control" id="add-gol" placeholder="Gol"><label>Gol</label></div></div>
            <div class="col-8"><div class="form-floating"><input type="text" class="form-control" id="add-jab" placeholder="Jabatan"><label>Jabatan</label></div></div>
          </div>
        </div>
        <div id="tab-excel" class="tab-content px-2 text-center">
          <!-- Bungkus area upload agar terisolasi -->
          <div class="file-drop-wrapper">
            <div class="file-drop-area">
              <i class="bi bi-cloud-arrow-up" style="font-size: 3rem; color: #cbd5e0;"></i>
              <span class="file-msg" id="file-name-display">Seret file Excel ke sini</span>
              <span class="file-submsg">atau klik untuk memilih file (.xlsx)</span>
              <!-- Input file ini yang sering meluber jika tidak dibungkus wrapper -->
              <input class="file-input" type="file" id="upload-excel" accept=".xlsx, .xls" onchange="updateFileName(this)">
            </div>
          </div>

          <!-- Tombol Template diletakkan di luar wrapper dan diberi class khusus -->
          <div class="mt-3">
            <button type="button" class="btn btn-sm btn-template-excel rounded-pill px-4" 
                    onclick="event.stopPropagation(); downloadTemplateExcel();">
              <i class="bi bi-download me-1"></i> Download Template
            </button>
          </div>
      </div>`,
    showCancelButton: true, 
    confirmButtonText: 'Simpan Data', 
    cancelButtonText: 'Batal', 
    buttonsStyling: false,
    preConfirm: () => {
      const isExcel = document.getElementById('tab-excel').classList.contains('active');
      if (isExcel) {
        const file = document.getElementById('upload-excel').files[0];
        if (!file) return Swal.showValidationMessage('Silakan pilih file Excel terlebih dahulu!');
        return { type: 'excel', file: file };
      } else {
        // VALIDASI MANUAL KETAT
        const nama = $('#add-nama').val().trim();
        const nip = $('#add-nip').val().trim();
        const gol = $('#add-gol').val().trim();
        const jab = $('#add-jab').val().trim();

        if (!nama) return Swal.showValidationMessage('Nama Lengkap wajib diisi!');
        if (!nip) return Swal.showValidationMessage('NIP wajib diisi!');
        if (!gol) return Swal.showValidationMessage('Golongan wajib diisi!');
        if (!jab) return Swal.showValidationMessage('Jabatan wajib diisi!');

        return { type: 'manual', data: { nama, nip, gol, jab } };
      }
    }
  });

  if (formResult) {
    startSmartLoading(); // Munculkan Animasi Pesawat

    // --- HANDLER UNTUK NOTIFIKASI SUKSES ---
    const successCallback = (msg) => {
      loadData(); // Refresh data & render ulang tabel
      Swal.fire({
        title: 'Berhasil!',
        text: msg,
        icon: 'success',
        confirmButtonColor: '#5850eb',
        confirmButtonText: 'Lihat Data Pegawai'
      }).then(() => {
        // TAMPILKAN TABEL PEGAWAI SETELAH KLIK OK
        $('#modalPegawai').modal('show');
      });
    };

    if (formResult.type === 'manual') {
      google.script.run
        .withSuccessHandler(() => successCallback("Data pegawai berhasil ditambahkan secara manual."))
        .withFailureHandler((err) => { finishSmartLoading(); Swal.fire("Gagal", err.message, "error"); })
        .tambahPegawai(formResult.data);
    } else {
      const reader = new FileReader();
      reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(firstSheet);
        
        const parsed = json.map(r => ({
          nama: r["Nama Lengkap"] || "",
          nip: r["NIP"] || "",
          gol: r["Golongan"] || "",
          jab: r["Jabatan"] || ""
        }));

        // 1. HITUNG JUMLAH DATA YANG AKAN DI-UPLOAD
        const jumlahData = parsed.length;

        google.script.run
          .withSuccessHandler(function() {
            // 2. REFRESH DATA & MATIKAN LOADING
            loadData(); 
            finishSmartLoading();

            // 3. TAMPILKAN POPUP DENGAN JUMLAH DATA
            Swal.fire({
              title: 'Berhasil!',
              html: `Sebanyak <b>${jumlahData}</b> data pegawai berhasil diimpor dari file Excel.`,
              icon: 'success',
              confirmButtonColor: '#5850eb',
              confirmButtonText: 'Lihat Data Pegawai'
            }).then(() => {
              // Tampilkan kembali modal tabel
              $('#modalPegawai').modal('show');
            });
          })
          .withFailureHandler((err) => { 
            finishSmartLoading(); 
            Swal.fire("Gagal Import", err.message, "error"); 
            $('#modalPegawai').modal('show');
          })
          .tambahPegawaiBulk(parsed);
      };
      reader.readAsArrayBuffer(formResult.file);
    }
  } else { 
    // Jika user menekan batal di form tambah, kembalikan ke modal tabel
    $('#modalPegawai').modal('show'); 
  }
}

  async function showEditPegawaiForm(index) {
  const p = dbPegawai[index]; 
  
  // 1. Sembunyikan modal tabel agar form edit terlihat jelas
  $('#modalPegawai').modal('hide');

  const { value: f } = await Swal.fire({
    customClass: { 
      popup: 'custom-swal-modal', 
      confirmButton: 'btn-main-action px-5 py-2', 
      cancelButton: 'btn-cancel-modern ms-2' 
    },
    title: '<div class="custom-swal-title"><i class="bi bi-pencil-square text-warning me-2"></i> Edit Pegawai</div>',
    html: `
      <div class="custom-swal-content mt-4 px-2" style="font-family: Arial, sans-serif;">
        <div class="form-floating mb-3">
          <input id="ed-nama" class="form-control" value="${p[0]}">
          <label>Nama Lengkap & Gelar</label>
        </div>
        <div class="form-floating mb-3">
          <input id="ed-nip" class="form-control" value="${p[1]}">
          <label>NIP</label>
        </div>
        <div class="row g-3">
          <div class="col-4">
            <div class="form-floating">
              <input id="ed-gol" class="form-control" value="${p[2]}">
              <label>Gol</label>
            </div>
          </div>
          <div class="col-8">
            <div class="form-floating">
              <input id="ed-jab" class="form-control" value="${p[3]}">
              <label>Jabatan</label>
            </div>
          </div>
        </div>
      </div>`,
    showCancelButton: true, 
    confirmButtonText: 'Update Data', 
    cancelButtonText: 'Batal',
    buttonsStyling: false,
    preConfirm: () => {
      const nama = $('#ed-nama').val();
      if (!nama) return Swal.showValidationMessage('Nama tidak boleh kosong!');
      return { 
        nama: nama, 
        nip: $('#ed-nip').val(), 
        gol: $('#ed-gol').val(), 
        jab: $('#ed-jab').val() 
      };
    }
  });

  if (f) {
    startSmartLoading(); // Jalankan animasi pesawat

    google.script.run
      .withSuccessHandler(function() {
        // Refresh data di background
        loadData(); 

        // Tampilkan Popup Sukses
        Swal.fire({
          title: 'Berhasil!',
          html: 'Data pegawai bernama <b>' + f.nama.toUpperCase() + '</b> telah di update.',
          icon: 'success',
          confirmButtonColor: '#5850eb',
          confirmButtonText: 'Kembali ke Tabel'
        }).then(() => {
          // TAMPILKAN KEMBALI MODAL TABEL PEGAWAI
          $('#modalPegawai').modal('show');
        });
      })
      .withFailureHandler(function(err) {
        finishSmartLoading();
        Swal.fire("Gagal Update", err.message, "error");
        $('#modalPegawai').modal('show');
      })
      .editPegawai(f, index);

  } else { 
    // Jika user klik Batal, kembalikan ke modal tabel
    $('#modalPegawai').modal('show'); 
  }
}

  async function showAddKegiatanForm() {
  $('#modalKegiatan').modal('hide');
  const { value: f } = await Swal.fire({
    customClass: { popup: 'custom-swal-modal', confirmButton: 'btn-main-action w-100 py-3', cancelButton: 'btn-cancel-modern w-100 mt-2' },
    title: '<div class="text-start px-2">Tambah Sub Kegiatan</div>',
    html: `
      <div class="px-2 mt-3">
        <div class="compact-form-group">
          <label class="compact-label">Kode Rekening</label>
          <input id="add-k-kode" class="compact-input" placeholder="Contoh: 5.03.02.1.01.0010">
        </div>
        <div class="compact-form-group">
          <label class="compact-label">Nama Sub Kegiatan</label>
          <textarea id="add-k-nama" class="compact-input" placeholder="Tuliskan nama sub kegiatan secara lengkap"></textarea>
        </div>
      </div>`,
    showCancelButton: true, confirmButtonText: 'Simpan Kegiatan', cancelButtonText: 'Batal', buttonsStyling: false,
    preConfirm: () => {
      const k = $('#add-k-kode').val().trim(); // atau #ed-k-kode untuk form edit
      const n = $('#add-k-nama').val().trim(); // atau #ed-k-nama untuk form edit

      if (!k) return Swal.showValidationMessage('Kode Rekening wajib diisi!');
      if (!n) return Swal.showValidationMessage('Nama Sub Kegiatan wajib diisi!');

      return [k, n];
    }
  });
  if (f) {
    startSmartLoading();
    google.script.run.withSuccessHandler(() => {
      loadData();
      Swal.fire({ title: 'Berhasil!', html: 'Kegiatan <b>' + f[1].toUpperCase() + '</b> telah ditambahkan.', icon: 'success', confirmButtonColor: '#5850eb' })
      .then(() => $('#modalKegiatan').modal('show'));
    }).tambahKegiatan({ kode: f[0], nama: f[1] });
  } else $('#modalKegiatan').modal('show');
}

  async function showEditKegiatanForm(index) {
  const k = dbKegiatan[index]; 
  $('#modalKegiatan').modal('hide');
  const { value: f } = await Swal.fire({
    customClass: { 
      popup: 'custom-swal-modal', 
      confirmButton: 'btn-main-action px-5 py-2', 
      cancelButton: 'btn-cancel-modern ms-2' 
    },
    title: '<div class="custom-swal-title"><i class="bi bi-pencil-fill text-warning me-2"></i> Edit Sub Kegiatan</div>',
    html: `
      <div class="custom-swal-content mt-4" style="font-family: Arial, sans-serif;">
        <div class="form-floating mb-3">
          <input id="ed-k-kode" class="form-control" value="${k[0]}">
          <label>Kode Rekening</label>
        </div>
        <div class="form-floating">
          <textarea id="ed-k-nama" class="form-control" style="height:100px">${k[1]}</textarea>
          <label>Nama Sub Kegiatan</label>
        </div>
      </div>`,
    showCancelButton: true, 
    confirmButtonText: 'Update Data', 
    cancelButtonText: 'Batal',
    buttonsStyling: false,
    preConfirm: () => {
      const kode = $('#ed-k-kode').val();
      const nama = $('#ed-k-nama').val();
      if (!kode || !nama) return Swal.showValidationMessage('Semua kolom wajib diisi!');
      return [kode, nama];
    }
  });

  if (f) {
    startSmartLoading(); // Munculkan Pesawat
    google.script.run.withSuccessHandler(function() {
      loadData(); // Refresh Tabel
      Swal.fire({
        title: 'Berhasil diperbarui!',
        html: 'Data kegiatan <b>' + f[1].toUpperCase() + '</b> telah di-update.',
        icon: 'success',
        confirmButtonColor: '#5850eb',
        confirmButtonText: 'OK'
      }).then(() => {
        $('#modalKegiatan').modal('show'); // Buka kembali modal tabel
      });
    }).editKegiatan({ kode: f[0], nama: f[1] }, index);
  } else { 
    $('#modalKegiatan').modal('show'); 
  }
}

  async function showAddPejabatForm() {
  $('#modalPejabat').modal('hide');
  const { value: f } = await Swal.fire({
    customClass: { popup: 'custom-swal-modal', confirmButton: 'btn-main-action w-100 py-3', cancelButton: 'btn-cancel-modern w-100 mt-2' },
    title: '<div class="text-start px-2">Tambah Pejabat</div>',
    html: `
      <div class="px-2 mt-3">
        <div class="compact-form-group">
          <label class="compact-label">Nama Lengkap Pejabat</label>
          <input id="pj-nama" class="compact-input" placeholder="Masukkan nama dan gelar">
        </div>
        <div class="compact-form-group">
          <label class="compact-label">Pangkat</label>
          <input id="pj-nip" class="compact-input" placeholder="Contoh: Pembina Utama Madya">
        </div>
        <div class="compact-form-group">
          <label class="compact-label">Jabatan</label>
          <input id="pj-jab" class="compact-input" placeholder="Contoh: Sekretaris BKD">
        </div>
      </div>`,
    showCancelButton: true, confirmButtonText: 'Simpan Data', cancelButtonText: 'Batal', buttonsStyling: false,
    preConfirm: () => {
      const n = $('#pj-nama').val().trim();
      const nip = $('#pj-nip').val().trim();
      const j = $('#pj-jab').val().trim();

      // VALIDASI: Cek satu per satu
      if (!n) return Swal.showValidationMessage('Nama Lengkap wajib diisi!');
      if (!nip) return Swal.showValidationMessage('Pangkat / NIP wajib diisi!');
      if (!j) return Swal.showValidationMessage('Jabatan wajib diisi!');

      return [n, nip, j];
    }
  });
  if (f) {
    startSmartLoading();
    google.script.run.withSuccessHandler(() => {
      loadData();
      Swal.fire({ title: 'Berhasil!', html: 'Pejabat <b>' + f[0].toUpperCase() + '</b> telah disimpan.', icon: 'success', confirmButtonColor: '#5850eb' })
      .then(() => $('#modalPejabat').modal('show'));
    }).tambahPejabat({ nama: f[0], nip: f[1], jab: f[2] });
  } else $('#modalPejabat').modal('show');
}

  async function showEditPejabatForm(index) {
  const p = dbPejabat[index]; 
  
  // 1. Sembunyikan modal tabel agar fokus ke form edit
  $('#modalPejabat').modal('hide');

  const { value: f } = await Swal.fire({
    customClass: { 
      popup: 'custom-swal-modal', 
      confirmButton: 'btn-main-action w-100 py-3', 
      cancelButton: 'btn-cancel-modern w-100 mt-2' 
    },
    title: '<div class="text-start px-2">Edit Pejabat</div>',
    html: `
      <div class="px-2 mt-3" style="font-family: Arial, sans-serif;">
        <div class="compact-form-group">
          <label class="compact-label">Nama Lengkap Pejabat</label>
          <input id="pj-nama" class="compact-input" value="${p[0]}" placeholder="Masukkan nama dan gelar">
        </div>
        <div class="compact-form-group">
          <label class="compact-label">Pangkat / NIP</label>
          <input id="pj-nip" class="compact-input" value="${p[1]}" placeholder="Contoh: Pembina Utama Madya / 1980...">
        </div>
        <div class="compact-form-group">
          <label class="compact-label">Jabatan</label>
          <input id="pj-jab" class="compact-input" value="${p[2]}" placeholder="Contoh: Sekretaris BKD">
        </div>
      </div>`,
    showCancelButton: true, 
    confirmButtonText: 'Update Data', 
    cancelButtonText: 'Batal',
    buttonsStyling: false,
    preConfirm: () => {
      const nama = $('#pj-nama').val();
      const nip = $('#pj-nip').val();
      const jab = $('#pj-jab').val();
      if (!nama) return Swal.showValidationMessage('Nama pejabat wajib diisi!');
      return [nama, nip, jab];
    }
  });

  if (f) {
    startSmartLoading(); // Jalankan animasi pesawat

    google.script.run
      .withSuccessHandler(function() {
        // Refresh data di latar belakang
        loadData(); 

        // Tampilkan Popup Sukses
        Swal.fire({
          title: 'Berhasil diperbarui!',
          html: 'Data pejabat <b>' + f[0].toUpperCase() + '</b> telah di-update.',
          icon: 'success',
          confirmButtonColor: '#5850eb',
          confirmButtonText: 'OK'
        }).then(() => {
          // TAMPILKAN KEMBALI MODAL TABEL PEJABAT
          $('#modalPejabat').modal('show');
        });
      })
      .withFailureHandler(function(err) {
        finishSmartLoading();
        Swal.fire("Gagal Update", err.message, "error");
        $('#modalPejabat').modal('show');
      })
      .editPejabat({ nama: f[0], nip: f[1], jab: f[2] }, index);

  } else { 
    // Jika user klik Batal, kembalikan ke modal tabel
    $('#modalPejabat').modal('show'); 
  }
}

  function hapusPegawai(index, nama) {
  // 1. Tutup modal data pegawai agar konfirmasi terlihat jelas
  $('#modalPegawai').modal('hide'); 

  Swal.fire({
    title: 'Hapus Data Pegawai?',
    text: "Nama : " + nama,
    icon: 'warning',
    showCancelButton: true,
    confirmButtonColor: '#ef4444',
    cancelButtonColor: '#64748b',
    confirmButtonText: 'Ya, Hapus!',
    cancelButtonText: 'Batal'
  }).then((result) => {
    if (result.isConfirmed) {
      // 2. Jalankan animasi pesawat
      startSmartLoading(); 
      
      google.script.run
        .withSuccessHandler(function() {
          // 1. Matikan loading pesawat
          finishSmartLoading();

          // 2. Refresh data di latar belakang
          loadData(); 

          // 3. Tampilkan Popup Berhasil (Bukan Toast)
          Swal.fire({
            title: 'Berhasil!',
            html: 'Nama <b>' + nama.toUpperCase() + '</b> telah dihapus dari database.',
            icon: 'success',
            confirmButtonColor: '#5850eb',
            confirmButtonText: 'OK'
          }).then(() => {
            // 4. Tampilkan kembali modal data pegawai setelah klik OK
            $('#modalPegawai').modal('show');
          });
        })
        .withFailureHandler(function(err) {
          // Jika gagal, matikan loading dan beri peringatan
          finishSmartLoading();
          Swal.fire("Gagal Menghapus", err.message, "error");
          $('#modalPegawai').modal('show'); 
        })
        .hapusPegawaiByIndex(index);
    } else {
      // 6. Jika batal, buka kembali modal pegawai
      $('#modalPegawai').modal('show');
    }
  });
}

  function hapusKegiatan(index) {
  // 1. Sembunyikan modal tabel kegiatan
  $('#modalKegiatan').modal('hide');

  Swal.fire({
    title: 'Hapus Sub Kegiatan?',
    text: "Data yang dihapus akan terhapus secara permanen!",
    icon: 'warning',
    showCancelButton: true,
    confirmButtonColor: '#ef4444',
    cancelButtonColor: '#64748b',
    confirmButtonText: 'Ya, Hapus!',
    cancelButtonText: 'Batal',
    customClass: {
      popup: 'custom-swal-modal'
    }
  }).then((result) => {
    if (result.isConfirmed) {
      // 2. Jalankan animasi pesawat
      startSmartLoading();

      google.script.run
        .withSuccessHandler(function() {
          // 3. Muat ulang data terbaru
          loadData();

          // 4. Tampilkan Popup Berhasil
          Swal.fire({
            title: 'Berhasil!',
            text: 'Sub Kegiatan telah dihapus dari database.',
            icon: 'success',
            confirmButtonColor: '#5850eb',
            confirmButtonText: 'OK'
          }).then(() => {
            // 5. Tampilkan kembali modal tabel kegiatan setelah klik OK
            $('#modalKegiatan').modal('show');
          });
        })
        .withFailureHandler(function(err) {
          finishSmartLoading();
          Swal.fire("Gagal Menghapus", err.message, "error");
          $('#modalKegiatan').modal('show');
        })
        .hapusKegiatanByIndex(index);
    } else {
      // 6. Jika batal, buka kembali modal kegiatan
      $('#modalKegiatan').modal('show');
    }
  });
}

  function hapusPejabat(index, nama) {
  // 1. Sembunyikan modal tabel pejabat
  $('#modalPejabat').modal('hide');

  Swal.fire({
    title: 'Hapus Pejabat?',
    html: 'Apakah Anda yakin ingin menghapus <b>' + nama.toUpperCase() + '</b> dari database?',
    icon: 'warning',
    showCancelButton: true,
    confirmButtonColor: '#ef4444',
    cancelButtonColor: '#64748b',
    confirmButtonText: 'Ya, Hapus!',
    cancelButtonText: 'Batal',
    customClass: {
      popup: 'custom-swal-modal'
    }
  }).then((result) => {
    if (result.isConfirmed) {
      // 2. Jalankan animasi pesawat
      startSmartLoading();

      google.script.run
        .withSuccessHandler(function() {
          // 3. Refresh data di background
          loadData();

          // 4. Tampilkan Popup Berhasil (Bold & Uppercase)
          Swal.fire({
            title: 'Berhasil!',
            html: 'Pejabat <b>' + nama.toUpperCase() + '</b> telah dihapus.',
            icon: 'success',
            confirmButtonColor: '#5850eb',
            confirmButtonText: 'OK'
          }).then(() => {
            // 5. Tampilkan kembali modal tabel pejabat
            $('#modalPejabat').modal('show');
          });
        })
        .withFailureHandler(function(err) {
          finishSmartLoading();
          Swal.fire("Gagal Menghapus", err.message, "error");
          $('#modalPejabat').modal('show');
        })
        .hapusPejabatByIndex(index);
    } else {
      // 6. Jika batal, buka kembali modal pejabat
      $('#modalPejabat').modal('show');
    }
  });
}

  // Fungsi Mandiri untuk membersihkan form tanpa popup konfirmasi
function executeClearForm() {
  // --- 1. MEMBERSIHKAN FORM KIRI ---
  $('#idRiwayat').val('');
  $('[name="nomorSP"]').val('');
  $('#inputUraian').val('');

  // Bersihkan Kalender Flatpickr secara aman
  const ids = ["input[name='tanggalSP']", "#tglBerangkat", "#tglKembali"];
  ids.forEach(id => {
    const el = document.querySelector(id);
    if (el && el._flatpickr) el._flatpickr.clear();
  });

  // Reset Dropdown & Pegawai
  $('#selectPejabat, #selectKegiatan').val('').trigger('change');
  const vs = document.querySelector('#selectPegawaiVirtual');
  if (vs) vs.reset(); 
  $('#containerPegawai').empty();
  selectedPegawaiOrder = [];

  // Muat Dasar Hukum Standar
  loadDefaultDasarHukum();

  // Matikan mode edit & aktifkan kembali navigasi
  toggleEditModeUI(false);


  // --- 2. MEMBERSIHKAN PRATINJAU KANAN (PDF) ---
  
  // KUNCI PERBAIKAN: Paksa bersihkan canvas dan reset visibility
  $('#canvasWrapper').html(''); // Hapus semua elemen canvas di dalamnya
  
  // Sembunyikan toolbar download/print (Paksa d-none)
  $('#actionArea').addClass('d-none').removeClass('d-flex').hide();
  
  // Munculkan kembali Placeholder (Paksa hapus d-none dan tampilkan)
  $('#placeholderInit').removeClass('d-none').css('display', 'flex').show();

  // Hapus jejak file di memori browser
  if (currentPdfBlobUrl) {
    URL.revokeObjectURL(currentPdfBlobUrl);
    currentPdfBlobUrl = null;
  }
  
  console.log("Form dan Pratinjau telah dibersihkan.");
}

// Fungsi untuk mengganti tampilan tombol antara mode "Buat Baru" dan "Update"
function toggleEditModeUI(isEdit) {
  const btnSubmit = $('#btnSubmitSP');
  const textSubmit = $('#textSubmit');
  const btnBatal = $('#btnBatalEdit');
  const icon = $('#iconSubmit');

  const navButtons = $('#navBtnPegawai, #navBtnKegiatan, #navBtnPejabat, #navBtnRiwayat, #navBtnBaru');

  if (isEdit) {
    // --- MODE EDIT AKTIF ---
    textSubmit.text('Update Data');
    btnBatal.addClass('active-mode'); // Munculkan tombol batal
    icon.removeClass('bi-file-earmark-check-fill').addClass('bi-arrow-repeat');
    btnSubmit.css('background', 'linear-gradient(135deg, #f59e0b 0%, #d97706 100%)'); 
    navButtons.addClass('btn-nav-disabled').prop('disabled', true);
  } else {
    // --- MODE BUAT BARU ---
    textSubmit.text('Buat Surat Perintah');
    btnBatal.removeClass('active-mode'); // Sembunyikan tombol batal
    icon.removeClass('bi-arrow-repeat').addClass('bi-file-earmark-check-fill');
    btnSubmit.css('background', ''); 
    navButtons.removeClass('btn-nav-disabled').prop('disabled', false);
  }
}

function cancelEditMode() {
  Swal.fire({
    title: 'Batal Edit?',
    text: "Perubahan Data akan dibatalkan",
    icon: 'question',
    showCancelButton: true,
    confirmButtonText: 'Ya, Batalkan',
    cancelButtonText: 'Kembali'
  }).then((result) => {
    if (result.isConfirmed) {
      executeClearForm(); // Gunakan fungsi clear form yang sudah kita buat sebelumnya
      toggleEditModeUI(false); // Kembalikan tombol ke mode Buat Baru
    }
  });
}

</script>
</body>
</html>
