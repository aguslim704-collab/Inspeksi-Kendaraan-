<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Sistem Inspeksi Armada Dwibahasa / 双语车队检查系统</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- SheetJS Library for Excel Export -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
    <!-- Font Awesome Icons -->
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
    <!-- Google Fonts Inter -->
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <script>
        tailwind.config = {
            theme: {
                extend: {
                    fontFamily: {
                        sans: ['Inter', 'sans-serif'],
                    }
                }
            }
        }
    </script>
    <style>
        ::-webkit-scrollbar { width: 6px; height: 6px; }
        ::-webkit-scrollbar-track { background: #f1f5f9; }
        ::-webkit-scrollbar-thumb { background: #cbd5e1; border-radius: 4px; }
        ::-webkit-scrollbar-thumb:hover { background: #94a3b8; }
    </style>
</head>
<body class="bg-slate-50 text-slate-800 font-sans min-h-screen flex flex-col">

    <!-- Header & Navbar Navigation -->
    <header class="bg-white border-b border-slate-200 sticky top-0 z-30 shadow-sm">
        <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4">
            <div class="flex flex-col md:flex-row md:items-center md:justify-between gap-4">
                <div class="flex items-center gap-2">
                    <div class="p-2 bg-blue-600 text-white rounded-lg shadow-sm">
                        <i class="fa-solid fa-truck-front text-xl"></i>
                    </div>
                    <div>
                        <h1 class="text-xl sm:text-2xl font-bold text-slate-900 leading-tight">Sistem Inspeksi Armada & Sparepart</h1>
                        <p class="text-xs sm:text-sm text-slate-500 font-medium">车队及零配件检查系统 <span class="bg-blue-100 text-blue-700 text-[10px] px-2 py-0.5 rounded-full ml-1">Offline-First IndexedDB</span></p>
                    </div>
                </div>

                <!-- Global Action Controls (Export & Backup Buttons) -->
                <div class="flex flex-wrap items-center gap-2">
                    <button onclick="exportToExcel()" class="inline-flex items-center justify-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-2 rounded-lg font-semibold text-xs shadow-sm transition">
                        <i class="fa-regular fa-file-excel text-sm"></i>
                        <span>Excel</span>
                    </button>
                    <button onclick="exportToWordTable()" class="inline-flex items-center justify-center gap-1.5 bg-sky-600 hover:bg-sky-700 text-white px-3 py-2 rounded-lg font-semibold text-xs shadow-sm transition">
                        <i class="fa-regular fa-file-word text-sm"></i>
                        <span>Word</span>
                    </button>
                    <button onclick="backupDataJSON()" class="inline-flex items-center justify-center gap-1.5 bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-2 rounded-lg font-semibold text-xs shadow-sm transition" title="Cadangkan seluruh data ke file JSON">
                        <i class="fa-solid fa-download text-sm"></i>
                        <span>Backup JSON</span>
                    </button>
                    <button onclick="document.getElementById('importJsonInput').click()" class="inline-flex items-center justify-center gap-1.5 bg-amber-600 hover:bg-amber-700 text-white px-3 py-2 rounded-lg font-semibold text-xs shadow-sm transition" title="Pulihkan data dari file JSON">
                        <i class="fa-solid fa-upload text-sm"></i>
                        <span>Restore</span>
                    </button>
                    <input type="file" id="importJsonInput" accept=".json" class="hidden" onchange="restoreDataJSON(event)">
                </div>
            </div>
        </div>
    </header>

    <!-- Main Container -->
    <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6 flex-grow w-full">

        <!-- Top Controls Panel -->
        <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-4 sm:p-5 mb-6">
            <div class="grid grid-cols-1 md:grid-cols-12 gap-4">
                <div class="md:col-span-4">
                    <label class="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                        Pilih Minggu Inspeksi <span class="text-slate-400 font-normal lowercase">/ 选择检查周</span>
                    </label>
                    <select id="weekSelect" onchange="loadWeekData()" class="w-full bg-slate-50 border border-slate-300 text-slate-900 rounded-lg p-2.5 text-sm font-medium focus:ring-2 focus:ring-blue-500 transition">
                    </select>
                </div>

                <div class="md:col-span-5">
                    <label class="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                        Cari Unit <span class="text-slate-400 font-normal lowercase">/ 搜索车号</span>
                    </label>
                    <div class="relative">
                        <input type="text" id="searchInput" placeholder="Cari nomor unit (contoh: LAS-200)..." oninput="renderUnits()" class="w-full bg-slate-50 border border-slate-300 text-slate-900 rounded-lg p-2.5 pl-9 text-sm focus:ring-2 focus:ring-blue-500 transition">
                        <div class="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-slate-400">
                            <i class="fa-solid fa-magnifying-glass text-xs"></i>
                        </div>
                    </div>
                </div>

                <div class="md:col-span-3">
                    <label class="block text-xs font-semibold uppercase tracking-wider text-slate-600 mb-1">
                        Filter Status <span class="text-slate-400 font-normal lowercase">/ 状态筛选</span>
                    </label>
                    <select id="statusFilterSelect" onchange="renderUnits()" class="w-full bg-slate-50 border border-slate-300 text-slate-900 rounded-lg p-2.5 text-sm font-medium focus:ring-2 focus:ring-blue-500 transition">
                        <option value="ALL">Semua Status / 全部</option>
                        <option value="DONE">Selesai / 已检</option>
                        <option value="PENDING">Belum Diperiksa / 未检</option>
                        <option value="SCRAPPED">Afkir / 报废</option>
                    </select>
                </div>
            </div>
        </div>

        <!-- Weekly Target Component Banner -->
        <div class="bg-gradient-to-r from-sky-50 via-blue-50 to-indigo-50 border border-sky-200 rounded-xl p-4 mb-6 shadow-sm flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div class="flex items-start gap-3">
                <div class="p-2 bg-sky-500 text-white rounded-lg shadow-sm">
                    <i class="fa-solid fa-bullseye text-lg"></i>
                </div>
                <div>
                    <h2 class="text-xs font-bold uppercase tracking-wider text-sky-800">TARGET INSPEKSI MINGGU INI / 本周检查目标:</h2>
                    <p id="targetDisplay" class="text-sm sm:text-base font-semibold text-sky-950 mt-0.5">Memuat data target...</p>
                </div>
            </div>
            <button onclick="openManualEditModal()" class="inline-flex items-center justify-center gap-1.5 bg-white border border-sky-300 text-sky-700 hover:bg-sky-100 px-3 py-1.5 rounded-lg text-xs font-semibold shadow-xs transition shrink-0">
                <i class="fa-solid fa-pen-to-square"></i>
                <span>Edit Bagian / 手动编辑</span>
            </button>
        </div>

        <!-- Dashboard Summary Statistics Bar -->
        <div class="grid grid-cols-2 md:grid-cols-4 gap-3 sm:gap-4 mb-6">
            <div class="bg-white rounded-xl p-4 border border-slate-200 shadow-sm flex items-center justify-between">
                <div>
                    <p class="text-xs font-medium text-slate-500 uppercase">Total Unit / 总数</p>
                    <p class="text-2xl font-bold text-slate-900 mt-1" id="statTotal">0</p>
                </div>
                <div class="w-10 h-10 rounded-full bg-slate-100 flex items-center justify-center text-slate-600"><i class="fa-solid fa-cubes"></i></div>
            </div>
            <div class="bg-white rounded-xl p-4 border border-slate-200 shadow-sm flex items-center justify-between">
                <div>
                    <p class="text-xs font-medium text-slate-500 uppercase">Selesai / 已检查</p>
                    <p class="text-2xl font-bold text-emerald-600 mt-1" id="statDone">0</p>
                </div>
                <div class="w-10 h-10 rounded-full bg-emerald-50 flex items-center justify-center text-emerald-600"><i class="fa-solid fa-circle-check"></i></div>
            </div>
            <div class="bg-white rounded-xl p-4 border border-slate-200 shadow-sm flex items-center justify-between">
                <div>
                    <p class="text-xs font-medium text-slate-500 uppercase">Belum / 未检查</p>
                    <p class="text-2xl font-bold text-amber-600 mt-1" id="statPending">0</p>
                </div>
                <div class="w-10 h-10 rounded-full bg-amber-50 flex items-center justify-center text-amber-600"><i class="fa-solid fa-clock"></i></div>
            </div>
            <div class="bg-white rounded-xl p-4 border border-slate-200 shadow-sm flex items-center justify-between">
                <div>
                    <p class="text-xs font-medium text-slate-500 uppercase">Afkir / 报废</p>
                    <p class="text-2xl font-bold text-rose-600 mt-1" id="statScrapped">0</p>
                </div>
                <div class="w-10 h-10 rounded-full bg-rose-50 flex items-center justify-center text-rose-600"><i class="fa-solid fa-ban"></i></div>
            </div>
        </div>

        <!-- Fleet Grid Display -->
        <div class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4" id="unitContainer"></div>

        <!-- Custom Empty State Container -->
        <div id="emptyState" class="hidden text-center py-12 bg-white rounded-xl border border-slate-200 shadow-sm">
            <i class="fa-solid fa-truck-monster text-4xl text-slate-300 mb-3"></i>
            <h3 class="text-base font-semibold text-slate-700">Tidak ada unit ditemukan</h3>
            <p class="text-xs text-slate-500 mt-1">Coba ubah kata kunci pencarian atau filter status Anda.</p>
        </div>
    </main>

    <!-- Footer -->
    <footer class="bg-white border-t border-slate-200 mt-auto py-4">
        <div class="max-w-7xl mx-auto px-4 text-center text-xs text-slate-500">
            &copy; 2026 Fleet Inspection System Dual-Language. Data disimpan secara offline pada browser (IndexedDB).
        </div>
    </footer>

    <!-- MODAL: Input / Edit Inspeksi Armada -->
    <div id="inputModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex items-center justify-center p-4 hidden">
        <div class="bg-white rounded-2xl shadow-xl border border-slate-100 max-w-lg w-full max-h-[90vh] overflow-y-auto flex flex-col">
            <div class="px-6 py-4 border-b border-slate-100 flex items-center justify-between sticky top-0 bg-white z-10">
                <div class="flex items-center gap-2">
                    <div class="w-8 h-8 rounded-lg bg-blue-50 text-blue-600 flex items-center justify-center font-bold text-sm">
                        <i class="fa-solid fa-clipboard-check"></i>
                    </div>
                    <div>
                        <h3 id="modalUnitTitle" class="text-base font-bold text-slate-900">Input Inspeksi / 录入检查</h3>
                        <p class="text-xs text-slate-500">Formulir pemeriksaan kondisi armada</p>
                    </div>
                </div>
                <button type="button" onclick="closeModal()" class="text-slate-400 hover:text-slate-600 p-1.5 rounded-lg hover:bg-slate-100 transition">
                    <i class="fa-solid fa-xmark text-lg"></i>
                </button>
            </div>

            <form id="inspectionForm" onsubmit="saveInspection(event)" class="p-6 space-y-4">
                <input type="hidden" id="modalUnitId">

                <div>
                    <label class="block text-xs font-bold text-slate-700 uppercase mb-1">
                        Komponen Yang Diperiksa <span class="text-slate-400 font-normal lowercase">/ 检查部位</span>
                    </label>
                    <select id="modalItemSelect" required class="w-full bg-slate-50 border border-slate-300 text-slate-900 rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-blue-500 transition">
                    </select>
                </div>

                <div>
                    <label class="block text-xs font-bold text-slate-700 uppercase mb-1">
                        Status Kondisi <span class="text-slate-400 font-normal lowercase">/ 状态</span>
                    </label>
                    <select id="modalStatusSelect" required class="w-full bg-slate-50 border border-slate-300 text-slate-900 rounded-lg p-2.5 text-sm font-medium focus:ring-2 focus:ring-blue-500 transition">
                        <option value="Normal / 正常">Normal / 正常</option>
                        <option value="Perlu Perbaikan / 需要维修">Perlu Perbaikan / 需要维修</option>
                        <option value="Rusak / 损坏">Rusak / 损坏</option>
                    </select>
                </div>

                <div>
                    <label class="block text-xs font-bold text-slate-700 uppercase mb-1">
                        Unggah Foto Sparepart <span class="text-slate-400 font-normal lowercase">/ 上传照片</span>
                    </label>
                    <div class="mt-1 flex flex-col items-center justify-center border-2 border-dashed border-slate-300 rounded-lg p-4 bg-slate-50 hover:bg-slate-100 transition cursor-pointer relative" onclick="document.getElementById('modalPhotoInput').click()">
                        <i class="fa-solid fa-cloud-arrow-up text-2xl text-slate-400 mb-1"></i>
                        <span class="text-xs text-slate-600 font-medium">Klik untuk ambil/unggah foto</span>
                        <input type="file" id="modalPhotoInput" accept="image/*" capture="environment" onchange="previewPhoto(event)" class="hidden">
                    </div>
                    <div id="photoPreviewContainer" class="hidden relative mt-3">
                        <img id="modalPhotoPreview" src="" alt="Preview Foto" class="w-full h-44 object-cover rounded-lg border border-slate-200">
                        <button type="button" onclick="removePhoto()" class="absolute top-2 right-2 bg-rose-600 text-white p-1.5 rounded-full hover:bg-rose-700 shadow-md transition" title="Hapus Foto">
                            <i class="fa-solid fa-trash text-xs"></i>
                        </button>
                    </div>
                </div>

                <div>
                    <label class="block text-xs font-bold text-slate-700 uppercase mb-1">
                        Catatan / Perbaikan <span class="text-slate-400 font-normal lowercase">/ 备注整改</span>
                    </label>
                    <textarea id="modalNotes" rows="3" placeholder="Masukkan catatan kondisi sparepart atau tindakan perbaikan..." class="w-full bg-slate-50 border border-slate-300 text-slate-900 rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-blue-500 transition resize-none"></textarea>
                </div>

                <div class="pt-3 flex items-center justify-end gap-3 border-t border-slate-100">
                    <button type="button" onclick="closeModal()" class="px-4 py-2 border border-slate-300 text-slate-700 rounded-lg text-xs font-semibold hover:bg-slate-50 transition">Batal / 取消</button>
                    <button type="submit" class="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-semibold shadow-sm transition inline-flex items-center gap-1.5">
                        <i class="fa-solid fa-floppy-disk"></i>
                        <span>Simpan Data / 保存</span>
                    </button>
                </div>
            </form>
        </div>
    </div>

    <!-- MODAL: Edit Bagian Mingguan Manual -->
    <div id="manualModal" class="fixed inset-0 bg-slate-900/60 backdrop-blur-xs z-50 flex items-center justify-center p-4 hidden">
        <div class="bg-white rounded-2xl shadow-xl border border-slate-100 max-w-md w-full p-6 space-y-4">
            <div class="flex items-center justify-between border-b border-slate-100 pb-3">
                <div class="flex items-center gap-2">
                    <i class="fa-solid fa-pen-to-square text-sky-600 text-lg"></i>
                    <h3 class="text-base font-bold text-slate-900">Edit Target Inspeksi</h3>
                </div>
                <button type="button" onclick="closeManualModal()" class="text-slate-400 hover:text-slate-600 p-1"><i class="fa-solid fa-xmark text-lg"></i></button>
            </div>
            <div>
                <label class="block text-xs font-bold text-slate-700 mb-1">Komponen Wajib (Pisahkan dengan Koma):</label>
                <textarea id="manualItemsInput" rows="4" placeholder="Contoh: Rem Depan / 前刹车, Ban Belakang / 后轮胎" class="w-full bg-slate-50 border border-slate-300 text-slate-900 rounded-lg p-2.5 text-sm focus:ring-2 focus:ring-blue-500 transition resize-none"></textarea>
            </div>
            <div class="flex items-center justify-end gap-2 pt-2">
                <button type="button" onclick="closeManualModal()" class="px-3.5 py-2 border border-slate-300 text-slate-700 rounded-lg text-xs font-semibold hover:bg-slate-50">Batal</button>
                <button type="button" onclick="saveManualItems()" class="px-4 py-2 bg-sky-600 hover:bg-sky-700 text-white rounded-lg text-xs font-semibold shadow-sm">Simpan Perubahan / 保存</button>
            </div>
        </div>
    </div>

    <!-- JavaScript Logic Section -->
    <script>
        const masterUnits = [
            { id: "LAS-200-01", type: "200T", isScrapped: false },
            { id: "LAS-200-02", type: "200T", isScrapped: false },
            { id: "LAS-200-03", type: "200T", isScrapped: false },
            { id: "LAS-200-04", type: "200T", isScrapped: false },
            { id: "LAS-200-05", type: "200T", isScrapped: false },
            { id: "LAS-200-06", type: "200T", isScrapped: false },
            { id: "LAS-200-07", type: "200T", isScrapped: false },
            { id: "LAS-200-08", type: "200T", isScrapped: false },
            { id: "LAS-200-09", type: "200T", isScrapped: false },
            { id: "LAS-200-10", type: "200T", isScrapped: false },
            { id: "LAS-260-11", type: "260T", isScrapped: false },
            { id: "LAS-260-12", type: "260T", isScrapped: false },
            { id: "LAS-260-13", type: "260T", isScrapped: false },
            { id: "LAS-260-14", type: "260T", isScrapped: false },
            { id: "LAS-260-15", type: "260T", isScrapped: false },
            { id: "LAS-260-16", type: "260T", isScrapped: false },
            { id: "LAS-260-17", type: "260T", isScrapped: false },
            { id: "LAS-260-18", type: "260T", isScrapped: false },
            { id: "LAS-260-19", type: "260T", isScrapped: false },
            { id: "LAS-260-20", type: "260T", isScrapped: false },
            { id: "LAS-260-21", type: "260T", isScrapped: false },
            { id: "LAS-260-22", type: "260T", isScrapped: false },
            { id: "LAS-260-23", type: "260T", isScrapped: false },
            { id: "LAS-260-24", type: "260T", isScrapped: false },
            { id: "LAS-260-25", type: "260T", isScrapped: false },
            { id: "LAS-260-26", type: "260T", isScrapped: false },
            { id: "LAS-260-27", type: "260T", isScrapped: false },
            { id: "LAS-260-28", type: "260T", isScrapped: false },
            { id: "LAS-120-29", type: "110T", isScrapped: true },
            { id: "LAS-120-30", type: "110T", isScrapped: false },
            { id: "LAS-120-31", type: "110T", isScrapped: false },
            { id: "LAS-120-32", type: "110T", isScrapped: false },
            { id: "LAS-120-33", type: "110T", isScrapped: false },
            { id: "LAS-120-34", type: "120T", isScrapped: false },
            { id: "LAS-120-35", type: "120T", isScrapped: false },
            { id: "LAS-120-36", type: "120T", isScrapped: false },
            { id: "LAS-120-37", type: "120T", isScrapped: false },
            { id: "LAS-120-38", type: "120T", isScrapped: false },
            { id: "LAS-120-39", type: "120T", isScrapped: false },
            { id: "LAS-120-40", type: "120T", isScrapped: false },
            { id: "LAS-240-41", type: "240T", isScrapped: false },
            { id: "LAS-240-42", type: "240T", isScrapped: false },
            { id: "LAS-240-43", type: "240T", isScrapped: false },
            { id: "LAS-240-44", type: "240T", isScrapped: false }
        ];

        const defaultSchedules = {
            41: ["前后拖车钩 - Kait Penderek Depan & Belakang", "驾驶室隔热板 - Penyekat Panas Kabin"],
            42: ["警示灯 - Lampu Peringatan", "发动机防护罩 - Pelindung Mesin"],
            43: ["发动机增压软管 - Selang Turbo Mesin", "发动机皮带 - Tali Kipas/V-Belt Mesin"],
            44: ["止档器 & 托架 - Ganjal Roda & Bracket", "车尾警示牌 - Papan Peringatan Belakang"],
            45: ["电箱总成 - Rakitan Kotak Listrik", "监控系统 - Sistem Kamera Monitor"],
            46: ["车辆轮胎 - Ban Kendaraan", "减速机马达 - Motor Reduser"],
            47: ["刹车分泵 - Silinder Rem", "应急转向 - Kemudi Darurat"],
            48: ["止档器 & 警示牌 - Ganjal Roda & Papan Peringatan"]
        };

        let db;
        let currentWeek = 41;
        let inspectionRecords = {};
        let tempPhotoBase64 = "";

        function initDB() {
            const req = indexedDB.open("InspectionFleetDB", 1);
            req.onupgradeneeded = (e) => {
                db = e.target.result;
                if (!db.objectStoreNames.contains("records")) {
                    db.createObjectStore("records", { keyPath: "id" });
                }
            };
            req.onsuccess = (e) => {
                db = e.target.result;
                populateWeekSelect();
                loadWeekData();
            };
        }

        function populateWeekSelect() {
            const select = document.getElementById("weekSelect");
            select.innerHTML = "";
            for (let w = 1; w <= 48; w++) {
                let m = Math.ceil(w / 4);
                let wInM = ((w - 1) % 4) + 1;
                let opt = document.createElement("option");
                opt.value = w;
                opt.text = `Bulan ${m} Minggu ${wInM} (W${w}) / ${m}月第${wInM}周`;
                if (w === 41) opt.selected = true;
                select.appendChild(opt);
            }
        }

        function getTargetItems(week) {
            let savedManual = localStorage.getItem(`manual_items_w${week}`);
            if (savedManual && savedManual.trim() !== "") {
                return savedManual.split(",").map(i => i.trim());
            }
            return defaultSchedules[week] || ["Bagian Manual / 手动编辑项"];
        }

        function loadWeekData() {
            currentWeek = parseInt(document.getElementById("weekSelect").value);
            const items = getTargetItems(currentWeek);
            document.getElementById("targetDisplay").innerText = items.join("  •  ");
            if (!db) return;

            const tx = db.transaction("records", "readonly");
            const store = tx.objectStore("records");
            const req = store.getAll();

            req.onsuccess = () => {
                inspectionRecords = {};
                req.result.forEach(r => {
                    if (r.week === currentWeek) {
                        inspectionRecords[r.unitId] = r;
                    }
                });
                renderUnits();
            };
        }

        function renderUnits() {
            const container = document.getElementById("unitContainer");
            const emptyState = document.getElementById("emptyState");
            const search = document.getElementById("searchInput").value.toLowerCase().trim();
            const statusFilter = document.getElementById("statusFilterSelect").value;
            container.innerHTML = "";

            let doneCount = 0;
            let totalCount = masterUnits.length;
            let scrappedCount = 0;
            let visibleCards = 0;

            masterUnits.forEach(unit => {
                if (unit.isScrapped) scrappedCount++;
                const rec = inspectionRecords[unit.id];
                const isDone = !!rec;
                if (isDone && !unit.isScrapped) doneCount++;

                if (search && !unit.id.toLowerCase().includes(search) && !unit.type.toLowerCase().includes(search)) return;
                if (statusFilter === "DONE" && (!isDone || unit.isScrapped)) return;
                if (statusFilter === "PENDING" && (isDone || unit.isScrapped)) return;
                if (statusFilter === "SCRAPPED" && !unit.isScrapped) return;

                visibleCards++;
                const card = document.createElement("div");
                card.className = `bg-white rounded-xl border ${unit.isScrapped ? 'border-rose-200 bg-rose-50/30' : isDone ? 'border-emerald-200' : 'border-slate-200'} shadow-xs p-4 flex flex-col justify-between`;

                let badgeHtml = unit.isScrapped 
                    ? `<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-rose-100 text-rose-700">Afkir / 报废</span>`
                    : isDone 
                        ? `<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-emerald-100 text-emerald-800"><i class="fa-solid fa-circle-check"></i> Selesai</span>`
                        : `<span class="px-2 py-0.5 rounded-full text-[11px] font-bold bg-amber-100 text-amber-800"><i class="fa-solid fa-clock"></i> Belum</span>`;

                card.innerHTML = `
                    <div>
                        <div class="flex items-start justify-between gap-2 border-b border-slate-100 pb-2.5 mb-3">
                            <div>
                                <h3 class="font-bold text-slate-900 text-base">${unit.id}</h3>
                                <span class="inline-block text-[11px] text-slate-500 font-medium bg-slate-100 px-2 py-0.5 rounded mt-0.5">Model: ${unit.type}</span>
                            </div>
                            ${badgeHtml}
                        </div>
                        ${isDone ? `
                            <div class="space-y-1.5 text-xs">
                                <div class="text-slate-600"><span class="font-semibold text-slate-700">Komponen:</span> ${rec.item}</div>
                                <div class="text-slate-600"><span class="font-semibold text-slate-700">Status:</span> ${rec.status}</div>
                                ${rec.notes ? `<div class="text-slate-600 bg-slate-50 p-2 rounded border border-slate-100 italic mt-1 text-[11px]">"${rec.notes}"</div>` : ''}
                                ${rec.photo ? `<div class="mt-2.5"><img src="${rec.photo}" class="w-full h-32 object-cover rounded-lg border border-slate-200"></div>` : ''}
                            </div>
                        ` : `<div class="py-4 text-center text-xs text-slate-400 italic">Belum ada data pemeriksaan</div>`}
                    </div>
                    <div class="mt-4 pt-2 border-t border-slate-100">
                        ${unit.isScrapped ? `<button disabled class="w-full py-1.5 bg-slate-100 text-slate-400 rounded-lg text-xs font-semibold cursor-not-allowed">Unit Non-Aktif</button>` : `
                            <button onclick="openInputModal('${unit.id}')" class="w-full py-2 ${isDone ? 'bg-slate-100 hover:bg-slate-200 text-slate-700' : 'bg-blue-600 hover:bg-blue-700 text-white'} rounded-lg text-xs font-semibold shadow-xs transition flex items-center justify-center gap-1.5">
                                <i class="fa-solid ${isDone ? 'fa-pen' : 'fa-plus'}"></i>
                                <span>${isDone ? 'Edit Data Inspeksi' : 'Input Inspeksi / 录入'}</span>
                            </button>
                        `}
                    </div>
                `;
                container.appendChild(card);
            });

            if (visibleCards === 0) emptyState.classList.remove("hidden");
            else emptyState.classList.add("hidden");

            document.getElementById("statTotal").innerText = totalCount;
            document.getElementById("statDone").innerText = doneCount;
            document.getElementById("statPending").innerText = (totalCount - scrappedCount) - doneCount;
            document.getElementById("statScrapped").innerText = scrappedCount;
        }

        function openInputModal(unitId) {
            document.getElementById("modalUnitId").value = unitId;
            document.getElementById("modalUnitTitle").innerText = `Inspeksi Unit ${unitId}`;
            const itemSelect = document.getElementById("modalItemSelect");
            itemSelect.innerHTML = "";
            getTargetItems(currentWeek).forEach(it => {
                let opt = document.createElement("option");
                opt.value = it;
                opt.text = it;
                itemSelect.appendChild(opt);
            });

            tempPhotoBase64 = "";
            const rec = inspectionRecords[unitId];
            if (rec) {
                document.getElementById("modalItemSelect").value = rec.item;
                document.getElementById("modalStatusSelect").value = rec.status;
                document.getElementById("modalNotes").value = rec.notes || "";
                if (rec.photo) {
                    tempPhotoBase64 = rec.photo;
                    document.getElementById("modalPhotoPreview").src = rec.photo;
                    document.getElementById("photoPreviewContainer").classList.remove("hidden");
                } else {
                    document.getElementById("photoPreviewContainer").classList.add("hidden");
                }
            } else {
                document.getElementById("inspectionForm").reset();
                document.getElementById("modalPhotoInput").value = "";
                document.getElementById("photoPreviewContainer").classList.add("hidden");
            }
            document.getElementById("inputModal").classList.remove("hidden");
        }

        function closeModal() { document.getElementById("inputModal").classList.add("hidden"); }

        function previewPhoto(e) {
            const file = e.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = function(evt) {
                    tempPhotoBase64 = evt.target.result;
                    document.getElementById("modalPhotoPreview").src = tempPhotoBase64;
                    document.getElementById("photoPreviewContainer").classList.remove("hidden");
                };
                reader.readAsDataURL(file);
            }
        }

        function removePhoto() {
            tempPhotoBase64 = "";
            document.getElementById("modalPhotoInput").value = "";
            document.getElementById("photoPreviewContainer").classList.add("hidden");
        }

        function saveInspection(e) {
            e.preventDefault();
            const unitId = document.getElementById("modalUnitId").value;
            const recordData = {
                id: `${currentWeek}_${unitId}`,
                week: currentWeek,
                unitId: unitId,
                item: document.getElementById("modalItemSelect").value,
                status: document.getElementById("modalStatusSelect").value,
                notes: document.getElementById("modalNotes").value.trim(),
                photo: tempPhotoBase64,
                timestamp: new Date().toISOString()
            };

            const tx = db.transaction("records", "readwrite");
            tx.objectStore("records").put(recordData);
            tx.oncomplete = () => {
                inspectionRecords[unitId] = recordData;
                closeModal();
                renderUnits();
            };
        }

        function openManualEditModal() {
            document.getElementById("manualItemsInput").value = getTargetItems(currentWeek).join(", ");
            document.getElementById("manualModal").classList.remove("hidden");
        }
        function closeManualModal() { document.getElementById("manualModal").classList.add("hidden"); }
        function saveManualItems() {
            localStorage.setItem(`manual_items_w${currentWeek}`, document.getElementById("manualItemsInput").value.trim());
            closeManualModal();
            loadWeekData();
        }

        // Fitur Backup dan Restore JSON
        function backupDataJSON() {
            const tx = db.transaction("records", "readonly");
            const req = tx.objectStore("records").getAll();
            req.onsuccess = () => {
                const dataBlob = new Blob([JSON.stringify(req.result, null, 2)], { type: 'application/json' });
                const url = URL.createObjectURL(dataBlob);
                const a = document.createElement('a');
                a.href = url;
                a.download = `Backup_Inspeksi_Armada_${new Date().toISOString().slice(0,10)}.json`;
                a.click();
                URL.revokeObjectURL(url);
            };
        }

        function restoreDataJSON(e) {
            const file = e.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(evt) {
                try {
                    const importedData = JSON.parse(evt.target.result);
                    const tx = db.transaction("records", "readwrite");
                    const store = tx.objectStore("records");
                    importedData.forEach(item => store.put(item));
                    tx.oncomplete = () => {
                        alert("Data berhasil dipulihkan!");
                        loadWeekData();
                    };
                } catch(err) {
                    alert("Format file JSON tidak valid!");
                }
            };
            reader.readAsText(file);
        }

        function exportToExcel() {
            const data = masterUnits.map(unit => {
                const rec = inspectionRecords[unit.id] || {};
                return {
                    "Nomor Unit / 车号": unit.id,
                    "Tipe Armada / 型号": unit.type,
                    "Minggu Ke / 周次": currentWeek,
                    "Komponen / 检查部位": rec.item || "-",
                    "Kondisi Status / 状态": unit.isScrapped ? "Afkir / 报废" : (rec.status || "Belum Diperiksa"),
                    "Catatan / 备注": rec.notes || "-"
                };
            });
            const ws = XLSX.utils.json_to_sheet(data);
            const wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, `Minggu ${currentWeek}`);
            XLSX.writeFile(wb, `Laporan_Minggu_${currentWeek}.xlsx`);
        }

        function exportToWordTable() {
            let rowsHtml = "";
            masterUnits.forEach(unit => {
                const rec = inspectionRecords[unit.id] || {};
                const photoHtml = rec.photo ? `<img src="${rec.photo}" width="120" style="border-radius:4px;">` : 'Tidak Ada Foto';
                rowsHtml += `<tr>
                    <td style="padding:8px; border:1px solid #ccc; text-align:center;">${unit.id}</td>
                    <td style="padding:8px; border:1px solid #ccc; text-align:center;">${unit.type}</td>
                    <td style="padding:8px; border:1px solid #ccc;">${rec.item || '-'}</td>
                    <td style="padding:8px; border:1px solid #ccc; text-align:center;">${unit.isScrapped ? 'Afkir' : (rec.status || 'Belum')}</td>
                    <td style="padding:8px; border:1px solid #ccc; text-align:center;">${photoHtml}</td>
                    <td style="padding:8px; border:1px solid #ccc;">${rec.notes || '-'}</td>
                </tr>`;
            });
            const htmlContent = `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
            <head><meta charset='utf-8'></head>
            <body>
                <h2 style="text-align:center;">LAPORAN INSPEKSI ARMADA - MINGGU ${currentWeek}</h2>
                <table style="width:100%; border-collapse:collapse; margin-top:20px;">
                    <thead><tr style="background:#f2f2f2;">
                        <th style="padding:8px; border:1px solid #ccc;">No. Unit</th>
                        <th style="padding:8px; border:1px solid #ccc;">Tipe</th>
                        <th style="padding:8px; border:1px solid #ccc;">Bagian</th>
                        <th style="padding:8px; border:1px solid #ccc;">Status</th>
                        <th style="padding:8px; border:1px solid #ccc;">Foto</th>
                        <th style="padding:8px; border:1px solid #ccc;">Catatan</th>
                    </tr></thead>
                    <tbody>${rowsHtml}</tbody>
                </table>
            </body></html>`;
            const blob = new Blob(['\ufeff', htmlContent], { type: 'application/msword' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = `Laporan_Foto_Minggu_${currentWeek}.doc`;
            a.click();
        }

        window.onload = initDB;
    </script>
</body>
</html>