/* Bulk product entry for the existing admin product section. Excel files are parsed in the browser. */
(() => {
    const DRAFT_KEY = 'attractive_admin_bulk_products_draft_v1';
    const MAX_SHEET_BYTES = 5 * 1024 * 1024;
    const MAX_ROWS = 250;
    const MAX_USED_IMAGES = 60; // matches /api/upload's per-IP rate limit
    const MAX_IMAGE_BYTES = 4.2 * 1024 * 1024; // leaves room below Vercel's request body cap
    let rows = [];
    let newCategories = [];
    let filesByName = new Map();
    let fileUrls = new Map();
    let previewReady = false;
    let saving = false;

    const $ = (id) => document.getElementById(id);
    const esc = (value) => ProCable.sanitize(String(value == null ? '' : value));
    const normalizeHeader = (value) => String(value == null ? '' : value).trim().toLowerCase().normalize('NFKC').replace(/[\u064B-\u065F\u0670\u0640]/g, '').replace(/[أإآ]/g, 'ا').replace(/[^\p{L}\p{N}]/gu, '');
    const normalizeName = (value) => String(value || '').trim().toLowerCase().normalize('NFKC').replace(/[\u064B-\u065F\u0670\u0640]/g, '').replace(/[أإآ]/g, 'ا').replace(/\s+/g, ' ');
    const normalizeFile = (value) => String(value || '').trim().replace(/\\/g, '/').split('/').pop().normalize('NFC').toLowerCase();
    const policy = () => document.querySelector('input[name="bulkDuplicatePolicy"]:checked')?.value || 'skip';

    const HEADER_ALIASES = {
        nameAr: ['اسم المنتج', 'اسم المنتج بالعربي', 'الاسم بالعربي', 'namear', 'name', 'productname', 'product'],
        nameEn: ['الاسم بالانجليزي', 'اسم المنتج بالانجليزي', 'nameen', 'englishname', 'productnameen'],
        price: ['السعر', 'السعر الحالي', 'price', 'currentprice'],
        originalPrice: ['السعر الاصلي', 'السعر قبل الخصم', 'originalprice', 'oldprice'],
        category: ['التصنيف', 'القسم', 'category', 'categoryid'],
        descAr: ['الوصف بالعربي', 'الوصف', 'المواصفات', 'descar', 'description'],
        descEn: ['الوصف بالانجليزي', 'descen', 'englishdescription', 'descriptionen'],
        tag: ['الموديل', 'الوسم', 'موديل او وسم', 'tag', 'model', 'sku'],
        imageFiles: ['اسماء الصور', 'اسماء ملفات الصور', 'اسم الصورة', 'اسم ملف الصورة', 'اسماء الصور مفصولة بفاصلة منقوطة', 'image', 'images', 'imagefile', 'imagefiles', 'photofiles'],
        sortOrder: ['الترتيب', 'ترتيب الظهور', 'sortorder', 'order'],
        badge: ['الشارة', 'شارة', 'badge'],
        inStock: ['متوفر', 'متوفر بالمخزون', 'instock', 'stock']
    };
    const FIELD_BY_HEADER = new Map();
    Object.entries(HEADER_ALIASES).forEach(([field, aliases]) => aliases.forEach(alias => FIELD_BY_HEADER.set(normalizeHeader(alias), field)));

    function message(text, kind = 'info') {
        const el = $('bulkProductsMessage');
        if (!el) return;
        const styles = { info: 'bg-blue-50 text-blue-800 border border-blue-200', error: 'bg-red-50 text-red-800 border border-red-200', success: 'bg-green-50 text-green-800 border border-green-200', warning: 'bg-amber-50 text-amber-900 border border-amber-200' };
        el.className = `rounded-xl p-3 mb-3 text-xs font-bold ${styles[kind] || styles.info}`;
        el.textContent = text;
    }

    function clearMessage() {
        const el = $('bulkProductsMessage');
        if (el) { el.className = 'hidden'; el.textContent = ''; }
    }

    function showModal() {
        const modal = $('bulkProductsModal');
        if (!modal) return;
        modal.classList.remove('hidden');
        modal.classList.add('flex');
        setTimeout(() => modal.classList.add('active'), 10);
        const hasDraft = !!localStorage.getItem(DRAFT_KEY);
        $('bulkLoadDraftBtn')?.classList.toggle('hidden', !hasDraft);
        if (rows.length) renderPreview();
    }

    function openBulkProductsModal() { showModal(); }

    function closeBulkProductsModal() {
        if (saving) return;
        const modal = $('bulkProductsModal');
        modal?.classList.remove('active');
        setTimeout(() => {
            modal?.classList.remove('flex');
            modal?.classList.add('hidden');
        }, 220);
        fileUrls.forEach(url => URL.revokeObjectURL(url));
        fileUrls = new Map();
        filesByName = new Map();
        rows = [];
        newCategories = [];
        previewReady = false;
        if ($('bulkProductsPreview')) $('bulkProductsPreview').classList.add('hidden');
        if ($('bulkProductsSheetInput')) $('bulkProductsSheetInput').value = '';
        if ($('bulkProductsImagesInput')) $('bulkProductsImagesInput').value = '';
        if ($('bulkSheetLabel')) $('bulkSheetLabel').textContent = 'اختيار ملف Excel (.xlsx)';
        if ($('bulkImagesLabel')) $('bulkImagesLabel').textContent = 'اختيار صور المنتجات';
        if ($('bulkProductsProgress')) $('bulkProductsProgress').textContent = '';
        $('bulkSaveDraftBtn')?.setAttribute('disabled', 'disabled');
        clearMessage();
    }

    function downloadBulkProductsTemplate() {
        const link = document.createElement('a');
        link.href = 'templates/product-import-template.xlsx';
        link.download = 'قالب-المنتجات-والتصنيفات-Attractive.xlsx';
        document.body.appendChild(link);
        link.click();
        link.remove();
    }

    async function handleBulkProductsSheet(file) {
        clearMessage();
        if (!file) return;
        rows = [];
        newCategories = [];
        previewReady = false;
        $('bulkProductsPreview')?.classList.add('hidden');
        $('bulkSaveDraftBtn')?.setAttribute('disabled', 'disabled');
        if (!window.XLSX || XLSX.version !== '0.20.3') return message('تعذر تحميل إصدار قارئ Excel الموثوق. أعد تحميل الصفحة.', 'error');
        if (!/\.xlsx$/i.test(file.name) || file.size > MAX_SHEET_BYTES || file.size < 4) return message('اختار ملف .xlsx صالحًا لا يتجاوز ٥ ميجابايت.', 'error');
        if ($('bulkSheetLabel')) $('bulkSheetLabel').textContent = file.name;
        try {
            const buf = await file.arrayBuffer();
            const bytes = new Uint8Array(buf);
            if (bytes[0] !== 0x50 || bytes[1] !== 0x4b) throw new Error('ملف Excel غير صالح.');
            const wb = XLSX.read(buf, { type: 'array', bookVBA: false, cellHTML: false, cellFormula: true, WTF: true });
            const sheetName = wb.SheetNames?.[0];
            if (!sheetName) throw new Error('ملف Excel لا يحتوي على ورقة بيانات.');
            const ws = wb.Sheets[sheetName];
            const cells = Object.keys(ws).filter(key => key[0] !== '!');
            if (cells.some(key => ws[key]?.f)) throw new Error('القالب لا يقبل معادلات Excel. اكتب قيمًا ثابتة في الخلايا.');
            const matrix = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false, blankrows: false, skipHidden: true });
            if (matrix.length < 2) throw new Error('الشيت فارغ أو لا يحتوي على صفوف منتجات.');
            if (matrix.length - 1 > MAX_ROWS) throw new Error(`الحد الأقصى ${MAX_ROWS} منتج في المرة الواحدة.`);
            const headerFields = matrix[0].map(value => FIELD_BY_HEADER.get(normalizeHeader(value)) || '');
            for (const required of ['nameAr', 'price', 'category', 'imageFiles']) {
                if (!headerFields.includes(required)) throw new Error(`عمود مطلوب غير موجود: ${required === 'nameAr' ? 'اسم المنتج' : required === 'price' ? 'السعر' : required === 'category' ? 'التصنيف' : 'أسماء الصور'}. نزّل القالب الجاهز.`);
            }
            const baseCategories = ProCable.getCategories().filter(c => c.id !== 'all');
            const categorySheetName = wb.SheetNames.find(name => normalizeHeader(name) === normalizeHeader('إضافة تصنيفات'));
            if (categorySheetName) {
                const categoryWs = wb.Sheets[categorySheetName];
                const categoryCells = Object.keys(categoryWs).filter(key => key[0] !== '!');
                if (categoryCells.some(key => categoryWs[key]?.f)) throw new Error('ورقة التصنيفات لا تقبل معادلات Excel. اكتب قيمًا ثابتة.');
                const categoryMatrix = XLSX.utils.sheet_to_json(categoryWs, { header: 1, defval: '', raw: false, blankrows: false, skipHidden: true });
                const categoryHeaders = (categoryMatrix[0] || []).map(normalizeHeader);
                const idColumn = categoryHeaders.findIndex(h => ['معرفالتصنيف', 'المعرف', 'categoryid', 'id', 'slug'].includes(h));
                const arColumn = categoryHeaders.findIndex(h => ['الاسمالعربي', 'اسمبالعربي', 'namear', 'arabicname'].includes(h));
                const enColumn = categoryHeaders.findIndex(h => ['الاسمبالانجليزي', 'nameen', 'englishname'].includes(h));
                const entered = categoryMatrix.slice(1).filter(cells => cells.some(value => String(value || '').trim()));
                if (entered.length && (idColumn < 0 || arColumn < 0)) throw new Error('أعمدة ورقة «إضافة تصنيفات» يجب أن تحتوي على «معرف التصنيف» و«الاسم بالعربي».');
                if (entered.length > 50) throw new Error('الحد الأقصى ٥٠ تصنيفًا جديدًا في الملف الواحد.');
                const ids = new Set(baseCategories.map(c => normalizeName(c.id)));
                newCategories = entered.map((cells, index) => {
                    const id = String(cells[idColumn] || '').trim().toLowerCase();
                    const nameAr = String(cells[arColumn] || '').trim();
                    const nameEn = enColumn >= 0 ? String(cells[enColumn] || '').trim() : '';
                    if (!/^[a-z0-9_-]{1,64}$/.test(id) || id === 'all') throw new Error(`معرف التصنيف في صف ${index + 2} غير صالح. استخدم حروفًا إنجليزية صغيرة أو أرقامًا أو - أو _.`);
                    if (!nameAr || nameAr.length > 120 || nameEn.length > 120) throw new Error(`اسم التصنيف في صف ${index + 2} مطلوب، والحد الأقصى ١٢٠ حرفًا.`);
                    if (ids.has(normalizeName(id))) throw new Error(`معرف التصنيف ${id} مستخدم بالفعل أو مكرر في الملف.`);
                    ids.add(normalizeName(id));
                    return { id, nameAr, nameEn: nameEn || nameAr, image: '', images: [], video: '', icon: 'category', hidden: false };
                });
            }
            const categories = [...baseCategories, ...newCategories];
            rows = matrix.slice(1).map((cells, i) => {
                const raw = {};
                headerFields.forEach((field, j) => { if (field) raw[field] = cells[j]; });
                const categoryValue = String(raw.category || '').trim();
                const category = categories.find(c => normalizeName(c.id) === normalizeName(categoryValue) || normalizeName(c.nameAr) === normalizeName(categoryValue) || normalizeName(c.nameEn) === normalizeName(categoryValue));
                const imageFiles = String(raw.imageFiles || '').split(/[;|\n]/).map(s => s.trim()).filter(Boolean);
                const stockVal = normalizeName(raw.inStock);
                return {
                    line: i + 2,
                    nameAr: String(raw.nameAr || '').trim(),
                    nameEn: String(raw.nameEn || '').trim(),
                    price: String(raw.price || '').trim(),
                    originalPrice: String(raw.originalPrice || '').trim(),
                    category: category?.id || categoryValue,
                    descAr: String(raw.descAr || '').trim(),
                    descEn: String(raw.descEn || '').trim(),
                    tag: String(raw.tag || '').trim(),
                    imageFiles,
                    sortOrder: String(raw.sortOrder || '').trim(),
                    badge: String(raw.badge || '').trim(),
                    inStock: stockVal ? !['لا', 'no', 'false', '0', 'غيرمتوفر'].includes(stockVal) : true,
                    inStockProvided: !!stockVal
                };
            }).filter(row => row.nameAr || row.price || row.imageFiles.length);
            if (!rows.length) throw new Error('لا توجد صفوف منتجات في ورقة Excel الأولى.');
            previewReady = true;
            $('bulkSaveDraftBtn')?.removeAttribute('disabled');
            renderPreview();
            $('bulkProductsPreview')?.classList.remove('hidden');
            if (!filesByName.size) message('اختار صور المنتجات أيضًا. أسماء الملفات في عمود الصور يجب أن تطابق أسماء الصور المختارة.', 'warning');
        } catch (error) {
            rows = [];
            newCategories = [];
            previewReady = false;
            $('bulkProductsPreview')?.classList.add('hidden');
            message(error.message || 'تعذر قراءة ملف Excel.', 'error');
        }
    }

    function handleBulkProductsImages(fileList) {
        clearMessage();
        fileUrls.forEach(url => URL.revokeObjectURL(url));
        fileUrls = new Map();
        filesByName = new Map();
        if ($('bulkImagesLabel')) $('bulkImagesLabel').textContent = 'اختيار صور المنتجات';
        const files = Array.from(fileList || []);
        if (files.length > 100) {
            $('bulkProductsImagesInput').value = '';
            if (rows.length) renderPreview();
            return message('اختار ١٠٠ صورة أو أقل في المرة الواحدة.', 'error');
        }
        const allowed = /\.(jpe?g|png|webp|gif)$/i;
        for (const file of files) {
            if (!allowed.test(file.name) || !/^image\/(jpeg|png|webp|gif)$/i.test(file.type || 'image/jpeg')) {
                $('bulkProductsImagesInput').value = '';
                filesByName.clear();
                fileUrls.forEach(url => URL.revokeObjectURL(url));
                fileUrls.clear();
                if (rows.length) renderPreview();
                return message(`نوع الصورة غير مدعوم: ${file.name}. استخدم JPG أو PNG أو WebP أو GIF.`, 'error');
            }
            if (!file.size || file.size > 20 * 1024 * 1024) {
                $('bulkProductsImagesInput').value = '';
                filesByName.clear();
                fileUrls.forEach(url => URL.revokeObjectURL(url));
                fileUrls.clear();
                if (rows.length) renderPreview();
                return message(`حجم الصورة ${file.name} غير صالح أو أكبر من ٢٠ ميجابايت.`, 'error');
            }
            const key = normalizeFile(file.name);
            if (filesByName.has(key)) filesByName.set(key, null); // ambiguous duplicate name
            else filesByName.set(key, file);
            fileUrls.set(file, URL.createObjectURL(file));
        }
        if ($('bulkImagesLabel')) $('bulkImagesLabel').textContent = files.length ? `${files.length} صورة مختارة` : 'اختيار صور المنتجات';
        if (rows.length) renderPreview();
        if (files.length) message('تم اختيار الصور. تأكد أن أسماءها في الشيت مطابقة لأسماء الملفات.', 'success');
    }

    function parseMoney(value) {
        let text = String(value == null ? '' : value).trim().replace(/[٠-٩]/g, d => String('٠١٢٣٤٥٦٧٨٩'.indexOf(d))).replace(/[٬,\s]/g, '').replace(/٫/g, '.').replace(/[^\d.\-]/g, '');
        const n = Number(text);
        return Number.isFinite(n) ? n : NaN;
    }

    function existingDuplicate(row) {
        const key = normalizeName(row.nameAr);
        if (!key) return null;
        return ProCable.getProducts().find(p => normalizeName(p.nameAr) === key) || null;
    }

    function rowValidation() {
        const categories = [...ProCable.getCategories().filter(c => c.id !== 'all'), ...newCategories];
        const names = new Map();
        rows.forEach((row, i) => {
            const key = normalizeName(row.nameAr);
            if (key) names.set(key, [...(names.get(key) || []), i]);
        });
        return rows.map((row, index) => {
            const errors = [];
            const name = String(row.nameAr || '').trim();
            const price = parseMoney(row.price);
            if (!name) errors.push('اسم المنتج مطلوب');
            if (name.length > 180) errors.push('الاسم أطول من ١٨٠ حرفًا');
            if (!Number.isFinite(price) || price <= 0 || price > 100000000) errors.push('السعر غير صالح');
            if (String(row.nameEn || '').length > 180 || String(row.tag || '').length > 120) errors.push('الاسم الإنجليزي أو الموديل أطول من الحد المسموح');
            if (String(row.descAr || '').length > 5000 || String(row.descEn || '').length > 5000) errors.push('الوصف أطول من ٥٠٠٠ حرف');
            if (!categories.some(c => String(c.id) === String(row.category))) errors.push('التصنيف غير موجود');
            if (!row.imageFiles?.length) errors.push('اسم صورة واحد على الأقل مطلوب');
            const resolved = [];
            (row.imageFiles || []).forEach(imageName => {
                const file = filesByName.get(normalizeFile(imageName));
                if (file === null) errors.push(`اسم الصورة مكرر في الملفات: ${imageName}`);
                else if (!file) errors.push(`الصورة غير مختارة أو اسمها مختلف: ${imageName}`);
                else resolved.push(file);
            });
            if (row.imageFiles?.length > 8) errors.push('الحد الأقصى ٨ صور لكل منتج');
            if (row.originalPrice && (!Number.isFinite(parseMoney(row.originalPrice)) || parseMoney(row.originalPrice) < 0)) errors.push('السعر الأصلي غير صالح');
            if (row.sortOrder && (!Number.isInteger(parseMoney(row.sortOrder)) || parseMoney(row.sortOrder) < 0)) errors.push('ترتيب الظهور يجب أن يكون رقمًا صحيحًا غير سالب');
            if (key && names.get(key)?.length > 1) errors.push('اسم المنتج مكرر داخل الشيت');
            const duplicate = existingDuplicate(row);
            return { errors, files: resolved, duplicate, index };
        });
    }

    function renderPreview() {
        if (!previewReady || !rows.length) return;
        const categories = [...ProCable.getCategories().filter(c => c.id !== 'all'), ...newCategories];
        const categorySummary = $('bulkProductsNewCategories');
        if (categorySummary) {
            categorySummary.classList.toggle('hidden', !newCategories.length);
            categorySummary.textContent = newCategories.length
                ? `سيُضاف ${newCategories.length} تصنيف جديد عند حفظ المنتجات: ${newCategories.map(c => `${c.nameAr} (${c.id})`).join('، ')}`
                : '';
        }
        const checks = rowValidation();
        const duplicatesSkipped = checks.filter(c => c.duplicate && policy() === 'skip').length;
        const errorCount = checks.reduce((n, c) => n + (c.duplicate && policy() === 'skip' ? 0 : c.errors.length), 0);
        const ready = checks.filter(c => !c.errors.length && (!c.duplicate || policy() === 'update')).length;
        $('bulkProductsSummary').innerHTML = `<span class="text-green-700">${ready} جاهز للحفظ</span> <span class="text-gray-400">·</span> <span class="${errorCount ? 'text-red-700' : 'text-gray-600'}">${errorCount} خطأ</span> <span class="text-gray-400">·</span> <span class="text-amber-700">${duplicatesSkipped} مكرر سيتم تخطيه</span>`;
        $('bulkProductsPreviewBody').innerHTML = rows.map((row, i) => {
            const check = checks[i];
            const options = categories.map(c => `<option value="${esc(c.id)}" ${String(c.id) === String(row.category) ? 'selected' : ''}>${esc(c.nameAr || c.id)} (${esc(c.id)})</option>`).join('');
            const firstFile = check.files[0];
            const preview = firstFile && fileUrls.get(firstFile) ? `<img src="${esc(fileUrls.get(firstFile))}" class="w-12 h-12 object-cover rounded-lg border border-outline-variant" alt="معاينة صورة المنتج"/>` : '<span class="text-gray-400 text-[10px]">لا توجد معاينة</span>';
            let status = check.duplicate && policy() === 'skip' ? `<span class="text-amber-800">مكرر: ${esc(check.duplicate.nameAr || row.nameAr)} (سيتم التخطي)</span>` : check.errors.length ? `<span class="text-red-700">${check.errors.map(esc).join('، ')}</span>` : check.duplicate ? `<span class="text-amber-800">مكرر: ${esc(check.duplicate.nameAr || row.nameAr)} (سيتم التحديث)</span>` : '<span class="text-green-700">جاهز للإضافة</span>';
            return `<tr class="align-top ${check.errors.length ? 'bg-red-50/50' : ''}">
                <td class="p-2 font-mono text-gray-500">${row.line}</td>
                <td class="p-2"><input aria-label="اسم المنتج في الصف ${row.line}" value="${esc(row.nameAr)}" onchange="updateBulkProductField(${i},'nameAr',this.value)" class="w-full min-w-48 p-2 border border-outline-variant rounded-lg"/></td>
                <td class="p-2"><input aria-label="السعر في الصف ${row.line}" type="number" min="0" step="0.01" value="${esc(row.price)}" onchange="updateBulkProductField(${i},'price',this.value)" class="w-24 p-2 border border-outline-variant rounded-lg"/></td>
                <td class="p-2"><select aria-label="التصنيف في الصف ${row.line}" onchange="updateBulkProductField(${i},'category',this.value)" class="min-w-36 p-2 border border-outline-variant rounded-lg">${options}</select></td>
                <td class="p-2"><input aria-label="أسماء الصور في الصف ${row.line}" value="${esc((row.imageFiles || []).join('; '))}" onchange="updateBulkProductImages(${i},this.value)" class="w-full min-w-44 p-2 border border-outline-variant rounded-lg dir-ltr"/><span class="block mt-1 text-[9px] text-on-surface-variant">افصل الصور بـ ;</span></td>
                <td class="p-2">${preview}</td>
                <td class="p-2 leading-5">
                    <details class="mb-2"><summary class="cursor-pointer font-bold text-primary">تعديل التفاصيل الاختيارية</summary><div class="grid grid-cols-2 gap-1.5 min-w-64 mt-2">
                        <input aria-label="الاسم بالإنجليزي" placeholder="الاسم بالإنجليزي" value="${esc(row.nameEn)}" onchange="updateBulkProductField(${i},'nameEn',this.value)" class="col-span-2 p-1.5 border border-outline-variant rounded-md"/>
                        <input aria-label="السعر الأصلي" type="number" min="0" step="0.01" placeholder="السعر الأصلي" value="${esc(row.originalPrice)}" onchange="updateBulkProductField(${i},'originalPrice',this.value)" class="p-1.5 border border-outline-variant rounded-md"/>
                        <input aria-label="الموديل أو الوسم" placeholder="الموديل / الوسم" value="${esc(row.tag)}" onchange="updateBulkProductField(${i},'tag',this.value)" class="p-1.5 border border-outline-variant rounded-md"/>
                        <input aria-label="الوصف بالعربي" placeholder="الوصف بالعربي" value="${esc(row.descAr)}" onchange="updateBulkProductField(${i},'descAr',this.value)" class="col-span-2 p-1.5 border border-outline-variant rounded-md"/>
                        <input aria-label="الوصف بالإنجليزي" placeholder="الوصف بالإنجليزي" value="${esc(row.descEn)}" onchange="updateBulkProductField(${i},'descEn',this.value)" class="col-span-2 p-1.5 border border-outline-variant rounded-md"/>
                        <input aria-label="ترتيب الظهور" type="number" min="0" step="1" placeholder="ترتيب الظهور" value="${esc(row.sortOrder)}" onchange="updateBulkProductField(${i},'sortOrder',this.value)" class="p-1.5 border border-outline-variant rounded-md"/>
                        <select aria-label="الشارة" onchange="updateBulkProductField(${i},'badge',this.value)" class="p-1.5 border border-outline-variant rounded-md"><option value="">بدون شارة</option><option value="bestseller" ${row.badge === 'bestseller' ? 'selected' : ''}>الأكثر مبيعًا</option><option value="sale" ${row.badge === 'sale' ? 'selected' : ''}>عرض خاص</option><option value="choice" ${row.badge === 'choice' ? 'selected' : ''}>اختيار المتجر</option><option value="new" ${row.badge === 'new' ? 'selected' : ''}>جديد</option></select>
                        <label class="col-span-2 inline-flex items-center gap-1.5"><input type="checkbox" ${row.inStock ? 'checked' : ''} onchange="updateBulkProductField(${i},'inStock',this.checked)"/>متوفر بالمخزون</label>
                    </div></details><div class="flex items-start justify-between gap-2"><span>${status}</span><button type="button" onclick="removeBulkProductRow(${i})" class="shrink-0 text-[9px] text-error underline">إزالة من الدفعة</button></div>
                </td>
            </tr>`;
        }).join('');
        const validError = checks.some(c => c.errors.length && (!c.duplicate || policy() === 'update'));
        const submit = $('bulkSaveProductsBtn');
        if (submit) submit.disabled = saving || ready === 0 || validError;
        if (submit) submit.textContent = policy() === 'update' ? 'إضافة وتحديث المنتجات' : 'إضافة المنتجات';
        $('bulkSaveDraftBtn')?.toggleAttribute('disabled', saving);
    }

    function renderBulkProductsPreview() { renderPreview(); }

    function updateBulkProductField(index, field, value) {
        if (!rows[index] || !['nameAr', 'nameEn', 'price', 'originalPrice', 'category', 'descAr', 'descEn', 'tag', 'sortOrder', 'badge', 'inStock'].includes(field)) return;
        rows[index][field] = field === 'inStock' ? !!value : String(value || '').trim();
        if (field === 'inStock') rows[index].inStockProvided = true;
        renderPreview();
    }

    function updateBulkProductImages(index, value) {
        if (!rows[index]) return;
        rows[index].imageFiles = String(value || '').split(/[;|\n]/).map(s => s.trim()).filter(Boolean);
        renderPreview();
    }

    function removeBulkProductRow(index) {
        if (saving || !Number.isInteger(index) || index < 0 || index >= rows.length) return;
        rows.splice(index, 1);
        if (!rows.length) {
            previewReady = false;
            $('bulkProductsPreview')?.classList.add('hidden');
            $('bulkSaveDraftBtn')?.setAttribute('disabled', 'disabled');
            return message('تمت إزالة كل الصفوف من الدفعة.', 'info');
        }
        renderPreview();
    }

    function previewBulkProducts() {
        clearMessage();
        if (!previewReady || !rows.length) return message('اختار ملف Excel صالحًا أولًا.', 'warning');
        renderPreview();
        $('bulkProductsPreview')?.classList.remove('hidden');
        const checks = rowValidation();
        if (checks.some(c => c.errors.length && (!c.duplicate || policy() === 'update'))) message('راجع الصفوف المعلّمة بالأحمر. لن يتم الحفظ قبل تصحيح أخطاء المنتجات التي ستُضاف أو تُحدّث.', 'warning');
    }

    function saveBulkProductsDraft() {
        if (!rows.length) return message('اقرأ ملف Excel قبل حفظ المسودة.', 'warning');
        try {
            localStorage.setItem(DRAFT_KEY, JSON.stringify({ version: 2, savedAt: new Date().toISOString(), rows, newCategories }));
            $('bulkLoadDraftBtn')?.classList.remove('hidden');
            message('تم حفظ المسودة على هذا الجهاز. عند استعادتها، اختر ملفات الصور مرة أخرى.', 'success');
        } catch (error) { message('تعذر حفظ المسودة في مساحة المتصفح.', 'error'); }
    }

    function loadBulkProductsDraft() {
        try {
            const draft = JSON.parse(localStorage.getItem(DRAFT_KEY) || 'null');
            if (!draft || ![1, 2].includes(draft.version) || !Array.isArray(draft.rows) || !draft.rows.length) return message('لا توجد مسودة صالحة للاستعادة.', 'warning');
            rows = draft.rows.slice(0, MAX_ROWS).map(row => Object.assign({}, row, { imageFiles: Array.isArray(row.imageFiles) ? row.imageFiles : [] }));
            const existingIds = new Set(ProCable.getCategories().map(category => normalizeName(category.id)));
            newCategories = draft.version === 2 && Array.isArray(draft.newCategories)
                ? draft.newCategories.filter(category => !existingIds.has(normalizeName(category.id))).slice(0, 50)
                : [];
            previewReady = true;
            $('bulkProductsPreview')?.classList.remove('hidden');
            $('bulkSaveDraftBtn')?.removeAttribute('disabled');
            renderPreview();
            message('تم استعادة المسودة. اختار صور المنتجات من الجهاز مرة أخرى لإكمال مطابقة الصور.', 'info');
        } catch (error) { message('تعذر قراءة المسودة المحفوظة.', 'error'); }
    }

    function fileToDataUrl(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = () => resolve(String(reader.result || ''));
            reader.onerror = () => reject(new Error(`تعذرت قراءة الصورة ${file.name}`));
            reader.readAsDataURL(file);
        });
    }

    async function optimizedImage(file) {
        if (file.size <= MAX_IMAGE_BYTES) return file;
        if (/\.gif$/i.test(file.name) || file.type === 'image/gif') throw new Error(`الصورة المتحركة ${file.name} أكبر من الحد المسموح (٤٫٢ ميجابايت). قلل حجمها قبل الرفع.`);
        const bitmap = await createImageBitmap(file);
        const scale = Math.min(1, 1800 / Math.max(bitmap.width, bitmap.height));
        const canvas = document.createElement('canvas');
        canvas.width = Math.max(1, Math.round(bitmap.width * scale));
        canvas.height = Math.max(1, Math.round(bitmap.height * scale));
        const ctx = canvas.getContext('2d');
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, canvas.width, canvas.height);
        ctx.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
        bitmap.close();
        const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', 0.82));
        if (!blob || blob.size > MAX_IMAGE_BYTES) throw new Error(`تعذر ضغط الصورة ${file.name} إلى الحجم المناسب. قلل أبعادها أو حجمها.`);
        const base = file.name.replace(/\.[^.]+$/, '').replace(/[^\w.-]+/g, '_').slice(0, 80) || 'product-image';
        return new File([blob], `${base}.jpg`, { type: 'image/jpeg', lastModified: Date.now() });
    }

    async function uploadProductImage(file, ordinal, total) {
        if (ProCable._isLocalAdminMode()) {
            $('bulkProductsProgress').textContent = `تجهيز الصور ${ordinal}/${total}: ${file.name}`;
            if (/\.gif$/i.test(file.name) || file.type === 'image/gif') {
                if (file.size > 1.2 * 1024 * 1024) throw new Error(`قلل حجم GIF ${file.name} إلى أقل من ١٫٢ ميجابايت في وضع الأدمن المحلي.`);
                return await fileToDataUrl(file);
            }
            return await ProCable.compressImage(file, 1200, 0.76);
        }
        const image = await optimizedImage(file);
        $('bulkProductsProgress').textContent = `رفع الصور ${ordinal}/${total}: ${file.name}`;
        const url = `/api/upload?folder=product-images&filename=${encodeURIComponent(image.name)}`;
        let response;
        try {
            response = await fetch(url, { method: 'POST', headers: ProCable._adminHeaders({ 'Content-Type': image.type }), body: image });
        } catch (error) { throw new Error(`تعذر الاتصال أثناء رفع ${file.name}.`); }
        const result = await response.json().catch(() => ({}));
        if (!response.ok || !result.url) throw new Error(result.error || `فشل رفع الصورة ${file.name} (${response.status}).`);
        return result.url;
    }

    function categoryFor(row) {
        return ProCable.getCategories().find(c => String(c.id) === String(row.category));
    }

    function rowProduct(row, imageUrls, id, sortOrder) {
        return {
            id,
            nameAr: row.nameAr.trim(), nameEn: row.nameEn.trim() || row.nameAr.trim(),
            descAr: row.descAr || '', descEn: row.descEn || row.descAr || '',
            category: row.category, sortOrder,
            badge: ['bestseller', 'sale', 'choice', 'new'].includes(row.badge) ? row.badge : '',
            tag: row.tag || '', icon: 'xlr-m',
            images: imageUrls, image: imageUrls[0] || '',
            video: '', videoType: 'file', coverType: 'image',
            heroMediaType: 'none', heroMediaUrl: '', heroPosterUrl: '',
            price: parseMoney(row.price),
            originalPrice: row.originalPrice ? parseMoney(row.originalPrice) : parseMoney(row.price),
            inStock: row.inStock !== false, hidden: false, datasheet: '', datasheetType: 'pdf'
        };
    }

    async function saveBulkProducts() {
        if (saving || !previewReady) return;
        clearMessage();
        const checks = rowValidation();
        const active = checks.filter(c => !c.duplicate || policy() === 'update');
        if (checks.some(c => c.errors.length && (!c.duplicate || policy() === 'update'))) return message('صحح أخطاء المنتجات التي ستُضاف أو تُحدّث أولًا.', 'error');
        if (!active.length) return message('كل المنتجات مكررة ومحدّد لها التخطي، لذلك لا توجد منتجات للحفظ.', 'warning');
        const uniqueFiles = Array.from(new Set(active.flatMap(c => c.files)));
        if (!uniqueFiles.length || uniqueFiles.length > MAX_USED_IMAGES) return message(`عدد الصور المستخدمة يجب أن يكون بين ١ و${MAX_USED_IMAGES}.`, 'error');
        if (!ProCable._isLocalAdminMode() && !ProCable._adminToken()) return message('جلسة الأدمن غير صالحة. سجّل الدخول مرة أخرى.', 'error');
        if (!ProCable._isLocalAdminMode() && !ProCable._serverSynced) return message('انتظر مزامنة بيانات الموقع من السيرفر، ثم حاول الحفظ.', 'warning');

        saving = true;
        const button = $('bulkSaveProductsBtn');
        if (button) { button.disabled = true; button.textContent = 'جارٍ الحفظ…'; }
        const uploaded = new Map();
        try {
            for (let i = 0; i < uniqueFiles.length; i++) uploaded.set(uniqueFiles[i], await uploadProductImage(uniqueFiles[i], i + 1, uniqueFiles.length));
            const activeCategoryIds = new Set(active.map(check => String(rows[check.index].category)));
            const categoriesToAdd = newCategories.filter(category => activeCategoryIds.has(String(category.id)));
            const current = ProCable.getProducts();
            const next = current.slice();
            const maxSort = current.reduce((m, p) => Math.max(m, Number(p.sortOrder) || 0), 0);
            let added = 0, updated = 0, skipped = 0;
            for (let i = 0; i < rows.length; i++) {
                const row = rows[i];
                const check = checks[i];
                if (!check.duplicate && check.errors.length) continue;
                if (check.duplicate && policy() === 'skip') { skipped++; continue; }
                const imageUrls = check.files.map(file => uploaded.get(file)).filter(Boolean);
                if (!imageUrls.length) throw new Error(`لم يتم تحميل صور الصف ${row.line}.`);
                const old = check.duplicate ? next.find(p => p.id === check.duplicate.id) : null;
                if (old) {
                    const update = rowProduct(row, imageUrls, old.id, row.sortOrder ? parseMoney(row.sortOrder) : old.sortOrder);
                    old.nameAr = update.nameAr;
                    old.nameEn = row.nameEn.trim() || old.nameEn || update.nameEn;
                    old.price = update.price;
                    old.category = update.category;
                    old.images = update.images;
                    old.image = update.image;
                    old.coverType = 'image';
                    old.video = '';
                    if (row.originalPrice) old.originalPrice = update.originalPrice;
                    if (row.descAr) old.descAr = row.descAr;
                    if (row.descEn) old.descEn = row.descEn;
                    if (row.tag) old.tag = row.tag;
                    if (row.badge) old.badge = update.badge;
                    if (row.inStockProvided) old.inStock = row.inStock;
                    updated++;
                } else {
                    let id = Date.now() + i;
                    while (next.some(p => Number(p.id) === id)) id++;
                    next.unshift(rowProduct(row, imageUrls, id, row.sortOrder ? parseMoney(row.sortOrder) : maxSort + added + 1));
                    added++;
                }
            }
            if (!added && !updated) return message('لا توجد تغييرات للحفظ.', 'warning');
            if (next.length > 5000) throw new Error('الكتالوج سيحتوي على أكثر من ٥٠٠٠ منتج، وهو الحد الأقصى المسموح.');
            if (JSON.stringify(next).length > 3.3 * 1024 * 1024) throw new Error('حجم بيانات الكتالوج بعد الإضافة كبير. قلل عدد المنتجات أو طول الأوصاف ثم حاول مرة أخرى.');
            if (categoriesToAdd.length) {
                $('bulkProductsProgress').textContent = 'حفظ التصنيفات الجديدة…';
                const currentCategories = ProCable.getCategories();
                const nextCategories = [...currentCategories, ...categoriesToAdd];
                const categoriesSaved = await ProCable.pushToServer('categories', nextCategories);
                if (!categoriesSaved) throw new Error('لم يتم حفظ التصنيفات الجديدة على الموقع، لذلك لم تُحفظ المنتجات. أعد المحاولة بعد التأكد من الاتصال.');
                localStorage.setItem('attractive_categories', JSON.stringify(nextCategories));
                newCategories = newCategories.filter(category => !categoriesToAdd.some(saved => saved.id === category.id));
            }
            $('bulkProductsProgress').textContent = 'حفظ الكتالوج…';
            const remoteSaved = await ProCable.pushToServer('products', next);
            if (!remoteSaved) throw new Error('لم يؤكد السيرفر حفظ الكتالوج. الصور التي رُفعت لم تُربط بمنتجات. أعد المحاولة بعد التأكد من الاتصال.');
            let cacheSaved = true;
            try {
                localStorage.setItem('attractive_products', JSON.stringify(next));
                try { localStorage.setItem('procable_products', JSON.stringify(next)); } catch (error) { localStorage.removeItem('procable_products'); }
            } catch (error) {
                if (ProCable._isLocalAdminMode()) throw new Error('مساحة التخزين المحلي ممتلئة، لذلك لم تُحفظ المنتجات. قلل حجم الصور أو عددها وحاول مرة أخرى.');
                cacheSaved = false; // The server has saved successfully; a reload will pull the catalog again.
            }
            try { localStorage.removeItem(DRAFT_KEY); } catch (error) { /* saved products are unaffected */ }
            renderAdminProducts();
            if (typeof renderAdminCategories === 'function') renderAdminCategories();
            $('bulkProductsProgress').textContent = '';
            saving = false;
            closeBulkProductsModal();
            ProCable.showToast(cacheSaved ? `تم حفظ ${categoriesToAdd?.length || 0} تصنيف جديد و${added} منتج${updated ? ` وتحديث ${updated}` : ''}${skipped ? ` وتخطي ${skipped} مكرر` : ''} بنجاح` : 'تم الحفظ على الموقع. أعد تحميل الصفحة لتحديث نسخة المتصفح.', cacheSaved ? 'success' : 'info');
        } catch (error) {
            $('bulkProductsProgress').textContent = '';
            message(error.message || 'تعذر حفظ المنتجات. لم يتم تعديل قائمة المنتجات.', 'error');
        } finally {
            saving = false;
            renderPreview();
        }
    }

    Object.assign(window, {
        openBulkProductsModal, closeBulkProductsModal, downloadBulkProductsTemplate,
        handleBulkProductsSheet, handleBulkProductsImages, previewBulkProducts,
        renderBulkProductsPreview,
        updateBulkProductField, updateBulkProductImages, saveBulkProductsDraft,
        loadBulkProductsDraft, removeBulkProductRow, saveBulkProducts
    });
})();
