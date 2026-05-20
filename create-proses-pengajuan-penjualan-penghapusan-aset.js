/* =============================================================================
 * CREATE FORM - PROSES PENGAJUAN PENJUALAN PENGHAPUSAN ASET (Liferay)
 * Migrasi dari versi SharePoint NewForm.aspx -> Liferay Object Form (Create)
 * =============================================================================
 *
 * Logic:
 *   1. Auto-fill Unit Kerja (Department) ke field Title → readonly
 *   2. Auto-fill Tanggal hari ini → readonly
 *   3. Hide semua field conditional saat load
 *   4. Watch perubahan Permohonan Tindaklanjut → show/hide field sesuai case
 *   5. PreSaveAction → validasi attachment wajib
 * ============================================================================= */

$(document).ready(function () {

    // ─── 1. AUTO-FILL UNIT KERJA & TANGGAL ──────────────────────────
    fetch('/o/headless-admin-user/v1.0/my-user-account', {
        credentials: 'include',
        headers: {
            'Accept': 'application/json',
            'X-CSRF-Token': Liferay.authToken
        }
    })
    .then(r => r.json())
    .then(data => {
        // --- Auto-fill Department ke Title ---
        const deptField = data.customFields?.find(f => f.name === 'department');
        const dept = deptField?.customValue?.data || '';

        if (dept) {
            const $titleInput = $('div[data-field-name="ObjectField_title"] input[type="text"]');
            $titleInput.val(dept);
            $titleInput[0].dispatchEvent(new Event('input', { bubbles: true }));
            $titleInput[0].dispatchEvent(new Event('change', { bubbles: true }));
            $titleInput.prop('readonly', true);
            $titleInput.css('background-color', '#e9ecef');
        }

        // --- Auto-fill Tanggal hari ini ---
        const d = new Date();
        const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
                        'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
        const tanggalStr = d.getDate() + ' ' + months[d.getMonth()] + ' ' + d.getFullYear();

        const $tanggalInput = $('div[data-field-name="ObjectField_tanggal"] input[type="text"]');
        if ($tanggalInput.length) {
            $tanggalInput.val(tanggalStr);
            $tanggalInput[0].dispatchEvent(new Event('input', { bubbles: true }));
            $tanggalInput[0].dispatchEvent(new Event('change', { bubbles: true }));
            $tanggalInput.prop('readonly', true);
            $tanggalInput.css('background-color', '#e9ecef');
        }
    })
    .catch(err => console.error('[CREATE] Error fetch user:', err));


    // ─── 2. HIDE SEMUA FIELD CONDITIONAL SAAT LOAD ──────────────────
    // Field approval (tidak ditampilkan di create form)
    hideField('ObjectField_approvalReviewerUn');
    hideField('ObjectField_keteranganApprovalReviewerUn');
    hideField('ObjectField_approvalPejabatUni');
    hideField('ObjectField_keteranganApprovalPejabatUni');
    hideField('ObjectField_approvalPICBOP');
    hideField('ObjectField_keteranganPICBOP');
    hideField('ObjectField_approvalReviewerBO');
    hideField('ObjectField_keteranganApprovalReviewerBO');
    hideField('ObjectField_approvalRFA1');
    hideField('ObjectField_keteranganApprovalRFA1');
    hideField('ObjectField_approvalRFA2');
    hideField('ObjectField_keteranganApprovalRFA2');
    hideField('ObjectField_approvalRFA3');
    hideField('ObjectField_keteranganApprovalRFA3');

    // Field status & tindak lanjut (tidak ditampilkan di create form)
    hideField('ObjectField_statusRequest');
    hideField('ObjectField_statusTindaklanjut');
    hideField('ObjectField_dokumenTindaklanjut');
    hideField('ObjectField_keteranganRequestDokumenTL');
    hideField('ObjectField_statusSubmitDokume');
    hideField('ObjectField_keteranganSubmitDokumenTL');
    hideField('ObjectField_nomorTransaksi');
    hideField('ObjectField_tanggalEfektif');
    hideField('ObjectField_tanggalPenyelesaian');
    hideField('ObjectField_jumlahHariLibur');
    hideField('ObjectField_sla');
    hideField('ObjectField_statusRequestPenjualanAsetTetap');
    hideField('ObjectField_nomorRequestLelangBOP');

    // Field data (conditional per tindaklanjut)
    hideField('ObjectField_jenisTransaksi');
    hideField('ObjectField_keteranganRequester');
    hideField('ObjectField_reviewerUnitKerja');
    hideField('ObjectField_pejabatUnitKerjaX');
    hideField('ObjectField_nomorRequest');
    hideField('ObjectField_isiInisialPIC');
    hideField('ObjectField_nilaiBukuTertinggi');
    hideField('ObjectField_hargaPenawaran');
    hideField('ObjectField_hargaPenawaran0');
    hideField('ObjectField_hargaPenawaran1');
    hideField('ObjectField_hargaPenawaran2');
    hideField('ObjectField_hargaPenawaran3');
    hideField('ObjectField_hargaPenawaran4');
    hideField('ObjectField_finalApproval');
    hideField('ObjectField_reviewerBOP');
    hideField('ObjectField_rFA1');
    hideField('ObjectField_rFA2');
    hideField('ObjectField_rFA3');


    // ─── 3. WATCH PERUBAHAN PERMOHONAN TINDAKLANJUT ─────────────────
    // Di SharePoint: $("select[title='Permohonan Tindaklanjut']").change(...)
    // Di Liferay Object Form: field picklist bisa berupa:
    //   a) <select> biasa
    //   b) hidden <input> yang di-update internal React
    //   c) custom dropdown (div + spans)
    // Kita pakai polling yang cek SEMUA kemungkinan value source
    watchPermohonanTindaklanjut();


    // ─── 4. PRE-SAVE VALIDATION (pengganti SharePoint PreSaveAction) ─
    // Di SharePoint: cek attachment wajib sebelum submit
    registerCreatePreSave();
});


/* =============================================================================
 * WATCH PERMOHONAN TINDAKLANJUT
 * Menggunakan beberapa strategi untuk detect value change di Liferay:
 *   1. Bind change/input event pada select/input
 *   2. Polling value dari semua possible sources
 *   3. MutationObserver pada container field
 * ============================================================================= */
function watchPermohonanTindaklanjut() {
    let lastVal = '';

    function getPermohonanValue() {
        const $container = $('div[data-field-name="ObjectField_permohonanTindaklanjut"]');
        if (!$container.length) return '';

        // Cek 1: <select> element
        const $select = $container.find('select');
        if ($select.length && $select.val()) return $select.val();

        // Cek 2: hidden input (Liferay sering pakai input[type="hidden"] untuk value picklist)
        const $hidden = $container.find('input[type="hidden"]');
        if ($hidden.length && $hidden.val()) return $hidden.val();

        // Cek 3: input[type="text"] (kadang value key disimpan di sini)
        const $text = $container.find('input[type="text"]');
        if ($text.length && $text.val()) return $text.val();

        // Cek 4: data attribute pada container atau child
        const dataVal = $container.attr('data-value') || $container.find('[data-value]').attr('data-value');
        if (dataVal) return dataVal;

        // Cek 5: selected option text → convert ke key
        const $selectedText = $container.find('.dropdown-item.active, .list-group-item.active, option:selected, .selected-label');
        if ($selectedText.length) {
            return mapDisplayTextToKey($selectedText.text().trim());
        }

        return '';
    }

    // Bind native events
    const $container = $('div[data-field-name="ObjectField_permohonanTindaklanjut"]');
    $container.on('change input', 'select, input', function () {
        const val = getPermohonanValue();
        if (val && val !== lastVal) {
            lastVal = val;
            console.log('[CREATE] permohonanTindaklanjut changed (event):', val);
            handlePermohonanChange(val);
        }
    });

    // MutationObserver: detect DOM changes inside the field container
    if ($container.length) {
        const observer = new MutationObserver(function () {
            const val = getPermohonanValue();
            if (val && val !== lastVal) {
                lastVal = val;
                console.log('[CREATE] permohonanTindaklanjut changed (mutation):', val);
                handlePermohonanChange(val);
            }
        });
        observer.observe($container[0], {
            childList: true,
            subtree: true,
            attributes: true,
            attributeFilter: ['value', 'data-value', 'class']
        });
    }

    // Polling fallback (setiap 500ms)
    setInterval(function () {
        const val = getPermohonanValue();
        if (val && val !== lastVal) {
            lastVal = val;
            console.log('[CREATE] permohonanTindaklanjut changed (poll):', val);
            handlePermohonanChange(val);
        }
    }, 500);
}

/**
 * Mapping display text (label) ke key value.
 * Kadang Liferay menyimpan display text, bukan key.
 * Tambahkan mapping sesuai Option List Anda.
 */
function mapDisplayTextToKey(text) {
    const map = {
        'Limit DLOG (nilai buku per unit <50jt)': 'LimitDLOGNilaiBukuPerUnit50jt',
        'Limit Direksi/ Mutasi UKKP /Penghapusan Ekskom (hilang)': 'LimitDireksiMutasiUKKPPenghapusanEkskomHilang',
        'Lelang Gudang BOP': 'LelangGudangBOP',
        'Penjualan Aset Tetap - Hasil Review Pengembalian Barang (Khusus UKKP)': 'PenjualanAsetTetapHasilReviewPengembalianBarangKhususUKKP'
    };
    return map[text] || text;
}


/* =============================================================================
 * HANDLER PERMOHONAN TINDAKLANJUT CHANGE
 * ============================================================================= */
function handlePermohonanChange(val) {

    // Reset semua ke hide dulu
    hideField('ObjectField_jenisTransaksi');
    hideField('ObjectField_keteranganRequester');
    hideField('ObjectField_reviewerUnitKerja');
    hideField('ObjectField_pejabatUnitKerjaX');
    hideField('ObjectField_nomorRequest');
    hideField('ObjectField_isiInisialPIC');
    hideField('ObjectField_nilaiBukuTertinggi');
    hideField('ObjectField_hargaPenawaran');
    hideField('ObjectField_hargaPenawaran0');
    hideField('ObjectField_hargaPenawaran1');
    hideField('ObjectField_hargaPenawaran2');
    hideField('ObjectField_hargaPenawaran3');
    hideField('ObjectField_hargaPenawaran4');
    hideField('ObjectField_finalApproval');
    hideField('ObjectField_reviewerBOP');
    hideField('ObjectField_rFA1');
    hideField('ObjectField_rFA2');
    hideField('ObjectField_rFA3');

    switch (val) {

        // ─── Limit DLOG (nilai buku per unit <50jt) ─────────────────
        case 'LimitDLOGNilaiBukuPerUnit50jt':
            showField('ObjectField_keteranganRequester');
            showField('ObjectField_jenisTransaksi');
            showField('ObjectField_reviewerUnitKerja');
            showField('ObjectField_pejabatUnitKerjaX');
            break;

        // ─── Limit Direksi / Mutasi UKKP / Penghapusan Ekskom ───────
        case 'LimitDireksiMutasiUKKPPenghapusanEkskomHilang':
            showField('ObjectField_keteranganRequester');
            showField('ObjectField_jenisTransaksi');
            // Reviewer & Pejabat Unit Kerja TIDAK ditampilkan
            break;

        // ─── Lelang Gudang BOP ──────────────────────────────────────
        case 'LelangGudangBOP':
            showField('ObjectField_keteranganRequester');
            showField('ObjectField_nomorRequest');
            showField('ObjectField_isiInisialPIC');
            showField('ObjectField_nilaiBukuTertinggi');
            showField('ObjectField_hargaPenawaran');
            showField('ObjectField_hargaPenawaran0');
            showField('ObjectField_hargaPenawaran1');
            showField('ObjectField_hargaPenawaran2');
            showField('ObjectField_hargaPenawaran3');
            showField('ObjectField_hargaPenawaran4');
            showField('ObjectField_finalApproval');
            showField('ObjectField_reviewerBOP');
            showField('ObjectField_rFA1');
            showField('ObjectField_rFA2');
            showField('ObjectField_rFA3');
            // Jenis Transaksi, Reviewer & Pejabat Unit Kerja TIDAK ditampilkan
            break;

        // ─── Penjualan Aset Tetap - Hasil Review Pengembalian Barang (Khusus UKKP) ─
        case 'PenjualanAsetTetapHasilReviewPengembalianBarangKhususUKKP':
            showField('ObjectField_keteranganRequester');
            showField('ObjectField_jenisTransaksi');
            showField('ObjectField_reviewerUnitKerja');
            showField('ObjectField_pejabatUnitKerjaX');
            // Semua field BOP TIDAK ditampilkan (sama seperti Limit DLOG)
            break;
    }
}


/* =============================================================================
 * PRE-SAVE VALIDATION
 * SharePoint: cek attachment wajib via idAttachmentsRow
 * Liferay: cek apakah ada file attachment yang diupload
 * ============================================================================= */
function registerCreatePreSave() {
    $(document).off('click.wfCreatePreSave').on('click.wfCreatePreSave',
        '.bca-form-container button[type="submit"], .lfr-ddm__form-submit, .button-submit',
        function (e) {
            if (!preSaveCreateValidate()) {
                e.preventDefault();
                e.stopImmediatePropagation();
                return false;
            }
        });
}

function preSaveCreateValidate() {
    // Cek apakah ada attachment yang diupload
    // Di Liferay, attachment biasanya ada di dalam container tertentu
    // Sesuaikan selector dengan implementasi form Anda
    const $attachmentContainer = $('.bca-form-container .attachment-list, .bca-form-container [data-field-name*="attachment"], .bca-form-container .document-library');
    
    // Cek apakah ada file yang sudah diupload
    const hasAttachment = $attachmentContainer.find('.file-entry, .list-group-item, img, a[href]').length > 0;
    
    // Alternatif: cek via input file
    const $fileInputs = $('.bca-form-container input[type="file"]');
    const hasFileSelected = $fileInputs.filter(function() {
        return this.files && this.files.length > 0;
    }).length > 0;

    if (!hasAttachment && !hasFileSelected) {
        // Cek juga apakah memang ada field attachment di form
        // Jika tidak ada field attachment sama sekali, skip validasi
        if ($attachmentContainer.length > 0 || $fileInputs.length > 0) {
            alert('Silahkan pilih attachment');
            return false;
        }
    }

    return true;
}


/* =============================================================================
 * HELPER FUNCTIONS
 * ============================================================================= */
function hideField(fieldName) {
    $('div[data-field-name="' + fieldName + '"]').hide();
}

function showField(fieldName) {
    $('div[data-field-name="' + fieldName + '"]').show();
}
