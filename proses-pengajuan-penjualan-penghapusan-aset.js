/* =============================================================================
 * PROSES PENGAJUAN PENJUALAN PENGHAPUSAN ASET (Liferay)
 * Migrasi dari versi SharePoint -> Liferay Object Form
 * =============================================================================
 *
 * Catatan field naming:
 *   - Field di Liferay diakses via: div[data-field-name="ObjectField_xxx"]
 *   - Picker user (reviewer/pejabat/RFA) di Liferay disimpan sebagai object
 *     dengan property: { key: "<userId/email>", name: "<displayName>" }
 *
 * Catatan role / group:
 *   - "PIC BOP - ProsesPengajuanPenjualanPenghapusan"
 *   - "Request Tindak Lanjut Aset Tetap - Open All Fields"
 *   Keduanya dicek lewat userData.roleBriefs[].name
 * ============================================================================= */

setTimeout(function () {
    const pathParts = window.location.pathname.split('/');
    const entryERC = pathParts[pathParts.length - 1];
    const scopeGroupId = Liferay.ThemeDisplay.getScopeGroupId();

    Promise.all([
        fetch('/o/c/prosespengajuanpenjualanpenghapusanasets/scopes/' + scopeGroupId +
              '/by-external-reference-code/' + entryERC + '?p_auth=' + Liferay.authToken, {
            credentials: 'include'
        }).then(r => r.json()),
        fetch('/o/headless-admin-user/v1.0/my-user-account', {
            credentials: 'include',
            headers: { 'Accept': 'application/json', 'X-CSRF-Token': Liferay.authToken }
        }).then(r => r.json())
    ])
    .then(([data, userData]) => {
        const ctx = buildContext(data, userData);
        console.log('[WF] context:', ctx);
        handleWorkflow(ctx);
        registerPreSaveAction(ctx);
    })
    .catch(err => console.error('[WF] Error:', err));
}, 500);


/* =============================================================================
 * BUILD CONTEXT - kumpulkan semua variabel yang dibutuhkan
 * ============================================================================= */
function buildContext(data, userData) {
    const userEmail = (userData.emailAddress || '').toLowerCase();
    const userName  = userData.name || userData.alternateName || '';
    const userRoles = (userData.roleBriefs || []).map(r => r.name);

    return {
        // raw
        data: data,
        userData: userData,

        // user info
        userEmail: userEmail,
        userName: userName,
        userRoles: userRoles,
        isPICBOP: userRoles.includes('PIC BOP - ProsesPengajuanPenjualanPenghapusan'),
        isOpenAllFields: userRoles.includes('Request Tindak Lanjut Aset Tetap - Open All Fields'),

        // workflow keys
        tindaklanjut:        data.permohonanTindaklanjut?.key || '',
        statusRequest:       data.statusRequest || '',
        jenisTransaksi:      data.jenisTransaksi?.key || '',

        // approval keys
        approvalReviewerUn:  data.approvalReviewerUn?.key  || '',
        approvalPejabatUni:  data.approvalPejabatUni?.key  || '',
        approvalPICBOP:      data.approvalPICBOP?.key      || '',
        approvalReviewerBO:  data.approvalReviewerBO?.key  || '',
        approvalRFA1:        data.approvalRFA1?.key        || '',
        approvalRFA2:        data.approvalRFA2?.key        || '',
        approvalRFA3:        data.approvalRFA3?.key        || '',

        // picker user (object {key,name})
        reviewerUnitKerja:   data.reviewerUnitKerja  || null,
        pejabatUnitKerjaX:   data.pejabatUnitKerjaX  || null,
        reviewerBOP:         data.reviewerBOP        || null,
        rFA1:                data.rFA1               || null,
        rFA2:                data.rFA2               || null,
        rFA3:                data.rFA3               || null,

        // status tindak lanjut & dokumen
        statusTindaklanjut:  data.statusTindaklanjut?.key || '',
        statusSubmitDokume:  data.statusSubmitDokume?.key || '',

        // Penjualan Aset Tetap UKKP only
        statusReqPenjualan:  data.statusRequestPenjualanAsetTetap?.key || '',
        nomorRequestLelang:  data.nomorRequestLelangBOP || ''
    };
}


/* =============================================================================
 * MAIN HANDLER
 * ============================================================================= */
function handleWorkflow(ctx) {
    // 1. Lock semua field approval & conditional
    lockAllFields();

    // 2. Hide semua field conditional
    hideAllConditionalFields();

    // 3. Field utama requester → SELALU unlocked & visible agar bisa mengisi data awal
    //    (Di SharePoint: Permohonan Tindaklanjut, Title, Unit Kerja Pemohon
    //     hanya di-disable setelah ada approval/reject)
    unlockField('ObjectField_permohonanTindaklanjut');
    unlockField('ObjectField_title');
    unlockField('ObjectField_keteranganRequester');

    // 4. Auto-enable picker yang masih kosong (mirip SharePoint:
    //    "if revbop.TotalUserCount == 0 enable picker")
    if (!ctx.reviewerBOP) unlockField('ObjectField_reviewerBOP');
    if (!ctx.rFA1)        unlockField('ObjectField_rFA1');
    if (!ctx.rFA2)        unlockField('ObjectField_rFA2');
    if (!ctx.rFA3)        unlockField('ObjectField_rFA3');

    // 5. Cek kondisi Reject/Completed -> disable semua, stop
    const isLocked = (ctx.approvalReviewerUn === 'Reject') ||
                     (ctx.approvalPejabatUni === 'Reject') ||
                     (ctx.approvalPICBOP     === 'Reject') ||
                     (ctx.approvalPICBOP     === 'Completed') ||
                     (ctx.approvalReviewerBO === 'Reject') ||
                     (ctx.approvalRFA1       === 'Reject') ||
                     (ctx.approvalRFA2       === 'Reject') ||
                     (ctx.approvalRFA3       === 'Reject');
    if (isLocked) {
        disableAll();
        return;
    }

    // 6. Jika sudah ada approval berjalan (reviewer sudah diisi dan statusnya bukan kosong),
    //    kunci field requester agar tidak bisa diubah lagi
    if (ctx.approvalReviewerUn && ctx.approvalReviewerUn !== '' &&
        ctx.approvalReviewerUn !== 'Correction') {
        lockField('ObjectField_permohonanTindaklanjut');
        lockField('ObjectField_title');
        lockField('ObjectField_keteranganRequester');
    }

    // 7. Routing per tindaklanjut
    switch (ctx.tindaklanjut) {
        case 'LimitDLOGNilaiBukuPerUnit50jt':
            handleLimitDLOG(ctx);
            break;

        case 'LimitDireksiMutasiUKKPPenghapusanEkskomHilang':
            handleLimitDireksi(ctx);
            break;

        case 'LelangGudangBOP':
            handleLelangGudangBOP(ctx);
            break;

        case 'PenjualanAsetTetapHasilReviewPengembalianBarangKhususUKKP':
            handlePenjualanAsetTetapUKKP(ctx);
            break;
    }

    // 8. Jika tindaklanjut belum dipilih → requester masih bisa isi data awal
    //    Buka juga reviewer unit kerja & pejabat agar bisa assign orang
    if (!ctx.tindaklanjut) {
        unlockField('ObjectField_permohonanTindaklanjut');
        unlockField('ObjectField_title');
        unlockField('ObjectField_keteranganRequester');
        showField('ObjectField_reviewerUnitKerja');
        showField('ObjectField_pejabatUnitKerjaX');
        unlockField('ObjectField_reviewerUnitKerja');
        unlockField('ObjectField_pejabatUnitKerjaX');
    }

    // 9. Status Tindaklanjut & Dokumen TL (berlaku across workflow)
    handleStatusDokumenTL(ctx);

    // 10. Group "Open All Fields" - override semua jika user punya role ini
    if (ctx.isOpenAllFields && ctx.statusRequest === 'Completed') {
        unlockAllMainFields();
    }
}


/* =============================================================================
 * 1. LIMIT DLOG (nilai buku per unit < 50jt)
 * ============================================================================= */
function handleLimitDLOG(ctx) {
    // Show semua field yang relevan
    showField('ObjectField_jenisTransaksi');
    showField('ObjectField_reviewerUnitKerja');
    showField('ObjectField_pejabatUnitKerjaX');
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

    // -------------------------------------------------------------------------
    // Reviewer Unit Kerja - Waiting Review
    // Hanya user yang ditunjuk sebagai reviewer yang bisa approve
    // -------------------------------------------------------------------------
    if (ctx.reviewerUnitKerja && ctx.approvalReviewerUn === 'WaitingReview' &&
        isCurrentUser(ctx.reviewerUnitKerja, ctx)) {
        unlockField('ObjectField_approvalReviewerUn');
        unlockField('ObjectField_keteranganApprovalReviewerUn');
    }

    // Reviewer Correction -> reset value & buka data fields
    if (ctx.reviewerUnitKerja && ctx.approvalReviewerUn === 'Correction') {
        resetApprovalToWaitingReview('ObjectField_approvalReviewerUn');

        unlockField('ObjectField_jenisTransaksi');
        unlockField('ObjectField_keteranganRequester');
        unlockField('ObjectField_reviewerUnitKerja');
        unlockField('ObjectField_pejabatUnitKerjaX');
    }

    // -------------------------------------------------------------------------
    // Pejabat Unit Kerja - Waiting Review (setelah reviewer approve)
    // -------------------------------------------------------------------------
    if (ctx.pejabatUnitKerjaX &&
        ctx.approvalPejabatUni === 'WaitingReview' &&
        ctx.approvalReviewerUn === 'Approve' &&
        isCurrentUser(ctx.pejabatUnitKerjaX, ctx)) {
        unlockField('ObjectField_approvalPejabatUni');
        unlockField('ObjectField_keteranganApprovalPejabatUni');
    }

    // Pejabat Correction -> reset value approval & buka data fields
    if (ctx.pejabatUnitKerjaX &&
        ctx.approvalPejabatUni === 'Correction' &&
        ctx.approvalReviewerUn === 'Approve') {
        resetApprovalToWaitingReview('ObjectField_approvalPejabatUni');
        resetApprovalToWaitingReview('ObjectField_approvalReviewerUn');

        unlockField('ObjectField_jenisTransaksi');
        unlockField('ObjectField_keteranganRequester');
        unlockField('ObjectField_reviewerUnitKerja');
        unlockField('ObjectField_pejabatUnitKerjaX');
    }

    // -------------------------------------------------------------------------
    // PIC BOP - buka semua data fields (setelah pejabat approve)
    // Di SharePoint: tidak ada cek role, semua orang bisa lihat & edit
    // selama approvalPICBOP kosong/OnProgress dan pejabat sudah approve
    // -------------------------------------------------------------------------
    if ((ctx.approvalPICBOP === '' || ctx.approvalPICBOP === 'OnProgress') &&
        ctx.approvalPejabatUni === 'Approve') {
        unlockField('ObjectField_nomorRequest');
        unlockField('ObjectField_isiInisialPIC');
        unlockField('ObjectField_approvalPICBOP');
        unlockField('ObjectField_keteranganPICBOP');
        unlockField('ObjectField_nilaiBukuTertinggi');
        unlockField('ObjectField_hargaPenawaran');
        unlockField('ObjectField_hargaPenawaran0');
        unlockField('ObjectField_hargaPenawaran1');
        unlockField('ObjectField_hargaPenawaran2');
        unlockField('ObjectField_hargaPenawaran3');
        unlockField('ObjectField_hargaPenawaran4');
        unlockField('ObjectField_finalApproval');
        unlockField('ObjectField_reviewerBOP');
        unlockField('ObjectField_rFA1');
        unlockField('ObjectField_rFA2');
        unlockField('ObjectField_rFA3');
    }

    // -------------------------------------------------------------------------
    // Reviewer BOP - Waiting Review
    // -------------------------------------------------------------------------
    if (ctx.approvalPICBOP === 'WaitingApproval' &&
        ctx.reviewerBOP &&
        ctx.approvalReviewerBO === 'WaitingReview' &&
        isCurrentUser(ctx.reviewerBOP, ctx)) {
        unlockField('ObjectField_approvalReviewerBO');
        unlockField('ObjectField_keteranganApprovalReviewerBO');
    }

    // Reviewer BOP Correction -> reset & buka PIC BOP fields
    if (ctx.reviewerBOP && ctx.approvalReviewerBO === 'Correction') {
        resetApprovalToWaitingReview('ObjectField_approvalReviewerBO');

        unlockField('ObjectField_nomorRequest');
        unlockField('ObjectField_isiInisialPIC');
        unlockField('ObjectField_approvalPICBOP');
        unlockField('ObjectField_keteranganPICBOP');
        unlockField('ObjectField_nilaiBukuTertinggi');
        unlockField('ObjectField_hargaPenawaran');
        unlockField('ObjectField_hargaPenawaran0');
        unlockField('ObjectField_hargaPenawaran1');
        unlockField('ObjectField_hargaPenawaran2');
        unlockField('ObjectField_hargaPenawaran3');
        unlockField('ObjectField_hargaPenawaran4');
        unlockField('ObjectField_finalApproval');
        unlockField('ObjectField_reviewerBOP');
        unlockField('ObjectField_rFA1');
        unlockField('ObjectField_rFA2');
        unlockField('ObjectField_rFA3');
    }

    // -------------------------------------------------------------------------
    // RFA 1 - Waiting Review (setelah Reviewer BOP approve)
    // -------------------------------------------------------------------------
    if (ctx.rFA1 && ctx.approvalRFA1 === 'WaitingReview' &&
        ctx.approvalReviewerBO === 'Approve' &&
        isCurrentUser(ctx.rFA1, ctx)) {
        unlockField('ObjectField_approvalRFA1');
        unlockField('ObjectField_keteranganApprovalRFA1');
    }

    // RFA 1 Correction -> reset semua approval di atasnya
    if (ctx.rFA1 && ctx.approvalRFA1 === 'Correction' &&
        ctx.approvalReviewerBO === 'Approve') {
        resetApprovalToWaitingReview('ObjectField_approvalRFA1');
        resetApprovalToWaitingReview('ObjectField_approvalReviewerBO');

        unlockField('ObjectField_nomorRequest');
        unlockField('ObjectField_isiInisialPIC');
        unlockField('ObjectField_approvalPICBOP');
        unlockField('ObjectField_keteranganPICBOP');
        unlockField('ObjectField_nilaiBukuTertinggi');
        unlockField('ObjectField_hargaPenawaran');
        unlockField('ObjectField_hargaPenawaran0');
        unlockField('ObjectField_hargaPenawaran1');
        unlockField('ObjectField_hargaPenawaran2');
        unlockField('ObjectField_hargaPenawaran3');
        unlockField('ObjectField_hargaPenawaran4');
        unlockField('ObjectField_finalApproval');
        unlockField('ObjectField_reviewerBOP');
        unlockField('ObjectField_rFA1');
        unlockField('ObjectField_rFA2');
        unlockField('ObjectField_rFA3');
    }

    // -------------------------------------------------------------------------
    // RFA 2 - Waiting Review (setelah RFA 1 approve)
    // -------------------------------------------------------------------------
    if (ctx.rFA2 && ctx.approvalRFA2 === 'WaitingReview' &&
        ctx.approvalRFA1 === 'Approve' &&
        isCurrentUser(ctx.rFA2, ctx)) {
        unlockField('ObjectField_approvalRFA2');
        unlockField('ObjectField_keteranganApprovalRFA2');
    }

    // RFA 2 Correction -> reset semua approval di atasnya
    if (ctx.rFA2 && ctx.approvalRFA2 === 'Correction' &&
        ctx.approvalRFA1 === 'Approve') {
        resetApprovalToWaitingReview('ObjectField_approvalRFA2');
        resetApprovalToWaitingReview('ObjectField_approvalRFA1');
        resetApprovalToWaitingReview('ObjectField_approvalReviewerBO');

        unlockField('ObjectField_nomorRequest');
        unlockField('ObjectField_isiInisialPIC');
        unlockField('ObjectField_approvalPICBOP');
        unlockField('ObjectField_keteranganPICBOP');
        unlockField('ObjectField_nilaiBukuTertinggi');
        unlockField('ObjectField_hargaPenawaran');
        unlockField('ObjectField_hargaPenawaran0');
        unlockField('ObjectField_hargaPenawaran1');
        unlockField('ObjectField_hargaPenawaran2');
        unlockField('ObjectField_hargaPenawaran3');
        unlockField('ObjectField_hargaPenawaran4');
        unlockField('ObjectField_finalApproval');
        unlockField('ObjectField_reviewerBOP');
        unlockField('ObjectField_rFA1');
        unlockField('ObjectField_rFA2');
        unlockField('ObjectField_rFA3');
    }

    // -------------------------------------------------------------------------
    // RFA 3 - Waiting Review (setelah RFA 2 approve)
    // -------------------------------------------------------------------------
    if (ctx.rFA3 && ctx.approvalRFA3 === 'WaitingReview' &&
        ctx.approvalRFA2 === 'Approve' &&
        isCurrentUser(ctx.rFA3, ctx)) {
        unlockField('ObjectField_approvalRFA3');
        unlockField('ObjectField_keteranganApprovalRFA3');
    }

    // RFA 3 Correction -> reset semua approval di atasnya
    if (ctx.rFA3 && ctx.approvalRFA3 === 'Correction' &&
        ctx.approvalRFA2 === 'Approve') {
        resetApprovalToWaitingReview('ObjectField_approvalRFA3');
        resetApprovalToWaitingReview('ObjectField_approvalRFA2');
        resetApprovalToWaitingReview('ObjectField_approvalRFA1');
        resetApprovalToWaitingReview('ObjectField_approvalReviewerBO');

        unlockField('ObjectField_nomorRequest');
        unlockField('ObjectField_isiInisialPIC');
        unlockField('ObjectField_approvalPICBOP');
        unlockField('ObjectField_keteranganPICBOP');
        unlockField('ObjectField_nilaiBukuTertinggi');
        unlockField('ObjectField_hargaPenawaran');
        unlockField('ObjectField_hargaPenawaran0');
        unlockField('ObjectField_hargaPenawaran1');
        unlockField('ObjectField_hargaPenawaran2');
        unlockField('ObjectField_hargaPenawaran3');
        unlockField('ObjectField_hargaPenawaran4');
        unlockField('ObjectField_finalApproval');
        unlockField('ObjectField_reviewerBOP');
        unlockField('ObjectField_rFA1');
        unlockField('ObjectField_rFA2');
        unlockField('ObjectField_rFA3');
    }

    // -------------------------------------------------------------------------
    // Auto-enable Final Approval & RFA 1 jika field-nya kosong
    // -------------------------------------------------------------------------
    if (!getFieldValue('ObjectField_finalApproval')) {
        unlockField('ObjectField_finalApproval');
    }
    if (!ctx.rFA1) {
        unlockField('ObjectField_rFA1');
    }
}


/* =============================================================================
 * 2. LIMIT DIREKSI / MUTASI UKKP / PENGHAPUSAN EKSKOM (HILANG)
 * ============================================================================= */
function handleLimitDireksi(ctx) {
    showField('ObjectField_jenisTransaksi');
    showField('ObjectField_nomorRequest');
    showField('ObjectField_isiInisialPIC');

    if (ctx.approvalPICBOP !== 'Reject' && ctx.isPICBOP) {
        unlockField('ObjectField_nomorRequest');
        unlockField('ObjectField_isiInisialPIC');
        unlockField('ObjectField_approvalPICBOP');
        unlockField('ObjectField_keteranganPICBOP');
        unlockField('ObjectField_nomorTransaksi');
        unlockField('ObjectField_tanggalEfektif');
        unlockField('ObjectField_tanggalPenyelesaian');
        unlockField('ObjectField_jumlahHariLibur');
        unlockField('ObjectField_sla');
    }
}


/* =============================================================================
 * 3. LELANG GUDANG BOP
 * ============================================================================= */
function handleLelangGudangBOP(ctx) {
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

    // -------------------------------------------------------------------------
    // Reviewer BOP - Waiting Review
    // -------------------------------------------------------------------------
    if (ctx.reviewerBOP && ctx.approvalReviewerBO === 'WaitingReview' &&
        isCurrentUser(ctx.reviewerBOP, ctx)) {
        unlockField('ObjectField_approvalReviewerBO');
        unlockField('ObjectField_keteranganApprovalReviewerBO');
    }

    // Reviewer BOP Correction
    if (ctx.reviewerBOP && ctx.approvalReviewerBO === 'Correction') {
        resetApprovalToWaitingReview('ObjectField_approvalReviewerBO');

        unlockField('ObjectField_approvalPICBOP');
        unlockField('ObjectField_keteranganRequester');
        unlockField('ObjectField_nomorRequest');
        unlockField('ObjectField_isiInisialPIC');
        unlockField('ObjectField_nilaiBukuTertinggi');
        unlockField('ObjectField_hargaPenawaran');
        unlockField('ObjectField_hargaPenawaran0');
        unlockField('ObjectField_hargaPenawaran1');
        unlockField('ObjectField_hargaPenawaran2');
        unlockField('ObjectField_hargaPenawaran3');
        unlockField('ObjectField_hargaPenawaran4');
        unlockField('ObjectField_finalApproval');
        unlockField('ObjectField_reviewerBOP');
        unlockField('ObjectField_rFA1');
        unlockField('ObjectField_rFA2');
        unlockField('ObjectField_rFA3');
    }

    // -------------------------------------------------------------------------
    // RFA 1 - Waiting Review (setelah Reviewer BOP approve)
    // -------------------------------------------------------------------------
    if (ctx.rFA1 && ctx.approvalRFA1 === 'WaitingReview' &&
        ctx.approvalReviewerBO === 'Approve' &&
        isCurrentUser(ctx.rFA1, ctx)) {
        unlockField('ObjectField_approvalRFA1');
        unlockField('ObjectField_keteranganApprovalRFA1');
    }

    // RFA 1 Correction
    if (ctx.rFA1 && ctx.approvalRFA1 === 'Correction' &&
        ctx.approvalReviewerBO === 'Approve') {
        resetApprovalToWaitingReview('ObjectField_approvalRFA1');
        resetApprovalToWaitingReview('ObjectField_approvalReviewerBO');

        unlockField('ObjectField_approvalPICBOP');
        unlockField('ObjectField_keteranganRequester');
        unlockField('ObjectField_nomorRequest');
        unlockField('ObjectField_isiInisialPIC');
        unlockField('ObjectField_nilaiBukuTertinggi');
        unlockField('ObjectField_hargaPenawaran');
        unlockField('ObjectField_hargaPenawaran0');
        unlockField('ObjectField_hargaPenawaran1');
        unlockField('ObjectField_hargaPenawaran2');
        unlockField('ObjectField_hargaPenawaran3');
        unlockField('ObjectField_hargaPenawaran4');
        unlockField('ObjectField_finalApproval');
        unlockField('ObjectField_reviewerBOP');
        unlockField('ObjectField_rFA1');
        unlockField('ObjectField_rFA2');
        unlockField('ObjectField_rFA3');
    }

    // -------------------------------------------------------------------------
    // RFA 2 - Waiting Review (setelah RFA 1 approve)
    // -------------------------------------------------------------------------
    if (ctx.rFA2 && ctx.approvalRFA2 === 'WaitingReview' &&
        ctx.approvalRFA1 === 'Approve' &&
        isCurrentUser(ctx.rFA2, ctx)) {
        unlockField('ObjectField_approvalRFA2');
        unlockField('ObjectField_keteranganApprovalRFA2');
    }

    // RFA 2 Correction
    if (ctx.rFA2 && ctx.approvalRFA2 === 'Correction' &&
        ctx.approvalRFA1 === 'Approve') {
        resetApprovalToWaitingReview('ObjectField_approvalRFA2');
        resetApprovalToWaitingReview('ObjectField_approvalRFA1');
        resetApprovalToWaitingReview('ObjectField_approvalReviewerBO');

        unlockField('ObjectField_approvalPICBOP');
        unlockField('ObjectField_keteranganRequester');
        unlockField('ObjectField_nomorRequest');
        unlockField('ObjectField_isiInisialPIC');
        unlockField('ObjectField_nilaiBukuTertinggi');
        unlockField('ObjectField_hargaPenawaran');
        unlockField('ObjectField_hargaPenawaran0');
        unlockField('ObjectField_hargaPenawaran1');
        unlockField('ObjectField_hargaPenawaran2');
        unlockField('ObjectField_hargaPenawaran3');
        unlockField('ObjectField_hargaPenawaran4');
        unlockField('ObjectField_finalApproval');
        unlockField('ObjectField_reviewerBOP');
        unlockField('ObjectField_rFA1');
        unlockField('ObjectField_rFA2');
        unlockField('ObjectField_rFA3');
    }

    // -------------------------------------------------------------------------
    // RFA 3 - Waiting Review (setelah RFA 2 approve)
    // -------------------------------------------------------------------------
    if (ctx.rFA3 && ctx.approvalRFA3 === 'WaitingReview' &&
        ctx.approvalRFA2 === 'Approve' &&
        isCurrentUser(ctx.rFA3, ctx)) {
        unlockField('ObjectField_approvalRFA3');
        unlockField('ObjectField_keteranganApprovalRFA3');
    }

    // RFA 3 Correction
    if (ctx.rFA3 && ctx.approvalRFA3 === 'Correction' &&
        ctx.approvalRFA2 === 'Approve') {
        resetApprovalToWaitingReview('ObjectField_approvalRFA3');
        resetApprovalToWaitingReview('ObjectField_approvalRFA2');
        resetApprovalToWaitingReview('ObjectField_approvalRFA1');
        resetApprovalToWaitingReview('ObjectField_approvalReviewerBO');

        unlockField('ObjectField_approvalPICBOP');
        unlockField('ObjectField_keteranganRequester');
        unlockField('ObjectField_nomorRequest');
        unlockField('ObjectField_isiInisialPIC');
        unlockField('ObjectField_nilaiBukuTertinggi');
        unlockField('ObjectField_hargaPenawaran');
        unlockField('ObjectField_hargaPenawaran0');
        unlockField('ObjectField_hargaPenawaran1');
        unlockField('ObjectField_hargaPenawaran2');
        unlockField('ObjectField_hargaPenawaran3');
        unlockField('ObjectField_hargaPenawaran4');
        unlockField('ObjectField_finalApproval');
        unlockField('ObjectField_reviewerBOP');
        unlockField('ObjectField_rFA1');
        unlockField('ObjectField_rFA2');
        unlockField('ObjectField_rFA3');
    }

    // -------------------------------------------------------------------------
    // PIC BOP - finalisasi (setelah RFA chain selesai)
    // -------------------------------------------------------------------------
    const rfa1OK = ctx.rFA1 && ctx.approvalRFA1 === 'Approve';
    const rfa2OK = ctx.rFA2 && ctx.approvalRFA2 === 'Approve';
    const rfa3OK = ctx.rFA3 && ctx.approvalRFA3 === 'Approve';
    const noRfa2 = !ctx.rFA2;
    const noRfa3 = !ctx.rFA3;

    const allRfaApprove =
        (rfa1OK && noRfa2 && noRfa3) ||
        (rfa1OK && rfa2OK && noRfa3) ||
        (rfa1OK && rfa2OK && rfa3OK);

    if (allRfaApprove && ctx.isPICBOP) {
        unlockField('ObjectField_approvalPICBOP');
        unlockField('ObjectField_keteranganPICBOP');
        unlockField('ObjectField_nomorTransaksi');
        unlockField('ObjectField_tanggalEfektif');
        unlockField('ObjectField_tanggalPenyelesaian');
        unlockField('ObjectField_jumlahHariLibur');
        unlockField('ObjectField_sla');
    }
}


/* =============================================================================
 * 4. PENJUALAN ASET TETAP - HASIL REVIEW PENGEMBALIAN BARANG (KHUSUS UKKP)
 * ============================================================================= */
function handlePenjualanAsetTetapUKKP(ctx) {
    // Reset semua hide -> show field yang dipakai case ini
    showField('ObjectField_jenisTransaksi');
    showField('ObjectField_reviewerUnitKerja');
    showField('ObjectField_pejabatUnitKerjaX');

    // Sembunyikan field yang TIDAK dipakai untuk case ini
    hideField('ObjectField_finalApproval');
    hideField('ObjectField_nomorRequest');
    hideField('ObjectField_isiInisialPIC');
    hideField('ObjectField_nilaiBukuTertinggi');
    hideField('ObjectField_hargaPenawaran');
    hideField('ObjectField_hargaPenawaran0');
    hideField('ObjectField_hargaPenawaran1');
    hideField('ObjectField_hargaPenawaran2');
    hideField('ObjectField_hargaPenawaran3');
    hideField('ObjectField_hargaPenawaran4');
    hideField('ObjectField_reviewerBOP');
    hideField('ObjectField_rFA1');
    hideField('ObjectField_rFA2');
    hideField('ObjectField_rFA3');
    hideField('ObjectField_statusTindaklanjut');
    hideField('ObjectField_statusRequestPenjualanAsetTetap');
    hideField('ObjectField_nomorRequestLelangBOP');

    // -------------------------------------------------------------------------
    // Reviewer UKKP Waiting Review -> buka Jenis Transaksi & Keterangan Requester
    // -------------------------------------------------------------------------
    if (ctx.approvalReviewerUn === 'WaitingReview') {
        unlockField('ObjectField_jenisTransaksi');
        unlockField('ObjectField_keteranganRequester');
    }

    // Reviewer UKKP - hanya user yang ditunjuk yang bisa approve
    if (ctx.reviewerUnitKerja && isCurrentUser(ctx.reviewerUnitKerja, ctx)) {
        showField('ObjectField_approvalReviewerUn');
        showField('ObjectField_keteranganApprovalReviewerUn');

        unlockField('ObjectField_approvalReviewerUn');
        unlockField('ObjectField_keteranganApprovalReviewerUn');

        // Sudah approve -> kunci kembali
        if (ctx.approvalReviewerUn === 'Approve') {
            lockField('ObjectField_approvalReviewerUn');
            lockField('ObjectField_keteranganApprovalReviewerUn');
        }

        // Tampilkan + buka Pejabat Unit Kerja jika reviewer sudah approve
        showField('ObjectField_approvalPejabatUni');
        showField('ObjectField_keteranganApprovalPejabatUni');
        if (ctx.approvalReviewerUn === 'Approve') {
            unlockField('ObjectField_approvalPejabatUni');
            unlockField('ObjectField_keteranganApprovalPejabatUni');
        }
    }

    // Pejabat sudah Approve -> kunci
    if (ctx.approvalReviewerUn === 'Approve' && ctx.approvalPejabatUni === 'Approve') {
        lockField('ObjectField_approvalPejabatUni');
        lockField('ObjectField_keteranganApprovalPejabatUni');
    }

    // -------------------------------------------------------------------------
    // PIC BOP -> buka Nomor Request, Inisial, Status Request, Nomor Lelang
    // -------------------------------------------------------------------------
    if (ctx.isPICBOP &&
        ctx.approvalReviewerUn === 'Approve' &&
        ctx.approvalPejabatUni === 'Approve') {

        showField('ObjectField_nomorRequest');
        showField('ObjectField_isiInisialPIC');
        showField('ObjectField_keteranganPICBOP');
        showField('ObjectField_statusRequestPenjualanAsetTetap');
        showField('ObjectField_nomorRequestLelangBOP');

        unlockField('ObjectField_nomorRequest');
        unlockField('ObjectField_isiInisialPIC');
        unlockField('ObjectField_keteranganPICBOP');
        unlockField('ObjectField_statusRequestPenjualanAsetTetap');

        // Nomor Request Lelang BOP -> hanya bisa diisi jika status = Completed
        if (ctx.statusReqPenjualan === 'Completed') {
            unlockField('ObjectField_nomorRequestLelangBOP');
        } else {
            lockField('ObjectField_nomorRequestLelangBOP');
        }

        // Listener: ketika user ubah Status Request, sesuaikan field Nomor Lelang
        bindStatusRequestListener();
    }

    // -------------------------------------------------------------------------
    // Status Request Completed/Rejected -> kunci semua
    // -------------------------------------------------------------------------
    if (ctx.statusRequest === 'Completed' || ctx.statusRequest === 'Rejected') {
        lockField('ObjectField_permohonanTindaklanjut');
        lockField('ObjectField_title');
        lockField('ObjectField_jenisTransaksi');
        lockField('ObjectField_keteranganRequester');
        lockField('ObjectField_approvalReviewerUn');
        lockField('ObjectField_keteranganApprovalReviewerUn');
        lockField('ObjectField_approvalPejabatUni');
        lockField('ObjectField_keteranganApprovalPejabatUni');
        lockField('ObjectField_nomorRequest');
        lockField('ObjectField_isiInisialPIC');
        lockField('ObjectField_keteranganPICBOP');
        lockField('ObjectField_statusRequestPenjualanAsetTetap');
        lockField('ObjectField_nomorRequestLelangBOP');
    }
}

// Listener: ubah perilaku field Nomor Request Lelang BOP berdasarkan Status Request
function bindStatusRequestListener() {
    const $field = $('div[data-field-name="ObjectField_statusRequestPenjualanAsetTetap"]');
    const $select = $field.find('select');
    if (!$select.length) return;

    $select.off('change.wfLelang').on('change.wfLelang', function () {
        const $lelang = $('div[data-field-name="ObjectField_nomorRequestLelangBOP"] input');
        $lelang.val('');
        if (this.value === 'Completed') {
            $lelang.prop('disabled', false);
        } else {
            $lelang.prop('disabled', true);
        }
    });
}


/* =============================================================================
 * STATUS TINDAKLANJUT & DOKUMEN TL (cross-workflow)
 * ============================================================================= */
function handleStatusDokumenTL(ctx) {
    // Cek apakah RFA chain sudah selesai untuk membuka Status Tindaklanjut
    const rfa1OK = ctx.rFA1 && ctx.approvalRFA1 === 'Approve';
    const rfa2OK = ctx.rFA2 && ctx.approvalRFA2 === 'Approve';
    const rfa3OK = ctx.rFA3 && ctx.approvalRFA3 === 'Approve';

    const chainComplete =
        (rfa1OK && !ctx.rFA2 && !ctx.rFA3) ||
        (rfa1OK && rfa2OK && !ctx.rFA3) ||
        (rfa1OK && rfa2OK && rfa3OK);

    // Status Tindaklanjut masih kosong -> PIC BOP set status & request dokumen
    if (chainComplete && ctx.statusTindaklanjut === '' && ctx.isPICBOP) {
        showField('ObjectField_statusTindaklanjut');
        showField('ObjectField_dokumenTindaklanjut');
        showField('ObjectField_keteranganRequestDokumenTL');

        unlockField('ObjectField_statusTindaklanjut');
        unlockField('ObjectField_keteranganRequestDokumenTL');
        // Checkbox dokumen tindaklanjut tetap dikunci di tahap ini
        lockField('ObjectField_dokumenTindaklanjut');
    }

    // Memerlukan Dokumen TL & belum lengkap -> User submit dokumen
    if (ctx.statusTindaklanjut === 'MemerlukanDokumenTL' &&
        ctx.statusSubmitDokume !== 'SudahLengkap') {

        showField('ObjectField_dokumenTindaklanjut');
        showField('ObjectField_statusSubmitDokume');
        showField('ObjectField_keteranganSubmitDokumenTL');

        // Checkbox dokumen di-disable, user hanya isi status submit + keterangan
        lockField('ObjectField_dokumenTindaklanjut');
        unlockField('ObjectField_statusSubmitDokume');
        unlockField('ObjectField_keteranganSubmitDokumenTL');
    }

    // Memerlukan Dokumen TL & sudah lengkap -> PIC BOP review checklist
    if (ctx.statusTindaklanjut === 'MemerlukanDokumenTL' &&
        ctx.statusSubmitDokume === 'SudahLengkap' &&
        ctx.isPICBOP) {

        showField('ObjectField_dokumenTindaklanjut');

        unlockField('ObjectField_dokumenTindaklanjut');
        unlockField('ObjectField_keteranganRequestDokumenTL');
        unlockField('ObjectField_statusTindaklanjut');
    }

    // PIC BOP final approval (setelah dokumen lengkap atau tidak diperlukan)
    const dokumenSiap =
        (ctx.statusTindaklanjut === 'TidakMemerlukanDokumenTL') ||
        (ctx.statusTindaklanjut === 'SudahLengkap') ||
        (ctx.statusSubmitDokume  === 'SudahLengkap') ||
        (ctx.statusSubmitDokume  === '');

    if (dokumenSiap &&
        ctx.approvalPICBOP === 'WaitingApproval' &&
        ctx.isPICBOP) {

        unlockField('ObjectField_approvalPICBOP');
        unlockField('ObjectField_keteranganPICBOP');
        unlockField('ObjectField_nomorTransaksi');
        unlockField('ObjectField_tanggalEfektif');
        unlockField('ObjectField_tanggalPenyelesaian');
        unlockField('ObjectField_jumlahHariLibur');
        unlockField('ObjectField_sla');
    }
}


/* =============================================================================
 * PRE-SAVE VALIDATION (pengganti SharePoint PreSaveAction)
 * ============================================================================= */
function registerPreSaveAction(ctx) {
    // Liferay Object Form submit biasanya pakai tombol .lfr-ddm__form-submit
    // atau button[type=submit] di dalam .bca-form-container
    $(document).off('click.wfPreSave').on('click.wfPreSave',
        '.bca-form-container button[type="submit"], .lfr-ddm__form-submit, .button-submit',
        function (e) {
            if (!preSaveValidate(ctx)) {
                e.preventDefault();
                e.stopImmediatePropagation();
                return false;
            }
        });
}

function preSaveValidate(ctx) {
    if (ctx.tindaklanjut !== 'PenjualanAsetTetapHasilReviewPengembalianBarangKhususUKKP') {
        return true;
    }
    if (!ctx.isPICBOP) {
        return true;
    }
    if (ctx.approvalReviewerUn !== 'Approve' || ctx.approvalPejabatUni !== 'Approve') {
        return true;
    }

    const inisial      = (getFieldValue('ObjectField_isiInisialPIC') || '').trim();
    const nomorRequest = (getFieldValue('ObjectField_nomorRequest')  || '').trim();
    if (inisial === '' && nomorRequest === '') {
        return true; // belum mulai diisi, biarkan
    }

    if (nomorRequest === '') {
        alert('Field "Nomor Request" wajib diisi.');
        return false;
    }
    if (inisial === '') {
        alert('Field "Isi Inisial PIC BOP" wajib diisi.');
        return false;
    }

    const statusReq = getFieldValue('ObjectField_statusRequestPenjualanAsetTetap') || '';
    if (statusReq === 'Completed') {
        const nomorLelang = (getFieldValue('ObjectField_nomorRequestLelangBOP') || '').trim();
        if (nomorLelang === '') {
            alert('Field "Nomor Request Lelang BOP" wajib diisi.');
            return false;
        }
    }
    return true;
}


/* =============================================================================
 * HELPER - LOCK/HIDE/SHOW
 * ============================================================================= */
function lockAllFields() {
    [
        'ObjectField_permohonanTindaklanjut',
        'ObjectField_title',
        'ObjectField_jenisTransaksi',
        'ObjectField_keteranganRequester',
        'ObjectField_statusRequest',
        'ObjectField_reviewerUnitKerja',
        'ObjectField_approvalReviewerUn',
        'ObjectField_keteranganApprovalReviewerUn',
        'ObjectField_pejabatUnitKerjaX',
        'ObjectField_approvalPejabatUni',
        'ObjectField_keteranganApprovalPejabatUni',
        'ObjectField_nomorRequest',
        'ObjectField_isiInisialPIC',
        'ObjectField_approvalPICBOP',
        'ObjectField_keteranganPICBOP',
        'ObjectField_nilaiBukuTertinggi',
        'ObjectField_hargaPenawaran',
        'ObjectField_hargaPenawaran0',
        'ObjectField_hargaPenawaran1',
        'ObjectField_hargaPenawaran2',
        'ObjectField_hargaPenawaran3',
        'ObjectField_hargaPenawaran4',
        'ObjectField_finalApproval',
        'ObjectField_reviewerBOP',
        'ObjectField_approvalReviewerBO',
        'ObjectField_keteranganApprovalReviewerBO',
        'ObjectField_rFA1',
        'ObjectField_approvalRFA1',
        'ObjectField_keteranganApprovalRFA1',
        'ObjectField_rFA2',
        'ObjectField_approvalRFA2',
        'ObjectField_keteranganApprovalRFA2',
        'ObjectField_rFA3',
        'ObjectField_approvalRFA3',
        'ObjectField_keteranganApprovalRFA3',
        'ObjectField_statusTindaklanjut',
        'ObjectField_dokumenTindaklanjut',
        'ObjectField_keteranganRequestDokumenTL',
        'ObjectField_statusSubmitDokume',
        'ObjectField_keteranganSubmitDokumenTL',
        'ObjectField_nomorTransaksi',
        'ObjectField_tanggalEfektif',
        'ObjectField_tanggalPenyelesaian',
        'ObjectField_jumlahHariLibur',
        'ObjectField_sla',
        'ObjectField_statusRequestPenjualanAsetTetap',
        'ObjectField_nomorRequestLelangBOP'
    ].forEach(lockField);
}

function hideAllConditionalFields() {
    [
        'ObjectField_jenisTransaksi',
        'ObjectField_reviewerUnitKerja',
        'ObjectField_approvalReviewerUn',
        'ObjectField_keteranganApprovalReviewerUn',
        'ObjectField_pejabatUnitKerjaX',
        'ObjectField_approvalPejabatUni',
        'ObjectField_keteranganApprovalPejabatUni',
        'ObjectField_statusRequest',
        'ObjectField_nomorRequest',
        'ObjectField_isiInisialPIC',
        'ObjectField_approvalPICBOP',
        'ObjectField_keteranganPICBOP',
        'ObjectField_nomorTransaksi',
        'ObjectField_tanggalEfektif',
        'ObjectField_tanggalPenyelesaian',
        'ObjectField_jumlahHariLibur',
        'ObjectField_sla',
        'ObjectField_approvalReviewerBO',
        'ObjectField_keteranganApprovalReviewerBO',
        'ObjectField_approvalRFA1',
        'ObjectField_keteranganApprovalRFA1',
        'ObjectField_approvalRFA2',
        'ObjectField_keteranganApprovalRFA2',
        'ObjectField_approvalRFA3',
        'ObjectField_keteranganApprovalRFA3',
        'ObjectField_statusTindaklanjut',
        'ObjectField_dokumenTindaklanjut',
        'ObjectField_keteranganRequestDokumenTL',
        'ObjectField_statusSubmitDokume',
        'ObjectField_keteranganSubmitDokumenTL',
        'ObjectField_nilaiBukuTertinggi',
        'ObjectField_hargaPenawaran',
        'ObjectField_hargaPenawaran0',
        'ObjectField_hargaPenawaran1',
        'ObjectField_hargaPenawaran2',
        'ObjectField_hargaPenawaran3',
        'ObjectField_hargaPenawaran4',
        'ObjectField_finalApproval',
        'ObjectField_reviewerBOP',
        'ObjectField_rFA1',
        'ObjectField_rFA2',
        'ObjectField_rFA3',
        'ObjectField_statusRequestPenjualanAsetTetap',
        'ObjectField_nomorRequestLelangBOP'
    ].forEach(hideField);
}

function unlockAllMainFields() {
    [
        'ObjectField_permohonanTindaklanjut',
        'ObjectField_title',
        'ObjectField_jenisTransaksi',
        'ObjectField_keteranganRequester',
        'ObjectField_approvalReviewerUn',
        'ObjectField_keteranganApprovalReviewerUn',
        'ObjectField_approvalPejabatUni',
        'ObjectField_keteranganApprovalPejabatUni',
        'ObjectField_nomorRequest',
        'ObjectField_isiInisialPIC',
        'ObjectField_keteranganPICBOP',
        'ObjectField_statusRequestPenjualanAsetTetap',
        'ObjectField_nomorRequestLelangBOP'
    ].forEach(unlockField);
}

function disableAll() {
    $('div.bca-form-container input:not([type="button"],[type="submit"],[type="reset"],[type="hidden"])').prop('disabled', true);
    $('div.bca-form-container textarea').prop('disabled', true);
    $('div.bca-form-container select').prop('disabled', true);
    $('div.bca-form-container button').prop('disabled', true);
    $('.button-submit').hide();
}


/* =============================================================================
 * HELPER - PER FIELD
 * ============================================================================= */
function lockField(fieldName) {
    const $f = $('div[data-field-name="' + fieldName + '"]');
    $f.find('input').prop('disabled', true);
    $f.find('select').prop('disabled', true);
    $f.find('textarea').prop('disabled', true);
    $f.find('button').prop('disabled', true);
}

function unlockField(fieldName) {
    const $f = $('div[data-field-name="' + fieldName + '"]');
    $f.find('input').prop('disabled', false);
    $f.find('select').prop('disabled', false);
    $f.find('textarea').prop('disabled', false);
    $f.find('button').prop('disabled', false);
}

function hideField(fieldName) {
    $('div[data-field-name="' + fieldName + '"]').hide();
}

function showField(fieldName) {
    $('div[data-field-name="' + fieldName + '"]').show();
}

function getFieldValue(fieldName) {
    const $f = $('div[data-field-name="' + fieldName + '"]');
    const $input = $f.find('input, select, textarea').first();
    return $input.length ? $input.val() : '';
}

function setFieldValue(fieldName, value) {
    const $f = $('div[data-field-name="' + fieldName + '"]');
    const $input = $f.find('input, select, textarea').first();
    if ($input.length) {
        $input.val(value).trigger('change');
    }
}

function resetApprovalToWaitingReview(fieldName) {
    setFieldValue(fieldName, 'WaitingReview');
}


/* =============================================================================
 * HELPER - VALIDASI USER LOGIN VS PICKER
 * ============================================================================= */
/**
 * Cek apakah user yang sedang login adalah picker target (reviewer/pejabat/RFA).
 * Picker di Liferay biasanya berbentuk { key, name } di mana:
 *   - key  = userId  (string angka)
 *   - name = display name / email
 * Kita compare dengan userData.id atau email.
 */
function isCurrentUser(picker, ctx) {
    if (!picker) return false;
    const pickerKey  = String(picker.key  || '').toLowerCase();
    const pickerName = String(picker.name || '').toLowerCase();

    const myId    = String(ctx.userData.id || '').toLowerCase();
    const myEmail = ctx.userEmail;
    const myName  = (ctx.userName || '').toLowerCase();

    if (!pickerKey && !pickerName) return false;

    if (pickerKey && (pickerKey === myId || pickerKey === myEmail)) return true;
    if (pickerName && (pickerName === myEmail || pickerName === myName)) return true;

    return false;
}
