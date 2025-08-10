import { product_data } from './data.js';

let supplementaryInsuredCount = 0;
let currentMainProductState = { product: null, age: null };

const MAX_ENTRY_AGE = {
    PUL_TRON_DOI: 70, PUL_15_NAM: 70, PUL_5_NAM: 70, KHOE_BINH_AN: 70, VUNG_TUONG_LAI: 70,
    TRON_TAM_AN: 60, AN_BINH_UU_VIET: 65,
    health_scl: 65, bhn: 70, accident: 64, hospital_support: 55
};

const MAX_RENEWAL_AGE = {
    health_scl: 74, bhn: 85, accident: 65, hospital_support: 59
};

const MAX_STBH = {
    bhn: 5_000_000_000,
    accident: 8_000_000_000
};

// Miễn Đóng Phí 3.0 (Section 5)
const WAIVER30 = { MIN_AGE: 18, MAX_AGE: 60, MAX_RENEWAL_AGE: 65 };

// Ngày tham chiếu tính tuổi: 09/08/2025
const REFERENCE_DATE = new Date(2025, 7, 9); // Tháng 8 là index 7

document.addEventListener('DOMContentLoaded', () => {
    initPerson(document.getElementById('main-person-container'), 'main');
    initMainProductLogic();
    initSupplementaryButton();
    initSummaryModal();
    initWaiver30Section();

    attachGlobalListeners();
    updateSupplementaryAddButtonState();
    calculateAll();
});

function attachGlobalListeners() {
    const allInputs = 'input, select';
    document.body.addEventListener('change', (e) => {
        const checkboxSelectors = [
            '.health-scl-checkbox',
            '.bhn-checkbox',
            '.accident-checkbox',
            '.hospital-support-checkbox'
        ];
        if (checkboxSelectors.some(selector => e.target.matches(selector))) {
            const section = e.target.closest('.product-section');
            const options = section.querySelector('.product-options');
            if (e.target.checked && !e.target.disabled) {
                options.classList.remove('hidden');
            } else {
                options.classList.add('hidden');
            }
            calculateAll();
        } else if (e.target.matches(allInputs)) {
            calculateAll();
        }
    });
    document.body.addEventListener('input', (e) => {
        if (e.target.matches('input[type="text"]') && !e.target.classList.contains('dob-input') && !e.target.classList.contains('occupation-input') && !e.target.classList.contains('name-input')) {
            formatNumberInput(e.target);
            calculateAll();
        } else if (e.target.matches('input[type="number"]')) {
            calculateAll();
        }
    });
}

/* ---------- Section 1: Người chính ---------- */
function initPerson(container, personId, isSupp = false) {
    if (!container) return;
    container.dataset.personId = personId;

    initDateFormatter(container.querySelector('.dob-input'));
    initOccupationAutocomplete(container.querySelector('.occupation-input'), container);

    // Nếu là NĐBH chính -> gắn validate khi blur/input
    if (!isSupp) {
        const nameInput = container.querySelector('.name-input');
        const dobInput = container.querySelector('.dob-input');
        const occInput = container.querySelector('.occupation-input');

        nameInput?.addEventListener('blur', validateMainPersonInputs);
        nameInput?.addEventListener('input', validateMainPersonInputs);

        dobInput?.addEventListener('blur', validateMainPersonInputs);
        dobInput?.addEventListener('input', validateMainPersonInputs);

        occInput?.addEventListener('input', validateMainPersonInputs);
        occInput?.addEventListener('blur', validateMainPersonInputs);
    }

    const suppProductsContainer = isSupp ? container.querySelector('.supplementary-products-container') : document.querySelector('#main-supp-container .supplementary-products-container');
    suppProductsContainer.innerHTML = generateSupplementaryProductsHtml(personId);

    const sclSection = suppProductsContainer.querySelector('.health-scl-section');
    if (sclSection) {
        const mainCheckbox = sclSection.querySelector('.health-scl-checkbox');
        const programSelect = sclSection.querySelector('.health-scl-program');
        const scopeSelect = sclSection.querySelector('.health-scl-scope');
        const outpatientCheckbox = sclSection.querySelector('.health-scl-outpatient');
        const dentalCheckbox = sclSection.querySelector('.health-scl-dental');

        const handleProgramChange = () => {
            const programChosen = programSelect.value !== '';
            outpatientCheckbox.disabled = !programChosen;
            dentalCheckbox.disabled = !programChosen;
            updateHealthSclStbhInfo(sclSection);
            if (!programChosen) {
                outpatientCheckbox.checked = false;
                dentalCheckbox.checked = false;
            }
            calculateAll();
        };

        const handleMainCheckboxChange = () => {
            const isChecked = mainCheckbox.checked && !mainCheckbox.disabled;
            programSelect.disabled = !isChecked;
            scopeSelect.disabled = !isChecked;
            const options = sclSection.querySelector('.product-options');
            options.classList.toggle('hidden', !isChecked);
            if (isChecked) {
                if (!programSelect.value) programSelect.value = 'nang_cao';
                if (!scopeSelect.value) scopeSelect.value = 'main_vn';
                updateHealthSclStbhInfo(sclSection);
            } else {
                programSelect.value = '';
                outpatientCheckbox.checked = false;
                dentalCheckbox.checked = false;
                updateHealthSclStbhInfo(sclSection);
            }
            handleProgramChange();
            calculateAll();
        };

        programSelect.addEventListener('change', handleProgramChange);
        mainCheckbox.addEventListener('change', handleMainCheckboxChange);
    }

    ['bhn', 'accident', 'hospital-support'].forEach(product => {
        const section = suppProductsContainer.querySelector(`.${product}-section`);
        if (section) {
            const checkbox = section.querySelector(`.${product}-checkbox`);
            const handleCheckboxChange = () => {
                const isChecked = checkbox.checked && !checkbox.disabled;
                const options = section.querySelector('.product-options');
                options.classList.toggle('hidden', !isChecked);
                calculateAll();
            };
            checkbox.addEventListener('change', handleCheckboxChange);
        }
    });

    const hsInput = suppProductsContainer.querySelector('.hospital-support-section .hospital-support-stbh');
    if (hsInput) {
        hsInput.addEventListener('blur', () => {
            const raw = parseFormattedNumber(hsInput.value || '0');
            if (raw <= 0) return;
            const rounded = Math.round(raw / 100000) * 100000;
            if (rounded !== raw) {
                hsInput.value = rounded.toLocaleString('vi-VN');
            }
            calculateAll();
        });
    }
}

function initMainProductLogic() {
    document.getElementById('main-product').addEventListener('change', () => {
        updateSupplementaryAddButtonState();
        calculateAll();
    });
}

/* ---------- Section 4: NĐBH bổ sung ---------- */
function getSupplementaryCount() {
    return document.querySelectorAll('#supplementary-insured-container .person-container').length;
}
function updateSupplementaryAddButtonState() {
    const btn = document.getElementById('add-supp-insured-btn');
    if (!btn) return;
    const mainProduct = document.getElementById('main-product')?.value || '';
    const count = getSupplementaryCount();
    const disabled = (mainProduct === 'TRON_TAM_AN') || (count >= 10);
    btn.disabled = disabled;
    btn.classList.toggle('opacity-50', disabled);
    btn.classList.toggle('cursor-not-allowed', disabled);
}
function initSupplementaryButton() {
    document.getElementById('add-supp-insured-btn').addEventListener('click', () => {
        if (getSupplementaryCount() >= 10) {
            updateSupplementaryAddButtonState();
            return;
        }
        supplementaryInsuredCount++;
        const personId = `supp${supplementaryInsuredCount}`;
        const container = document.getElementById('supplementary-insured-container');
        const newPersonDiv = document.createElement('div');
        newPersonDiv.className = 'person-container space-y-6 bg-gray-100 p-4 rounded-lg mt-4';
        newPersonDiv.id = `person-container-${personId}`;
        newPersonDiv.innerHTML = generateSupplementaryPersonHtml(personId, supplementaryInsuredCount);
        container.appendChild(newPersonDiv);
        initPerson(newPersonDiv, personId, true);
        updateSupplementaryAddButtonState();
        calculateAll();
    });
}

/* ---------- Section 5: Miễn Đóng Phí 3.0 ---------- */
function initWaiver30Section() {
    const cont = document.getElementById('waiver30-container');
    if (!cont) return;
    // render lần đầu, phần còn lại gọi trong calculateAll
    cont.innerHTML = '';
}

function renderWaiver30Section(mainPersonInfo) {
    const cont = document.getElementById('waiver30-container');
    if (!cont) return;

    const mainProduct = document.getElementById('main-product')?.value || '';
    if (mainProduct === 'TRON_TAM_AN') {
        cont.classList.add('hidden');
        cont.innerHTML = '';
        return;
    }
    cont.classList.remove('hidden');

    // Lấy danh sách NĐBH bổ sung đủ điều kiện (18-60)
    const eligibleSupps = [];
    document.querySelectorAll('#supplementary-insured-container .person-container').forEach(p => {
        const info = getCustomerInfo(p, false);
        if (info.age >= WAIVER30.MIN_AGE && info.age <= WAIVER30.MAX_AGE) {
            eligibleSupps.push({ id: p.id, name: info.name || 'NĐBH bổ sung', age: info.age, gender: info.gender, container: p });
        }
    });

    // Giữ lại lựa chọn cũ nếu có
    const prevSelected = cont.querySelector('input[name="waiver30-insurer"]:checked')?.value || cont.dataset.selected || (eligibleSupps[0]?.id || 'other');

    let radiosHtml = '';
    if (eligibleSupps.length > 0) {
        radiosHtml += `<div class="space-y-2">`;
        eligibleSupps.forEach(p => {
            radiosHtml += `
            <label class="flex items-center space-x-3 cursor-pointer">
                <input type="radio" name="waiver30-insurer" class="form-checkbox" value="${p.id}" ${prevSelected === p.id ? 'checked' : ''}>
                <span>${sanitizeHtml(p.name)} (Tuổi ${p.age})</span>
            </label>`;
        });
        radiosHtml += `</div>`;
    }

    // Radio Người khác
    radiosHtml += `
    <label class="flex items-center space-x-3 cursor-pointer mt-2">
        <input type="radio" name="waiver30-insurer" class="form-checkbox" value="other" ${prevSelected === 'other' ? 'checked' : ''}>
        <span>Người khác</span>
    </label>`;

    const otherFormHtml = `
    <div id="waiver30-other-form" class="${prevSelected === 'other' ? '' : 'hidden'} mt-3">
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
                <label class="font-medium text-gray-700 block mb-1">Họ và Tên</label>
                <input type="text" class="form-input waiver-other-name" placeholder="Nguyễn Văn C">
            </div>
            <div>
                <label class="font-medium text-gray-700 block mb-1">Ngày sinh</label>
                <input type="text" class="form-input waiver-other-dob dob-input" placeholder="DD/MM/YYYY">
            </div>
            <div>
                <label class="font-medium text-gray-700 block mb-1">Giới tính</label>
                <select class="form-select waiver-other-gender">
                    <option value="Nam">Nam</option>
                    <option value="Nữ">Nữ</option>
                </select>
            </div>
            <div class="relative">
                <label class="font-medium text-gray-700 block mb-1">Nghề nghiệp</label>
                <input type="text" class="form-input waiver-other-occupation occupation-input" placeholder="Gõ để tìm nghề nghiệp...">
                <div class="occupation-autocomplete absolute z-10 w-full bg-white border border-gray-300 rounded-md mt-1 hidden max-h-60 overflow-y-auto"></div>
            </div>
        </div>
    </div>`;

    cont.innerHTML = `
        <h3 class="text-lg font-bold text-gray-800">Miễn Đóng Phí 3.0 (Bên mua BH)</h3>
        <div class="space-y-3">
            ${eligibleSupps.length > 0 ? `<p class="text-sm text-gray-600">Chọn người áp dụng (18–60 tuổi) hoặc chọn “Người khác”</p>` : `<p class="text-sm text-gray-600">Không có NĐBH bổ sung 18–60 tuổi. Vui lòng chọn “Người khác”.</p>`}
            ${radiosHtml}
            ${otherFormHtml}
            <div class="text-right font-semibold text-aia-red waiver30-fee-display min-h-[1.5rem]"></div>
        </div>
    `;

    // Gắn events: show/hide Other form
    cont.querySelectorAll('input[name="waiver30-insurer"]').forEach(r => {
        r.addEventListener('change', () => {
            cont.dataset.selected = r.value;
            const otherForm = cont.querySelector('#waiver30-other-form');
            otherForm?.classList.toggle('hidden', r.value !== 'other');
            calculateAll();
        });
    });

    // Init formatter + autocomplete cho "Người khác"
    const otherDob = cont.querySelector('.waiver-other-dob');
    const otherOcc = cont.querySelector('.waiver-other-occupation');
    if (otherDob) initDateFormatter(otherDob);
    if (otherOcc) initOccupationAutocomplete(otherOcc, cont);
}

function getWaiverOtherInfo() {
    const cont = document.getElementById('waiver30-container');
    if (!cont) return null;
    const nameInput = cont.querySelector('.waiver-other-name');
    const dobInput = cont.querySelector('.waiver-other-dob');
    const genderSelect = cont.querySelector('.waiver-other-gender');
    const occInput = cont.querySelector('.waiver-other-occupation');

    if (!dobInput || !genderSelect || !occInput) return null;

    // Tính tuổi theo REFERENCE_DATE
    let age = 0;
    const dobStr = (dobInput.value || '').trim();
    if (/^\d{2}\/\d{2}\/\d{4}$/.test(dobStr)) {
        const [dd, mm, yyyy] = dobStr.split('/').map(n => parseInt(n, 10));
        const birth = new Date(yyyy, mm - 1, dd);
        const isValid = birth.getFullYear() === yyyy && birth.getMonth() === (mm - 1) && birth.getDate() === dd && birth <= REFERENCE_DATE;
        if (isValid) {
            age = REFERENCE_DATE.getFullYear() - birth.getFullYear();
            const m = REFERENCE_DATE.getMonth() - birth.getMonth();
            if (m < 0 || (m === 0 && REFERENCE_DATE.getDate() < birth.getDate())) age--;
        }
    }
    const riskGroup = parseInt(occInput.dataset.group || '0', 10) || 0;

    return {
        name: (nameInput?.value || '').trim() || 'Người khác',
        age,
        gender: genderSelect.value,
        riskGroup
    };
}

function validateWaiverOtherForm() {
    const cont = document.getElementById('waiver30-container');
    if (!cont) return false;
    const nameInput = cont.querySelector('.waiver-other-name');
    const dobInput = cont.querySelector('.waiver-other-dob');
    const occInput = cont.querySelector('.waiver-other-occupation');

    let ok = true;

    // Tên
    if (nameInput) {
        const v = (nameInput.value || '').trim();
        if (!v) { setFieldError(nameInput, 'Vui lòng nhập họ và tên'); ok = false; }
        else { clearFieldError(nameInput); }
    }

    // DOB
    if (dobInput) {
        const v = (dobInput.value || '').trim();
        const re = /^\d{2}\/\d{2}\/\d{4}$/;
        if (!re.test(v)) { setFieldError(dobInput, 'Ngày sinh không hợp lệ, nhập DD/MM/YYYY'); ok = false; }
        else {
            const [dd, mm, yyyy] = v.split('/').map(n => parseInt(n, 10));
            const d = new Date(yyyy, mm - 1, dd);
            const valid = d.getFullYear() === yyyy && d.getMonth() === (mm - 1) && d.getDate() === dd && d <= REFERENCE_DATE;
            if (!valid) { setFieldError(dobInput, 'Ngày sinh không hợp lệ, nhập DD/MM/YYYY'); ok = false; }
            else clearFieldError(dobInput);
        }
    }

    // Occupation chọn từ danh sách
    if (occInput) {
        const typed = (occInput.value || '').trim().toLowerCase();
        const match = product_data.occupations.find(o => o.group > 0 && o.name.toLowerCase() === typed);
        const group = parseInt(occInput.dataset.group, 10);
        if (!match || !(group >= 1 && group <= 4)) {
            setFieldError(occInput, 'Chọn nghề nghiệp từ danh sách');
            ok = false;
        } else {
            clearFieldError(occInput);
        }
    }
    // Tuổi 18-60
    const other = getWaiverOtherInfo();
    if (!other || other.age < WAIVER30.MIN_AGE || other.age > WAIVER30.MAX_AGE) {
        if (dobInput) setFieldError(dobInput, 'Độ tuổi phải từ 18 đến 60');
        ok = false;
    }
    return ok;
}

function calculateWaiver30Premium(mainPersonInfo, mainPremiumDisplay) {
    const cont = document.getElementById('waiver30-container');
    if (!cont) return 0;
    const mainProduct = document.getElementById('main-product')?.value || '';
    const feeEl = cont.querySelector('.waiver30-fee-display');
    if (mainProduct === 'TRON_TAM_AN') {
        if (feeEl) feeEl.textContent = '';
        return 0;
    }

    const selected = cont.querySelector('input[name="waiver30-insurer"]:checked')?.value || 'other';

    // Tính tổng phí SP bổ sung của NĐBH chính
    const mainSuppContainer = document.querySelector('#main-supp-container .supplementary-products-container');
    let suppMain = 0;
    if (mainSuppContainer) {
        suppMain += calculateHealthSclPremium(mainPersonInfo, mainSuppContainer);
        suppMain += calculateBhnPremium(mainPersonInfo, mainSuppContainer);
        suppMain += calculateAccidentPremium(mainPersonInfo, mainSuppContainer);
        // Hospital support cần hạn mức theo phí chính (không tính extra)
        const baseMainPremium = mainPremiumDisplay - getExtraPremiumValue();
        suppMain += calculateHospitalSupportPremium(mainPersonInfo, baseMainPremium, mainSuppContainer, 0);
    }

    // Tính tổng phí SP bổ sung của tất cả NĐBH bổ sung và đồng thời lấy phí của người đang chọn (nếu chọn từ danh sách)
    let sumSuppAllSupp = 0;
    let suppOfSelected = 0;

    document.querySelectorAll('#supplementary-insured-container .person-container').forEach(p => {
        const personInfo = getCustomerInfo(p, false);
        const container = p.querySelector('.supplementary-products-container');
        if (!container) return;
        let sub = 0;
        sub += calculateHealthSclPremium(personInfo, container);
        sub += calculateBhnPremium(personInfo, container);
        sub += calculateAccidentPremium(personInfo, container);
        // Hospital support theo phí chính (không extra)
        const baseMainPremium = mainPremiumDisplay - getExtraPremiumValue();
        sub += calculateHospitalSupportPremium(personInfo, baseMainPremium, container, 0);
        sumSuppAllSupp += sub;
        if (selected === p.id) {
            suppOfSelected = sub;
        }
    });

    // STBH theo quy tắc
    // Nếu chọn "other": không trừ; nếu chọn một NĐBH bổ sung thì trừ phần phí SP bổ sung của người đó.
    const stbh = Math.max(0, mainPremiumDisplay + suppMain + sumSuppAllSupp - (selected === 'other' ? 0 : suppOfSelected));

    // Xác định người được chọn để tính biểu phí (tuổi/giới tính)
    let insuredAge = 0;
    let insuredGender = 'Nam';
    if (selected === 'other') {
        // validate other
        if (!validateWaiverOtherForm()) {
            if (feeEl) feeEl.textContent = '';
            return 0;
        }
        const other = getWaiverOtherInfo();
        insuredAge = other.age;
        insuredGender = other.gender;
    } else {
        const pEl = document.getElementById(selected);
        const pInfo = pEl ? getCustomerInfo(pEl, false) : null;
        if (!pInfo || pInfo.age < WAIVER30.MIN_AGE || pInfo.age > WAIVER30.MAX_AGE) {
            if (feeEl) feeEl.textContent = '';
            return 0;
        }
        insuredAge = pInfo.age;
        insuredGender = pInfo.gender;
    }

    // Tra biểu phí Miễn đóng phí 3.0
    const rateRow = product_data.waiver30_rates?.find(r => insuredAge >= r.ageMin && insuredAge <= r.ageMax);
    if (!rateRow) {
        throw new Error(`Không có biểu phí Miễn đóng phí 3.0 cho tuổi ${insuredAge}.`);
    }
    const rate = insuredGender === 'Nữ' ? (rateRow.nu || 0) : (rateRow.nam || 0);
    const premium = (stbh / 1000) * rate;

    if (feeEl) feeEl.textContent = premium > 0 ? `Phí: ${formatCurrency(premium)}` : '';

    return premium;
}

/* ---------- Section 6 / Modal ---------- */
function initSummaryModal() {
    const modal = document.getElementById('summary-modal');
    document.getElementById('view-summary-btn').addEventListener('click', generateSummaryTable);
    document.getElementById('close-summary-modal-btn').addEventListener('click', () => modal.classList.add('hidden'));
    modal.addEventListener('click', (e) => {
        if (e.target === modal) modal.classList.add('hidden');
    });

    // Xử lý input target-age-input
    const targetAgeInput = document.getElementById('target-age-input');
    const mainPersonContainer = document.getElementById('main-person-container');
    const mainPersonInfo = getCustomerInfo(mainPersonContainer, true);
    const mainProduct = mainPersonInfo.mainProduct;

    if (mainProduct === 'TRON_TAM_AN') {
        targetAgeInput.value = mainPersonInfo.age + 10 - 1;
        targetAgeInput.disabled = true;
    } else if (mainProduct === 'AN_BINH_UU_VIET') {
        const termSelect = document.getElementById('abuv-term');
        const term = parseInt(termSelect?.value || '15', 10);
        targetAgeInput.value = mainPersonInfo.age + term - 1;
        targetAgeInput.disabled = true;
    } else {
        const paymentTermInput = document.getElementById('payment-term');
        const paymentTerm = paymentTermInput ? parseInt(paymentTermInput.value, 10) || 0 : 0;
        targetAgeInput.disabled = false;
        targetAgeInput.min = mainPersonInfo.age + paymentTerm - 1;
        if (!targetAgeInput.value || parseInt(targetAgeInput.value, 10) < mainPersonInfo.age + paymentTerm - 1) {
            targetAgeInput.value = mainPersonInfo.age + paymentTerm - 1;
        }
    }

    const abuvTermSelect = document.getElementById('abuv-term');
    document.getElementById('main-product').addEventListener('change', () => {
        updateTargetAge();
        if (document.getElementById('summary-modal').classList.contains('hidden')) {
            calculateAll();
        } else {
            generateSummaryTable();
        }
    });

    const mainDobInput = document.querySelector('#main-person-container .dob-input');
    if (mainDobInput) {
        mainDobInput.addEventListener('input', () => {
            updateTargetAge();
            if (document.getElementById('summary-modal').classList.contains('hidden')) {
                calculateAll();
            } else {
                generateSummaryTable();
            }
        });
    }

    if (abuvTermSelect) {
        abuvTermSelect.addEventListener('change', () => {
            updateTargetAge();
            if (document.getElementById('summary-modal').classList.contains('hidden')) {
                calculateAll();
            } else {
                generateSummaryTable();
            }
        });
    }
    document.getElementById('payment-term')?.addEventListener('change', () => {
        updateTargetAge();
        if (document.getElementById('summary-modal').classList.contains('hidden')) {
            calculateAll();
        } else {
            generateSummaryTable();
        }
    });
}

function updateTargetAge() {
    const mainPersonContainer = document.getElementById('main-person-container');
    const mainPersonInfo = getCustomerInfo(mainPersonContainer, true);
    const mainProduct = mainPersonInfo.mainProduct;
    const targetAgeInput = document.getElementById('target-age-input');

    if (mainProduct === 'TRON_TAM_AN') {
        targetAgeInput.value = mainPersonInfo.age + 10 - 1;
        targetAgeInput.disabled = true;
    } else if (mainProduct === 'AN_BINH_UU_VIET') {
        const termSelect = document.getElementById('abuv-term');
        const term = termSelect ? parseInt(termSelect.value || '15', 10) : 15;
        targetAgeInput.value = mainPersonInfo.age + term - 1;
        targetAgeInput.disabled = true;
    } else {
        const paymentTermInput = document.getElementById('payment-term');
        const paymentTerm = paymentTermInput ? parseInt(paymentTermInput.value, 10) || 0 : 0;
        targetAgeInput.disabled = false;
        targetAgeInput.min = mainPersonInfo.age + paymentTerm - 1;
        if (!targetAgeInput.value || parseInt(targetAgeInput.value, 10) < mainPersonInfo.age + paymentTerm - 1) {
            targetAgeInput.value = mainPersonInfo.age + paymentTerm - 1;
        }
    }
}

/* ---------- Formatters & Autocomplete ---------- */
function initDateFormatter(input) {
    if (!input) return;
    input.addEventListener('input', (e) => {
        let value = e.target.value.replace(/\D/g, '');
        if (value.length > 2) value = value.slice(0, 2) + '/' + value.slice(2);
        if (value.length > 5) value = value.slice(0, 5) + '/' + value.slice(5, 9);
        e.target.value = value.slice(0, 10);
    });
}

// Autocomplete nghề: dùng mousedown để chọn trước blur
function initOccupationAutocomplete(input, container) {
    if (!input) return;
    const autocompleteContainer = container.querySelector('.occupation-autocomplete');
    const riskGroupSpan = container.querySelector('.risk-group-span');

    const applyOccupation = (occ) => {
        input.value = occ.name;
        input.dataset.group = occ.group;
        if (riskGroupSpan) riskGroupSpan.textContent = occ.group;
        clearFieldError(input);
        autocompleteContainer.classList.add('hidden');
        calculateAll();
    };

    const renderList = (filtered) => {
        autocompleteContainer.innerHTML = '';
        if (filtered.length === 0) {
            autocompleteContainer.classList.add('hidden');
            return;
        }
        filtered.forEach(occ => {
            const item = document.createElement('div');
            item.className = 'autocomplete-item';
            item.textContent = occ.name;
            item.addEventListener('mousedown', (ev) => {
                ev.preventDefault();
                applyOccupation(occ);
            });
            autocompleteContainer.appendChild(item);
        });
        autocompleteContainer.classList.remove('hidden');
    };

    input.addEventListener('input', () => {
        const value = input.value.trim().toLowerCase();
        if (value.length < 2) {
            autocompleteContainer.classList.add('hidden');
            return;
        }
        const filtered = product_data.occupations
            .filter(o => o.group > 0 && o.name.toLowerCase().includes(value));
        renderList(filtered);
    });

    input.addEventListener('blur', () => {
        setTimeout(() => {
            const typed = (input.value || '').trim().toLowerCase();
            const match = product_data.occupations.find(o => o.group > 0 && o.name.toLowerCase() === typed);
            if (typed && match) {
                applyOccupation(match);
            } else {
                input.dataset.group = '';
                if (riskGroupSpan) riskGroupSpan.textContent = '...';
                setFieldError(input, 'Chọn nghề nghiệp từ danh sách');
                autocompleteContainer.classList.add('hidden');
                calculateAll();
            }
        }, 0);
    });

    document.addEventListener('click', (e) => {
        if (!container.contains(e.target)) {
            autocompleteContainer.classList.add('hidden');
        }
    });
}

/* ---------- Helpers chung ---------- */
function getCustomerInfo(container, isMain = false) {
    const dobInput = container.querySelector('.dob-input');
    const genderSelect = container.querySelector('.gender-select');
    const occupationInput = container.querySelector('.occupation-input');
    const ageSpan = container.querySelector('.age-span');
    const riskGroupSpan = container.querySelector('.risk-group-span');
    const nameInput = container.querySelector('.name-input');

    let age = 0;
    let daysFromBirth = 0;

    const dobStr = dobInput ? dobInput.value : '';
    if (dobStr && /^\d{2}\/\d{2}\/\d{4}$/.test(dobStr)) {
        const [dd, mm, yyyy] = dobStr.split('/').map(n => parseInt(n, 10));
        const birthDate = new Date(yyyy, mm - 1, dd);
        const isValidDate = birthDate.getFullYear() === yyyy && (birthDate.getMonth() === (mm - 1)) && birthDate.getDate() === dd;
        if (isValidDate && birthDate <= REFERENCE_DATE) {
            const diffMs = REFERENCE_DATE - birthDate;
            daysFromBirth = Math.floor(diffMs / (1000 * 60 * 60 * 24));

            age = REFERENCE_DATE.getFullYear() - birthDate.getFullYear();
            const m = REFERENCE_DATE.getMonth() - birthDate.getMonth();
            if (m < 0 || (m === 0 && REFERENCE_DATE.getDate() < birthDate.getDate())) {
                age--;
            }
        }
    }

    if (ageSpan) ageSpan.textContent = age;
    const riskGroup = occupationInput ? parseInt(occupationInput.dataset.group, 10) || 0 : 0;
    if (riskGroupSpan) riskGroupSpan.textContent = riskGroup > 0 ? riskGroup : '...';

    const info = {
        age,
        daysFromBirth,
        gender: genderSelect ? genderSelect.value : 'Nam',
        riskGroup,
        container,
        name: nameInput ? nameInput.value : 'NĐBH Chính'
    };

    if (isMain) {
        info.mainProduct = document.getElementById('main-product').value;
    }

    return info;
}

function calculateAll() {
    try {
        clearError();
        validateMainPersonInputs();

        const mainPersonContainer = document.getElementById('main-person-container');
        const mainPersonInfo = getCustomerInfo(mainPersonContainer, true);

        updateMainProductVisibility(mainPersonInfo);
        // Render lại Section 5 theo người & sản phẩm
        renderWaiver30Section(mainPersonInfo);

        validateSection2FieldsPreCalc(mainPersonInfo);

        const baseMainPremium = calculateMainPremium(mainPersonInfo);
        validateExtraPremiumLimit(baseMainPremium);
        const extraPremium = getExtraPremiumValue();
        const mainPremiumDisplay = baseMainPremium + extraPremium;

        updateMainProductFeeDisplay(baseMainPremium, extraPremium);

        updateSupplementaryProductVisibility(mainPersonInfo, baseMainPremium, document.querySelector('#main-supp-container .supplementary-products-container'));

        let totalSupplementaryPremium = 0;
        let totalHospitalSupportStbh = 0; // Theo dõi tổng STBH viện phí (tất cả người)

        document.querySelectorAll('.person-container').forEach(container => {
            const isMain = container.id === 'main-person-container';
            const personInfo = getCustomerInfo(container, isMain);
            const suppProductsContainer = isMain ? document.querySelector('#main-supp-container .supplementary-products-container') : container.querySelector('.supplementary-products-container');

            if (!suppProductsContainer) return;

            updateSupplementaryProductVisibility(personInfo, baseMainPremium, suppProductsContainer);

            totalSupplementaryPremium += calculateHealthSclPremium(personInfo, suppProductsContainer);
            totalSupplementaryPremium += calculateBhnPremium(personInfo, suppProductsContainer);
            totalSupplementaryPremium += calculateAccidentPremium(personInfo, suppProductsContainer);
            totalSupplementaryPremium += calculateHospitalSupportPremium(personInfo, baseMainPremium, suppProductsContainer, totalHospitalSupportStbh);
            const hospitalSupportStbh = parseFormattedNumber(suppProductsContainer.querySelector('.hospital-support-stbh')?.value || '0');
            if (suppProductsContainer.querySelector('.hospital-support-checkbox')?.checked && hospitalSupportStbh > 0) {
                totalHospitalSupportStbh += hospitalSupportStbh;
            }
        });

        // Tính phí Miễn Đóng Phí 3.0 và cộng vào tổng bổ sung
        const waiverPremium = calculateWaiver30Premium(mainPersonInfo, mainPremiumDisplay);
        totalSupplementaryPremium += waiverPremium;

        const totalPremium = mainPremiumDisplay + totalSupplementaryPremium;
        updateSummaryUI({ mainPremium: mainPremiumDisplay, totalSupplementaryPremium, totalPremium });

    } catch (error) {
        showError(error.message);
        updateSummaryUI({ mainPremium: 0, totalSupplementaryPremium: 0, totalPremium: 0 });
    }
}

function updateMainProductVisibility(customer) {
    const { age, daysFromBirth, gender, riskGroup } = customer;
    const mainProductSelect = document.getElementById('main-product');

    document.querySelectorAll('#main-product option').forEach(option => {
        const productKey = option.value;
        if (!productKey) return;

        let isEligible = true;

        const PUL_MUL = ['PUL_TRON_DOI', 'PUL_15_NAM', 'PUL_5_NAM', 'KHOE_BINH_AN', 'VUNG_TUONG_LAI'];
        if (PUL_MUL.includes(productKey)) {
            isEligible = (daysFromBirth >= 30) && (age <= 70);
        }
        if (productKey === 'TRON_TAM_AN') {
            const withinAgeByGender = (gender === 'Nam')
                ? (age >= 12 && age <= 60)
                : (age >= 28 && age <= 60);
            isEligible = withinAgeByGender && (riskGroup !== 4);
        }
        if (productKey === 'AN_BINH_UU_VIET') {
            const minOk = (gender === 'Nam') ? age >= 12 : age >= 28;
            isEligible = minOk && (age <= 65);
        }

        option.disabled = !isEligible;
        option.classList.toggle('hidden', !isEligible);
    });

    if (mainProductSelect.options[mainProductSelect.selectedIndex]?.disabled) {
        mainProductSelect.value = "";
    }

    const newProduct = mainProductSelect.value;

    if (newProduct === 'TRON_TAM_AN') {
        document.getElementById('supplementary-insured-container').classList.add('hidden');
        document.getElementById('add-supp-insured-btn').classList.add('hidden');
        supplementaryInsuredCount = 0;
        document.getElementById('supplementary-insured-container').innerHTML = '';
    } else {
        document.getElementById('supplementary-insured-container').classList.remove('hidden');
        document.getElementById('add-supp-insured-btn').classList.remove('hidden');
    }

    if (currentMainProductState.product !== newProduct || currentMainProductState.age !== age) {
        renderMainProductOptions(customer);
        currentMainProductState.product = newProduct;
        currentMainProductState.age = age;
    }

    updateSupplementaryAddButtonState();

    // Ẩn Section 5 nếu là Trọn Tâm An
    const waiverCont = document.getElementById('waiver30-container');
    if (waiverCont) {
        if (newProduct === 'TRON_TAM_AN') waiverCont.classList.add('hidden');
        else waiverCont.classList.remove('hidden');
    }
}

function updateSupplementaryProductVisibility(customer, mainPremium, container) {
    const { age, riskGroup, daysFromBirth } = customer;
    const mainProduct = document.getElementById('main-product').value;

    const showOrHide = (sectionId, productKey, condition) => {
        const section = container.querySelector(`.${sectionId}-section`);
        if (!section) {
            console.error(`Không tìm thấy section ${sectionId}`);
            return;
        }
        const checkbox = section.querySelector('input[type="checkbox"]');
        const options = section.querySelector('.product-options');
        const finalCondition = condition
            && daysFromBirth >= 30
            && age >= 0 && age <= MAX_ENTRY_AGE[productKey]
            && (sectionId !== 'health-scl' || riskGroup !== 4);

        if (finalCondition) {
            section.classList.remove('hidden');
            checkbox.disabled = false;
            options.classList.toggle('hidden', !checkbox.checked || checkbox.disabled);

            if (sectionId === 'health-scl') {
                const programSelect = section.querySelector('.health-scl-program');
                const scopeSelect = section.querySelector('.health-scl-scope');
                const outpatient = section.querySelector('.health-scl-outpatient');
                const dental = section.querySelector('.health-scl-dental');

                if (mainProduct === 'TRON_TAM_AN') {
                    checkbox.checked = true;
                    checkbox.disabled = true;
                    options.classList.remove('hidden');
                    programSelect.disabled = false;
                    scopeSelect.disabled = false;

                    Array.from(programSelect.options).forEach(opt => { if (opt.value) opt.disabled = false; });
                    if (!programSelect.value || programSelect.options[programSelect.selectedIndex]?.disabled) {
                        if (!programSelect.querySelector('option[value="nang_cao"]').disabled) {
                            programSelect.value = 'nang_cao';
                        }
                    }
                    if (!scopeSelect.value) scopeSelect.value = 'main_vn';
                    // Cho tick Ngoại trú/Nha khoa khi là TTA
                    outpatient.disabled = false;
                    dental.disabled = false;

                    updateHealthSclStbhInfo(section);
                } else {
                    programSelect.disabled = false;
                    scopeSelect.disabled = false;
                    programSelect.querySelectorAll('option').forEach(opt => {
                        if (opt.value === '') return;
                        if (mainPremium >= 15000000) {
                            opt.disabled = false;
                        } else if (mainPremium >= 10000000) {
                            opt.disabled = !['co_ban', 'nang_cao', 'toan_dien'].includes(opt.value);
                        } else if (mainPremium >= 5000000) {
                            opt.disabled = !['co_ban', 'nang_cao'].includes(opt.value);
                        } else {
                            opt.disabled = true;
                        }
                    });
                    if (!programSelect.value || programSelect.options[programSelect.selectedIndex]?.disabled) {
                        const nangCao = programSelect.querySelector('option[value="nang_cao"]');
                        if (nangCao && !nangCao.disabled) {
                            programSelect.value = 'nang_cao';
                        } else {
                            const firstEnabled = Array.from(programSelect.options).find(opt => opt.value && !opt.disabled);
                            programSelect.value = firstEnabled ? firstEnabled.value : '';
                        }
                    }
                    if (!scopeSelect.value) scopeSelect.value = 'main_vn';

                    const hasProgram = programSelect.value !== '';
                    outpatient.disabled = !hasProgram;
                    dental.disabled = !hasProgram;

                    updateHealthSclStbhInfo(section);
                }
            }
        } else {
            section.classList.add('hidden');
            checkbox.checked = false;
            checkbox.disabled = true;
            options.classList.add('hidden');
        }
    };

    const baseCondition = ['PUL_TRON_DOI', 'PUL_15_NAM', 'PUL_5_NAM', 'KHOE_BINH_AN', 'VUNG_TUONG_LAI', 'AN_BINH_UU_VIET', 'TRON_TAM_AN'].includes(mainProduct);

    showOrHide('health-scl', 'health_scl', baseCondition);
    showOrHide('bhn', 'bhn', baseCondition);
    showOrHide('accident', 'accident', baseCondition);
    showOrHide('hospital-support', 'hospital_support', baseCondition);

    if (mainProduct === 'TRON_TAM_AN') {
        ['bhn', 'accident', 'hospital-support'].forEach(id => {
            const section = container.querySelector(`.${id}-section`);
            if (section) {
                section.classList.add('hidden');
                section.querySelector('input[type="checkbox"]').checked = false;
                section.querySelector('.product-options').classList.add('hidden');
            }
        });
    }
}

/* ---------- Section 2: SP chính ---------- */
function renderMainProductOptions(customer) {
    const container = document.getElementById('main-product-options');
    const { mainProduct, age } = customer;

    let currentStbh = container.querySelector('#main-stbh')?.value || '';
    let currentPremium = container.querySelector('#main-premium-input')?.value || '';
    let currentPaymentTerm = container.querySelector('#payment-term')?.value || '';
    let currentExtra = container.querySelector('#extra-premium-input')?.value || '';

    container.innerHTML = '';
    if (!mainProduct) return;

    let optionsHtml = '';

    if (mainProduct === 'TRON_TAM_AN') {
        optionsHtml = `
            <div>
                <label for="main-stbh" class="font-medium text-gray-700 block mb-1">Số tiền bảo hiểm (STBH)</label>
                <input type="text" id="main-stbh" class="form-input bg-gray-100" value="100.000.000" disabled>
            </div>
            <div>
                <p class="text-sm text-gray-600 mt-1">Thời hạn đóng phí: 10 năm (bằng thời hạn hợp đồng). Thời gian bảo vệ: 10 năm.</p>
            </div>`;
    } else if (mainProduct === 'AN_BINH_UU_VIET') {
        optionsHtml = `
            <div>
                <label for="main-stbh" class="font-medium text-gray-700 block mb-1">Số tiền bảo hiểm (STBH)</label>
                <input type="text" id="main-stbh" class="form-input" value="${currentStbh}" placeholder="VD: 1.000.000.000">
            </div>`;
        let termOptions = '';
        if (age <= 55) termOptions += '<option value="15">15 năm</option>';
        if (age <= 60) termOptions += '<option value="10">10 năm</option>';
        if (age <= 65) termOptions += '<option value="5">5 năm</option>';
        if (!termOptions) termOptions = '<option value="" disabled>Không có kỳ hạn phù hợp (tuổi vượt quá 65)</option>';
        optionsHtml += `
            <div>
                <label for="abuv-term" class="font-medium text-gray-700 block mb-1">Thời hạn đóng phí</label>
                <select id="abuv-term" class="form-select">${termOptions}</select>
                <p class="text-sm text-gray-500 mt-1">Thời hạn đóng phí bằng thời hạn hợp đồng.</p>
            </div>`;
    } else if (['KHOE_BINH_AN', 'VUNG_TUONG_LAI', 'PUL_TRON_DOI', 'PUL_15_NAM', 'PUL_5_NAM'].includes(mainProduct)) {
        optionsHtml = `
            <div>
                <label for="main-stbh" class="font-medium text-gray-700 block mb-1">Số tiền bảo hiểm (STBH)</label>
                <input type="text" id="main-stbh" class="form-input" value="${currentStbh}" placeholder="VD: 1.000.000.000">
            </div>`;

        if (['KHOE_BINH_AN', 'VUNG_TUONG_LAI'].includes(mainProduct)) {
            optionsHtml += `
                <div>
                    <label for="main-premium-input" class="font-medium text-gray-700 block mb-1">Phí sản phẩm chính</label>
                    <input type="text" id="main-premium-input" class="form-input" value="${currentPremium}" placeholder="Nhập phí">
                    <div id="mul-fee-range" class="text-sm text-gray-500 mt-1"></div>
                </div>`;
        }

        const { min, max } = getPaymentTermBoundsByProduct(mainProduct, age);
        optionsHtml += `
            <div>
                <label for="payment-term" class="font-medium text-gray-700 block mb-1">Thời gian đóng phí (năm)</label>
                <input type="number" id="payment-term" class="form-input" value="${currentPaymentTerm}" placeholder="VD: 20" min="${min}" max="${max}">
                <div id="payment-term-hint" class="text-sm text-gray-500 mt-1"></div>
            </div>`;

        optionsHtml += `
            <div>
                <label for="extra-premium-input" class="font-medium text-gray-700 block mb-1">Phí đóng thêm</label>
                <input type="text" id="extra-premium-input" class="form-input" value="${currentExtra || ''}" placeholder="VD: 10.000.000">
                <div class="text-sm text-gray-500 mt-1">Tối đa 5 lần phí chính.</div>
            </div>`;
    }

    container.innerHTML = optionsHtml;

    if (['KHOE_BINH_AN', 'VUNG_TUONG_LAI', 'PUL_TRON_DOI', 'PUL_15_NAM', 'PUL_5_NAM'].includes(mainProduct)) {
        setPaymentTermHint(mainProduct, age);
    }
}

function calculateMainPremium(customer, ageOverride = null) {
    const ageToUse = ageOverride ?? customer.age;
    const { gender, mainProduct } = customer;
    let premium = 0;

    if (mainProduct.startsWith('PUL') || mainProduct === 'AN_BINH_UU_VIET' || mainProduct === 'TRON_TAM_AN') {
        let stbh = 0;
        let rate = 0;
        const stbhEl = document.getElementById('main-stbh');
        if (stbhEl) stbh = parseFormattedNumber(stbhEl.value);

        if (stbh === 0 && mainProduct !== 'TRON_TAM_AN') {
            return 0;
        }

        const genderKey = gender === 'Nữ' ? 'nu' : 'nam';

        if (mainProduct.startsWith('PUL')) {
            const pulRate = product_data.pul_rates[mainProduct]?.find(r => r.age === customer.age)?.[genderKey] || 0;
            if (pulRate === 0 && !ageOverride) throw new Error(`Không có biểu phí PUL cho tuổi ${customer.age}.`);
            rate = pulRate;

            premium = (stbh / 1000) * rate;

            if (!ageOverride && stbh > 0 && stbh < 100000000) {
                setFieldError(stbhEl, 'STBH nhỏ hơn 100 triệu');
                throw new Error('STBH nhỏ hơn 100 triệu');
            } else {
                clearFieldError(stbhEl);
            }

            if (!ageOverride && premium > 0 && premium < 5000000) {
                setFieldError(stbhEl, 'Phí chính nhỏ hơn 5 triệu');
                throw new Error('Phí chính nhỏ hơn 5 triệu');
            }
        } else if (mainProduct === 'AN_BINH_UU_VIET') {
            const term = document.getElementById('abuv-term')?.value;
            if (!term) return 0;
            const abuvRate = product_data.an_binh_uu_viet_rates[term]?.find(r => r.age === customer.age)?.[genderKey] || 0;
            if (abuvRate === 0 && !ageOverride) {
                throw new Error(`Không có biểu phí An Bình Ưu Việt cho tuổi ${customer.age}, kỳ hạn ${term} năm.`);
            }
            rate = abuvRate;
            premium = (stbh / 1000) * rate;

            const stbhEl2 = document.getElementById('main-stbh');
            if (!ageOverride) {
                if (stbh > 0 && stbh < 100000000) {
                    setFieldError(stbhEl2, 'STBH nhỏ hơn 100 triệu');
                    throw new Error('STBH nhỏ hơn 100 triệu');
                } else {
                    clearFieldError(stbhEl2);
                }
                if (premium > 0 && premium < 5000000) {
                    setFieldError(stbhEl2, 'Phí chính nhỏ hơn 5 triệu');
                    throw new Error('Phí chính nhỏ hơn 5 triệu');
                }
            }
        } else if (mainProduct === 'TRON_TAM_AN') {
            stbh = 100000000;
            const term = '10';
            const ttaRate = product_data.an_binh_uu_viet_rates[term]?.find(r => r.age === customer.age)?.[genderKey] || 0;
            if (ttaRate === 0 && !ageOverride) {
                throw new Error(`Không có biểu phí Trọn Tâm An cho tuổi ${customer.age}.`);
            }
            rate = ttaRate;
            premium = (stbh / 1000) * rate;
        }
    } else if (['KHOE_BINH_AN', 'VUNG_TUONG_LAI'].includes(mainProduct)) {
        const stbh = parseFormattedNumber(document.getElementById('main-stbh')?.value || '0');

        if (!ageOverride && stbh > 0 && stbh < 100000000) {
            setFieldError(document.getElementById('main-stbh'), 'STBH nhỏ hơn 100 triệu');
            throw new Error('STBH nhỏ hơn 100 triệu');
        } else {
            clearFieldError(document.getElementById('main-stbh'));
        }

        const factorRow = product_data.mul_factors.find(f => ageToUse >= f.ageMin && ageToUse <= f.ageMax);
        if (!factorRow) throw new Error(`Không có hệ số MUL cho tuổi ${ageToUse}.`);

        const minFee = stbh / factorRow.maxFactor;
        const maxFee = stbh / factorRow.minFactor;
        const rangeEl = document.getElementById('mul-fee-range');
        if (!ageOverride && rangeEl) {
            rangeEl.textContent = `Phí hợp lệ từ ${formatCurrency(minFee, '')} đến ${formatCurrency(maxFee, '')}.`;
        }

        const enteredPremium = parseFormattedNumber(document.getElementById('main-premium-input')?.value || '0');
        if (!ageOverride) {
            const feeInput = document.getElementById('main-premium-input');
            let feeInvalid = false;
            if (stbh > 0 && enteredPremium > 0) {
                if (enteredPremium < minFee || enteredPremium > maxFee) feeInvalid = true;
                if (enteredPremium < 5000000) feeInvalid = true;
            }
            if (feeInvalid) {
                setFieldError(feeInput, 'Phí không hợp lệ');
                throw new Error('Phí không hợp lệ');
            } else {
                clearFieldError(feeInput);
            }
        }

        premium = enteredPremium;
    }

    return premium;
}

/* ---------- Section 3: SP bổ sung (tính phí) ---------- */
function calculateHealthSclPremium(customer, container, ageOverride = null) {
    const section = container.querySelector('.health-scl-section');
    if (!section || !section.querySelector('.health-scl-checkbox')?.checked) {
        if (section && !ageOverride) section.querySelector('.fee-display').textContent = '';
        return 0;
    }
    const ageToUse = ageOverride ?? customer.age;
    if (ageToUse > MAX_RENEWAL_AGE.health_scl) return 0;

    const program = section.querySelector('.health-scl-program').value;
    const scope = section.querySelector('.health-scl-scope').value;
    const hasOutpatient = section.querySelector('.health-scl-outpatient').checked;
    const hasDental = section.querySelector('.health-scl-dental').checked;

    const ageBandIndex = product_data.health_scl_rates.age_bands.findIndex(b => ageToUse >= b.min && ageToUse <= b.max);
    if (ageBandIndex === -1) return 0;

    let totalPremium = 0;
    totalPremium += product_data.health_scl_rates[scope]?.[ageBandIndex]?.[program] || 0;
    if (hasOutpatient) totalPremium += product_data.health_scl_rates.outpatient?.[ageBandIndex]?.[program] || 0;
    if (hasDental) totalPremium += product_data.health_scl_rates.dental?.[ageBandIndex]?.[program] || 0;

    if (!ageOverride) section.querySelector('.fee-display').textContent = totalPremium > 0 ? `Phí: ${formatCurrency(totalPremium)}` : '';
    return totalPremium;
}

function calculateBhnPremium(customer, container, ageOverride = null) {
    const section = container.querySelector('.bhn-section');
    if (!section || !section.querySelector('.bhn-checkbox')?.checked) {
        if (section && !ageOverride) section.querySelector('.fee-display').textContent = '';
        return 0;
    }
    const ageToUse = ageOverride ?? customer.age;
    if (ageToUse > MAX_RENEWAL_AGE.bhn) return 0;

    const { gender } = customer;
    const stbhInput = section.querySelector('.bhn-stbh');
    const stbh = parseFormattedNumber(stbhInput?.value || '0');
    if (stbh === 0) {
        if (!ageOverride) section.querySelector('.fee-display').textContent = '';
        return 0;
    }
    if (stbh < 100000000 || stbh > MAX_STBH.bhn) {
        setFieldError(stbhInput, 'STBH không hợp lệ, từ 100 triệu đến 5 tỷ');
        throw new Error('STBH không hợp lệ, từ 100 triệu đến 5 tỷ');
    } else {
        clearFieldError(stbhInput);
    }

    const rate = product_data.bhn_rates.find(r => ageToUse >= r.ageMin && ageToUse <= r.ageMax)?.[gender === 'Nữ' ? 'nu' : 'nam'] || 0;
    const premium = (stbh / 1000) * rate;
    if (!ageOverride) section.querySelector('.fee-display').textContent = `Phí: ${formatCurrency(premium)}`;
    return premium;
}

function calculateAccidentPremium(customer, container, ageOverride = null) {
    const section = container.querySelector('.accident-section');
    if (!section || !section.querySelector('.accident-checkbox')?.checked) {
        if (section && !ageOverride) section.querySelector('.fee-display').textContent = '';
        return 0;
    }
    const ageToUse = ageOverride ?? customer.age;
    if (ageToUse > MAX_RENEWAL_AGE.accident) return 0;

    const { riskGroup } = customer;
    if (riskGroup === 0) return 0;
    const stbhInput = section.querySelector('.accident-stbh');
    const stbh = parseFormattedNumber(stbhInput?.value || '0');
    if (stbh === 0) {
        if (!ageOverride) section.querySelector('.fee-display').textContent = '';
        return 0;
    }
    if (stbh < 100000000 || stbh > MAX_STBH.accident) {
        setFieldError(stbhInput, 'STBH không hợp lệ, từ 100 triệu đến 8 tỷ');
        throw new Error('STBH không hợp lệ, từ 100 triệu đến 8 tỷ');
    } else {
        clearFieldError(stbhInput);
    }

    const rate = product_data.accident_rates[riskGroup] || 0;
    const premium = (stbh / 1000) * rate;
    if (!ageOverride) section.querySelector('.fee-display').textContent = `Phí: ${formatCurrency(premium)}`;
    return premium;
}

function calculateHospitalSupportPremium(customer, mainPremium, container, totalHospitalSupportStbh = 0, ageOverride = null) {
    const section = container.querySelector('.hospital-support-section');
    if (!section || !section.querySelector('.hospital-support-checkbox')?.checked) {
        if (section && !ageOverride) section.querySelector('.fee-display').textContent = '';
        return 0;
    }
    const ageToUse = ageOverride ?? customer.age;
    if (ageToUse > MAX_RENEWAL_AGE.hospital_support) return 0;

    const hsInput = section.querySelector('.hospital-support-stbh');

    const totalMaxSupport = Math.floor(mainPremium / 4000000) * 100000;
    const maxSupportByAge = ageToUse >= 18 ? 1_000_000 : 300_000;
    const remainingSupport = totalMaxSupport - totalHospitalSupportStbh;

    if (!ageOverride) {
        section.querySelector('.hospital-support-validation').textContent =
            `Tối đa: ${formatCurrency(Math.min(maxSupportByAge, remainingSupport), 'đ/ngày')}. Phải là bội số của 100.000.`;
    }

    const stbh = parseFormattedNumber(hsInput?.value || '0');
    if (stbh === 0) {
        if (!ageOverride) section.querySelector('.fee-display').textContent = '';
        clearFieldError(hsInput);
        return 0;
    }

    if (stbh % 100000 !== 0) {
        setFieldError(hsInput, 'STBH không hợp lệ, phải là bội số 100.000');
        throw new Error('STBH không hợp lệ, phải là bội số 100.000');
    }
    if (stbh > maxSupportByAge || stbh > remainingSupport) {
        setFieldError(hsInput, 'Vượt quá giới hạn cho phép');
        throw new Error('Vượt quá giới hạn cho phép');
    }
    clearFieldError(hsInput);

    const rate = product_data.hospital_fee_support_rates.find(r => ageToUse >= r.ageMin && ageToUse <= r.ageMax)?.rate || 0;
    const premium = (stbh / 100) * rate;
    if (!ageOverride) section.querySelector('.fee-display').textContent = `Phí: ${formatCurrency(premium)}`;
    return premium;
}

/* ---------- Tổng hợp phí hiển thị ---------- */
function updateSummaryUI(premiums) {
    document.getElementById('main-premium-result').textContent = formatCurrency(premiums.mainPremium);

    const suppContainer = document.getElementById('supplementary-premiums-results');
    suppContainer.innerHTML = '';
    if (premiums.totalSupplementaryPremium > 0) {
        suppContainer.innerHTML = `<div class="flex justify-between items-center py-2 border-b"><span class="text-gray-600">Tổng phí SP bổ sung:</span><span class="font-bold text-gray-900">${formatCurrency(premiums.totalSupplementaryPremium)}</span></div>`;
    }

    document.getElementById('total-premium-result').textContent = formatCurrency(premiums.totalPremium);
}

/* ---------- Bảng Minh Hoạ ---------- */
function generateSummaryTable() {
    const modal = document.getElementById('summary-modal');
    const container = document.getElementById('summary-content-container');
    container.innerHTML = '';

    try {
        const targetAgeInput = document.getElementById('target-age-input');
        const targetAge = parseInt(targetAgeInput.value, 10);
        const mainPersonContainer = document.getElementById('main-person-container');
        const mainPersonInfo = getCustomerInfo(mainPersonContainer, true);
        const mainProduct = mainPersonInfo.mainProduct;

        if (isNaN(targetAge) || targetAge <= mainPersonInfo.age) {
            throw new Error("Vui lòng nhập một độ tuổi mục tiêu hợp lệ, lớn hơn tuổi hiện tại của NĐBH chính.");
        }

        if (mainProduct === 'TRON_TAM_AN') {
            const mainSuppContainer = document.querySelector('#main-supp-container .supplementary-products-container');
            const healthSclSection = mainSuppContainer?.querySelector('.health-scl-section');
            const healthSclCheckbox = healthSclSection?.querySelector('.health-scl-checkbox');
            const healthSclPremium = calculateHealthSclPremium(mainPersonInfo, mainSuppContainer);
            if (!healthSclCheckbox?.checked || healthSclPremium === 0) {
                throw new Error('Sản phẩm Trọn Tâm An bắt buộc phải tham gia kèm Sức Khỏe Bùng Gia Lực với phí hợp lệ.');
            }
        }

        let paymentTerm = 999;
        const paymentTermInput = document.getElementById('payment-term');
        if (paymentTermInput) {
            paymentTerm = parseInt(paymentTermInput.value, 10) || 999;
        } else if (mainPersonInfo.mainProduct === 'AN_BINH_UU_VIET') {
            paymentTerm = parseInt(document.getElementById('abuv-term')?.value, 10);
        } else if (mainPersonInfo.mainProduct === 'TRON_TAM_AN') {
            paymentTerm = 10;
        }

        if (['PUL_TRON_DOI', 'PUL_5_NAM', 'PUL_15_NAM', 'KHOE_BINH_AN', 'VUNG_TUONG_LAI'].includes(mainPersonInfo.mainProduct) && targetAge < mainPersonInfo.age + paymentTerm - 1) {
            throw new Error(`Độ tuổi mục tiêu phải lớn hơn hoặc bằng ${mainPersonInfo.age + paymentTerm - 1} đối với ${mainPersonInfo.mainProduct}.`);
        }

        // Thu thập NĐBH bổ sung
        const suppPersons = [];
        document.querySelectorAll('.person-container').forEach(pContainer => {
            if (pContainer.id !== 'main-person-container') {
                const personInfo = getCustomerInfo(pContainer, false);
                suppPersons.push(personInfo);
            }
        });

        // Trạng thái Miễn đóng phí 3.0
        const waiverCont = document.getElementById('waiver30-container');
        const waiverSelected = waiverCont ? waiverCont.querySelector('input[name="waiver30-insurer"]:checked')?.value || 'other' : null;
        const otherInfo = waiverSelected === 'other' ? getWaiverOtherInfo() : null;

        let tableHtml = `<table class="w-full text-left border-collapse"><thead class="bg-gray-100"><tr>`;
        tableHtml += `<th class="p-2 border">Năm HĐ</th>`;
        tableHtml += `<th class="p-2 border">Tuổi NĐBH Chính<br>(${sanitizeHtml(mainPersonInfo.name)})</th>`;
        tableHtml += `<th class="p-2 border">Phí SP Chính<br>(${sanitizeHtml(mainPersonInfo.name)})</th>`;
        tableHtml += `<th class="p-2 border">Phí SP Bổ Sung<br>(${sanitizeHtml(mainPersonInfo.name)})</th>`;
        suppPersons.forEach(person => {
            tableHtml += `<th class="p-2 border">Phí SP Bổ Sung<br>(${sanitizeHtml(person.name)})</th>`;
        });
        if (waiverSelected === 'other') {
            tableHtml += `<th class="p-2 border">Miễn đóng phí 3.0<br>(Người khác)</th>`;
        }
        tableHtml += `<th class="p-2 border">Tổng Phí Năm</th>`;
        tableHtml += `</tr></thead><tbody>`;

        let totalMainAcc = 0;
        let totalSuppAccMain = 0;
        let totalSuppAccAll = 0;
        let totalWaiverAcc = 0;

        const initialBaseMainPremium = calculateMainPremium(mainPersonInfo);
        const extraPremium = getExtraPremiumValue();
        const initialMainPremiumWithExtra = initialBaseMainPremium + extraPremium;

        const totalMaxSupport = Math.floor(initialBaseMainPremium / 4000000) * 100000;

        // Tính phí Miễn Đóng Phí 3.0 (cố định theo năm)
        const waiverPerYear = calculateWaiver30Premium(mainPersonInfo, initialMainPremiumWithExtra);
        let waiverStartAge = 0; // tuổi người được chọn
        if (waiverSelected && waiverPerYear > 0) {
            if (waiverSelected === 'other') {
                waiverStartAge = otherInfo?.age || 0;
            } else {
                const pEl = document.getElementById(waiverSelected);
                const pInfo = pEl ? getCustomerInfo(pEl, false) : null;
                waiverStartAge = pInfo?.age || 0;
            }
        }

        for (let i = 0; (mainPersonInfo.age + i) <= targetAge; i++) {
            const currentAgeMain = mainPersonInfo.age + i;
            const contractYear = i + 1;

            const mainPremiumForYear = (contractYear <= paymentTerm) ? initialMainPremiumWithExtra : 0;
            totalMainAcc += mainPremiumForYear;

            let suppPremiumMain = 0;
            let totalHospitalSupportStbh = 0;
            const mainSuppContainer = document.querySelector('#main-supp-container .supplementary-products-container');
            if (mainSuppContainer) {
                suppPremiumMain += calculateHealthSclPremium({ ...mainPersonInfo, age: currentAgeMain }, mainSuppContainer, currentAgeMain);
                suppPremiumMain += calculateBhnPremium({ ...mainPersonInfo, age: currentAgeMain }, mainSuppContainer, currentAgeMain);
                suppPremiumMain += calculateAccidentPremium({ ...mainPersonInfo, age: currentAgeMain }, mainSuppContainer, currentAgeMain);
                suppPremiumMain += calculateHospitalSupportPremium({ ...mainPersonInfo, age: currentAgeMain }, initialBaseMainPremium, mainSuppContainer, totalHospitalSupportStbh, currentAgeMain);
                const hospitalSupportStbh = parseFormattedNumber(mainSuppContainer.querySelector('.hospital-support-stbh')?.value || '0');
                if (mainSuppContainer.querySelector('.hospital-support-checkbox')?.checked && hospitalSupportStbh > 0) {
                    totalHospitalSupportStbh += hospitalSupportStbh;
                }
            }

            // Waiver năm này (áp dụng tới 65 tuổi của người chọn)
            let waiverThisYear = 0;
            if (waiverPerYear > 0 && waiverStartAge > 0 && (waiverStartAge + i) <= WAIVER30.MAX_RENEWAL_AGE) {
                waiverThisYear = waiverPerYear;
            }

            totalSuppAccMain += suppPremiumMain;

            // Phí SP BS của từng NĐBH BS (cộng waiver nếu chọn người đó)
            const suppPremiums = suppPersons.map(person => {
                const currentPersonAge = person.age + i;
                const suppProductsContainer = person.container.querySelector('.supplementary-products-container');
                let suppPremium = 0;
                if (suppProductsContainer) {
                    suppPremium += calculateHealthSclPremium({ ...person, age: currentPersonAge }, suppProductsContainer, currentPersonAge);
                    suppPremium += calculateBhnPremium({ ...person, age: currentPersonAge }, suppProductsContainer, currentPersonAge);
                    suppPremium += calculateAccidentPremium({ ...person, age: currentPersonAge }, suppProductsContainer, currentPersonAge);
                    suppPremium += calculateHospitalSupportPremium({ ...person, age: currentPersonAge }, initialBaseMainPremium, suppProductsContainer, totalHospitalSupportStbh, currentPersonAge);
                    const hospitalSupportStbh = parseFormattedNumber(suppProductsContainer.querySelector('.hospital-support-stbh')?.value || '0');
                    if (suppProductsContainer.querySelector('.hospital-support-checkbox')?.checked && hospitalSupportStbh > 0) {
                        totalHospitalSupportStbh += hospitalSupportStbh;
                    }
                }
                // Nếu chọn người này cho Waiver, cộng vào suppPremium của họ
                if (waiverSelected === person.container.id && waiverThisYear > 0) {
                    suppPremium += waiverThisYear;
                }
                totalSuppAccAll += suppPremium;
                return suppPremium;
            });

            if (totalHospitalSupportStbh > totalMaxSupport) {
                throw new Error(`Tổng số tiền Hỗ trợ viện phí vượt quá hạn mức chung: ${formatCurrency(totalMaxSupport, 'đ/ngày')}.`);
            }

            tableHtml += `<tr>
                <td class="p-2 border text-center">${contractYear}</td>
                <td class="p-2 border text-center">${currentAgeMain}</td>
                <td class="p-2 border text-right">${formatCurrency(mainPremiumForYear)}</td>
                <td class="p-2 border text-right">${formatCurrency(suppPremiumMain)}</td>`;
            suppPremiums.forEach(suppPremium => {
                tableHtml += `<td class="p-2 border text-right">${formatCurrency(suppPremium)}</td>`;
            });

            // Nếu chọn "Người khác", thêm cột riêng
            if (waiverSelected === 'other') {
                tableHtml += `<td class="p-2 border text-right">${formatCurrency(waiverThisYear)}</td>`;
                totalWaiverAcc += waiverThisYear;
            }

            const rowTotal = mainPremiumForYear + suppPremiumMain + suppPremiums.reduce((sum, p) => sum + p, 0) + (waiverSelected === 'other' ? waiverThisYear : 0);
            tableHtml += `<td class="p-2 border text-right font-semibold">${formatCurrency(rowTotal)}</td>`;
            tableHtml += `</tr>`;
        }

        // Tổng cộng
        tableHtml += `<tr class="bg-gray-200 font-bold"><td class="p-2 border" colspan="2">Tổng cộng</td>`;
        tableHtml += `<td class="p-2 border text-right">${formatCurrency(totalMainAcc)}</td>`;
        tableHtml += `<td class="p-2 border text-right">${formatCurrency(totalSuppAccMain)}</td>`;
        suppPersons.forEach((_) => {
            // Tổng cộng của mỗi NĐBH bổ sung đã cộng trong totalSuppAccAll
            tableHtml += `<td class="p-2 border text-right">—</td>`;
        });
        if (waiverSelected === 'other') {
            tableHtml += `<td class="p-2 border text-right">${formatCurrency(totalWaiverAcc)}</td>`;
        }
        const grandTotal = totalMainAcc + totalSuppAccMain + totalSuppAccAll + (waiverSelected === 'other' ? totalWaiverAcc : 0);
        tableHtml += `<td class="p-2 border text-right">${formatCurrency(grandTotal)}</td>`;
        tableHtml += `</tr></tbody></table>`;
        tableHtml += `<div class="mt-4 text-center"><button id="export-html-btn" class="bg-blue-500 text-white px-4 py-2 rounded hover:bg-blue-600">Xuất HTML</button></div>`;
        container.innerHTML = tableHtml;

        document.getElementById('export-html-btn').addEventListener('click', () => exportToHTML(mainPersonInfo, suppPersons, targetAge, initialBaseMainPremium + extraPremium, paymentTerm));

    } catch (e) {
        container.innerHTML = `<p class="text-red-600 font-semibold text-center">${e.message}</p>`;
    } finally {
        modal.classList.remove('hidden');
    }
}

/* ---------- Export HTML ---------- */
function sanitizeHtml(str) {
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function exportToHTML(mainPersonInfo, suppPersons, targetAge, initialMainPremiumWithExtra, paymentTerm) {
    // Gộp xuất cơ bản (không tách riêng waiver ở export để tránh dài, có thể nâng cấp nếu cần)
    const initialBaseMainPremium = calculateMainPremium(mainPersonInfo);
    const totalMaxSupport = Math.floor(initialBaseMainPremium / 4000000) * 100000;

    let tableHtml = `
        <table style="width: 100%; border-collapse: collapse; font-family: Arial, sans-serif;">
            <thead style="background-color: #f3f4f6;">
                <tr>
                    <th style="padding: 8px; border: 1px solid #d1d5db; text-align: center;">Năm HĐ</th>
                    <th style="padding: 8px; border: 1px solid #d1d5db; text-align: center;">Tuổi NĐBH Chính<br>(${sanitizeHtml(mainPersonInfo.name)})</th>
                    <th style="padding: 8px; border: 1px solid #d1d5db; text-align: right;">Phí SP Chính<br>(${sanitizeHtml(mainPersonInfo.name)})</th>
                    <th style="padding: 8px; border: 1px solid #d1d5db; text-align: right;">Phí SP Bổ Sung<br>(${sanitizeHtml(mainPersonInfo.name)})</th>
                    ${suppPersons.map(person => `<th style="padding: 8px; border: 1px solid #d1d5db; text-align: right;">Phí SP Bổ Sung<br>(${sanitizeHtml(person.name)})</th>`).join('')}
                    <th style="padding: 8px; border: 1px solid #d1d5db; text-align: right;">Tổng Phí Năm</th>
                </tr>
            </thead>
            <tbody>
    `;

    let totalMainAcc = 0;
    let totalSuppAccMain = 0;
    let totalSuppAccAll = 0;

    for (let i = 0; (mainPersonInfo.age + i) <= targetAge; i++) {
        const currentAgeMain = mainPersonInfo.age + i;
        const contractYear = i + 1;

        const mainPremiumForYear = (contractYear <= paymentTerm) ? initialMainPremiumWithExtra : 0;
        totalMainAcc += mainPremiumForYear;

        let suppPremiumMain = 0;
        let totalHospitalSupportStbh = 0;
        const mainSuppContainer = document.querySelector('#main-supp-container .supplementary-products-container');
        if (mainSuppContainer) {
            suppPremiumMain += calculateHealthSclPremium({ ...mainPersonInfo, age: currentAgeMain }, mainSuppContainer, currentAgeMain);
            suppPremiumMain += calculateBhnPremium({ ...mainPersonInfo, age: currentAgeMain }, mainSuppContainer, currentAgeMain);
            suppPremiumMain += calculateAccidentPremium({ ...mainPersonInfo, age: currentAgeMain }, mainSuppContainer, currentAgeMain);
            suppPremiumMain += calculateHospitalSupportPremium({ ...mainPersonInfo, age: currentAgeMain }, initialBaseMainPremium, mainSuppContainer, totalHospitalSupportStbh, currentAgeMain);
            const hospitalSupportStbh = parseFormattedNumber(mainSuppContainer.querySelector('.hospital-support-stbh')?.value || '0');
            if (mainSuppContainer.querySelector('.hospital-support-checkbox')?.checked && hospitalSupportStbh > 0) {
                totalHospitalSupportStbh += hospitalSupportStbh;
            }
        }
        totalSuppAccMain += suppPremiumMain;

        const suppPremiums = suppPersons.map(person => {
            const currentPersonAge = person.age + i;
            const suppProductsContainer = person.container.querySelector('.supplementary-products-container');
            let suppPremium = 0;
            if (suppProductsContainer) {
                suppPremium += calculateHealthSclPremium({ ...person, age: currentPersonAge }, suppProductsContainer, currentPersonAge);
                suppPremium += calculateBhnPremium({ ...person, age: currentPersonAge }, suppProductsContainer, currentPersonAge);
                suppPremium += calculateAccidentPremium({ ...person, age: currentPersonAge }, suppProductsContainer, currentPersonAge);
                suppPremium += calculateHospitalSupportPremium({ ...person, age: currentPersonAge }, initialBaseMainPremium, suppProductsContainer, totalHospitalSupportStbh, currentPersonAge);
                const hospitalSupportStbh = parseFormattedNumber(suppProductsContainer.querySelector('.hospital-support-stbh')?.value || '0');
                if (suppProductsContainer.querySelector('.hospital-support-checkbox')?.checked && hospitalSupportStbh > 0) {
                    totalHospitalSupportStbh += hospitalSupportStbh;
                }
            }
            totalSuppAccAll += suppPremium;
            return suppPremium;
        });

        const rowTotal = mainPremiumForYear + suppPremiumMain + suppPremiums.reduce((sum, p) => sum + p, 0);

        tableHtml += `
            <tr>
                <td style="padding: 8px; border: 1px solid #d1d5db; text-align: center;">${contractYear}</td>
                <td style="padding: 8px; border: 1px solid #d1d5db; text-align: center;">${currentAgeMain}</td>
                <td style="padding: 8px; border: 1px solid #d1d5db; text-align: right;">${formatCurrency(mainPremiumForYear)}</td>
                <td style="padding: 8px; border: 1px solid #d1d5db; text-align: right;">${formatCurrency(suppPremiumMain)}</td>
                ${suppPremiums.map(suppPremium => `<td style="padding: 8px; border: 1px solid #d1d5db; text-align: right;">${formatCurrency(suppPremium)}</td>`).join('')}
                <td style="padding: 8px; border: 1px solid #d1d5db; text-align: right; font-weight: 600;">${formatCurrency(rowTotal)}</td>
            </tr>
        `;
    }

    tableHtml += `
        <tr style="background-color: #e5e7eb; font-weight: bold;">
            <td style="padding: 8px; border: 1px solid #d1d5db;" colspan="2">Tổng cộng</td>
            <td style="padding: 8px; border: 1px solid #d1d5db; text-align: right;">${formatCurrency(totalMainAcc)}</td>
            <td style="padding: 8px; border: 1px solid #d1d5db; text-align: right;">${formatCurrency(totalSuppAccMain)}</td>
            ${suppPersons.map(() => `<td style="padding: 8px; border: 1px solid #d1d5db; text-align: right;"></td>`).join('')}
            <td style="padding: 8px; border: 1px solid #d1d5db; text-align: right;">${formatCurrency(totalMainAcc + totalSuppAccMain + totalSuppAccAll)}</td>
        </tr>
    </tbody></table>`;

    const htmlContent = `
<!DOCTYPE html>
<html lang="vi">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Bảng Minh Họa Phí Bảo Hiểm</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 40px; }
        h1 { text-align: center; color: #1f2937; }
        table { width: 100%; border-collapse: collapse; margin-top: 20px; }
        th, td { padding: 8px; border: 1px solid #d1d5db; }
        th { background-color: #f3f4f6; }
        tr:nth-child(even) { background-color: #f9fafb; }
        .text-right { text-align: right; }
        .text-center { text-align: center; }
        .font-bold { font-weight: bold; }
        @media print {
            body { margin: 0; }
            .no-print { display: none; }
        }
    </style>
</head>
<body>
    <h1>Bảng Minh Họa Phí Bảo Hiểm</h1>
    ${tableHtml}
    <div style="margin-top: 20px; text-align: center;" class="no-print">
        <button onclick="window.print()" style="background-color: #3b82f6; color: white; padding: 8px 16px; border-radius: 4px; border: none; cursor: pointer;">In thành PDF</button>
    </div>
</body>
</html>
    `;

    const blob = new Blob([htmlContent], { type: 'text/html' });
    const url = window.URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'bang_minh_hoa_phi_bao_hiem.html';
    document.body.appendChild(a);
    a.click();
    window.URL.revokeObjectURL(url);
    document.body.removeChild(a);
}

/* ---------- Utils ---------- */
function formatCurrency(value, suffix = ' VNĐ') {
    if (isNaN(value)) return '0' + suffix;
    return Math.round(value).toLocaleString('vi-VN') + suffix;
}
function formatNumberInput(input) {
    if (!input || !input.value) return;
    let value = input.value.replace(/[.,]/g, '');
    if (!isNaN(value) && value.length > 0) {
        input.value = parseInt(value, 10).toLocaleString('vi-VN');
    } else if (input.value !== '') {
        input.value = '';
    }
}
function parseFormattedNumber(formattedString) {
    return parseInt(String(formattedString).replace(/[.,]/g, ''), 10) || 0;
}
function showError(message) {
    document.getElementById('error-message').textContent = message;
}
function clearError() {
    document.getElementById('error-message').textContent = '';
}

// Helpers: hiển thị lỗi trường cho Section 1
function setFieldError(inputOrEl, message) {
    const input = typeof inputOrEl === 'string' ? document.querySelector(inputOrEl) : inputOrEl;
    if (!input) return;
    let err = input.parentElement.querySelector('.field-error');
    if (!err) {
        err = document.createElement('p');
        err.className = 'field-error text-sm text-red-600 mt-1';
        input.parentElement.appendChild(err);
    }
    err.textContent = message || '';
    if (message) {
        input.classList.add('border-red-500');
    } else {
        input.classList.remove('border-red-500');
    }
}
function clearFieldError(inputOrEl) {
    setFieldError(inputOrEl, '');
}

function validateMainPersonInputs() {
    const container = document.getElementById('main-person-container');
    if (!container) return true;

    const nameInput = container.querySelector('.name-input');
    const dobInput = container.querySelector('.dob-input');
    const occupationInput = container.querySelector('.occupation-input');

    let ok = true;

    if (nameInput) {
        const v = (nameInput.value || '').trim();
        if (!v) {
            setFieldError(nameInput, 'Vui lòng nhập họ và tên');
            ok = false;
        } else {
            clearFieldError(nameInput);
        }
    }

    if (dobInput) {
        const v = (dobInput.value || '').trim();
        const re = /^\d{2}\/\d{2}\/\d{4}$/;
        if (!re.test(v)) {
            setFieldError(dobInput, 'Ngày sinh không hợp lệ, nhập DD/MM/YYYY');
            ok = false;
        } else {
            const [dd, mm, yyyy] = v.split('/').map(n => parseInt(n, 10));
            const d = new Date(yyyy, mm - 1, dd);
            const valid = d.getFullYear() === yyyy && d.getMonth() === (mm - 1) && d.getDate() === dd && d <= REFERENCE_DATE;
            if (!valid) {
                setFieldError(dobInput, 'Ngày sinh không hợp lệ, nhập DD/MM/YYYY');
                ok = false;
            } else {
                clearFieldError(dobInput);
            }
        }
    }

    if (occupationInput) {
        const typed = (occupationInput.value || '').trim().toLowerCase();
        const match = product_data.occupations.find(o => o.group > 0 && o.name.toLowerCase() === typed);
        const group = parseInt(occupationInput.dataset.group, 10);
        if (!match || !(group >= 1 && group <= 4)) {
            setFieldError(occupationInput, 'Chọn nghề nghiệp từ danh sách');
            ok = false;
        } else {
            clearFieldError(occupationInput);
        }
    }

    return ok;
}

// Section 2 helpers
function getProductMinPaymentTerm(mainProduct) {
    if (mainProduct === 'PUL_5_NAM') return 5;
    if (mainProduct === 'PUL_15_NAM') return 15;
    return 4;
}
function getPaymentTermBounds(age) {
    const min = 4;
    const max = Math.max(0, 100 - age - 1);
    return { min, max };
}
function getPaymentTermBoundsByProduct(mainProduct, age) {
    const prodMin = getProductMinPaymentTerm(mainProduct);
    const max = Math.max(0, 100 - age - 1);
    return { min: prodMin, max };
}
function setPaymentTermHint(mainProduct, age) {
    const hintEl = document.getElementById('payment-term-hint');
    if (!hintEl) return;
    const { min, max } = getPaymentTermBoundsByProduct(mainProduct, age);
    let hint = `Nhập từ ${min} đến ${max} năm`;
    hintEl.textContent = hint;
}
function validateSection2FieldsPreCalc(customer) {
    const mainProduct = customer.mainProduct;

    if (mainProduct && mainProduct !== 'TRON_TAM_AN') {
        const stbhEl = document.getElementById('main-stbh');
        if (stbhEl) {
            const stbh = parseFormattedNumber(stbhEl.value || '0');
            if (stbh > 0 && stbh < 100000000) {
                setFieldError(stbhEl, 'STBH nhỏ hơn 100 triệu');
            } else {
                clearFieldError(stbhEl);
            }
        }
    }

    if (['KHOE_BINH_AN', 'VUNG_TUONG_LAI', 'PUL_TRON_DOI', 'PUL_15_NAM', 'PUL_5_NAM'].includes(mainProduct)) {
        const el = document.getElementById('payment-term');
        if (el) {
            const { min, max } = getPaymentTermBoundsByProduct(mainProduct, customer.age);
            const val = parseInt(el.value, 10);
            if (el.value && (isNaN(val) || val < min || val > max)) {
                setFieldError(el, `Thời hạn không hợp lệ, từ ${min} đến ${max}`);
            } else {
                clearFieldError(el);
            }
        }
    }

    if (['KHOE_BINH_AN', 'VUNG_TUONG_LAI'].includes(mainProduct)) {
        const stbh = parseFormattedNumber(document.getElementById('main-stbh')?.value || '0');
        const feeInput = document.getElementById('main-premium-input');
        const factorRow = product_data.mul_factors.find(f => customer.age >= f.ageMin && customer.age <= f.ageMax);
        if (factorRow && stbh > 0) {
            const minFee = stbh / factorRow.maxFactor;
            const maxFee = stbh / factorRow.minFactor;
            const rangeEl = document.getElementById('mul-fee-range');
            if (rangeEl) rangeEl.textContent = `Phí hợp lệ từ ${formatCurrency(minFee, '')} đến ${formatCurrency(maxFee, '')}.`;

            const entered = parseFormattedNumber(feeInput?.value || '0');
            if (entered > 0 && (entered < minFee || entered > maxFee || entered < 5000000)) {
                setFieldError(feeInput, 'Phí không hợp lệ');
            } else {
                clearFieldError(feeInput);
            }
        }
    }
}

/* ---------- HTML generators ---------- */
function generateSupplementaryPersonHtml(personId, count) {
    return `
        <button class="w-full text-right text-sm text-red-600 font-semibold" onclick="this.closest('.person-container').remove(); updateSupplementaryAddButtonState(); calculateAll();">Xóa NĐBH này</button>
        <h3 class="text-lg font-bold text-gray-700 mb-2 border-t pt-4">NĐBH Bổ Sung ${count}</h3>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
                <label for="name-${personId}" class="font-medium text-gray-700 block mb-1">Họ và Tên</label>
                <input type="text" id="name-${personId}" class="form-input name-input" placeholder="Trần Thị B">
            </div>
            <div>
                <label for="dob-${personId}" class="font-medium text-gray-700 block mb-1">Ngày sinh</label>
                <input type="text" id="dob-${personId}" class="form-input dob-input" placeholder="DD/MM/YYYY">
            </div>
            <div>
                <label for="gender-${personId}" class="font-medium text-gray-700 block mb-1">Giới tính</label>
                <select id="gender-${personId}" class="form-select gender-select">
                    <option value="Nam">Nam</option>
                    <option value="Nữ">Nữ</option>
                </select>
            </div>
            <div class="flex items-end space-x-4">
                <p class="text-lg">Tuổi: <span id="age-${personId}" class="font-bold text-aia-red age-span">0</span></p>
            </div>
            <div class="relative">
                <label for="occupation-input-${personId}" class="font-medium text-gray-700 block mb-1">Nghề nghiệp</label>
                <input type="text" id="occupation-input-${personId}" class="form-input occupation-input" placeholder="Gõ để tìm nghề nghiệp...">
                <div class="occupation-autocomplete absolute z-10 w-full bg-white border border-gray-300 rounded-md mt-1 hidden max-h-60 overflow-y-auto"></div>
            </div>
            <div class="flex items-end space-x-4">
                <p class="text-lg">Nhóm nghề: <span id="risk-group-${personId}" class="font-bold text-aia-red risk-group-span">...</span></p>
            </div>
        </div>
        <div class="mt-4">
            <h4 class="text-md font-semibold text-gray-800 mb-2">Sản phẩm bổ sung cho người này</h4>
            <div class="supplementary-products-container space-y-6"></div>
        </div>
    `;
}

function generateSupplementaryProductsHtml(personId) {
    return `
        <div class="product-section health-scl-section hidden">
            <label class="flex items-center space-x-3 cursor-pointer">
                <input type="checkbox" class="form-checkbox health-scl-checkbox">
                <span class="text-lg font-medium text-gray-800">Sức khỏe Bùng Gia Lực</span>
            </label>
            <div class="product-options hidden mt-3 pl-8 space-y-4 border-l-2 border-gray-200">
                <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                        <label class="font-medium text-gray-700 block mb-1">Quyền lợi chính (Bắt buộc)</label>
                        <select class="form-select health-scl-program" disabled>
                            <option value="">-- Chọn chương trình --</option>
                            <option value="co_ban">Cơ bản</option>
                            <option value="nang_cao">Nâng cao</option>
                            <option value="toan_dien">Toàn diện</option>
                            <option value="hoan_hao">Hoàn hảo</option>
                        </select>
                        <div class="text-sm text-gray-600 mt-1 health-scl-stbh-info"></div>
                    </div>
                    <div>
                        <label class="font-medium text-gray-700 block mb-1">Phạm vi địa lý</label>
                        <select class="form-select health-scl-scope" disabled>
                            <option value="main_vn">Việt Nam</option>
                            <option value="main_global">Nước ngoài</option>
                        </select>
                    </div>
                </div>
                <div>
                    <span class="font-medium text-gray-700 block mb-2">Quyền lợi tùy chọn:</span>
                    <div class="space-y-2">
                        <label class="flex items-center space-x-3 cursor-pointer"><input type="checkbox" class="form-checkbox health-scl-outpatient" disabled> <span>Điều trị ngoại trú</span></label>
                        <label class="flex items-center space-x-3 cursor-pointer"><input type="checkbox" class="form-checkbox health-scl-dental" disabled> <span>Chăm sóc nha khoa</span></label>
                    </div>
                </div>
                <div class="text-right font-semibold text-aia-red fee-display min-h-[1.5rem]"></div>
            </div>
        </div>
        <div class="product-section bhn-section hidden">
            <label class="flex items-center space-x-3 cursor-pointer">
                <input type="checkbox" class="form-checkbox bhn-checkbox"> <span class="text-lg font-medium text-gray-800">Bảo hiểm Bệnh Hiểm Nghèo 2.0</span>
            </label>
            <div class="product-options hidden mt-3 pl-8 space-y-3 border-l-2 border-gray-200">
                <div><label class="font-medium text-gray-700 block mb-1">Số tiền bảo hiểm (STBH)</label><input type="text" class="form-input bhn-stbh" placeholder="VD: 500.000.000"></div>
                <div class="text-right font-semibold text-aia-red fee-display min-h-[1.5rem]"></div>
            </div>
        </div>
        <div class="product-section accident-section hidden">
            <label class="flex items-center space-x-3 cursor-pointer">
                <input type="checkbox" class="form-checkbox accident-checkbox"> <span class="text-lg font-medium text-gray-800">Bảo hiểm Tai nạn</span>
            </label>
            <div class="product-options hidden mt-3 pl-8 space-y-3 border-l-2 border-gray-200">
                <div><label class="font-medium text-gray-700 block mb-1">Số tiền bảo hiểm (STBH)</label><input type="text" class="form-input accident-stbh" placeholder="VD: 200.000.000"></div>
                <div class="text-right font-semibold text-aia-red fee-display min-h-[1.5rem]"></div>
            </div>
        </div>
        <div class="product-section hospital-support-section hidden">
            <label class="flex items-center space-x-3 cursor-pointer">
                <input type="checkbox" class="form-checkbox hospital-support-checkbox"> <span class="text-lg font-medium text-gray-800">Hỗ trợ chi phí nằm viện</span>
            </label>
            <div class="product-options hidden mt-3 pl-8 space-y-3 border-l-2 border-gray-200">
                <div>
                    <label class="font-medium text-gray-700 block mb-1">Số tiền hỗ trợ/ngày</label>
                    <input type="text" class="form-input hospital-support-stbh" placeholder="VD: 300.000">
                    <p class="hospital-support-validation text-sm text-gray-500 mt-1"></p>
                </div>
                <div class="text-right font-semibold text-aia-red fee-display min-h-[1.5rem]"></div>
            </div>
        </div>
    `;
}
