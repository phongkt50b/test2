import { product_data } from './data.js';

let supplementaryInsuredCount = 0;
let currentMainProductState = { product: null, age: null };

const MAX_ENTRY_AGE = {
  PUL_TRON_DOI: 70, PUL_15_NAM: 70, PUL_5_NAM: 70, KHOE_BINH_AN: 70, VUNG_TUONG_LAI: 70,
  TRON_TAM_AN: 60, AN_BINH_UU_VIET: 65,
  health_scl: 65, bhn: 70, accident: 64, hospital_support: 55
};
const MAX_RENEWAL_AGE = { health_scl: 74, bhn: 85, accident: 65, hospital_support: 59 };
const MAX_STBH = { bhn: 5_000_000_000, accident: 8_000_000_000 };
const WAIVER30 = { MIN_AGE: 18, MAX_AGE: 60, MAX_RENEWAL_AGE: 65 };
const REFERENCE_DATE = new Date(2025, 7, 9); // 09/08/2025

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
      '.health-scl-checkbox', '.bhn-checkbox', '.accident-checkbox', '.hospital-support-checkbox'
    ];
    if (checkboxSelectors.some(selector => e.target.matches(selector))) {
      const section = e.target.closest('.product-section');
      const options = section.querySelector('.product-options');
      if (e.target.checked && !e.target.disabled) options.classList.remove('hidden');
      else options.classList.add('hidden');
      calculateAll();
    } else if (e.target.matches(allInputs)) {
      calculateAll();
    }
  });
  document.body.addEventListener('input', (e) => {
    if (e.target.matches('input[type="text"]') &&
        !e.target.classList.contains('dob-input') &&
        !e.target.classList.contains('occupation-input') &&
        !e.target.classList.contains('name-input')) {
      formatNumberInput(e.target);
      calculateAll();
    } else if (e.target.matches('input[type="number"]')) {
      calculateAll();
    }
  });
}

/* Section 1 */
function initPerson(container, personId, isSupp = false) {
  if (!container) return;
  container.dataset.personId = personId;

  initDateFormatter(container.querySelector('.dob-input'));
  initOccupationAutocomplete(container.querySelector('.occupation-input'), container);

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

  const suppProductsContainer = isSupp
    ? container.querySelector('.supplementary-products-container')
    : document.querySelector('#main-supp-container .supplementary-products-container');
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
      if (rounded !== raw) hsInput.value = rounded.toLocaleString('vi-VN');
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

/* Section 4 */
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

/* Section 5: Miễn Đóng Phí 3.0 */
function initWaiver30Section() {
  const cont = document.getElementById('waiver30-container');
  if (!cont) return;
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

  // Danh sách NĐBH bổ sung 18–60 tuổi
  const eligibleSupps = [];
  document.querySelectorAll('#supplementary-insured-container .person-container').forEach(p => {
    const info = getCustomerInfo(p, false);
    if (info.age >= WAIVER30.MIN_AGE && info.age <= WAIVER30.MAX_AGE) {
      eligibleSupps.push({ id: p.id, name: info.name || 'NĐBH bổ sung', age: info.age, gender: info.gender, container: p });
    }
  });

  // Giữ selection
  const prev = cont.querySelector('input[name="waiver30-insurer"]:checked')?.value || cont.dataset.selected || '';

  let radiosHtml = '';
  if (eligibleSupps.length > 0) {
    radiosHtml += `<div class="space-y-2">`;
    eligibleSupps.forEach(p => {
      radiosHtml += `
      <label class="flex items-center space-x-3 cursor-pointer">
        <input type="radio" name="waiver30-insurer" class="form-checkbox" value="${p.id}" ${prev === p.id ? 'checked' : ''}>
        <span>${sanitizeHtml(p.name)} (Tuổi ${p.age})</span>
      </label>`;
    });
    radiosHtml += `</div>`;
  }
  // "Người khác"
  radiosHtml += `
  <label class="flex items-center space-x-3 cursor-pointer mt-2">
    <input type="radio" name="waiver30-insurer" class="form-checkbox" value="other" ${prev === 'other' ? 'checked' : ''}>
    <span>Người khác</span>
  </label>`;

  const otherFormHtml = `
  <div id="waiver30-other-form" class="${prev === 'other' ? '' : 'hidden'} mt-3">
    <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
      <div><label class="font-medium text-gray-700 block mb-1">Họ và Tên</label><input type="text" class="form-input waiver-other-name" placeholder="Nguyễn Văn C"></div>
      <div><label class="font-medium text-gray-700 block mb-1">Ngày sinh</label><input type="text" class="form-input waiver-other-dob dob-input" placeholder="DD/MM/YYYY"></div>
      <div><label class="font-medium text-gray-700 block mb-1">Giới tính</label><select class="form-select waiver-other-gender"><option value="Nam">Nam</option><option value="Nữ">Nữ</option></select></div>
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
      ${eligibleSupps.length > 0 ? `<p class="text-sm text-gray-600">Chọn người áp dụng (18–60 tuổi) hoặc chọn “Người khác”.</p>` : `<p class="text-sm text-gray-600">Không có NĐBH bổ sung 18–60 tuổi. Vui lòng chọn “Người khác”.</p>`}
      ${radiosHtml}
      ${otherFormHtml}
      <div class="text-right font-semibold text-aia-red waiver30-fee-display min-h-[1.5rem]"></div>
    </div>`;

  cont.querySelectorAll('input[name="waiver30-insurer"]').forEach(r => {
    r.addEventListener('change', () => {
      cont.dataset.selected = r.value;
      cont.querySelector('#waiver30-other-form')?.classList.toggle('hidden', r.value !== 'other');
      calculateAll();
    });
  });

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
  return { name: (nameInput?.value || '').trim() || 'Người khác', age, gender: genderSelect.value, riskGroup };
}
function validateWaiverOtherForm() {
  const cont = document.getElementById('waiver30-container');
  if (!cont) return false;
  const nameInput = cont.querySelector('.waiver-other-name');
  const dobInput = cont.querySelector('.waiver-other-dob');
  const occInput = cont.querySelector('.waiver-other-occupation');

  let ok = true;
  if (nameInput) {
    const v = (nameInput.value || '').trim();
    if (!v) { setFieldError(nameInput, 'Vui lòng nhập họ và tên'); ok = false; } else clearFieldError(nameInput);
  }
  if (dobInput) {
    const v = (dobInput.value || '').trim();
    const re = /^\d{2}\/\d{2}\/\d{4}$/;
    if (!re.test(v)) { setFieldError(dobInput, 'Ngày sinh không hợp lệ, nhập DD/MM/YYYY'); ok = false; }
    else {
      const [dd, mm, yyyy] = v.split('/').map(n => parseInt(n, 10));
      const d = new Date(yyyy, mm - 1, dd);
      const valid = d.getFullYear() === yyyy && d.getMonth() === (mm - 1) && d.getDate() === dd && d <= REFERENCE_DATE;
      if (!valid) { setFieldError(dobInput, 'Ngày sinh không hợp lệ, nhập DD/MM/YYYY'); ok = false; } else clearFieldError(dobInput);
    }
  }
  if (occInput) {
    const typed = (occInput.value || '').trim().toLowerCase();
    const match = product_data.occupations.find(o => o.group > 0 && o.name.toLowerCase() === typed);
    const group = parseInt(occInput.dataset.group, 10);
    if (!match || !(group >= 1 && group <= 4)) { setFieldError(occInput, 'Chọn nghề nghiệp từ danh sách'); ok = false; } else clearFieldError(occInput);
  }
  const other = getWaiverOtherInfo();
  if (!other || other.age < WAIVER30.MIN_AGE || other.age > WAIVER30.MAX_AGE) {
    if (dobInput) setFieldError(dobInput, 'Độ tuổi phải từ 18 đến 60');
    ok = false;
  }
  return ok;
}
function calculateWaiver30Premium(mainPersonInfo, mainPremiumDisplay) {
  const cont = document.getElementById('waiver30-container');
  if (!cont) return { fee: 0, selected: '' };
  const feeEl = cont.querySelector('.waiver30-fee-display');
  const mainProduct = document.getElementById('main-product')?.value || '';
  if (mainProduct === 'TRON_TAM_AN') { if (feeEl) feeEl.textContent = ''; return { fee: 0, selected: '' }; }

  const selected = cont.querySelector('input[name="waiver30-insurer"]:checked')?.value || '';
  if (!selected) { if (feeEl) feeEl.textContent = ''; return { fee: 0, selected: '' }; }

  // Phí bổ sung của NĐBH chính
  const mainSuppContainer = document.querySelector('#main-supp-container .supplementary-products-container');
  let suppMain = 0;
  if (mainSuppContainer) {
    suppMain += calculateHealthSclPremium(mainPersonInfo, mainSuppContainer);
    suppMain += calculateBhnPremium(mainPersonInfo, mainSuppContainer);
    suppMain += calculateAccidentPremium(mainPersonInfo, mainSuppContainer);
    const baseMainPremium = mainPremiumDisplay - getExtraPremiumValue();
    suppMain += calculateHospitalSupportPremium(mainPersonInfo, baseMainPremium, mainSuppContainer, 0);
  }

  // Tổng phí bổ sung của tất cả NĐBH bổ sung, và phí của người được chọn (nếu là 1 trong danh sách)
  let sumSuppAllSupp = 0;
  let suppOfSelected = 0;
  document.querySelectorAll('#supplementary-insured-container .person-container').forEach(p => {
    const info = getCustomerInfo(p, false);
    const container = p.querySelector('.supplementary-products-container');
    if (!container) return;
    let sub = 0;
    sub += calculateHealthSclPremium(info, container);
    sub += calculateBhnPremium(info, container);
    sub += calculateAccidentPremium(info, container);
    const baseMainPremium = mainPremiumDisplay - getExtraPremiumValue();
    sub += calculateHospitalSupportPremium(info, baseMainPremium, container, 0);
    sumSuppAllSupp += sub;
    if (selected === p.id) suppOfSelected = sub;
  });

  const stbh = Math.max(0, mainPremiumDisplay + suppMain + sumSuppAllSupp - (selected === 'other' ? 0 : suppOfSelected));

  // Lấy tuổi/giới tính người chọn
  let insuredAge = 0;
  let insuredGender = 'Nam';
  if (selected === 'other') {
    if (!validateWaiverOtherForm()) { if (feeEl) feeEl.textContent = ''; return { fee: 0, selected }; }
    const other = getWaiverOtherInfo();
    insuredAge = other.age;
    insuredGender = other.gender;
  } else {
    const pEl = document.getElementById(selected);
    const pInfo = pEl ? getCustomerInfo(pEl, false) : null;
    if (!pInfo || pInfo.age < WAIVER30.MIN_AGE || pInfo.age > WAIVER30.MAX_AGE) { if (feeEl) feeEl.textContent = ''; return { fee: 0, selected: '' }; }
    insuredAge = pInfo.age;
    insuredGender = pInfo.gender;
  }

  const rateRow = product_data.waiver30_rates?.find(r => insuredAge >= r.ageMin && insuredAge <= r.ageMax);
  if (!rateRow) { if (feeEl) feeEl.textContent = ''; return { fee: 0, selected }; }
  const rate = insuredGender === 'Nữ' ? (rateRow.nu || 0) : (rateRow.nam || 0);
  const premium = (stbh / 1000) * rate;

  if (feeEl) feeEl.textContent = premium > 0 ? `Phí: ${formatCurrency(premium)}` : '';
  return { fee: premium, selected };
}

/* Modal Section 6 */
function initSummaryModal() {
  const modal = document.getElementById('summary-modal');
  document.getElementById('view-summary-btn').addEventListener('click', generateSummaryTable);
  document.getElementById('close-summary-modal-btn').addEventListener('click', () => modal.classList.add('hidden'));
  modal.addEventListener('click', (e) => { if (e.target === modal) modal.classList.add('hidden'); });

  const targetAgeInput = document.getElementById('target-age-input');
  const mainPersonInfo = getCustomerInfo(document.getElementById('main-person-container'), true);
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
    if (document.getElementById('summary-modal').classList.contains('hidden')) calculateAll();
    else generateSummaryTable();
  });

  const mainDobInput = document.querySelector('#main-person-container .dob-input');
  if (mainDobInput) {
    mainDobInput.addEventListener('input', () => {
      updateTargetAge();
      if (document.getElementById('summary-modal').classList.contains('hidden')) calculateAll();
      else generateSummaryTable();
    });
  }

  if (abuvTermSelect) {
    abuvTermSelect.addEventListener('change', () => {
      updateTargetAge();
      if (document.getElementById('summary-modal').classList.contains('hidden')) calculateAll();
      else generateSummaryTable();
    });
  }
  document.getElementById('payment-term')?.addEventListener('change', () => {
    updateTargetAge();
    if (document.getElementById('summary-modal').classList.contains('hidden')) calculateAll();
    else generateSummaryTable();
  });
}
function updateTargetAge() {
  const mainPersonInfo = getCustomerInfo(document.getElementById('main-person-container'), true);
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

/* Autocomplete + formatters */
function initDateFormatter(input) {
  if (!input) return;
  input.addEventListener('input', (e) => {
    let value = e.target.value.replace(/\D/g, '');
    if (value.length > 2) value = value.slice(0, 2) + '/' + value.slice(2);
    if (value.length > 5) value = value.slice(0, 5) + '/' + value.slice(5, 9);
    e.target.value = value.slice(0, 10);
  });
}
function initOccupationAutocomplete(input, container) {
  if (!input) return;
  const ac = container.querySelector('.occupation-autocomplete');
  const riskGroupSpan = container.querySelector('.risk-group-span');

  const apply = (occ) => {
    input.value = occ.name;
    input.dataset.group = occ.group;
    if (riskGroupSpan) riskGroupSpan.textContent = occ.group;
    clearFieldError(input);
    ac.classList.add('hidden');
    calculateAll();
  };
  const renderList = (filtered) => {
    ac.innerHTML = '';
    if (filtered.length === 0) { ac.classList.add('hidden'); return; }
    filtered.forEach(occ => {
      const item = document.createElement('div');
      item.className = 'autocomplete-item';
      item.textContent = occ.name;
      item.addEventListener('mousedown', (ev) => { ev.preventDefault(); apply(occ); });
      ac.appendChild(item);
    });
    ac.classList.remove('hidden');
  };

  input.addEventListener('input', () => {
    const value = input.value.trim().toLowerCase();
    if (value.length < 2) { ac.classList.add('hidden'); return; }
    const filtered = product_data.occupations.filter(o => o.group > 0 && o.name.toLowerCase().includes(value));
    renderList(filtered);
  });

  input.addEventListener('blur', () => {
    setTimeout(() => {
      const typed = (input.value || '').trim().toLowerCase();
      const match = product_data.occupations.find(o => o.group > 0 && o.name.toLowerCase() === typed);
      if (typed && match) apply(match);
      else {
        input.dataset.group = '';
        if (riskGroupSpan) riskGroupSpan.textContent = '...';
        setFieldError(input, 'Chọn nghề nghiệp từ danh sách');
        ac.classList.add('hidden');
        calculateAll();
      }
    }, 0);
  });

  document.addEventListener('click', (e) => { if (!container.contains(e.target)) ac.classList.add('hidden'); });
}

/* Helpers chung */
function getCustomerInfo(container, isMain = false) {
  const dobInput = container.querySelector('.dob-input');
  const genderSelect = container.querySelector('.gender-select');
  const occupationInput = container.querySelector('.occupation-input');
  const ageSpan = container.querySelector('.age-span');
  const riskGroupSpan = container.querySelector('.risk-group-span');
  const nameInput = container.querySelector('.name-input');

  let age = 0, daysFromBirth = 0;
  const dobStr = dobInput ? dobInput.value : '';
  if (dobStr && /^\d{2}\/\d{2}\/\d{4}$/.test(dobStr)) {
    const [dd, mm, yyyy] = dobStr.split('/').map(n => parseInt(n, 10));
    const birthDate = new Date(yyyy, mm - 1, dd);
    const isValidDate = birthDate.getFullYear() === yyyy && birthDate.getMonth() === (mm - 1) && birthDate.getDate() === dd;
    if (isValidDate && birthDate <= REFERENCE_DATE) {
      const diffMs = REFERENCE_DATE - birthDate;
      daysFromBirth = Math.floor(diffMs / (1000 * 60 * 60 * 24));
      age = REFERENCE_DATE.getFullYear() - birthDate.getFullYear();
      const m = REFERENCE_DATE.getMonth() - birthDate.getMonth();
      if (m < 0 || (m === 0 && REFERENCE_DATE.getDate() < birthDate.getDate())) age--;
    }
  }
  if (ageSpan) ageSpan.textContent = age;
  const riskGroup = occupationInput ? parseInt(occupationInput.dataset.group, 10) || 0 : 0;
  if (riskGroupSpan) riskGroupSpan.textContent = riskGroup > 0 ? riskGroup : '...';

  const info = {
    age, daysFromBirth,
    gender: genderSelect ? genderSelect.value : 'Nam',
    riskGroup, container, name: nameInput ? nameInput.value : 'NĐBH Chính'
  };
  if (isMain) info.mainProduct = document.getElementById('main-product').value;
  return info;
}

function calculateAll() {
  try {
    clearError();
    validateMainPersonInputs();

    const mainPersonInfo = getCustomerInfo(document.getElementById('main-person-container'), true);
    updateMainProductVisibility(mainPersonInfo);

    validateSection2FieldsPreCalc(mainPersonInfo);

    const baseMainPremium = calculateMainPremium(mainPersonInfo);
    validateExtraPremiumLimit(baseMainPremium);
    const extraPremium = getExtraPremiumValue();
    const mainPremiumDisplay = baseMainPremium + extraPremium;

    updateMainProductFeeDisplay(baseMainPremium, extraPremium);

    updateSupplementaryProductVisibility(mainPersonInfo, baseMainPremium, document.querySelector('#main-supp-container .supplementary-products-container'));

    let totalSupplementaryPremium = 0;
    let totalHospitalSupportStbh = 0;
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

      const hs = parseFormattedNumber(suppProductsContainer.querySelector('.hospital-support-stbh')?.value || '0');
      if (suppProductsContainer.querySelector('.hospital-support-checkbox')?.checked && hs > 0) totalHospitalSupportStbh += hs;
    });

    // Render Section 5 sau khi đã có phí chính + bổ sung
    renderWaiver30Section(mainPersonInfo);

    // Tính phí miễn đóng phí 3.0 nếu có chọn
    let waiverData = { fee: 0, selected: '' };
    try { waiverData = calculateWaiver30Premium(mainPersonInfo, mainPremiumDisplay); } catch (e) { waiverData = { fee: 0, selected: '' }; }
    totalSupplementaryPremium += (waiverData.fee || 0);

    const totalPremium = mainPremiumDisplay + totalSupplementaryPremium;
    // Hiển thị riêng nếu "Người khác"
    updateSummaryUI({
      mainPremium: mainPremiumDisplay,
      totalSupplementaryPremium,
      totalPremium,
      waiverOtherFee: (waiverData.selected === 'other') ? waiverData.fee : 0
    });

  } catch (e) {
    showError(e.message);
    updateSummaryUI({ mainPremium: 0, totalSupplementaryPremium: 0, totalPremium: 0, waiverOtherFee: 0 });
  }
}

function updateMainProductVisibility(customer) {
  const { age, daysFromBirth, gender, riskGroup } = customer;
  const mainProductSelect = document.getElementById('main-product');

  document.querySelectorAll('#main-product option').forEach(option => {
    const productKey = option.value; if (!productKey) return;
    let ok = true;
    const PUL_MUL = ['PUL_TRON_DOI','PUL_15_NAM','PUL_5_NAM','KHOE_BINH_AN','VUNG_TUONG_LAI'];
    if (PUL_MUL.includes(productKey)) ok = (daysFromBirth >= 30) && (age <= 70);
    if (productKey === 'TRON_TAM_AN') {
      const byGender = (gender === 'Nam') ? (age >= 12 && age <= 60) : (age >= 28 && age <= 60);
      ok = byGender && (riskGroup !== 4);
    }
    if (productKey === 'AN_BINH_UU_VIET') {
      const minOk = (gender === 'Nam') ? age >= 12 : age >= 28;
      ok = minOk && (age <= 65);
    }
    option.disabled = !ok;
    option.classList.toggle('hidden', !ok);
  });

  if (mainProductSelect.options[mainProductSelect.selectedIndex]?.disabled) mainProductSelect.value = '';

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

  // Ẩn Section 5 nếu TTA
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
    if (!section) return;
    const checkbox = section.querySelector('input[type="checkbox"]');
    const options = section.querySelector('.product-options');
    const ok = condition && daysFromBirth >= 30 && age >= 0 && age <= MAX_ENTRY_AGE[productKey] && (sectionId !== 'health-scl' || riskGroup !== 4);

    if (ok) {
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
            if (!programSelect.querySelector('option[value="nang_cao"]').disabled) programSelect.value = 'nang_cao';
          }
          if (!scopeSelect.value) scopeSelect.value = 'main_vn';
          outpatient.disabled = false;
          dental.disabled = false;
          updateHealthSclStbhInfo(section);
        } else {
          programSelect.disabled = false;
          scopeSelect.disabled = false;
          programSelect.querySelectorAll('option').forEach(opt => {
            if (opt.value === '') return;
            if (mainPremium >= 15000000) opt.disabled = false;
            else if (mainPremium >= 10000000) opt.disabled = !['co_ban','nang_cao','toan_dien'].includes(opt.value);
            else if (mainPremium >= 5000000) opt.disabled = !['co_ban','nang_cao'].includes(opt.value);
            else opt.disabled = true;
          });
          if (!programSelect.value || programSelect.options[programSelect.selectedIndex]?.disabled) {
            const nc = programSelect.querySelector('option[value="nang_cao"]');
            if (nc && !nc.disabled) programSelect.value = 'nang_cao';
            else {
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

  const baseCondition = ['PUL_TRON_DOI','PUL_15_NAM','PUL_5_NAM','KHOE_BINH_AN','VUNG_TUONG_LAI','AN_BINH_UU_VIET','TRON_TAM_AN'].includes(mainProduct);
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

/* Section 2: SP chính */
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
      <div><label class="font-medium text-gray-700 block mb-1">Số tiền bảo hiểm (STBH)</label>
      <input type="text" id="main-stbh" class="form-input bg-gray-100" value="100.000.000" disabled></div>
      <div><p class="text-sm text-gray-600 mt-1">Thời hạn đóng phí: 10 năm (bằng thời hạn hợp đồng). Thời gian bảo vệ: 10 năm.</p></div>`;
  } else if (mainProduct === 'AN_BINH_UU_VIET') {
    optionsHtml = `
      <div><label class="font-medium text-gray-700 block mb-1">Số tiền bảo hiểm (STBH)</label>
      <input type="text" id="main-stbh" class="form-input" value="${currentStbh}" placeholder="VD: 1.000.000.000"></div>`;
    let termOptions = '';
    if (age <= 55) termOptions += '<option value="15">15 năm</option>';
    if (age <= 60) termOptions += '<option value="10">10 năm</option>';
    if (age <= 65) termOptions += '<option value="5">5 năm</option>';
    if (!termOptions) termOptions = '<option value="" disabled>Không có kỳ hạn phù hợp (tuổi vượt quá 65)</option>';
    optionsHtml += `
      <div><label class="font-medium text-gray-700 block mb-1">Thời hạn đóng phí</label>
      <select id="abuv-term" class="form-select">${termOptions}</select>
      <p class="text-sm text-gray-500 mt-1">Thời hạn đóng phí bằng thời hạn hợp đồng.</p></div>`;
  } else if (['KHOE_BINH_AN','VUNG_TUONG_LAI','PUL_TRON_DOI','PUL_15_NAM','PUL_5_NAM'].includes(mainProduct)) {
    optionsHtml = `
      <div><label class="font-medium text-gray-700 block mb-1">Số tiền bảo hiểm (STBH)</label>
      <input type="text" id="main-stbh" class="form-input" value="${currentStbh}" placeholder="VD: 1.000.000.000"></div>`;
    if (['KHOE_BINH_AN','VUNG_TUONG_LAI'].includes(mainProduct)) {
      optionsHtml += `
        <div><label class="font-medium text-gray-700 block mb-1">Phí sản phẩm chính</label>
        <input type="text" id="main-premium-input" class="form-input" value="${currentPremium}" placeholder="Nhập phí">
        <div id="mul-fee-range" class="text-sm text-gray-500 mt-1"></div></div>`;
    }
    const { min, max } = getPaymentTermBounds(age);
    const effMin = mainProduct === 'PUL_5_NAM' ? 5 : (mainProduct === 'PUL_15_NAM' ? 15 : 4);
    optionsHtml += `
      <div><label class="font-medium text-gray-700 block mb-1">Thời gian đóng phí (năm)</label>
      <input type="number" id="payment-term" class="form-input" value="${currentPaymentTerm}" placeholder="VD: 20" min="${effMin}" max="${max}">
      <div id="payment-term-hint" class="text-sm text-gray-500 mt-1"></div></div>`;
    optionsHtml += `
      <div><label class="font-medium text-gray-700 block mb-1">Phí đóng thêm</label>
      <input type="text" id="extra-premium-input" class="form-input" value="${currentExtra || ''}" placeholder="VD: 10.000.000">
      <div class="text-sm text-gray-500 mt-1">Tối đa 5 lần phí chính.</div></div>`;
  }
  container.innerHTML = optionsHtml;
  if (['KHOE_BINH_AN','VUNG_TUONG_LAI','PUL_TRON_DOI','PUL_15_NAM','PUL_5_NAM'].includes(mainProduct)) {
    setPaymentTermHint(mainProduct, age);
  }
}

function calculateMainPremium(customer, ageOverride = null) {
  const ageToUse = ageOverride ?? customer.age;
  const { gender, mainProduct } = customer;
  let premium = 0;

  if (mainProduct.startsWith('PUL') || mainProduct === 'AN_BINH_UU_VIET' || mainProduct === 'TRON_TAM_AN') {
    let stbh = 0; let rate = 0;
    const stbhEl = document.getElementById('main-stbh');
    if (stbhEl) stbh = parseFormattedNumber(stbhEl.value);
    if (mainProduct !== 'TRON_TAM_AN' && stbh === 0) return 0;

    const genderKey = (gender === 'Nữ') ? 'nu' : 'nam';

    if (mainProduct.startsWith('PUL')) {
      const row = product_data.pul_rates[mainProduct]?.find(r => r.age === customer.age);
      const pulRate = row ? (row[genderKey] || 0) : 0;
      if (pulRate === 0 && !ageOverride) return 0;
      rate = pulRate;
      premium = (stbh / 1000) * rate;

      if (!ageOverride) {
        if (stbh > 0 && stbh < 100000000) setFieldError(stbhEl, 'STBH nhỏ hơn 100 triệu'); else clearFieldError(stbhEl);
        if (premium > 0 && premium < 5000000) setFieldError(stbhEl, 'Phí chính nhỏ hơn 5 triệu');
      }
    } else if (mainProduct === 'AN_BINH_UU_VIET') {
      const term = document.getElementById('abuv-term')?.value;
      if (!term) return 0;
      const row = product_data.an_binh_uu_viet_rates[term]?.find(r => r.age === customer.age);
      const abuvRate = row ? (row[genderKey] || 0) : 0;
      if (abuvRate === 0 && !ageOverride) return 0;
      rate = abuvRate;
      premium = (stbh / 1000) * rate;

      if (!ageOverride) {
        if (stbh > 0 && stbh < 100000000) setFieldError(stbhEl, 'STBH nhỏ hơn 100 triệu'); else clearFieldError(stbhEl);
        if (premium > 0 && premium < 5000000) setFieldError(stbhEl, 'Phí chính nhỏ hơn 5 triệu');
      }
    } else if (mainProduct === 'TRON_TAM_AN') {
      stbh = 100000000;
      const row = product_data.an_binh_uu_viet_rates['10']?.find(r => r.age === customer.age);
      const ttaRate = row ? (row[genderKey] || 0) : 0;
      if (ttaRate === 0 && !ageOverride) return 0;
      premium = (stbh / 1000) * ttaRate;
    }
  } else if (['KHOE_BINH_AN', 'VUNG_TUONG_LAI'].includes(mainProduct)) {
    const stbh = parseFormattedNumber(document.getElementById('main-stbh')?.value || '0');
    const factorRow = product_data.mul_factors.find(f => ageToUse >= f.ageMin && ageToUse <= f.ageMax);
    if (!factorRow) return 0;

    const minFee = stbh / factorRow.maxFactor;
    const maxFee = stbh / factorRow.minFactor;
    const rangeEl = document.getElementById('mul-fee-range');
    if (!ageOverride && rangeEl) {
      rangeEl.textContent = `Phí hợp lệ từ ${formatCurrency(minFee, '')} đến ${formatCurrency(maxFee, '')}.`;
    }

    const enteredPremium = parseFormattedNumber(document.getElementById('main-premium-input')?.value || '0');
    if (!ageOverride) {
      const feeInput = document.getElementById('main-premium-input');
      if (stbh > 0 && enteredPremium > 0) {
        const invalid = (enteredPremium < minFee || enteredPremium > maxFee || enteredPremium < 5000000);
        if (invalid) setFieldError(feeInput, 'Phí không hợp lệ'); else clearFieldError(feeInput);
      } else clearFieldError(feeInput);
    }
    premium = enteredPremium;
  }

  return premium;
}

/* Section 3: tính phí */
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

  const idx = product_data.health_scl_rates.age_bands.findIndex(b => ageToUse >= b.min && ageToUse <= b.max);
  if (idx === -1) return 0;

  let premium = 0;
  premium += product_data.health_scl_rates[scope]?.[idx]?.[program] || 0;
  if (hasOutpatient) premium += product_data.health_scl_rates.outpatient?.[idx]?.[program] || 0;
  if (hasDental) premium += product_data.health_scl_rates.dental?.[idx]?.[program] || 0;

  if (!ageOverride) section.querySelector('.fee-display').textContent = premium > 0 ? `Phí: ${formatCurrency(premium)}` : '';
  return premium;
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
  if (stbh === 0) { if (!ageOverride) section.querySelector('.fee-display').textContent = ''; return 0; }
  if (stbh < 100000000 || stbh > MAX_STBH.bhn) { setFieldError(stbhInput, 'STBH không hợp lệ, từ 100 triệu đến 5 tỷ'); throw new Error('STBH không hợp lệ, từ 100 triệu đến 5 tỷ'); } else clearFieldError(stbhInput);

  const rate = product_data.bhn_rates.find(r => ageToUse >= r.ageMin && ageToUse <= r.ageMax)?.[(gender === 'Nữ') ? 'nu' : 'nam'] || 0;
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
  if (stbh === 0) { if (!ageOverride) section.querySelector('.fee-display').textContent = ''; return 0; }
  if (stbh < 100000000 || stbh > MAX_STBH.accident) { setFieldError(stbhInput, 'STBH không hợp lệ, từ 100 triệu đến 8 tỷ'); throw new Error('STBH không hợp lệ, từ 100 triệu đến 8 tỷ'); } else clearFieldError(stbhInput);

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
  if (stbh === 0) { if (!ageOverride) section.querySelector('.fee-display').textContent = ''; clearFieldError(hsInput); return 0; }

  if (stbh % 100000 !== 0) { setFieldError(hsInput, 'STBH không hợp lệ, phải là bội số 100.000'); throw new Error('STBH không hợp lệ, phải là bội số 100.000'); }
  if (stbh > maxSupportByAge || stbh > remainingSupport) { setFieldError(hsInput, 'Vượt quá giới hạn cho phép'); throw new Error('Vượt quá giới hạn cho phép'); }
  clearFieldError(hsInput);

  const rate = product_data.hospital_fee_support_rates.find(r => ageToUse >= r.ageMin && ageToUse <= r.ageMax)?.rate || 0;
  const premium = (stbh / 100) * rate;
  if (!ageOverride) section.querySelector('.fee-display').textContent = `Phí: ${formatCurrency(premium)}`;
  return premium;
}

/* Tổng kết & UI tổng */
function updateSummaryUI({ mainPremium, totalSupplementaryPremium, totalPremium, waiverOtherFee = 0 }) {
  document.getElementById('main-premium-result').textContent = formatCurrency(mainPremium);

  const suppContainer = document.getElementById('supplementary-premiums-results');
  let html = '';
  if (waiverOtherFee > 0) {
    html += `<div class="flex justify-between items-center py-2 border-b">
      <span class="text-gray-600">Miễn đóng phí 3.0 (Người khác):</span>
      <span class="font-bold text-gray-900">${formatCurrency(waiverOtherFee)}</span>
    </div>`;
  }
  if (totalSupplementaryPremium > 0) {
    html += `<div class="flex justify-between items-center py-2 border-b">
      <span class="text-gray-600">Tổng phí SP bổ sung:</span>
      <span class="font-bold text-gray-900">${formatCurrency(totalSupplementaryPremium)}</span>
    </div>`;
  }
  suppContainer.innerHTML = html;

  document.getElementById('total-premium-result').textContent = formatCurrency(totalPremium);
}

/* Bảng Minh Họa */
function generateSummaryTable() {
  const modal = document.getElementById('summary-modal');
  const container = document.getElementById('summary-content-container');
  container.innerHTML = '';

  try {
    const targetAge = parseInt(document.getElementById('target-age-input').value, 10);
    const mainPersonInfo = getCustomerInfo(document.getElementById('main-person-container'), true);
    const mainProduct = mainPersonInfo.mainProduct;
    if (isNaN(targetAge) || targetAge <= mainPersonInfo.age) throw new Error('Vui lòng nhập một độ tuổi mục tiêu hợp lệ, lớn hơn tuổi hiện tại của NĐBH chính.');

    if (mainProduct === 'TRON_TAM_AN') {
      const mainSuppContainer = document.querySelector('#main-supp-container .supplementary-products-container');
      const healthSclPremium = calculateHealthSclPremium(mainPersonInfo, mainSuppContainer);
      const chk = mainSuppContainer?.querySelector('.health-scl-checkbox');
      if (!chk?.checked || healthSclPremium === 0) throw new Error('Sản phẩm Trọn Tâm An bắt buộc phải tham gia kèm Sức Khỏe Bùng Gia Lực với phí hợp lệ.');
    }

    let paymentTerm = 999;
    const pt = document.getElementById('payment-term');
    if (pt) paymentTerm = parseInt(pt.value, 10) || 999;
    else if (mainProduct === 'AN_BINH_UU_VIET') paymentTerm = parseInt(document.getElementById('abuv-term')?.value, 10);
    else if (mainProduct === 'TRON_TAM_AN') paymentTerm = 10;

    if (['PUL_TRON_DOI','PUL_5_NAM','PUL_15_NAM','KHOE_BINH_AN','VUNG_TUONG_LAI'].includes(mainProduct)
      && targetAge < mainPersonInfo.age + paymentTerm - 1) {
      throw new Error(`Độ tuổi mục tiêu phải lớn hơn hoặc bằng ${mainPersonInfo.age + paymentTerm - 1} đối với ${mainProduct}.`);
    }

    const suppPersons = [];
    document.querySelectorAll('.person-container').forEach(pContainer => {
      if (pContainer.id !== 'main-person-container') suppPersons.push(getCustomerInfo(pContainer, false));
    });

    // Waiver status
    const waiverCont = document.getElementById('waiver30-container');
    const waiverSelected = waiverCont ? waiverCont.querySelector('input[name="waiver30-insurer"]:checked')?.value || '' : '';
    const otherInfo = (waiverSelected === 'other') ? getWaiverOtherInfo() : null;

    let html = `<table class="w-full text-left border-collapse"><thead class="bg-gray-100"><tr>`;
    html += `<th class="p-2 border">Năm HĐ</th>`;
    html += `<th class="p-2 border">Tuổi NĐBH Chính<br>(${sanitizeHtml(mainPersonInfo.name)})</th>`;
    html += `<th class="p-2 border">Phí SP Chính<br>(${sanitizeHtml(mainPersonInfo.name)})</th>`;
    html += `<th class="p-2 border">Phí SP Bổ Sung<br>(${sanitizeHtml(mainPersonInfo.name)})</th>`;
    suppPersons.forEach(person => { html += `<th class="p-2 border">Phí SP Bổ Sung<br>(${sanitizeHtml(person.name)})</th>`; });
    if (waiverSelected === 'other') html += `<th class="p-2 border">Miễn đóng phí 3.0<br>(Người khác)</th>`;
    html += `<th class="p-2 border">Tổng Phí Năm</th></tr></thead><tbody>`;

    let totalMainAcc = 0, totalSuppAccMain = 0, totalSuppAccAll = 0, totalWaiverAcc = 0;

    const initialBaseMainPremium = calculateMainPremium(mainPersonInfo);
    const extraPremium = getExtraPremiumValue();
    const initialMainPremiumWithExtra = initialBaseMainPremium + extraPremium;
    const totalMaxSupport = Math.floor(initialBaseMainPremium / 4000000) * 100000;

    // Phí Miễn đóng phí 3.0 cố định theo năm
    const waiverData = calculateWaiver30Premium(mainPersonInfo, initialMainPremiumWithExtra);
    const waiverPerYear = waiverData.fee || 0;
    let waiverStartAge = 0;
    if (waiverPerYear > 0) {
      if (waiverData.selected === 'other') waiverStartAge = otherInfo?.age || 0;
      else {
        const pEl = waiverData.selected ? document.getElementById(waiverData.selected) : null;
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
        const hs = parseFormattedNumber(mainSuppContainer.querySelector('.hospital-support-stbh')?.value || '0');
        if (mainSuppContainer.querySelector('.hospital-support-checkbox')?.checked && hs > 0) totalHospitalSupportStbh += hs;
      }
      totalSuppAccMain += suppPremiumMain;

      // Waiver năm này (áp dụng tới 65 tuổi của người chọn)
      let waiverThisYear = 0;
      if (waiverPerYear > 0 && waiverStartAge > 0 && (waiverStartAge + i) <= WAIVER30.MAX_RENEWAL_AGE) {
        waiverThisYear = waiverPerYear;
      }

      const suppPremiums = suppPersons.map(person => {
        const currentPersonAge = person.age + i;
        const suppProductsContainer = person.container.querySelector('.supplementary-products-container');
        let sp = 0;
        if (suppProductsContainer) {
          sp += calculateHealthSclPremium({ ...person, age: currentPersonAge }, suppProductsContainer, currentPersonAge);
          sp += calculateBhnPremium({ ...person, age: currentPersonAge }, suppProductsContainer, currentPersonAge);
          sp += calculateAccidentPremium({ ...person, age: currentPersonAge }, suppProductsContainer, currentPersonAge);
          sp += calculateHospitalSupportPremium({ ...person, age: currentPersonAge }, initialBaseMainPremium, suppProductsContainer, totalHospitalSupportStbh, currentPersonAge);
          const hs = parseFormattedNumber(suppProductsContainer.querySelector('.hospital-support-stbh')?.value || '0');
          if (suppProductsContainer.querySelector('.hospital-support-checkbox')?.checked && hs > 0) totalHospitalSupportStbh += hs;
        }
        // Nếu chọn người này cho waiver, cộng vào phí của họ
        if (waiverData.selected === person.container.id && waiverThisYear > 0) sp += waiverThisYear;
        totalSuppAccAll += sp;
        return sp;
      });

      if (totalHospitalSupportStbh > totalMaxSupport) throw new Error(`Tổng STBH hỗ trợ viện phí vượt hạn mức chung: ${formatCurrency(totalMaxSupport, 'đ/ngày')}.`);

      html += `<tr>
        <td class="p-2 border text-center">${contractYear}</td>
        <td class="p-2 border text-center">${currentAgeMain}</td>
        <td class="p-2 border text-right">${formatCurrency(mainPremiumForYear)}</td>
        <td class="p-2 border text-right">${formatCurrency(suppPremiumMain)}</td>`;
      suppPremiums.forEach(sp => { html += `<td class="p-2 border text-right">${formatCurrency(sp)}</td>`; });
      if (waiverData.selected === 'other') { html += `<td class="p-2 border text-right">${formatCurrency(waiverThisYear)}</td>`; totalWaiverAcc += waiverThisYear; }
      const rowTotal = mainPremiumForYear + suppPremiumMain + suppPremiums.reduce((s, v) => s + v, 0) + (waiverData.selected === 'other' ? waiverThisYear : 0);
      html += `<td class="p-2 border text-right font-semibold">${formatCurrency(rowTotal)}</td></tr>`;
    }

    html += `<tr class="bg-gray-200 font-bold"><td class="p-2 border" colspan="2">Tổng cộng</td>`;
    html += `<td class="p-2 border text-right">${formatCurrency(totalMainAcc)}</td>`;
    html += `<td class="p-2 border text-right">${formatCurrency(totalSuppAccMain)}</td>`;
    suppPersons.forEach(() => { html += `<td class="p-2 border text-right">—</td>`; });
    if (waiverData.selected === 'other') html += `<td class="p-2 border text-right">${formatCurrency(totalWaiverAcc)}</td>`;
    const grandTotal = totalMainAcc + totalSuppAccMain + totalSuppAccAll + (waiverData.selected === 'other' ? totalWaiverAcc : 0);
    html += `<td class="p-2 border text-right">${formatCurrency(grandTotal)}</td></tr></tbody></table>`;
    html += `<div class="mt-4 text-center"><button id="export-html-btn" class="bg-blue-500 text-white px-4 py-2 rounded hover:bg-blue-600">Xuất HTML</button></div>`;
    container.innerHTML = html;

    document.getElementById('export-html-btn').addEventListener('click', () => exportToHTML(mainPersonInfo, suppPersons, targetAge, initialBaseMainPremium + extraPremium, paymentTerm));
  } catch (e) {
    container.innerHTML = `<p class="text-red-600 font-semibold text-center">${e.message}</p>`;
  } finally {
    document.getElementById('summary-modal').classList.remove('hidden');
  }
}

/* Export (gọn) */
function sanitizeHtml(str) {
  return str.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}
function exportToHTML() { window.print(); }

/* Utils */
function formatCurrency(value, suffix = ' VNĐ') {
  if (isNaN(value)) return '0' + suffix;
  return Math.round(value).toLocaleString('vi-VN') + suffix;
}
function formatNumberInput(input) {
  if (!input || !input.value) return;
  const v = input.value.replace(/[.,]/g, '');
  if (!isNaN(v) && v.length > 0) input.value = parseInt(v, 10).toLocaleString('vi-VN');
  else if (input.value !== '') input.value = '';
}
function parseFormattedNumber(s) {
  return parseInt(String(s).replace(/[.,]/g, ''), 10) || 0;
}
function showError(message) { document.getElementById('error-message').textContent = message; }
function clearError() { document.getElementById('error-message').textContent = ''; }

/* Field errors (Section 1) */
function setFieldError(input, message) {
  if (!input) return;
  let err = input.parentElement.querySelector('.field-error');
  if (!err) {
    err = document.createElement('p');
    err.className = 'field-error text-sm text-red-600 mt-1';
    input.parentElement.appendChild(err);
  }
  err.textContent = message || '';
  if (message) input.classList.add('border-red-500');
  else input.classList.remove('border-red-500');
}
function clearFieldError(input) { setFieldError(input, ''); }

/* Section 2 helpers */
function getPaymentTermBounds(age) {
  const min = 4;
  const max = Math.max(0, 100 - age - 1);
  return { min, max };
}
function setPaymentTermHint(mainProduct, age) {
  const hintEl = document.getElementById('payment-term-hint');
  if (!hintEl) return;
  const { min, max } = getPaymentTermBounds(age);
  let hint = `Nhập từ ${min} đến ${max} năm`;
  if (mainProduct === 'PUL_5_NAM') hint += ' (tối thiểu 5 năm)';
  if (mainProduct === 'PUL_15_NAM') hint += ' (tối thiểu 15 năm)';
  hintEl.textContent = hint;
}
function validateSection2FieldsPreCalc(customer) {
  const mainProduct = customer.mainProduct;

  if (mainProduct && mainProduct !== 'TRON_TAM_AN') {
    const stbhEl = document.getElementById('main-stbh');
    if (stbhEl) {
      const stbh = parseFormattedNumber(stbhEl.value || '0');
      if (stbh > 0 && stbh < 100000000) setFieldError(stbhEl, 'STBH nhỏ hơn 100 triệu'); else clearFieldError(stbhEl);
    }
  }
  if (['KHOE_BINH_AN','VUNG_TUONG_LAI','PUL_TRON_DOI','PUL_15_NAM','PUL_5_NAM'].includes(mainProduct)) {
    const el = document.getElementById('payment-term');
    if (el) {
      const { min, max } = getPaymentTermBounds(customer.age);
      const effMin = mainProduct === 'PUL_5_NAM' ? 5 : (mainProduct === 'PUL_15_NAM' ? 15 : 4);
      const val = parseInt(el.value, 10);
      if (el.value && (isNaN(val) || val < effMin || val > max)) setFieldError(el, `Thời hạn không hợp lệ, từ ${effMin} đến ${max}`);
      else clearFieldError(el);
    }
  }
  if (['KHOE_BINH_AN','VUNG_TUONG_LAI'].includes(mainProduct)) {
    const stbh = parseFormattedNumber(document.getElementById('main-stbh')?.value || '0');
    const feeInput = document.getElementById('main-premium-input');
    const factorRow = product_data.mul_factors.find(f => customer.age >= f.ageMin && customer.age <= f.ageMax);
    if (factorRow && stbh > 0) {
      const minFee = stbh / factorRow.maxFactor;
      const maxFee = stbh / factorRow.minFactor;
      const rangeEl = document.getElementById('mul-fee-range');
      if (rangeEl) rangeEl.textContent = `Phí hợp lệ từ ${formatCurrency(minFee, '')} đến ${formatCurrency(maxFee, '')}.`;
      const entered = parseFormattedNumber(feeInput?.value || '0');
      if (entered > 0 && (entered < minFee || entered > maxFee || entered < 5000000)) setFieldError(feeInput, 'Phí không hợp lệ');
      else clearFieldError(feeInput);
    }
  }
}

/* Generators */
function getHealthSclStbhByProgram(program) {
  switch (program) {
    case 'co_ban': return 100_000_000;
    case 'nang_cao': return 250_000_000;
    case 'toan_dien': return 500_000_000;
    case 'hoan_hao': return 1_000_000_000;
    default: return 0;
  }
}
function updateHealthSclStbhInfo(section) {
  const info = section.querySelector('.health-scl-stbh-info'); if (!info) return;
  const program = section.querySelector('.health-scl-program')?.value || '';
  const stbh = getHealthSclStbhByProgram(program);
  info.textContent = program ? `STBH: ${formatCurrency(stbh, '')}` : '';
}
function generateSupplementaryPersonHtml(personId, count) {
  return `
    <button class="w-full text-right text-sm text-red-600 font-semibold" onclick="this.closest('.person-container').remove(); updateSupplementaryAddButtonState(); calculateAll();">Xóa NĐBH này</button>
    <h3 class="text-lg font-bold text-gray-700 mb-2 border-t pt-4">NĐBH Bổ Sung ${count}</h3>
    <div class="grid grid-cols-1 md:grid-cols-2 gap-4">
      <div><label class="font-medium text-gray-700 block mb-1">Họ và Tên</label><input type="text" class="form-input name-input" placeholder="Trần Thị B"></div>
      <div><label class="font-medium text-gray-700 block mb-1">Ngày sinh</label><input type="text" class="form-input dob-input" placeholder="DD/MM/YYYY"></div>
      <div><label class="font-medium text-gray-700 block mb-1">Giới tính</label><select class="form-select gender-select"><option value="Nam">Nam</option><option value="Nữ">Nữ</option></select></div>
      <div class="flex items-end space-x-4"><p class="text-lg">Tuổi: <span class="font-bold text-aia-red age-span">0</span></p></div>
      <div class="relative"><label class="font-medium text-gray-700 block mb-1">Nghề nghiệp</label><input type="text" class="form-input occupation-input" placeholder="Gõ để tìm nghề nghiệp..."><div class="occupation-autocomplete absolute z-10 w-full bg-white border border-gray-300 rounded-md mt-1 hidden max-h-60 overflow-y-auto"></div></div>
      <div class="flex items-end space-x-4"><p class="text-lg">Nhóm nghề: <span class="font-bold text-aia-red risk-group-span">...</span></p></div>
    </div>
    <div class="mt-4"><h4 class="text-md font-semibold text-gray-800 mb-2">Sản phẩm bổ sung cho người này</h4><div class="supplementary-products-container space-y-6"></div></div>
  `;
}
function generateSupplementaryProductsHtml() {
  return `
    <div class="product-section health-scl-section hidden">
      <label class="flex items-center space-x-3 cursor-pointer"><input type="checkbox" class="form-checkbox health-scl-checkbox"><span class="text-lg font-medium text-gray-800">Sức khỏe Bùng Gia Lực</span></label>
      <div class="product-options hidden mt-3 pl-8 space-y-4 border-l-2 border-gray-200">
        <div class="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <div><label class="font-medium text-gray-700 block mb-1">Quyền lợi chính (Bắt buộc)</label>
            <select class="form-select health-scl-program" disabled>
              <option value="">-- Chọn chương trình --</option>
              <option value="co_ban">Cơ bản</option>
              <option value="nang_cao">Nâng cao</option>
              <option value="toan_dien">Toàn diện</option>
              <option value="hoan_hao">Hoàn hảo</option>
            </select>
            <div class="text-sm text-gray-600 mt-1 health-scl-stbh-info"></div>
          </div>
          <div><label class="font-medium text-gray-700 block mb-1">Phạm vi địa lý</label>
            <select class="form-select health-scl-scope" disabled><option value="main_vn">Việt Nam</option><option value="main_global">Toàn cầu (trừ Hoa Kỳ)</option></select>
          </div>
        </div>
        <div><span class="font-medium text-gray-700 block mb-2">Quyền lợi tùy chọn:</span>
          <div class="space-y-2">
            <label class="flex items-center space-x-3 cursor-pointer"><input type="checkbox" class="form-checkbox health-scl-outpatient" disabled> <span>Điều trị ngoại trú</span></label>
            <label class="flex items-center space-x-3 cursor-pointer"><input type="checkbox" class="form-checkbox health-scl-dental" disabled> <span>Chăm sóc nha khoa</span></label>
          </div>
        </div>
        <div class="text-right font-semibold text-aia-red fee-display min-h-[1.5rem]"></div>
      </div>
    </div>
    <div class="product-section bhn-section hidden">
      <label class="flex items-center space-x-3 cursor-pointer"><input type="checkbox" class="form-checkbox bhn-checkbox"> <span class="text-lg font-medium text-gray-800">Bệnh Hiểm Nghèo 2.0</span></label>
      <div class="product-options hidden mt-3 pl-8 space-y-3 border-l-2 border-gray-200">
        <div><label class="font-medium text-gray-700 block mb-1">Số tiền bảo hiểm (STBH)</label><input type="text" class="form-input bhn-stbh" placeholder="VD: 500.000.000"></div>
        <div class="text-right font-semibold text-aia-red fee-display min-h-[1.5rem]"></div>
      </div>
    </div>
    <div class="product-section accident-section hidden">
      <label class="flex items-center space-x-3 cursor-pointer"><input type="checkbox" class="form-checkbox accident-checkbox"> <span class="text-lg font-medium text-gray-800">Bảo hiểm Tai nạn</span></label>
      <div class="product-options hidden mt-3 pl-8 space-y-3 border-l-2 border-gray-200">
        <div><label class="font-medium text-gray-700 block mb-1">Số tiền bảo hiểm (STBH)</label><input type="text" class="form-input accident-stbh" placeholder="VD: 200.000.000"></div>
        <div class="text-right font-semibold text-aia-red fee-display min-h-[1.5rem]"></div>
      </div>
    </div>
    <div class="product-section hospital-support-section hidden">
      <label class="flex items-center space-x-3 cursor-pointer"><input type="checkbox" class="form-checkbox hospital-support-checkbox"> <span class="text-lg font-medium text-gray-800">Hỗ trợ chi phí nằm viện</span></label>
      <div class="product-options hidden mt-3 pl-8 space-y-3 border-l-2 border-gray-200">
        <div><label class="font-medium text-gray-700 block mb-1">Số tiền hỗ trợ/ngày</label><input type="text" class="form-input hospital-support-stbh" placeholder="VD: 300.000"><p class="hospital-support-validation text-sm text-gray-500 mt-1"></p></div>
        <div class="text-right font-semibold text-aia-red fee-display min-h-[1.5rem]"></div>
      </div>
    </div>
  `;
}
