
/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

// --- Data Structures ---
interface MaterialItem {
    id: string; // Unique ID for UI key purposes during form editing
    material: string;
    quantity: number;
    rate: number;
    gstPercentage: number; // GST percentage for this item
}

interface PurchaseOrder {
    id: string; // Now a simple sequential number string, e.g., "1", "2"
    externalPoNumber?: string; // Optional user-defined PO number
    partyName: string;
    gstin: string;
    salesmanName: string;
    siteAddress: string;
    destination: string;
    items: MaterialItem[]; // Note: Saved items won't need the form 'id', but will have gstPercentage
    createdAt: string; // Full ISO Date string (e.g., "2023-10-26T07:30:00.000Z")
    status: 'Pending' | 'Partially Dispatched' | 'Completed' | 'Cancelled'; // Added 'Cancelled'
    totalAmount: number; // Now includes GST
    // Tracks total dispatched quantity for each material in this PO
    dispatchedQuantityByMaterial: { [materialName: string]: number };
}

interface Dispatch {
    id: string; // Format: D-YYYYMMDD-XXXX
    poId: string;
    vehicleNumber: string;
    driverContact: string;
    invoiceNumber?: string; // Optional: Invoice number for the dispatch
    transporterName?: string; // Optional: Name of the transporter
    dispatchedItems: { material: string; quantity: number }[]; // Material name and quantity dispatched in this event
    dispatchedAt: string; // Full ISO Date string (e.g., "2023-10-26T07:30:00.000Z") - User can set this date
}

// --- Application State ---
type View = 'create-po' | 'view-po' | 'view-dispatches' | 'pending-orders';
let currentView: View = 'create-po';
let purchaseOrders: PurchaseOrder[] = [];
let dispatches: Dispatch[] = [];

// State for PO form's dynamic material items
let poFormMaterialItems: MaterialItem[] = []; // Used for create and edit PO forms
const DEFAULT_GST_RATE = 18; // Default GST rate, e.g., 18%

// State for Dispatch Log filters
let dispatchFilterStartDate: string = '';
let dispatchFilterEndDate: string = '';

const DEFAULT_PREDEFINED_MATERIALS: string[] = [
    "HITECH READYMIX PLASTER",
    "HITECH READYMIX PLASTER 1 :3",
    "FLOOR SCREED",
    "BLOCK JOINING MORATAR",
    "GYPSUM",
    "TILES ADHESIVE GOLD",
    "TILES ADHESIVE SILVER",
    "TILES ADHESIVE PLATINUM",
    "ROOFIT READYMIX PLASTER"
];
let PREDEFINED_MATERIALS: string[] = [...DEFAULT_PREDEFINED_MATERIALS];

// Current items being edited (if any)
let currentEditingPOId: string | null = null;
let currentEditingDispatchId: string | null = null;

// Temporary state for Create PO page mode and prefill data
let _formModeForCreatePage: 'create' | 'edit' | 'revise' = 'create';
let _formDataForCreatePage: PurchaseOrder | undefined = undefined;
let _originalIdForRevisionOnCreatePage: string | undefined = undefined;


// --- DOM Elements ---
const mainContent = document.getElementById('main-content')!;
const navbar = document.getElementById('navbar')!;

// --- Utility Functions ---
function generateId(prefix: 'PO' | 'D' | 'item'): string {
    if (prefix === 'PO') {
        const maxId = purchaseOrders.reduce((max, po) => Math.max(max, parseInt(po.id, 10) || 0), 0);
        return (maxId + 1).toString();
    }
    const now = new Date();
    const datePart = `${now.getFullYear()}${(now.getMonth() + 1).toString().padStart(2, '0')}${now.getDate().toString().padStart(2, '0')}`;
    const randomPart = Math.random().toString(36).substring(2, 6).toUpperCase();
    if (prefix === 'item') return `${prefix}-${randomPart}`;
    return `${prefix}-${datePart}-${randomPart}`;
}

function formatToDDMMYY(isoString?: string): string {
    if (!isoString) return 'N/A';
    try {
        const date = new Date(isoString);
        const day = date.getUTCDate().toString().padStart(2, '0');
        const month = (date.getUTCMonth() + 1).toString().padStart(2, '0'); // Month is 0-indexed
        const year = date.getUTCFullYear().toString().slice(-2); // Get last two digits of year
        return `${day}-${month}-${year}`;
    } catch (e) {
        console.error("Error formatting date to DDMMYY:", isoString, e);
        return 'Invalid Date';
    }
}

function formatToDDMMYY_HHMM(isoString?: string): string {
    if (!isoString) return 'N/A';
    try {
        const date = new Date(isoString);
        // Displaying in local time as it's more relevant for creation/modification time
        const day = date.getDate().toString().padStart(2, '0');
        const month = (date.getMonth() + 1).toString().padStart(2, '0');
        const year = date.getFullYear().toString().slice(-2);
        const hours = date.getHours().toString().padStart(2, '0');
        const minutes = date.getMinutes().toString().padStart(2, '0');
        return `${day}-${month}-${year} (${hours}:${minutes})`;
    } catch (e) {
        console.error("Error formatting date to DDMMYY_HHMM:", isoString, e);
        return 'Invalid Date';
    }
}


function escapeHTML(str: string | number | undefined | null): string {
    if (str === null || typeof str === 'undefined') {
        return '';
    }
    const text = String(str);
    return text.replace(/[&<>"']/g, function (match) {
        switch (match) {
            case '&': return '&amp;';
            case '<': return '&lt;';
            case '>': return '&gt;';
            case '"': return '&quot;';
            case "'": return '&#39;';
            default: return match;
        }
    });
}

function escapeCSVField(field: any): string {
    if (field === null || typeof field === 'undefined') {
        return '';
    }
    const stringField = String(field);
    if (stringField.includes(',') || stringField.includes('"') || stringField.includes('\n') || stringField.includes('\r')) {
        return `"${stringField.replace(/"/g, '""')}"`;
    }
    return stringField;
}

// Auto-uppercase listener
function toUpperCaseListener(event: Event) {
    const inputElement = event.target as HTMLInputElement | HTMLTextAreaElement;
    const originalSelectionStart = inputElement.selectionStart;
    const originalSelectionEnd = inputElement.selectionEnd;
    inputElement.value = inputElement.value.toUpperCase();
    if (originalSelectionStart !== null && originalSelectionEnd !== null) {
        // Restore cursor position
        inputElement.setSelectionRange(originalSelectionStart, originalSelectionEnd);
    }
}


// --- LocalStorage Persistence ---
function saveData(): void {
    localStorage.setItem('HITECH_purchaseOrders', JSON.stringify(purchaseOrders));
    localStorage.setItem('HITECH_dispatches', JSON.stringify(dispatches));
    localStorage.setItem('HITECH_dispatchFilterStartDate', dispatchFilterStartDate);
    localStorage.setItem('HITECH_dispatchFilterEndDate', dispatchFilterEndDate);
    localStorage.setItem('HITECH_predefinedMaterials', JSON.stringify(PREDEFINED_MATERIALS));
}

function loadData(): void {
    const storedPOs = localStorage.getItem('HITECH_purchaseOrders');
    if (storedPOs) {
        purchaseOrders = JSON.parse(storedPOs).map((po: PurchaseOrder) => ({
             ...po,
             // Ensure dispatchedQuantityByMaterial exists, for older data
             dispatchedQuantityByMaterial: po.dispatchedQuantityByMaterial || {}
        }));
    }
    const storedDispatches = localStorage.getItem('HITECH_dispatches');
    if (storedDispatches) {
        dispatches = JSON.parse(storedDispatches);
    }
    dispatchFilterStartDate = localStorage.getItem('HITECH_dispatchFilterStartDate') || '';
    dispatchFilterEndDate = localStorage.getItem('HITECH_dispatchFilterEndDate') || '';

    const storedPredefinedMaterials = localStorage.getItem('HITECH_predefinedMaterials');
    if (storedPredefinedMaterials) {
        try {
            const parsedMaterials = JSON.parse(storedPredefinedMaterials);
            if (Array.isArray(parsedMaterials) && parsedMaterials.every(m => typeof m === 'string')) {
                PREDEFINED_MATERIALS = parsedMaterials;
            } else {
                PREDEFINED_MATERIALS = [...DEFAULT_PREDEFINED_MATERIALS];
                localStorage.setItem('HITECH_predefinedMaterials', JSON.stringify(PREDEFINED_MATERIALS));
            }
        } catch (e) {
            console.error("Error parsing predefined materials from localStorage:", e);
            PREDEFINED_MATERIALS = [...DEFAULT_PREDEFINED_MATERIALS];
            localStorage.setItem('HITECH_predefinedMaterials', JSON.stringify(PREDEFINED_MATERIALS));
        }
    } else {
        PREDEFINED_MATERIALS = [...DEFAULT_PREDEFINED_MATERIALS];
        localStorage.setItem('HITECH_predefinedMaterials', JSON.stringify(PREDEFINED_MATERIALS));
    }
}

// --- Core Rendering Logic ---
function renderApp(): void {
    renderNavbar();
    updateNavActiveState();

    // Reset editing states when view changes, unless a modal is specifically managing them
    // also ensure temporary form state is not prematurely cleared if a modal is open and form is behind it.
    if (!document.querySelector('.modal[style*="display: block"]')) {
        if (currentView !== 'create-po') { // Only reset if navigating away from create-po
             currentEditingPOId = null; // Should be cleared by cancelPOFormEdit or by successful submit
             _formModeForCreatePage = 'create';
             _formDataForCreatePage = undefined;
             _originalIdForRevisionOnCreatePage = undefined;
        }
        currentEditingDispatchId = null;
    }


    switch (currentView) {
        case 'create-po':
            const mode = _formModeForCreatePage;
            const data = _formDataForCreatePage;
            const originalId = _originalIdForRevisionOnCreatePage;
            renderCreatePOForm(data, mode, originalId);

            // Reset temporary state only if it was consumed by renderCreatePOForm
            // This reset is now handled by the caller (promptEditPO, promptRevisePO) or renderCreatePOForm exit
            // _formModeForCreatePage = 'create';
            // _formDataForCreatePage = undefined;
            // _originalIdForRevisionOnCreatePage = undefined;
            break;
        case 'view-po':
            renderPOList('all');
            break;
        case 'view-dispatches':
            renderDispatchList();
            break;
        case 'pending-orders':
            renderPOList('pending');
            break;
        default:
            mainContent.innerHTML = '<p>Error: View not found.</p>';
    }
}

function renderNavbar(): void {
    navbar.innerHTML = `
        <button data-view="create-po" aria-label="Create New Purchase Order">Create PO</button>
        <button data-view="view-po" aria-label="View All Purchase Orders">All POs</button>
        <button data-view="pending-orders" aria-label="View Pending Purchase Orders">Pending Orders</button>
        <button data-view="view-dispatches" aria-label="View Dispatch Log">Dispatch Log</button>
    `;
    navbar.querySelectorAll('button').forEach(button => {
        button.addEventListener('click', () => {
            const newView = button.dataset.view as View;
            if (currentView === 'create-po' && newView !== 'create-po') {
                // If leaving create-po form, ensure temp states are reset
                 _formModeForCreatePage = 'create';
                 _formDataForCreatePage = undefined;
                 _originalIdForRevisionOnCreatePage = undefined;
                 currentEditingPOId = null; // Ensure edit state is also cleared
            }
            currentView = newView;
            renderApp();
        });
    });
}

function updateNavActiveState(): void {
    navbar.querySelectorAll('button').forEach(button => {
        if (button.dataset.view === currentView) {
            button.classList.add('active');
        } else {
            button.classList.remove('active');
        }
    });
}

// --- Purchase Order Creation / Editing ---
function renderCreatePOForm(
    poPrefillData?: PurchaseOrder,
    mode: 'create' | 'edit' | 'revise' = 'create',
    originalCancelledPOId?: string
): void {

    let formTitle: string;
    let submitButtonText: string;
    let isEditingOrRevising = mode === 'edit' || mode === 'revise';

    // Set currentEditingPOId only for actual 'edit' mode, for submit handler.
    // For 'revise', currentEditingPOId must be null to trigger create logic.
    if (mode === 'edit' && poPrefillData) {
        currentEditingPOId = poPrefillData.id;
        formTitle = `Edit Purchase Order: ${escapeHTML(poPrefillData.id)}`;
        submitButtonText = 'Update Purchase Order';
    } else if (mode === 'revise' && poPrefillData && originalCancelledPOId) {
        currentEditingPOId = null; // Critical: ensures handleCreatePOSubmit is called
        formTitle = `Revise Purchase Order (New from Cancelled PO: ${escapeHTML(originalCancelledPOId)})`;
        submitButtonText = 'Create Revised PO';
    } else { // mode === 'create'
        currentEditingPOId = null;
        formTitle = 'Create New Purchase Order';
        submitButtonText = 'Create Purchase Order';
        // poPrefillData will be undefined here
    }

    if (poPrefillData && poPrefillData.items) {
        // Deep clone items for form editing to avoid direct mutation
        poFormMaterialItems = JSON.parse(JSON.stringify(poPrefillData.items.map(item => ({...item, id: item.id || generateId('item') }))));
    } else {
        poFormMaterialItems = [{ id: generateId('item'), material: '', quantity: 0, rate: 0, gstPercentage: DEFAULT_GST_RATE }];
    }

    mainContent.innerHTML = `
        <div class="form-container">
            <h2 id="po-form-heading">${formTitle}</h2>
            <form id="po-form" aria-labelledby="po-form-heading">
                <div class="form-group">
                    <label for="partyName">Party Name:</label>
                    <input type="text" id="partyName" name="partyName" value="${poPrefillData ? escapeHTML(poPrefillData.partyName) : ''}" required>
                </div>
                <div class="form-group">
                    <label for="gstin">GSTIN:</label>
                    <input type="text" id="gstin" name="gstin" value="${poPrefillData ? escapeHTML(poPrefillData.gstin) : ''}" pattern="^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$" title="Enter valid GSTIN (e.g., 22AAAAA0000A1Z5)">
                </div>
                <div class="form-group">
                    <label for="salesmanName">Salesman Name:</label>
                    <input type="text" id="salesmanName" name="salesmanName" value="${poPrefillData ? escapeHTML(poPrefillData.salesmanName) : ''}" required>
                </div>
                <div class="form-group">
                    <label for="siteAddress">Site Address:</label>
                    <textarea id="siteAddress" name="siteAddress" rows="3" required>${poPrefillData ? escapeHTML(poPrefillData.siteAddress) : ''}</textarea>
                </div>
                <div class="form-group">
                    <label for="destination">Destination:</label>
                    <input type="text" id="destination" name="destination" value="${poPrefillData ? escapeHTML(poPrefillData.destination) : ''}" required>
                </div>
                <div class="form-group">
                    <label for="externalPoNumber">Optional PO Number (External Ref):</label>
                    <input type="text" id="externalPoNumber" name="externalPoNumber" value="${poPrefillData && poPrefillData.externalPoNumber ? escapeHTML(poPrefillData.externalPoNumber) : ''}">
                </div>

                <h3>Material Items</h3>
                <div id="material-items-container">
                    ${renderPOFormMaterialItemsHTML(mode === 'edit' ? poPrefillData : undefined)}
                </div>
                <button type="button" id="add-material-item" class="secondary" style="margin-top: 10px; margin-bottom:20px;" ${ (mode === 'edit' && poPrefillData && poPrefillData.status === 'Partially Dispatched') ? '' : ''}>+ Add Material</button>
                <hr style="margin: 20px 0;">
                <div class="form-group" style="text-align: right; font-weight: bold; font-size: 1.2em;">
                    Grand Total (incl. GST): ₹<span id="po-grand-total">0.00</span>
                </div>
                <button type="submit" class="primary">${submitButtonText}</button>
                ${isEditingOrRevising ? `<button type="button" onclick="window.cancelPOFormEdit()" class="secondary">Cancel</button>` : ''}
            </form>
        </div>
    `;

    // Attach uppercase listeners
    document.getElementById('partyName')?.addEventListener('input', toUpperCaseListener);
    document.getElementById('gstin')?.addEventListener('input', toUpperCaseListener);
    document.getElementById('salesmanName')?.addEventListener('input', toUpperCaseListener);
    document.getElementById('destination')?.addEventListener('input', toUpperCaseListener);
    document.getElementById('externalPoNumber')?.addEventListener('input', toUpperCaseListener);

    attachPOFormMaterialItemListeners(mode === 'edit' ? poPrefillData : undefined);
    poFormMaterialItems.forEach(item => updateItemCalculationsInForm(item.id)); // Initial calculation
    updatePOGrandTotalInForm();
    document.getElementById('add-material-item')?.addEventListener('click', () => handleAddPOFormMaterialItem(mode === 'edit' ? poPrefillData : undefined));

    document.getElementById('po-form')?.addEventListener('submit', (e) => {
        e.preventDefault();
        // currentEditingPOId is set by mode logic at the start of this function
        if (currentEditingPOId && mode === 'edit') {
            handleUpdatePOSubmit(currentEditingPOId, e);
        } else { // 'create' or 'revise'
            handleCreatePOSubmit(e);
        }
    });
}

(window as any).cancelPOFormEdit = ():void => {
    currentEditingPOId = null;
    _formModeForCreatePage = 'create';
    _formDataForCreatePage = undefined;
    _originalIdForRevisionOnCreatePage = undefined;
    currentView = 'view-po';
    renderApp();
};


function renderPOFormMaterialItemsHTML(poForEditContext?: PurchaseOrder): string {
    const gstOptions = [0, 3, 5, 12, 18, 28];
    // poForEditContext is only relevant if we are actually editing (not revising a cancelled one for new, not creating new)
    const isPartiallyDispatchedEdit = poForEditContext?.status === 'Partially Dispatched';

    return poFormMaterialItems.map((item, index) => {
        // For edit mode, check dispatched quantities to lock fields. For revise/create, this isn't applicable.
        const originalPOItem = poForEditContext?.items.find(i => i.material === item.material);
        const dispatchedQty = (poForEditContext && originalPOItem) ? (poForEditContext.dispatchedQuantityByMaterial[originalPOItem.material] || 0) : 0;
        const itemIsDispatched = dispatchedQty > 0;

        const disableMaterial = isPartiallyDispatchedEdit && itemIsDispatched;
        const disableRate = isPartiallyDispatchedEdit && itemIsDispatched;
        const disableGst = isPartiallyDispatchedEdit && itemIsDispatched;
        const minQuantity = (isPartiallyDispatchedEdit && itemIsDispatched) ? dispatchedQty : 0.01;


        return `
        <div class="material-item" data-item-id="${item.id}">
            <h4>Material ${index + 1} ${(poForEditContext && itemIsDispatched) ? `<span class="badge badge-info-light">Dispatched: ${dispatchedQty.toFixed(2)}</span>` : ''}</h4>
            <div class="form-group">
                <label for="material-${item.id}">Material Name:</label>
                <input type="text" id="material-${item.id}" class="material-name" list="material-suggestions-${item.id}" value="${escapeHTML(item.material)}" required data-item-id="${item.id}" ${disableMaterial ? 'readonly style="background-color:#e9ecef;"' : ''}>
                <datalist id="material-suggestions-${item.id}">
                    ${PREDEFINED_MATERIALS.map(mat => `<option value="${escapeHTML(mat)}"></option>`).join('')}
                </datalist>
            </div>
            <div class="form-group">
                <label for="quantity-${item.id}">Quantity (Min: ${minQuantity.toFixed(2)}):</label>
                <input type="number" id="quantity-${item.id}" class="material-quantity" value="${item.quantity}" min="${minQuantity.toFixed(2)}" step="0.01" required data-item-id="${item.id}">
            </div>
            <div class="form-group">
                <label for="rate-${item.id}">Rate:</label>
                <input type="number" id="rate-${item.id}" class="material-rate" value="${item.rate}" min="0.01" step="0.01" required data-item-id="${item.id}" ${disableRate ? 'readonly style="background-color:#e9ecef;"' : ''}>
            </div>
            <div class="form-group">
                <label for="gst-${item.id}">GST Percentage:</label>
                <select id="gst-${item.id}" class="material-gst" data-item-id="${item.id}" ${disableGst ? 'disabled style="background-color:#e9ecef;"' : ''}>
                    ${gstOptions.map(opt => `<option value="${opt}" ${item.gstPercentage === opt ? 'selected' : ''}>${opt}%</option>`).join('')}
                </select>
            </div>
            <div class="item-calculations">
                <p>Taxable Amount: ₹<span id="taxable-amount-${item.id}">0.00</span></p>
                <p>GST Amount: ₹<span id="gst-amount-${item.id}">0.00</span></p>
                <p><strong>Line Total (incl. GST): ₹<span id="line-total-${item.id}">0.00</span></strong></p>
            </div>
            ${poFormMaterialItems.length > 1 && !(isPartiallyDispatchedEdit && itemIsDispatched) ? `<button type="button" class="remove-material-item danger" data-item-id="${item.id}" aria-label="Remove Material ${index + 1}">Remove</button>` : ''}
        </div>
    `}).join('');
}

function attachPOFormMaterialItemListeners(poForEditContext?: PurchaseOrder): void {
    const container = document.getElementById('material-items-container')!;
    container.querySelectorAll('.remove-material-item').forEach(button => {
        button.addEventListener('click', (e) => {
            const itemId = (e.target as HTMLElement).dataset.itemId;
            if (itemId) handleRemovePOFormMaterialItem(itemId, poForEditContext);
        });
    });

    container.querySelectorAll('input[data-item-id], select[data-item-id]').forEach(inputEl => {
        const input = inputEl as HTMLInputElement | HTMLSelectElement;
        const eventType = (input.classList.contains('material-name') || input.classList.contains('material-quantity') || input.classList.contains('material-rate')) ? 'input' : 'change';

        input.addEventListener(eventType, () => {
            const itemId = input.dataset.itemId!;
            const itemIndex = poFormMaterialItems.findIndex(i => i.id === itemId);
            if (itemIndex > -1) {
                if (input.classList.contains('material-name')) poFormMaterialItems[itemIndex].material = input.value;
                if (input.classList.contains('material-quantity')) poFormMaterialItems[itemIndex].quantity = parseFloat(input.value) || 0;
                if (input.classList.contains('material-rate')) poFormMaterialItems[itemIndex].rate = parseFloat(input.value) || 0;
                if (input.classList.contains('material-gst')) poFormMaterialItems[itemIndex].gstPercentage = parseFloat(input.value) || 0;

                updateItemCalculationsInForm(itemId);
                updatePOGrandTotalInForm();
            }
        });
        if (input.classList.contains('material-name')) {
            input.addEventListener('input', toUpperCaseListener);
        }
        if (input.dataset.itemId) updateItemCalculationsInForm(input.dataset.itemId);
    });
}

function updateItemCalculationsInForm(itemId: string): void {
    const item = poFormMaterialItems.find(i => i.id === itemId);
    if (!item) return;

    const qty = item.quantity;
    const rate = item.rate;
    const gstPerc = item.gstPercentage;

    const taxableAmount = qty * rate;
    const gstAmount = (taxableAmount * gstPerc) / 100;
    const lineTotal = taxableAmount + gstAmount;

    document.getElementById(`taxable-amount-${itemId}`)!.textContent = taxableAmount.toFixed(2);
    document.getElementById(`gst-amount-${itemId}`)!.textContent = gstAmount.toFixed(2);
    document.getElementById(`line-total-${itemId}`)!.textContent = lineTotal.toFixed(2);
}

function updatePOGrandTotalInForm(): void {
    let grandTotal = 0;
    poFormMaterialItems.forEach(item => {
        const taxableAmount = item.quantity * item.rate;
        const gstAmount = (taxableAmount * item.gstPercentage) / 100;
        grandTotal += taxableAmount + gstAmount;
    });
    document.getElementById('po-grand-total')!.textContent = grandTotal.toFixed(2);
}


function refreshPOFormMaterialItemsUI(poForEditContext?: PurchaseOrder): void {
    const container = document.getElementById('material-items-container');
    if (container) {
        container.innerHTML = renderPOFormMaterialItemsHTML(poForEditContext);
        attachPOFormMaterialItemListeners(poForEditContext);
        poFormMaterialItems.forEach(item => updateItemCalculationsInForm(item.id));
        updatePOGrandTotalInForm();
    }
}

function handleAddPOFormMaterialItem(poForEditContext?: PurchaseOrder): void {
    poFormMaterialItems.push({ id: generateId('item'), material: '', quantity: 0, rate: 0, gstPercentage: DEFAULT_GST_RATE });
    refreshPOFormMaterialItemsUI(poForEditContext);
}

function handleRemovePOFormMaterialItem(itemId: string, poForEditContext?: PurchaseOrder): void {
    poFormMaterialItems = poFormMaterialItems.filter(item => item.id !== itemId);
    refreshPOFormMaterialItemsUI(poForEditContext);
}

function handleCreatePOSubmit(event: Event): void {
    event.preventDefault();
    const form = (event.target as HTMLFormElement);
    const formData = new FormData(form);

    const validItemsFromForm = poFormMaterialItems
        .filter(item => item.material.trim() !== '' && item.quantity > 0 && item.rate >= 0)
        .map(({ id, material, quantity, rate, gstPercentage }) => ({ material: material.toUpperCase(), quantity, rate, gstPercentage, id: id || generateId('item') }));


    if (validItemsFromForm.length === 0) {
        alert('Please add at least one valid material item with name and quantity.');
        return;
    }

    let totalAmountWithGST = 0;
    validItemsFromForm.forEach(item => {
        const taxable = item.quantity * item.rate;
        const gst = (taxable * item.gstPercentage) / 100;
        totalAmountWithGST += taxable + gst;
    });

    const dispatchedQuantityByMaterial: { [materialName: string]: number } = {};
    validItemsFromForm.forEach(item => {
        dispatchedQuantityByMaterial[item.material.toUpperCase()] = 0;
    });

    const externalPoNumberValue = formData.get('externalPoNumber') as string;

    const newPO: PurchaseOrder = {
        id: generateId('PO'),
        externalPoNumber: externalPoNumberValue ? externalPoNumberValue.trim() : undefined,
        partyName: (formData.get('partyName') as string).toUpperCase(),
        gstin: (formData.get('gstin') as string).toUpperCase(),
        salesmanName: (formData.get('salesmanName') as string).toUpperCase(),
        siteAddress: formData.get('siteAddress') as string,
        destination: (formData.get('destination') as string).toUpperCase(),
        items: validItemsFromForm,
        createdAt: new Date().toISOString(),
        status: 'Pending',
        totalAmount: totalAmountWithGST,
        dispatchedQuantityByMaterial: dispatchedQuantityByMaterial
    };

    newPO.items.forEach(item => {
        const materialNameUpper = item.material;
        if (!PREDEFINED_MATERIALS.includes(materialNameUpper)) {
            PREDEFINED_MATERIALS.push(materialNameUpper);
        }
    });

    purchaseOrders.unshift(newPO);
    saveData();
    alert('Purchase Order Created Successfully! System PO ID: ' + newPO.id);
    form.reset();
    poFormMaterialItems = []; // Reset for next creation
    currentEditingPOId = null; // Clear any potential editing ID
    _formModeForCreatePage = 'create'; // Reset form mode
    _formDataForCreatePage = undefined;
    _originalIdForRevisionOnCreatePage = undefined;
    currentView = 'view-po';
    renderApp();
}

function handleUpdatePOSubmit(poId: string, event: Event): void {
    event.preventDefault();
    const form = event.target as HTMLFormElement;
    const formData = new FormData(form);
    const poIndex = purchaseOrders.findIndex(p => p.id === poId);
    if (poIndex === -1) {
        alert('Error: Purchase Order not found for update.');
        return;
    }
    const existingPO = purchaseOrders[poIndex];

    const editedItemsFromForm = poFormMaterialItems
        .map(({ id, material, quantity, rate, gstPercentage }) => ({ material: material.toUpperCase(), quantity, rate, gstPercentage, id: id || generateId('item') }));

    if (editedItemsFromForm.length === 0) {
        alert('A Purchase Order must have at least one material item.');
        return;
    }

    // Validation for partially dispatched POs
    if (existingPO.status === 'Partially Dispatched' || existingPO.status === 'Completed') { // Also check completed in case it was over-dispatched
        for (const editedItem of editedItemsFromForm) {
            const originalItem = existingPO.items.find(i => i.material === editedItem.material);
            const dispatchedQty = existingPO.dispatchedQuantityByMaterial[editedItem.material] || 0;

            if (originalItem && dispatchedQty > 0) {
                if (editedItem.quantity < dispatchedQty) {
                    alert(`Error for ${editedItem.material}: Ordered quantity (${editedItem.quantity}) cannot be less than already dispatched quantity (${dispatchedQty}).`);
                    return;
                }
            }
        }
        for (const originalItem of existingPO.items) {
            const dispatchedQty = existingPO.dispatchedQuantityByMaterial[originalItem.material] || 0;
            if (dispatchedQty > 0 && !editedItemsFromForm.find(ei => ei.material === originalItem.material)) {
                alert(`Error: Cannot remove material ${originalItem.material} as it has been dispatched.`);
                return;
            }
        }
    }


    let totalAmountWithGST = 0;
    editedItemsFromForm.forEach(item => {
        const taxable = item.quantity * item.rate;
        const gst = (taxable * item.gstPercentage) / 100;
        totalAmountWithGST += taxable + gst;
    });


    const updatedPO: PurchaseOrder = {
        ...existingPO,
        externalPoNumber: (formData.get('externalPoNumber') as string)?.trim() || undefined,
        partyName: (formData.get('partyName') as string).toUpperCase(),
        gstin: (formData.get('gstin')as string).toUpperCase(),
        salesmanName: (formData.get('salesmanName')as string).toUpperCase(),
        siteAddress: formData.get('siteAddress') as string,
        destination: (formData.get('destination') as string).toUpperCase(),
        items: editedItemsFromForm,
        totalAmount: totalAmountWithGST,
    };

     updatedPO.items.forEach(item => {
        if (!PREDEFINED_MATERIALS.includes(item.material)) {
            PREDEFINED_MATERIALS.push(item.material);
        }
    });

    purchaseOrders[poIndex] = updatedPO;
    updatePOStatus(poId); // Re-evaluate status after edit, esp. if quantities changed
    saveData();
    alert(`Purchase Order ${poId} Updated Successfully!`);
    form.reset();
    poFormMaterialItems = [];
    currentEditingPOId = null;
    _formModeForCreatePage = 'create'; // Reset form mode
    _formDataForCreatePage = undefined;
    _originalIdForRevisionOnCreatePage = undefined;
    currentView = 'view-po';
    renderApp();
}

(window as any).promptEditPO = (poId: string): void => {
    const po = purchaseOrders.find(p => p.id === poId);
    if (!po) {
        alert('PO not found for editing.');
        return;
    }
    if (po.status === 'Cancelled') {
        alert(`Cannot edit a PO that is ${po.status}. Use 'Revise' to create a new PO from this one.`);
        return;
    }
    closeModal('po-details-modal');

    _formModeForCreatePage = 'edit';
    _formDataForCreatePage = po; // Pass the actual PO object for prefill
    currentEditingPOId = poId; // Set this for handleUpdatePOSubmit
    currentView = 'create-po';
    renderApp();
};

(window as any).promptRevisePO = (poId: string): void => {
    const sourcePO = purchaseOrders.find(p => p.id === poId);
    if (!sourcePO) {
        alert('PO not found for revision.');
        return;
    }
    if (sourcePO.status !== 'Cancelled') {
        alert(`PO is not Cancelled. Cannot revise. Current status: ${sourcePO.status}`);
        return;
    }
    closeModal('po-details-modal'); // Close if open

    // Create a deep copy for the form, reset/modify fields for a new PO
    const poDataForForm: PurchaseOrder = JSON.parse(JSON.stringify(sourcePO));

    // Reset fields that should be new for a revised PO
    // poDataForForm.id will be generated by handleCreatePOSubmit
    // poDataForForm.createdAt will be set by handleCreatePOSubmit
    poDataForForm.status = 'Pending'; // Initial status for the new PO
    poDataForForm.dispatchedQuantityByMaterial = {}; // Reset dispatches

    _formModeForCreatePage = 'revise';
    _formDataForCreatePage = poDataForForm; // This is the COPIED and MODIFIED data
    _originalIdForRevisionOnCreatePage = sourcePO.id; // For heading
    currentEditingPOId = null; // CRITICAL: Ensures handleCreatePOSubmit is called
    currentView = 'create-po';
    renderApp();
};


(window as any).promptCancelPO = (poId: string): void => {
    const po = purchaseOrders.find(p => p.id === poId);
    if (!po) {
        alert('PO not found for cancellation.');
        return;
    }
    if (po.status === 'Cancelled') {
        alert(`PO is already ${po.status} and cannot be cancelled again.`);
        return;
    }
    // Removed: if (po.status === 'Completed') ... to allow cancelling completed POs
    handleCancelPO(poId);
};

function handleCancelPO(poId: string): void {
    const poIndex = purchaseOrders.findIndex(p => p.id === poId);
    if (poIndex > -1) {
        purchaseOrders[poIndex].status = 'Cancelled';
        saveData();
        alert(`Purchase Order ${poId} has been cancelled.`);
        closeModal('po-details-modal');
        renderApp();
    }
}


// --- Purchase Order Listing ---
function renderPOList(type: 'all' | 'pending'): void {
    let title = '';
    let posToList: PurchaseOrder[] = [];

    if (type === 'all') {
        title = 'All Purchase Orders';
        posToList = purchaseOrders;
    } else { // 'pending'
        title = 'Pending Orders (Awaiting Initial Dispatch)';
        posToList = purchaseOrders.filter(po => po.status === 'Pending');
    }


    let content = `<div class="list-container"><h2>${escapeHTML(title)} (${posToList.length})</h2>`;

    if (posToList.length === 0) {
        if (type === 'all') {
             content += `<p>No purchase orders found. <button type="button" class="primary" id="go-create-po">Create one now?</button></p>`;
        } else { // 'pending'
             content += `<p>No orders awaiting initial dispatch. Check "All POs" for partially dispatched or completed orders.</p>`;
        }
    } else {
        content += `
            <div class="table-responsive-wrapper">
            <table aria-label="${escapeHTML(title)}">
                <thead>
                    <tr>
                        <th>System PO ID</th>
                        <th>Ext. PO Ref</th>
                        <th>Date</th>
                        <th>Party Name</th>
                        <th>Destination</th>
                        <th>Total (₹, incl. GST)</th>
                        <th>Status</th>
                        <th class="text-right">Actions</th>
                    </tr>
                </thead>
                <tbody>
                    ${posToList.map(po => `
                        <tr>
                            <td>${escapeHTML(po.id)}</td>
                            <td>${escapeHTML(po.externalPoNumber || 'N/A')}</td>
                            <td>${escapeHTML(formatToDDMMYY_HHMM(po.createdAt))}</td>
                            <td>${escapeHTML(po.partyName)}</td>
                            <td>${escapeHTML(po.destination)}</td>
                            <td class="text-right">${po.totalAmount.toFixed(2)}</td>
                            <td><span class="badge ${getBadgeClass(po.status)}">${escapeHTML(po.status)}</span></td>
                            <td class="actions-column text-right">
                                <button class="info" onclick="window.showPODetailsModal('${escapeHTML(po.id)}')" aria-label="View details for PO ${escapeHTML(po.id)}">Details</button>
                                ${(po.status !== 'Completed' && po.status !== 'Cancelled') ? `<button class="primary" onclick="window.showAddDispatchModal('${escapeHTML(po.id)}')" aria-label="Add dispatch for PO ${escapeHTML(po.id)}">Dispatch</button>` : ''}
                                ${(po.status !== 'Cancelled') ? `<button class="secondary" onclick="window.promptEditPO('${escapeHTML(po.id)}')" aria-label="Edit PO ${escapeHTML(po.id)}">Edit</button>` : ''}
                                ${(po.status === 'Cancelled') ? `<button class="info" onclick="window.promptRevisePO('${escapeHTML(po.id)}')" aria-label="Revise PO ${escapeHTML(po.id)}">Revise</button>` : ''}
                                ${(po.status !== 'Cancelled') ? `<button class="danger" onclick="window.promptCancelPO('${escapeHTML(po.id)}')" aria-label="Cancel PO ${escapeHTML(po.id)}">Cancel</button>` : ''}
                            </td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
            </div>
        `;
    }
    content += `</div>`;
    content += getPODetailsModalHTML() + getAddDispatchModalHTML() + getEditDispatchModalHTML();
    mainContent.innerHTML = content;

    document.getElementById('go-create-po')?.addEventListener('click', () => {
        _formModeForCreatePage = 'create'; // Ensure clean state for manual navigation
        _formDataForCreatePage = undefined;
        _originalIdForRevisionOnCreatePage = undefined;
        currentEditingPOId = null;
        currentView = 'create-po';
        renderApp();
    });
}

function getBadgeClass(status: PurchaseOrder['status']): string {
    switch (status) {
        case 'Pending': return 'badge-pending';
        case 'Partially Dispatched': return 'badge-partial';
        case 'Completed': return 'badge-completed';
        case 'Cancelled': return 'badge-cancelled';
        default: return '';
    }
}

// --- Modal Generic Functions ---
function closeModal(modalId: string): void {
    const modal = document.getElementById(modalId);
    if (modal) {
        modal.style.display = 'none';
        modal.setAttribute('aria-hidden', 'true');
        if (modalId === 'po-details-modal') {
            const contentEl = document.getElementById('po-details-content');
            if(contentEl) contentEl.innerHTML = '';
        } else if (modalId === 'add-dispatch-modal') {
            const formContainerEl = document.getElementById('add-dispatch-form-container');
            if(formContainerEl) formContainerEl.innerHTML = '';
             currentEditingDispatchId = null;
        } else if (modalId === 'edit-dispatch-modal') {
            const formContainerEl = document.getElementById('edit-dispatch-form-container');
            if(formContainerEl) formContainerEl.innerHTML = '';
            currentEditingDispatchId = null;
        }
    }
}
(window as any).closeModal = closeModal;


// --- PO Details Modal ---
function getPODetailsModalHTML(): string {
    return `
        <div id="po-details-modal" class="modal" aria-labelledby="po-details-modal-title" aria-hidden="true">
            <div class="modal-content large">
                <div class="modal-header">
                    <h2 id="po-details-modal-title">Purchase Order Details</h2>
                    <span class="close-button" onclick="window.closeModal('po-details-modal')" aria-label="Close PO Details">&times;</span>
                </div>
                <div id="po-details-content"></div>
                 <div class="modal-footer" id="po-details-modal-footer" style="text-align: right; margin-top: 20px;">
                    <!-- Action buttons can be added here by showPODetailsModal -->
                </div>
            </div>
        </div>
    `;
}

(window as any).showPODetailsModal = (poId: string): void => {
    const po = purchaseOrders.find(p => p.id === poId);
    if (!po) return;

    const modal = document.getElementById('po-details-modal')!;
    const contentEl = document.getElementById('po-details-content')!;
    const footerEl = document.getElementById('po-details-modal-footer')!;
    (document.getElementById('po-details-modal-title')!).textContent = `Details for System PO: ${escapeHTML(po.id)}`;

    let dispatchesHtml = '<h4>Dispatches for this PO:</h4>';
    const relatedDispatches = dispatches.filter(d => d.poId === po.id);
    if (relatedDispatches.length > 0) {
        dispatchesHtml += `
            <div class="table-responsive-wrapper">
            <table class="po-details-table">
                <thead><tr><th>Dispatch ID</th><th>Date</th><th>Vehicle</th><th>Items</th><th>Actions</th></tr></thead>
                <tbody>
                ${relatedDispatches.map(d => `
                    <tr>
                        <td>${escapeHTML(d.id)}</td>
                        <td>${escapeHTML(formatToDDMMYY(d.dispatchedAt))}</td>
                        <td>${escapeHTML(d.vehicleNumber)}</td>
                        <td><ul>${d.dispatchedItems.map(i => `<li>${escapeHTML(i.material)}: ${i.quantity.toFixed(2)}</li>`).join('')}</ul></td>
                        <td>
                            ${po.status !== 'Cancelled' ? `<button class="info small" onclick="window.showEditDispatchModal('${escapeHTML(d.id)}')" aria-label="Edit dispatch ${escapeHTML(d.id)}">Edit</button>` : 'N/A'}
                        </td>
                    </tr>
                `).join('')}
                </tbody>
            </table>
            </div>`;
    } else {
        dispatchesHtml += '<p>No dispatches recorded for this PO yet.</p>';
    }

    contentEl.innerHTML = `
        <p><strong>System PO ID:</strong> ${escapeHTML(po.id)}</p>
        ${po.externalPoNumber ? `<p><strong>External PO Ref:</strong> ${escapeHTML(po.externalPoNumber)}</p>` : ''}
        <p><strong>Party Name:</strong> ${escapeHTML(po.partyName)}</p>
        <p><strong>GSTIN:</strong> ${escapeHTML(po.gstin || 'N/A')}</p>
        <p><strong>Salesman:</strong> ${escapeHTML(po.salesmanName || 'N/A')}</p>
        <p><strong>Site Address:</strong> ${escapeHTML(po.siteAddress)}</p>
        <p><strong>Destination:</strong> ${escapeHTML(po.destination)}</p>
        <p><strong>Created At:</strong> ${escapeHTML(formatToDDMMYY_HHMM(po.createdAt))}</p>
        <p><strong>Status:</strong> <span class="badge ${getBadgeClass(po.status)}">${escapeHTML(po.status)}</span></p>
        <p><strong>Grand Total (incl. GST):</strong> ₹${po.totalAmount.toFixed(2)}</p>
        <h4>Material Items:</h4>
        <div class="table-responsive-wrapper">
        <table class="po-details-table">
            <thead>
                <tr>
                    <th>Material</th>
                    <th>Qty Ordered</th>
                    <th>Rate (₹)</th>
                    <th>Taxable (₹)</th>
                    <th>GST %</th>
                    <th>GST Amt (₹)</th>
                    <th>Line Total (₹)</th>
                    <th>Dispatched Qty</th>
                    <th>Pending Qty</th>
                </tr>
            </thead>
            <tbody>
                ${po.items.map(item => {
                    const taxableAmount = item.quantity * item.rate;
                    const gstAmount = (taxableAmount * item.gstPercentage) / 100;
                    const lineTotal = taxableAmount + gstAmount;
                    const dispatchedQty = po.dispatchedQuantityByMaterial[item.material.toUpperCase()] || 0;
                    const pendingQty = item.quantity - dispatchedQty;
                    return `
                    <tr>
                        <td>${escapeHTML(item.material)}</td>
                        <td class="text-right">${item.quantity.toFixed(2)}</td>
                        <td class="text-right">${item.rate.toFixed(2)}</td>
                        <td class="text-right">${taxableAmount.toFixed(2)}</td>
                        <td class="text-right">${item.gstPercentage}%</td>
                        <td class="text-right">${gstAmount.toFixed(2)}</td>
                        <td class="text-right">${lineTotal.toFixed(2)}</td>
                        <td class="text-right">${dispatchedQty.toFixed(2)}</td>
                        <td class="text-right ${pendingQty < 0 ? 'text-danger' : ''}">${pendingQty.toFixed(2)}${pendingQty < 0 ? ' (Over)' : ''}</td>
                    </tr>
                `}).join('')}
            </tbody>
        </table>
        </div>
        <hr style="margin: 20px 0;">
        ${dispatchesHtml}
    `;

    footerEl.innerHTML = ''; // Clear previous buttons
    if (po.status !== 'Completed' && po.status !== 'Cancelled') {
        const dispatchButton = document.createElement('button');
        dispatchButton.className = 'primary';
        dispatchButton.textContent = 'Add Dispatch';
        dispatchButton.setAttribute('aria-label', `Add dispatch for PO ${po.id}`);
        dispatchButton.onclick = () => (window as any).showAddDispatchModal(po.id);
        footerEl.appendChild(dispatchButton);
    }
     if (po.status !== 'Cancelled') {
        const editButton = document.createElement('button');
        editButton.className = 'secondary';
        editButton.textContent = 'Edit PO';
        editButton.setAttribute('aria-label', `Edit PO ${po.id}`);
        editButton.onclick = () => (window as any).promptEditPO(po.id);
        footerEl.appendChild(editButton);
    }

    if (po.status === 'Cancelled') {
        const reviseButton = document.createElement('button');
        reviseButton.className = 'info';
        reviseButton.textContent = 'Revise PO';
        reviseButton.setAttribute('aria-label', `Revise PO ${po.id}`);
        reviseButton.onclick = () => (window as any).promptRevisePO(po.id);
        footerEl.appendChild(reviseButton);
    } else { // Only show Cancel PO if not already Cancelled
        const cancelButton = document.createElement('button');
        cancelButton.className = 'danger';
        cancelButton.textContent = 'Cancel PO';
        cancelButton.setAttribute('aria-label', `Cancel PO ${po.id}`);
        cancelButton.onclick = () => (window as any).promptCancelPO(po.id);
        footerEl.appendChild(cancelButton);
    }


    modal.style.display = 'block';
    modal.setAttribute('aria-hidden', 'false');
    (modal.querySelector('.close-button') as HTMLElement)?.focus();
};


// --- Add Dispatch Modal ---
function getAddDispatchModalHTML(): string {
    return `
        <div id="add-dispatch-modal" class="modal" aria-labelledby="add-dispatch-modal-title" aria-hidden="true">
            <div class="modal-content">
                <div class="modal-header">
                    <h2 id="add-dispatch-modal-title">Add Dispatch</h2>
                    <span class="close-button" onclick="window.closeModal('add-dispatch-modal')" aria-label="Close Add Dispatch Form">&times;</span>
                </div>
                <div id="add-dispatch-form-container"></div>
            </div>
        </div>
    `;
}

(window as any).showAddDispatchModal = (poId: string): void => {
    const po = purchaseOrders.find(p => p.id === poId);
    if (!po) {
        alert('Error: Purchase Order not found.');
        return;
    }
    if (po.status === 'Completed' || po.status === 'Cancelled') {
         alert(`Cannot add dispatch to a PO that is ${po.status}.`);
        return;
    }


    const modal = document.getElementById('add-dispatch-modal')!;
    const formContainer = document.getElementById('add-dispatch-form-container')!;
    (document.getElementById('add-dispatch-modal-title')!).textContent = `Add Dispatch for PO: ${escapeHTML(po.id)}`;

    let itemsHtml = '';
    po.items.forEach(item => {
        const materialKey = item.material.toUpperCase();
        const totalDispatchedForMaterial = po.dispatchedQuantityByMaterial[materialKey] || 0;
        const orderedQty = item.quantity;
        // const remainingOnPO = orderedQty - totalDispatchedForMaterial; // Can be negative if over-dispatched

        // Show all items from PO, regardless if fully dispatched or over-dispatched, user can still add more.
        itemsHtml += `
            <div class="form-group material-dispatch-item">
                <label for="dispatch-qty-${materialKey.replace(/\s+/g, '-').replace(/[^a-zA-Z0-9-]/g, '')}">
                    ${escapeHTML(item.material)}
                    (Ordered: ${orderedQty.toFixed(2)}, Total Dispatched: ${totalDispatchedForMaterial.toFixed(2)})
                </label>
                <input type="number" id="dispatch-qty-${materialKey.replace(/\s+/g, '-').replace(/[^a-zA-Z0-9-]/g, '')}"
                       name="dispatch_qty_${escapeHTML(materialKey)}"
                       min="0" step="0.01" value="0" required
                       aria-label="Dispatch quantity for ${escapeHTML(item.material)}">
                <input type="hidden" name="material_name" value="${escapeHTML(materialKey)}">
            </div>
        `;
    });

    if (po.items.length === 0) {
        itemsHtml = '<p>No materials listed in this PO to dispatch against.</p>';
    }

    const today = new Date().toISOString().split('T')[0];

    formContainer.innerHTML = `
        <form id="add-dispatch-form" data-poid="${escapeHTML(poId)}">
            <p><strong>Party:</strong> ${escapeHTML(po.partyName)}</p>
            <p><strong>Destination:</strong> ${escapeHTML(po.destination)}</p>
             <div class="form-group">
                <label for="dispatchDate">Dispatch Date:</label>
                <input type="date" id="dispatchDate" name="dispatchDate" value="${today}" required>
            </div>
            <div class="form-group">
                <label for="vehicleNumber">Vehicle Number:</label>
                <input type="text" id="vehicleNumber" name="vehicleNumber" required>
            </div>
            <div class="form-group">
                <label for="driverContact">Driver Contact Number:</label>
                <input type="tel" id="driverContact" name="driverContact">
            </div>
            <div class="form-group">
                <label for="invoiceNumber">Invoice Number (Optional):</label>
                <input type="text" id="invoiceNumber" name="invoiceNumber">
            </div>
            <div class="form-group">
                <label for="transporterName">Transporter Name (Optional):</label>
                <input type="text" id="transporterName" name="transporterName">
            </div>
            <h4>Dispatch Quantities:</h4>
            ${itemsHtml}
            <div class="form-group text-right mt-2">
                <button type="submit" class="primary" ${ po.items.length === 0 ? 'disabled' : ''}>Confirm Dispatch</button>
            </div>
        </form>
    `;

    document.getElementById('vehicleNumber')?.addEventListener('input', toUpperCaseListener);
    document.getElementById('invoiceNumber')?.addEventListener('input', toUpperCaseListener);
    document.getElementById('transporterName')?.addEventListener('input', toUpperCaseListener);

    document.getElementById('add-dispatch-form')?.addEventListener('submit', handleAddDispatchSubmit);

    modal.style.display = 'block';
    modal.setAttribute('aria-hidden', 'false');
    (modal.querySelector('.close-button, input:not([disabled]):not([readonly]), button:not([disabled])') as HTMLElement)?.focus();
};

function handleAddDispatchSubmit(event: Event): void {
    event.preventDefault();
    const form = event.target as HTMLFormElement;
    const poId = form.dataset.poid;
    const po = purchaseOrders.find(p => p.id === poId);

    if (!po) {
        alert('Critical Error: PO not found during dispatch submission.');
        return;
    }
     if (po.status === 'Completed' || po.status === 'Cancelled') {
         alert(`Cannot add dispatch: PO is ${po.status}.`);
        return;
    }


    const dispatchedItemsFromForm: { material: string; quantity: number }[] = [];
    let totalDispatchedThisTime = 0;
    let validationError = false;

    form.querySelectorAll('.material-dispatch-item').forEach(itemDiv => {
        const materialNameInput = itemDiv.querySelector('input[name="material_name"]') as HTMLInputElement;
        const quantityInput = itemDiv.querySelector('input[type="number"]') as HTMLInputElement;
        const label = itemDiv.querySelector('label');


        if (materialNameInput && quantityInput) {
            const materialName = materialNameInput.value;
            const quantity = parseFloat(quantityInput.value);
            // const maxQuantity = parseFloat(quantityInput.max); // Max attribute removed

            if (label) label.classList.remove('error-text');


            if (quantity < 0) {
                 alert(`Error for ${materialName}: Dispatch quantity cannot be negative.`);
                 if (label) label.classList.add('error-text');
                 quantityInput.focus();
                 validationError = true;
                 return; // Exit forEach callback for this item
            }
            // Removed max quantity validation:
            // if (quantity > maxQuantity) { ... }

            if (quantity > 0) {
                dispatchedItemsFromForm.push({ material: materialName, quantity: quantity });
                totalDispatchedThisTime += quantity;
            }
        }
    });

    if (validationError) {
      return; // Stop submission if any item had validation error
    }

    if (totalDispatchedThisTime === 0 && dispatchedItemsFromForm.length === 0 && po.items.length > 0) {
        alert('Please enter a dispatch quantity greater than zero for at least one material.');
        return;
    }
    if (po.items.length === 0 && dispatchedItemsFromForm.length === 0) {
         alert('Cannot dispatch as PO has no items.');
        closeModal('add-dispatch-modal');
        return;
    }


    const dispatchDateValue = (form.elements.namedItem('dispatchDate') as HTMLInputElement).value;
    const [year, month, day] = dispatchDateValue.split('-').map(Number);
    const dispatchDateTimeUTC = new Date(Date.UTC(year, month - 1, day, 0, 0, 0, 0));
    const dispatchedAtISO = dispatchDateTimeUTC.toISOString();


    const newDispatch: Dispatch = {
        id: generateId('D'),
        poId: poId!,
        vehicleNumber: ((form.elements.namedItem('vehicleNumber') as HTMLInputElement).value).toUpperCase(),
        driverContact: (form.elements.namedItem('driverContact') as HTMLInputElement).value,
        invoiceNumber: ((form.elements.namedItem('invoiceNumber') as HTMLInputElement).value.trim() || undefined)?.toUpperCase(),
        transporterName: ((form.elements.namedItem('transporterName') as HTMLInputElement).value.trim() || undefined)?.toUpperCase(),
        dispatchedItems: dispatchedItemsFromForm,
        dispatchedAt: dispatchedAtISO
    };

    if (newDispatch.dispatchedItems.length === 0 && totalDispatchedThisTime === 0) {
        alert('No items were dispatched. Please enter quantities greater than 0.');
        return;
    }


    dispatches.unshift(newDispatch);
    dispatches.sort((a, b) => new Date(b.dispatchedAt).getTime() - new Date(a.dispatchedAt).getTime());


    dispatchedItemsFromForm.forEach(dispItem => {
        const materialKey = dispItem.material.toUpperCase();
        po.dispatchedQuantityByMaterial[materialKey] = (po.dispatchedQuantityByMaterial[materialKey] || 0) + dispItem.quantity;
    });

    updatePOStatus(po.id);

    saveData();
    alert('Dispatch added successfully!');
    closeModal('add-dispatch-modal');
    renderApp();
}


// --- Edit Dispatch Modal & Logic ---
function getEditDispatchModalHTML(): string {
    return `
        <div id="edit-dispatch-modal" class="modal" aria-labelledby="edit-dispatch-modal-title" aria-hidden="true">
            <div class="modal-content">
                <div class="modal-header">
                    <h2 id="edit-dispatch-modal-title">Edit Dispatch</h2>
                    <span class="close-button" onclick="window.closeModal('edit-dispatch-modal')" aria-label="Close Edit Dispatch Form">&times;</span>
                </div>
                <div id="edit-dispatch-form-container"></div>
            </div>
        </div>
    `;
}

(window as any).showEditDispatchModal = (dispatchId: string): void => {
    currentEditingDispatchId = dispatchId;
    const dispatch = dispatches.find(d => d.id === dispatchId);
    if (!dispatch) {
        alert('Error: Dispatch not found.');
        currentEditingDispatchId = null;
        return;
    }
    const po = purchaseOrders.find(p => p.id === dispatch.poId);
    if (!po) {
        alert('Error: Associated Purchase Order not found.');
        currentEditingDispatchId = null;
        return;
    }
    if (po.status === 'Cancelled') {
        alert('Cannot edit dispatches for a Cancelled PO.');
        currentEditingDispatchId = null;
        return;
    }


    const modal = document.getElementById('edit-dispatch-modal')!;
    const formContainer = document.getElementById('edit-dispatch-form-container')!;
    (document.getElementById('edit-dispatch-modal-title')!).textContent = `Edit Dispatch: ${escapeHTML(dispatch.id)} (for PO: ${escapeHTML(po.id)})`;

    let itemsHtml = '';
    po.items.forEach(poItem => {
        const materialKey = poItem.material.toUpperCase();
        const dispatchItem = dispatch.dispatchedItems.find(di => di.material.toUpperCase() === materialKey);
        const currentDispatchQtyForItem = dispatchItem ? dispatchItem.quantity : 0;
        const orderedQty = poItem.quantity;
        const totalDispatchedForMaterialInPO = po.dispatchedQuantityByMaterial[materialKey] || 0;
        // const dispatchedByOtherDispatches = totalDispatchedForMaterialInPO - currentDispatchQtyForItem;
        // const maxEditableConsideringPOOrder = orderedQty - dispatchedByOtherDispatches; // This logic is removed as we allow over-dispatch

        if (dispatchItem) { // Only allow editing items that were part of this original dispatch
             itemsHtml += `
                <div class="form-group material-dispatch-item">
                    <label for="edit-dispatch-qty-${materialKey.replace(/\s+/g, '-').replace(/[^a-zA-Z0-9-]/g, '')}">
                        ${escapeHTML(poItem.material)}
                        (Ordered: ${orderedQty.toFixed(2)}, This Dispatch Had: ${currentDispatchQtyForItem.toFixed(2)})
                    </label>
                    <input type="number" id="edit-dispatch-qty-${materialKey.replace(/\s+/g, '-').replace(/[^a-zA-Z0-9-]/g, '')}"
                           name="dispatch_qty_${escapeHTML(materialKey)}"
                           min="0" step="0.01" value="${currentDispatchQtyForItem.toFixed(2)}" required
                           aria-label="New dispatch quantity for ${escapeHTML(poItem.material)}">
                    <input type="hidden" name="material_name" value="${escapeHTML(materialKey)}">
                    <input type="hidden" name="original_dispatch_qty_${escapeHTML(materialKey)}" value="${currentDispatchQtyForItem.toFixed(2)}">
                </div>
            `;
        }
    });
     if (itemsHtml === '') {
        itemsHtml = '<p>No editable items found in this dispatch. This might indicate an empty dispatch or items that are no longer in the PO.</p>';
    }


    const dispatchDate = new Date(dispatch.dispatchedAt).toISOString().split('T')[0];

    formContainer.innerHTML = `
        <form id="edit-dispatch-form" data-dispatchid="${escapeHTML(dispatchId)}">
            <p><strong>Party:</strong> ${escapeHTML(po.partyName)}</p>
            <p><strong>Destination:</strong> ${escapeHTML(po.destination)}</p>
             <div class="form-group">
                <label for="editDispatchDate">Dispatch Date:</label>
                <input type="date" id="editDispatchDate" name="dispatchDate" value="${dispatchDate}" required>
            </div>
            <div class="form-group">
                <label for="editVehicleNumber">Vehicle Number:</label>
                <input type="text" id="editVehicleNumber" name="vehicleNumber" value="${escapeHTML(dispatch.vehicleNumber)}" required>
            </div>
            <div class="form-group">
                <label for="editDriverContact">Driver Contact Number:</label>
                <input type="tel" id="editDriverContact" name="driverContact" value="${escapeHTML(dispatch.driverContact)}">
            </div>
            <div class="form-group">
                <label for="editInvoiceNumber">Invoice Number (Optional):</label>
                <input type="text" id="editInvoiceNumber" name="invoiceNumber" value="${escapeHTML(dispatch.invoiceNumber || '')}">
            </div>
            <div class="form-group">
                <label for="editTransporterName">Transporter Name (Optional):</label>
                <input type="text" id="editTransporterName" name="transporterName" value="${escapeHTML(dispatch.transporterName || '')}">
            </div>
            <h4>Dispatch Quantities:</h4>
            ${itemsHtml}
            <div class="form-group text-right mt-2">
                <button type="submit" class="primary" ${itemsHtml.startsWith("<p>No editable items") ? 'disabled' : ''}>Update Dispatch</button>
            </div>
        </form>
    `;

    document.getElementById('editVehicleNumber')?.addEventListener('input', toUpperCaseListener);
    document.getElementById('editInvoiceNumber')?.addEventListener('input', toUpperCaseListener);
    document.getElementById('editTransporterName')?.addEventListener('input', toUpperCaseListener);

    document.getElementById('edit-dispatch-form')?.addEventListener('submit', handleUpdateDispatchSubmit);

    modal.style.display = 'block';
    modal.setAttribute('aria-hidden', 'false');
    (modal.querySelector('.close-button, input:not([disabled]):not([readonly]), button:not([disabled])') as HTMLElement)?.focus();
};

function handleUpdateDispatchSubmit(event: Event): void {
    event.preventDefault();
    const form = event.target as HTMLFormElement;
    const dispatchId = form.dataset.dispatchid;

    const dispatchIndex = dispatches.findIndex(d => d.id === dispatchId);
    if (dispatchIndex === -1) {
        alert('Critical Error: Dispatch not found for update.');
        return;
    }
    const originalDispatch = dispatches[dispatchIndex];
    const po = purchaseOrders.find(p => p.id === originalDispatch.poId);
    if (!po) {
        alert('Critical Error: Associated PO not found.');
        return;
    }
     if (po.status === 'Cancelled') {
        alert('Cannot update dispatch: PO is Cancelled.');
        return;
    }

    const updatedDispatchedItems: { material: string; quantity: number }[] = [];
    let validationError = false;
    const quantityChanges: { material: string; delta: number }[] = []; // To update PO's total dispatched

    form.querySelectorAll('.material-dispatch-item').forEach(itemDiv => {
        const materialNameInput = itemDiv.querySelector('input[name="material_name"]') as HTMLInputElement;
        const quantityInput = itemDiv.querySelector('input[type="number"]') as HTMLInputElement;
        const originalQtyInput = itemDiv.querySelector(`input[name^="original_dispatch_qty_"]`) as HTMLInputElement;
        const label = itemDiv.querySelector('label');

        if (materialNameInput && quantityInput && originalQtyInput) {
            const materialName = materialNameInput.value;
            const newQuantity = parseFloat(quantityInput.value);
            const originalQuantityInThisDispatch = parseFloat(originalQtyInput.value);
            // const maxQuantity = parseFloat(quantityInput.max); // Max attribute removed

            if(label) label.classList.remove('error-text');

            if (newQuantity < 0) {
                alert(`Error for ${materialName}: Dispatch quantity cannot be negative.`);
                if (label) label.classList.add('error-text');
                quantityInput.focus();
                validationError = true;
                return;
            }
            // Removed max quantity validation:
            // if (newQuantity > maxQuantity) { ... }

            if (newQuantity > 0) {
                updatedDispatchedItems.push({ material: materialName, quantity: newQuantity });
            }
            quantityChanges.push({ material: materialName, delta: newQuantity - originalQuantityInThisDispatch });
        }
    });

    if (validationError) return;

    if (updatedDispatchedItems.length === 0 && originalDispatch.dispatchedItems.some(item => item.quantity > 0) ) {
        if (!confirm("You are setting all item quantities to zero for this dispatch. This will effectively remove items from this dispatch. Continue?")) {
            return;
        }
    }


    const dispatchDateValue = (form.elements.namedItem('dispatchDate') as HTMLInputElement).value;
    const [year, month, day] = dispatchDateValue.split('-').map(Number);
    const dispatchDateTimeUTC = new Date(Date.UTC(year, month - 1, day, 0, 0, 0, 0));

    const updatedDispatch: Dispatch = {
        ...originalDispatch,
        vehicleNumber: ((form.elements.namedItem('vehicleNumber') as HTMLInputElement).value).toUpperCase(),
        driverContact: (form.elements.namedItem('driverContact') as HTMLInputElement).value,
        invoiceNumber: ((form.elements.namedItem('invoiceNumber') as HTMLInputElement).value.trim() || undefined)?.toUpperCase(),
        transporterName: ((form.elements.namedItem('transporterName') as HTMLInputElement).value.trim() || undefined)?.toUpperCase(),
        dispatchedItems: updatedDispatchedItems,
        dispatchedAt: dispatchDateTimeUTC.toISOString()
    };

    dispatches[dispatchIndex] = updatedDispatch;
    dispatches.sort((a, b) => new Date(b.dispatchedAt).getTime() - new Date(a.dispatchedAt).getTime());

    // Update PO's dispatchedQuantityByMaterial
    quantityChanges.forEach(change => {
        const materialKey = change.material.toUpperCase();
        po.dispatchedQuantityByMaterial[materialKey] = (po.dispatchedQuantityByMaterial[materialKey] || 0) + change.delta;
        if (po.dispatchedQuantityByMaterial[materialKey] < 0) po.dispatchedQuantityByMaterial[materialKey] = 0;
    });

    updatePOStatus(po.id);

    saveData();
    alert(`Dispatch ${dispatchId} updated successfully!`);
    closeModal('edit-dispatch-modal');
    currentEditingDispatchId = null;

    const poDetailsModal = document.getElementById('po-details-modal');
    if (poDetailsModal && poDetailsModal.style.display === 'block') {
        const displayedPOIdText = (document.getElementById('po-details-modal-title') as HTMLElement).textContent;
        const poIdMatch = displayedPOIdText?.match(/PO: (\S+)\)?/); // Adjusted regex to find PO ID
        const displayedPOId = poIdMatch ? poIdMatch[1] : null;

        if (displayedPOId === po.id) {
            (window as any).showPODetailsModal(po.id);
        } else {
            renderApp();
        }
    } else {
        renderApp();
    }
}


function updatePOStatus(poId: string): void {
    const poIndex = purchaseOrders.findIndex(p => p.id === poId);
    if (poIndex === -1) return;
    const po = purchaseOrders[poIndex];

    if (po.status === 'Cancelled') return;

    let allItemsMeetOrExceedOrder = true;
    if (po.items.length > 0) {
        for (const item of po.items) {
            const materialKey = item.material.toUpperCase();
            if (typeof po.dispatchedQuantityByMaterial[materialKey] === 'undefined') {
                 po.dispatchedQuantityByMaterial[materialKey] = 0;
            }
            if ((po.dispatchedQuantityByMaterial[materialKey] || 0) < item.quantity) {
                allItemsMeetOrExceedOrder = false;
                break;
            }
        }
    } else {
        allItemsMeetOrExceedOrder = true; // An empty PO is considered "completed" if no items to dispatch
    }

    let totalDispatchedEver = 0;
    for (const materialKey in po.dispatchedQuantityByMaterial) {
        totalDispatchedEver += po.dispatchedQuantityByMaterial[materialKey];
    }

    if (allItemsMeetOrExceedOrder) {
        po.status = 'Completed';
    } else if (totalDispatchedEver > 0) {
        po.status = 'Partially Dispatched';
    } else {
        po.status = 'Pending';
    }
}


// --- Dispatch Log ---
interface DisplayDispatchLine {
    dispatchId: string;
    poId: string;
    partyName: string;
    salesmanName: string;
    dispatchedAt: string;
    originalDispatchedAtISO: string;
    invoiceNumber?: string;
    vehicleNumber: string;
    driverContact?: string;
    materialName: string;
    quantityDispatched: number;
    transporterName?: string;
    itemTotalAmount: number;
    destination: string;
}

function downloadDispatchLogAsCSV(linesToDownload: DisplayDispatchLine[]): void {
    if (linesToDownload.length === 0) {
        alert('No dispatch data to download.');
        return;
    }

    const headers = [
        "Dispatch ID", "PO ID", "Party Name", "Salesman", "Date", "Invoice No.",
        "Vehicle No.", "Driver Contact", "Item Name", "Quantity",
        "Transporter", "Destination", "Item Line Amt. (₹)"
    ];

    const csvRows = [
        headers.map(header => escapeCSVField(header)).join(',')
    ];

    const sortedLinesToDownload = [...linesToDownload].sort((a,b) =>
        new Date(b.originalDispatchedAtISO).getTime() - new Date(a.originalDispatchedAtISO).getTime()
    );


    sortedLinesToDownload.forEach(line => {
        const row = [
            escapeCSVField(line.dispatchId),
            escapeCSVField(line.poId),
            escapeCSVField(line.partyName),
            escapeCSVField(line.salesmanName),
            escapeCSVField(line.dispatchedAt),
            escapeCSVField(line.invoiceNumber),
            escapeCSVField(line.vehicleNumber),
            escapeCSVField(line.driverContact),
            escapeCSVField(line.materialName),
            escapeCSVField(line.quantityDispatched.toFixed(2)),
            escapeCSVField(line.transporterName),
            escapeCSVField(line.destination),
            escapeCSVField(line.itemTotalAmount.toFixed(2))
        ];
        csvRows.push(row.join(','));
    });

    const csvString = csvRows.join('\n');
    const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');

    if (link.download !== undefined) {
        const url = URL.createObjectURL(blob);
        link.setAttribute('href', url);
        const today = new Date().toISOString().split('T')[0].replace(/-/g,'');
        const filterStart = dispatchFilterStartDate ? dispatchFilterStartDate.replace(/-/g,'') : 'all';
        const filterEnd = dispatchFilterEndDate ? dispatchFilterEndDate.replace(/-/g,'') : 'all';

        link.setAttribute('download', `dispatch_log_${filterStart}_to_${filterEnd}_${today}.csv`);
        link.style.visibility = 'hidden';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
    } else {
        alert('CSV download is not supported by your browser.');
    }
}

function downloadItemSummaryCSV(summaryData: { [materialName: string]: number }): void {
    const sortedEntries = Object.entries(summaryData).sort((a, b) => b[1] - a[1]);
    if (sortedEntries.length === 0) {
        alert('No item summary data to download.');
        return;
    }

    const headers = ["Material Name", "Total Quantity Dispatched"];
    const csvRows = [headers.map(header => escapeCSVField(header)).join(',')];

    sortedEntries.forEach(([material, quantity]) => {
        const row = [
            escapeCSVField(material),
            escapeCSVField(quantity.toFixed(2))
        ];
        csvRows.push(row.join(','));
    });

    const csvString = csvRows.join('\n');
    const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const today = new Date().toISOString().split('T')[0].replace(/-/g, '');
    link.setAttribute('download', `item_dispatch_summary_${today}.csv`);
    link.href = URL.createObjectURL(blob);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
}

function downloadSalesmanSummaryCSV(summaryData: { [salesmanName: string]: number }): void {
    const sortedEntries = Object.entries(summaryData).sort((a, b) => b[1] - a[1]);
    if (sortedEntries.length === 0) {
        alert('No salesman summary data to download.');
        return;
    }

    const headers = ["Salesman Name", "Total Quantity Dispatched"];
    const csvRows = [headers.map(header => escapeCSVField(header)).join(',')];

    sortedEntries.forEach(([salesman, quantity]) => {
        const row = [
            escapeCSVField(salesman),
            escapeCSVField(quantity.toFixed(2))
        ];
        csvRows.push(row.join(','));
    });

    const csvString = csvRows.join('\n');
    const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const today = new Date().toISOString().split('T')[0].replace(/-/g, '');
    link.setAttribute('download', `salesman_dispatch_summary_${today}.csv`);
    link.href = URL.createObjectURL(blob);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
}

function downloadPartySummaryCSV(summaryData: { [partyName: string]: number }): void {
    const sortedEntries = Object.entries(summaryData).sort((a, b) => b[1] - a[1]);
    if (sortedEntries.length === 0) {
        alert('No party summary data to download.');
        return;
    }

    const headers = ["Party Name", "Total Quantity Dispatched"];
    const csvRows = [headers.map(header => escapeCSVField(header)).join(',')];

    sortedEntries.forEach(([party, quantity]) => {
        const row = [
            escapeCSVField(party),
            escapeCSVField(quantity.toFixed(2))
        ];
        csvRows.push(row.join(','));
    });

    const csvString = csvRows.join('\n');
    const blob = new Blob([csvString], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const today = new Date().toISOString().split('T')[0].replace(/-/g, '');
    link.setAttribute('download', `party_dispatch_summary_${today}.csv`);
    link.href = URL.createObjectURL(blob);
    link.style.visibility = 'hidden';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
}


function renderDispatchList(): void {
    let filteredDispatches = [...dispatches];

    if (dispatchFilterStartDate || dispatchFilterEndDate) {
        const filterStart = dispatchFilterStartDate ? new Date(dispatchFilterStartDate + "T00:00:00.000Z") : null;
        const filterEnd = dispatchFilterEndDate ? new Date(dispatchFilterEndDate + "T23:59:59.999Z") : null;

        filteredDispatches = dispatches.filter(dispatch => {
            const dispatchDate = new Date(dispatch.dispatchedAt);
            let satisfiesStart = true;
            if (filterStart) {
                satisfiesStart = dispatchDate >= filterStart;
            }
            let satisfiesEnd = true;
            if (filterEnd) {
                satisfiesEnd = dispatchDate <= filterEnd;
            }
            return satisfiesStart && satisfiesEnd;
        });
    }

    filteredDispatches.sort((a, b) => new Date(b.dispatchedAt).getTime() - new Date(a.dispatchedAt).getTime());

    const displayLines: DisplayDispatchLine[] = [];
    let overallTotalQuantity = 0;
    const itemDispatchSummary: { [materialName: string]: number } = {};
    const salesmanDispatchSummary: { [salesmanName: string]: number } = {};
    const partyDispatchSummary: { [partyName: string]: number } = {};


    filteredDispatches.forEach(dispatch => {
        const po = purchaseOrders.find(p => p.id === dispatch.poId);

        // If PO is cancelled, skip its dispatches from the log and summaries
        if (po && po.status === 'Cancelled') {
            return; // Skip to the next dispatch
        }

        let partyName = 'N/A (PO Data Missing)';
        let destination = 'N/A';
        let salesmanName = 'N/A';

        if (po) {
            partyName = po.partyName;
            destination = po.destination;
            salesmanName = po.salesmanName || 'N/A';
        } else {
             console.warn(`PO with ID ${dispatch.poId} not found for dispatch ${dispatch.id}`);
        }


        dispatch.dispatchedItems.forEach(dispItem => {
            const materialKey = dispItem.material.toUpperCase();
            overallTotalQuantity += dispItem.quantity;
            itemDispatchSummary[materialKey] = (itemDispatchSummary[materialKey] || 0) + dispItem.quantity;

            if (salesmanName !== 'N/A' && salesmanName !== 'N/A (PO Data Missing)') {
                 salesmanDispatchSummary[salesmanName] = (salesmanDispatchSummary[salesmanName] || 0) + dispItem.quantity;
            }
            if (partyName !== 'N/A' && partyName !== 'N/A (PO Data Missing)') {
                 partyDispatchSummary[partyName] = (partyDispatchSummary[partyName] || 0) + dispItem.quantity;
            }


            let itemTotalAmount = 0;
            if (po) {
                const poItem = po.items.find(i => i.material.toUpperCase() === materialKey);
                if (poItem) {
                    const taxableAmount = dispItem.quantity * poItem.rate;
                    const gstAmount = (taxableAmount * poItem.gstPercentage) / 100;
                    itemTotalAmount = taxableAmount + gstAmount;
                }
            }

            displayLines.push({
                dispatchId: dispatch.id,
                poId: dispatch.poId,
                partyName: partyName,
                salesmanName: salesmanName,
                dispatchedAt: formatToDDMMYY(dispatch.dispatchedAt),
                originalDispatchedAtISO: dispatch.dispatchedAt,
                invoiceNumber: dispatch.invoiceNumber,
                vehicleNumber: dispatch.vehicleNumber,
                driverContact: dispatch.driverContact,
                materialName: dispItem.material,
                quantityDispatched: dispItem.quantity,
                transporterName: dispatch.transporterName,
                itemTotalAmount: itemTotalAmount,
                destination: destination,
            });
        });
    });

    let content = `<div class="list-container"><h2>Dispatch Log (${displayLines.length} item lines)</h2>`;

    content += `
        <div class="dispatch-filter-controls" style="display: flex; flex-wrap: wrap; align-items: flex-end; gap: 15px; margin-bottom: 20px; padding: 15px; background-color: #f8f9fa; border-radius: 5px; box-shadow: 0 2px 4px rgba(0,0,0,0.05);">
            <div class="form-group" style="margin-bottom: 0;">
                <label for="filterStartDate">From Date:</label>
                <input type="date" id="filterStartDate" name="filterStartDate" value="${escapeHTML(dispatchFilterStartDate)}" aria-label="Filter start date">
            </div>
            <div class="form-group" style="margin-bottom: 0;">
                <label for="filterEndDate">To Date:</label>
                <input type="date" id="filterEndDate" name="filterEndDate" value="${escapeHTML(dispatchFilterEndDate)}" aria-label="Filter end date">
            </div>
            <button id="filterDispatchButton" class="primary" type="button">Filter</button>
            <button id="clearDispatchFilterButton" class="secondary" type="button">Clear Filters</button>
        </div>
    `;

    if (displayLines.length === 0) {
        content += '<p>No dispatches recorded for the selected criteria, or all relevant dispatches belong to cancelled POs.</p>';
    } else {
        content += `
            <div class="table-responsive-wrapper">
            <table aria-label="Dispatch Log">
                <thead>
                    <tr>
                        <th>Dispatch ID</th>
                        <th>PO ID</th>
                        <th>Date</th>
                        <th>Item Name</th>
                        <th class="text-right">Quantity</th>
                        <th>Vehicle No.</th>
                        <th>Actions</th>
                    </tr>
                </thead>
                <tbody>
                    ${displayLines.map(line => `
                        <tr>
                            <td>${escapeHTML(line.dispatchId)}</td>
                            <td><a href="#" onclick="window.showPODetailsModalWrapper('${escapeHTML(line.poId)}'); return false;" aria-label="View details for PO ${escapeHTML(line.poId)}">${escapeHTML(line.poId)}</a></td>
                            <td>${escapeHTML(line.dispatchedAt)}</td>
                            <td>${escapeHTML(line.materialName)}</td>
                            <td class="text-right">${line.quantityDispatched.toFixed(2)}</td>
                            <td>${escapeHTML(line.vehicleNumber)}</td>
                            <td>
                                ${purchaseOrders.find(p => p.id === line.poId)?.status !== 'Cancelled' ? `<button class="info small" onclick="window.showEditDispatchModal('${escapeHTML(line.dispatchId)}')" aria-label="Edit dispatch ${escapeHTML(line.dispatchId)}">Edit</button>`: 'N/A'}
                            </td>
                        </tr>
                    `).join('')}
                </tbody>
                <tfoot>
                    <tr>
                        <td colspan="4" style="text-align: right; font-weight: bold;">Total Dispatched Quantity (Units):</td>
                        <td style="font-weight: bold;" class="text-right">${overallTotalQuantity.toFixed(2)}</td>
                        <td colspan="2"></td>
                    </tr>
                </tfoot>
            </table>
            </div>
             <p style="font-size: 0.9em; margin-top: 10px;">For full dispatch details (Party, Salesman, Invoice etc.), click PO ID then view dispatches, or use Excel download.</p>
            <div style="text-align: right; margin-top: 20px;">
                <button id="download-dispatch-csv" class="secondary" aria-label="Download Dispatch Log as CSV">Download Full Log as Excel (CSV)</button>
            </div>
        `;

        content += `<div class="dispatch-summary-container">`;
        content += `
            <div class="summary-table-wrapper">
                <h3>Item-wise Dispatch Summary</h3>
        `;
        if (Object.keys(itemDispatchSummary).length > 0) {
            content += `
                <div class="table-responsive-wrapper">
                <table aria-label="Item-wise Dispatch Summary">
                    <thead><tr><th>Material Name</th><th class="text-right">Total Quantity Dispatched</th></tr></thead>
                    <tbody>
                        ${Object.entries(itemDispatchSummary).sort((a,b) => b[1] - a[1]).map(([material, quantity]) => `
                            <tr><td>${escapeHTML(material)}</td><td class="text-right">${quantity.toFixed(2)}</td></tr>`).join('')}
                    </tbody>
                </table>
                </div>
                <div style="text-align: right; margin-top: 10px;">
                    <button id="download-item-summary-csv" class="secondary small" aria-label="Download Item Summary as CSV">Download Item Summary (CSV)</button>
                </div>`;
        } else {
            content += `<p>No item dispatch data to summarize for the selected criteria.</p>`;
        }
        content += `</div>`;

        content += `
            <div class="summary-table-wrapper">
                <h3>Salesman-wise Dispatch Summary</h3>
        `;
        if (Object.keys(salesmanDispatchSummary).length > 0) {
            content += `
                <div class="table-responsive-wrapper">
                <table aria-label="Salesman-wise Dispatch Summary">
                    <thead><tr><th>Salesman Name</th><th class="text-right">Total Quantity Dispatched</th></tr></thead>
                    <tbody>
                        ${Object.entries(salesmanDispatchSummary).sort((a,b) => b[1] - a[1]).map(([salesman, quantity]) => `
                            <tr><td>${escapeHTML(salesman)}</td><td class="text-right">${quantity.toFixed(2)}</td></tr>`).join('')}
                    </tbody>
                </table>
                </div>
                <div style="text-align: right; margin-top: 10px;">
                    <button id="download-salesman-summary-csv" class="secondary small" aria-label="Download Salesman Summary as CSV">Download Salesman Summary (CSV)</button>
                </div>`;
        } else {
            content += `<p>No salesman dispatch data to summarize for the selected criteria.</p>`;
        }
        content += `</div>`;

        content += `
            <div class="summary-table-wrapper">
                <h3>Party-wise Dispatch Summary</h3>
        `;
        if (Object.keys(partyDispatchSummary).length > 0) {
            content += `
                <div class="table-responsive-wrapper">
                <table aria-label="Party-wise Dispatch Summary">
                    <thead><tr><th>Party Name</th><th class="text-right">Total Quantity Dispatched</th></tr></thead>
                    <tbody>
                        ${Object.entries(partyDispatchSummary).sort((a,b) => b[1] - a[1]).map(([party, quantity]) => `
                            <tr><td>${escapeHTML(party)}</td><td class="text-right">${quantity.toFixed(2)}</td></tr>`).join('')}
                    </tbody>
                </table>
                </div>
                <div style="text-align: right; margin-top: 10px;">
                    <button id="download-party-summary-csv" class="secondary small" aria-label="Download Party Summary as CSV">Download Party Summary (CSV)</button>
                </div>`;
        } else {
            content += `<p>No party dispatch data to summarize for the selected criteria.</p>`;
        }
        content += `</div>`;

        content += `</div>`;
    }
    content += `</div>`;
    content += getPODetailsModalHTML() + getEditDispatchModalHTML();
    mainContent.innerHTML = content;

    document.getElementById('filterDispatchButton')?.addEventListener('click', () => {
        dispatchFilterStartDate = (document.getElementById('filterStartDate') as HTMLInputElement).value;
        dispatchFilterEndDate = (document.getElementById('filterEndDate') as HTMLInputElement).value;
        saveData();
        renderDispatchList();
    });

    document.getElementById('clearDispatchFilterButton')?.addEventListener('click', () => {
        dispatchFilterStartDate = '';
        dispatchFilterEndDate = '';
        (document.getElementById('filterStartDate') as HTMLInputElement).value = '';
        (document.getElementById('filterEndDate') as HTMLInputElement).value = '';
        saveData();
        renderDispatchList();
    });

    if (displayLines.length > 0) {
        document.getElementById('download-dispatch-csv')?.addEventListener('click', () => {
            downloadDispatchLogAsCSV(displayLines);
        });
    }
    if (Object.keys(itemDispatchSummary).length > 0) {
        document.getElementById('download-item-summary-csv')?.addEventListener('click', () => {
            downloadItemSummaryCSV(itemDispatchSummary);
        });
    }
    if (Object.keys(salesmanDispatchSummary).length > 0) {
        document.getElementById('download-salesman-summary-csv')?.addEventListener('click', () => {
            downloadSalesmanSummaryCSV(salesmanDispatchSummary);
        });
    }
    if (Object.keys(partyDispatchSummary).length > 0) {
        document.getElementById('download-party-summary-csv')?.addEventListener('click', () => {
            downloadPartySummaryCSV(partyDispatchSummary);
        });
    }
}

// Wrapper for showPODetailsModal called from dispatch list to ensure it's available
(window as any).showPODetailsModalWrapper = (poId: string): void => {
    const po = purchaseOrders.find(p => p.id === poId);
    if (po) {
        (window as any).showPODetailsModal(poId);
    } else {
        alert(`Purchase Order ${poId} not found.`);
    }
};


// Close modal if user clicks outside of it or presses Escape
window.addEventListener('click', (event: MouseEvent) => {
    document.querySelectorAll('.modal').forEach(modal => {
        if (event.target === modal) {
            const modalId = modal.id;
            if (modalId) {
                 closeModal(modalId);
            } else {
                (modal as HTMLElement).style.display = "none";
                modal.setAttribute('aria-hidden', 'true');
            }
        }
    });
});
window.addEventListener('keydown', (event: KeyboardEvent) => {
    if (event.key === 'Escape') {
        document.querySelectorAll('.modal[style*="display: block"]').forEach(modal => {
             const modalId = modal.id;
             if (modalId) {
                 closeModal(modalId);
             } else {
                  (modal as HTMLElement).style.display = "none";
                  modal.setAttribute('aria-hidden', 'true');
             }
        });
    }
});


// --- Initial Load ---
document.addEventListener('DOMContentLoaded', () => {
    loadData();
    dispatches.sort((a, b) => new Date(b.dispatchedAt).getTime() - new Date(a.dispatchedAt).getTime());
    purchaseOrders.forEach(po => updatePOStatus(po.id)); // Ensure statuses are correct on load
    saveData(); // Save any status updates
    // Set initial view. If create-po, ensure state is clean.
    if (currentView === 'create-po') {
        _formModeForCreatePage = 'create';
        _formDataForCreatePage = undefined;
        _originalIdForRevisionOnCreatePage = undefined;
        currentEditingPOId = null;
    }
    renderApp();
});