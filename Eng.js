/**
 * Plant Transfer Calendar Web Application
 * NOTE: Replace APPS_SCRIPT_URL with your actual published Google Apps Script Web App URL
 */
const APPS_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbyqm07HvfEj7RPmKG2d2qs5gNTQH23M1krMLiyaa9osX2vAIGa6hiIvC-dXE4ApGSfKBA/exec"; 

// State Management
let currentDate = new Date(); // Represents the currently viewed month
let systemDate = new Date(); // Actual today
let masterPlants = []; // Data from 'PlantTranfer' sheet
let bookings = []; // Data from 'Bookings' sheet
let selectedDateStr = "";
let currentUser = null; // Store logged-in user

// DOM Elements
const loginOverlay = document.getElementById("login-overlay");
const appContainer = document.getElementById("app-container");
const loginUserInput = document.getElementById("login-user");
const loginPassInput = document.getElementById("login-pass");
const btnLogin = document.getElementById("btn-login");
const btnLogout = document.getElementById("btn-logout");
const loginError = document.getElementById("login-error");
const btnUserProfile = document.getElementById("btn-user-profile");
const userProfileModal = document.getElementById("user-profile-modal");
const closeProfileModal = document.getElementById("close-profile-modal");
const userProfileDetails = document.getElementById("user-profile-details");
const calendarDaysEl = document.getElementById("calendar-days");
const monthYearEl = document.getElementById("current-month-year");
const prevBtn = document.getElementById("prev-month");
const nextBtn = document.getElementById("next-month");

const modal = document.getElementById("booking-modal");
const closeBtn = document.getElementById("close-modal");
const modalDateTitle = document.getElementById("modal-date-title");
const noPlantsMsg = document.getElementById("no-plants-msg");
const bookingsForm = document.getElementById("bookings-form");
const plantsFormsContainer = document.getElementById("plants-forms-container");
const connectionStatus = document.getElementById("connection-status");
const statusDot = connectionStatus.querySelector(".status-dot");
const statusText = connectionStatus.querySelector(".status-text");

// Handle Login
btnLogin.addEventListener("click", async () => {
    const username = loginUserInput.value.trim();
    const password = loginPassInput.value.trim();
    
    if (!username || !password) {
        loginError.textContent = "กรุณากรอกข้อมูลให้ครบถ้วน";
        loginError.classList.remove("hidden");
        return;
    }
    
    loginError.classList.add("hidden");
    const originalBtnText = btnLogin.textContent;
    btnLogin.textContent = "Authenticating...";
    btnLogin.disabled = true;
    
    try {
        const SHEET_ID = "1tYlXDTrIoMYajLDJZa_7apaEpvVzB0rRld7K6Ymbhck";
        const response = await fetch(`https://opensheet.elk.sh/${SHEET_ID}/Username`);
        
        if (!response.ok) throw new Error("Could not reach authentication server.");
        
        const users = await response.json();
        const user = users.find(u => String(u.IDRec) === username);
        
        if (user) {
            const expectedPassword = String(user.IDRec).slice(-4);
            if (password === expectedPassword) {
                // Success!
                currentUser = user;
                
                // Show App
                loginOverlay.classList.remove("active");
                appContainer.style.display = "flex";
                
                // Show User Profile Button
                btnUserProfile.style.display = "flex";
                
                // Prepare Profile Details
                userProfileDetails.innerHTML = `
                    <div style="display: grid; grid-template-columns: 120px 1fr; gap: 0.5rem; margin-bottom: 0.5rem; border-bottom: 1px solid rgba(255,255,255,0.1); padding-bottom: 0.5rem;">
                        <span style="font-weight: 500;">IDRec:</span>
                        <span>${user.IDRec || "-"}</span>
                        
                        <span style="font-weight: 500;">Name:</span>
                        <span>${user.Name || "-"}</span>
                        
                        <span style="font-weight: 500;">หน่วยงาน:</span>
                        <span>${user["หน่วยงาน"] || user.Unit || "-"}</span>
                        
                        <span style="font-weight: 500;">Plant:</span>
                        <span>${user.Plant || "-"}</span>
                        
                        <span style="font-weight: 500;">Status:</span>
                        <span style="color: ${String(user.Status).toLowerCase() === 'admin' ? '#fde047' : '#93c5fd'}; font-weight: 600;">${user.Status || "-"}</span>
                    </div>
                `;
                
                // Init data fetch
                init();
                logUserAction("Login", "User logged into the system.");
                return;
            }
        }
        
        // Failed
        loginError.textContent = "รหัสพนักงาน หรือ รหัสผ่าน (4 ตัวท้ายของรหัส) ไม่ถูกต้อง";
        loginError.classList.remove("hidden");
    } catch (e) {
        console.error(e);
        loginError.textContent = "ระบบมีปัญหาขัดข้อง ไม่สามารถล็อกอินได้";
        loginError.classList.remove("hidden");
    } finally {
        btnLogin.textContent = originalBtnText;
        btnLogin.disabled = false;
    }
});

btnUserProfile.addEventListener("click", () => {
    userProfileModal.classList.add("active");
});

closeProfileModal.addEventListener("click", () => {
    userProfileModal.classList.remove("active");
});
window.addEventListener("click", (e) => {
    if (e.target === userProfileModal) userProfileModal.classList.remove("active");
});

btnLogout.addEventListener("click", () => {
    logUserAction("Logout", "User logged out.");
    currentUser = null;
    appContainer.style.display = "none";
    userProfileModal.classList.remove("active");
    btnUserProfile.style.display = "none";
    loginOverlay.classList.add("active");
    loginUserInput.value = "";
    loginPassInput.value = "";
    loginError.classList.add("hidden");
});

// Initialize
async function init() {
    updateConnectionStatus("Fetching data...", "warning");
    renderCalendar(); // Render optimistic empty calendar
    
    if (APPS_SCRIPT_URL !== "YOUR_APPS_SCRIPT_WEB_APP_URL_HERE") {
        await fetchData();
    } else {
        updateConnectionStatus("Please set APPS_SCRIPT_URL in your script", "error");
        // Inject dummy data for preview purposes if URL is not set
        setTimeout(() => {
            loadDummyData();
            renderCalendar();
            updateConnectionStatus("Preview Mode (Local Logic only)", "success");
        }, 1000);
    }
}

function updateConnectionStatus(msg, type) {
    statusText.textContent = msg;
    statusDot.className = `status-dot ${type}`;
}

// Data Fetching
async function fetchData() {
    try {
        const SHEET_ID = "1tYlXDTrIoMYajLDJZa_7apaEpvVzB0rRld7K6Ymbhck";
        
        // Fetch Master Plants from OpenSheet
        const plantsRes = await fetch(`https://opensheet.elk.sh/${SHEET_ID}/PlantTranfer`);
        if (!plantsRes.ok) throw new Error("Failed to load Plants");
        const plantsData = await plantsRes.json();
        
        // Fetch Bookings from OpenSheet
        const bookingsRes = await fetch(`https://opensheet.elk.sh/${SHEET_ID}/Bookings`);
        let bookingsData = [];
        if (bookingsRes.ok) {
            bookingsData = await bookingsRes.json();
            if (bookingsData.error) bookingsData = [];
        }
        
        // Map OpenSheet data format
        let tempPlants = plantsData.map(p => ({
            idPlant: String(p.IDPlant || p.idplant || ""),
            plant: String(p.Plant || p.plant || ""),
            round: String(p.Round || p.round || "Non"),
            dateStart: String(p.DateStart || p.datestart || ""),
            dateBegin: String(p.Datebegin || p.datebegin || "")
        }));
        
        // Auth Filter Logic 
        if (currentUser && String(currentUser.Status).toLowerCase() !== "admin") {
            const userPlantId = String(currentUser.Plant); // Note: "Plant" in Username sheet = IDPlant in PlantTranfer
            tempPlants = tempPlants.filter(p => p.idPlant === userPlantId);
        }
        masterPlants = tempPlants;
        
        bookings = bookingsData.map(b => ({
            id: String(b.ID || b.id || ""),
            date: String(b.Date || b.date || ""),
            plant: String(b.Plant || b.plant || ""),
            deliveryNumber: String(b.DeliveryNumber || b.deliverynumber || ""),
            scrapNumber: String(b.ScrapNumber || b.scrapnumber || ""),
            timestamp: String(b.Timestamp || b.timestamp || "")
        }));
        
        updateConnectionStatus("Data Loaded (OpenSheet API)", "success");
        renderCalendar(); // Re-render with real data
    } catch (error) {
        console.error("Error fetching data:", error);
        updateConnectionStatus("Failed to Load Data", "error");
    }
}

// Dummy Data for immediate local preview
function loadDummyData() {
    masterPlants = [
        { idPlant: "0101", plant: "Plant A", round: "1", dateStart: "2,3,4,5", dateBegin: "2026-03-10" },   // Mon-Thu, Every week
        { idPlant: "0201", plant: "Plant B", round: "2", dateStart: "6", dateBegin: "" },         // Fri, Every other week
        { idPlant: "0301", plant: "Stock กทม", round: "1", dateStart: "1,7", dateBegin: "" },       // Sun/Sat, Every week
        { idPlant: "0401", plant: "Plant Non", round: "Non", dateStart: "2", dateBegin: "" }      // Ignored
    ];
}

// Calendar Rendering Logic
function renderCalendar() {
    calendarDaysEl.innerHTML = "";
    
    const year = currentDate.getFullYear();
    const month = currentDate.getMonth(); // 0-11
    
    // Set Header
    const monthNames = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
    monthYearEl.textContent = `${monthNames[month]} ${year}`;
    
    // First day of month (0 = Sun, 1 = Mon, etc.)
    const firstDay = new Date(year, month, 1).getDay();
    // Number of days in month
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    
    // Empty prefix cells
    for (let i = 0; i < firstDay; i++) {
        const emptyCell = document.createElement("div");
        emptyCell.classList.add("day-cell", "empty");
        calendarDaysEl.appendChild(emptyCell);
    }
    
    // Day cells
    for (let i = 1; i <= daysInMonth; i++) {
        const cellDate = new Date(year, month, i);
        const dateStr = formatDate(cellDate);
        const dayOfWeek = cellDate.getDay() + 1; // 1=Sun, ..., 7=Sat
        
        const cell = document.createElement("div");
        cell.classList.add("day-cell");
        if (dayOfWeek === 1) cell.classList.add("sunday-cell");
        
        // Highlight Today
        if (
            i === systemDate.getDate() && 
            month === systemDate.getMonth() && 
            year === systemDate.getFullYear()
        ) {
            cell.classList.add("today");
        }
        
        // Render Numbers
        const numberEl = document.createElement("div");
        numberEl.classList.add("day-number");
        numberEl.textContent = i;
        cell.appendChild(numberEl);
        
        // Render badges based on active Plants and existing bookings
        const contentEl = document.createElement("div");
        contentEl.classList.add("day-content");
        
        const activePlantsForDay = getActivePlantsForDate(cellDate, dayOfWeek);
        
        activePlantsForDay.forEach(plantObj => {
            const hasBooking = bookings.find(b => b.date === dateStr && b.plant === plantObj.plant);
            const btn = document.createElement("button");
            btn.classList.add("plant-btn");
            
            const displayName = formatDisplayName(plantObj);
            
            if (hasBooking) {
                btn.classList.add("success");
                btn.textContent = `✓ ${displayName}`;
            } else {
                btn.textContent = `+ ${displayName}`;
            }
            
            btn.addEventListener("click", (e) => {
                e.stopPropagation();
                openModal(cellDate, [plantObj], dateStr);
            });
            
            contentEl.appendChild(btn);
        });
        
        cell.appendChild(contentEl);
        
        // Click handler
        cell.addEventListener("click", () => openModal(cellDate, activePlantsForDay, dateStr));
        
        calendarDaysEl.appendChild(cell);
    }
}

// Helper to parse dates like "25/03/2026"
function parseDateStr(ds) {
    if (!ds) return null;
    let clean = ds.split(" ")[0].trim();
    if (clean.includes("/")) {
        const parts = clean.split("/");
        if (parts.length >= 3) {
            return new Date(parts[2], parseInt(parts[1])-1, parseInt(parts[0]));
        }
    }
    return new Date(ds);
}

// Format plant name for display (e.g. remove "Stock ")
function formatDisplayName(plantObj) {
    if (!plantObj || !plantObj.plant) return "";
    return plantObj.plant.replace(/Stock\s*/i, "").trim();
}

// Logic to determine if a Plant is active on a given date
function getActivePlantsForDate(dateObj, dayOfWeek) { // dayOfWeek: 1-7
    const result = [];
    
    const weekNumber = getWeekNumber(dateObj);
    const isEvenWeek = weekNumber % 2 === 0;
    
    // Check date (ignore time)
    const checkingTime = new Date(dateObj.getFullYear(), dateObj.getMonth(), dateObj.getDate()).getTime();
    
    masterPlants.forEach(p => {
        if (!p.round || String(p.round).toLowerCase() === "non") return;
        
        let beginWeek = -1;
        
        // Datebegin Check
        if (p.dateBegin) {
            const bDate = parseDateStr(p.dateBegin);
            if (bDate && !isNaN(bDate.getTime())) {
                const beginTime = new Date(bDate.getFullYear(), bDate.getMonth(), bDate.getDate()).getTime();
                beginWeek = getWeekNumber(bDate);
                
                if (checkingTime < beginTime) {
                    return; // Skip if calendar date is before Datebegin
                }
            }
        }
        
        // Check days. "DateStart" could be "5,7"
        if (!p.dateStart) return;
        const validDays = String(p.dateStart).split(",").map(d => parseInt(d.trim()));
        
        if (validDays.includes(dayOfWeek)) {
            if (String(p.round) === "1") {
                // Every week
                result.push(p);
            } else if (String(p.round) === "2") {
                // Every other week relative to Datebegin
                if (beginWeek !== -1) {
                    const weekDiff = Math.abs(weekNumber - beginWeek);
                    if (weekDiff % 2 === 0) {
                        result.push(p);
                    }
                } else {
                    if (isEvenWeek) result.push(p);
                }
            }
        }
    });
    
    return result;
}

// Format Date as YYYY-MM-DD
function formatDate(d) {
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

// Get ISO Week Number
function getWeekNumber(d) {
    const date = new Date(d.getTime());
    date.setHours(0, 0, 0, 0);
    date.setDate(date.getDate() + 3 - (date.getDay() + 6) % 7);
    const week1 = new Date(date.getFullYear(), 0, 4);
    return 1 + Math.round(((date.getTime() - week1.getTime()) / 86400000 - 3 + (week1.getDay() + 6) % 7) / 7);
}

// Navigation Listeners
prevBtn.addEventListener("click", () => {
    currentDate.setMonth(currentDate.getMonth() - 1);
    renderCalendar();
});

nextBtn.addEventListener("click", () => {
    currentDate.setMonth(currentDate.getMonth() + 1);
    renderCalendar();
});

// Modal Logic
function openModal(dateObj, activePlants, dateStr) {
    selectedDateStr = dateStr;
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric' };
    modalDateTitle.textContent = dateObj.toLocaleDateString('en-US', options); // Or 'th-TH' for Thai
    
    plantsFormsContainer.innerHTML = "";
    
    if (activePlants.length === 0) {
        noPlantsMsg.classList.remove("hidden");
        bookingsForm.classList.add("hidden");
    } else {
        noPlantsMsg.classList.add("hidden");
        bookingsForm.classList.remove("hidden");
        
        activePlants.forEach(plantObj => renderPlantForm(plantObj, dateStr));
    }
    
    modal.classList.add("active");
}

function renderPlantForm(plantObj, dateStr) {
    const plantName = plantObj.plant;
    const displayName = formatDisplayName(plantObj);
    
    // Find existing booking
    const existing = bookings.find(b => b.date === dateStr && b.plant === plantName);
    const formId = existing ? existing.id : "";
    const deliveryVal = existing ? existing.deliveryNumber : "";
    const scrapVal = existing ? existing.scrapNumber : "";
    
    const card = document.createElement("div");
    card.classList.add("plant-form-card");
    
    card.innerHTML = `
        <div class="plant-form-header">
            <h4><span class="plant-badge">${displayName}</span></h4>
        </div>
        
        <div class="form-group">
            <label>Delivery Reservation (เลขที่ใบจองรถ ส่งของ)</label>
            <input type="text" id="deliv-${plantName}" value="${deliveryVal}" placeholder="Enter delivery reservation NO..." />
        </div>
        
        <div class="form-group">
            <label>Scrap Reservation (เลขที่รับซาก)</label>
            <input type="text" id="scrap-${plantName}" value="${scrapVal}" placeholder="Enter scrap reservation NO..." />
        </div>
        
        <div class="plant-actions">
            ${existing ? `<button type="button" class="btn btn-danger" onclick="handleDelete('${formId}')">Delete</button>` : ''}
            <button type="button" class="btn btn-primary" onclick="handleSave('${plantName}', '${formId}')">Save</button>
        </div>
    `;
    
    plantsFormsContainer.appendChild(card);
}

closeBtn.addEventListener("click", () => {
    modal.classList.remove("active");
});
window.addEventListener("click", (e) => {
    if (e.target === modal) modal.classList.remove("active");
});

// API Interactions (Save, Delete)
async function handleSave(plantName, existingId) {
    const dVal = document.getElementById(`deliv-${plantName}`).value.trim();
    const sVal = document.getElementById(`scrap-${plantName}`).value.trim();
    
    // Validation: Starts with 205 and 10 digits total
    const regex = /^205\d{7}$/;
    
    if (dVal && !regex.test(dVal)) {
        await showCustomAlert("ข้อมูลไม่ถูกต้อง", "ข้อมูลใบจองรถ ต้องมีตัวเลข 10 หลัก และขึ้นต้นด้วย '205' เท่านั้นครับ", "error");
        return;
    }
    if (sVal && !regex.test(sVal)) {
        await showCustomAlert("ข้อมูลไม่ถูกต้อง", "ข้อมูลใบรับซาก ต้องมีตัวเลข 10 หลัก และขึ้นต้นด้วย '205' เท่านั้นครับ", "error");
        return;
    }
    if (!dVal && !sVal) {
        await showCustomAlert("ข้อมูลไม่ครบถ้วน", "กรุณากรอกข้อมูลอย่างน้อย 1 ช่องเพื่อบันทึกครับ", "warning");
        return;
    }
    
    const payload = {
        id: existingId,
        date: selectedDateStr,
        plant: plantName,
        deliveryNumber: dVal,
        scrapNumber: sVal
    };
    
    if (APPS_SCRIPT_URL === "YOUR_APPS_SCRIPT_WEB_APP_URL_HERE") {
        updateConnectionStatus("Mock Save Success!", "success");
        // Update local state mocking
        let idx = bookings.findIndex(b => b.id === existingId && existingId !== "");
        if (idx !== -1) {
            bookings[idx].deliveryNumber = dVal;
            bookings[idx].scrapNumber = sVal;
        } else {
            bookings.push({...payload, id: "mock_" + Date.now()});
        }
        renderCalendar();
        openModal(new Date(selectedDateStr), getActivePlantsForDate(new Date(selectedDateStr), new Date(selectedDateStr).getDay() + 1), selectedDateStr);
        return;
    }
    
    updateConnectionStatus("Saving...", "warning");
    
    try {
        const payloadStr = encodeURIComponent(JSON.stringify(payload));
        const response = await fetch(`${APPS_SCRIPT_URL}?action=saveBooking&payload=${payloadStr}`, {
            method: 'GET'
        });
        const resData = await response.json();
        
        if (resData.status === "success") {
            updateConnectionStatus("Saved successfully", "success");
            
            // Optimistic update to avoid OpenSheet Cache delay
            const payloadId = resData.id || existingId;
            const newBooking = {
                id: String(payloadId),
                date: selectedDateStr,
                plant: plantName,
                deliveryNumber: dVal,
                scrapNumber: sVal,
                timestamp: new Date().toISOString()
            };
            const idx = bookings.findIndex(b => b.id === payloadId);
            if (idx !== -1) {
                bookings[idx] = newBooking;
            } else {
                bookings.push(newBooking);
            }
            renderCalendar();
            
            logUserAction(existingId ? "Edit Booking" : "Create Booking", `Plant: ${plantName}, Delivery: ${dVal}, Scrap: ${sVal}`);
            
            await showCustomAlert("สำเร็จ", "บันทึกข้อมูลเรียบร้อยแล้ว!", "success");
            modal.classList.remove("active");
        } else {
            await showCustomAlert("เกิดข้อผิดพลาด", resData.message, "error");
            updateConnectionStatus("Error saving", "error");
        }
    } catch (e) {
        console.error(e);
        updateConnectionStatus("Request failed", "error");
    }
}

// Ensure function is attached to window so inline onclick works
window.handleSave = handleSave;

async function handleDelete(id) {
    const isConfirm = await showCustomConfirm("ยืนยันการลบ", "คุณต้องการลบข้อมูลการจองนี้ใช่หรือไม่?");
    if (!isConfirm) return;
    
    if (APPS_SCRIPT_URL === "YOUR_APPS_SCRIPT_WEB_APP_URL_HERE") {
        updateConnectionStatus("Mock Delete Success!", "success");
        bookings = bookings.filter(b => b.id !== id);
        renderCalendar();
        modal.classList.remove("active");
        return;
    }
    
    updateConnectionStatus("Deleting...", "warning");
    
    try {
        const payloadStr = encodeURIComponent(JSON.stringify({ id: id }));
        const response = await fetch(`${APPS_SCRIPT_URL}?action=deleteBooking&payload=${payloadStr}`, {
            method: 'GET'
        });
        const resData = await response.json();
        
        if (resData.status === "success") {
            updateConnectionStatus("Deleted successfully", "success");
            
            // Optimistic deletion
            bookings = bookings.filter(b => b.id !== id);
            renderCalendar();
            
            logUserAction("Delete Booking", `Deleted booking ID: ${id}`);
            
            await showCustomAlert("สำเร็จ", "ลบข้อมูลเรียบร้อยแล้ว!", "success");
            modal.classList.remove("active");
        } else {
            await showCustomAlert("เกิดข้อผิดพลาด", resData.message, "error");
            updateConnectionStatus("Error deleting", "error");
        }
    } catch (e) {
        console.error(e);
        updateConnectionStatus("Request failed", "error");
    }
}

// Ensure function is attached to window so inline onclick works
window.handleDelete = handleDelete;

// Custom Alert/Confirm Modals
const customAlertModal = document.getElementById("custom-alert-modal");
const alertTitle = document.getElementById("alert-title");
const alertMessage = document.getElementById("alert-message");
const alertIcon = document.getElementById("alert-icon");
const alertOkBtn = document.getElementById("alert-ok-btn");
const alertCancelBtn = document.getElementById("alert-cancel-btn");

function showCustomAlert(title, message, type = "success") {
    return new Promise((resolve) => {
        alertTitle.textContent = title;
        alertMessage.textContent = message;
        
        let svg = '';
        alertIcon.className = `alert-icon ${type}`;
        if (type === 'success') {
            svg = `<svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"></path><polyline points="22 4 12 14.01 9 11.01"></polyline></svg>`;
        } else if (type === 'warning') {
            svg = `<svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>`;
        } else if (type === 'error') {
            svg = `<svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="15" y1="9" x2="9" y2="15"></line><line x1="9" y1="9" x2="15" y2="15"></line></svg>`;
        }
        alertIcon.innerHTML = svg;
        
        alertCancelBtn.classList.add("hidden");
        customAlertModal.classList.add("active");
        
        const handleOk = () => {
            cleanup();
            resolve(true);
        };
        
        const cleanup = () => {
            alertOkBtn.removeEventListener("click", handleOk);
            customAlertModal.classList.remove("active");
        };
        
        alertOkBtn.addEventListener("click", handleOk);
    });
}

function showCustomConfirm(title, message) {
    return new Promise((resolve) => {
        alertTitle.textContent = title;
        alertMessage.textContent = message;
        
        alertIcon.className = `alert-icon warning`;
        alertIcon.innerHTML = `<svg width="32" height="32" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"></circle><line x1="12" y1="8" x2="12" y2="12"></line><line x1="12" y1="16" x2="12.01" y2="16"></line></svg>`;
        
        alertCancelBtn.classList.remove("hidden");
        customAlertModal.classList.add("active");
        
        const handleOk = () => { cleanup(); resolve(true); };
        const handleCancel = () => { cleanup(); resolve(false); };
        
        const cleanup = () => {
            alertOkBtn.removeEventListener("click", handleOk);
            alertCancelBtn.removeEventListener("click", handleCancel);
            customAlertModal.classList.remove("active");
        };
        
        alertOkBtn.addEventListener("click", handleOk);
        alertCancelBtn.addEventListener("click", handleCancel);
    });
}

function logUserAction(actionType, details) {
    if (!currentUser) return;
    try {
        const payload = {
            userId: String(currentUser.IDRec),
            name: String(currentUser.Name || currentUser.IDRec),
            action: actionType,
            details: String(details || "")
        };
        const payloadStr = encodeURIComponent(JSON.stringify(payload));
        fetch(`${APPS_SCRIPT_URL}?action=logAction&payload=${payloadStr}`, { method: 'GET' })
            .catch(e => console.error("Logging error", e));
    } catch(e) {}
}

// Application is initialized after successful login via btnLogin.addEventListener
