// =========================================================
// 1. CONFIGURATION
// =========================================================

// 🔴 BACKEND: Your Render URL
const API_URL = "https://booking-a-room-poc.onrender.com"; 

// 🔴 FRONTEND: Your Azure AD "Frontend" App Registration
const msalConfig = {
    auth: {
        clientId: "PASTE_YOUR_FRONTEND_CLIENT_ID_HERE", // e.g. "a1b2c3d4-..."
        authority: "https://login.microsoftonline.com/PASTE_YOUR_TENANT_ID_HERE",
        redirectUri: window.location.origin, // Auto-detects localhost or vercel
    },
    cache: {
        cacheLocation: "sessionStorage", // Securely store tokens in the tab
        storeAuthStateInCookie: false,
    }
};

const HOURS_TO_SHOW = 12;

// =========================================================
// 2. AUTHENTICATION LOGIC (MSAL + PKCE)
// =========================================================

const msalInstance = new msal.PublicClientApplication(msalConfig);
let username = ""; // Stores the logged-in user's email

// Initialize App
document.addEventListener("DOMContentLoaded", async () => {
    initModalTimes();
    fetchRooms();
    
    // Initialize MSAL
    await msalInstance.initialize();

    // Check if user is already signed in (from previous session)
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        handleLoginSuccess(accounts[0]);
    }
});

async function signIn() {
    try {
        // This triggers the Popup with PKCE Flow
        const loginResponse = await msalInstance.loginPopup({
            scopes: ["User.Read"] // Just asking to read their profile
        });
        handleLoginSuccess(loginResponse.account);
    } catch (error) {
        console.error("Login Failed:", error);
    }
}

function signOut() {
    const logoutRequest = {
        account: msalInstance.getAccountByUsername(username),
        mainWindowRedirectUri: window.location.origin
    };
    msalInstance.logoutPopup(logoutRequest);
    username = "";
    updateUI(false);
}

function handleLoginSuccess(account) {
    username = account.username;
    console.log("✅ Authenticated as:", username);
    updateUI(true);
}

function updateUI(isLoggedIn) {
    const userDisplay = document.getElementById("userWelcome");
    const loginBtn = document.getElementById("loginBtn");
    const logoutBtn = document.getElementById("logoutBtn");

    if (userDisplay && isLoggedIn) {
        userDisplay.textContent = `👤 ${username}`;
        userDisplay.style.display = "inline";
        if(loginBtn) loginBtn.style.display = "none";
        if(logoutBtn) logoutBtn.style.display = "inline-block";
    } else if (userDisplay) {
        userDisplay.style.display = "none";
        if(loginBtn) loginBtn.style.display = "inline-block";
        if(logoutBtn) logoutBtn.style.display = "none";
    }
}

// =========================================================
// 3. BOOKING LOGIC
// =========================================================

// Triggered when user clicks "New Booking"
async function handleBookClick() {
    // 1. Check Auth
    if (!username) {
        await signIn(); // Force Login
        // If they close the popup without logging in, stop
        if (!username) return; 
    }
    
    // 2. Pre-fill the modal with their locked email
    const displayEmail = document.getElementById('displayEmail');
    if(displayEmail) displayEmail.value = username;

    // 3. Show Modal
    const modal = new bootstrap.Modal(document.getElementById('bookingModal'));
    modal.show();
}

async function createBooking() {
    const roomEmail = document.getElementById('roomSelect').value;
    const start = new Date(document.getElementById('startTime').value);
    const end = new Date(document.getElementById('endTime').value);
    const subject = document.getElementById('subject').value;

    if (!roomEmail) return alert("Please select a room.");
    if (!username) return alert("Security Check Failed: You are not logged in.");

    try {
        const res = await fetch(`${API_URL}/book`, {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({ 
                subject: subject, 
                room_email: roomEmail, 
                start_time: start.toISOString(), 
                end_time: end.toISOString(), 
                organizer_email: username // ✅ Uses the verified token email
            })
        });
        
        if (res.ok) {
            alert(`✅ Booking Confirmed for ${username}!`);
            // Close Modal safely
            const modalEl = document.getElementById('bookingModal');
            const modal = bootstrap.Modal.getInstance(modalEl);
            modal.hide();
            loadAvailability(); // Refresh grid
        } else {
            const err = await res.json();
            if (res.status === 409) alert("⛔ " + err.detail);
            else alert("❌ Error: " + JSON.stringify(err));
        }
    } catch (e) { alert(e.message); }
}

// =========================================================
// 4. TIMELINE & DATA FETCHING (Existing Logic)
// =========================================================

function initModalTimes() {
    const now = new Date();
    now.setMinutes(now.getMinutes() - now.getTimezoneOffset());
    document.getElementById('startTime').value = now.toISOString().slice(0,16);
    now.setMinutes(now.getMinutes() + 30);
    document.getElementById('endTime').value = now.toISOString().slice(0,16);
}

async function fetchRooms() {
    try {
        const res = await fetch(`${API_URL}/rooms`);
        const data = await res.json();
        const select = document.getElementById('roomSelect');
        select.innerHTML = '<option value="" disabled selected>Select a room...</option>';
        if (data.value) data.value.forEach(r => {
            const opt = document.createElement('option');
            opt.value = r.emailAddress;
            opt.textContent = r.displayName;
            select.appendChild(opt);
        });
    } catch (e) { console.error("API Error:", e); }
}

async function loadAvailability() {
    const roomEmail = document.getElementById('roomSelect').value;
    if (!roomEmail) return;

    const spinner = document.getElementById('loadingSpinner');
    if(spinner) spinner.style.display = "inline";

    const now = new Date();
    const viewStart = new Date(now);
    viewStart.setMinutes(0, 0, 0); 
    const viewEnd = new Date(viewStart.getTime() + HOURS_TO_SHOW * 60 * 60 * 1000);

    try {
        const res = await fetch(`${API_URL}/availability`, {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({ room_email: roomEmail, start_time: viewStart.toISOString(), end_time: viewEnd.toISOString(), time_zone: "UTC" })
        });
        const data = await res.json();
        renderTimeline(data, viewStart, viewEnd);
    } catch (err) { console.error(err); } 
    finally { if(spinner) spinner.style.display = "none"; }
}

function renderTimeline(data, viewStart, viewEnd) {
    const timelineContainer = document.getElementById('timeline');
    timelineContainer.innerHTML = ''; 

    const totalDurationMs = viewEnd - viewStart;
    const totalSlots = HOURS_TO_SHOW * 2; 
    const slotWidthPct = 100 / totalSlots;

    let headerHtml = `<div class="timeline-header">`;
    for (let i = 0; i < totalSlots; i++) {
        let slotTime = new Date(viewStart.getTime() + i * 30 * 60 * 1000);
        let timeStr = slotTime.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
        headerHtml += `<div class="timeline-time-label" style="width:${slotWidthPct}%">${timeStr}</div>`;
    }
    headerHtml += `</div>`;

    let trackHtml = `<div class="timeline-track">`;
    for (let i = 1; i < totalSlots; i++) {
        trackHtml += `<div class="grid-line" style="left: ${i * slotWidthPct}%"></div>`;
    }

    const now = new Date();
    if (now >= viewStart && now <= viewEnd) {
        const nowPct = ((now - viewStart) / totalDurationMs) * 100;
        trackHtml += `<div class="current-time-line" style="left: ${nowPct}%"><div class="current-time-label">NOW</div></div>`;
    }

    const schedule = (data.value && data.value[0]) ? data.value[0] : null; 
    if (schedule && schedule.scheduleItems) {
        schedule.scheduleItems.forEach(item => {
            if (item.status === 'busy') {
                const start = new Date(item.start.dateTime + 'Z');
                const end = new Date(item.end.dateTime + 'Z');
                const leftPct = ((start - viewStart) / totalDurationMs) * 100;
                const widthPct = ((end - start) / totalDurationMs) * 100;

                if (leftPct < 100 && (leftPct + widthPct) > 0) {
                    const safeLeft = Math.max(0, leftPct);
                    const safeWidth = Math.min(widthPct, 100 - safeLeft);
                    trackHtml += `<div class="event-block" style="left:${safeLeft}%; width:${safeWidth}%;" title="${item.subject}"><span>🚫 Occupied</span></div>`;
                }
            }
        });
    }
    trackHtml += `</div>`;
    timelineContainer.innerHTML = headerHtml + trackHtml;
}
