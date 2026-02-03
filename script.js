// =========================================================
// 1. CONFIGURATION
// =========================================================

// 🔴 1. BACKEND URL (Check your Render Dashboard)
const API_URL = "https://booking-a-room-poc.onrender.com"; 

// 🔴 2. AZURE FRONTEND CONFIG (The "SPA" App Registration)
const msalConfig = {
    auth: {
        clientId: "0f759785-1ba8-449d-ba6f-9ba5e8f479d8", 
        authority: "https://login.microsoftonline.com/2b2369a3-0061-401b-97d9-c8c8d92b76f6",
        redirectUri: window.location.origin, 
    },
    cache: {
        cacheLocation: "sessionStorage",
        storeAuthStateInCookie: false,
    }
};

const HOURS_TO_SHOW = 12;
const msalInstance = new msal.PublicClientApplication(msalConfig);
let username = ""; 

// =========================================================
// 2. INITIALIZATION (Runs immediately)
// =========================================================

document.addEventListener("DOMContentLoaded", async () => {
    initModalTimes();
    fetchRooms(); // Load rooms immediately
    
    // Initialize MSAL & Handle Redirects
    try {
        await msalInstance.initialize();
        
        // Check if returning from a login redirect
        const response = await msalInstance.handleRedirectPromise();
        if (response) {
            handleLoginSuccess(response.account);
        } else {
            // Check if already logged in
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) {
                handleLoginSuccess(accounts[0]);
            }
        }
    } catch (e) {
        console.error("MSAL Init Error:", e);
    }
});

// =========================================================
// 3. AUTH LOGIC (PKCE + Redirect Fallback)
// =========================================================

async function signIn() {
    try {
        // Try Popup first
        const loginResponse = await msalInstance.loginPopup({ scopes: ["User.Read"] });
        handleLoginSuccess(loginResponse.account);
    } catch (error) {
        console.warn("Popup blocked or failed. Trying Redirect...", error);
        // Fallback to Redirect (Browser cannot block this)
        await msalInstance.loginRedirect({ scopes: ["User.Read"] });
    }
}

function signOut() {
    const logoutRequest = { 
        account: msalInstance.getAccountByUsername(username),
        mainWindowRedirectUri: window.location.origin 
    };
    msalInstance.logoutPopup(logoutRequest); // Or logoutRedirect
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
// 4. BOOKING LOGIC
// =========================================================

async function handleBookClick() {
    if (!username) {
        await signIn(); 
        if (!username) return; // If invite failed/closed
    }
    const displayEmail = document.getElementById('displayEmail');
    if(displayEmail) displayEmail.value = username;
    const modal = new bootstrap.Modal(document.getElementById('bookingModal'));
    modal.show();
}

async function createBooking() {
    const roomEmail = document.getElementById('roomSelect').value;
    const start = new Date(document.getElementById('startTime').value);
    const end = new Date(document.getElementById('endTime').value);
    const subject = document.getElementById('subject').value;
    
    // Parse Attendees
    const attendeesRaw = document.getElementById('attendees').value;
    let attendeeList = [];
    if (attendeesRaw.trim()) {
        attendeeList = attendeesRaw.split(',').map(email => email.trim());
    }

    if (!roomEmail) return alert("Select a room.");
    if (!username) return alert("Please sign in first.");

    try {
        const res = await fetch(`${API_URL}/book`, {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({ 
                subject: subject, 
                room_email: roomEmail, 
                start_time: start.toISOString(), 
                end_time: end.toISOString(), 
                organizer_email: username,
                attendees: attendeeList
            })
        });
        
        if (res.ok) {
            alert(`✅ Booking Confirmed for ${username}! Invites sent.`);
            bootstrap.Modal.getInstance(document.getElementById('bookingModal')).hide();
            loadAvailability(); 
        } else {
            const err = await res.json();
            if (res.status === 409) alert("⛔ " + err.detail);
            else alert("❌ Error: " + JSON.stringify(err));
        }
    } catch (e) { alert(e.message); }
}

// =========================================================
// 5. TIMELINE & DATA FETCHING
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
        
        if (data.value && data.value.length > 0) {
            data.value.forEach(r => {
                const opt = document.createElement('option');
                opt.value = r.emailAddress;
                opt.textContent = r.displayName;
                select.appendChild(opt);
            });
        } else {
            select.innerHTML = '<option disabled>No rooms found</option>';
        }
    } catch (e) { 
        console.error("Fetch Rooms Error:", e); 
    }
}

async function loadAvailability() {
    const roomEmail = document.getElementById('roomSelect').value;
    if (!roomEmail) return;

    document.getElementById('loadingSpinner').style.display = "inline";

    const now = new Date();
    const viewStart = new Date(now);
    viewStart.setMinutes(0, 0, 0); 
    const viewEnd = new Date(viewStart.getTime() + HOURS_TO_SHOW * 60 * 60 * 1000);

    try {
        const res = await fetch(`${API_URL}/availability`, {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({ 
                room_email: roomEmail, 
                start_time: viewStart.toISOString(), 
                end_time: viewEnd.toISOString(), 
                time_zone: "UTC" 
            })
        });
        const data = await res.json();
        renderTimeline(data, viewStart, viewEnd);
    } catch (err) { console.error(err); } 
    finally { document.getElementById('loadingSpinner').style.display = "none"; }
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
