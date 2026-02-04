// =========================================================
// 1. CONFIGURATION
// =========================================================

const API_URL = "https://booking-a-room-poc.onrender.com"; 

const msalConfig = {
    auth: {
        clientId: "0f759785-1ba8-449d-ba6f-9ba5e8f479d8",
        authority: "https://login.microsoftonline.com/2b2369a3-0061-401b-97d9-c8c8d92b76f6",
        redirectUri: window.location.origin, 
    },
    cache: { cacheLocation: "sessionStorage" }
};

// 🔴 REQUEST PERMISSION TO WRITE TO *ANY* USER'S CALENDAR WHO LOGS IN
const loginRequest = {
    scopes: ["User.Read", "Calendars.ReadWrite"] 
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
let username = ""; 
let availableRooms = []; 

// =========================================================
// 2. INITIALIZATION & AUTH
// =========================================================

document.addEventListener("DOMContentLoaded", async () => {
    initModalTimes();
    await fetchRooms();
    
    try {
        await msalInstance.initialize();
        const response = await msalInstance.handleRedirectPromise();
        if (response) {
            handleLoginSuccess(response.account);
        } else {
            const accounts = msalInstance.getAllAccounts();
            if (accounts.length > 0) handleLoginSuccess(accounts[0]);
        }
    } catch (e) { console.error(e); }
});

async function signIn() { 
    try { 
        // Use the new scope request (ReadWrite)
        await msalInstance.loginRedirect(loginRequest); 
    } catch (e) { console.error(e); } 
}

function signOut() { msalInstance.logoutPopup(); }

function handleLoginSuccess(acc) { 
    username = acc.username; 
    document.getElementById("userWelcome").textContent = `👤 ${username}`; 
    document.getElementById("userWelcome").style.display="inline"; 
    document.getElementById("loginBtn").style.display="none"; 
    document.getElementById("logoutBtn").style.display="inline-block"; 
}

// =========================================================
// 3. BOOKING LOGIC (DYNAMIC USER)
// =========================================================

async function createBooking() {
    // 1. Get current form values
    const index = document.getElementById('roomSelect').value;
    if (!index) return alert("Select a room.");
    
    const room = availableRooms[index];
    const roomEmail = room.emailAddress;
    const subject = document.getElementById('subject').value;
    const filiale = document.getElementById('filiale').value; 
    const desc = document.getElementById('description').value;
    const startInput = document.getElementById('startTime').value;
    const endInput = document.getElementById('endTime').value;
    
    const attendeesRaw = document.getElementById('attendees').value;
    let attendeeList = attendeesRaw.trim() ? attendeesRaw.split(',').map(e => e.trim()) : [];

    if (!username) return alert("Please sign in first.");
    if (!filiale) return alert("Please enter the Filiale name.");

    // 2. Get the Current User's Token
    let accessToken = "";
    try {
        const account = msalInstance.getAllAccounts()[0];
        const tokenResp = await msalInstance.acquireTokenSilent({
            ...loginRequest,
            account: account
        });
        accessToken = tokenResp.accessToken;
    } catch (e) {
        console.warn("Silent token failed", e);
        return alert("Session expired. Please sign in again.");
    }

    // 3. Send to Backend
    try {
        const res = await fetch(`${API_URL}/book`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${accessToken}`
            },
            body: JSON.stringify({ 
                subject: subject, 
                room_email: roomEmail, 
                start_time: new Date(startInput).toISOString(), 
                end_time: new Date(endInput).toISOString(), 
                organizer_email: username,
                attendees: attendeeList,
                filiale: filiale,       
                description: desc       
            })
        });
        
        if (res.ok) {
            // 4. ✅ CUSTOMIZED CONFIRMATION MESSAGE
            const startTimeFormatted = new Date(startInput).toLocaleString();
            const endTimeFormatted = new Date(endInput).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
            const inviteesMsg = attendeeList.length > 0 ? attendeeList.join(", ") : "None";

            alert(
                `✅ BOOKING CONFIRMED\n\n` +
                `📅 Time: ${startTimeFormatted} - ${endTimeFormatted}\n` +
                `🏢 Unit: ${filiale}\n` +
                `📝 Subject: ${subject}\n` +
                `💡 Reason: ${desc || "N/A"}\n` +
                `👥 Invitees: ${inviteesMsg}\n\n` +
                `The meeting has been added to your calendar.`
            );

            bootstrap.Modal.getInstance(document.getElementById('bookingModal')).hide();
            loadAvailability(roomEmail); 
        } else {
            const err = await res.json();
            alert("Error: " + (err.detail || JSON.stringify(err)));
        }
    } catch (e) { alert("Network Error: " + e.message); }
}

// =========================================================
// 4. HELPERS
// =========================================================
async function fetchRooms() { try { const res = await fetch(`${API_URL}/rooms`); const data = await res.json(); if (data.value) { availableRooms = data.value; const select = document.getElementById('roomSelect'); select.innerHTML = '<option value="" disabled selected>Select a room...</option>'; availableRooms.forEach((r, index) => { const opt = document.createElement('option'); opt.value = index; opt.textContent = `${r.displayName}  [ ${r.department} - ${r.floor} ]`; select.appendChild(opt); }); } } catch (e) { console.error("Fetch Rooms Error:", e); } }
function handleRoomChange() { const index = document.getElementById('roomSelect').value; const room = availableRooms[index]; if (room) { loadAvailability(room.emailAddress); } }
async function loadAvailability(email) { if (!email) return; document.getElementById('loadingSpinner').style.display = "inline"; const now = new Date(); const viewStart = new Date(now); viewStart.setMinutes(0, 0, 0); const viewEnd = new Date(viewStart.getTime() + 12 * 60 * 60 * 1000); try { const res = await fetch(`${API_URL}/availability`, { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({ room_email: email, start_time: viewStart.toISOString(), end_time: viewEnd.toISOString(), time_zone: "UTC" }) }); const data = await res.json(); renderTimeline(data, viewStart, viewEnd); } catch (err) { console.error(err); } finally { document.getElementById('loadingSpinner').style.display = "none"; } }
function handleBookClick() { if(!username) { signIn(); return; } document.getElementById('displayEmail').value = username; new bootstrap.Modal(document.getElementById('bookingModal')).show(); }
function initModalTimes() { const now=new Date(); now.setMinutes(now.getMinutes()-now.getTimezoneOffset()); document.getElementById('startTime').value=now.toISOString().slice(0,16); now.setMinutes(now.getMinutes()+30); document.getElementById('endTime').value=now.toISOString().slice(0,16); }
function renderTimeline(data, viewStart, viewEnd) { const timelineContainer = document.getElementById('timeline'); timelineContainer.innerHTML = ''; const totalDurationMs = viewEnd - viewStart; const totalSlots = 12 * 2; const slotWidthPct = 100 / totalSlots; let headerHtml = `<div class="timeline-header">`; for (let i = 0; i < totalSlots; i++) { let slotTime = new Date(viewStart.getTime() + i * 30 * 60 * 1000); headerHtml += `<div class="timeline-time-label" style="width:${slotWidthPct}%">${slotTime.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}</div>`; } headerHtml += `</div><div class="timeline-track">`; let trackHtml = ''; for (let i = 1; i < totalSlots; i++) trackHtml += `<div class="grid-line" style="left: ${i * slotWidthPct}%"></div>`; const schedule = (data.value && data.value[0]) ? data.value[0] : null; if (schedule && schedule.scheduleItems) { schedule.scheduleItems.forEach(item => { if (item.status === 'busy') { const start = new Date(item.start.dateTime + 'Z'); const end = new Date(item.end.dateTime + 'Z'); const leftPct = ((start - viewStart) / totalDurationMs) * 100; const widthPct = ((end - start) / totalDurationMs) * 100; if (leftPct < 100 && (leftPct + widthPct) > 0) { trackHtml += `<div class="event-block" style="left:${Math.max(0, leftPct)}%; width:${Math.min(widthPct, 100 - Math.max(0, leftPct))}%;" title="${item.subject}"><span>🚫 Busy</span></div>`; } } }); } timelineContainer.innerHTML = headerHtml + trackHtml + `</div>`; }
