// ================= CONFIGURATION =================
const API_URL = "https://booking-a-room-poc.onrender.com"; 

const msalConfig = {
    auth: {
        clientId: "0f759785-1ba8-449d-ba6f-9ba5e8f479d8", 
        authority: "https://login.microsoftonline.com/2b2369a3-0061-401b-97d9-c8c8d92b76f6",
        redirectUri: window.location.origin, 
    },
    cache: { cacheLocation: "sessionStorage" }
};
const loginRequest = { scopes: ["User.Read", "Calendars.ReadWrite"] };
const msalInstance = new msal.PublicClientApplication(msalConfig);

let username = ""; 
let availableRooms = []; 
let currentLockedEvent = null; 
let checkInCountdown = null;
let meetingEndInterval = null;
let sessionTimeout = null;

// ================= INITIALIZATION =================
document.addEventListener("DOMContentLoaded", async () => {
    initModalTimes();
    await fetchRooms();
    setInterval(checkForActiveMeeting, 5000); 
    
    try {
        await msalInstance.initialize();
        const response = await msalInstance.handleRedirectPromise();
        if (response) handleLoginSuccess(response.account);
    } catch (e) { console.error(e); }
});

async function signIn() { try { await msalInstance.loginRedirect(loginRequest); } catch (e) { console.error(e); } }
function signOut() { 
    username = ""; 
    document.getElementById("userWelcome").style.display="none"; 
    document.getElementById("loginBtn").style.display="inline-block"; 
    document.getElementById("logoutBtn").style.display="none"; 
    if(sessionTimeout) clearTimeout(sessionTimeout);
}

function handleLoginSuccess(acc) { 
    username = acc.username; 
    document.getElementById("userWelcome").textContent = `👤 ${username}`; 
    document.getElementById("userWelcome").style.display="inline"; 
    document.getElementById("loginBtn").style.display="none"; 
    document.getElementById("logoutBtn").style.display="inline-block"; 
    if(sessionTimeout) clearTimeout(sessionTimeout);
    sessionTimeout = setTimeout(() => { alert("Session Expired."); signOut(); }, 120000);
}

// ================= CHECK-IN LOGIC (Strict Per Room) =================
async function checkForActiveMeeting() {
    const index = document.getElementById('roomSelect').value;
    if (!index) return;
    const roomEmail = availableRooms[index].emailAddress;

    try {
        const res = await fetch(`${API_URL}/active-meeting?room_email=${roomEmail}`);
        const event = await res.json();
        
        const banner = document.getElementById('checkInBanner');
        const overlay = document.getElementById('meetingInProgressOverlay');

        if (event) {
            // Check Auto-Exit (Time Over?)
            const now = new Date();
            const endTime = new Date(event.end.dateTime + 'Z');
            if (now >= endTime) {
                banner.style.display = "none";
                overlay.classList.add('d-none');
                stopCheckInCountdown();
                stopMeetingEndTimer();
                return;
            }

            // Populate Info (Banner & Overlay)
            const subjectText = event.subject; // "Filiale : Description"
            const organizerText = event.organizer?.emailAddress?.name || "Unknown";

            // CASE 1: ALREADY CHECKED IN
            if (event.categories && event.categories.includes("Checked-In")) {
                 banner.style.display = "none";
                 stopCheckInCountdown();
                 
                 if (overlay.classList.contains('d-none')) {
                     showMeetingMode(event);
                 }
            } 
            // CASE 2: WAITING FOR CHECK-IN
            else {
                 banner.style.display = "block";
                 overlay.classList.add('d-none');
                 
                 document.getElementById('bannerSubject').textContent = subjectText;
                 document.getElementById('bannerOrganizer').textContent = organizerText;
                 
                 const btn = document.getElementById('realCheckInBtn');
                 btn.onclick = () => performCheckIn(roomEmail, event.id, event);
                 
                 startCheckInCountdown(event.start.dateTime);
            }
        } else {
            banner.style.display = "none";
            overlay.classList.add('d-none');
            stopCheckInCountdown();
            stopMeetingEndTimer();
        }
    } catch (e) { console.error(e); }
}

function startCheckInCountdown(startTimeStr) {
    if (checkInCountdown) clearInterval(checkInCountdown);
    const startTime = new Date(startTimeStr + 'Z').getTime();
    const deadline = startTime + (5 * 60 * 1000); 

    checkInCountdown = setInterval(() => {
        const now = new Date().getTime();
        const distance = deadline - now;
        const timerEl = document.getElementById('checkInTimer');
        if (distance < 0) { clearInterval(checkInCountdown); timerEl.textContent = "EXPIRED"; } 
        else {
            const m = Math.floor((distance % (1000 * 60 * 60)) / (1000 * 60));
            const s = Math.floor((distance % (1000 * 60)) / 1000);
            timerEl.textContent = `${m}m ${s}s`;
        }
    }, 1000);
}
function stopCheckInCountdown() { if (checkInCountdown) clearInterval(checkInCountdown); }

async function performCheckIn(roomEmail, eventId, eventDetails) {
    stopCheckInCountdown();
    document.getElementById('checkInBanner').style.display = "none";
    try {
        const res = await fetch(`${API_URL}/checkin`, {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({ room_email: roomEmail, event_id: eventId })
        });
        if (res.ok) showMeetingMode(eventDetails);
        else { alert("Check-in failed."); checkForActiveMeeting(); }
    } catch (e) { alert(e.message); }
}

// ================= RED SCREEN & AUTO EXIT =================
function showMeetingMode(event) {
    currentLockedEvent = event;
    const overlay = document.getElementById('meetingInProgressOverlay');
    const start = new Date(event.start.dateTime + 'Z').toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    const end = new Date(event.end.dateTime + 'Z').toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    
    document.getElementById('overlaySubject').textContent = event.subject;
    document.getElementById('overlayOrganizer').textContent = `Booked by: ${event.organizer?.emailAddress?.name}`;
    document.getElementById('overlayTime').textContent = `${start} - ${end}`;
    
    overlay.classList.remove('d-none');
    startMeetingEndTimer(event.end.dateTime);
}

function startMeetingEndTimer(endTimeStr) {
    if (meetingEndInterval) clearInterval(meetingEndInterval);
    const endTime = new Date(endTimeStr + 'Z').getTime();

    meetingEndInterval = setInterval(() => {
        const now = new Date().getTime();
        const distance = endTime - now;
        const timerEl = document.getElementById('meetingEndTimer');
        if (distance < 0) {
            clearInterval(meetingEndInterval);
            document.getElementById('meetingInProgressOverlay').classList.add('d-none');
            checkForActiveMeeting();
        } else {
            const m = Math.floor((distance % (1000 * 60 * 60)) / (1000 * 60));
            const s = Math.floor((distance % (1000 * 60)) / 1000);
            timerEl.textContent = `${m}m ${s}s`;
        }
    }, 1000);
}
function stopMeetingEndTimer() { if (meetingEndInterval) clearInterval(meetingEndInterval); }

async function secureExitMeetingMode() {
    const organizerEmail = currentLockedEvent.organizer.emailAddress.address.toLowerCase();
    try {
        const loginResp = await msalInstance.loginPopup({ scopes: ["User.Read"], prompt: "login" });
        const verifiedEmail = loginResp.account.username.toLowerCase();
        if (verifiedEmail === organizerEmail) {
            document.getElementById('meetingInProgressOverlay').classList.add('d-none');
            currentLockedEvent = null;
            stopMeetingEndTimer();
            checkForActiveMeeting();
        } else {
            alert(`⛔ ACCESS DENIED\n\nVerified: ${verifiedEmail}\nOrganizer: ${organizerEmail}`);
        }
    } catch (e) { console.error(e); }
}

// ================= BOOKING FUNCTION (FIXED) =================
async function createBooking() {
    // 1. Check Login
    if (!username) return alert("Please sign in.");
    
    // 2. Check Room Selection
    const index = document.getElementById('roomSelect').value;
    if (!index) return alert("Select a room.");
    const roomEmail = availableRooms[index].emailAddress;
    
    // 3. Get Form Data
    const subject = document.getElementById('subject').value;
    const filiale = document.getElementById('filiale').value; 
    const desc = document.getElementById('description').value;
    const startInput = document.getElementById('startTime').value;
    const endInput = document.getElementById('endTime').value;
    const attendeesRaw = document.getElementById('attendees').value;
    let attendeeList = attendeesRaw.trim() ? attendeesRaw.split(',').map(e => e.trim()) : [];

    // 4. Get Token (Robust)
    let accessToken = "";
    try {
        const account = msalInstance.getAllAccounts()[0];
        const tokenResp = await msalInstance.acquireTokenSilent({ ...loginRequest, account: account });
        accessToken = tokenResp.accessToken;
    } catch (e) {
        try {
            const tokenResp = await msalInstance.acquireTokenPopup(loginRequest);
            accessToken = tokenResp.accessToken;
            handleLoginSuccess(tokenResp.account);
        } catch (err) { return alert("Permission denied. Could not acquire token."); }
    }

    // 5. Send to Backend
    try {
        const res = await fetch(`${API_URL}/book`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Authorization': `Bearer ${accessToken}` },
            body: JSON.stringify({ 
                subject: subject, room_email: roomEmail, 
                start_time: new Date(startInput).toISOString(), end_time: new Date(endInput).toISOString(), 
                organizer_email: username, attendees: attendeeList, filiale: filiale, description: desc 
            })
        });
        
        if (res.ok) {
            alert(`✅ Booking Confirmed!`);
            bootstrap.Modal.getInstance(document.getElementById('bookingModal')).hide();
            loadAvailability(roomEmail); 
        } else {
            // FIX: Ensure error is readable
            const err = await res.json();
            alert("Error: " + (err.detail || JSON.stringify(err)));
        }
    } catch (e) { 
        console.error("Booking Error:", e);
        alert("Network Error: Check console for details."); 
    }
}

async function fetchRooms() { 
    try { 
        const res = await fetch(`${API_URL}/rooms`); 
        const data = await res.json(); 
        if (data.value) { 
            availableRooms = data.value; 
            const select = document.getElementById('roomSelect'); 
            select.innerHTML = '<option value="" disabled selected>Select a room...</option>'; 
            availableRooms.forEach((r, index) => { 
                const opt = document.createElement('option'); 
                opt.value = index; 
                opt.textContent = `${r.displayName}  [ ${r.department} - ${r.floor} ]`; 
                select.appendChild(opt); 
            }); 
        } 
    } catch (e) { console.error(e); } 
}

// ... (Keep helpers: handleRoomChange, loadAvailability, handleBookClick, initModalTimes, renderTimeline) ...
function handleRoomChange() { const index = document.getElementById('roomSelect').value; const room = availableRooms[index]; if (room) { loadAvailability(room.emailAddress); checkForActiveMeeting(); } }
async function loadAvailability(email) { if (!email) return; document.getElementById('loadingSpinner').style.display = "inline"; const now = new Date(); const viewStart = new Date(now); viewStart.setMinutes(0, 0, 0); const viewEnd = new Date(viewStart.getTime() + 12 * 60 * 60 * 1000); try { const res = await fetch(`${API_URL}/availability`, { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({ room_email: email, start_time: viewStart.toISOString(), end_time: viewEnd.toISOString(), time_zone: "UTC" }) }); const data = await res.json(); renderTimeline(data, viewStart, viewEnd); } catch (err) { console.error(err); } finally { document.getElementById('loadingSpinner').style.display = "none"; } }
function handleBookClick() { if(!username) { signIn(); return; } document.getElementById('displayEmail').value = username; new bootstrap.Modal(document.getElementById('bookingModal')).show(); }
function initModalTimes() { const now=new Date(); now.setMinutes(now.getMinutes()-now.getTimezoneOffset()); document.getElementById('startTime').value=now.toISOString().slice(0,16); now.setMinutes(now.getMinutes()+30); document.getElementById('endTime').value=now.toISOString().slice(0,16); }
function renderTimeline(data, viewStart, viewEnd) { const timelineContainer = document.getElementById('timeline'); timelineContainer.innerHTML = ''; const totalDurationMs = viewEnd - viewStart; const totalSlots = 12 * 2; const slotWidthPct = 100 / totalSlots; let headerHtml = `<div class="timeline-header">`; for (let i = 0; i < totalSlots; i++) { let slotTime = new Date(viewStart.getTime() + i * 30 * 60 * 1000); headerHtml += `<div class="timeline-time-label" style="width:${slotWidthPct}%">${slotTime.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}</div>`; } headerHtml += `</div><div class="timeline-track">`; let trackHtml = ''; for (let i = 1; i < totalSlots; i++) trackHtml += `<div class="grid-line" style="left: ${i * slotWidthPct}%"></div>`; const schedule = (data.value && data.value[0]) ? data.value[0] : null; if (schedule && schedule.scheduleItems) { schedule.scheduleItems.forEach(item => { if (item.status === 'busy') { const start = new Date(item.start.dateTime + 'Z'); const end = new Date(item.end.dateTime + 'Z'); const leftPct = ((start - viewStart) / totalDurationMs) * 100; const widthPct = ((end - start) / totalDurationMs) * 100; if (leftPct < 100 && (leftPct + widthPct) > 0) { trackHtml += `<div class="event-block" style="left:${Math.max(0, leftPct)}%; width:${Math.min(widthPct, 100 - Math.max(0, leftPct))}%;" title="${item.subject}"><span>🚫 Busy</span></div>`; } } }); } timelineContainer.innerHTML = headerHtml + trackHtml + `</div>`; }
