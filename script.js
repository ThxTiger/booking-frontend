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

// ================= LOGIC: FIND NEXT MEETING =================
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
            const now = new Date();
            const start = new Date(event.start.dateTime + 'Z');
            const end = new Date(event.end.dateTime + 'Z');

            // 1. Is meeting OVER?
            if (now >= end) {
                banner.style.display = "none";
                overlay.classList.add('d-none');
                stopCheckInCountdown(); stopMeetingEndTimer();
                return;
            }

            // 2. Is meeting ALREADY CHECKED IN?
            if (event.categories && event.categories.includes("Checked-In")) {
                 banner.style.display = "none";
                 stopCheckInCountdown();
                 if (overlay.classList.contains('d-none')) showMeetingMode(event);
                 return;
            } 
            
            // 3. SHOW BANNER (Active OR Upcoming)
            banner.style.display = "block";
            overlay.classList.add('d-none');
            
            document.getElementById('bannerSubject').textContent = event.subject;
            document.getElementById('bannerOrganizer').textContent = event.organizer?.emailAddress?.name || "Unknown";
            const btn = document.getElementById('realCheckInBtn');
            btn.onclick = () => performCheckIn(roomEmail, event.id, event);

            // 🔴 STATE CHECK: Is it Active Now or Future?
            // "Active" = Starts within 15 mins OR already started
            const timeDiff = start - now;
            const minutesUntilStart = Math.floor(timeDiff / 60000);

            if (minutesUntilStart > 15) {
                // FUTURE MODE: "Next Meeting"
                document.getElementById('bannerStatusTitle').textContent = "📅 Next Meeting";
                document.getElementById('bannerBadge').textContent = "STARTS IN";
                document.getElementById('bannerBadge').className = "badge bg-info mb-1";
                // Disable button if too early? (Optional, but user asked for check-in option)
                // We leave it enabled or you can disable: btn.disabled = true;
                
                // Countdown to START
                startGenericCountdown(start, "checkInTimer");
            } else {
                // ACTION MODE: "Check-In Required"
                document.getElementById('bannerStatusTitle').textContent = "⚠️ Check-In Required";
                document.getElementById('bannerBadge').textContent = "DEADLINE";
                document.getElementById('bannerBadge').className = "badge bg-danger mb-1";
                
                // Deadline = Start + 5 mins
                const deadline = new Date(start.getTime() + 5*60000);
                startGenericCountdown(deadline, "checkInTimer", "EXPIRED");
            }

        } else {
            banner.style.display = "none";
            overlay.classList.add('d-none');
            stopCheckInCountdown(); stopMeetingEndTimer();
        }
    } catch (e) { console.error(e); }
}

function startGenericCountdown(targetDate, elementId, expireText="00:00") {
    if (checkInCountdown) clearInterval(checkInCountdown);
    
    checkInCountdown = setInterval(() => {
        const now = new Date().getTime();
        const distance = targetDate.getTime() - now;
        const timerEl = document.getElementById(elementId);
        
        if (distance < 0) {
            timerEl.textContent = expireText;
        } else {
            const h = Math.floor((distance % (1000 * 60 * 60 * 24)) / (1000 * 60 * 60));
            const m = Math.floor((distance % (1000 * 60 * 60)) / (1000 * 60));
            const s = Math.floor((distance % (1000 * 60)) / 1000);
            
            if (h > 0) timerEl.textContent = `${h}h ${m}m`;
            else timerEl.textContent = `${m}m ${s}s`;
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
            currentLockedEvent = null; stopMeetingEndTimer(); checkForActiveMeeting();
        } else { alert(`⛔ ACCESS DENIED\nVerified: ${verifiedEmail}`); }
    } catch (e) { console.error(e); }
}

// ================= BOOKING (FIXED) =================
async function createBooking() {
    if (!username) return alert("Please sign in first.");
    const index = document.getElementById('roomSelect').value;
    if (!index) return alert("Select a room.");
    const roomEmail = availableRooms[index].emailAddress;
    
    // ... form values ...
    const subject = document.getElementById('subject').value;
    const filiale = document.getElementById('filiale').value; 
    const desc = document.getElementById('description').value;
    const startInput = document.getElementById('startTime').value;
    const endInput = document.getElementById('endTime').value;
    const attendeesRaw = document.getElementById('attendees').value;
    let attendeeList = attendeesRaw.trim() ? attendeesRaw.split(',').map(e => e.trim()) : [];

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
        } catch (err) { return alert("Permission denied: " + err.message); }
    }

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
            // FORCE CLOSE MODAL manually via DOM if bootstrap instance fails
            const modalEl = document.getElementById('bookingModal');
            const modal = bootstrap.Modal.getInstance(modalEl);
            if (modal) modal.hide();
            else modalEl.classList.remove('show'); // Fallback
            
            loadAvailability(roomEmail); 
        } else {
            const err = await res.json();
            alert("Error: " + (err.detail || JSON.stringify(err)));
        }
    } catch (e) { alert("Network Error: " + e.message); }
}

// ... (Helpers: handleRoomChange, loadAvailability, fetchRooms etc. same as before) ...
async function fetchRooms() { try { const res = await fetch(`${API_URL}/rooms`); const data = await res.json(); if (data.value) { availableRooms = data.value; const select = document.getElementById('roomSelect'); select.innerHTML = '<option value="" disabled selected>Select a room...</option>'; availableRooms.forEach((r, index) => { const opt = document.createElement('option'); opt.value = index; opt.textContent = `${r.displayName}  [ ${r.department} - ${r.floor} ]`; select.appendChild(opt); }); } } catch (e) { console.error(e); } }
function handleRoomChange() { const index = document.getElementById('roomSelect').value; const room = availableRooms[index]; if (room) { loadAvailability(room.emailAddress); checkForActiveMeeting(); } }
async function loadAvailability(email) { if (!email) return; document.getElementById('loadingSpinner').style.display = "inline"; const now = new Date(); const viewStart = new Date(now); viewStart.setMinutes(0, 0, 0); const viewEnd = new Date(viewStart.getTime() + 12 * 60 * 60 * 1000); try { const res = await fetch(`${API_URL}/availability`, { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({ room_email: email, start_time: viewStart.toISOString(), end_time: viewEnd.toISOString(), time_zone: "UTC" }) }); const data = await res.json(); renderTimeline(data, viewStart, viewEnd); } catch (err) { console.error(err); } finally { document.getElementById('loadingSpinner').style.display = "none"; } }
function handleBookClick() { if(!username) { signIn(); return; } document.getElementById('displayEmail').value = username; new bootstrap.Modal(document.getElementById('bookingModal')).show(); }
function initModalTimes() { const now=new Date(); now.setMinutes(now.getMinutes()-now.getTimezoneOffset()); document.getElementById('startTime').value=now.toISOString().slice(0,16); now.setMinutes(now.getMinutes()+30); document.getElementById('endTime').value=now.toISOString().slice(0,16); }
function renderTimeline(data, viewStart, viewEnd) { const timelineContainer = document.getElementById('timeline'); timelineContainer.innerHTML = ''; const totalDurationMs = viewEnd - viewStart; const totalSlots = 12 * 2; const slotWidthPct = 100 / totalSlots; let headerHtml = `<div class="timeline-header">`; for (let i = 0; i < totalSlots; i++) { let slotTime = new Date(viewStart.getTime() + i * 30 * 60 * 1000); headerHtml += `<div class="timeline-time-label" style="width:${slotWidthPct}%">${slotTime.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'})}</div>`; } headerHtml += `</div><div class="timeline-track">`; let trackHtml = ''; for (let i = 1; i < totalSlots; i++) trackHtml += `<div class="grid-line" style="left: ${i * slotWidthPct}%"></div>`; const schedule = (data.value && data.value[0]) ? data.value[0] : null; if (schedule && schedule.scheduleItems) { schedule.scheduleItems.forEach(item => { if (item.status === 'busy') { const start = new Date(item.start.dateTime + 'Z'); const end = new Date(item.end.dateTime + 'Z'); const leftPct = ((start - viewStart) / totalDurationMs) * 100; const widthPct = ((end - start) / totalDurationMs) * 100; if (leftPct < 100 && (leftPct + widthPct) > 0) { trackHtml += `<div class="event-block" style="left:${Math.max(0, leftPct)}%; width:${Math.min(widthPct, 100 - Math.max(0, leftPct))}%;" title="${item.subject}"><span>🚫 Busy</span></div>`; } } }); } timelineContainer.innerHTML = headerHtml + trackHtml + `</div>`; }
