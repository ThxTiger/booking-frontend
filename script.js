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

const myMSALObj = new msal.PublicClientApplication(msalConfig);

let username = ""; 
let availableRooms = []; 
let currentLockedEvent = null; 
let checkInCountdown = null;
let meetingEndInterval = null;
let sessionTimeout = null;
let isAuthInProgress = false; 
let manuallyUnlockedEventId = null;

// Track changes to auto-refresh timeline
let lastKnownEventId = "init"; 

// ================= INITIALIZATION =================
document.addEventListener("DOMContentLoaded", async () => {
    initModalTimes();
    await fetchRooms();
    
    // Heartbeat (5s)
    setInterval(checkForActiveMeeting, 5000); 
    
    // Timeline Auto-Refresh (60s backup)
    setInterval(() => {
        const index = document.getElementById('roomSelect').value;
        if (index) loadAvailability(availableRooms[index].emailAddress);
    }, 60000);

    try {
        await myMSALObj.initialize();
        const response = await myMSALObj.handleRedirectPromise();
        if (response) handleLoginSuccess(response.account);
    } catch (e) { console.error(e); }
});

async function signIn() { 
    if(isAuthInProgress) return;
    isAuthInProgress = true;
    try { await myMSALObj.loginRedirect(loginRequest); } 
    catch (e) { console.error(e); } 
    finally { isAuthInProgress = false; }
}

function signOut() { 
    username = ""; 
    document.getElementById("userWelcome").style.display="none"; 
    document.getElementById("loginBtn").style.display="inline-block"; 
    document.getElementById("logoutBtn").style.display="none"; 
    
    if(sessionTimeout) clearTimeout(sessionTimeout);
    stopCheckInCountdown();
    stopMeetingEndTimer();

    // 🧠 MEMORY TRICK: Keep details visible if Red Screen is active
    const overlay = document.getElementById('meetingInProgressOverlay');
    if (overlay.classList.contains('d-none')) {
        document.getElementById('bannerSubject').textContent = "";
        document.getElementById('bannerOrganizer').textContent = "";
    }

    checkForActiveMeeting();
}

function handleLoginSuccess(acc) { 
    username = acc.username; 
    document.getElementById("userWelcome").textContent = `👤 ${username}`; 
    document.getElementById("userWelcome").style.display="inline"; 
    document.getElementById("loginBtn").style.display="none"; 
    document.getElementById("logoutBtn").style.display="inline-block"; 
    
    if(sessionTimeout) clearTimeout(sessionTimeout);
    
    // Silent Session Expiry
    sessionTimeout = setTimeout(() => { 
        console.log("Session timed out. Locking Kiosk."); 
        signOut(); 
    }, 120000); 
}

// ================= CORE LOGIC =================

async function getAuthToken() {
    try {
        const account = myMSALObj.getAllAccounts()[0];
        if (!account) return null;
        const response = await myMSALObj.acquireTokenSilent({ scopes: ["User.Read"], account: account });
        return response.accessToken;
    } catch (error) { return null; }
}

async function checkForActiveMeeting() {
    const index = document.getElementById('roomSelect').value;
    if (!index) return;
    const roomEmail = availableRooms[index].emailAddress;

    try {
        const token = await getAuthToken();
        const headers = { "Content-Type": "application/json" };
        if (token) headers["Authorization"] = `Bearer ${token}`;

        const res = await fetch(`${API_URL}/active-meeting?room_email=${roomEmail}`, {
            method: "GET", headers: headers
        });

        if (res.status === 401) return;

        let event = await res.json();

        // Auto-Refresh Timeline on Status Change
        const currentId = event ? event.id : "free";
        if (lastKnownEventId !== "init" && lastKnownEventId !== currentId) {
            console.log("🔄 Status Change Detected! Refreshing Timeline...");
            loadAvailability(roomEmail);
        }
        lastKnownEventId = currentId;

        const banner = document.getElementById('checkInBanner');
        const overlay = document.getElementById('meetingInProgressOverlay');

        // CASE 1: NO MEETING
        if (!event) {
            banner.style.display = "none";
            overlay.classList.add('d-none');
            stopCheckInCountdown(); stopMeetingEndTimer();
            return;
        }

        const now = new Date();
        const start = new Date(event.start.dateTime + 'Z');
        const end = new Date(event.end.dateTime + 'Z');

        // CASE 2: ENDED
        if (now >= end) {
            banner.style.display = "none";
            overlay.classList.add('d-none');
            return;
        }

        // --- MEMORY PROTECTION LOGIC ---
        const isRedScreenActive = !overlay.classList.contains('d-none');
        const isSameEvent = (currentLockedEvent && currentLockedEvent.id === event.id);
        
        let displaySubject = event.subject;
        let displayOrganizer = event.organizer?.emailAddress?.name || "Unknown";

        if (isRedScreenActive && isSameEvent && event.subject === "Busy") {
            displaySubject = document.getElementById('overlaySubject').textContent;
            displayOrganizer = document.getElementById('overlayOrganizer').textContent.replace("Booked by: ", "");
        } else {
             let cleanSubject = (event.subject || "").trim();
             const cleanOrg = displayOrganizer.trim();
             if (cleanSubject.toLowerCase() === cleanOrg.toLowerCase() || cleanSubject === "") {
                if (event.bodyPreview && event.bodyPreview.includes("Filiale")) {
                    const match = event.bodyPreview.match(/Filiale\s*:\s*(.*?)(\n|$)/i);
                    if (match) cleanSubject = match[1].trim();
                } else {
                    cleanSubject = "Private Meeting";
                }
             }
             displaySubject = cleanSubject;
        }
        
        document.getElementById('bannerSubject').textContent = displaySubject;
        document.getElementById('bannerOrganizer').textContent = displayOrganizer;

        // FUTURE
        if (now < start) {
            banner.style.display = "block";
            overlay.classList.add('d-none'); 
            document.getElementById('bannerStatusTitle').textContent = "📅 Next Meeting";
            const badge = document.getElementById('bannerBadge');
            badge.className = "badge bg-info mb-1"; badge.textContent = "STARTS IN";
            document.getElementById('realCheckInBtn').style.display = "none";
            startGenericCountdown(start, "checkInTimer", "STARTING...");
            return; 
        }

        // ACTIVE & CHECKED IN
        if (event.categories && event.categories.includes("Checked-In")) {
             banner.style.display = "none"; 
             stopCheckInCountdown();
             if (overlay.classList.contains('d-none') && event.id !== manuallyUnlockedEventId) {
                 showMeetingMode(event);
             }
             return;
        } 
        
        // ACTIVE & PENDING CHECK-IN
        if (event.id !== manuallyUnlockedEventId) manuallyUnlockedEventId = null;
        banner.style.display = "block";
        overlay.classList.add('d-none'); 
        document.getElementById('bannerStatusTitle').textContent = "⚠️ Meeting Started! Confirm Presence";
        const badge = document.getElementById('bannerBadge');
        badge.className = "badge bg-danger mb-1"; badge.textContent = "AUTO-CANCEL IN";
        const btn = document.getElementById('realCheckInBtn');
        btn.style.display = "inline-block";
        btn.onclick = () => performCheckIn(roomEmail, event.id, event);
        const deadline = new Date(start.getTime() + 5*60000); 
        startGenericCountdown(deadline, "checkInTimer", "EXPIRED");

    } catch (e) { console.error(e); }
}

function startGenericCountdown(targetDate, elementId, expireText="00:00") {
    if (checkInCountdown) clearInterval(checkInCountdown);
    checkInCountdown = setInterval(() => {
        const now = new Date().getTime();
        const distance = targetDate.getTime() - now;
        const timerEl = document.getElementById(elementId);
        if (distance < 0) timerEl.textContent = expireText; 
        else {
            const m = Math.floor((distance % (1000 * 60 * 60)) / (1000 * 60));
            const s = Math.floor((distance % (1000 * 60)) / 1000);
            timerEl.textContent = `${m}m ${s}s`;
        }
    }, 1000);
}
function stopCheckInCountdown() { if (checkInCountdown) clearInterval(checkInCountdown); }

// 🔒 SECURE CHECK-IN 🔒
async function performCheckIn(roomEmail, eventId, eventDetails) {
    stopCheckInCountdown();
    
    const organizerEmail = eventDetails.organizer?.emailAddress?.address?.toLowerCase() || "";
    const attendees = eventDetails.attendees || []; 
    const allowedEmails = attendees.map(a => a.emailAddress?.address?.toLowerCase());
    allowedEmails.push(organizerEmail);

    let userEmail = "";
    try {
        const loginResp = await myMSALObj.loginPopup({
            scopes: ["User.Read"],
            prompt: "select_account"
        });
        userEmail = loginResp.account.username.toLowerCase();
    } catch (e) {
        console.log("Check-in cancelled:", e);
        checkForActiveMeeting(); 
        return;
    }

    if (!allowedEmails.includes(userEmail)) {
        alert(`⛔ ACCESS DENIED\n\nYou (${userEmail}) are not invited to this meeting.`);
        signOut();
        return;
    }

    try {
        const res = await fetch(`${API_URL}/checkin`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${await getAuthToken()}` 
            },
            body: JSON.stringify({ room_email: roomEmail, event_id: eventId })
        });
        
        if (res.ok) {
            showMeetingMode(eventDetails); 
        } else {
            alert("System Error: Check-in failed.");
            checkForActiveMeeting();
        }
    } catch (e) { alert("Network Error: " + e.message); }
}

function showMeetingMode(event) {
    currentLockedEvent = event;
    const overlay = document.getElementById('meetingInProgressOverlay');
    const start = new Date(event.start.dateTime + 'Z').toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    const end = new Date(event.end.dateTime + 'Z').toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    
    const safeSubject = document.getElementById('bannerSubject').textContent;
    document.getElementById('overlaySubject').textContent = safeSubject; 
    document.getElementById('overlayOrganizer').textContent = `Booked by: ${event.organizer?.emailAddress?.name}`;
    document.getElementById('overlayTime').textContent = `${start} - ${end}`;
    
    overlay.classList.remove('d-none');
    if (sessionTimeout) clearTimeout(sessionTimeout); 
    startMeetingEndTimer(event.end.dateTime);
}

function startMeetingEndTimer(endTimeStr) {
    if (meetingEndInterval) clearInterval(meetingEndInterval);
    const endTime = new Date(endTimeStr + 'Z').getTime();
    meetingEndInterval = setInterval(() => {
        const now = new Date().getTime();
        const distance = endTime - now;
        if (distance < 0) {
            clearInterval(meetingEndInterval);
            document.getElementById('meetingInProgressOverlay').classList.add('d-none');
            checkForActiveMeeting();
        } else {
            const m = Math.floor((distance % (1000 * 60 * 60)) / (1000 * 60));
            const s = Math.floor((distance % (1000 * 60)) / 1000);
            document.getElementById('meetingEndTimer').textContent = `${m}m ${s}s`;
        }
    }, 1000);
}
function stopMeetingEndTimer() { if (meetingEndInterval) clearInterval(meetingEndInterval); }

// 🔒 SECURE END MEETING 🔒
async function secureEndMeeting() {
    if (isAuthInProgress) return;
    
    const organizerEmail = currentLockedEvent.organizer?.emailAddress?.address?.toLowerCase() || "";
    const attendees = currentLockedEvent.attendees || [];
    const allowedEmails = attendees.map(a => a.emailAddress?.address?.toLowerCase());
    allowedEmails.push(organizerEmail);

    const roomIndex = document.getElementById('roomSelect').value;
    const roomEmail = availableRooms[roomIndex].emailAddress;
    const eventId = currentLockedEvent.id;

    isAuthInProgress = true;

    try {
        const loginResp = await myMSALObj.loginPopup({
            scopes: ["User.Read"],
            prompt: "select_account"
        });
        
        const userEmail = loginResp.account.username.toLowerCase();
        
        if (!allowedEmails.includes(userEmail)) {
            alert(`⛔ ACCESS DENIED\n\nYou (${userEmail}) are not authorized to end this meeting.`);
            return;
        }

        const res = await fetch(`${API_URL}/end-meeting`, {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${await getAuthToken()}` 
            },
            body: JSON.stringify({ room_email: roomEmail, event_id: eventId })
        });

        if (res.ok) {
            manuallyUnlockedEventId = eventId;
            document.getElementById('meetingInProgressOverlay').classList.add('d-none');
            currentLockedEvent = null; 
            stopMeetingEndTimer(); 
            checkForActiveMeeting();
        } else {
            alert("Error ending meeting. Please try again.");
        }
        
    } catch (e) { 
        if (e.errorCode !== "user_cancelled") alert("Authentication failed."); 
    } finally {
        isAuthInProgress = false;
    }
}

// ================= BOOKING & TIMELINE =================

// 🆕 NEW: Custom Kiosk Alert Function (Auto-Close)
function showKioskAlert(message, duration, callback) {
    // 1. Create Overlay
    const overlay = document.createElement('div');
    overlay.id = "kiosk-success-overlay";
    overlay.style.cssText = `
        position: fixed; top: 0; left: 0; width: 100%; height: 100%;
        background: rgba(0,0,0,0.85); z-index: 99999;
        display: flex; flex-direction: column; justify-content: center; align-items: center;
        color: white; font-family: sans-serif; text-align: center;
    `;
    
    // 2. Create Content
    const content = document.createElement('div');
    content.innerHTML = `
        <div style="font-size: 4rem; margin-bottom: 20px;">✅</div>
        <h2 style="font-size: 2.5rem;">Booking Confirmed</h2>
        <div style="font-size: 1.5rem; margin-top: 20px; text-align: left; background: rgba(255,255,255,0.1); padding: 20px; border-radius: 10px;">
            ${message}
        </div>
        <button id="kiosk-close-btn" class="btn btn-success btn-lg mt-4" style="font-size: 1.5rem; padding: 10px 40px;">
            OK (Closing in <span id="kiosk-timer">4</span>s)
        </button>
    `;
    
    overlay.appendChild(content);
    document.body.appendChild(overlay);

    // 3. Logic to Close
    let secondsLeft = duration / 1000;
    const timerSpan = document.getElementById('kiosk-timer');
    
    const interval = setInterval(() => {
        secondsLeft--;
        if(timerSpan) timerSpan.innerText = secondsLeft;
    }, 1000);

    const close = () => {
        clearInterval(interval);
        if(document.body.contains(overlay)) document.body.removeChild(overlay);
        if(callback) callback();
    };

    // Auto-Close
    const timeout = setTimeout(close, duration);

    // Manual Close
    document.getElementById('kiosk-close-btn').onclick = () => {
        clearTimeout(timeout);
        close();
    };
}

async function createBooking() {
    if (!username) return alert("Please sign in first.");
    const index = document.getElementById('roomSelect').value;
    if (!index) return alert("Select a room.");
    const roomEmail = availableRooms[index].emailAddress;
    
    const subject = document.getElementById('subject').value;
    const filiale = document.getElementById('filiale').value; 
    const desc = document.getElementById('description').value;
    const startInput = document.getElementById('startTime').value;
    const endInput = document.getElementById('endTime').value;
    const attendeesRaw = document.getElementById('attendees').value;
    let attendeeList = attendeesRaw.trim() ? attendeesRaw.split(',').map(e => e.trim()) : [];

    let accessToken = "";
    try {
        const account = myMSALObj.getAllAccounts()[0];
        const tokenResp = await myMSALObj.acquireTokenSilent({ ...loginRequest, account: account });
        accessToken = tokenResp.accessToken;
    } catch (e) { return alert("Permission denied. Relogin required."); }

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
            // Close the form modal first
            const modalEl = document.getElementById('bookingModal');
            const modal = bootstrap.Modal.getInstance(modalEl);
            if(modal) modal.hide(); else modalEl.classList.remove('show');

            loadAvailability(roomEmail); 
            
            // 🆕 Construct Detailed Message
            const startFmt = new Date(startInput).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
            const endFmt = new Date(endInput).toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
            const msg = `
                <strong>Subject:</strong> ${subject}<br>
                <strong>Unit:</strong> ${filiale}<br>
                <strong>Time:</strong> ${startFmt} - ${endFmt}<br>
                <strong>Invitees:</strong> ${attendeesRaw || "None"}
            `;

            // 🆕 Show Auto-Closing Kiosk Alert
            showKioskAlert(msg, 4000, () => {
                signOut(); // Auto-logout after alert closes
            });

        } else {
            const err = await res.json();
            alert("Error: " + (err.detail || JSON.stringify(err)));
        }
    } catch (e) { alert("Network Error: " + e.message); }
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

function handleRoomChange() { 
    const index = document.getElementById('roomSelect').value; 
    const room = availableRooms[index]; 
    if (room) { 
        lastKnownEventId = "init"; 
        loadAvailability(room.emailAddress); 
        checkForActiveMeeting(); 
    } 
}

async function loadAvailability(email) { 
    if (!email) return; 
    document.getElementById('loadingSpinner').style.display = "inline"; 
    const now = new Date(); 
    const viewStart = new Date(now); 
    viewStart.setHours(viewStart.getHours() - 2); 
    viewStart.setMinutes(0, 0, 0); 
    const viewEnd = new Date(viewStart.getTime() + 12 * 60 * 60 * 1000); 
    
    try { 
        const res = await fetch(`${API_URL}/availability`, { 
            method: 'POST', 
            headers: {'Content-Type': 'application/json'}, 
            body: JSON.stringify({ 
                room_email: email, 
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

function handleBookClick() { 
    if(!username) { signIn(); return; } 
    document.getElementById('displayEmail').value = username; 
    new bootstrap.Modal(document.getElementById('bookingModal')).show(); 
}

function initModalTimes() { 
    const now = new Date(); 
    now.setMinutes(now.getMinutes()-now.getTimezoneOffset()); 
    document.getElementById('startTime').value = now.toISOString().slice(0,16); 
    now.setMinutes(now.getMinutes()+30); 
    document.getElementById('endTime').value = now.toISOString().slice(0,16); 
}

function renderTimeline(data, viewStart, viewEnd) { 
    const timelineContainer = document.getElementById('timeline'); 
    timelineContainer.innerHTML = ''; 
    const totalDurationMs = viewEnd - viewStart; 
    
    const track = document.createElement('div'); 
    track.className = 'timeline-track'; 
    
    const numHours = 12; 
    for (let i = 0; i <= numHours; i++) { 
        let slotTime = new Date(viewStart.getTime() + i * 60 * 60 * 1000); 
        let pct = (i / numHours) * 100; 
        
        const label = document.createElement('div'); 
        label.className = 'timeline-time-label'; 
        
        // 🛠️ THE FIX: Force Zig-Zag using JavaScript
        // If 'i' is an odd number (1, 3, 5...), add the 'stagger-up' class
        if (i % 2 !== 0) {
            label.classList.add('stagger-up');
        }

        label.style.left = `${pct}%`; 
        label.innerText = slotTime.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'}); 
        track.appendChild(label); 
        
        if (i > 0 && i < numHours) { 
            const line = document.createElement('div'); 
            line.className = 'grid-line'; 
            line.style.left = `${pct}%`; 
            track.appendChild(line); 
        } 
    } 
    
    const schedule = (data.value && data.value[0]) ? data.value[0] : null; 
    if (schedule && schedule.scheduleItems) { 
        schedule.scheduleItems.forEach(item => { 
            if (item.status === 'busy' || item.status === 'tentative') { 
                const start = new Date(item.start.dateTime + 'Z'); 
                const end = new Date(item.end.dateTime + 'Z'); 
                const leftPct = ((start - viewStart) / totalDurationMs) * 100; 
                const widthPct = ((end - start) / totalDurationMs) * 100; 
                
                if (leftPct < 100 && (leftPct + widthPct) > 0) { 
                    const block = document.createElement('div'); 
                    block.className = 'event-block'; 
                    block.style.left = `${Math.max(0, leftPct)}%`; 
                    block.style.width = `${Math.min(widthPct, 100 - Math.max(0, leftPct))}%`; 
                    block.innerHTML = '<span class="event-label">Busy</span>'; 
                    
                    block.addEventListener('mouseenter', (e) => showTooltip(e, item));
                    block.addEventListener('mousemove', (e) => moveTooltip(e));
                    block.addEventListener('mouseleave', hideTooltip);
                    track.appendChild(block); 
                } 
            } 
        }); 
    } 
    timelineContainer.appendChild(track); 
}
function showTooltip(e, item) {
    const tooltip = document.getElementById('timelineTooltip');
    const subject = item.subject || "Private Meeting";
    const start = new Date(item.start.dateTime + 'Z').toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    const end = new Date(item.end.dateTime + 'Z').toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    document.getElementById('tooltipSubject').innerText = subject;
    document.getElementById('tooltipTime').innerText = `🕒 ${start} - ${end}`;
    tooltip.style.display = 'block';
    moveTooltip(e);
}

function moveTooltip(e) {
    const tooltip = document.getElementById('timelineTooltip');
    tooltip.style.left = (e.pageX + 15) + 'px';
    tooltip.style.top = (e.pageY + 15) + 'px';
}

function hideTooltip() { document.getElementById('timelineTooltip').style.display = 'none'; }
