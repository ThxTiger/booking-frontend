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
let isAuthInProgress = false; 
let manuallyUnlockedEventId = null;

// ================= INITIALIZATION =================
document.addEventListener("DOMContentLoaded", async () => {
    initModalTimes();
    await fetchRooms();
    setInterval(checkForActiveMeeting, 5000); // Poll every 5 seconds
    
    try {
        await msalInstance.initialize();
        const response = await msalInstance.handleRedirectPromise();
        if (response) handleLoginSuccess(response.account);
    } catch (e) { console.error(e); }
});

async function signIn() { 
    if(isAuthInProgress) return;
    isAuthInProgress = true;
    try { await msalInstance.loginRedirect(loginRequest); } 
    catch (e) { console.error(e); } 
    finally { isAuthInProgress = false; }
}

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

// ================= CORE LOGIC (BUG FIXES HERE) =================

async function checkForActiveMeeting() {
    const index = document.getElementById('roomSelect').value;
    if (!index) return;
    const roomEmail = availableRooms[index].emailAddress;

    try {
        // Fetch the CURRENT or NEXT meeting from backend
        const res = await fetch(`${API_URL}/active-meeting?room_email=${roomEmail}`);
        let event = await res.json();
        
        const banner = document.getElementById('checkInBanner');
        const overlay = document.getElementById('meetingInProgressOverlay');

        // CASE 1: NO MEETING FOUND
        if (!event) {
            banner.style.display = "none";
            overlay.classList.add('d-none');
            stopCheckInCountdown(); stopMeetingEndTimer();
            return;
        }

        const now = new Date();
        const start = new Date(event.start.dateTime + 'Z');
        const end = new Date(event.end.dateTime + 'Z');

        // CASE 2: MEETING ENDED (Double check to be safe)
        if (now >= end) {
            banner.style.display = "none";
            overlay.classList.add('d-none');
            return;
        }

        // --- 🛠️ AUTO-REPAIR LOGIC (THE NAME FIXER) 🛠️ ---
        const organizerName = event.organizer?.emailAddress?.name || "Unknown";
        let cleanSubject = (event.subject || "").trim();
        const cleanOrg = organizerName.trim();

        // Check if Outlook overwrote the subject with the organizer's name
        if (cleanSubject.toLowerCase() === cleanOrg.toLowerCase() || cleanSubject === "") {
            console.log("⚠️ Corrupted Subject Detected. Fixing...");
            
            // Try to find "Filiale :" inside the body preview
            if (event.bodyPreview) {
                const filialeMatch = event.bodyPreview.match(/Filiale\s*:\s*(.*?)(\n|$|Reason)/i);
                if (filialeMatch && filialeMatch[1]) {
                    event.subject = filialeMatch[1].trim(); 
                } else if (event.bodyPreview.length > 5) {
                    // Fallback: use first 40 chars of body
                    event.subject = event.bodyPreview.substring(0, 40) + "...";
                } else {
                    event.subject = "Private Meeting";
                }
            } else {
                event.subject = "Private Meeting";
            }
        }
        
        // Update Variables
        document.getElementById('bannerSubject').textContent = event.subject;
        document.getElementById('bannerOrganizer').textContent = organizerName;

        // --- 🚥 LOGIC BRANCHING: FUTURE vs ACTIVE 🚥 ---

        // CASE 3: FUTURE MEETING (Upcoming)
        // BUG FIX: Strictly check if Now < Start
        if (now < start) {
            // Setup Blue Banner (Info Mode)
            banner.style.display = "block";
            overlay.classList.add('d-none'); // Ensure Red Screen is hidden
            
            // UI: "STARTS IN"
            document.getElementById('bannerStatusTitle').textContent = "📅 Next Meeting";
            const badge = document.getElementById('bannerBadge');
            badge.className = "badge bg-info mb-1";
            badge.textContent = "STARTS IN";
            
            // Hide Check-in Button (Can't check in yet)
            document.getElementById('realCheckInBtn').style.display = "none";
            
            // Start Countdown to START TIME
            startGenericCountdown(start, "checkInTimer", "STARTING...");
            return; 
        }

        // CASE 4: ACTIVE MEETING (Happening Now)
        // (now >= start && now < end)

        // 4A: ALREADY CHECKED IN?
        if (event.categories && event.categories.includes("Checked-In")) {
             banner.style.display = "none"; // Hide Yellow/Blue Banner
             stopCheckInCountdown();
             
             // Show Red Overlay (Do Not Disturb)
             if (overlay.classList.contains('d-none') && event.id !== manuallyUnlockedEventId) {
                 showMeetingMode(event);
             }
             return;
        } 
        
        // 4B: NOT CHECKED IN (Ghost Buster Danger Zone)
        // Show Yellow Banner
        if (event.id !== manuallyUnlockedEventId) manuallyUnlockedEventId = null;

        banner.style.display = "block";
        overlay.classList.add('d-none'); 
        
        // UI: "DEADLINE"
        document.getElementById('bannerStatusTitle').textContent = "⚠️ Meeting Started! Confirm Presence";
        const badge = document.getElementById('bannerBadge');
        badge.className = "badge bg-danger mb-1";
        badge.textContent = "AUTO-CANCEL IN";

        // Show Check-In Button
        const btn = document.getElementById('realCheckInBtn');
        btn.style.display = "inline-block";
        btn.onclick = () => performCheckIn(roomEmail, event.id, event);

        // Timer: 5 Minutes from Start Time
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
        
        if (distance < 0) { 
            timerEl.textContent = expireText; 
        } else {
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

function showMeetingMode(event) {
    currentLockedEvent = event;
    const overlay = document.getElementById('meetingInProgressOverlay');
    const start = new Date(event.start.dateTime + 'Z').toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    const end = new Date(event.end.dateTime + 'Z').toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    
    // UI Update
    const safeSubject = document.getElementById('bannerSubject').textContent;
    document.getElementById('overlaySubject').textContent = safeSubject; // Use the repaired subject
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

async function secureEndMeeting() {
    if (isAuthInProgress) return;
    
    const organizerEmail = currentLockedEvent.organizer.emailAddress.address.toLowerCase();
    const roomIndex = document.getElementById('roomSelect').value;
    const roomEmail = availableRooms[roomIndex].emailAddress;
    const eventId = currentLockedEvent.id;

    isAuthInProgress = true;

    try {
        const loginResp = await msalInstance.loginPopup({ 
            scopes: ["User.Read"], 
            prompt: "login" 
        });
        
        const verifiedEmail = loginResp.account.username.toLowerCase();
        
        if (verifiedEmail === organizerEmail) {
            const res = await fetch(`${API_URL}/end-meeting`, {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
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
        } else { alert(`⛔ ACCESS DENIED\nVerified: ${verifiedEmail}`); }
        
    } catch (e) { 
        if (e.errorCode !== "user_cancelled") alert("Authentication failed."); 
    } finally {
        isAuthInProgress = false;
    }
}

// ================= BOOKING & TIMELINE (EXISTING CODE) =================

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
        const account = msalInstance.getAllAccounts()[0];
        const tokenResp = await msalInstance.acquireTokenSilent({ ...loginRequest, account: account });
        accessToken = tokenResp.accessToken;
    } catch (e) {
        try {
            const tokenResp = await msalInstance.acquireTokenPopup(loginRequest);
            accessToken = tokenResp.accessToken;
            handleLoginSuccess(tokenResp.account);
        } catch (err) { return alert("Permission denied."); }
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
            const startFmt = new Date(startInput).toLocaleString();
            const endFmt = new Date(endInput).toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'});
            
            alert(
                `✅ BOOKING CONFIRMED\n\n` +
                `📅 Time: ${startFmt} - ${endFmt}\n` +
                `🏢 Unit: ${filiale}\n` +
                `📝 Subject: ${subject}\n` +
                `💡 Reason: ${desc}\n` +
                `👥 Invitees: ${attendeesRaw || "None"}\n\n` +
                `The meeting has been added to your calendar.`
            );

            const modalEl = document.getElementById('bookingModal');
            const modal = bootstrap.Modal.getInstance(modalEl);
            if(modal) modal.hide(); else modalEl.classList.remove('show');
            loadAvailability(roomEmail); 
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
            if (item.status === 'busy') { 
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
