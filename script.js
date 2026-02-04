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
    setInterval(checkForActiveMeeting, 5000); // Check for meeting updates every 5s
    
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
    
    // Auto-Logout Timer (2 Minutes)
    if(sessionTimeout) clearTimeout(sessionTimeout);
    sessionTimeout = setTimeout(() => { 
        alert("Session Expired: You have been logged out of Booking Mode."); 
        signOut(); 
    }, 120000);
}

// ================= 📅 TIMELINE RENDERER (GANTT STYLE) =================
function renderTimeline(data, viewStart, viewEnd) {
    const timelineContainer = document.getElementById('timeline');
    timelineContainer.innerHTML = ''; // Clear existing timeline

    const totalDurationMs = viewEnd - viewStart; 
    
    // Create the main track div
    const track = document.createElement('div');
    track.className = 'timeline-track';
    
    // Add Time Labels & Grid Lines (Every 1 Hour)
    const numHours = 12;
    for (let i = 0; i <= numHours; i++) {
        let slotTime = new Date(viewStart.getTime() + i * 60 * 60 * 1000);
        let pct = (i / numHours) * 100;

        // Label
        const label = document.createElement('div');
        label.className = 'timeline-time-label';
        label.style.left = `${pct}%`;
        label.innerText = slotTime.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
        track.appendChild(label);

        // Grid Line
        if (i > 0 && i < numHours) {
            const line = document.createElement('div');
            line.className = 'grid-line';
            line.style.left = `${pct}%`;
            track.appendChild(line);
        }
    }

    // Add Meeting Blocks
    const schedule = (data.value && data.value[0]) ? data.value[0] : null; 
    if (schedule && schedule.scheduleItems) { 
        schedule.scheduleItems.forEach(item => { 
            if (item.status === 'busy') { 
                const start = new Date(item.start.dateTime + 'Z'); 
                const end = new Date(item.end.dateTime + 'Z'); 
                
                // Calculate Position & Width
                const leftPct = ((start - viewStart) / totalDurationMs) * 100; 
                const widthPct = ((end - start) / totalDurationMs) * 100; 
                
                // Only render if visible in this 12h window
                if (leftPct < 100 && (leftPct + widthPct) > 0) { 
                    const block = document.createElement('div');
                    block.className = 'event-block';
                    // Clamp visual range so it doesn't overflow
                    block.style.left = `${Math.max(0, leftPct)}%`;
                    block.style.width = `${Math.min(widthPct, 100 - Math.max(0, leftPct))}%`;
                    
                    // Label inside block
                    block.innerHTML = '<span class="event-label">Busy</span>';

                    // 🛑 ADD TOOLTIP EVENTS
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

// Tooltip Helpers
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
    // Position tooltip 15px to the right/bottom of mouse cursor
    tooltip.style.left = (e.pageX + 15) + 'px';
    tooltip.style.top = (e.pageY + 15) + 'px';
}

function hideTooltip() {
    document.getElementById('timelineTooltip').style.display = 'none';
}


// ================= 🔍 CHECK-IN & BANNER LOGIC =================
async function checkForActiveMeeting() {
    const index = document.getElementById('roomSelect').value;
    if (!index) return; // Exit if no room selected
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

            // 🔴 PREVENT FLICKER: Only update text if changed
            const newSubject = event.subject;
            const newOrganizer = event.organizer?.emailAddress?.name || "Unknown";

            if (document.getElementById('bannerSubject').innerText !== newSubject) {
                document.getElementById('bannerSubject').textContent = newSubject;
                document.getElementById('bannerOrganizer').textContent = newOrganizer;
            }

            // 2. Is meeting ALREADY CHECKED IN?
            if (event.categories && event.categories.includes("Checked-In")) {
                 banner.style.display = "none";
                 stopCheckInCountdown();
                 // Show Red Screen if not already there
                 if (overlay.classList.contains('d-none')) showMeetingMode(event);
                 return;
            } 
            
            // 3. SHOW BANNER (Active OR Upcoming)
            banner.style.display = "block";
            overlay.classList.add('d-none');
            
            const btn = document.getElementById('realCheckInBtn');
            btn.onclick = () => performCheckIn(roomEmail, event.id, event);

            // Calculate "Starts In" vs "Deadline"
            const minsUntil = Math.floor((start - now) / 60000);

            if (minsUntil > 15) {
                // FUTURE: Show "Next Meeting"
                document.getElementById('bannerStatusTitle').textContent = "📅 Next Meeting";
                const badge = document.getElementById('bannerBadge');
                badge.className = "badge bg-info mb-1";
                badge.textContent = "STARTS IN";
                startGenericCountdown(start, "checkInTimer");
            } else {
                // ACTIVE: Show "Check-In Required"
                document.getElementById('bannerStatusTitle').textContent = "⚠️ Check-In Required";
                const badge = document.getElementById('bannerBadge');
                badge.className = "badge bg-danger mb-1";
                badge.textContent = "DEADLINE";
                const deadline = new Date(start.getTime() + 5*60000); // Start + 5m
                startGenericCountdown(deadline, "checkInTimer", "EXPIRED");
            }
        } else {
            // No Active Meeting
            banner.style.display = "none";
            overlay.classList.add('d-none');
            stopCheckInCountdown(); stopMeetingEndTimer();
        }
    } catch (e) { console.error(e); }
}

// Timer Helpers
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

// ================= ⛔ RED SCREEN LOGIC =================
function showMeetingMode(event) {
    currentLockedEvent = event;
    const overlay = document.getElementById('meetingInProgressOverlay');
    const start = new Date(event.start.dateTime + 'Z').toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    const end = new Date(event.end.dateTime + 'Z').toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
    
    document.getElementById('overlaySubject').textContent = event.subject;
    document.getElementById('overlayOrganizer').textContent = `Booked by: ${event.organizer?.emailAddress?.name}`;
    document.getElementById('overlayTime').textContent = `${start} - ${end}`;
    overlay.classList.remove('d-
