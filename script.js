// ═══════════════════════════════════════════
//  CONFIGURATION
// ═══════════════════════════════════════════
const API_URL = "https://booking-a-room-poc.onrender.com";

const msalConfig = {
    auth: {
        clientId: "0f759785-1ba8-449d-ba6f-9ba5e8f479d8",
        authority: "https://login.microsoftonline.com/2b2369a3-0061-401b-97d9-c8c8d92b76f6",
        redirectUri: window.location.origin,
    },
    // CRITICAL: Must be localStorage so the Kiosk popup can share the token with the main screen
    cache: { cacheLocation: "localStorage" } 
};

// prompt: "select_account" guarantees the next employee isn't auto-logged in as you
const loginRequest = { 
    scopes: ["User.Read", "Calendars.ReadWrite"],
    prompt: "select_account" 
};

const myMSALObj = new msal.PublicClientApplication(msalConfig);

// ── State ──
let username = "";
let availableRooms = [];
let currentLockedEvent = null;
let checkInInterval = null;
let meetingEndInterval = null;
let clockInterval = null;
let sessionTimeout = null;
let isAuthInProgress = false;
let manuallyUnlockedEventId = null;
let lastKnownEventId = "init";
let currentAppState = "available"; // "available" | "pending" | "occupied"

// ═══════════════════════════════════════════
//  STATE MACHINE
// ═══════════════════════════════════════════
function setAppState(state) {
    if (currentAppState === state) return;
    currentAppState = state;

    document.body.classList.remove("state-available", "state-pending", "state-occupied");
    document.body.classList.add(`state-${state}`);

    const labels = { available: "Available", pending: "Pending Check-In", occupied: "In Use" };
    const el = document.getElementById("statusLabel");
    if (el) el.textContent = labels[state] || "Available";
}

function showView(viewId) {
    ["viewAvailable", "viewPending", "viewFuture"].forEach(id => {
        const el = document.getElementById(id);
        if (el) { el.classList.remove("active"); }
    });
    const target = document.getElementById(viewId);
    if (target) target.classList.add("active");
}

// ═══════════════════════════════════════════
//  INITIALIZATION
// ═══════════════════════════════════════════
document.addEventListener("DOMContentLoaded", async () => {
    initModalTimes();
    startClock();
    await fetchRooms();

    // Heartbeat every 5s
    setInterval(checkForActiveMeeting, 5000);
    // Timeline refresh every 60s
    setInterval(() => {
        const idx = document.getElementById("roomSelect").value;
        if (idx !== "") loadAvailability(availableRooms[idx].emailAddress);
    }, 60000);

    try {
        await myMSALObj.initialize();
        const accounts = myMSALObj.getAllAccounts();
        if (accounts.length > 0) {
            handleLoginSuccess(accounts[0]);
        }
    } catch (e) { 
        console.error("Auth init error: ", e); 
    }
});

// ═══════════════════════════════════════════
//  CLOCK
// ═══════════════════════════════════════════
function startClock() {
    function tick() {
        const now = new Date();
        const t = now.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
        const d = now.toLocaleDateString([], { weekday: "long", month: "long", day: "numeric" });
        const lc = document.getElementById("liveClock");
        const oc = document.getElementById("occClock");
        const od = document.getElementById("occDate");
        if (lc) lc.textContent = t;
        if (oc) oc.textContent = t;
        if (od) od.textContent = d;
    }
    tick();
    clockInterval = setInterval(tick, 1000);
}

// ═══════════════════════════════════════════
//  AUTH (POPUP FLOW WITH KIOSK POLLER)
// ═══════════════════════════════════════════
async function signIn() {
    if (isAuthInProgress) return;
    isAuthInProgress = true;
    let authCompleted = false;

    // KIOSK POLLER: Actively watch storage because the Kiosk WebView blocks popup messages
    const kioskPoller = setInterval(() => {
        const accounts = myMSALObj.getAllAccounts();
        if (accounts.length > 0 && !authCompleted) {
            authCompleted = true;
            clearInterval(kioskPoller);
            handleLoginSuccess(accounts[0]);
            isAuthInProgress = false;
        }
    }, 1000);

    try { 
        const response = await myMSALObj.loginPopup(loginRequest);
        if (!authCompleted) {
            authCompleted = true;
            clearInterval(kioskPoller);
            handleLoginSuccess(response.account);
        }
    } catch (e) { 
        clearInterval(kioskPoller);
        console.error(e); 
    } finally { 
        isAuthInProgress = false; 
    }
}

function signOut() {
    username = "";
    const badge = document.getElementById("userBadge");
    const loginBtn = document.getElementById("loginBtn");
    if (badge) badge.style.display = "none";
    if (loginBtn) loginBtn.style.display = "inline-block";
    if (sessionTimeout) clearTimeout(sessionTimeout);
    
    // Silently wipe the tablet's local memory instead of triggering the MS Popup
    localStorage.clear();
    sessionStorage.clear();
    
    stopCountdowns();
    checkForActiveMeeting();
}

function handleLoginSuccess(acc) {
    username = acc.username;
    const welcome = document.getElementById("userWelcome");
    const badge = document.getElementById("userBadge");
    const loginBtn = document.getElementById("loginBtn");
    if (welcome) welcome.textContent = username;
    if (badge) badge.style.display = "flex";
    if (loginBtn) loginBtn.style.display = "none";

    if (sessionTimeout) clearTimeout(sessionTimeout);
    sessionTimeout = setTimeout(() => signOut(), 120000); // 2min auto-logout
}

async function getAuthToken() {
    try {
        const account = myMSALObj.getAllAccounts()[0];
        if (!account) return null;
        const r = await myMSALObj.acquireTokenSilent({ scopes: ["User.Read"], account });
        return r.accessToken;
    } catch { return null; }
}

// ═══════════════════════════════════════════
//  AUTH GATE (shown before booking if not logged in)
// ═══════════════════════════════════════════
function handleBookClick() {
    if (!username) {
        openAuthGate();
    } else {
        openBookingModal();
    }
}

function openAuthGate() {
    document.getElementById("authGateOverlay").classList.remove("hidden");
}

function closeAuthGate() {
    document.getElementById("authGateOverlay").classList.add("hidden");
}

async function triggerSignInThenBook() {
    closeAuthGate();
    if (isAuthInProgress) return;
    isAuthInProgress = true;
    let authCompleted = false;

    // KIOSK POLLER: Ensures the booking modal opens even if the popup promise hangs
    const kioskPoller = setInterval(() => {
        const accounts = myMSALObj.getAllAccounts();
        if (accounts.length > 0 && !authCompleted) {
            authCompleted = true;
            clearInterval(kioskPoller);
            handleLoginSuccess(accounts[0]);
            openBookingModal();
            isAuthInProgress = false;
        }
    }, 1000);

    try {
        const response = await myMSALObj.loginPopup(loginRequest);
        if (!authCompleted) {
            authCompleted = true;
            clearInterval(kioskPoller);
            handleLoginSuccess(response.account);
            openBookingModal();
        }
    } catch (e) {
        clearInterval(kioskPoller);
        console.error("Sign in failed before booking: ", e);
    } finally {
        isAuthInProgress = false;
    }
}

// ═══════════════════════════════════════════
//  BOOKING MODAL
// ═══════════════════════════════════════════
function openBookingModal() {
    document.getElementById("displayEmail").value = username;
    initModalTimes();
    document.getElementById("bookingOverlay").classList.remove("hidden");
    setTimeout(refreshBookingTimeline, 100);
}

function closeBookingModal() {
    document.getElementById("bookingOverlay").classList.add("hidden");
}

// ═══════════════════════════════════════════
//  AGENDA MODAL
// ═══════════════════════════════════════════
async function openAgenda() {
    document.getElementById("agendaOverlay").classList.remove("hidden");
    const content = document.getElementById("agendaContent");
    const idx = document.getElementById("roomSelect").value;
    if (idx === "") {
        content.innerHTML = `<div class="occ-agenda-empty">Please select a room first.</div>`;
        return;
    }
    const roomEmail = availableRooms[idx].emailAddress;
    content.innerHTML = `<div class="occ-agenda-empty">Loading…</div>`;

    try {
        const now = new Date();
        const dayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
        const dayEnd   = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);
        const res = await fetch(`${API_URL}/availability`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                room_email: roomEmail,
                start_time: dayStart.toISOString(),
                end_time: dayEnd.toISOString(),
                time_zone: "UTC"
            })
        });
        const data = await res.json();
        const items = data?.value?.[0]?.scheduleItems || [];
        const busy = items.filter(i => i.status === "busy");

        if (busy.length === 0) {
            content.innerHTML = `<div class="occ-agenda-empty">No meetings scheduled today.</div>`;
            return;
        }

        content.innerHTML = busy.map(item => {
            const itemStart = new Date(item.start.dateTime + "Z");
            const itemEnd   = new Date(item.end.dateTime + "Z");
            const s = itemStart.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            const e = itemEnd.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            const isNow  = now >= itemStart && now < itemEnd;
            const isPast = now >= itemEnd;
            const badge  = isNow  ? `<span class="agenda-badge now">NOW</span>` :
                           isPast ? `<span class="agenda-badge past">DONE</span>` : "";
            return `
                <div class="agenda-modal-item${isPast ? " past" : isNow ? " active-now" : ""}">
                    <div class="agenda-modal-time">${s} – ${e}</div>
                    <div style="flex:1">
                        <div class="agenda-modal-subject">${item.subject || "Meeting"} ${badge}</div>
                    </div>
                </div>`;
        }).join("");
    } catch (e) {
        content.innerHTML = `<div class="occ-agenda-empty">Failed to load.</div>`;
    }
}

function closeAgenda() {
    document.getElementById("agendaOverlay").classList.add("hidden");
}

// ═══════════════════════════════════════════
//  HEARTBEAT — ACTIVE MEETING CHECK
// ═══════════════════════════════════════════
async function checkForActiveMeeting() {
    const idx = document.getElementById("roomSelect").value;
    if (idx === "") return;
    const roomEmail = availableRooms[idx].emailAddress;

    try {
        const token = await getAuthToken();
        const headers = { "Content-Type": "application/json" };
        if (token) headers["Authorization"] = `Bearer ${token}`;

        const res = await fetch(`${API_URL}/active-meeting?room_email=${roomEmail}`, { headers });
        if (res.status === 401) return;
        const event = await res.json();

        const cid = event ? event.id : "free";
        if (lastKnownEventId !== "init" && lastKnownEventId !== cid) {
            loadAvailability(roomEmail);
        }
        lastKnownEventId = cid;

        const occupied = document.getElementById("occupiedScreen");

        // NO MEETING
        if (!event) {
            setAppState("available");
            showOccupied(false);
            showView("viewAvailable");
            stopCountdowns();
            updateNextMeetingPreview(null);
            return;
        }

        const now = new Date();
        const start = new Date(event.start.dateTime + "Z");
        const end = new Date(event.end.dateTime + "Z");

        if (now >= end) {
            setAppState("available");
            showOccupied(false);
            showView("viewAvailable");
            return;
        }

        let displaySubject = event.subject || "Meeting";
        let displayOrg = event.organizer?.emailAddress?.name || "Unknown";

        if (event.subject === "Busy" && !occupied.classList.contains("hidden")) {
            const existing = document.getElementById("occSubject").textContent;
            if (existing && existing !== "—") displaySubject = existing;
        }

        if (displaySubject === displayOrg || !displaySubject) {
            displaySubject = "Private Meeting";
        }

        const startFmt = start.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
        const endFmt = end.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });

        // FUTURE
        if (now < start) {
            setAppState("available");
            showOccupied(false);
            showView("viewFuture");
            document.getElementById("futureSubject").textContent = displaySubject;
            document.getElementById("futureTime").textContent = `${startFmt} – ${endFmt}`;
            startCountdown(start, "futureTimer", "STARTING…");
            updateNextMeetingPreview({ subject: displaySubject, startFmt, endFmt });
            return;
        }

        // ACTIVE + CHECKED IN (Timer freeze fix applied here)
        if (event.categories?.includes("Checked-In")) {
            setAppState("occupied");
            
            // FIX: Only kill the 5-min warning timer, keep the meeting end timer running!
            if (checkInInterval) { clearInterval(checkInInterval); checkInInterval = null; }
            
            if (occupied.classList.contains("hidden") && event.id !== manuallyUnlockedEventId) {
                showMeetingMode(event, displaySubject, displayOrg, startFmt, endFmt);
            }
            return;
        }

        // ACTIVE + PENDING
        if (event.id !== manuallyUnlockedEventId) manuallyUnlockedEventId = null;
        setAppState("pending");
        showOccupied(false);
        showView("viewPending");

        document.getElementById("pendingSubject").textContent = displaySubject;
        document.getElementById("pendingTime").textContent = `${startFmt} – ${endFmt}`;
        document.getElementById("pendingOrganizer").textContent = `Organized by ${displayOrg}`;

        const deadline = new Date(start.getTime() + 5 * 60000);
        startCountdown(deadline, "checkInTimer", "EXPIRED");

        const btn = document.getElementById("realCheckInBtn");
        btn.onclick = () => performCheckIn(roomEmail, event.id, event);

    } catch (e) { console.error(e); }
}

function showOccupied(show) {
    const occ = document.getElementById("occupiedScreen");
    const main = document.getElementById("mainScreen");
    if (show) {
        occ.classList.remove("hidden");
        main.classList.add("hidden");
    } else {
        occ.classList.add("hidden");
        main.classList.remove("hidden");
    }
}

function updateNextMeetingPreview(data) {
    const preview = document.getElementById("nextMeetingPreview");
    if (!data || !preview) { if (preview) preview.style.display = "none"; return; }
    preview.style.display = "block";
    document.getElementById("nextSubject").textContent = data.subject;
    document.getElementById("nextTime").textContent = `${data.startFmt} – ${data.endFmt}`;
}

// ═══════════════════════════════════════════
//  COUNTDOWNS
// ═══════════════════════════════════════════
function startCountdown(targetDate, elementId, expireText) {
    if (checkInInterval) clearInterval(checkInInterval);
    checkInInterval = setInterval(() => {
        const dist = targetDate - new Date();
        const el = document.getElementById(elementId);
        if (!el) return;
        if (dist <= 0) { el.textContent = expireText; return; }
        const m = Math.floor(dist / 60000);
        const s = Math.floor((dist % 60000) / 1000);
        el.textContent = `${m}:${String(s).padStart(2, "0")}`;
    }, 1000);
}

function stopCountdowns() {
    if (checkInInterval) { clearInterval(checkInInterval); checkInInterval = null; }
    if (meetingEndInterval) { clearInterval(meetingEndInterval); meetingEndInterval = null; }
}

// ═══════════════════════════════════════════
//  CHECK-IN (No Auth)
// ═══════════════════════════════════════════
async function performCheckIn(roomEmail, eventId, eventDetails) {
    if (checkInInterval) clearInterval(checkInInterval);

    try {
        const res = await fetch(`${API_URL}/checkin`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ room_email: roomEmail, event_id: eventId })
        });
        if (res.ok) {
            const startFmt = new Date(eventDetails.start.dateTime + "Z").toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            const endFmt = new Date(eventDetails.end.dateTime + "Z").toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            showMeetingMode(eventDetails, document.getElementById("pendingSubject").textContent,
                eventDetails.organizer?.emailAddress?.name, startFmt, endFmt);
        } else {
            showToast("Check-in failed. Try again.", true);
            checkForActiveMeeting();
        }
    } catch (e) { showToast("Network error.", true); }
}

// ═══════════════════════════════════════════
//  MEETING MODE (Red Screen)
// ═══════════════════════════════════════════
function showMeetingMode(event, subject, organizer, startFmt, endFmt) {
    currentLockedEvent = event;
    setAppState("occupied");
    showOccupied(true);

    document.getElementById("occSubject").textContent = subject || "Meeting";
    document.getElementById("occTime").textContent = `${startFmt} – ${endFmt}`;
    document.getElementById("occOrganizer").textContent = `Organized by ${organizer || "Unknown"}`;

    startMeetingEndTimer(event.end.dateTime);
    updateEndsIn(new Date(event.end.dateTime + "Z"));
    loadOccupiedAgenda(availableRooms[document.getElementById("roomSelect").value]?.emailAddress, event.end.dateTime);

    if (sessionTimeout) clearTimeout(sessionTimeout); 
}

function updateEndsIn(endDate) {
    const mins = Math.max(0, Math.round((endDate - new Date()) / 60000));
    const el = document.getElementById("occEndsIn");
    if (el) el.textContent = mins > 0 ? `Ends in ${mins} min` : "Ending now";
}

function startMeetingEndTimer(endTimeStr) {
    if (meetingEndInterval) clearInterval(meetingEndInterval);
    const endTime = new Date(endTimeStr + "Z").getTime();
    meetingEndInterval = setInterval(() => {
        const dist = endTime - Date.now();
        if (dist <= 0) {
            clearInterval(meetingEndInterval);
            showOccupied(false);
            setAppState("available");
            showView("viewAvailable");
            checkForActiveMeeting();
        } else {
            const m = Math.floor(dist / 60000);
            const s = Math.floor((dist % 60000) / 1000);
            const el = document.getElementById("meetingEndTimer");
            if (el) el.textContent = `${m}m ${String(s).padStart(2, "0")}s`;
            updateEndsIn(new Date(endTimeStr + "Z"));
        }
    }, 1000);
}

async function loadOccupiedAgenda(roomEmail, currentMeetingEndStr) {
    if (!roomEmail) return;
    const list = document.getElementById("occAgenda");
    if (!list) return;

    try {
        const windowStart = currentMeetingEndStr ? new Date(currentMeetingEndStr + "Z") : new Date();
        const windowEnd = new Date(windowStart.getFullYear(), windowStart.getMonth(), windowStart.getDate(), 23, 59);

        const res = await fetch(`${API_URL}/availability`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                room_email: roomEmail,
                start_time: windowStart.toISOString(),
                end_time: windowEnd.toISOString(),
                time_zone: "UTC"
            })
        });
        const data = await res.json();
        const upcoming = (data?.value?.[0]?.scheduleItems || []).filter(i => i.status === "busy");

        if (upcoming.length === 0) {
            list.innerHTML = `<div class="occ-agenda-empty">No more meetings today.</div>`;
            return;
        }
        list.innerHTML = upcoming.map(item => {
            const s = new Date(item.start.dateTime + "Z").toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            const e = new Date(item.end.dateTime + "Z").toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            return `
                <div class="occ-agenda-item">
                    <div class="occ-agenda-item-time">${s} – ${e}</div>
                    <div class="occ-agenda-item-subj">${item.subject || "Meeting"}</div>
                </div>`;
        }).join("");
    } catch { }
}

// ═══════════════════════════════════════════
//  +15 MIN EXTENSION (No Auth Required)
// ═══════════════════════════════════════════
async function extendMeeting(minutes) {
    if (!currentLockedEvent) return;

    const roomIdx = document.getElementById("roomSelect").value;
    const roomEmail = availableRooms[roomIdx].emailAddress;
    const currentEnd = new Date(currentLockedEvent.end.dateTime + "Z");
    const newEnd = new Date(currentEnd.getTime() + minutes * 60000);

    try {
        const res = await fetch(`${API_URL}/availability`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ room_email: roomEmail, start_time: currentEnd.toISOString(), end_time: newEnd.toISOString(), time_zone: "UTC" })
        });
        const data = await res.json();
        const isBusy = (data?.value?.[0]?.scheduleItems || []).some(i => i.status === "busy");
        if (isBusy) {
            showToast("⛔ Cannot extend — another meeting follows immediately.", true);
            return;
        }
    } catch { }

    try {
        const res = await fetch(`${API_URL}/extend-meeting`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ room_email: roomEmail, event_id: currentLockedEvent.id, extend_minutes: minutes })
        });

        if (res.ok) {
            currentLockedEvent.end.dateTime = newEnd.toISOString().replace("Z", "");
            startMeetingEndTimer(currentLockedEvent.end.dateTime);

            const startFmt = new Date(currentLockedEvent.start.dateTime + "Z").toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            const newEndFmt = newEnd.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            document.getElementById("occTime").textContent = `${startFmt} – ${newEndFmt}`;

            showToast(`✅ Extended by ${minutes} min — now ends at ${newEndFmt}`);
            loadAvailability(roomEmail);
        } else {
            const err = await res.json().catch(() => ({}));
            showToast(err.detail || "Extension failed.", true);
        }
    } catch (e) { showToast("Network error.", true); }
}

// ═══════════════════════════════════════════
//  SECURE END MEETING (Popup Flow with Poller)
// ═══════════════════════════════════════════
async function secureEndMeeting() {
    if (isAuthInProgress || !currentLockedEvent) return;

    const roomIdx = document.getElementById("roomSelect").value;
    const roomEmail = availableRooms[roomIdx].emailAddress;
    
    const organizerEmail = currentLockedEvent.organizer?.emailAddress?.address?.toLowerCase() || "";
    const attendees = currentLockedEvent.attendees || [];
    const allowed = [...attendees.map(a => a.emailAddress?.address?.toLowerCase()), organizerEmail];

    isAuthInProgress = true;
    let authCompleted = false;

    // KIOSK POLLER
    const kioskPoller = setInterval(async () => {
        const accounts = myMSALObj.getAllAccounts();
        if (accounts.length > 0 && !authCompleted) {
            authCompleted = true;
            clearInterval(kioskPoller);
            await processSecureEnd(accounts[0].username, allowed, roomEmail, accounts[0]);
        }
    }, 1000);

    try {
        const response = await myMSALObj.loginPopup({ scopes: ["User.Read"], prompt: "select_account" });
        if (!authCompleted) {
            authCompleted = true;
            clearInterval(kioskPoller);
            await processSecureEnd(response.account.username, allowed, roomEmail, response.account);
        }
    } catch (e) {
        clearInterval(kioskPoller);
        if (e.errorCode !== "user_cancelled") showToast("Authentication failed.", true);
        isAuthInProgress = false;
    }
}

async function processSecureEnd(username, allowed, roomEmail, accountObj) {
    try {
        const userEmail = username.toLowerCase();
        if (!allowed.includes(userEmail)) {
            showToast(`⛔ Access denied — you are not authorized to end this meeting.`, true);
            localStorage.clear();
            sessionStorage.clear();
            return;
        }

        const token = await getAuthToken();
        const res = await fetch(`${API_URL}/end-meeting`, {
            method: "POST",
            headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
            body: JSON.stringify({ room_email: roomEmail, event_id: currentLockedEvent.id })
        });

        if (res.ok) {
            manuallyUnlockedEventId = currentLockedEvent.id;
            currentLockedEvent = null;
            stopCountdowns();
            showOccupied(false);
            setAppState("available");
            showView("viewAvailable");
            loadAvailability(roomEmail);
            showToast("✅ Meeting ended successfully.");
        } else {
            showToast("Failed to end meeting.", true);
        }
        
        localStorage.clear();
        sessionStorage.clear();
    } catch (e) {
        showToast("Network error.", true);
        console.error(e);
    } finally {
        isAuthInProgress = false;
    }
}

// ═══════════════════════════════════════════
//  BOOKING
// ═══════════════════════════════════════════
async function createBooking() {
    if (!username) { openAuthGate(); return; }
    const idx = document.getElementById("roomSelect").value;
    if (idx === "") { showToast("Please select a room first.", true); return; }

    const roomEmail = availableRooms[idx].emailAddress;
    const subject = document.getElementById("subject").value.trim();
    const filiale = document.getElementById("filiale").value.trim();
    const desc = document.getElementById("description").value.trim();
    const startVal = document.getElementById("startTime").value;
    const endVal = document.getElementById("endTime").value;
    const attendeesRaw = document.getElementById("attendees").value;
    const attendeeList = attendeesRaw.trim() ? attendeesRaw.split(",").map(e => e.trim()).filter(Boolean) : [];

    if (!subject || !filiale || !startVal || !endVal) {
        showToast("Please fill in all required fields.", true); return;
    }

    let accessToken = "";
    try {
        const account = myMSALObj.getAllAccounts()[0];
        const r = await myMSALObj.acquireTokenSilent({ ...loginRequest, account });
        accessToken = r.accessToken;
    } catch { showToast("Session expired. Please sign in again.", true); return; }

    try {
        const res = await fetch(`${API_URL}/book`, {
            method: "POST",
            headers: { "Content-Type": "application/json", "Authorization": `Bearer ${accessToken}` },
            body: JSON.stringify({
                subject, room_email: roomEmail,
                start_time: new Date(startVal).toISOString(),
                end_time: new Date(endVal).toISOString(),
                organizer_email: username, attendees: attendeeList, filiale, description: desc
            })
        });

        if (res.ok) {
            closeBookingModal();
            const startFmt = new Date(startVal).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            const endFmt = new Date(endVal).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            showBookingSuccess(subject, filiale, `${startFmt} – ${endFmt}`, attendeesRaw);
            loadAvailability(roomEmail);
        } else {
            const err = await res.json().catch(() => ({}));
            showToast(err.detail || "Booking failed.", true);
        }
    } catch (e) { showToast("Network error: " + e.message, true); }
}

function showBookingSuccess(subject, filiale, timeRange, invitees) {
    const overlay = document.createElement("div");
    overlay.style.cssText = `
        position:fixed;inset:0;z-index:9999;
        background:rgba(5,20,10,0.97);
        display:flex;flex-direction:column;justify-content:center;align-items:center;
        font-family:'Sora',sans-serif;text-align:center;padding:40px;
        animation:fadeIn .3s ease;
    `;
    overlay.innerHTML = `
        <div style="font-size:3.5rem;margin-bottom:20px;">✅</div>
        <div style="font-size:1.8rem;font-weight:800;color:#fff;margin-bottom:6px;">Booking Confirmed</div>
        <div style="font-size:0.9rem;color:rgba(255,255,255,.4);margin-bottom:32px;">Added to your Outlook calendar.</div>
        <div style="background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);border-radius:16px;padding:24px 36px;text-align:left;min-width:280px;line-height:2.2;font-size:.9rem;color:rgba(255,255,255,.75);">
            <div><strong style="color:#22c46e;">Subject</strong>  ${subject}</div>
            <div><strong style="color:#22c46e;">Unit</strong>     ${filiale}</div>
            <div><strong style="color:#22c46e;">Time</strong>     ${timeRange}</div>
            <div><strong style="color:#22c46e;">Invitees</strong> ${invitees || "None"}</div>
        </div>
        <button id="successClose" style="margin-top:28px;padding:12px 40px;border-radius:999px;background:#22c46e;color:#05200e;border:none;font-family:'Sora',sans-serif;font-weight:700;font-size:0.95rem;cursor:pointer;">
            OK · Closing in <span id="successCountdown">5</span>s
        </button>
    `;
    document.body.appendChild(overlay);

    let n = 5;
    const iv = setInterval(() => {
        n--;
        const el = document.getElementById("successCountdown");
        if (el) el.textContent = n;
        if (n <= 0) { clearInterval(iv); close(); }
    }, 1000);

    const close = () => {
        if (document.body.contains(overlay)) document.body.removeChild(overlay);
        signOut(); 
    };

    document.getElementById("successClose").onclick = () => { clearInterval(iv); close(); };
}

// ═══════════════════════════════════════════
//  ROOMS & TIMELINE
// ═══════════════════════════════════════════
async function fetchRooms() {
    try {
        const res = await fetch(`${API_URL}/rooms`);
        const data = await res.json();
        if (data.value) {
            availableRooms = data.value;
            const select = document.getElementById("roomSelect");
            select.innerHTML = `<option value="" disabled selected>Select a room…</option>`;
            availableRooms.forEach((r, i) => {
                const opt = document.createElement("option");
                opt.value = i;
                opt.textContent = `${r.displayName}  [${r.department} · ${r.floor}]`;
                select.appendChild(opt);
            });
        }
    } catch (e) { console.error(e); }
}

function handleRoomChange() {
    const idx = document.getElementById("roomSelect").value;
    if (idx === "") return;
    const room = availableRooms[idx];

    const floorEl = document.getElementById("roomFloor");
    const deptEl = document.getElementById("roomDept");
    const capEl = document.getElementById("roomCapacity");
    const locEl = document.getElementById("roomLocation");
    if (floorEl) floorEl.querySelector("span").textContent = room.floor || "—";
    if (deptEl) deptEl.querySelector("span").textContent = room.department || "—";
    if (capEl) capEl.querySelector("span").textContent = (room.capacity || 8) + " persons";
    if (locEl) locEl.querySelector("span").textContent = room.location || "Casablanca HQ";

    lastKnownEventId = "init";
    loadAvailability(room.emailAddress);
    checkForActiveMeeting();
}

async function loadAvailability(email) {
    if (!email) return;
    const spinner = document.getElementById("loadingSpinner");
    if (spinner) spinner.style.display = "inline";

    const now = new Date();
    const dayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
    const dayEnd   = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);

    try {
        const res = await fetch(`${API_URL}/availability`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ room_email: email, start_time: dayStart.toISOString(), end_time: dayEnd.toISOString(), time_zone: "UTC" })
        });
        const data = await res.json();
        const hasMeetings = (data?.value?.[0]?.scheduleItems || []).some(i => i.status === "busy");
        const calBtn = document.getElementById("roomCalendarBtn");
        if (calBtn) calBtn.style.display = hasMeetings ? "flex" : "none";
    } catch (e) { console.error(e); }
    finally { if (spinner) spinner.style.display = "none"; }
}

// ═══════════════════════════════════════════
//  BOOKING AVAILABILITY STRIP
// ═══════════════════════════════════════════
let stripFetchTimeout = null;
let lastStripDate = null;
let lastStripData = null;

function refreshBookingTimeline() {
    const startVal = document.getElementById("startTime").value;
    const endVal   = document.getElementById("endTime").value;
    if (!startVal) return;

    const startDate = new Date(startVal);
    const dateKey = startDate.toDateString();

    clearTimeout(stripFetchTimeout);
    stripFetchTimeout = setTimeout(async () => {
        if (dateKey !== lastStripDate) {
            lastStripDate = dateKey;
            await fetchStripData(startDate);
        }
        renderStrip(startVal, endVal);
    }, 250);
}

async function fetchStripData(forDate) {
    const idx = document.getElementById("roomSelect").value;
    if (idx === "") return;
    const roomEmail = availableRooms[idx].emailAddress;

    const dayStart = new Date(forDate.getFullYear(), forDate.getMonth(), forDate.getDate(), 0, 0, 0);
    const dayEnd   = new Date(forDate.getFullYear(), forDate.getMonth(), forDate.getDate(), 23, 59, 59);

    const track = document.getElementById("availStripTrack");
    if (track) track.innerHTML = `<div class="avail-strip-loading">Loading…</div>`;

    try {
        const res = await fetch(`${API_URL}/availability`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                room_email: roomEmail,
                start_time: dayStart.toISOString(),
                end_time: dayEnd.toISOString(),
                time_zone: "UTC"
            })
        });
        const data = await res.json();
        lastStripData = data?.value?.[0]?.scheduleItems || [];

        const dateLabel = document.getElementById("availStripDate");
        const today = new Date();
        if (dateLabel) {
            if (forDate.toDateString() === today.toDateString()) dateLabel.textContent = "Today";
            else if (forDate.toDateString() === new Date(today.getTime() + 86400000).toDateString()) dateLabel.textContent = "Tomorrow";
            else dateLabel.textContent = forDate.toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" });
        }
    } catch (e) {
        lastStripData = [];
    }
}

function renderStrip(startVal, endVal) {
    const track  = document.getElementById("availStripTrack");
    const hours  = document.getElementById("availStripHours");
    const conflict = document.getElementById("availConflict");
    if (!track) return;

    const STRIP_START_H = 7;
    const STRIP_END_H   = 22;
    const TOTAL_MINS    = (STRIP_END_H - STRIP_START_H) * 60;

    const baseDate = startVal ? new Date(startVal) : new Date();
    const stripStart = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate(), STRIP_START_H, 0, 0);
    const stripEnd   = new Date(baseDate.getFullYear(), baseDate.getMonth(), baseDate.getDate(), STRIP_END_H, 0, 0);

    track.innerHTML = "";

    if (hours) {
        hours.innerHTML = "";
        for (let h = STRIP_START_H; h <= STRIP_END_H; h += 2) {
            const pct = ((h - STRIP_START_H) / (STRIP_END_H - STRIP_START_H)) * 100;
            const lbl = document.createElement("span");
            lbl.className = "avail-hour-label";
            lbl.style.left = pct + "%";
            lbl.textContent = `${String(h).padStart(2,"0")}:00`;
            hours.appendChild(lbl);
        }
    }

    let hasConflict = false;
    const selStart = startVal ? new Date(startVal) : null;
    const selEnd   = endVal   ? new Date(endVal)   : null;

    (lastStripData || []).forEach(item => {
        if (item.status !== "busy") return;
        const bs = new Date(item.start.dateTime + "Z");
        const be = new Date(item.end.dateTime + "Z");
        const leftMs  = Math.max(bs - stripStart, 0);
        const rightMs = Math.min(be - stripStart, TOTAL_MINS * 60000);
        if (rightMs <= 0 || leftMs >= TOTAL_MINS * 60000) return;

        const leftPct  = (leftMs  / (TOTAL_MINS * 60000)) * 100;
        const widthPct = ((rightMs - leftMs) / (TOTAL_MINS * 60000)) * 100;

        const block = document.createElement("div");
        block.className = "avail-busy-block";
        block.style.left  = leftPct  + "%";
        block.style.width = widthPct + "%";

        const st = bs.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
        const et = be.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
        block.title = `${item.subject || "Busy"}: ${st} – ${et}`;
        track.appendChild(block);

        if (selStart && selEnd && bs < selEnd && be > selStart) hasConflict = true;
    });

    if (selStart && selEnd && selEnd > selStart) {
        const sleftMs  = Math.max(selStart - stripStart, 0);
        const srightMs = Math.min(selEnd   - stripStart, TOTAL_MINS * 60000);
        if (srightMs > 0 && sleftMs < TOTAL_MINS * 60000) {
            const sleftPct  = (sleftMs  / (TOTAL_MINS * 60000)) * 100;
            const swidthPct = ((srightMs - sleftMs) / (TOTAL_MINS * 60000)) * 100;
            const sel = document.createElement("div");
            sel.className = "avail-selected-block" + (hasConflict ? " conflict" : "");
            sel.style.left  = sleftPct  + "%";
            sel.style.width = swidthPct + "%";
            track.appendChild(sel);
        }
    }

    if (conflict) {
        conflict.classList.toggle("hidden", !hasConflict);
    }
}

function initModalTimes() {
    const now = new Date();
    const offset = now.getTimezoneOffset();
    now.setMinutes(now.getMinutes() - offset);
    const start = document.getElementById("startTime");
    const end = document.getElementById("endTime");
    if (!start || !end) return;
    start.value = now.toISOString().slice(0, 16);
    now.setMinutes(now.getMinutes() + 30);
    end.value = now.toISOString().slice(0, 16);
}

function showToast(msg, isError = false) {
    const t = document.getElementById("toastBar");
    if (!t) return;
    t.textContent = msg;
    t.className = "toast-bar" + (isError ? " error" : "");
    t.classList.remove("hidden");
    clearTimeout(t._timeout);
    t._timeout = setTimeout(() => t.classList.add("hidden"), 3500);
}
