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
    cache: { cacheLocation: "sessionStorage" }
};
const loginRequest = { scopes: ["User.Read", "Calendars.ReadWrite"] };
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
        const resp = await myMSALObj.handleRedirectPromise();
        if (resp) {
            handleLoginSuccess(resp.account);
            // If we were redirected mid-book, restore room selection then open modal
            if (sessionStorage.getItem("pendingBook") === "1") {
                sessionStorage.removeItem("pendingBook");
                const savedRoom = sessionStorage.getItem("pendingBookRoom");
                sessionStorage.removeItem("pendingBookRoom");
                if (savedRoom !== null && availableRooms[savedRoom]) {
                    const sel = document.getElementById("roomSelect");
                    sel.value = savedRoom;
                    handleRoomChange();
                }
                // Small delay to let room load before opening modal
                setTimeout(openBookingModal, 300);
            }
        }
    } catch (e) { console.error(e); }
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
//  AUTH
// ═══════════════════════════════════════════
async function signIn() {
    if (isAuthInProgress) return;
    isAuthInProgress = true;
    try { await myMSALObj.loginRedirect(loginRequest); }
    catch (e) { console.error(e); }
    finally { isAuthInProgress = false; }
}

function signOut() {
    username = "";
    const badge = document.getElementById("userBadge");
    const loginBtn = document.getElementById("loginBtn");
    if (badge) badge.style.display = "none";
    if (loginBtn) loginBtn.style.display = "inline-block";
    if (sessionTimeout) clearTimeout(sessionTimeout);
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
    // Save both the pending flag AND the currently selected room index
    const idx = document.getElementById("roomSelect").value;
    sessionStorage.setItem("pendingBook", "1");
    if (idx !== "") sessionStorage.setItem("pendingBookRoom", idx);
    await signIn();
}

// ═══════════════════════════════════════════
//  BOOKING MODAL
// ═══════════════════════════════════════════
function openBookingModal() {
    document.getElementById("displayEmail").value = username;
    initModalTimes();
    document.getElementById("bookingOverlay").classList.remove("hidden");
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
        const end = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59);
        const res = await fetch(`${API_URL}/availability`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                room_email: roomEmail,
                start_time: now.toISOString(),
                end_time: end.toISOString(),
                time_zone: "UTC"
            })
        });
        const data = await res.json();
        const items = data?.value?.[0]?.scheduleItems || [];
        const busy = items.filter(i => i.status === "busy");

        if (busy.length === 0) {
            content.innerHTML = `<div class="occ-agenda-empty">No more meetings today.</div>`;
            return;
        }

        content.innerHTML = busy.map(item => {
            const s = new Date(item.start.dateTime + "Z").toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            const e = new Date(item.end.dateTime + "Z").toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            return `
                <div class="agenda-modal-item">
                    <div class="agenda-modal-time">${s} – ${e}</div>
                    <div>
                        <div class="agenda-modal-subject">${item.subject || "Meeting"}</div>
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

        // Timeline auto-refresh on change
        const cid = event ? event.id : "free";
        if (lastKnownEventId !== "init" && lastKnownEventId !== cid) {
            loadAvailability(roomEmail);
        }
        lastKnownEventId = cid;

        const occupied = document.getElementById("occupiedScreen");
        const main = document.getElementById("mainScreen");

        // ── NO MEETING ──
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

        // Build clean display subject
        let displaySubject = event.subject || "Meeting";
        let displayOrg = event.organizer?.emailAddress?.name || "Unknown";

        // If privacy-masked, preserve what's already on screen
        if (event.subject === "Busy" && !occupied.classList.contains("hidden")) {
            const existing = document.getElementById("occSubject").textContent;
            if (existing && existing !== "—") displaySubject = existing;
        }

        if (displaySubject === displayOrg || !displaySubject) {
            displaySubject = "Private Meeting";
        }

        const startFmt = start.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
        const endFmt = end.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });

        // ── FUTURE ──
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

        // ── ACTIVE + CHECKED IN → RED ──
        if (event.categories?.includes("Checked-In")) {
            setAppState("occupied");
            stopCountdowns();
            if (occupied.classList.contains("hidden") && event.id !== manuallyUnlockedEventId) {
                showMeetingMode(event, displaySubject, displayOrg, startFmt, endFmt);
            }
            return;
        }

        // ── ACTIVE + PENDING ──
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
//  CHECK-IN (No Auth — tap to confirm presence)
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

    if (sessionTimeout) clearTimeout(sessionTimeout); // Don't auto-logout while in meeting
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
        // Start window = end of current meeting so we only get TRULY upcoming ones
        const windowStart = currentMeetingEndStr
            ? new Date(currentMeetingEndStr + "Z")
            : new Date();
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
        // No slice needed — window starts after current meeting ends
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

    // Check adjacent slot availability
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

    // Call backend extend
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
//  SECURE END MEETING (SSO Required)
// ═══════════════════════════════════════════
async function secureEndMeeting() {
    if (isAuthInProgress || !currentLockedEvent) return;

    const organizerEmail = currentLockedEvent.organizer?.emailAddress?.address?.toLowerCase() || "";
    const attendees = currentLockedEvent.attendees || [];
    const allowed = [...attendees.map(a => a.emailAddress?.address?.toLowerCase()), organizerEmail];

    const roomIdx = document.getElementById("roomSelect").value;
    const roomEmail = availableRooms[roomIdx].emailAddress;
    const eventId = currentLockedEvent.id;

    isAuthInProgress = true;
    try {
        const r = await myMSALObj.loginPopup({ scopes: ["User.Read"], prompt: "select_account" });
        const userEmail = r.account.username.toLowerCase();

        if (!allowed.includes(userEmail)) {
            showToast(`⛔ Access denied — you are not authorized to end this meeting.`, true);
            return;
        }

        const res = await fetch(`${API_URL}/end-meeting`, {
            method: "POST",
            headers: { "Content-Type": "application/json", "Authorization": `Bearer ${await getAuthToken()}` },
            body: JSON.stringify({ room_email: roomEmail, event_id: eventId })
        });

        if (res.ok) {
            manuallyUnlockedEventId = eventId;
            currentLockedEvent = null;
            stopCountdowns();
            showOccupied(false);
            setAppState("available");
            showView("viewAvailable");
            loadAvailability(roomEmail);
        } else {
            showToast("Failed to end meeting.", true);
        }
    } catch (e) {
        if (e.errorCode !== "user_cancelled") showToast("Authentication failed.", true);
    } finally { isAuthInProgress = false; }
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
    // Build a full-screen success flash then auto-dismiss
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
            <div><strong style="color:#22c46e;">Subject</strong>&nbsp;&nbsp;${subject}</div>
            <div><strong style="color:#22c46e;">Unit</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;${filiale}</div>
            <div><strong style="color:#22c46e;">Time</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;${timeRange}</div>
            <div><strong style="color:#22c46e;">Invitees</strong>&nbsp;${invitees || "None"}</div>
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
        signOut(); // Auto-logout after booking (kiosk practice)
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

    // Update sidebar meta
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
    const viewStart = new Date(now);
    viewStart.setHours(viewStart.getHours() - 2, 0, 0, 0);
    const viewEnd = new Date(viewStart.getTime() + 12 * 3600000);

    try {
        const res = await fetch(`${API_URL}/availability`, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ room_email: email, start_time: viewStart.toISOString(), end_time: viewEnd.toISOString(), time_zone: "UTC" })
        });
        const data = await res.json();
        renderTimeline(data, viewStart, viewEnd);
    } catch (e) { console.error(e); }
    finally { if (spinner) spinner.style.display = "none"; }
}

function renderTimeline(data, viewStart, viewEnd) {
    // Timeline is embedded in the agenda overlay only; main screen doesn't need it
    // (It's available via openAgenda)
}

// ═══════════════════════════════════════════
//  UTILITIES
// ═══════════════════════════════════════════
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
