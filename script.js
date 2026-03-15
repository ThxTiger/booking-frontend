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
    cache: { cacheLocation: "localStorage" } // required — sessionStorage dies on kiosk redirect
};

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
let agendaRefreshInterval = null;
let clockInterval = null;
let sessionTimeout = null;
let isAuthInProgress = false;
let manuallyUnlockedEventId = null;
let lastKnownEventId = "init";
let currentAppState = "available";
let announcedMeetings = [];
let announcedEndings = [];

// ═══════════════════════════════════════════
//  FIX-03 (VULN-03): HTML ESCAPE HELPER
//  Applied to every piece of API-sourced text
//  before it is injected into innerHTML.
//  Prevents XSS via crafted meeting subjects.
// ═══════════════════════════════════════════
function escapeHtml(str) {
    return String(str)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

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
        if (el) el.classList.remove("active");
    });
    const target = document.getElementById(viewId);
    if (target) target.classList.add("active");
}

// ═══════════════════════════════════════════
//  AUDIO WARNING SYSTEMS
// ═══════════════════════════════════════════
function playEvictionWarning(minutes) {
    const file = minutes >= 14 ? './alerte-15min.mp3' : './alerte-imminente.mp3';
    const audio = new Audio(file);
    audio.play().catch(e => console.error("Audio blocked:", e));
    setTimeout(() => {
        new Audio(file).play().catch(e => console.error("Audio blocked:", e));
    }, 10000);
}

function playMeetingEndWarning() {
    new Audio('./alerte-5min.mp3').play().catch(e => console.error("Audio blocked:", e));
}

// ═══════════════════════════════════════════
//  DATA SAVER
// ═══════════════════════════════════════════
function saveFormDataToStorage() {
    const idx = document.getElementById("roomSelect").value;
    if (idx === "") return;
    localStorage.setItem("pendingBookRoom", idx);
    const fields = { pbSubj: "subject", pbFil: "filiale", pbDesc: "description", pbAtt: "attendees", pbStart: "startTime", pbEnd: "endTime" };
    Object.entries(fields).forEach(([key, id]) => {
        const val = document.getElementById(id).value.trim();
        if (val) localStorage.setItem(key, val);
    });
}

// ═══════════════════════════════════════════
//  INITIALIZATION & REDIRECT ROUTING
// ═══════════════════════════════════════════
document.addEventListener("DOMContentLoaded", async () => {

    document.querySelectorAll("input").forEach(input => {
        input.setAttribute("autocomplete", "new-password");
        input.setAttribute("data-lpignore", "true");
        input.setAttribute("spellcheck", "false");
        input.removeAttribute("list");
    });

    initModalTimes();
    startClock();
    await fetchRooms();

    setInterval(checkForActiveMeeting, 5000);
    setInterval(() => {
        const idx = document.getElementById("roomSelect").value;
        if (idx !== "") loadAvailability(availableRooms[idx].emailAddress);
    }, 60000);

    try {
        await myMSALObj.initialize();
        const redirectResponse = await myMSALObj.handleRedirectPromise();

        if (redirectResponse) {
            handleLoginSuccess(redirectResponse.account);

            const pendingRoom = localStorage.getItem("pendingBookRoom");
            if (pendingRoom !== null) {
                const restored = {
                    subj:  localStorage.getItem("pbSubj"),
                    fil:   localStorage.getItem("pbFil"),
                    desc:  localStorage.getItem("pbDesc"),
                    att:   localStorage.getItem("pbAtt"),
                    start: localStorage.getItem("pbStart"),
                    end:   localStorage.getItem("pbEnd"),
                };
                ["pendingBookRoom","pbSubj","pbFil","pbDesc","pbAtt","pbStart","pbEnd"]
                    .forEach(k => localStorage.removeItem(k));

                document.getElementById("roomSelect").value = pendingRoom;
                handleRoomChange();

                setTimeout(() => {
                    openBookingModal();
                    if (restored.subj)  document.getElementById("subject").value     = restored.subj;
                    if (restored.fil)   document.getElementById("filiale").value     = restored.fil;
                    if (restored.desc)  document.getElementById("description").value = restored.desc;
                    if (restored.att)   document.getElementById("attendees").value   = restored.att;
                    if (restored.start) document.getElementById("startTime").value   = restored.start;
                    if (restored.end)   document.getElementById("endTime").value     = restored.end;
                }, 500);
            }

            const pendingEndId = localStorage.getItem("pendingEndEventId");
            if (pendingEndId) {
                const roomIdx = localStorage.getItem("pendingEndRoomIdx");
                const allowed = JSON.parse(localStorage.getItem("pendingEndAllowed") || "[]");
                ["pendingEndEventId","pendingEndRoomIdx","pendingEndAllowed"]
                    .forEach(k => localStorage.removeItem(k));
                if (availableRooms[roomIdx]) {
                    document.getElementById("roomSelect").value = roomIdx;
                    handleRoomChange();
                    setTimeout(() => {
                        processSecureEnd(
                            redirectResponse.account.username,
                            allowed,
                            availableRooms[roomIdx].emailAddress,
                            pendingEndId
                        );
                    }, 800);
                }
            }

        } else {
            const accounts = myMSALObj.getAllAccounts();
            if (accounts.length > 0) handleLoginSuccess(accounts[0]);
            ["pendingBookRoom","pbSubj","pbFil","pbDesc","pbAtt","pbStart","pbEnd",
             "pendingEndEventId","pendingEndRoomIdx","pendingEndAllowed"]
                .forEach(k => localStorage.removeItem(k));
        }
    } catch (e) {
        console.error("Auth init error:", e);
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
//  AUTH — REDIRECT FLOW
// ═══════════════════════════════════════════
async function signIn() {
    if (isAuthInProgress) return;
    isAuthInProgress = true;
    try {
        await myMSALObj.loginRedirect(loginRequest);
    } catch (e) {
        console.error(e);
        isAuthInProgress = false;
    }
}

async function signOut() {
    username = "";
    const badge    = document.getElementById("userBadge");
    const loginBtn = document.getElementById("loginBtn");
    if (badge)    badge.style.display = "none";
    if (loginBtn) loginBtn.style.display = "inline-block";
    if (sessionTimeout) clearTimeout(sessionTimeout);

    try {
        const account = myMSALObj.getAllAccounts()[0];
        if (account) {
            await myMSALObj.logoutRedirect({
                account,
                onRedirectNavigate: () => false // kill MS session, stay on page
            });
        }
    } catch (e) {
        console.error("Silent logout error:", e);
    }

    localStorage.clear();
    sessionStorage.clear();
    stopCountdowns();
    checkForActiveMeeting();
}

function handleLoginSuccess(acc) {
    username = acc.username;
    const welcome  = document.getElementById("userWelcome");
    const badge    = document.getElementById("userBadge");
    const loginBtn = document.getElementById("loginBtn");
    if (welcome)  welcome.textContent = username;
    if (badge)    badge.style.display = "flex";
    if (loginBtn) loginBtn.style.display = "none";
    if (sessionTimeout) clearTimeout(sessionTimeout);
    sessionTimeout = setTimeout(() => signOut(), 120000);
}

async function getAuthToken() {
    try {
        const account = myMSALObj.getAllAccounts()[0];
        if (!account) return null;
        const r = await myMSALObj.acquireTokenSilent({ scopes: ["User.Read", "Calendars.ReadWrite"], account });
        return r.accessToken;
    } catch { return null; }
}

// ═══════════════════════════════════════════
//  AUTH GATE
// ═══════════════════════════════════════════
function handleBookClick() {
    if (!username) openAuthGate();
    else           openBookingModal();
}

function openAuthGate()  { document.getElementById("authGateOverlay").classList.remove("hidden"); }
function closeAuthGate() { document.getElementById("authGateOverlay").classList.add("hidden"); }

async function triggerSignInThenBook() {
    closeAuthGate();
    saveFormDataToStorage();
    await signIn();
}

// ═══════════════════════════════════════════
//  BOOKING MODAL
// ═══════════════════════════════════════════
function openBookingModal() {
    document.getElementById("displayEmail").value = username;
    if (!document.getElementById("startTime").value) initModalTimes();
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
    const content   = document.getElementById("agendaContent");
    const idx       = document.getElementById("roomSelect").value;
    if (idx === "") { content.innerHTML = `<div class="occ-agenda-empty">Please select a room first.</div>`; return; }
    const roomEmail = availableRooms[idx].emailAddress;
    content.innerHTML = `<div class="occ-agenda-empty">Loading…</div>`;
    try {
        const now      = new Date();
        const dayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
        const dayEnd   = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);
        const res = await fetch(`${API_URL}/availability`, {
            method: "POST", headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ room_email: roomEmail, start_time: dayStart.toISOString(), end_time: dayEnd.toISOString(), time_zone: "UTC" })
        });
        const data  = await res.json();
        const busy  = (data?.value?.[0]?.scheduleItems || []).filter(i => i.status === "busy");
        if (busy.length === 0) { content.innerHTML = `<div class="occ-agenda-empty">No meetings scheduled today.</div>`; return; }
        // ── FIX-03: escapeHtml on item.subject (XSS fix — VULN-03, lines 329-342) ──
        content.innerHTML = busy.map(item => {
            const iS    = new Date(item.start.dateTime + "Z");
            const iE    = new Date(item.end.dateTime   + "Z");
            const s     = iS.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            const e     = iE.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            const isNow  = now >= iS && now < iE;
            const isPast = now >= iE;
            const badge  = isNow  ? `<span class="agenda-badge now">NOW</span>`
                         : isPast ? `<span class="agenda-badge past">DONE</span>` : "";
            return `<div class="agenda-modal-item${isPast?" past":isNow?" active-now":""}">
                <div class="agenda-modal-time">${s} – ${e}</div>
                <div style="flex:1"><div class="agenda-modal-subject">${escapeHtml(item.subject||"Meeting")} ${badge}</div></div>
            </div>`;
        }).join("");
    } catch { content.innerHTML = `<div class="occ-agenda-empty">Failed to load.</div>`; }
}

function closeAgenda() { document.getElementById("agendaOverlay").classList.add("hidden"); }

// ═══════════════════════════════════════════
//  HEARTBEAT
// ═══════════════════════════════════════════
async function checkForActiveMeeting() {
    const idx = document.getElementById("roomSelect").value;
    if (idx === "") return;
    const roomEmail = availableRooms[idx].emailAddress;
    try {
        const token   = await getAuthToken();
        const headers = { "Content-Type": "application/json" };
        if (token) headers["Authorization"] = `Bearer ${token}`;
        const res   = await fetch(`${API_URL}/active-meeting?room_email=${roomEmail}`, { headers });
        if (res.status === 401) return;
        const event = await res.json();

        const cid = event ? event.id : "free";
        if (lastKnownEventId !== "init" && lastKnownEventId !== cid) loadAvailability(roomEmail);
        lastKnownEventId = cid;

        const occupied = document.getElementById("occupiedScreen");

        if (!event) {
            setAppState("available"); showOccupied(false); showView("viewAvailable");
            stopCountdowns(); updateNextMeetingPreview(null); return;
        }

        const now   = new Date();
        const start = new Date(event.start.dateTime + "Z");
        const end   = new Date(event.end.dateTime   + "Z");

        if (now >= end) { setAppState("available"); showOccupied(false); showView("viewAvailable"); return; }

        let displaySubject = event.subject || "Meeting";
        let displayOrg     = event.organizer?.emailAddress?.name || "Unknown";

        if (event.subject === "Busy" && !occupied.classList.contains("hidden")) {
            const existing = document.getElementById("occSubject").textContent;
            if (existing && existing !== "—") displaySubject = existing;
        }
        if (displaySubject === displayOrg || !displaySubject) displaySubject = "Private Meeting";

        const startFmt = start.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
        const endFmt   = end.toLocaleTimeString([],   { hour: "2-digit", minute: "2-digit" });

        // FUTURE
        if (now < start) {
            setAppState("available"); showOccupied(false); showView("viewFuture");
            document.getElementById("futureSubject").textContent = displaySubject;
            document.getElementById("futureTime").textContent    = `${startFmt} – ${endFmt}`;
            startCountdown(start, "futureTimer", "STARTING…");
            updateNextMeetingPreview({ subject: displaySubject, startFmt, endFmt });
            const minsToStart = Math.round((start - now) / 60000);
            if (minsToStart <= 15 && minsToStart > 0 && !announcedMeetings.includes(event.id)) {
                announcedMeetings.push(event.id);
                playEvictionWarning(minsToStart);
            }
            return;
        }

        // CHECKED IN → RED
        if (event.categories?.includes("Checked-In")) {
            setAppState("occupied");
            if (checkInInterval) { clearInterval(checkInInterval); checkInInterval = null; }
            if (occupied.classList.contains("hidden") && event.id !== manuallyUnlockedEventId)
                showMeetingMode(event, displaySubject, displayOrg, startFmt, endFmt);
            return;
        }

        // PENDING → ORANGE
        if (event.id !== manuallyUnlockedEventId) manuallyUnlockedEventId = null;
        setAppState("pending"); showOccupied(false); showView("viewPending");
        document.getElementById("pendingSubject").textContent   = displaySubject;
        document.getElementById("pendingTime").textContent      = `${startFmt} – ${endFmt}`;
        document.getElementById("pendingOrganizer").textContent = `Organized by ${displayOrg}`;
        const deadline = new Date(start.getTime() + 5 * 60000);
        startCountdown(deadline, "checkInTimer", "EXPIRED");
        document.getElementById("realCheckInBtn").onclick = () => performCheckIn(roomEmail, event.id, event);

    } catch (e) { console.error(e); }
}

function showOccupied(show) {
    const occ  = document.getElementById("occupiedScreen");
    const main = document.getElementById("mainScreen");
    if (show) { occ.classList.remove("hidden"); main.classList.add("hidden"); }
    else       { occ.classList.add("hidden");   main.classList.remove("hidden"); }
}

function updateNextMeetingPreview(data) {
    const preview = document.getElementById("nextMeetingPreview");
    if (!data || !preview) { if (preview) preview.style.display = "none"; return; }
    preview.style.display = "block";
    document.getElementById("nextSubject").textContent = data.subject;
    document.getElementById("nextTime").textContent    = `${data.startFmt} – ${data.endFmt}`;
}

// ═══════════════════════════════════════════
//  COUNTDOWNS
// ═══════════════════════════════════════════
function startCountdown(targetDate, elementId, expireText) {
    if (checkInInterval) clearInterval(checkInInterval);
    checkInInterval = setInterval(() => {
        const dist = targetDate - new Date();
        const el   = document.getElementById(elementId);
        if (!el) return;
        if (dist <= 0) { el.textContent = expireText; return; }
        const m = Math.floor(dist / 60000);
        const s = Math.floor((dist % 60000) / 1000);
        el.textContent = `${m}:${String(s).padStart(2, "0")}`;
    }, 1000);
}

function stopCountdowns() {
    if (checkInInterval)      { clearInterval(checkInInterval);      checkInInterval      = null; }
    if (meetingEndInterval)   { clearInterval(meetingEndInterval);   meetingEndInterval   = null; }
    if (agendaRefreshInterval){ clearInterval(agendaRefreshInterval); agendaRefreshInterval = null; }
}

// ═══════════════════════════════════════════
//  CHECK-IN (no auth — physical presence)
// ═══════════════════════════════════════════
async function performCheckIn(roomEmail, eventId, eventDetails) {
    if (checkInInterval) clearInterval(checkInInterval);
    try {
        const res = await fetch(`${API_URL}/checkin`, {
            method: "POST", headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ room_email: roomEmail, event_id: eventId })
        });
        if (res.ok) {
            const startFmt = new Date(eventDetails.start.dateTime + "Z").toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            const endFmt   = new Date(eventDetails.end.dateTime   + "Z").toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            showMeetingMode(eventDetails, document.getElementById("pendingSubject").textContent,
                eventDetails.organizer?.emailAddress?.name, startFmt, endFmt);
        } else { showToast("Check-in failed. Try again.", true); checkForActiveMeeting(); }
    } catch { showToast("Network error.", true); }
}

// ═══════════════════════════════════════════
//  MEETING MODE (Red Screen)
// ═══════════════════════════════════════════
function showMeetingMode(event, subject, organizer, startFmt, endFmt) {
    currentLockedEvent = event;
    setAppState("occupied");
    showOccupied(true);
    document.getElementById("occSubject").textContent   = subject   || "Meeting";
    document.getElementById("occTime").textContent      = `${startFmt} – ${endFmt}`;
    document.getElementById("occOrganizer").textContent = `Organized by ${organizer || "Unknown"}`;
    startMeetingEndTimer(event.end.dateTime);
    updateEndsIn(new Date(event.end.dateTime + "Z"));
    loadOccupiedAgenda(
        availableRooms[document.getElementById("roomSelect").value]?.emailAddress,
        event.end.dateTime
    );
    if (sessionTimeout) clearTimeout(sessionTimeout);

    // Refresh "Coming Up" every 60 s so newly booked meetings appear automatically
    if (agendaRefreshInterval) clearInterval(agendaRefreshInterval);
    agendaRefreshInterval = setInterval(() => {
        const idx = document.getElementById("roomSelect").value;
        const email = idx !== "" ? availableRooms[idx]?.emailAddress : null;
        if (email && currentLockedEvent) {
            loadOccupiedAgenda(email, currentLockedEvent.end.dateTime);
        }
    }, 60000);
}

function updateEndsIn(endDate) {
    const mins = Math.max(0, Math.round((endDate - new Date()) / 60000));
    const el   = document.getElementById("occEndsIn");
    if (el) el.textContent = mins > 0 ? `Ends in ${mins} min` : "Ending now";
}

function startMeetingEndTimer(endTimeStr) {
    if (meetingEndInterval) clearInterval(meetingEndInterval);
    const endTime = new Date(endTimeStr + "Z").getTime();
    meetingEndInterval = setInterval(() => {
        const dist = endTime - Date.now();
        if (dist <= 5 * 60000 && dist > 0 && currentLockedEvent && !announcedEndings.includes(currentLockedEvent.id)) {
            announcedEndings.push(currentLockedEvent.id);
            playMeetingEndWarning();
        }
        if (dist <= 0) {
            clearInterval(meetingEndInterval);
            showOccupied(false); setAppState("available"); showView("viewAvailable");
            checkForActiveMeeting();
        } else {
            const m  = Math.floor(dist / 60000);
            const s  = Math.floor((dist % 60000) / 1000);
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
        const windowEnd   = new Date(windowStart.getFullYear(), windowStart.getMonth(), windowStart.getDate(), 23, 59);
        const res = await fetch(`${API_URL}/availability`, {
            method: "POST", headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ room_email: roomEmail, start_time: windowStart.toISOString(), end_time: windowEnd.toISOString(), time_zone: "UTC" })
        });
        const data     = await res.json();
        // Show events that end after the current meeting ends,
        // but exclude the current event itself (matched by time overlap with windowStart).
        const upcoming = (data?.value?.[0]?.scheduleItems || []).filter(i => {
            if (i.status !== "busy") return false;
            const itemEnd   = new Date(i.end.dateTime   + "Z");
            const itemStart = new Date(i.start.dateTime + "Z");
            // Must end after the current meeting ends (excludes the current meeting itself)
            if (itemEnd <= windowStart) return false;
            // Exclude the current locked event by matching its exact start time
            if (currentLockedEvent) {
                const curStart = new Date(currentLockedEvent.start.dateTime + "Z");
                if (Math.abs(itemStart - curStart) < 60000) return false;
            }
            return true;
        });
        if (upcoming.length === 0) { list.innerHTML = `<div class="occ-agenda-empty">No more meetings today.</div>`; return; }
        // ── FIX-03: escapeHtml on item.subject (XSS fix — VULN-03, lines 546-553) ──
        list.innerHTML = upcoming.map(item => {
            const s = new Date(item.start.dateTime + "Z").toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            const e = new Date(item.end.dateTime   + "Z").toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            return `<div class="occ-agenda-item">
                <div class="occ-agenda-item-time">${s} – ${e}</div>
                <div class="occ-agenda-item-subj">${escapeHtml(item.subject || "Meeting")}</div>
            </div>`;
        }).join("");
    } catch {}
}

// ═══════════════════════════════════════════
//  +15 MIN EXTENSION
// ═══════════════════════════════════════════
async function extendMeeting(minutes) {
    if (!currentLockedEvent) return;
    const roomIdx   = document.getElementById("roomSelect").value;
    const roomEmail = availableRooms[roomIdx].emailAddress;
    const curEnd    = new Date(currentLockedEvent.end.dateTime + "Z");
    const newEnd    = new Date(curEnd.getTime() + minutes * 60000);

    try {
        const res  = await fetch(`${API_URL}/availability`, {
            method: "POST", headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ room_email: roomEmail, start_time: curEnd.toISOString(), end_time: newEnd.toISOString(), time_zone: "UTC" })
        });
        const data   = await res.json();
        const isBusy = (data?.value?.[0]?.scheduleItems || []).some(item => {
            if (item.status !== "busy") return false;
            const iS = new Date(item.start.dateTime + "Z");
            const iE = new Date(item.end.dateTime   + "Z");
            return iS < newEnd && iE > curEnd;
        });
        if (isBusy) { showToast("⛔ Cannot extend — another meeting follows immediately.", true); return; }
    } catch (e) { console.error("Availability check failed:", e); }

    try {
        const res = await fetch(`${API_URL}/extend-meeting`, {
            method: "POST", headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ room_email: roomEmail, event_id: currentLockedEvent.id, extend_minutes: minutes })
        });
        if (res.ok) {
            currentLockedEvent.end.dateTime = newEnd.toISOString().replace("Z", "");
            announcedEndings = announcedEndings.filter(id => id !== currentLockedEvent.id);
            startMeetingEndTimer(currentLockedEvent.end.dateTime);
            const startFmt  = new Date(currentLockedEvent.start.dateTime + "Z").toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            const newEndFmt = newEnd.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            document.getElementById("occTime").textContent = `${startFmt} – ${newEndFmt}`;
            showToast(`✅ Extended by ${minutes} min — now ends at ${newEndFmt}`);
            loadAvailability(roomEmail);
        } else {
            const err = await res.json().catch(() => ({}));
            showToast(err.detail || "Extension failed.", true);
        }
    } catch { showToast("Network error.", true); }
}

// ═══════════════════════════════════════════
//  SECURE END MEETING
// ═══════════════════════════════════════════
async function secureEndMeeting() {
    if (isAuthInProgress || !currentLockedEvent) return;
    const roomIdx   = document.getElementById("roomSelect").value;
    const roomEmail = availableRooms[roomIdx].emailAddress;

    const organizerEmail = currentLockedEvent.organizer?.emailAddress?.address?.toLowerCase() || "";
    const attendees      = currentLockedEvent.attendees || [];
    const allowed        = [...attendees.map(a => a.emailAddress?.address?.toLowerCase()), organizerEmail];

    localStorage.setItem("pendingEndEventId", currentLockedEvent.id);
    localStorage.setItem("pendingEndRoomIdx", roomIdx);
    localStorage.setItem("pendingEndAllowed", JSON.stringify(allowed));

    isAuthInProgress = true;
    try {
        await myMSALObj.loginRedirect({ scopes: ["User.Read"], prompt: "select_account" });
    } catch (e) {
        isAuthInProgress = false;
        console.error(e);
    }
}

async function processSecureEnd(userEmail, allowedList, roomEmail, eventId) {
    if (!allowedList.includes(userEmail.toLowerCase())) {
        showToast("⛔ Access denied — you are not authorized to end this meeting.", true);
        localStorage.clear(); sessionStorage.clear();
        return;
    }
    try {
        const token = await getAuthToken();
        const res   = await fetch(`${API_URL}/end-meeting`, {
            method: "POST",
            headers: { "Content-Type": "application/json", "Authorization": `Bearer ${token}` },
            body: JSON.stringify({ room_email: roomEmail, event_id: eventId })
        });
        if (res.ok) {
            manuallyUnlockedEventId = eventId;
            currentLockedEvent = null;
            stopCountdowns();
            showOccupied(false); setAppState("available"); showView("viewAvailable");
            loadAvailability(roomEmail);
            showToast("✅ Meeting ended. Signing out…");
            // Security: sign out immediately so no one can book under the previous user's identity
            setTimeout(() => signOut(), 1500);
        } else {
            const err = await res.json().catch(() => ({}));
            showToast(err.detail || "Failed to end meeting.", true);
        }
        localStorage.clear(); sessionStorage.clear();
    } catch (e) {
        showToast("Network error.", true); console.error(e);
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

    const roomEmail    = availableRooms[idx].emailAddress;
    const subject      = document.getElementById("subject").value.trim();
    const filiale      = document.getElementById("filiale").value.trim();
    const desc         = document.getElementById("description").value.trim();
    const startVal     = document.getElementById("startTime").value;
    const endVal       = document.getElementById("endTime").value;
    const attendeesRaw = document.getElementById("attendees").value;
    const attendeeList = attendeesRaw.trim() ? attendeesRaw.split(",").map(e => e.trim()).filter(Boolean) : [];

    if (!subject || !filiale || !startVal || !endVal) {
        showToast("Please fill in all required fields.", true); return;
    }

    let accessToken = "";
    try {
        const account = myMSALObj.getAllAccounts()[0];
        if (!account) throw new Error("No active account");
        const r = await myMSALObj.acquireTokenSilent({ ...loginRequest, account });
        accessToken = r.accessToken;
    } catch {
        saveFormDataToStorage();
        myMSALObj.loginRedirect(loginRequest);
        return;
    }

    try {
        const res = await fetch(`${API_URL}/book`, {
            method: "POST",
            headers: { "Content-Type": "application/json", "Authorization": `Bearer ${accessToken}` },
            body: JSON.stringify({
                subject, room_email: roomEmail,
                start_time: new Date(startVal).toISOString(),
                end_time:   new Date(endVal).toISOString(),
                organizer_email: username, attendees: attendeeList, filiale, description: desc
            })
        });
        if (res.ok) {
            closeBookingModal();
            const startFmt = new Date(startVal).toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
            const endFmt   = new Date(endVal).toLocaleTimeString([],   { hour: "2-digit", minute: "2-digit" });
            showBookingSuccess(subject, filiale, `${startFmt} – ${endFmt}`, attendeesRaw);
            loadAvailability(roomEmail);
        } else {
            const err = await res.json().catch(() => ({}));
            showToast(err.detail || "Booking failed.", true);
        }
    } catch (e) { showToast("Network error: " + e.message, true); }
}

// ── FIX-03: escapeHtml on subject, filiale, invitees (XSS fix — VULN-03, lines 727-743) ──
function showBookingSuccess(subject, filiale, timeRange, invitees) {
    const overlay = document.createElement("div");
    overlay.style.cssText = `
        position:fixed;inset:0;z-index:9999;background:rgba(5,20,10,0.97);
        display:flex;flex-direction:column;justify-content:center;align-items:center;
        font-family:'Sora',sans-serif;text-align:center;padding:40px;animation:fadeIn .3s ease;
    `;
    overlay.innerHTML = `
        <div style="font-size:3.5rem;margin-bottom:20px;">✅</div>
        <div style="font-size:1.8rem;font-weight:800;color:#fff;margin-bottom:6px;">Booking Confirmed</div>
        <div style="font-size:0.9rem;color:rgba(255,255,255,.4);margin-bottom:32px;">Added to your Outlook calendar.</div>
        <div style="background:rgba(255,255,255,.06);border:1px solid rgba(255,255,255,.1);border-radius:16px;
             padding:24px 36px;text-align:left;min-width:280px;line-height:2.2;font-size:.9rem;color:rgba(255,255,255,.75);">
            <div><strong style="color:#22c46e;">Subject</strong>&nbsp;&nbsp;${escapeHtml(subject)}</div>
            <div><strong style="color:#22c46e;">Unit</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;${escapeHtml(filiale)}</div>
            <div><strong style="color:#22c46e;">Time</strong>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;${timeRange}</div>
            <div><strong style="color:#22c46e;">Invitees</strong>&nbsp;${escapeHtml(invitees || "None")}</div>
        </div>
        <button id="successClose" style="margin-top:28px;padding:12px 40px;border-radius:999px;
            background:#22c46e;color:#05200e;border:none;font-family:'Sora',sans-serif;
            font-weight:700;font-size:0.95rem;cursor:pointer;">
            OK · Closing in <span id="successCountdown">5</span>s
        </button>
    `;
    document.body.appendChild(overlay);
    let n = 5;
    const iv = setInterval(() => {
        n--;
        const el = document.getElementById("successCountdown");
        if (el) el.textContent = n;
        if (n <= 0) { clearInterval(iv); closeSuccess(); }
    }, 1000);
    const closeSuccess = () => {
        if (document.body.contains(overlay)) document.body.removeChild(overlay);
        signOut();
    };
    document.getElementById("successClose").onclick = () => { clearInterval(iv); closeSuccess(); };
}

// ═══════════════════════════════════════════
//  ROOMS & TIMELINE
// ═══════════════════════════════════════════
async function fetchRooms() {
    try {
        const res  = await fetch(`${API_URL}/rooms`);
        const data = await res.json();
        if (data.value) {
            availableRooms = data.value;
            const select   = document.getElementById("roomSelect");
            select.innerHTML = `<option value="" disabled selected>Select a room…</option>`;
            availableRooms.forEach((r, i) => {
                const opt       = document.createElement("option");
                opt.value       = i;
                opt.textContent = `${r.displayName}  [${r.department} · ${r.floor}]`;
                select.appendChild(opt);
            });
        }
    } catch (e) { console.error(e); }
}

function handleRoomChange() {
    const idx  = document.getElementById("roomSelect").value;
    if (idx === "") return;
    const room = availableRooms[idx];
    const floorEl = document.getElementById("roomFloor");
    const deptEl  = document.getElementById("roomDept");
    const capEl   = document.getElementById("roomCapacity");
    const locEl   = document.getElementById("roomLocation");
    if (floorEl) floorEl.querySelector("span").textContent = room.floor      || "—";
    if (deptEl)  deptEl.querySelector("span").textContent  = room.department || "—";
    if (capEl)   capEl.querySelector("span").textContent   = (room.capacity  || 8) + " persons";
    if (locEl)   locEl.querySelector("span").textContent   = room.location   || "Casablanca HQ";
    lastKnownEventId = "init";
    loadAvailability(room.emailAddress);
    checkForActiveMeeting();
}

async function loadAvailability(email) {
    if (!email) return;
    const spinner  = document.getElementById("loadingSpinner");
    if (spinner) spinner.style.display = "inline";
    const now      = new Date();
    const dayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 0, 0, 0);
    const dayEnd   = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);
    try {
        const res  = await fetch(`${API_URL}/availability`, {
            method: "POST", headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ room_email: email, start_time: dayStart.toISOString(), end_time: dayEnd.toISOString(), time_zone: "UTC" })
        });
        const data        = await res.json();
        const hasMeetings = (data?.value?.[0]?.scheduleItems || []).some(i => i.status === "busy");
        const calBtn      = document.getElementById("roomCalendarBtn");
        if (calBtn) calBtn.style.display = hasMeetings ? "flex" : "none";
    } catch (e) { console.error(e); }
    finally { if (spinner) spinner.style.display = "none"; }
}

// ═══════════════════════════════════════════
//  BOOKING AVAILABILITY STRIP
// ═══════════════════════════════════════════
let stripFetchTimeout = null;
let lastStripDate     = null;
let lastStripData     = null;

function refreshBookingTimeline() {
    const startVal = document.getElementById("startTime").value;
    const endVal   = document.getElementById("endTime").value;
    if (!startVal) return;
    const startDate = new Date(startVal);
    const dateKey   = startDate.toDateString();
    clearTimeout(stripFetchTimeout);
    stripFetchTimeout = setTimeout(async () => {
        if (dateKey !== lastStripDate) { lastStripDate = dateKey; await fetchStripData(startDate); }
        renderStrip(startVal, endVal);
    }, 250);
}

async function fetchStripData(forDate) {
    const idx = document.getElementById("roomSelect").value;
    if (idx === "") return;
    const roomEmail = availableRooms[idx].emailAddress;
    const dayStart  = new Date(forDate.getFullYear(), forDate.getMonth(), forDate.getDate(), 0, 0, 0);
    const dayEnd    = new Date(forDate.getFullYear(), forDate.getMonth(), forDate.getDate(), 23, 59, 59);
    const track     = document.getElementById("availStripTrack");
    if (track) track.innerHTML = `<div class="avail-strip-loading">Loading…</div>`;
    try {
        const res  = await fetch(`${API_URL}/availability`, {
            method: "POST", headers: { "Content-Type": "application/json" },
            body: JSON.stringify({ room_email: roomEmail, start_time: dayStart.toISOString(), end_time: dayEnd.toISOString(), time_zone: "UTC" })
        });
        const data    = await res.json();
        lastStripData = data?.value?.[0]?.scheduleItems || [];
        const dateLabel = document.getElementById("availStripDate");
        const today     = new Date();
        if (dateLabel) {
            if      (forDate.toDateString() === today.toDateString())
                dateLabel.textContent = "Today";
            else if (forDate.toDateString() === new Date(today.getTime() + 86400000).toDateString())
                dateLabel.textContent = "Tomorrow";
            else
                dateLabel.textContent = forDate.toLocaleDateString([], { weekday: "short", month: "short", day: "numeric" });
        }
    } catch { lastStripData = []; }
}

function renderStrip(startVal, endVal) {
    const track    = document.getElementById("availStripTrack");
    const hours    = document.getElementById("availStripHours");
    const conflict = document.getElementById("availConflict");
    if (!track) return;
    const SH    = 7, EH = 22, TOTAL = (EH - SH) * 60;
    const base       = startVal ? new Date(startVal) : new Date();
    const stripStart = new Date(base.getFullYear(), base.getMonth(), base.getDate(), SH, 0, 0);
    track.innerHTML  = "";
    if (hours) {
        hours.innerHTML = "";
        for (let h = SH; h <= EH; h += 2) {
            const pct = ((h - SH) / (EH - SH)) * 100;
            const lbl = document.createElement("span");
            lbl.className = "avail-hour-label"; lbl.style.left = pct + "%";
            lbl.textContent = `${String(h).padStart(2,"0")}:00`;
            hours.appendChild(lbl);
        }
    }
    let hasConflict = false;
    const selStart  = startVal ? new Date(startVal) : null;
    const selEnd    = endVal   ? new Date(endVal)   : null;
    (lastStripData || []).forEach(item => {
        if (item.status !== "busy") return;
        const bs  = new Date(item.start.dateTime + "Z");
        const be  = new Date(item.end.dateTime   + "Z");
        const lMs = Math.max(bs - stripStart, 0);
        const rMs = Math.min(be - stripStart, TOTAL * 60000);
        if (rMs <= 0 || lMs >= TOTAL * 60000) return;
        const block = document.createElement("div");
        block.className   = "avail-busy-block";
        block.style.left  = (lMs / (TOTAL * 60000)) * 100 + "%";
        block.style.width = ((rMs - lMs) / (TOTAL * 60000)) * 100 + "%";
        const st = bs.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
        const et = be.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" });
        block.title = `${item.subject || "Busy"}: ${st} – ${et}`;
        track.appendChild(block);
        if (selStart && selEnd && bs < selEnd && be > selStart) hasConflict = true;
    });
    if (selStart && selEnd && selEnd > selStart) {
        const slMs = Math.max(selStart - stripStart, 0);
        const srMs = Math.min(selEnd   - stripStart, TOTAL * 60000);
        if (srMs > 0 && slMs < TOTAL * 60000) {
            const sel = document.createElement("div");
            sel.className   = "avail-selected-block" + (hasConflict ? " conflict" : "");
            sel.style.left  = (slMs / (TOTAL * 60000)) * 100 + "%";
            sel.style.width = ((srMs - slMs) / (TOTAL * 60000)) * 100 + "%";
            track.appendChild(sel);
        }
    }
    if (conflict) conflict.classList.toggle("hidden", !hasConflict);
}

// ═══════════════════════════════════════════
//  UTILITIES
// ═══════════════════════════════════════════
function initModalTimes() {
    const now    = new Date();
    const offset = now.getTimezoneOffset();
    now.setMinutes(now.getMinutes() - offset);
    const start = document.getElementById("startTime");
    const end   = document.getElementById("endTime");
    if (!start || !end) return;
    start.value = now.toISOString().slice(0, 16);
    now.setMinutes(now.getMinutes() + 30);
    end.value = now.toISOString().slice(0, 16);
}

function showToast(msg, isError = false) {
    const t = document.getElementById("toastBar");
    if (!t) return;
    t.textContent = msg;
    t.className   = "toast-bar" + (isError ? " error" : "");
    t.classList.remove("hidden");
    clearTimeout(t._timeout);
    t._timeout = setTimeout(() => t.classList.add("hidden"), 3500);
}
