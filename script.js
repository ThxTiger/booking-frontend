// 🔴 REPLACE THIS WITH YOUR RENDER URL
const API_URL = "https://booking-a-room-poc.onrender.com"; 

// --- Configuration ---
const HOURS_TO_SHOW = 12; // Use 12 hours instead of 8

document.addEventListener("DOMContentLoaded", () => {
    initModalTimes();
    fetchRooms();
});

function initModalTimes() {
    const now = new Date();
    now.setMinutes(now.getMinutes() - now.getTimezoneOffset());
    document.getElementById('startTime').value = now.toISOString().slice(0,16);
    
    now.setMinutes(now.getMinutes() + 30);
    document.getElementById('endTime').value = now.toISOString().slice(0,16);
}

// --- 1. Fetch Rooms ---
async function fetchRooms() {
    try {
        const res = await fetch(`${API_URL}/rooms`);
        const data = await res.json();
        const select = document.getElementById('roomSelect');
        
        select.innerHTML = '<option value="" disabled selected>Select a room...</option>';
        
        if (data.value && data.value.length > 0) {
            data.value.forEach(room => {
                const opt = document.createElement('option');
                opt.value = room.emailAddress;
                opt.textContent = room.displayName;
                select.appendChild(opt);
            });
        }
    } catch (err) {
        console.error("Error fetching rooms:", err);
    }
}

// --- 2. Load Availability ---
async function loadAvailability() {
    const roomEmail = document.getElementById('roomSelect').value;
    if (!roomEmail) return;

    document.getElementById('loadingSpinner').style.display = "inline";
    
    // Calculate View Range: Snap to the PREVIOUS hour (e.g., 15:23 -> 15:00)
    const now = new Date();
    const viewStart = new Date(now);
    viewStart.setMinutes(0, 0, 0); 
    
    const viewEnd = new Date(viewStart.getTime() + HOURS_TO_SHOW * 60 * 60 * 1000);

    try {
        const res = await fetch(`${API_URL}/availability`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                room_email: roomEmail,
                start_time: viewStart.toISOString(),
                end_time: viewEnd.toISOString(),
                time_zone: "UTC"
            })
        });

        const data = await res.json();
        renderTimeline(data, viewStart, viewEnd);
    } catch (err) {
        document.getElementById('timeline').innerHTML = `<p class="text-danger p-3">Error: ${err.message}</p>`;
    } finally {
        document.getElementById('loadingSpinner').style.display = "none";
    }
}

// --- 3. Render Professional Gantt Timeline ---
function renderTimeline(data, viewStart, viewEnd) {
    const timelineContainer = document.getElementById('timeline');
    timelineContainer.innerHTML = ''; 

    const totalDurationMs = viewEnd - viewStart;

    // A. Header Row (e.g. 08:00, 09:00)
    let headerHtml = `<div class="timeline-header">`;
    for (let i = 0; i < HOURS_TO_SHOW; i++) {
        let hourDate = new Date(viewStart.getTime() + i * 60 * 60 * 1000);
        let timeLabel = hourDate.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
        headerHtml += `<div class="timeline-hour-label">${timeLabel}</div>`;
    }
    headerHtml += `</div>`;

    // B. Track Row (Grid + Events)
    let trackHtml = `<div class="timeline-track">`;
    
    // Draw Vertical Grid Lines
    for (let i = 1; i < HOURS_TO_SHOW; i++) {
        let leftPct = (i / HOURS_TO_SHOW) * 100;
        trackHtml += `<div class="grid-line" style="left: ${leftPct}%"></div>`;
    }

    // Draw "NOW" Indicator
    const now = new Date();
    if (now >= viewStart && now <= viewEnd) {
        const nowOffset = now - viewStart;
        const nowPct = (nowOffset / totalDurationMs) * 100;
        trackHtml += `
            <div class="current-time-line" style="left: ${nowPct}%">
                <div class="current-time-label">NOW</div>
            </div>`;
    }

    // Draw Events
    const schedule = (data.value && data.value[0]) ? data.value[0] : null; 
    
    if (schedule && schedule.scheduleItems) {
        schedule.scheduleItems.forEach(item => {
            if (item.status === 'busy') {
                const eventStart = new Date(item.start.dateTime + 'Z'); 
                const eventEnd = new Date(item.end.dateTime + 'Z');

                // Math
                const offsetMs = eventStart - viewStart;
                const durationMs = eventEnd - eventStart;
                
                // % Calculation
                const leftPct = (offsetMs / totalDurationMs) * 100;
                const widthPct = (durationMs / totalDurationMs) * 100;
                
                // Render if visible
                if (leftPct < 100 && (leftPct + widthPct) > 0) {
                    const safeLeft = Math.max(0, leftPct);
                    const safeWidth = Math.min(widthPct, 100 - safeLeft);

                    // Create nice label text
                    const startStr = eventStart.toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'});
                    const endStr = eventEnd.toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'});

                    trackHtml += `
                        <div class="event-block" 
                             style="left:${safeLeft}%; width:${safeWidth}%;" 
                             title="${item.subject || 'Busy'} (${startStr} - ${endStr})">
                             <span class="event-title">${item.subject || 'Busy'}</span>
                             <span class="event-time">${startStr} - ${endStr}</span>
                        </div>`;
                }
            }
        });
    }

    trackHtml += `</div>`; // End track
    timelineContainer.innerHTML = headerHtml + trackHtml;
}

// --- 4. Create Booking ---
async function createBooking() {
    const roomEmail = document.getElementById('roomSelect').value;
    const organizer = document.getElementById('organizerEmail').value;
    const start = new Date(document.getElementById('startTime').value);
    const end = new Date(document.getElementById('endTime').value);
    const subject = document.getElementById('subject').value;

    if (!roomEmail) return alert("Select a room first.");
    if (!organizer) return alert("Enter organizer email.");

    try {
        const res = await fetch(`${API_URL}/book`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                subject: subject,
                room_email: roomEmail,
                start_time: start.toISOString(),
                end_time: end.toISOString(),
                organizer_email: organizer
            })
        });
        
        if (res.ok) {
            alert("✅ Booking Created!");
            bootstrap.Modal.getInstance(document.getElementById('bookingModal')).hide();
            loadAvailability(); 
        } else {
            const err = await res.json();
            alert("❌ Failed: " + JSON.stringify(err));
        }
    } catch (e) {
        alert("Error: " + e.message);
    }
}
