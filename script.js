// 🔴 REPLACE THIS WITH YOUR RENDER URL
const API_URL = "https://booking-a-room-poc.onrender.com"; 

const HOURS_TO_SHOW = 12; // Total span to display

document.addEventListener("DOMContentLoaded", () => {
    initModalTimes();
    fetchRooms();
});

// Initialize Date/Time inputs in the Modal
function initModalTimes() {
    const now = new Date();
    now.setMinutes(now.getMinutes() - now.getTimezoneOffset());
    document.getElementById('startTime').value = now.toISOString().slice(0,16);
    
    now.setMinutes(now.getMinutes() + 30);
    document.getElementById('endTime').value = now.toISOString().slice(0,16);
}

// 1. Fetch Rooms List
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

// 2. Load Availability (Visual Timeline)
async function loadAvailability() {
    const roomEmail = document.getElementById('roomSelect').value;
    if (!roomEmail) return;

    document.getElementById('loadingSpinner').style.display = "inline";
    
    // Calculate View Range: Snap to the PREVIOUS hour
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

// 3. Render Timeline (Green Grid + Red Events)
function renderTimeline(data, viewStart, viewEnd) {
    const timelineContainer = document.getElementById('timeline');
    timelineContainer.innerHTML = ''; 

    const totalDurationMs = viewEnd - viewStart;
    
    // We want 30-minute slots instead of 1-hour
    // 12 hours * 2 slots/hour = 24 slots
    const totalSlots = HOURS_TO_SHOW * 2; 
    const slotWidthPct = 100 / totalSlots;

    // A. Header Row (15:00, 15:30...)
    let headerHtml = `<div class="timeline-header">`;
    for (let i = 0; i < totalSlots; i++) {
        let slotTime = new Date(viewStart.getTime() + i * 30 * 60 * 1000);
        let timeStr = slotTime.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
        
        // Only show text for full hours to avoid clutter, or keep small if desired
        // Here we show every 30 mins but you can filter with (i % 2 === 0)
        headerHtml += `<div class="timeline-time-label" style="width:${slotWidthPct}%">${timeStr}</div>`;
    }
    headerHtml += `</div>`;

    // B. Track Row (Green Bg + Red Events)
    let trackHtml = `<div class="timeline-track">`;
    
    // Draw Vertical Grid Lines (every 30 mins)
    for (let i = 1; i < totalSlots; i++) {
        trackHtml += `<div class="grid-line" style="left: ${i * slotWidthPct}%"></div>`;
    }

    // Draw "NOW" Indicator Line
    const now = new Date();
    if (now >= viewStart && now <= viewEnd) {
        const nowPct = ((now - viewStart) / totalDurationMs) * 100;
        trackHtml += `
            <div class="current-time-line" style="left: ${nowPct}%">
                <div class="current-time-label">NOW</div>
            </div>`;
    }

    // Draw Occupied Events (Red Blocks)
    const schedule = (data.value && data.value[0]) ? data.value[0] : null; 
    
    if (schedule && schedule.scheduleItems) {
        schedule.scheduleItems.forEach(item => {
            if (item.status === 'busy') {
                const start = new Date(item.start.dateTime + 'Z');
                const end = new Date(item.end.dateTime + 'Z');
                
                // Calculate Position & Width
                const offsetMs = start - viewStart;
                const durationMs = end - start;
                
                const leftPct = (offsetMs / totalDurationMs) * 100;
                const widthPct = (durationMs / totalDurationMs) * 100;

                // Render if visible
                if (leftPct < 100 && (leftPct + widthPct) > 0) {
                    const safeLeft = Math.max(0, leftPct);
                    const safeWidth = Math.min(widthPct, 100 - safeLeft);
                    
                    trackHtml += `
                        <div class="event-block" 
                             style="left:${safeLeft}%; width:${safeWidth}%;" 
                             title="${item.subject || 'Busy'}">
                             <span>🚫 Occupied</span>
                        </div>`;
                }
            }
        });
    }

    trackHtml += `</div>`; // End track
    timelineContainer.innerHTML = headerHtml + trackHtml;
}

// 4. Create Booking (With Error Handling)
async function createBooking() {
    const roomEmail = document.getElementById('roomSelect').value;
    const organizer = document.getElementById('organizerEmail').value;
    const start = new Date(document.getElementById('startTime').value);
    const end = new Date(document.getElementById('endTime').value);
    const subject = document.getElementById('subject').value;

    if (!roomEmail) return alert("Please select a room first.");
    if (!organizer) return alert("Please enter organizer email.");

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
            alert("✅ Booking Successful!");
            const modalEl = document.getElementById('bookingModal');
            const modal = bootstrap.Modal.getInstance(modalEl);
            modal.hide();
            loadAvailability(); // Refresh timeline
        } else {
            // Handle the conflict error
            const err = await res.json();
            if (res.status === 409) {
                alert("⛔ STOP: " + err.detail); 
            } else {
                alert("❌ Booking Failed: " + JSON.stringify(err));
            }
        }
    } catch (e) {
        alert("Error: " + e.message);
    }
}
