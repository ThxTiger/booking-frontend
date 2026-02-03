// Keep your existing API_URL and fetchRooms() code...
// ONLY REPLACE THE TIMELINE FUNCTIONS BELOW:

async function loadAvailability() {
    const roomEmail = document.getElementById('roomSelect').value;
    if (!roomEmail) return;

    document.getElementById('loadingSpinner').style.display = "inline";
    
    // 1. Calculate View Range: Snap to the previous Hour
    // Example: If it's 15:19, view starts at 15:00
    const now = new Date();
    const viewStart = new Date(now);
    viewStart.setMinutes(0, 0, 0); // Snap to top of hour
    
    const viewEnd = new Date(viewStart.getTime() + 8 * 60 * 60 * 1000); // +8 Hours

    try {
        const res = await fetch(`${API_URL}/availability`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                room_email: roomEmail,
                start_time: viewStart.toISOString(),
                end_time: viewEnd.toISOString(),
                time_zone: "UTC" // Force UTC to keep backend math simple
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

function renderTimeline(data, viewStart, viewEnd) {
    const timelineContainer = document.getElementById('timeline');
    timelineContainer.innerHTML = ''; // Clear previous

    const totalDurationMs = viewEnd - viewStart;

    // --- A. Build Header (15:00, 16:00...) ---
    let headerHtml = `<div class="timeline-header">`;
    for (let i = 0; i < 8; i++) {
        let hourDate = new Date(viewStart.getTime() + i * 60 * 60 * 1000);
        let timeLabel = hourDate.toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
        headerHtml += `<div class="timeline-hour-label">${timeLabel}</div>`;
    }
    headerHtml += `</div>`;

    // --- B. Build Track (Grid Lines + Events) ---
    let trackHtml = `<div class="timeline-track">`;
    
    // 1. Draw Vertical Grid Lines (Visual separators for hours)
    for (let i = 1; i < 8; i++) {
        let leftPct = (i / 8) * 100;
        trackHtml += `<div class="grid-line" style="left: ${leftPct}%"></div>`;
    }

    // 2. Draw "Current Time" Line
    const now = new Date();
    if (now >= viewStart && now <= viewEnd) {
        const nowOffset = now - viewStart;
        const nowPct = (nowOffset / totalDurationMs) * 100;
        trackHtml += `
            <div class="current-time-line" style="left: ${nowPct}%">
                <div class="current-time-label">Now</div>
            </div>`;
    }

    // 3. Draw Events
    const schedule = (data.value && data.value[0]) ? data.value[0] : null; 
    
    if (schedule && schedule.scheduleItems) {
        schedule.scheduleItems.forEach(item => {
            if (item.status === 'busy') {
                const eventStart = new Date(item.start.dateTime + 'Z'); 
                const eventEnd = new Date(item.end.dateTime + 'Z');

                // Math to position the block
                const offsetMs = eventStart - viewStart;
                const durationMs = eventEnd - eventStart;
                
                // Calculate Percentages
                const leftPct = (offsetMs / totalDurationMs) * 100;
                const widthPct = (durationMs / totalDurationMs) * 100;
                
                // Only render if it's visible or partially visible
                if (leftPct < 100 && (leftPct + widthPct) > 0) {
                    // Clip visuals if sticking out
                    const safeLeft = Math.max(0, leftPct);
                    const safeWidth = Math.min(widthPct, 100 - safeLeft);

                    trackHtml += `
                        <div class="event-block" 
                             style="left:${safeLeft}%; width:${safeWidth}%;" 
                             title="${item.subject || 'Busy'} • ${eventStart.toLocaleTimeString([], {hour:'2-digit', minute:'2-digit'})}">
                             ${item.subject || 'Busy'}
                        </div>`;
                }
            }
        });
    }

    trackHtml += `</div>`; // Close track
    timelineContainer.innerHTML = headerHtml + trackHtml;
}
