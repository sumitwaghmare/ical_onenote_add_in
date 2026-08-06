/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, localStorage, fetch */

import html2canvas from "html2canvas";
import ICAL from "ical.js";

Office.onReady((info) => {
    if (info.host === Office.HostType.OneNote) {
        // Hide sideload message and show app body
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";

        // Initialize default date to today
        const dateInput = document.getElementById("planner-date");
        const today = new Date();
        const yyyy = today.getFullYear();
        const mm = String(today.getMonth() + 1).padStart(2, "0");
        const dd = String(today.getDate()).padStart(2, "0");
        dateInput.value = `${yyyy}-${mm}-${dd}`;

        // Initialize Tab Navigation
        setupTabs();

        // Load saved settings
        loadSettings();

        // Setup Button Listeners
        document.getElementById("btn-generate").onclick = generatePlanner;
        document.getElementById("btn-save-settings").onclick = saveSettings;
    }
});

// Navigation Tab Controller
function setupTabs() {
    const tabPlanner = document.getElementById("tab-btn-planner");
    const tabSettings = document.getElementById("tab-btn-settings");
    const contentPlanner = document.getElementById("tab-content-planner");
    const contentSettings = document.getElementById("tab-content-settings");

    tabPlanner.onclick = () => {
        tabPlanner.classList.add("active");
        tabSettings.classList.remove("active");
        contentPlanner.classList.add("active");
        contentSettings.classList.remove("active");
    };

    tabSettings.onclick = () => {
        tabSettings.classList.add("active");
        tabPlanner.classList.remove("active");
        contentSettings.classList.add("active");
        contentPlanner.classList.remove("active");
    };
}

// Load configurations from localStorage
function loadSettings() {
    const ical = localStorage.getItem("ical_url") || "";
    const todoist = localStorage.getItem("todoist_api") || "";

    document.getElementById("setting-ical").value = ical;
    document.getElementById("setting-todoist").value = todoist;
}

// Save configurations to localStorage
function saveSettings() {
    const ical = document.getElementById("setting-ical").value.trim();
    const todoist = document.getElementById("setting-todoist").value.trim();

    localStorage.setItem("ical_url", ical);
    localStorage.setItem("todoist_api", todoist);

    const statusEl = document.getElementById("settings-status");
    statusEl.style.display = "block";
    setTimeout(() => {
        statusEl.style.display = "none";
    }, 3000);
}

// Logging Helpers for UI
function clearLog() {
    const logEl = document.getElementById("status-log");
    logEl.innerHTML = "";
    document.getElementById("status-container").style.display = "flex";
}

function writeLog(message, type = "info") {
    const logEl = document.getElementById("status-log");
    const time = new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    const item = document.createElement("div");
    item.className = `log-item ${type}`;
    item.innerText = `[${time}] ${message}`;
    logEl.appendChild(item);
    logEl.scrollTop = logEl.scrollHeight;
}

// Helper to Format Month & Days for the mini calendars
function renderMiniCalendar(containerId, titleId, year, month, highlightedDay) {
    const daysContainer = document.getElementById(containerId);
    const titleContainer = document.getElementById(titleId);
    daysContainer.innerHTML = "";

    const monthNames = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ];
    titleContainer.innerText = `${monthNames[month]} ${year}`;

    // Get first day of the month and number of days
    const firstDay = new Date(year, month, 1).getDay();
    const totalDays = new Date(year, month + 1, 0).getDate();

    // Previous month padding
    const prevMonthTotalDays = new Date(year, month, 0).getDate();
    for (let i = firstDay - 1; i >= 0; i--) {
        const dayEl = document.createElement("span");
        dayEl.className = "mini-cal-day other-month";
        dayEl.innerText = prevMonthTotalDays - i;
        daysContainer.appendChild(dayEl);
    }

    // Current month days
    for (let day = 1; day <= totalDays; day++) {
        const dayEl = document.createElement("span");
        dayEl.className = "mini-cal-day";
        dayEl.innerText = day;
        if (day === highlightedDay) {
            dayEl.classList.add("today");
        }
        daysContainer.appendChild(dayEl);
    }

    // Next month padding to fill grid (6 rows * 7 columns = 42 cells total)
    const filledCells = firstDay + totalDays;
    const remaining = 42 - filledCells;
    for (let i = 1; i <= remaining; i++) {
        const dayEl = document.createElement("span");
        dayEl.className = "mini-cal-day other-month";
        dayEl.innerText = i;
        daysContainer.appendChild(dayEl);
    }
}

// Generate the A5 background page
async function generatePlanner() {
    clearLog();
    
    // Toggle Loading Spinners
    const btn = document.getElementById("btn-generate");
    const spinner = btn.querySelector(".btn-spinner");
    btn.disabled = true;
    spinner.style.display = "inline-block";

    try {
        const dateVal = document.getElementById("planner-date").value;
        if (!dateVal) {
            writeLog("Error: Please select a valid date.", "error");
            return;
        }

        const targetDate = new Date(dateVal + "T00:00:00");
        const targetDateStr = dateVal; // YYYY-MM-DD
        
        writeLog(`Initializing generation for: ${targetDate.toDateString()}`, "info");

        const icalUrl = localStorage.getItem("ical_url");
        const todoistToken = localStorage.getItem("todoist_api");

        if (!icalUrl) {
            writeLog("Warning: Outlook iCal URL is not configured in settings. Schedule will be empty.", "error");
        }
        if (!todoistToken) {
            writeLog("Warning: Todoist API token is not configured in settings. Task list will be empty.", "error");
        }

        // 1. Update Date Titles on Template
        const options = { month: 'long', day: 'numeric', year: 'numeric' };
        document.getElementById("tmpl-date").innerText = targetDate.toLocaleDateString("en-US", options);
        document.getElementById("tmpl-day").innerText = targetDate.toLocaleDateString("en-US", { weekday: 'long' });

        // 2. Generate Mini Calendars (Current and Next Month)
        const currentYear = targetDate.getFullYear();
        const currentMonth = targetDate.getMonth();
        const currentDay = targetDate.getDate();

        // Calendar 1 (Current Month): Highlight selected day
        renderMiniCalendar("tmpl-cal1-days", "tmpl-cal1-title", currentYear, currentMonth, currentDay);

        // Calendar 2 (Next Month): No highlighted day (set to -1)
        const nextMonthDate = new Date(currentYear, currentMonth + 1, 1);
        renderMiniCalendar("tmpl-cal2-days", "tmpl-cal2-title", nextMonthDate.getFullYear(), nextMonthDate.getMonth(), -1);

        // Clear previous grid contents
        for (let h = 7; h <= 18; h++) {
            document.getElementById(`event-slot-${h}`).innerHTML = "";
        }
        document.getElementById("tmpl-task-list").innerHTML = "";

        // 3. Fetch and Parse iCal Data
        if (icalUrl) {
            writeLog("Fetching calendar events from Outlook iCal...", "info");
            try {
                // Fetch the iCal file
                const response = await fetch(icalUrl);
                if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
                const icsText = await response.text();

                writeLog("Parsing calendar feed...", "info");
                const jcalData = ICAL.parse(icsText);
                const comp = new ICAL.Component(jcalData);
                const vevents = comp.getAllSubcomponents("vevent");

                const eventsToday = [];
                const rangeStart = ICAL.Time.fromString(targetDateStr + "T00:00:00");
                const rangeEnd = ICAL.Time.fromString(targetDateStr + "T23:59:59");

                vevents.forEach(vevent => {
                    const event = new ICAL.Event(vevent);
                    if (event.isRecurring()) {
                        const iterator = event.iterator();
                        let nextTime;
                        // Limit loop in case of unbounded recurring series
                        let count = 0;
                        while ((nextTime = iterator.next()) && nextTime.compare(rangeEnd) <= 0 && count < 100) {
                            count++;
                            if (nextTime.compare(rangeStart) >= 0) {
                                const duration = event.duration;
                                const start = nextTime.clone();
                                const end = nextTime.clone();
                                end.addDuration(duration);
                                eventsToday.push({
                                    summary: event.summary,
                                    start: start.toJSDate(),
                                    end: end.toJSDate()
                                });
                            }
                        }
                    } else {
                        const start = event.startDate;
                        const end = event.endDate;
                        if (start.toISODateString() === targetDateStr) {
                            eventsToday.push({
                                    summary: event.summary,
                                    start: start.toJSDate(),
                                    end: end.toJSDate()
                            });
                        }
                    }
                });

                writeLog(`Found ${eventsToday.length} events for this day.`, "success");

                // Overlay events onto Schedule Slots
                eventsToday.forEach(evt => {
                    const startHour = evt.start.getHours();
                    const startMin = evt.start.getMinutes().toString().padStart(2, "0");
                    const endHour = evt.end.getHours();
                    const endMin = evt.end.getMinutes().toString().padStart(2, "0");

                    // Only map hours within 7 AM to 6 PM (7 to 18)
                    if (startHour >= 7 && startHour <= 18) {
                        const slot = document.getElementById(`event-slot-${startHour}`);
                        const eventEl = document.createElement("div");
                        eventEl.className = "schedule-event";
                        eventEl.innerText = `${startHour}:${startMin} - ${endHour}:${endMin}: ${evt.summary}`;
                        slot.appendChild(eventEl);
                    }
                });

            } catch (err) {
                writeLog(`iCal Fetch/Parse Error: ${err.message}. Generating template with empty schedule.`, "error");
            }
        }

        // 4. Fetch Todoist Tasks
        if (todoistToken) {
            writeLog("Fetching active tasks from Todoist...", "info");
            try {
                const response = await fetch("https://api.todoist.com/rest/v2/tasks", {
                    headers: {
                        "Authorization": `Bearer ${todoistToken}`
                    }
                });
                if (!response.ok) throw new Error(`HTTP error! status: ${response.status}`);
                const tasks = await response.json();
                writeLog(`Fetched ${tasks.length} total tasks. Filtering...`, "info");

                // Filter tasks: incomplete and due today or overdue (due <= targetDateStr)
                const plannerTasks = tasks.filter(t => {
                    return !t.checked && t.due && t.due.date <= targetDateStr;
                });

                // Sort by due date ascending
                plannerTasks.sort((a, b) => a.due.date.localeCompare(b.due.date));

                writeLog(`Found ${plannerTasks.length} tasks due today or overdue.`, "success");

                // Populate template (limit to 8 items to fit nicely in the A5 box)
                const taskContainer = document.getElementById("tmpl-task-list");
                const itemsToShow = plannerTasks.slice(0, 8);
                
                itemsToShow.forEach(task => {
                    const item = document.createElement("div");
                    item.className = "task-item";

                    const checkbox = document.createElement("div");
                    checkbox.className = "task-checkbox";

                    const text = document.createElement("div");
                    text.className = "task-text";
                    
                    // Format due date in short form like (7/17)
                    const dueDate = new Date(task.due.date + "T00:00:00");
                    const dateStr = `(${dueDate.getMonth() + 1}/${dueDate.getDate()})`;

                    text.innerText = `${task.content} ${dateStr}`;

                    item.appendChild(checkbox);
                    item.appendChild(text);
                    taskContainer.appendChild(item);
                });

                if (plannerTasks.length > 8) {
                    writeLog(`Daily Task List truncated to fit A5 page (+${plannerTasks.length - 8} more).`, "info");
                }

            } catch (err) {
                writeLog(`Todoist API Error: ${err.message}. Generating template with empty task list.`, "error");
            }
        }

        // Update footer generation timestamp
        const now = new Date();
        const footerTime = now.toLocaleDateString([], { month: 'numeric', day: 'numeric', year: 'numeric' }) + ", " + now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        document.getElementById("tmpl-footer-timestamp").innerText = footerTime;

        // 5. Render HTML to Canvas Image
        writeLog("Rendering premium background image...", "info");
        const element = document.getElementById("planner-template");
        
        // Wait a small bit for browser layout updates
        await new Promise(resolve => setTimeout(resolve, 300));

        const canvas = await html2canvas(element, {
            scale: 2, // Double scale for crisp text resolution
            logging: false,
            useCORS: true,
            backgroundColor: "#ffffff"
        });

        writeLog("Exporting to image format...", "info");
        const dataUrl = canvas.toDataURL("image/png");
        const base64Data = dataUrl.split(",")[1];

        // 6. Insert Image into Active Page via OfficeJS
        writeLog("Inserting background image into OneNote...", "info");
        
        await OneNote.run(async (context) => {
            const page = context.application.getActivePage();
            
            // Add outline at top-left of page
            const outline = page.addOutline(0, 0);
            
            // Append HTML with image
            const imgHtml = `<img src="data:image/png;base64,${base64Data}" width="595" height="842" />`;
            outline.appendHtml(imgHtml);
            
            await context.sync();
            writeLog("Planner background inserted successfully!", "success");
            writeLog("Remember to right-click the image and select 'Set Picture as Background' to lock it.", "success");
        });

    } catch (err) {
        writeLog(`System Failure: ${err.message}`, "error");
        console.error(err);
    } finally {
        // Toggle Loading Spinners Off
        btn.disabled = false;
        spinner.style.display = "none";
    }
}
