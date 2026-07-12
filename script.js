const milestones = [
  { 
    id: 11, 
    name: "Shire", 
    cx: 21.10, 
    cy: 27.42, 
    distance: 0,
    book: "FOTR, Book I, Chapter 1",
    quote: "He used to say there was only one Road; that it was like a great muddy river, and that its springs were at every doorstep, and every path was its tributary."
  },
  { 
    id: 12, 
    name: "Rivendell", 
    cx: 47.73, 
    cy: 26.82, 
    distance: 458,
    book: "FOTR, Book II, Chapter 1",
    quote: "Rivendell was the perfect house, whether you liked food, or sleep, or work, or storytelling, or singing, or just sitting and thinking best."
  },
  { 
    id: 13, 
    name: "Moria", 
    cx: 45.96, 
    cy: 38.66, 
    distance: 708,
    book: "FOTR, Book II, Chapter 4",
    quote: "The world is grey, the mountains old, the forge's fire is ashen-cold; no harp is wrung, no hammer falls: in Moria, in Khazad-dûm."
  },
  { 
    id: 14, 
    name: "Lothlórien", 
    cx: 52.66, 
    cy: 41.22, 
    distance: 920,
    book: "FOTR, Book II, Chapter 6",
    quote: "It is a land of blossoms and flowers, where the leaves do not fall, but turn to gold. No shadow has yet fallen upon Lórien."
  },
  { 
    id: 15, 
    name: "Falls of Rauros", 
    cx: 62.52, 
    cy: 61.54, 
    distance: 1309,
    book: "FOTR, Book II, Chapter 10",
    quote: "The great river bore them down to the spray-drenched chasm of Rauros, where the waters plunged with a dull roar into the mist."
  },
  { 
    id: 16, 
    name: "Dead Marshes", 
    cx: 66.47, 
    cy: 60.55, 
    distance: 1519,
    book: "TT, Book IV, Chapter 2",
    quote: "Mires and fens where the water is cold and yellow, and there are lights, pale and flickering, like candles in the wind."
  },
  { 
    id: 17, 
    name: "Shelob's Lair", 
    cx: 73.18, 
    cy: 68.44, 
    distance: 1669,
    book: "TT, Book IV, Chapter 9",
    quote: "In that dark place, nothing lived, and no light entered, save for the dreadful eyes of the ancient spider waiting in the deep."
  },
  { 
    id: 18, 
    name: "Mount Doom", 
    cx: 77.71, 
    cy: 65.68, 
    distance: 1779,
    book: "ROTK, Book VI, Chapter 3",
    quote: "The land of Mordor lay before them, and in the midst of it stood the mountain of fire, burning with a red and terrible glow."
  },
  { 
    id: 34, 
    name: "Minas Tirith", 
    cx: 68.49, 
    cy: 70.49, 
    distance: 1779,
    book: "ROTK, Book V, Chapter 1",
    quote: "Seven walls of stone, so strong and old that they seemed to have been carved by giants out of the bones of the earth."
  },
  { 
    id: 35, 
    name: "Isengard", 
    cx: 41.97, 
    cy: 53.92, 
    distance: 2314,
    book: "TT, Book III, Chapter 9",
    quote: "Once it was green and fair, but Saruman broke the trees and dug deep pits, filling the valley with fire and iron."
  },
  { 
    id: 36, 
    name: "Hobbiton", 
    cx: 17.56, 
    cy: 27.81, 
    distance: 3404,
    book: "ROTK, Book VI, Chapter 8",
    quote: "They rode back into the valley of Hobbiton, where the water-mill stood, and the trees were turning green under the autumn sun."
  },
  { 
    id: 37, 
    name: "Grey Havens", 
    cx: 6.42, 
    cy: 26.795, 
    distance: 3659,
    book: "ROTK, Book VI, Chapter 9",
    quote: "And the ship went out into the High Sea and passed into the West, until at last on a night of rain Frodo smelled a sweet fragrance on the air."
  }
];

const STEPS_PER_MILE = 2112;
const TOTAL_JOURNEY_MILES = 3659;

let members = [];
let selectedMemberName = "Team Combined"; // Default selection

// DOM elements
const totalStepsEl = document.getElementById("stat-total-steps");
const totalMilesEl = document.getElementById("stat-total-miles");
const locationEl = document.getElementById("stat-location");
const progressPctEl = document.getElementById("progress-pct");
const progressFillEl = document.getElementById("progress-fill");
const nextMilestoneEl = document.getElementById("next-milestone-name");
const nextMilestoneStepsEl = document.getElementById("next-milestone-steps");
const nextMilestoneMilesEl = document.getElementById("next-milestone-miles");
const timelineEl = document.getElementById("timeline-list");
const mapContent = document.getElementById("map-content");
const tokensContainer = document.getElementById("tokens-container");
const memberSelect = document.getElementById("member-select");
const leaderboardList = document.getElementById("leaderboard-list");

let pzInstance = null;

// Initialize panzoom
function initPanzoom() {
  if (window.panzoom && mapContent) {
    pzInstance = window.panzoom(mapContent, {
      bounds: true,
      maxZoom: 5,
      minZoom: 0.8,
      autocenter: false,
      smoothScroll: false
    });
  }
}

function getSheetId(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return match ? match[1] : null;
}

// Parse Excel ArrayBuffer via SheetJS
function parseExcelData(arrayBuffer) {
  const data = new Uint8Array(arrayBuffer);
  const workbook = XLSX.read(data, { type: 'array' });
  
  // Find member sheet
  const memberSheetName = workbook.SheetNames.find(name => 
    !['raw_data', 'leaderboard', 'team_totals', 'route_pct'].includes(name)
  );
  
  if (!memberSheetName) {
    throw new Error("No member data sheet found in Excel workbook.");
  }
  
  const worksheet = workbook.Sheets[memberSheetName];
  const jsonData = XLSX.utils.sheet_to_json(worksheet);
  
  const colors = [
    "#3b82f6", // Blue
    "#10b981", // Emerald
    "#ec4899", // Pink
    "#8b5cf6", // Violet
    "#f43f5e", // Rose
    "#06b6d4", // Cyan
    "#14b8a6", // Teal
    "#84cc16", // Lime
    "#eab308", // Yellow
    "#a855f7", // Purple
    "#f97316"  // Orange
  ];
  
  const membersList = [];
  jsonData.forEach((row, idx) => {
    const name = row['Name'] ? String(row['Name']).trim() : '';
    if (!name || name.toLowerCase() === 'nan' || name === '') return;
    
    const steps = parseInt(row['Total Steps']) || 0;
    const miles = parseFloat(row['Total Distance (mi)']) || (steps / STEPS_PER_MILE);
    
    const initial = name[0] ? name[0].toUpperCase() : "?";
    const color = colors[membersList.length % colors.length];
    
    membersList.push({
      name: name,
      steps: steps,
      miles: miles,
      initial: initial,
      color: color
    });
  });
  
  return membersList;
}

// Interpolate token coordinates along the path
function getUserCoordinates(totalMiles) {
  if (totalMiles <= 0) return { x: milestones[0].cx, y: milestones[0].cy };
  if (totalMiles >= TOTAL_JOURNEY_MILES) {
    const last = milestones[milestones.length - 1];
    return { x: last.cx, y: last.cy };
  }
  
  // Find active segment
  for (let i = 0; i < milestones.length - 1; i++) {
    const mStart = milestones[i];
    const mEnd = milestones[i+1];
    
    if (totalMiles >= mStart.distance && totalMiles <= mEnd.distance) {
      const distRange = mEnd.distance - mStart.distance;
      if (distRange === 0) {
        return { x: mEnd.cx, y: mEnd.cy };
      }
      const t = (totalMiles - mStart.distance) / distRange;
      const x = mStart.cx + t * (mEnd.cx - mStart.cx);
      const y = mStart.cy + t * (mEnd.cy - mStart.cy);
      return { x, y };
    }
  }
  return { x: milestones[0].cx, y: milestones[0].cy };
}

// Fetch spreadsheet data from Google Sheets directly or fallback to local API
// Helper to process loaded members, inject Fellowship total, sort, and update select dropdown
function processLoadedMembers() {
  // Inject Fellowship (Total) combined tracker if it's not present
  const hasTotal = members.some(m => m.name === "Fellowship (Total)");
  if (!hasTotal && members.length > 0) {
    // Exclude existing total rows if they exist
    const individuals = members.filter(m => m.name !== "Fellowship (Total)");
    const totalSteps = individuals.reduce((sum, m) => sum + m.steps, 0);
    const totalMiles = individuals.reduce((sum, m) => sum + m.miles, 0);
    members.push({
      name: "Fellowship (Total)",
      steps: totalSteps,
      miles: totalMiles,
      initial: "F",
      color: "#ffd700"
    });
  }
  
  // Sort members by miles descending
  members.sort((a, b) => b.miles - a.miles);
  
  // Save selected member before updating dropdown
  const previousSelection = selectedMemberName;
  
  // Populate dropdown selector
  populateDropdown();
  
  if (members.some(m => m.name === previousSelection) || previousSelection === "Team Combined") {
    selectedMemberName = previousSelection;
    memberSelect.value = previousSelection;
  } else {
    selectedMemberName = "Team Combined";
    memberSelect.value = "Team Combined";
  }
}

// Fetch spreadsheet data from local Python server, Netlify proxy, or static fallback
async function fetchTrackerData() {
  try {
    let config = { google_sheet_url: "" };
    try {
      const configRes = await fetch('config.json');
      config = await configRes.json();
    } catch (err) {
      console.log("No local config.json available.");
    }
    
    const sheetUrl = config.google_sheet_url ? config.google_sheet_url.trim() : "";
    let dataLoaded = false;
    
    // Fallback 1: Try local Python server API first (if running on localhost).
    // The Python server handles Google Sheet downloads backend-side, avoiding CORS.
    const isLocal = window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1";
    if (isLocal) {
      try {
        console.log("Running locally. Attempting to fetch from local Python server API...");
        const response = await fetch('/api/data');
        if (response.ok) {
          const data = await response.json();
          if (data && data.members) {
            members = data.members;
            dataLoaded = true;
            console.log("Successfully loaded data from local Python server.");
          }
        }
      } catch (localApiErr) {
        console.log("Local Python server not running or failed. Falling back...", localApiErr);
      }
    }
    
    // Fallback 2: Try direct Google Sheet fetch (via Netlify proxy in production to bypass CORS)
    if (!dataLoaded && sheetUrl) {
      try {
        const sheetId = getSheetId(sheetUrl);
        if (sheetId) {
          // If local, fetching spreadsheet directly in browser will fail due to CORS.
          // In production on Netlify, we route through our /api/sheet/ proxy rule.
          const exportUrl = isLocal 
            ? `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=xlsx`
            : `/api/sheet/${sheetId}`;
            
          console.log("Attempting to fetch Google Sheet data from:", exportUrl);
          const res = await fetch(exportUrl);
          if (res.ok) {
            const buffer = await res.arrayBuffer();
            members = parseExcelData(buffer);
            dataLoaded = true;
            console.log("Successfully loaded Google Sheet data.");
          } else {
            console.log("Failed to fetch Google Sheet (HTTP status " + res.status + ")");
          }
        }
      } catch (sheetFetchErr) {
        console.log("Google Sheet fetch failed (expected locally due to CORS):", sheetFetchErr);
      }
    }
    
    // Fallback 3: Load the static 'Step Tracker.xlsx' file from root folder (failsafe)
    if (!dataLoaded) {
      console.log("Fetching static Step Tracker.xlsx file directly from site root...");
      const res = await fetch('Step Tracker.xlsx');
      if (!res.ok) {
        throw new Error("Could not find local static Step Tracker.xlsx file.");
      }
      const buffer = await res.arrayBuffer();
      members = parseExcelData(buffer);
      dataLoaded = true;
      console.log("Successfully loaded local static Step Tracker.xlsx file.");
    }
    
    processLoadedMembers();
    
    // Calculate total team progress and populate splash screen on first load
    const individuals = members.filter(m => m.name !== "Fellowship (Total)");
    const teamMiles = individuals.reduce((sum, m) => sum + m.miles, 0);
    populateSplashScreen(teamMiles);
    
    updateUI();
  } catch (e) {
    console.error("Failed to load tracker data:", e);
    // Don't show blocking alerts on background sync failures, just log to console.
  }
}

// Helper to map milestone name to local generated asset
function getMilestoneImage(name) {
  const mapping = {
    "Shire": "assets/shire.jpg",
    "Rivendell": "assets/rivendell.jpg",
    "Moria": "assets/moria.jpg",
    "Lothlórien": "assets/lothlorien.jpg",
    "Falls of Rauros": "assets/rauros.jpg",
    "Dead Marshes": "assets/marshes.jpg",
    "Shelob's Lair": "assets/shelob.jpg",
    "Mount Doom": "assets/doom.jpg",
    "Minas Tirith": "assets/tirith.jpg",
    "Isengard": "assets/isengard.jpg",
    "Hobbiton": "assets/hobbiton.jpg",
    "Grey Havens": "assets/havens.jpg"
  };
  return mapping[name] || "assets/shire.jpg";
}

// Show milestone detail modal
function showMilestoneModal(m) {
  const modal = document.getElementById("milestone-modal");
  const modalImg = document.getElementById("modal-img");
  const modalTitle = document.getElementById("modal-title");
  const modalDist = document.getElementById("modal-dist");
  const modalQuote = document.getElementById("modal-quote");
  const modalRef = document.getElementById("modal-ref");
  
  if (modal && modalImg && modalTitle && modalDist && modalQuote && modalRef) {
    modalImg.src = getMilestoneImage(m.name);
    modalTitle.textContent = m.name;
    modalDist.textContent = `${m.distance} miles from Hobbiton`;
    modalQuote.textContent = `"${m.quote}"`;
    modalRef.textContent = m.book;
    
    modal.classList.add("active");
  }
}

// Close milestone detail modal
function closeMilestoneModal() {
  const modal = document.getElementById("milestone-modal");
  if (modal) {
    modal.classList.remove("active");
  }
}

// Helper to bind click and mobile tap events to an element, bypassing panzoom swallowing
function bindClickAndTap(element, callback) {
  // Desktop click listener
  element.addEventListener("click", (e) => {
    // If it's a touch pointer type, ignore standard click to avoid double execution
    if (e.pointerType === "touch" || (e.clientX === 0 && e.clientY === 0 && e.screenX === 0)) {
      return; 
    }
    e.preventDefault();
    e.stopPropagation();
    callback(e);
  });

  // Mobile tap listener (uses short touchstart/touchend delta calculations)
  let touchStartX = 0;
  let touchStartY = 0;
  let touchStartTime = 0;

  element.addEventListener("touchstart", (e) => {
    const touch = e.changedTouches[0];
    touchStartX = touch.clientX;
    touchStartY = touch.clientY;
    touchStartTime = Date.now();
  }, { passive: true });

  element.addEventListener("touchend", (e) => {
    const touch = e.changedTouches[0];
    const dx = touch.clientX - touchStartX;
    const dy = touch.clientY - touchStartY;
    const duration = Date.now() - touchStartTime;

    // Trigger callback if the movement was small (not a pan drag) and quick (a tap)
    if (Math.abs(dx) < 12 && Math.abs(dy) < 12 && duration < 350) {
      e.preventDefault();
      e.stopPropagation();
      callback(e);
    }
  });
}

// Add click listeners to SVG milestone markers on the map
function initMilestoneClickListeners() {
  console.log("Setting up map pin click listeners...");
  const markers = document.querySelectorAll(".milestone-marker");
  markers.forEach(marker => {
    marker.style.cursor = "pointer";
    bindClickAndTap(marker, () => {
      const name = marker.getAttribute("data-name");
      console.log("Map pin tapped/clicked:", name);
      const milestone = milestones.find(m => m.name === name);
      if (milestone) {
        showMilestoneModal(milestone);
      }
    });
  });
}

let splashPopulated = false;

// Populate splash screen card with milestone based on team combined progress
function populateSplashScreen(teamMiles) {
  if (splashPopulated) return;
  
  const splashMiles = document.getElementById("splash-miles");
  const splashImg = document.getElementById("splash-img");
  const splashTitle = document.getElementById("splash-title");
  const splashQuote = document.getElementById("splash-quote");
  const splashRef = document.getElementById("splash-ref");
  
  // Find current landmark of the combined team
  let currentMilestone = milestones[0];
  for (let i = 0; i < milestones.length; i++) {
    if (teamMiles >= milestones[i].distance) {
      currentMilestone = milestones[i];
    }
  }
  
  if (splashMiles && splashImg && splashTitle && splashQuote && splashRef) {
    splashMiles.textContent = teamMiles.toFixed(1);
    splashImg.src = getMilestoneImage(currentMilestone.name);
    splashTitle.textContent = currentMilestone.name;
    splashQuote.textContent = `"${currentMilestone.quote}"`;
    splashRef.textContent = currentMilestone.book;
    splashPopulated = true;
  }
}

// Populate dropdown selector
function populateDropdown() {
  memberSelect.innerHTML = '<option value="Team Combined">Team Combined (Fellowship)</option>';
  members.forEach(m => {
    if (m.name === "Fellowship (Total)") return; // Exclude redundant total option
    const opt = document.createElement("option");
    opt.value = m.name;
    opt.textContent = `${m.name}`;
    memberSelect.appendChild(opt);
  });
}

// Update all UI elements
function updateUI() {
  let activeSteps = 0;
  let activeMiles = 0;
  
  // Calculate team totals by summing individual profiles (excluding Fellowship row)
  const individuals = members.filter(m => m.name !== "Fellowship (Total)");
  const teamSteps = individuals.reduce((sum, m) => sum + m.steps, 0);
  const teamMiles = individuals.reduce((sum, m) => sum + m.miles, 0);

  if (selectedMemberName === "Team Combined") {
    activeSteps = teamSteps;
    activeMiles = teamMiles;
  } else {
    const selected = members.find(m => m.name === selectedMemberName);
    if (selected) {
      activeSteps = selected.steps;
      activeMiles = selected.miles;
    }
  }
  
  // Update stats layout
  totalStepsEl.textContent = activeSteps.toLocaleString();
  totalMilesEl.textContent = activeMiles.toFixed(1) + " mi";
  
  // Overall Progress
  const pct = Math.min(100, (activeMiles / TOTAL_JOURNEY_MILES) * 100);
  progressPctEl.textContent = pct.toFixed(1) + "%";
  progressFillEl.style.width = pct + "%";
  
  // Find current and next milestone
  let currentMilestone = milestones[0];
  let nextMilestone = null;
  
  for (let i = 0; i < milestones.length; i++) {
    if (activeMiles >= milestones[i].distance) {
      currentMilestone = milestones[i];
    }
    if (activeMiles < milestones[i].distance && !nextMilestone) {
      nextMilestone = milestones[i];
    }
  }
  
  locationEl.textContent = currentMilestone.name;
  
  // Next milestone progress
  if (nextMilestone) {
    const milesRemaining = nextMilestone.distance - activeMiles;
    const stepsRemaining = Math.ceil(milesRemaining * STEPS_PER_MILE);
    
    nextMilestoneEl.textContent = nextMilestone.name;
    nextMilestoneStepsEl.textContent = stepsRemaining.toLocaleString() + " steps";
    nextMilestoneMilesEl.textContent = milesRemaining.toFixed(1) + " miles";
  } else {
    nextMilestoneEl.textContent = "Grey Havens Reached!";
    nextMilestoneStepsEl.textContent = "0 steps";
    nextMilestoneMilesEl.textContent = "0.0 miles";
  }
  
  // Update SVG tokens and coordinates
  renderTokens();
  
  // Render active path timeline
  renderTimeline(activeMiles);
  
  // Render sidebar leaderboard
  renderLeaderboard();
}

// Render dynamic map SVG tokens for all members
function renderTokens() {
  tokensContainer.innerHTML = "";
  members.forEach((member, index) => {
    const coords = getUserCoordinates(member.miles);
    
    const isSelected = selectedMemberName === member.name;
    const pulseRing = isSelected ? `
      <circle class="token-pulse" cx="${coords.x}" cy="${coords.y}" r="1.5" fill="none" stroke="${member.color}" stroke-width="0.25"></circle>
    ` : '';
    
    // Add offset clustering for overlapping tokens
    const offsetCoords = { x: coords.x, y: coords.y };
    const overlapping = members.filter((m, i) => i < index && Math.abs(m.miles - member.miles) < 0.1);
    if (overlapping.length > 0) {
      const angle = index * (2 * Math.PI / 8);
      offsetCoords.x += 0.45 * Math.cos(angle);
      offsetCoords.y += 0.45 * Math.sin(angle);
    }
    
    const tokenGroup = document.createElementNS("http://www.w3.org/2000/svg", "g");
    tokenGroup.setAttribute("class", "user-token");
    tokenGroup.style.cursor = "pointer";
    
    tokenGroup.innerHTML = `
      ${pulseRing}
      <circle cx="${offsetCoords.x}" cy="${offsetCoords.y}" r="0.9" fill="${member.color}" stroke="white" stroke-width="0.15" style="transition: all 0.5s ease-out;"></circle>
      <text x="${offsetCoords.x}" y="${offsetCoords.y + 0.3}" font-size="0.85" fill="black" font-weight="bold" text-anchor="middle" style="transition: all 0.5s ease-out; font-family: sans-serif;">${member.initial}</text>
      <title>${member.name}: ${member.miles.toFixed(1)} mi (${member.steps.toLocaleString()} steps)</title>
    `;
    
    // Select member on click or tap
    bindClickAndTap(tokenGroup, () => {
      console.log("Token tapped/clicked:", member.name);
      selectMember(member.name);
    });
    
    tokensContainer.appendChild(tokenGroup);
  });
}

// Render milestones timeline
function renderTimeline(activeMiles) {
  timelineEl.innerHTML = "";
  milestones.forEach(m => {
    const item = document.createElement("div");
    item.className = "timeline-item";
    
    let statusClass = "locked";
    let isCurrentNext = false;
    if (activeMiles >= m.distance) {
      statusClass = "completed";
    } else {
      const isNext = milestones.find(x => activeMiles < x.distance) === m;
      if (isNext) {
        statusClass = "active";
        isCurrentNext = true;
      }
    }
    
    item.classList.add(statusClass);
    
    // Auto-expand only the active milestone on load
    const detailsDisplay = isCurrentNext ? "block" : "none";
    
    item.innerHTML = `
      <div class="timeline-dot"></div>
      <div class="timeline-content" style="cursor: pointer; display: flex; flex-direction: column; align-items: stretch; width: 100%;">
        <div class="timeline-header" style="display: flex; justify-content: space-between; align-items: center; width: 100%;">
          <span class="timeline-name">${m.name}</span>
          <span class="timeline-dist">${m.distance} mi</span>
        </div>
        <div class="timeline-details" style="margin-top: 0.5rem; border-top: 1px solid #333; padding-top: 0.5rem; display: ${detailsDisplay};">
          <div class="timeline-thumb-container" style="width: 100%; height: 110px; border-radius: 6px; overflow: hidden; margin-bottom: 0.5rem; border: 1px solid #444;">
            <img src="${getMilestoneImage(m.name)}" alt="${m.name} thumbnail" style="width: 100%; height: 100%; object-fit: cover;">
          </div>
          <p class="timeline-quote" style="font-style: italic; font-size: 0.75rem; color: var(--color-text-secondary); line-height: 1.35; margin-bottom: 0.35rem; font-family: Georgia, serif;">"${m.quote}"</p>
          <p class="timeline-ref" style="font-size: 0.65rem; color: var(--color-accent); font-weight: 600; text-align: right; text-transform: uppercase; letter-spacing: 0.5px;">${m.book}</p>
        </div>
      </div>
    `;
    
    // Toggle details and center coordinate on click (no modal popup, inline expansion only)
    const contentEl = item.querySelector(".timeline-content");
    contentEl.addEventListener("click", (e) => {
      const detailsEl = contentEl.querySelector(".timeline-details");
      const isExpanded = detailsEl.style.display === "block";
      
      // Toggle details visibility
      detailsEl.style.display = isExpanded ? "none" : "block";
      
      // Zoom map to milestone coordinates
      focusOnCoordinates(m.cx, m.cy);
    });
    
    timelineEl.appendChild(item);
  });
}

// Render sidebar leaderboard of members
function renderLeaderboard() {
  leaderboardList.innerHTML = "";
  
  // Calculate team totals first
  const individuals = members.filter(m => m.name !== "Fellowship (Total)");
  const teamSteps = individuals.reduce((sum, m) => sum + m.steps, 0);
  const teamMiles = individuals.reduce((sum, m) => sum + m.miles, 0);
  const teamPct = ((teamMiles / TOTAL_JOURNEY_MILES) * 100).toFixed(1);
  
  // 1. Add a special "Team Combined" row at the top of the leaderboard
  const teamRow = document.createElement("div");
  teamRow.className = "history-item";
  if (selectedMemberName === "Team Combined") {
    teamRow.style.borderColor = "var(--color-success)";
    teamRow.style.backgroundColor = "rgba(255, 255, 255, 0.05)";
  }
  teamRow.style.cursor = "pointer";
  teamRow.style.borderLeft = "4px solid var(--color-success)";
  teamRow.style.marginBottom = "10px";
  
  teamRow.innerHTML = `
    <div class="history-details" style="flex: 1;">
      <span style="font-weight: 700; display: flex; align-items: center; gap: 8px;">
        <span>👥 Team Combined (Fellowship)</span>
      </span>
      <span class="history-date">${teamSteps.toLocaleString()} steps (Sum of individuals)</span>
    </div>
    <div style="text-align: right;">
      <div style="font-weight: bold; color: var(--color-accent);">${teamMiles.toFixed(1)} mi</div>
      <div style="font-size: 0.7rem; color: var(--color-accent);">${teamPct}% complete</div>
    </div>
  `;
  
  teamRow.addEventListener("click", () => {
    selectMember("Team Combined");
  });
  leaderboardList.appendChild(teamRow);
  
  // 2. Render individual members ranked by miles (excluding total row)
  individuals.forEach((m, idx) => {
    const row = document.createElement("div");
    row.className = "history-item";
    if (selectedMemberName === m.name) {
      row.style.borderColor = m.color;
      row.style.backgroundColor = "rgba(255, 255, 255, 0.05)";
    }
    row.style.cursor = "pointer";
    row.style.borderLeft = `4px solid ${m.color}`;
    
    const progressPct = ((m.miles / TOTAL_JOURNEY_MILES) * 100).toFixed(1);
    
    row.innerHTML = `
      <div class="history-details" style="flex: 1;">
        <span style="font-weight: 600; display: flex; align-items: center; gap: 8px;">
          <span>#${idx + 1} ${m.name}</span>
        </span>
        <span class="history-date">${m.steps.toLocaleString()} steps</span>
      </div>
      <div style="text-align: right;">
        <div style="font-weight: bold; color: var(--color-text-primary);">${m.miles.toFixed(1)} mi</div>
        <div style="font-size: 0.7rem; color: var(--color-accent);">${progressPct}% complete</div>
      </div>
    `;
    
    row.addEventListener("click", () => {
      selectMember(m.name);
    });
    
    leaderboardList.appendChild(row);
  });
}

// Select a member and refresh
function selectMember(name) {
  selectedMemberName = name;
  memberSelect.value = name;
  updateUI();
  
  // Center camera focus on the selected member after manual selection
  let activeMiles = 0;
  if (selectedMemberName === "Team Combined") {
    const individuals = members.filter(m => m.name !== "Fellowship (Total)");
    activeMiles = individuals.reduce((sum, m) => sum + m.miles, 0);
  } else {
    const selected = members.find(m => m.name === selectedMemberName);
    if (selected) activeMiles = selected.miles;
  }
  focusOnSelectedToken(activeMiles);
}

// Focus panzoom on selected token
function focusOnSelectedToken(miles) {
  if (selectedMemberName === "Team Combined") {
    const coords = getUserCoordinates(miles);
    focusOnCoordinates(coords.x, coords.y);
  } else {
    const selected = members.find(m => m.name === selectedMemberName);
    if (selected) {
      const coords = getUserCoordinates(selected.miles);
      focusOnCoordinates(coords.x, coords.y);
    }
  }
}

// Center map completely zoomed out
function centerMapZoomedOut() {
  if (!pzInstance || !mapContent) return;
  
  console.log("Centering map zoomed out...");
  const viewport = document.getElementById("map-viewport").getBoundingClientRect();
  const contentWidth = mapContent.offsetWidth;
  const contentHeight = mapContent.offsetHeight;
  
  // Calculate offsets to center the map at scale 0.95 to fit the viewport nicely
  const scale = 0.95;
  const translateX = (viewport.width - contentWidth * scale) / 2;
  const translateY = (viewport.height - contentHeight * scale) / 2;
  
  pzInstance.zoomAbs(0, 0, scale);
  pzInstance.moveTo(translateX, translateY);
}

// Centralized helper to pan and zoom to coordinates
function focusOnCoordinates(cx, cy) {
  if (!pzInstance || !mapContent) return;
  
  const contentWidth = mapContent.offsetWidth;
  const contentHeight = mapContent.offsetHeight;
  const scale = 2.2;
  
  const pointX = (cx / 100) * contentWidth;
  const pointY = (cy / 100) * contentHeight;
  
  const viewport = document.getElementById("map-viewport").getBoundingClientRect();
  const targetX = viewport.width / 2;
  const targetY = viewport.height / 2;
  
  const translateX = targetX - (pointX * scale);
  const translateY = targetY - (pointY * scale);
  
  pzInstance.zoomAbs(0, 0, scale);
  pzInstance.moveTo(translateX, translateY);
}

// Change selected member listener
memberSelect.addEventListener("change", (e) => {
  selectMember(e.target.value);
});


// Controls binding
document.getElementById("zoom-in").addEventListener("click", () => {
  if (pzInstance) pzInstance.smoothZoom(window.innerWidth / 2, window.innerHeight / 2, 1.4);
});

document.getElementById("zoom-out").addEventListener("click", () => {
  if (pzInstance) pzInstance.smoothZoom(window.innerWidth / 2, window.innerHeight / 2, 0.7);
});

document.getElementById("zoom-reset").addEventListener("click", () => {
  if (pzInstance) {
    pzInstance.zoomAbs(0, 0, 1.0);
    pzInstance.moveTo(0, 0);
  }
});

document.getElementById("zoom-focus").addEventListener("click", () => {
  let activeMiles = 0;
  if (selectedMemberName === "Team Combined") {
    const individuals = members.filter(m => m.name !== "Fellowship (Total)");
    activeMiles = individuals.reduce((sum, m) => sum + m.miles, 0);
  } else {
    const selected = members.find(m => m.name === selectedMemberName);
    if (selected) activeMiles = selected.miles;
  }
  focusOnSelectedToken(activeMiles);
});

// Initialize application and wire up mobile tabs
function initializeApp() {
  console.log("Initializing Walk to Mordor Tracker...");
  initPanzoom();
  fetchTrackerData();
  
  // Fix map centering on initial load & window resize (starts zoomed out)
  const mapImg = document.querySelector(".map-image");
  const defaultCenterMap = () => {
    centerMapZoomedOut();
  };

  if (mapImg) {
    if (mapImg.complete) {
      setTimeout(defaultCenterMap, 300); // Wait for panzoom initialization and layout calculations
    } else {
      mapImg.addEventListener("load", defaultCenterMap);
    }
  }
  window.addEventListener("resize", defaultCenterMap);
  
  // Setup mobile tabs toggling
  const tabMap = document.getElementById("tab-map");
  const tabStats = document.getElementById("tab-stats");
  const appContainer = document.querySelector("main.app-container");
  
  console.log("Mobile Tab Elements Found:", {
    tabMap: !!tabMap,
    tabStats: !!tabStats,
    appContainer: !!appContainer
  });
  
  if (tabMap && tabStats && appContainer) {
    // Ensure default view state is active
    if (!appContainer.classList.contains("show-map") && !appContainer.classList.contains("show-stats")) {
      appContainer.classList.add("show-map");
    }
    
    tabMap.addEventListener("click", (e) => {
      console.log("Map Tab Clicked");
      e.preventDefault();
      
      tabMap.classList.add("active");
      tabStats.classList.remove("active");
      appContainer.classList.add("show-map");
      appContainer.classList.remove("show-stats");
      
      // Auto-recenter map view when shifting back to map tab
      if (pzInstance) {
        setTimeout(() => {
          let activeMiles = 0;
          if (selectedMemberName === "Team Combined") {
            const individuals = members.filter(m => m.name !== "Fellowship (Total)");
            activeMiles = individuals.reduce((sum, m) => sum + m.miles, 0);
          } else {
            const selected = members.find(m => m.name === selectedMemberName);
            if (selected) activeMiles = selected.miles;
          }
          focusOnSelectedToken(activeMiles);
        }, 80);
      }
    });
    
    tabStats.addEventListener("click", (e) => {
      console.log("Stats Tab Clicked");
      e.preventDefault();
      
      tabStats.classList.add("active");
      tabMap.classList.remove("active");
      appContainer.classList.add("show-stats");
      appContainer.classList.remove("show-map");
    });
  }
  
  // Setup modal close listeners
  const modalCloseBtn = document.getElementById("modal-close-btn");
  const modalOverlay = document.getElementById("milestone-modal");
  if (modalCloseBtn) {
    modalCloseBtn.addEventListener("click", closeMilestoneModal);
  }
  if (modalOverlay) {
    modalOverlay.addEventListener("click", (e) => {
      if (e.target === modalOverlay) {
        closeMilestoneModal();
      }
    });
  }
  
  // Setup splash screen enter button listener
  const splashScreen = document.getElementById("splash-screen");
  const splashEnterBtn = document.getElementById("splash-enter-btn");
  if (splashEnterBtn && splashScreen) {
    splashEnterBtn.addEventListener("click", (e) => {
      e.preventDefault();
      console.log("Dismissing splash screen...");
      splashScreen.classList.add("fade-out");
      setTimeout(() => {
        splashScreen.remove();
      }, 500);
    });
  }
  
  // Set up milestone marker click listeners (pins on the map)
  initMilestoneClickListeners();
  
  // Automatically sync/pull spreadsheet data every 60 seconds
  setInterval(fetchTrackerData, 60000);
}

// Run immediately if DOM is already parsed, otherwise wait for event
if (document.readyState === "complete" || document.readyState === "interactive") {
  initializeApp();
} else {
  document.addEventListener("DOMContentLoaded", initializeApp);
}
