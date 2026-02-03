// ======================================
// 1. CONFIG
// ======================================

// Carousels (SAME engine/behavior) — gallery + reviews
const CAROUSEL_CONFIGS = [
  {
    name: "imageGallery",
    rootSelector: '[data-carousel="gallery"]',
    trackSelector: ".gallery-track",
    slideSelector: ".gallery-image",
    dotSelector: ".subway-dot",
    blocksEachSide: 5,
    dotModulo: true,
  },
];

const SIMPLE_SCROLL_GALLERIES = [
  {
    name: "reviewCarousel",
    rootSelector: '[data-carousel="reviews"]',
    trackSelector: ".reviews-track",
    slideSelector: ".review-card",
    dotSelector: ".subway-dot",
  },
];


function initSimpleScrollGallery(config) {
  const { rootSelector, trackSelector, slideSelector, dotSelector } = config;

  const root = rootSelector ? document.querySelector(rootSelector) : document;
  if (!root) return;

  const track = root.querySelector(trackSelector);
  if (!track) return;

  const dots = Array.from(root.querySelectorAll(dotSelector));
  const slides = Array.from(track.querySelectorAll(slideSelector)).filter(
    (el) => !el.classList.contains("clone")
  );
  if (!slides.length) return;

  function getClosestSlideIndex() {
    const denom = track.scrollWidth - track.clientWidth;
    if (denom <= 1) return 0;
    const t = Math.min(1, Math.max(0, track.scrollLeft / denom));
    return Math.round(t * (slides.length - 1));
  }

  function setActiveDot(idx) {
    if (!dots.length) return;
    dots.forEach((d, i) => d.classList.toggle("active", i === idx));
  }

  track.addEventListener(
    "scroll",
    () => requestAnimationFrame(() => setActiveDot(getClosestSlideIndex())),
    { passive: true }
  );

  dots.forEach((dot, idx) => {
    dot.addEventListener("click", () => {
      track.scrollTo({
  left: slides[idx].offsetLeft - (track.clientWidth - slides[idx].clientWidth) / 2,
  behavior: "smooth",
});
      setActiveDot(idx);
    });
  });

  setActiveDot(0);
  // Center first review on load (NO vertical page scroll)
requestAnimationFrame(() => {
  track.scrollTo({
    left: slides[0].offsetLeft - (track.clientWidth - slides[0].clientWidth) / 2,
    behavior: "auto",
  });
});
}




// ======================================
// 2. HELPERS – CAROUSELS
// ======================================

/**
 * Infinite-style horizontal carousel with clones:
 *  - clones N full blocks of slides before and after originals
 *  - keeps active dot in sync
 *  - centers clicked slide
 */
function initInfiniteCarousel(config) {
  const {
    rootSelector,
    trackSelector,
    slideSelector,
    dotSelector,
    blocksEachSide = 3,
    dotModulo = false,
  } = config;

  const root = rootSelector ? document.querySelector(rootSelector) : document;
  if (!root) return;

  const track = root.querySelector(trackSelector);
  if (!track) return;

  // IMPORTANT: scope dots to THIS carousel only
  const dots = Array.from(root.querySelectorAll(dotSelector));

  // Originals only (ignore any previous clones)
  const originals = Array.from(track.querySelectorAll(slideSelector)).filter(
    (el) => !el.classList.contains("clone")
  );
  const N = originals.length;
  if (N === 0) return;

  // ---- clone blocks before/after ----
  const fragStart = document.createDocumentFragment();
  const fragEnd = document.createDocumentFragment();

  for (let b = 0; b < blocksEachSide; b++) {
    originals.forEach((card) => {
      const cloneBefore = card.cloneNode(true);
      cloneBefore.classList.add("clone");
      fragStart.appendChild(cloneBefore);
    });

    originals.forEach((card) => {
      const cloneAfter = card.cloneNode(true);
      cloneAfter.classList.add("clone");
      fragEnd.appendChild(cloneAfter);
    });
  }

  track.insertBefore(fragStart, track.firstChild);
  track.appendChild(fragEnd);

  const slides = Array.from(track.querySelectorAll(slideSelector));
  const offsetStart = blocksEachSide * N;

  const physicalToLogical = (p) => {
    const x = (p - offsetStart) % N;
    return (x + N) % N;
  };

  const targetScrollForSlide = (idx) => {
    const s = slides[idx];
    const left = s.offsetLeft;
    const centerOffset = (track.clientWidth - s.clientWidth) / 2;
    return left - centerOffset;
  };

  const getClosestSlideIndex = () => {
    const center = track.scrollLeft + track.clientWidth / 2;

    let best = 0;
    let bestDist = Infinity;

    slides.forEach((s, i) => {
      const c = s.offsetLeft + s.clientWidth / 2;
      const d = Math.abs(c - center);
      if (d < bestDist) {
        bestDist = d;
        best = i;
      }
    });

    return best;
  };

  const setActiveDotFromPhysicalIndex = (physicalIndex) => {
    if (!dots.length) return;
    const logicalIndex = physicalToLogical(physicalIndex);
    const activeIdx = dotModulo ? (logicalIndex % dots.length) : logicalIndex;
    dots.forEach((d, i) => d.classList.toggle("active", i === activeIdx));
  };

  let rafPending = false;
  const updateActiveDot = () => {
    if (rafPending) return;
    rafPending = true;
    requestAnimationFrame(() => {
      rafPending = false;
      const idx = getClosestSlideIndex();
      setActiveDotFromPhysicalIndex(idx);
    });
  };

  // ---- initial position: center first original ----
  track.scrollLeft = targetScrollForSlide(offsetStart);
  setActiveDotFromPhysicalIndex(offsetStart);

  // ---- keep position inside the middle clone region ----
  const maybeRecenter = () => {
    const idx = getClosestSlideIndex();
    const logical = physicalToLogical(idx);
    const desiredPhysical = offsetStart + logical;

    // If we drifted far into clones, snap back to the middle copy (same logical slide)
    if (Math.abs(idx - desiredPhysical) > N) {
      track.scrollLeft = targetScrollForSlide(desiredPhysical);
      setActiveDotFromPhysicalIndex(desiredPhysical);
    }
  };

  track.addEventListener(
    "scroll",
    () => {
      updateActiveDot();
      maybeRecenter();
    },
    { passive: true }
  );

  // ---- dot click -> go to nearest matching slide in the middle region ----
  dots.forEach((dot, dotIndex) => {
    dot.addEventListener("click", () => {
      const currentPhysical = getClosestSlideIndex();
      const currentLogical = physicalToLogical(currentPhysical);

      const D = dots.length || 1;
      let targetLogical;

      if (!dotModulo) {
        targetLogical = dotIndex;
      } else {
        const curMod = currentLogical % D;
        const forward = (dotIndex - curMod + D) % D;
        const backward = (curMod - dotIndex + D) % D;
        const step = backward < forward ? -backward : forward; // tie -> forward
        targetLogical = (currentLogical + step + N) % N;
      }

      const targetPhysical = offsetStart + targetLogical;
      track.scrollLeft = targetScrollForSlide(targetPhysical);
      setActiveDotFromPhysicalIndex(targetPhysical);
    });
  });
}



// (removed) initSimpleScrollGallery — reviews now use the same infinite carousel engine





// ======================================
// 3. HELPERS – NAV TABS / HOVER
// ======================================

/**
 * Highlight active nav tab based on current page.
 */
function initNavActiveTab() {
  const currentPage = window.location.pathname.split("/").pop() || "home.html";

  document.querySelectorAll(".nav-tabs .tab").forEach((tab) => {
    const tabPage = tab.getAttribute("href");
    if (tabPage === currentPage) {
      tab.classList.add("active");
    } else {
      tab.classList.remove("active");
    }
  });
}




// ======================================
// 4. INITIALIZATION
// ======================================

document.addEventListener("DOMContentLoaded", () => {

    // --- iOS/Instagram in-app browser scroll failsafe ---
  (function ensureScrollablePage(){
    const ua = navigator.userAgent || "";
    const isIG = /Instagram/i.test(ua);
    const isIOS = /iPhone|iPad|iPod/i.test(ua);
    if (!(isIG && isIOS)) return;

    const html = document.documentElement;
    const body = document.body;

    html.style.overflowY = "auto";
    html.style.overflowX = "hidden";
    html.style.height = "auto";

    body.style.overflowY = "auto";
    body.style.overflowX = "hidden";
    body.style.height = "auto";
    body.style.position = "relative";
  })();

  // Nav tabs active state
  initNavActiveTab();

  // Header: hide on scroll down, show on scroll up
  (function initHeaderHideOnScroll(){
    const header = document.querySelector(".top-banner");
    if (!header) return;

    let lastY = window.scrollY || 0;
    let lastToggleY = lastY;
    const threshold = parseInt(getComputedStyle(document.documentElement).getPropertyValue("--header-hide-threshold"), 10) || 10;

    window.addEventListener("scroll", () => {
      const y = window.scrollY || 0;
      const dy = y - lastY;

      // ignore tiny jitter
      if (Math.abs(y - lastToggleY) < threshold) {
        lastY = y;
        return;
      }

      if (dy > 0) header.classList.add("is-hidden");    // scrolling down
      else header.classList.remove("is-hidden");        // scrolling up

      lastToggleY = y;
      lastY = y;
    }, { passive: true });
  })();

  // Defer non-critical UI work (improves first paint in IG in-app browser)
  const defer = (fn) => {
    if ("requestIdleCallback" in window) requestIdleCallback(fn, { timeout: 1200 });
    else setTimeout(fn, 200);
  };

  // Infinite carousels (gallery)
  defer(() => {
    CAROUSEL_CONFIGS.forEach((cfg) => initInfiniteCarousel(cfg));
  });

  // Simple scroll galleries (reviews)
  defer(() => {
    SIMPLE_SCROLL_GALLERIES.forEach((cfg) => initSimpleScrollGallery(cfg));
  });

  

  // Itinerary: show first 6 stops, expand/collapse on button
  (function initItineraryToggle(){
    const list = document.querySelector(".stop-list");
    const btn  = document.querySelector(".itinerary-toggle");
    if (!list || !btn) return;

    // Default collapsed (CSS hides items 7+)
    list.classList.add("is-collapsed");
    btn.setAttribute("aria-expanded", "false");
    btn.textContent = "Show full itinerary";

    btn.addEventListener("click", () => {
      const expanded = btn.getAttribute("aria-expanded") === "true";
      btn.setAttribute("aria-expanded", String(!expanded));
      list.classList.toggle("is-collapsed", expanded);
      btn.textContent = expanded ? "Show full itinerary" : "Show less";
    });
  })();

});



// ======================================
// X. Booking List Updater in Google Sheets
// ======================================

// ======================================
// X. Booking + Availability Calendar
// ======================================
document.addEventListener("DOMContentLoaded", () => {
  const SCRIPT_URL =
    "https://script.google.com/macros/s/AKfycbynjG8HWWc7i9g7ZmL_jPBvUof4ZfGEZqxIRKUWQd0MUO5ImWw1jJAxW_e6t6Mydc7n/exec";

  // Availability endpoint (expects ?ym=YYYY-MM and returns { ym, days: { "YYYY-MM-DD": "AVAILABLE"/"CLOSED"/... } })
  const AVAILABILITY_URL = 
    "https://script.google.com/macros/s/AKfycbwiMTeUI77O0rqRKTLaCZowBNpRzQCbA3GXLNS-KCNGRj440HrSInMxwCEA7eh0BDqW/exec";

  const form = document.getElementById("booking-form");
  const messageEl = document.getElementById("booking-message");
  const submitBtn = form ? form.querySelector(".booking-submit") : null;

  // Calendar elements
  const calGrid = document.getElementById("rr-cal-grid");
  const calTitle = document.getElementById("rr-cal-title");
  const calPrev = document.getElementById("rr-cal-prev");
  const calNext = document.getElementById("rr-cal-next");
  const calSelected = document.getElementById("rr-cal-selected");
  const tourDateHidden = document.getElementById("tour_date");

  if (!form || !submitBtn) return;
  if (!calGrid || !calTitle || !calPrev || !calNext || !calSelected || !tourDateHidden) return;


  // --------- calendar state ----------
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let viewYear = today.getFullYear();
  let viewMonth = today.getMonth(); // 0-11
  let selectedISO = "";
  let monthAvailability = {}; // { "YYYY-MM-DD": "AVAILABLE"/"CLOSED"/... }

  function pad2(n) { return String(n).padStart(2, "0"); }
  function monthKey(y, m) { return `${y}-${pad2(m + 1)}`; }
  function isoDate(y, m, d) { return `${y}-${pad2(m + 1)}-${pad2(d)}`; }

  function monthName(y, m) {
    return new Date(y, m, 1).toLocaleString(undefined, { month: "long", year: "numeric" });
  }

  // Monday-first index: 0..6 where 0=Mon
  function mondayFirstIndex(date) {
    const js = date.getDay(); // 0=Sun..6=Sat
    return (js + 6) % 7;
  }

  async function fetchAvailabilityForMonth(ym) {
    if (!AVAILABILITY_URL) return {};
    try {
      const res = await fetch(`${AVAILABILITY_URL}?ym=${encodeURIComponent(ym)}`);
      const json = await res.json();
      return (json && json.days && typeof json.days === "object") ? json.days : {};
    } catch (e) {
      console.warn("Availability fetch failed:", e);
      return {};
    }
  }

  function applyDayState(btn, iso) {
    const d = new Date(iso);
    d.setHours(0, 0, 0, 0);
    const isPast = d < today;

    btn.classList.toggle("is-past", isPast);
    btn.disabled = isPast;

    const raw = (monthAvailability[iso] || "AVAILABLE").toString().trim().toUpperCase();
    const isClosed = raw.startsWith("CLOSED"); // CLOSED, CLOSED_FULL, CLOSED_EMPTY, etc.

    btn.classList.toggle("is-open", !isClosed);
    btn.classList.toggle("is-closed", isClosed);
  }

  async function renderCalendar() {
    if (!calGrid || !calTitle) return;

    const key = monthKey(viewYear, viewMonth);
    calTitle.textContent = monthName(viewYear, viewMonth);

    monthAvailability = await fetchAvailabilityForMonth(key);

    calGrid.innerHTML = "";

    const first = new Date(viewYear, viewMonth, 1);
    const firstIndex = mondayFirstIndex(first);
    const daysInMonth = new Date(viewYear, viewMonth + 1, 0).getDate();

    // blanks
    for (let i = 0; i < firstIndex; i++) {
      const blank = document.createElement("button");
      blank.type = "button";
      blank.className = "rr-cal-day is-blank";
      blank.tabIndex = -1;
      calGrid.appendChild(blank);
    }

    // days
    for (let day = 1; day <= daysInMonth; day++) {
      const iso = isoDate(viewYear, viewMonth, day);

      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "rr-cal-day";
      btn.textContent = String(day);
      btn.setAttribute("data-date", iso);

      applyDayState(btn, iso);

      btn.addEventListener("click", () => {
        if (btn.classList.contains("is-past")) return;

        selectedISO = iso;
        if (tourDateHidden) tourDateHidden.value = selectedISO;

        calGrid.querySelectorAll(".rr-cal-day").forEach((b) => {
          const bISO = b.getAttribute("data-date");
          if (!bISO) return;
          b.classList.toggle("is-selected", bISO === selectedISO);
        });

        if (calSelected) {
          const pretty = new Date(selectedISO).toLocaleDateString(undefined, {
            weekday: "long", year: "numeric", month: "long", day: "numeric"
          });
          calSelected.textContent = `Selected: ${pretty}`;
        }
      });

      calGrid.appendChild(btn);
    }
  }

  if (calPrev) {
    calPrev.addEventListener("click", async () => {
      viewMonth -= 1;
      if (viewMonth < 0) { viewMonth = 11; viewYear -= 1; }
      await renderCalendar();
    });
  }

  if (calNext) {
    calNext.addEventListener("click", async () => {
      viewMonth += 1;
      if (viewMonth > 11) { viewMonth = 0; viewYear += 1; }
      await renderCalendar();
    });
  }

  // Initial render
  renderCalendar();

  // --------- submit ----------
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    clearErrors();
    if (messageEl) messageEl.textContent = "";

    const data = new FormData(form);

    // --- sanitize phone BEFORE sending to Apps Script ---
  const phoneClean = (data.get("phone") || "")
    .toString()
    .replace(/\s+/g, "")   // remove any spaces
    .replace(/^\+/, "");   // remove leading + (prevents Sheets formula parse)
  data.set("phone", phoneClean);

    const name = (data.get("name") || "").toString().trim();
    const email = (data.get("email") || "").toString().trim();
    const guests = Number(data.get("guests"));
    const tourDate = (data.get("tour_date") || "").toString().trim();
    const consent = data.get("consent");

    let valid = true;

    if (!name) { showError("name", "Please add your name."); valid = false; }
    if (!email || !email.includes("@")) { showError("email", "Please enter a valid email."); valid = false; }
    if (!guests || guests < 1) { showError("guests", "Guests must be at least 1."); valid = false; }
    if (!tourDate) { showError("tour_date", "Please choose a date."); valid = false; }
    if (!consent) { showError("consent", "We need your consent to store your details."); valid = false; }

    if (!valid) return;

    submitBtn.disabled = true;
    submitBtn.textContent = "Sending...";

    try {
      await fetch(SCRIPT_URL, { method: "POST", body: data, mode: "no-cors" });
      form.reset();
      if (messageEl) {
        messageEl.textContent = "Thank you! We’ve received your request. We’ll email you shortly to confirm.";
      }
      // re-render (optional, keeps calendar view consistent)
      await renderCalendar();
    } catch (err) {
      console.error(err);
      if (messageEl) {
        messageEl.textContent = "Something went wrong. Please try again or email us directly.";
      }
    } finally {
      submitBtn.disabled = false;
      submitBtn.textContent = "Reserve my spot";
    }
  });

  function showError(fieldName, msg) {
    const errorEl = document.querySelector(`.field-error[data-error-for="${fieldName}"]`);
    if (errorEl) errorEl.textContent = msg;
  }

  function clearErrors() {
    document.querySelectorAll(".field-error").forEach((el) => (el.textContent = ""));
  }
});



document.addEventListener("DOMContentLoaded", () => {
  const track = document.querySelector(".reviews-track");
  if (!track) return;

  function needsReadMore(card) {
    const text = card.querySelector(".review-text");
    if (!text) return false;

    // Clone the text node and measure natural height (no clamp)
    const clone = text.cloneNode(true);
    clone.style.position = "absolute";
    clone.style.visibility = "hidden";
    clone.style.pointerEvents = "none";
    clone.style.height = "auto";
    clone.style.maxHeight = "none";
    clone.style.overflow = "visible";
    clone.style.display = "block";
    clone.style.webkitLineClamp = "unset";
    clone.style.width = `${text.clientWidth}px`;

    document.body.appendChild(clone);
    const natural = clone.scrollHeight;
    document.body.removeChild(clone);

    const available = text.clientHeight;
    return natural > available + 6; // tolerance to avoid false positives
  }

  function updateButtons() {
    track.querySelectorAll(".review-card:not(.clone)").forEach(card => {
      const btn = card.querySelector(".review-more");
      if (!btn) return;

      if (card.classList.contains("is-expanded")) {
        btn.hidden = false;
        btn.textContent = "Read less";
        btn.setAttribute("aria-expanded", "true");
        return;
      }

      btn.hidden = !needsReadMore(card);
      btn.textContent = "Read more";
      btn.setAttribute("aria-expanded", "false");
    });
  }

  // IMPORTANT: run after your cloning code has executed + layout settled
  requestAnimationFrame(() => requestAnimationFrame(updateButtons));
  window.addEventListener("resize", () => requestAnimationFrame(updateButtons));

  // Toggle only clicked card, then re-check buttons
  track.addEventListener("click", (e) => {
    const btn = e.target.closest(".review-more");
    if (!btn) return;

    const card = btn.closest(".review-card");
    if (!card) return;

    if (card.classList.contains("clone")) return;

    card.classList.toggle("is-expanded");
    requestAnimationFrame(updateButtons);
  });
});



