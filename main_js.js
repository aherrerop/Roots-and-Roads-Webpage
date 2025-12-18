// ======================================
// 1. CONFIG
// ======================================

// Infinite carousels (clones) — keep ONLY for the image gallery
const CAROUSEL_CONFIGS = [
  {
    name: "imageGallery",
    trackSelector: ".gallery-track",
    slideSelector: ".gallery-image",
    dotSelector: ".subway-dot",
    blocksEachSide: 5,
  },
];

// Simple scroll galleries (NO CLONES) — reviews go here
const SIMPLE_SCROLL_GALLERIES = [
  {
    name: "reviewCarousel",
    trackSelector: ".reviews-track",
    slideSelector: ".review-card",
    dotSelector: ".review-dot",
    hidePartialSlides: true, // IMPORTANT: only show fully visible cards
  },
];

// Hover behaviors – for future use
const HOVER_CONFIGS = [];



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
  const { trackSelector, slideSelector, dotSelector, blocksEachSide } = config;

  const track = document.querySelector(trackSelector);
  if (!track) return;

  const dots = Array.from(document.querySelectorAll(dotSelector));

  const originals = Array.from(track.querySelectorAll(slideSelector));
  const N = originals.length;
  if (N === 0) return;

  const BLOCKS_EACH_SIDE = blocksEachSide ?? 5;

  const fragStart = document.createDocumentFragment();
  const fragEnd = document.createDocumentFragment();

  // Clone full blocks BEFORE and AFTER originals
  for (let b = 0; b < BLOCKS_EACH_SIDE; b++) {
    originals.forEach((slide) => {
      const cloneBefore = slide.cloneNode(true);
      cloneBefore.classList.add("clone");
      fragStart.appendChild(cloneBefore);
    });

    originals.forEach((slide) => {
      const cloneAfter = slide.cloneNode(true);
      cloneAfter.classList.add("clone");
      fragEnd.appendChild(cloneAfter);
    });
  }

  // [ clones-before ][ originals ][ clones-after ]
  track.insertBefore(fragStart, originals[0]);
  track.appendChild(fragEnd);

  const slides = Array.from(track.querySelectorAll(slideSelector));
  const offsetStart = BLOCKS_EACH_SIDE * N; // first original index

  // map physical index -> logical 0..N-1
  function physicalToLogical(i) {
    let idx = (i - offsetStart) % N;
    if (idx < 0) idx += N;
    return idx;
  }

  function setActiveDotFromPhysicalIndex(physicalIndex) {
    const logicalIndex = physicalToLogical(physicalIndex);
    dots.forEach((dot, i) => {
      dot.classList.toggle("active", i === logicalIndex);
    });
  }

  // slide whose center is closest to track center
  function getClosestSlideIndex() {
    const centerX = track.scrollLeft + track.clientWidth / 2;
    let closestIndex = 0;
    let minDist = Infinity;

    slides.forEach((slide, i) => {
      const slideCenter = slide.offsetLeft + slide.offsetWidth / 2;
      const dist = Math.abs(slideCenter - centerX);
      if (dist < minDist) {
        minDist = dist;
        closestIndex = i;
      }
    });

    return closestIndex;
  }

  function updateActiveDot() {
    const idx = getClosestSlideIndex();
    setActiveDotFromPhysicalIndex(idx);
  }

  function targetScrollForSlide(slideIndex) {
    const slide = slides[slideIndex];
    if (!slide) return track.scrollLeft;
    const slideCenter = slide.offsetLeft + slide.offsetWidth / 2;
    const half = track.clientWidth / 2;
    return slideCenter - half;
  }

  // Scroll: user drives, we only update the dot
  track.addEventListener("scroll", () => {
    requestAnimationFrame(updateActiveDot);
  });

  // Dot clicks: center logical slide in the middle block
  dots.forEach((dot, logicalIndex) => {
    dot.addEventListener("click", () => {
      const targetIndex = offsetStart + logicalIndex; // middle originals
      track.scrollLeft = targetScrollForSlide(targetIndex); // instant
      updateActiveDot();
    });
  });

  // Initial position: center first original once on load
  function recenterFirst() {
    const firstIndex = offsetStart;
    track.scrollLeft = targetScrollForSlide(firstIndex);
    updateActiveDot();
  }

  requestAnimationFrame(recenterFirst);

  // On resize: re-center that first original
  window.addEventListener("resize", () => {
    requestAnimationFrame(recenterFirst);
  });
}

/**
 * Simple scroll gallery (no clones).
 * Keeps dots in sync with closest slide and scrolls to slide on dot click.
 * Currently unused; here for future galleries.
 */
function initSimpleScrollGallery(config) {
  const { trackSelector, slideSelector, dotSelector } = config;

  const track = document.querySelector(trackSelector);
  if (!track) return;

  const dots = Array.from(document.querySelectorAll(dotSelector));
  const slides = Array.from(track.querySelectorAll(slideSelector));
  if (slides.length === 0) return;

  function getClosestSlideIndex() {
    const left = track.getBoundingClientRect().left;
    let bestIdx = 0;
    let bestDist = Infinity;

    slides.forEach((slide, idx) => {
      const r = slide.getBoundingClientRect();
      const dist = Math.abs(r.left - left);
      if (dist < bestDist) {
        bestDist = dist;
        bestIdx = idx;
      }
    });
    return bestIdx;
  }

  function setActiveDot(idx) {
    if (!dots.length) return;
    dots.forEach((d, i) => d.classList.toggle("active", i === idx));
  }

  track.addEventListener("scroll", () => {
    window.requestAnimationFrame(() => {
      setActiveDot(getClosestSlideIndex());
    });
  });

  dots.forEach((dot, idx) => {
    dot.addEventListener("click", () => {
      slides[idx].scrollIntoView({ behavior: "smooth", inline: "start", block: "nearest" });
      setActiveDot(idx);
    });
  });

  setActiveDot(0);
}




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

/**
 * Generic hover handler – for future components.
 */
function initHoverEffects(configs) {
  configs.forEach((cfg) => {
    const { triggerSelector, hoverClass } = cfg;
    if (!triggerSelector || !hoverClass) return;

    const elements = document.querySelectorAll(triggerSelector);
    elements.forEach((el) => {
      el.addEventListener("mouseenter", () => el.classList.add(hoverClass));
      el.addEventListener("mouseleave", () => el.classList.remove(hoverClass));
    });
  });
}


// ======================================
// 4. INITIALIZATION
// ======================================

document.addEventListener("DOMContentLoaded", () => {
  // Nav tabs active state
  initNavActiveTab();

  // Infinite carousels (the ones you already have)
  CAROUSEL_CONFIGS.forEach((cfg) => initInfiniteCarousel(cfg));

  // Simple galleries (future use)
  SIMPLE_SCROLL_GALLERIES.forEach((cfg) => initSimpleScrollGallery(cfg));

  // Hover behaviors (future use)
  initHoverEffects(HOVER_CONFIGS);

  const rTrack = document.querySelector(".reviews-track");
if (rTrack) {
  const updateFade = () => {
    rTrack.classList.toggle("is-scrolled", rTrack.scrollLeft > 2);
  };
  updateFade();
  rTrack.addEventListener("scroll", updateFade, { passive: true });
}

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
    track.querySelectorAll(".review-card").forEach(card => {
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

    card.classList.toggle("is-expanded");
    requestAnimationFrame(updateButtons);
  });
});



