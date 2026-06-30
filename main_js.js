// ======================================
// 1. Booking form + availability calendar
// ======================================
// This section:
// - Loads availability from Google Apps Script.
// - Shows only dates/times available for the selected language.
// - Sends booking requests to Google Apps Script.
// - Refreshes the calendar after a booking is submitted.

document.addEventListener("DOMContentLoaded", () => {
  // --------------------------------------
  // Google Apps Script endpoints
  // --------------------------------------
  // SCRIPT_URL receives website bookings through doPost(e).
  // AVAILABILITY_URL returns month availability through doGet(e).
  const SCRIPT_URL =
    "https://script.google.com/macros/s/AKfycby8qYjPhWntY9iW2_eh168Um_XrFTNB9veuH9gJSu3p8FYpSwhSU14p81ZiVTrs8kZ6/exec";

  const AVAILABILITY_URL =
    "https://script.google.com/macros/s/AKfycby8qYjPhWntY9iW2_eh168Um_XrFTNB9veuH9gJSu3p8FYpSwhSU14p81ZiVTrs8kZ6/exec";

  // --------------------------------------
  // Form elements
  // --------------------------------------
  const form = document.getElementById("booking-form");
  const messageEl = document.getElementById("booking-message");
  const submitBtn = form ? form.querySelector(".booking-submit") : null;

  const languageSelect = document.getElementById("tour_language");
  const tourDateHidden = document.getElementById("tour_date");
  const tourTimeHidden = document.getElementById("tour_time");
  const tourTimeOptions = document.getElementById("tour-time-options");

  // --------------------------------------
  // Calendar elements
  // --------------------------------------
  const calGrid = document.getElementById("rr-cal-grid");
  const calTitle = document.getElementById("rr-cal-title");
  const calPrev = document.getElementById("rr-cal-prev");
  const calNext = document.getElementById("rr-cal-next");
  const calSelected = document.getElementById("rr-cal-selected");
  const calendarBox = document.getElementById("rr-calendar");

  // Stop if this page does not contain the booking form.
  if (!form || !submitBtn) return;
  if (!languageSelect || !tourDateHidden || !tourTimeHidden || !tourTimeOptions) return;
  if (!calGrid || !calTitle || !calPrev || !calNext || !calSelected) return;

  // --------------------------------------
  // Calendar state
  // --------------------------------------
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let viewYear = today.getFullYear();
  let viewMonth = today.getMonth();

  let selectedISO = "";
  let selectedTime = "";

  // This stores the availability returned by Apps Script.
  // Shape:
  // {
  //   "2026-06-24": [
  //     { language: "English", time: "5:00 PM", spotsLeft: 0, available: false }
  //   ]
  // }
  let monthSlots = {};

  // ======================================
  // Small date/time helpers
  // ======================================

  function pad2(n) {
    return String(n).padStart(2, "0");
  }

  function monthKey(year, monthIndex) {
    return `${year}-${pad2(monthIndex + 1)}`;
  }

  function isoDate(year, monthIndex, day) {
    return `${year}-${pad2(monthIndex + 1)}-${pad2(day)}`;
  }

  function isoToLocalDate(iso) {
    const [year, month, day] = iso.split("-").map(Number);
    return new Date(year, month - 1, day);
  }

  function monthName(year, monthIndex) {
    return new Date(year, monthIndex, 1).toLocaleString(undefined, {
      month: "long",
      year: "numeric",
    });
  }

  function mondayFirstIndex(date) {
    return (date.getDay() + 6) % 7;
  }

  function selectedLanguage() {
    return String(languageSelect.value || "English").trim();
  }

  function timeToMinutes(time) {
    const s = String(time || "").trim();

    const ampm = s.match(/^(\d{1,2}):(\d{2})\s*(AM|PM)$/i);
    if (ampm) {
      let h = Number(ampm[1]);
      const m = Number(ampm[2]);
      const ap = ampm[3].toUpperCase();

      if (ap === "PM" && h !== 12) h += 12;
      if (ap === "AM" && h === 12) h = 0;

      return h * 60 + m;
    }

    const hhmm = s.match(/^(\d{1,2}):(\d{2})$/);
    if (hhmm) {
      return Number(hhmm[1]) * 60 + Number(hhmm[2]);
    }

    return 99999;
  }

  // ======================================
  // Availability loading
  // ======================================

  function normalizeSlot(slot) {
    const language = String(slot.language || slot.lang || "").trim();
    const time = String(slot.time || slot.tour_time || "").trim();

    const spotsLeft = Number(
      slot.spotsLeft ??
      slot.spots_left ??
      slot.remaining ??
      0
    );

    const status = String(slot.status || "").toUpperCase();
    const closed = status.startsWith("CLOSED");

    return {
      language,
      time,
      spotsLeft,
      available: slot.available === true && !closed && spotsLeft > 0,
    };
  }

async function fetchAvailabilityForMonth(ym) {
  try {
    const url = `${AVAILABILITY_URL}?ym=${encodeURIComponent(ym)}&t=${Date.now()}`;

    console.log("[R&R] Fetching availability URL:", url);

    const res = await fetch(url, { cache: "no-store" });
    const json = await res.json();

    console.log("[R&R] Raw June 24 from backend:", json.days?.["2026-06-24"]);

    const out = {};

    if (json && json.days && typeof json.days === "object") {
      Object.keys(json.days).forEach((iso) => {
        const slots = json.days[iso];
        out[iso] = Array.isArray(slots) ? slots.map(normalizeSlot) : [];
      });
    }

    console.log("[R&R] Normalized June 24 in frontend:", out["2026-06-24"]);

    return out;
  } catch (err) {
    console.warn("[R&R] Availability fetch failed:", err);
    return {};
  }
}

  // IMPORTANT:
  // This function decides what is available for the selected language.
  // Example:
  // - English full, Spanish open => English calendar must show red.
  // - Spanish open => Spanish calendar can still show green.
function openSlotsForDate(iso) {
  const lang = selectedLanguage().toLowerCase();
  const slots = Array.isArray(monthSlots[iso]) ? monthSlots[iso] : [];

  const checked = slots.map((slot) => {
    const slotLang = String(slot.language || "").trim().toLowerCase();
    const spots = Number(slot.spotsLeft || 0);

    return {
      iso,
      selectedLanguage: lang,
      slotLanguage: slot.language,
      slotTime: slot.time,
      slotAvailable: slot.available,
      slotSpotsLeft: slot.spotsLeft,
      languageMatches: slotLang === lang,
      availabilityPasses: slot.available === true,
      spotsPasses: spots > 0,
      finalPasses:
        slotLang === lang &&
        slot.available === true &&
        spots > 0,
    };
  });

  if (iso === "2026-06-24") {
    console.log("[R&R] Slot filter check June 24:", JSON.stringify(checked, null, 2));
  }

  return checked
    .filter((x) => x.finalPasses)
    .map((x) => ({
      language: x.slotLanguage,
      time: x.slotTime,
      available: x.slotAvailable,
      spotsLeft: x.slotSpotsLeft,
    }))
    .sort((a, b) => timeToMinutes(a.time) - timeToMinutes(b.time));
}

  // ======================================
  // Calendar rendering
  // ======================================

  function applyDayState(btn, iso) {
    const d = isoToLocalDate(iso);
    d.setHours(0, 0, 0, 0);

    const isPast = d < today;
    const openSlots = openSlotsForDate(iso);
    const isClosed = openSlots.length === 0;

    btn.classList.toggle("is-past", isPast);
    btn.classList.toggle("is-open", !isPast && !isClosed);
    btn.classList.toggle("is-closed", !isPast && isClosed);

    btn.disabled = isPast || isClosed;
  }

  function renderTimeOptions() {
    selectedTime = "";
    tourTimeHidden.value = "";
    tourTimeOptions.innerHTML = "";

    if (!selectedLanguage() || !selectedISO) {
      tourTimeOptions.innerHTML =
        '<p class="tour-time-placeholder">Choose a language and date first.</p>';
      return;
    }

    const slots = openSlotsForDate(selectedISO);

    if (!slots.length) {
      tourTimeOptions.innerHTML =
        '<p class="tour-time-placeholder">No available times for this date.</p>';
      return;
    }

    slots.forEach((slot) => {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "tour-time-btn";
      btn.textContent = slot.time;
      btn.dataset.time = slot.time;

      btn.addEventListener("click", () => {
        selectedTime = slot.time;
        tourTimeHidden.value = slot.time;

        tourTimeOptions.querySelectorAll(".tour-time-btn").forEach((b) => {
          b.classList.toggle("is-selected", b.dataset.time === selectedTime);
        });
      });

      tourTimeOptions.appendChild(btn);
    });
  }

  async function renderCalendar() {
    const ym = monthKey(viewYear, viewMonth);

    calTitle.textContent = monthName(viewYear, viewMonth);

    if (calendarBox) calendarBox.classList.add("is-loading");
    calGrid.setAttribute("aria-busy", "true");

    monthSlots = await fetchAvailabilityForMonth(ym);

    calGrid.innerHTML = "";

    const first = new Date(viewYear, viewMonth, 1);
    const firstIndex = mondayFirstIndex(first);
    const daysInMonth = new Date(viewYear, viewMonth + 1, 0).getDate();

    // Blank cells before day 1.
    for (let i = 0; i < firstIndex; i++) {
      const blank = document.createElement("button");
      blank.type = "button";
      blank.className = "rr-cal-day is-blank";
      blank.tabIndex = -1;
      blank.disabled = true;
      calGrid.appendChild(blank);
    }

    // Real calendar days.
    for (let day = 1; day <= daysInMonth; day++) {
      const iso = isoDate(viewYear, viewMonth, day);

      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "rr-cal-day";
      btn.textContent = String(day);
      btn.dataset.date = iso;

      applyDayState(btn, iso);

      btn.addEventListener("click", () => {
        if (btn.disabled) return;

        selectedISO = iso;
        tourDateHidden.value = iso;

        calGrid.querySelectorAll(".rr-cal-day").forEach((b) => {
          b.classList.toggle("is-selected", b.dataset.date === selectedISO);
        });

        const pretty = isoToLocalDate(selectedISO).toLocaleDateString(undefined, {
          weekday: "long",
          year: "numeric",
          month: "long",
          day: "numeric",
        });

        calSelected.textContent = `Selected: ${pretty}`;

        renderTimeOptions();
      });

      calGrid.appendChild(btn);
    }

    if (calendarBox) calendarBox.classList.remove("is-loading");
    calGrid.setAttribute("aria-busy", "false");
  }

  // ======================================
  // Calendar controls
  // ======================================

  languageSelect.addEventListener("change", async () => {
    selectedISO = "";
    selectedTime = "";
    tourDateHidden.value = "";
    tourTimeHidden.value = "";
    calSelected.textContent = "";

    renderTimeOptions();
    await renderCalendar();
  });

  calPrev.addEventListener("click", async () => {
    viewMonth -= 1;

    if (viewMonth < 0) {
      viewMonth = 11;
      viewYear -= 1;
    }

    selectedISO = "";
    selectedTime = "";
    tourDateHidden.value = "";
    tourTimeHidden.value = "";
    calSelected.textContent = "";

    renderTimeOptions();
    await renderCalendar();
  });

  calNext.addEventListener("click", async () => {
    viewMonth += 1;

    if (viewMonth > 11) {
      viewMonth = 0;
      viewYear += 1;
    }

    selectedISO = "";
    selectedTime = "";
    tourDateHidden.value = "";
    tourTimeHidden.value = "";
    calSelected.textContent = "";

    renderTimeOptions();
    await renderCalendar();
  });

  // Initial render.
  renderTimeOptions();
  renderCalendar();

  // ======================================
  // Booking form submission
  // ======================================

  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    clearErrors();

    if (messageEl) messageEl.textContent = "";

    const data = new FormData(form);

    const phoneRaw = String(data.get("phone") || "");
    data.set("phone", phoneRaw.replace(/\s+/g, ""));

    const language = String(data.get("language") || "").trim();
    const name = String(data.get("name") || "").trim();
    const email = String(data.get("email") || "").trim();
    const guests = Number(data.get("guests"));
    const tourDate = String(data.get("tour_date") || "").trim();
    const tourTime = String(data.get("tour_time") || "").trim();
    const consent = data.get("consent");

    let valid = true;

    if (!language) {
      showError("language", "Please choose a language.");
      valid = false;
    }

    if (!name) {
      showError("name", "Please add your name.");
      valid = false;
    }

    if (!email || !email.includes("@")) {
      showError("email", "Please enter a valid email.");
      valid = false;
    }

    if (!guests || guests < 1) {
      showError("guests", "Guests must be at least 1.");
      valid = false;
    }

    if (!tourDate) {
      showError("tour_date", "Please choose a date.");
      valid = false;
    }

    if (!tourTime) {
      showError("tour_time", "Please choose a time.");
      valid = false;
    }

    if (!consent) {
      showError("consent", "We need your consent to store your details.");
      valid = false;
    }

    const chosenSlot = openSlotsForDate(tourDate).find((slot) => {
      return slot.time === tourTime;
    });

    if (!chosenSlot) {
      showError("tour_time", "This time is no longer available.");
      valid = false;
    } else if (guests > chosenSlot.spotsLeft) {
      showError("guests", `Only ${chosenSlot.spotsLeft} spots left.`);
      valid = false;
    }

    if (!valid) return;

    submitBtn.disabled = true;
    submitBtn.textContent = "Sending...";

    try {
      await fetch(SCRIPT_URL, {
        method: "POST",
        body: data,
        mode: "no-cors",
      });

      form.reset();

      selectedISO = "";
      selectedTime = "";
      tourDateHidden.value = "";
      tourTimeHidden.value = "";
      calSelected.textContent = "";

      renderTimeOptions();

      if (messageEl) {
        messageEl.textContent =
          "Thank you! We’ve received your request. We’ll email you shortly to confirm.";
      }

      if (typeof window.gtag === "function") {
        window.gtag("event", "booking_submit", {
          event_category: "booking",
          event_label: "free_walking_tour",
        });
      }

      await renderCalendar();
    } catch (err) {
      console.error(err);

      if (messageEl) {
        messageEl.textContent =
          "Something went wrong. Please try again or email us directly.";
      }
    } finally {
      submitBtn.disabled = false;
      submitBtn.textContent = "Reserve your spot";
    }
  });

  // ======================================
  // Form error helpers
  // ======================================

  function showError(fieldName, msg) {
    const errorEl = document.querySelector(
      `.field-error[data-error-for="${fieldName}"]`
    );

    if (errorEl) errorEl.textContent = msg;
  }

  function clearErrors() {
    document.querySelectorAll(".field-error").forEach((el) => {
      el.textContent = "";
    });
  }
});

// ======================================
// Gallery modal
// Opens the full-screen image gallery when a gallery image is clicked.
// Scrolls the modal to the matching image.
// Closes the modal when the X button is clicked.
// ======================================
const galleryImages = document.querySelectorAll(".see-gallery img");
const modal = document.getElementById("gallery-modal");
const modalContent = document.querySelector(".gallery-modal-content");
const closeBtn = document.querySelector(".gallery-close");

galleryImages.forEach((img) => {
  img.addEventListener("click", () => {
    modal.style.display = "flex";

    const src = img.getAttribute("src");
    const target = modalContent.querySelector(`img[src="${src}"]`);

    if (target) {
      modalContent.scrollTo({
        left: target.offsetLeft - (modalContent.clientWidth - target.clientWidth) / 2,
        behavior: "smooth"
      });
    }
  });
});

closeBtn.addEventListener("click", () => {
  modal.style.display = "none";
});


// ======================================
// Reviews carousel arrows
// Moves the review cards left/right by one card width.
// Hides the left arrow at the beginning.
// Hides the right arrow at the end.
// ======================================
document.addEventListener("DOMContentLoaded", () => {

const reviewsTrack = document.querySelector(".reviews-track");
const leftArrow = document.querySelector(".reviews-arrow-left");
const rightArrow = document.querySelector(".reviews-arrow-right");

if(reviewsTrack && leftArrow && rightArrow){

const cardWidth = document.querySelector(".review-card").offsetWidth;
const gap = parseInt(getComputedStyle(reviewsTrack).gap) || 0;
const step = cardWidth + gap;

function updateArrows(){

leftArrow.classList.toggle(
"hidden",
reviewsTrack.scrollLeft <= 5
);

rightArrow.classList.toggle(
"hidden",
reviewsTrack.scrollLeft + reviewsTrack.clientWidth >= reviewsTrack.scrollWidth - 5
);

}

leftArrow.addEventListener("click", ()=>{
reviewsTrack.scrollBy({left:-step, behavior:"smooth"});
});

rightArrow.addEventListener("click", ()=>{
reviewsTrack.scrollBy({left:step, behavior:"smooth"});
});

reviewsTrack.addEventListener("scroll", updateArrows);

updateArrows();

}

});