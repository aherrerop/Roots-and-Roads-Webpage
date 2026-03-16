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

    const phoneRaw = (data.get("phone") || "").toString();
    const phoneClean = phoneRaw.replace(/\s+/g, ""); // remove spaces ONLY
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
    messageEl.textContent =
      "Thank you! We’ve received your request. We’ll email you shortly to confirm.";
  }

  if (typeof window.gtag === "function") {
    window.gtag("event", "booking_submit", {
      event_category: "booking",
      event_label: "free_walking_tour",
    });
  }



      // re-render (optional, keeps calendar view consistent)
      try { 
        await renderCalendar(); 
      } catch (e) { 
        console.warn("Calendar rerender failed:", e); 
      };
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


const galleryImages = document.querySelectorAll(".see-gallery img");
const modal = document.getElementById("gallery-modal");
const modalImages = document.querySelectorAll(".gallery-modal img");
const closeBtn = document.querySelector(".gallery-close");

galleryImages.forEach((img, index) => {

  img.addEventListener("click", () => {

    modal.style.display = "flex";

    const src = img.getAttribute("src");
    const target = document.querySelector(`.gallery-modal img[src="${src}"]`);
    target.scrollIntoView({
      block: "nearest",
      inline: "center"
    });

  });

});

closeBtn.addEventListener("click", () => {
  modal.style.display = "none";
});

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