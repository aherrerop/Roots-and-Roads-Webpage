// Add this only after the booking has been successfully saved and confirmed.
if (typeof gtag === "function") {
  gtag("event", "booking_success", {
    event_category: "booking",
    event_label: "website_booking"
  });
}
