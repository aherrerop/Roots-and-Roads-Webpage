# Roots & Roads AI/search tracking setup

## GA4 traffic tracking

Go to: Google Analytics → Reports → Acquisition → Traffic acquisition.

Set the dimension to: Session source / medium.

Search weekly for:

- chatgpt
- perplexity
- copilot
- bing
- google

Record:

- Sessions
- Total users
- Key events

## Booking tracking

Add `tracking/booking_success_gtag_snippet.js` inside `main_js.js` only after the booking has been successfully submitted.

Then in GA4:

Admin → Data display → Events → find `booking_success` → mark as Key event.

Optional: if you later redirect successful bookings to `/booking-success/`, that page already fires `booking_success` and is marked noindex.

## Google Search Console setup

1. Go to Search Console.
2. Add property using URL prefix: https://rootsandroadsbcn.com/
3. Try verification with Google Analytics.
4. If needed, use the HTML meta tag verification method and paste the tag into the `<head>` of `index.html`.
5. Submit sitemap: `sitemap.xml`.
6. Use URL Inspection and request indexing for the homepage and main landing pages.

## Weekly dashboard

- GA4 total sessions
- GA4 chatgpt.com sessions
- GA4 perplexity.ai sessions
- GA4 copilot sessions
- GA4 google / organic sessions
- GA4 bing / organic sessions
- booking_success key events
- Search Console impressions
- Search Console clicks
- Top query
- Top landing page
