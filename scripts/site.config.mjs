/**
 * The one place to change Offizin's public address, version, and publisher.
 *
 * SITE_URL is where docs/ is published (GitHub Pages with a custom domain). The Word manifest
 * points every icon and page at this address, so change it here and run `npm run build`.
 * People who installed the Word add-in from an older manifest need to install it again
 * after the address changes, so pick the long-term address before sharing the add-in.
 */
export const SITE_URL = "https://offizin.alyssaagard.dev";

/** The Word add-in's permanent ID. Keep it the same for the life of the add-in. */
export const ADDIN_ID = "e0b7a428-5265-42f2-aa60-f3629f4e70fe";

export const VERSION = "2.1.0";

/** The release date of this version, used for the timestamps inside the download packages. */
export const RELEASE_DATE = "2026-11-12";

/** Shown as the publisher in Word. Change to the company name after the transfer. */
export const PROVIDER = "Alyssa I. Agard";

export const REPO_URL = "https://github.com/alyssaagard/offizin";

/**
 * Optional: the "make a copy" link for a Google Doc that already has Offizin installed.
 * Create it once (see MAINTAINING.md, "Google Docs starter document"), share it as "Anyone with the
 * link can view", and paste its address here ending in /copy. Leave empty to hide the button.
 */
export const GOOGLE_STARTER_COPY_URL = "";
