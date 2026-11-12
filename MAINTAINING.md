# Maintaining Offizin

Notes for publishing Offizin and keeping it running. Everything here is done by hand in a browser or with the commands shown.

## Release checklist

1. Make changes in `docs/app/`, `google-docs/`, or `scripts/`.
2. If it is a release, raise `VERSION` in `scripts/site.config.mjs` and `OZ_VERSION` in `google-docs/Code.gs`, and add a line to `CHANGELOG.md`.
3. Run `npm run build` and `npm run check`. Both must pass.
4. Commit and push to `main`. GitHub Pages publishes `docs/` within a minute or two.

Word users get the new version automatically, because Word loads Offizin from the website. Google Docs users who pasted the code update by pasting it again (the install page says so), so mention Google changes in the changelog.

## GitHub Pages and the custom domain

The site is published from the `main` branch, `/docs` folder (Settings, Pages).

The permanent address is `https://offizin.alyssaagard.dev`. The Word manifest contains this address, so it should not change once people install Offizin. To connect it:

1. At Namecheap (the DNS host for alyssaagard.dev), open Domain List, Manage, then Advanced DNS, and add a record:
   - Type: `CNAME Record`
   - Host: `offizin`
   - Value: `alyssaagard.github.io.`
   - TTL: Automatic
2. Wait until `offizin.alyssaagard.dev` resolves (usually minutes, sometimes an hour).
3. In the repository, open Settings, then Pages. Under Custom domain, enter `offizin.alyssaagard.dev` and select Save. GitHub adds a `docs/CNAME` file.
4. When the certificate is ready, check Enforce HTTPS. Because `.dev` domains require HTTPS, the site only loads once this step is done.
5. Recommended: verify the domain for your GitHub account (your profile Settings, then Pages, then Add a domain) so no one else can claim the subdomain.

Until the custom domain is set, the site is also reachable at `https://alyssaagard.github.io/offizin/`, but the Word add-in only works at the permanent address.

To use a different address, change `SITE_URL` in `scripts/site.config.mjs`, run `npm run build`, and update the CNAME. Anyone who already installed the Word add-in must install it again.

## Google Docs starter document

The quickest Google install is a shared document that already contains Offizin. To create it:

1. Create a Google Doc named "Offizin starter" (Pages format, not pageless).
2. Add Offizin by following the steps on the Google Docs install page.
3. Open Offizin once and approve it, so you know it works.
4. Share the document: General access, Anyone with the link, Viewer.
5. Copy its link, change the ending from `/edit...` to `/copy`, and paste that into `GOOGLE_STARTER_COPY_URL` in `scripts/site.config.mjs`.
6. Run `npm run build`, commit, and push. The install page then shows a Copy the starter document button.

When Code.gs changes, update the three files in the starter document as well.

## Store listings (after the company developer accounts exist)

- **Microsoft AppSource**: needs a Microsoft Partner Center account for the company. Submit `docs/word/manifest.xml` with the privacy (`/privacy/`), terms (`/terms/`), and support (`/support/`) pages, and the icons in `docs/assets/icons/`. A listing also enables Word for iPad.
- **Google Workspace Marketplace**: needs a Google Cloud project owned by the company, an OAuth consent screen (the scopes are `documents.currentonly` and `script.container.ui`), the Marketplace SDK, and a store listing using the same pages and `icon-32.png` and `icon-128.png`.

## Moving the project to the company

1. Create the company's GitHub organization, then in this repository open Settings, General, Transfer ownership. GitHub redirects the old repository address.
2. Set up GitHub Pages and the custom domain again in the new organization (the steps above). Keep `offizin.alyssaagard.dev`, or choose a company subdomain and follow "To use a different address".
3. Change `PROVIDER` in `scripts/site.config.mjs` to the company's legal name, and update the copyright line in `LICENSE` if the company now holds the copyright.
4. Keep `ADDIN_ID` the same: it identifies the add-in to Word.

## Brand assets

`npm install` then `npm run brand` regenerates every logo, icon, and favicon from `brand/fonts/CormorantGaramond-500-italic.woff` (the same typeface as Cachet). Outputs go to `docs/assets/`, `docs/favicon.*`, and `docs/app/ui/mark.js`. For the GitHub social preview, upload `docs/assets/brand/offizin-social.png` under Settings, General, Social preview.
