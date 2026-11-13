# Changelog

## 2.1.3 (2026-11-13)

- The website is ready for search engines: every page has its own title, description, address, and link preview, the home page and the live demo describe Offizin as free software in a form search engines read, and the site has a sitemap, a robots file, and favicons in the sizes search results use.
- Pages that are not meant to be found on their own (the Word task pane, the Google sidebar file, and the old demo address) ask search engines to leave them out.

## 2.1.2 (2026-11-13)

- The preview shows the two-page resume and the CVs on two pages, the way Word and Google Docs lay them out, instead of warning that text runs past the first page. Headings stay with the text that follows them.

## 2.1.1 (2026-11-13)

- Headings use a typewriter-like code face. Publications get the new Typewriter font scheme by default: IBM Plex Mono headings in Google Docs and the live demo, Consolas in Word (the closest face every copy of Word has), with Arial body text. Settings that still use the old default fonts (Editorial) move to Typewriter; choose Editorial again on the Design tab if you prefer it.
- The website and the live demo use IBM Plex Mono for their headings. The Offizin wordmark is unchanged.

## 2.1.0 (2026-11-12)

- Offizin's own Undo and Redo, with a history list: they reverse only what Offizin did (publications, page parts, designs, page setup, and columns) and leave your typing and the editor's own Undo alone. Keyboard: Ctrl or Command Z, and Ctrl or Command Shift Z.
- Undo of a publication puts back the headers and footers, the page setup of each section, and the styles that were there before, and it says so when something cannot be restored. Undo of a page part remembers what you typed in it, and Redo puts that back.
- Customize every publication: paper size, orientation, margins with Word's presets, text size, line spacing, paragraph spacing, padding, and gaps. Shrink to fit lowers the text and spacing step by step until every page fits. Default spacing and text are on the Design tab.
- 15 new publications (33 in all): six plain resumes and CVs that screening software reads reliably, booklets of 4 to 16 pages, name badges, tickets, gift certificates, bookmarks, place cards, envelopes, an invoice, and a photo page.
- Mail merge for labels, envelopes, badges, tickets, place cards, and certificates: paste a list or open a .csv file. A column your list does not have leaves its line blank.
- Black and white color schemes (Ink, the new default, and Graphite), and your own fonts for Word and Google Docs. Settings saved by Offizin 2.0 with its untouched default colors (Harbor) now open in Ink. Choose Harbor again on the Design tab if you prefer it.
- Apply design switches every color exactly, so a restyled publication matches a new one in that scheme, and Offizin's Undo brings every color back. Colored and gray text is dark enough to read in every scheme.
- Example text is plain and short: placeholders for names, dates, prices, and contact details, and Middle High German filler for running text. The Design Checker finds all of it, including short placeholders, text in capitals, and page parts you inserted.
- Native styles (Title, Heading 1, Heading 2, Normal), text columns, and page number footers in Word. The plain resumes and CVs take their name and heading look from those styles, so changing a style changes every matching line.
- Booklets print in the right order in Word (book fold) and Google Docs (printer spreads).
- Print settings kept in the document where Word allows it (mirror margins, two pages per sheet, book fold, embedded fonts), plus print steps for each editor and each version of Word.
- Word: a publication made with Open in a new document remembers its design, so Apply design and the Design Checker work there too. Page parts carry a hidden marker and no longer show a frame.
- Google Docs: if a publication cannot be finished, the document is put back as it was, and a document too large to copy for Undo stops before anything changes. The sidebar waits for your saved settings before it saves, so a quick change can no longer replace them or remove your logo.
- Finish publication removes Offizin's invisible markers when a publication is done.
- Keyboard and screen reader use: the focus stays where you were after dialogs and redraws, every button shows its focus, and previews are announced once instead of line by line.
- Download packages for Word and Google Docs, each with a plain text install guide.
- Redesigned live demo with pages, rulers, zoom, guides, and print preview.
- Google Docs: paste the three new files into your starter document to update.

## 2.0.0 (2026-11-11)

Layout Toolkit is now **Offizin**, in one repository for Word and Google Docs.

- New name, logo, and website at offizin.alyssaagard.dev, with a live demo that needs no install.
- One template library shared by Word, Google Docs, and the demo: 18 publications (up from 6), each for Letter or A4 paper.
- Word: publications are built as complete Word files that carry their own page size, margins, and panels. Choose to replace the document or open a new one.
- Word: an Offizin ribbon tab with Publications, Design, Page Parts, Business Info, and Check and Print.
- Color schemes (12) and font schemes (8), plus your own colors, with restyling of finished publications.
- Business information sets and a logo, filled into every template.
- 14 page parts that insert at the cursor.
- Picture placeholders that you replace with Change Picture or Replace image.
- Design checker for leftover sample text, placeholders, and missing business information.
- Folding and printing guides for each publication type.
- Brand kit files for sharing designs with a team.
- Word page setup and columns tools where Word allows add-ins to use them (Windows and Mac).
- Google Docs keeps the narrowest permission: access to the current document only.
- Install guides for personal, work, school, and organization-wide deployment, plus privacy, terms, and help pages.

## 1.0.0 (2026-04-06)

Layout Toolkit for Word (web and desktop editions) and Google Docs: six templates, layout tools, brand profiles, and folding guides.
