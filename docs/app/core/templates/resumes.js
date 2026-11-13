/**
 * Resumes and CVs for Offizin 2.1.
 *
 * The plain resumes follow common guidance for applicant tracking systems (the screening
 * software many employers use to read applications):
 *   - one column, read from top to bottom, with no tables, text boxes, pictures, or rules;
 *   - contact details in the body of the page, never in a header or footer;
 *   - standard section names such as Summary, Experience, Education, and Skills;
 *   - simple round bullets and dates written the same way everywhere ("March 2021 to present").
 * Sources checked: Jobscan, "Resume tables, columns, and ATS"
 * (https://www.jobscan.co/blog/resume-tables-columns-ats/), and USAJOBS Help, "What should I
 * include in my federal resume?" (https://help.usajobs.gov/faq/application/documents/resume/what-to-include).
 *
 * Job titles, employers, schools, and citations below are plain placeholders. Summaries,
 * achievements, and other running text use the Middle High German filler from filler.js.
 */
import { p, list } from "../dsl.js";
import { alt, altList, altShort } from "../filler.js";

/** The heading and name color. Dark text is the default because it prints and copies most reliably. */
const HEADING_OPTION = {
  key: "headingColor",
  label: "Name and heading color",
  type: "select",
  choices: [
    { value: "dark", label: "Dark text (recommended)" },
    { value: "primary", label: "Theme main color" },
    { value: "accent", label: "Theme accent color" },
  ],
  default: "dark",
  help: "Screening software ignores color. Dark text is the safest choice when a resume may be printed in black and white.",
};

const PAGE = { size: "paper", orientation: "portrait", margins: { top: 48, right: 54, bottom: 48, left: 54 } };
const CUSTOMIZE = { paper: ["letter", "a4"], orientation: false, margins: true, spacing: true, text: true };
const PRINT = { paper: "Plain white paper. Upload the file itself (a .docx or a PDF) when you apply online." };

/**
 * The name and heading look, written into the Title, Heading 1, and Heading 2 styles. The kit's
 * paragraphs use exactly these values, so they carry no direct formatting, and changing one of
 * those styles in Word or Google Docs changes every matching line of the resume.
 */
function resumeStyles(ctx) {
  const { hc } = kit(ctx);
  return {
    title: { size: 24, color: hc, line: 1.1, after: 2 },
    h1: { size: 12.5, color: hc, before: 8, after: 4 },
    h2: { size: 10.5, color: "dark", before: 4, after: 3 },
  };
}

/** Writing helpers shared by every resume and CV. */
function kit(ctx) {
  const hc = ["primary", "accent"].includes(ctx.opt.headingColor) ? ctx.opt.headingColor : "dark";
  return {
    hc,
    name: (o = {}) => p("{name}", "title", { size: 24, color: hc, line: 1.1, after: 2, ...o }),
    headline: (text) => p(text, "lead", { size: 11.5, color: "dark", after: 4 }),
    contact: (text) => p(text, "body", { after: 6, line: 1.3 }),
    heading: (text) => p(text, "h1", { size: 12.5, color: hc, before: 8, after: 4, keepNext: true, fixed: true }),
    sub: (text) => p(text, "h2", { size: 10.5, color: "dark", before: 4, after: 3, keepNext: true, fixed: true }),
    text: (text, o = {}) => p(text, "body", { after: 6, ...o }),
    bullets: (items, o = {}) => list(items, { bulletColor: "dark", indent: 14, gap: 2, after: 8, ...o }),
    /** Job title, then employer, place, and dates on one line, then bullets. */
    job: (title, where, dates, bullets, extra) => [
      p(`**${title}**`, "body", { after: 0, keepNext: true }),
      p(`${where}  |  ${dates}`, "body", { after: extra ? 0 : 3, keepNext: true }),
      extra ? p(extra, "body", { after: 3, keepNext: true }) : null,
      list(bullets, { bulletColor: "dark", indent: 14, gap: 2, after: 8 }),
    ],
    /** A degree or course line with its school and dates. */
    school: (degree, where, dates, note) => [
      p(`**${degree}**`, "body", { after: 0, keepNext: true, sample: true }),
      p(`${where}  |  ${dates}`, "body", { after: note ? 2 : 6 }),
      note ? p(note, "body", { after: 6 }) : null,
    ],
    /** A citation with a hanging indent, as academic style guides ask. */
    cite: (text) => p(text, "body", { indent: { left: 22, hanging: 22 }, after: 5 }),
  };
}

const CONTACT = "{city}  |  {phone}  |  {email}\n{profile}";

export const RESUME_TEMPLATES = [
  {
    id: "resume-plain",
    styles: resumeStyles,
    name: "Plain resume, chronological",
    category: "resumes",
    description:
      "A one page resume that screening software reads reliably: one column, standard section names, simple bullets, and your contact details in the body. Most recent job first.",
    sizeLabel: "Portrait, 1 page",
    fold: "none",
    plain: true,
    page: PAGE,
    customize: CUSTOMIZE,
    print: PRINT,
    options: [HEADING_OPTION],
    build(ctx) {
      const k = kit(ctx);
      return [
        k.name(),
        k.headline("{title}"),
        k.contact(CONTACT),
        k.heading("Summary"),
        k.text(alt(2, 0)),
        k.heading("Experience"),
        k.job("Job title", "Organization name, City, ST", "March 2021 to present", altList(3, 0)),
        k.job("Job title", "Organization name, City, ST", "June 2017 to February 2021", altList(2, 5)),
        k.job("Job title", "Organization name, City, ST", "August 2015 to May 2017", altList(1, 7)),
        k.heading("Education"),
        k.school("Degree", "School name, City, ST", "May 2015"),
        k.heading("Skills"),
        k.bullets([
          "Planning and scheduling",
          "Budgets and contracts",
          "Spreadsheets and databases",
          "Team leadership and training",
        ]),
        k.heading("Certifications"),
        k.text("Certification name, 2022", { after: 0 }),
      ];
    },
  },

  {
    id: "resume-plain-skills",
    styles: resumeStyles,
    name: "Plain resume, skills first",
    category: "resumes",
    description:
      "A one page hybrid resume that leads with grouped skills, then lists experience by date. Good for career changers. One column and standard section names, so screening software reads it reliably.",
    sizeLabel: "Portrait, 1 page",
    fold: "none",
    plain: true,
    page: PAGE,
    customize: CUSTOMIZE,
    print: PRINT,
    options: [HEADING_OPTION],
    build(ctx) {
      const k = kit(ctx);
      return [
        k.name(),
        k.headline("{title}"),
        k.contact(CONTACT),
        k.heading("Summary"),
        k.text(alt(2, 2)),
        k.heading("Skills"),
        k.bullets([
          "**Planning:** schedules, budgets, and reports",
          "**Training:** workshops and coaching",
          "**Tools:** spreadsheets and survey software",
          "**Communication:** writing and public speaking",
        ]),
        k.heading("Experience"),
        k.job("Job title (Contract)", "Organization name, remote", "January 2024 to present", altList(3, 4)),
        k.job("Job title", "Organization name, City, ST", "August 2018 to December 2023", altList(3, 7)),
        k.heading("Education"),
        k.school("Certificate name", "School name, online", "December 2023"),
        k.school("Degree", "School name, City, ST", "May 2018"),
        k.heading("Certifications"),
        k.text("License or certification name, 2018", { after: 0 }),
      ];
    },
  },

  {
    id: "resume-plain-early",
    styles: resumeStyles,
    name: "Plain resume, early career",
    category: "resumes",
    description:
      "A one page resume for students and recent graduates that puts education, internships, and projects first. One column and standard section names, so screening software reads it reliably.",
    sizeLabel: "Portrait, 1 page",
    fold: "none",
    plain: true,
    page: PAGE,
    customize: CUSTOMIZE,
    print: PRINT,
    options: [HEADING_OPTION],
    build(ctx) {
      const k = kit(ctx);
      return [
        k.name(),
        k.headline("{title}"),
        k.contact(CONTACT),
        k.heading("Summary"),
        k.text(alt(2, 4)),
        k.heading("Education"),
        k.school("Degree", "School name, City, ST", "May 2026", "Grade point average 3.7. Relevant coursework: Course name, Course name, Course name."),
        k.heading("Experience"),
        k.job("Job title (Internship)", "Organization name, City, ST", "January 2026 to May 2026", altList(2, 8)),
        k.job("Job title (Part time)", "Organization name, City, ST", "September 2023 to December 2025", altList(2, 10)),
        k.heading("Projects"),
        k.bullets([`**Project name:** ${altShort(5)}`, `**Project name:** ${altShort(6)}`]),
        k.heading("Skills"),
        k.bullets([
          "Social media and email",
          "Office software and survey tools",
          "Spanish, professional working proficiency",
        ]),
        k.heading("Activities"),
        k.text("Role, Organization name, 2024 to 2026. Role, Organization name, 2022 to present.", { after: 0 }),
      ];
    },
  },

  {
    id: "resume-plain-federal",
    styles: resumeStyles,
    name: "Plain resume, federal",
    category: "resumes",
    description:
      "A plain resume for United States government jobs. Each job lists month and year dates and hours per week, as USAJOBS asks. Check the job announcement for any page limit.",
    sizeLabel: "Portrait, 2 pages",
    fold: "none",
    plain: true,
    page: PAGE,
    customize: CUSTOMIZE,
    print: PRINT,
    options: [HEADING_OPTION],
    build(ctx) {
      const k = kit(ctx);
      return [
        k.name(),
        k.headline("{title}"),
        k.contact("{address}, {city}\n{phone}  |  {email}"),
        k.text("Citizenship: United States citizen\nVeterans' preference: None claimed\nAvailability: Full time, permanent", { after: 4 }),
        k.heading("Summary"),
        k.text(alt(2, 6)),
        k.heading("Work Experience"),
        k.job(
          "Job title",
          "Organization name, City, ST",
          "March 2021 to present",
          altList(4, 0),
          "40 hours per week. Supervisor: Name, phone. May contact."
        ),
        k.job(
          "Job title",
          "Organization name, City, ST",
          "July 2017 to February 2021",
          altList(3, 7),
          "40 hours per week. Supervisor: Name, phone. May contact."
        ),
        k.job(
          "Job title (Part time)",
          "Organization name, City, ST",
          "September 2015 to June 2017",
          altList(1, 10),
          "20 hours per week."
        ),
        k.heading("Education"),
        k.school("Master's degree", "School name, City, ST", "May 2017", "Coursework: Course name, Course name, Course name."),
        k.school("Bachelor's degree", "School name, City, ST", "May 2015"),
        k.heading("Certifications and Training"),
        k.bullets(["Certification name, 2022", "Course name, 24 hours, 2020", "Course name, 8 hours, 2019"]),
        k.heading("Skills"),
        k.bullets([
          "Budgets and grants management",
          "Data analysis and reporting",
          "Writing briefings and reports",
        ]),
        k.heading("Awards"),
        k.text("Award name, 2023", { after: 0 }),
      ];
    },
  },

  {
    id: "cv-academic",
    styles: resumeStyles,
    name: "Academic CV",
    category: "resumes",
    description:
      "A plain curriculum vitae for academic jobs, grants, and promotion files, with standard sections, citations with hanging indents, and your name and page number in the footer.",
    sizeLabel: "Portrait, 2 or more pages",
    fold: "none",
    plain: true,
    page: { size: "paper", orientation: "portrait", margins: { top: 54, right: 60, bottom: 54, left: 60 } },
    customize: CUSTOMIZE,
    print: { duplex: "long", paper: "Plain white paper, printed on both sides for a committee copy." },
    footer: { text: "{name}  |  Curriculum vitae  |  Page {page} of {pages}", align: "right" },
    options: [HEADING_OPTION],
    build(ctx) {
      const k = kit(ctx);
      return [
        k.name(),
        k.headline("{title}, Department name, University name"),
        k.contact("{address}, {city}\n{phone}  |  {email}  |  {web}"),
        k.heading("Education"),
        k.school("Doctoral degree", "University name", "2018", "Dissertation: Title of the dissertation. Advisor: Professor Name."),
        k.school("Master's degree", "University name", "2013"),
        k.school("Bachelor's degree", "University name", "2011"),
        k.heading("Academic Appointments"),
        k.text("**Job title**, Department name, University name, 2021 to present", { after: 3 }),
        k.text("**Job title**, Department name, University name, 2018 to 2021"),
        k.heading("Research Interests"),
        k.text(alt(1, 8)),
        k.heading("Publications"),
        k.sub("Peer reviewed journal articles"),
        k.cite("Author, A., Author, B., & Author, C. (2025). Title of the article. Journal Name, 41(2), 112 to 129."),
        k.cite("Author, A., & Author, B. (2024). Title of the article. Journal Name, 18(4), 455 to 470."),
        k.cite("Author, A., Author, B., & Author, C. (2023). Title of the article. Journal Name, 9, 23 to 38."),
        k.cite("Author, A. (2021). Title of the article. Journal Name, 37(1), 1 to 19."),
        k.sub("Book chapters"),
        k.cite("Author, A. (2024). Title of the chapter. In A. Editor & B. Editor (Eds.), Title of the book (pp. 77 to 98). Publisher Name."),
        k.sub("Manuscripts under review"),
        k.cite("Author, A., & Author, B. Title of the manuscript. Submitted to Journal Name."),
        k.heading("Grants and Fellowships"),
        k.bullets([
          "Grant name, Principal Investigator, 2024 to 2027, 248,000 dollars",
          "Grant name, University name, 2022, 15,000 dollars",
          "Fellowship name, University name, 2014 to 2017",
        ]),
        k.heading("Honors and Awards"),
        k.bullets(["Award name, University name, 2024", "Award name, Conference name, 2017"]),
        k.heading("Invited Talks"),
        k.bullets([
          "Title of the talk. Department seminar, University name, 2025",
          "Title of the talk. Public lecture series, Organization name, 2023",
        ]),
        k.heading("Conference Presentations"),
        k.cite("Author, A., & Author, B. (2025, April). Title of the presentation [Paper presentation]. Conference Name, City, ST."),
        k.cite("Author, A., & Author, B. (2024, October). Title of the poster [Poster]. Conference Name."),
        k.heading("Teaching Experience"),
        k.text("**University name**", { after: 2 }),
        k.bullets([
          "Course name (DEPT 101), about 120 students, every fall since 2021",
          "Course name (DEPT 340), with labs, every spring since 2022",
          "Course name (DEPT 610), 2023 and 2025",
        ]),
        k.heading("Mentoring"),
        k.bullets(["Graduate students advised: 3 master's students (2 completed)", "Undergraduate researchers supervised: 11 since 2021"]),
        k.heading("Service"),
        k.bullets([
          "Reviewer for Journal Name and Journal Name",
          "Committee name, member, 2022 to present",
          "Volunteer role, Organization name, 2021 to present",
        ]),
        k.heading("Professional Memberships"),
        k.text("Organization name; Organization name"),
        k.heading("References"),
        k.text("Professor Name, University name, email address\nProfessor Name, University name, email address\nDr. Name, Organization name, email address", { after: 0 }),
      ];
    },
  },

  {
    id: "cv-europe",
    styles: resumeStyles,
    name: "UK and Europe CV",
    category: "resumes",
    description:
      "A two page CV in the style used in the United Kingdom and much of Europe: a personal profile, employment history, qualifications, and language levels on the common European scale. No photo or date of birth, as UK guidance advises.",
    sizeLabel: "Portrait, 2 pages",
    fold: "none",
    plain: true,
    page: PAGE,
    customize: CUSTOMIZE,
    print: PRINT,
    footer: { text: "{name}  |  CV  |  Page {page} of {pages}", align: "right" },
    options: [HEADING_OPTION],
    build(ctx) {
      const k = kit(ctx);
      return [
        k.name(),
        k.headline("{title}"),
        k.contact("{address}, {city}\n{phone}  |  {email}  |  {profile}"),
        k.heading("Personal Profile"),
        k.text(alt(2, 12)),
        k.heading("Key Skills"),
        k.bullets([
          "Project planning and budgeting",
          "Partnership building",
          "Event management",
          "Office and survey software",
        ]),
        k.heading("Employment History"),
        k.job("Job title", "Organisation name, City, Country", "September 2022 to present", altList(3, 3)),
        k.job("Job title", "Organisation name, City, Country", "March 2019 to August 2022", altList(3, 6)),
        k.job("Job title", "Organisation name, City, Country", "July 2017 to February 2019", altList(1, 9)),
        k.heading("Education and Qualifications"),
        k.school("Master's degree", "University name, Country", "September 2016 to September 2017", "Merit. Dissertation: Title of the dissertation."),
        k.school("Bachelor's degree (Hons)", "University name, Country", "September 2012 to June 2016", "Upper second class honours (2:1). Year abroad in City, Country."),
        k.heading("Languages"),
        k.bullets(["English: native", "French: C1 (proficient)", "Spanish: B2 (upper intermediate)"]),
        k.heading("Professional Development"),
        k.bullets(["Certificate name, 2023", "Course name, 2022", "Course name, renewed 2024"]),
        k.heading("Interests"),
        k.text(alt(1, 10)),
        k.heading("References"),
        k.text("Available on request.", { after: 0 }),
      ];
    },
  },
];
