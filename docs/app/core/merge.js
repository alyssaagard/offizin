/**
 * Mail merge: turns a pasted list into records for templates that print one item per person
 * (labels, name badges, certificates, envelopes, tickets).
 *
 * People paste from Excel or Google Sheets (tab separated), from a .csv file (comma or
 * semicolon separated), or type one name per line. If the first row looks like headings,
 * it is used to match columns to the template's fields; otherwise columns are taken in order.
 */

/** Splits one line of delimited text, honoring double quotes ("Smith, Jr."). */
function splitLine(line, delim) {
  const out = [];
  let cur = "";
  let quoted = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (quoted) {
      if (ch === '"' && line[i + 1] === '"') {
        cur += '"';
        i++;
      } else if (ch === '"') {
        quoted = false;
      } else {
        cur += ch;
      }
    } else if (ch === '"' && cur.trim() === "") {
      quoted = true;
      cur = "";
    } else if (ch === delim) {
      out.push(cur.trim());
      cur = "";
    } else {
      cur += ch;
    }
  }
  out.push(cur.trim());
  return out;
}

function detectDelimiter(lines) {
  const sample = lines.slice(0, 5).join("\n");
  if (sample.includes("\t")) return "\t";
  const commas = (sample.match(/,/g) || []).length;
  const semis = (sample.match(/;/g) || []).length;
  if (semis > commas) return ";";
  if (commas) return ",";
  return null; // one value per line
}

const simplify = (s) => String(s || "").toLowerCase().replace(/[^a-z0-9]/g, "");

/**
 * Parses pasted text into records for `fields` ([{ key, label, aliases? }]).
 * Returns { records, matched: [field keys in column order], header: boolean, warnings: [] }.
 */
export function parseRecords(text, fields) {
  const warnings = [];
  const lines = String(text || "")
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .filter((l) => l.trim() !== "");
  if (!lines.length) return { records: [], matched: [], header: false, warnings };
  const delim = detectDelimiter(lines);
  const rows = lines.map((l) => (delim ? splitLine(l, delim) : [l.trim()]));

  // Does the first row name the fields?
  const first = rows[0].map(simplify);
  const names = fields.map((f) => [f.key, f.label, ...(f.aliases || [])].map(simplify));
  const hits = first.map((h) => names.findIndex((n) => n.includes(h)));
  const header = hits.filter((i) => i >= 0).length >= Math.min(2, fields.length) || (fields.length === 1 && hits[0] === 0 && rows.length > 1);
  let matched;
  let body = rows;
  if (header) {
    matched = hits.map((i) => (i >= 0 ? fields[i].key : null));
    body = rows.slice(1);
    const missing = fields.filter((f) => !matched.includes(f.key)).map((f) => f.label);
    if (missing.length) warnings.push(`No column found for: ${missing.join(", ")}.`);
  } else {
    matched = fields.map((f) => f.key).slice(0, Math.max(...rows.map((r) => r.length)));
  }

  const records = [];
  for (const row of body) {
    const rec = {};
    row.forEach((value, i) => {
      const key = matched[i];
      if (key) rec[key] = rec[key] ? `${rec[key]}\n${value}` : value;
    });
    if (Object.values(rec).some((v) => v)) records.push(rec);
  }
  if (records.length > 1000) {
    warnings.push(`Only the first 1,000 of ${records.length} rows are used.`);
    records.length = 1000;
  }
  return { records, matched, header, warnings };
}

/** Turns records back into tab-separated text with a header row, for editing in the panel. */
export function recordsToText(records, fields) {
  const head = fields.map((f) => f.label).join("\t");
  const rows = records.map((r) => fields.map((f) => String(r[f.key] || "").replace(/[\t\n]+/g, " ")).join("\t"));
  return [head, ...rows].join("\n");
}
