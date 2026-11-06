/* =====================================================================
   CSA HUB - Import from PDF
   ---------------------------------------------------------------------
   Reads a BHC Proposal letter or a BHC Consulting Services Agreement
   (CSA) PDF and turns it into the same fields the New Proposal / New
   Agreement forms use. Nothing here talks to SharePoint: the page fills
   the form and the person checks it before clicking Create.

   Works in the browser (window.CSAImport) and in Node (module.exports)
   so the parsing can be tested without a browser.

   Pattern-matching only - no AI, nothing leaves the browser.
     - CSAs follow the BHC template (labelled boxes, fee table), so they
       read reliably.
     - Proposals are letters, so the date, recipient, fees, total and
       payment split read reliably; introduction and scope usually need
       a light tidy.
   ===================================================================== */
(function (global) {
  "use strict";

  const MONTHS = ["january","february","march","april","may","june","july",
                  "august","september","october","november","december"];

  // Words kept in capitals when converting "ALL CAPS" text to Title Case.
  const ACRONYMS = new Set(["CEO","CFO","COO","HR","ICT","IT","PNG","NCD","PO","GST","SRC","SCMC","JE",
    "BHC","KF","REM","LMS","SLA","CSA","ERP","OHS","HRIS","MD","GM","ID","II","III","IV","AGM","SOE"]);
  const SMALL_WORDS = new Set(["of","and","the","for","to","in","on","at","by","a","an","or","with"]);

  /* ---------------- PDF -> text (layout kept) ---------------- */

  // Rebuilds each page as lines of text, placing words at character
  // columns from their x position so table columns stay lined up.
  async function textFromPdf(pdfjsLib, data) {
    const doc = await pdfjsLib.getDocument({ data, isEvalSupported: false, disableFontFace: true }).promise;
    const pages = [];
    for (let n = 1; n <= doc.numPages; n++) {
      const page = await doc.getPage(n);
      const content = await page.getTextContent();
      const items = content.items
        .filter(it => it.str && it.str.trim() !== "")
        .map(it => ({ str: it.str, x: it.transform[4], y: it.transform[5], w: it.width,
                      h: Math.abs(it.transform[3]) || Math.abs(it.transform[0]) || 10 }));
      if (!items.length) { pages.push(""); continue; }

      // Typical character width on this page.
      const widths = items.filter(i => i.str.length > 2 && i.w > 0).map(i => i.w / i.str.length).sort((a, b) => a - b);
      const cw = widths.length ? widths[Math.floor(widths.length / 2)] : 5;

      // Group into lines by baseline (top of page first).
      items.sort((a, b) => (b.y - a.y) || (a.x - b.x));
      const lines = [];
      for (const it of items) {
        const line = lines.find(l => Math.abs(l.y - it.y) < Math.max(2, it.h * 0.35));
        if (line) line.items.push(it); else lines.push({ y: it.y, items: [it] });
      }
      lines.sort((a, b) => b.y - a.y);

      const out = [];
      let prevY = null;
      for (const l of lines) {
        l.items.sort((a, b) => a.x - b.x);
        let text = "", endX = null;
        for (const it of l.items) {
          let col = Math.max(0, Math.round(it.x / cw));
          if (endX !== null) {
            const gap = it.x - endX;
            if (gap > cw * 1.5) col = Math.max(col, text.length + 2);   // a real column gap
            else if (gap > cw * 0.25) col = text.length + 1;           // ordinary word space
            else col = text.length;                                     // same word continues
          }
          if (col > text.length) text += " ".repeat(col - text.length);
          text += it.str;
          endX = it.x + it.w;
        }
        // Blank line where there is a visible vertical gap.
        if (prevY !== null && prevY - l.y > (l.items[0].h || 10) * 1.9) out.push("");
        out.push(text.replace(/\s+$/, ""));
        prevY = l.y;
      }
      pages.push(out.join("\n"));
    }
    return pages.join("\n\f\n");
  }

  /* ---------------- small helpers ---------------- */

  function clean(s) {
    return String(s || "")
      .replace(/[‘’]/g, "'").replace(/[“”]/g, '"')
      .replace(/[–—]/g, "-").replace(/ /g, " ").replace(/|●|▪|•/g, "•")
      .replace(/\r/g, "")
      .replace(/\b((?:PNG\s*)?K)\s{2,}(?=\d)/g, "$1 ");
  }

  function parseMoney(s) {
    const m = /(?:PNG\s*)?K\s*([\d,]+(?:\.\d+)?)/i.exec(String(s || ""));
    return m ? parseFloat(m[1].replace(/,/g, "")) : null;
  }

  // "10 July 2026" / "8 September 2026" -> "2026-07-10" (built from parts, no Date objects)
  function parseDate(s) {
    const m = /(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]+)\s*,?\s+(\d{4})/.exec(String(s || ""));
    if (!m) return "";
    const mon = MONTHS.findIndex(x => x.startsWith(m[2].toLowerCase().slice(0, 3)));
    if (mon < 0) return "";
    return `${m[3]}-${String(mon + 1).padStart(2, "0")}-${String(m[1]).padStart(2, "0")}`;
  }

  // "NATIONAL PETROLEUM AUTHORITY CEO REMUNERATION REVIEW" -> "National Petroleum Authority CEO Remuneration Review"
  function titleCase(s, keep) {
    const extra = new Set((keep || []).map(k => String(k).toUpperCase()));
    const words = String(s || "").trim().split(/\s+/);
    return words.map((w, i) => {
      const bare = w.replace(/[^A-Za-z0-9&]/g, "").toUpperCase();
      if (!bare) return w;
      if (ACRONYMS.has(bare) || extra.has(bare)) return w.toUpperCase();
      if (/\d/.test(w)) return w.toUpperCase();
      const lower = w.toLowerCase();
      if (i > 0 && SMALL_WORDS.has(lower)) return lower;
      return lower.replace(/(^|[-(/'"])([a-z])/g, (_, p, c) => p + c.toUpperCase());
    }).join(" ");
  }
  const isMostlyUpper = s => { const l = s.replace(/[^A-Za-z]/g, ""); return l.length > 3 && l === l.toUpperCase(); };
  const tidyCase = (s, keep) => isMostlyUpper(s) ? titleCase(s, keep) : String(s || "").trim();

  // Join wrapped lines back into paragraphs; numbered / bulleted items stay on their own lines.
  function reflow(rawLines) {
    // A row like "•  MRA      •  CCDA      •  ICAC" becomes three bullet lines.
    const lines = [];
    for (const raw of rawLines) {
      if ((raw.match(/•/g) || []).length > 1) {
        const ind = raw.match(/^\s*/)[0];
        raw.trim().split(/\s{2,}(?=•)/).forEach(p => lines.push(ind + p));
      } else lines.push(raw);
    }
    const out = [];
    let itemIndent = null;     // indent of the last numbered / bulleted item
    for (const raw of lines) {
      const line = raw.trim();
      const indent = raw.length - raw.replace(/^\s+/, "").length;
      if (!line) { if (out.length && out[out.length - 1] !== "") out.push(""); itemIndent = null; continue; }
      const startsItem = /^(\d+[.)]|[a-z][.)]|•|-\s|\*)\s*/.test(line);
      const prev = out.length ? out[out.length - 1] : "";
      const leftItem = !startsItem && itemIndent !== null && indent <= itemIndent;
      if (!prev || startsItem || leftItem || (/[:.;]$/.test(prev) && /^[A-Z]/.test(line))) out.push(line);
      else out[out.length - 1] = prev + " " + line;
      itemIndent = startsItem ? indent : (leftItem ? null : itemIndent);
    }
    while (out.length && out[out.length - 1] === "") out.pop();
    return out.join("\n").replace(/\n{3,}/g, "\n\n").replace(/^•\s*/gm, "• ").trim();
  }

  // Lines repeated on several pages (letterhead, running header/footer).
  function repeatedLines(pages) {
    const seen = new Map();
    pages.forEach(p => new Set(p.split("\n").map(l => l.trim()).filter(Boolean)).forEach(l => seen.set(l, (seen.get(l) || 0) + 1)));
    return new Set([...seen].filter(([, n]) => n >= 2 && pages.length >= 2).map(([l]) => l));
  }

  // Splits a layout line into its text blocks (separated by 2+ spaces) with their columns.
  function segments(line) {
    const out = [], re = /\S+(?: \S+)*/g;
    let m;
    while ((m = re.exec(line))) out.push({ text: m[0], start: m.index, end: m.index + m[0].length });
    return out;
  }

  function splitClient(raw) {
    // "NATIONAL PETROLEUM AUTHORITY ("NPA")" -> { name, abbr }
    const s = clean(raw).trim();
    const m = /^(.*?)\s*\(\s*"?\s*([A-Za-z0-9&.\- ]{2,15}?)\s*"?\s*\)\s*$/.exec(s);
    const abbr = m ? m[2].replace(/\s+/g, "").toUpperCase() : "";
    const name = tidyCase(m ? m[1] : s, [abbr]);
    return { name, abbr, suggestedTitle: abbr ? `${name} (${abbr})` : name };
  }

  /* ---------------- detect document type ---------------- */
  function detectType(text) {
    const t = clean(text);
    if (/Consulting Services Agreement/i.test(t) && /AGREEMENT No\s*:/i.test(t)) return "csa";
    if (/Dear\s+\S+/i.test(t) && /(Professional Fees|Scope of Works|proposal)/i.test(t)) return "proposal";
    return "";
  }

  /* ---------------- CSA ---------------- */
  function parseCsa(text) {
    const t = clean(text);
    const lines = t.split("\n");
    const warnings = [];
    const findIdx = (re, from = 0) => { for (let i = from; i < lines.length; i++) if (re.test(lines[i])) return i; return -1; };

    // Agreement number + short label, e.g. "BCSA 2026-0809 - NPA - CEO REMUNERATION REVIEW"
    const noLine = (/AGREEMENT No\s*:\s*(.+)/i.exec(t) || [])[1] || "";
    const noParts = noLine.split(/\s+-\s+/).map(s => s.trim()).filter(Boolean);
    const originalNumber = noParts[0] || "";

    // Client (AND:) and its address on the next line
    let client = { name: "", abbr: "", suggestedTitle: "" }, clientAddress = "";
    const andIdx = findIdx(/^\s*AND\s*:/i);
    if (andIdx >= 0) {
      client = splitClient(lines[andIdx].replace(/^\s*AND\s*:/i, ""));
      const next = (lines[andIdx + 1] || "").trim();
      if (next && !/^AGREEMENT\s*:/i.test(next)) clientAddress = tidyCase(next, [client.abbr]);
    } else warnings.push("Couldn't find the client (AND:) line.");

    // Project name: between SERVICES SCHEDULE and SCOPE OF WORKS, minus the label words
    let projectName = "";
    const ssIdx = findIdx(/^\s*SERVICES SCHEDULE\s*$/);
    const sowIdx = findIdx(/^\s*SCOPE OF WORKS/i, ssIdx + 1);
    if (ssIdx >= 0 && sowIdx > ssIdx) {
      projectName = lines.slice(ssIdx + 1, sowIdx).join(" ")
        .replace(/\bPROJECT\b/, "").replace(/\bNAME\s*:/i, "").replace(/\s+/g, " ").trim();
      projectName = tidyCase(projectName, [client.abbr]);
    }
    if (!projectName) warnings.push("Couldn't read the project name.");

    // Scope of works: from SCOPE OF WORKS to the end of the schedule, without the fee bullets
    let scope = "", feeLines = [];
    if (sowIdx >= 0) {
      let end = findIdx(/^\s*(Commencement|FEE BASIS|©|\f)/i, sowIdx + 1);
      if (end < 0) end = lines.length;
      const body = [], fees = [];
      let inFees = false;
      for (let i = sowIdx; i < end; i++) {
        let l = lines[i].replace(/^\s*SCOPE OF WORKS/i, "");
        if (/Professional Fees\s*:?/i.test(l)) { inFees = true; continue; }
        if (inFees) {
          if (!l.trim()) { if (fees.length) inFees = false; continue; }
          if (/K\s*[\d,]+/.test(l)) { fees.push(l); continue; }
          if (/^\s*•/.test(l) || fees.length === 0) { fees.push(l); continue; }
          inFees = false;
        }
        body.push(l);
      }
      scope = reflow(body);
      // Fee bullets: "• Job Evaluation & Remuneration Review for CEO Role:   K 13,500 plus GST"
      let pending = "";
      for (const f of fees) {
        const amt = parseMoney(f);
        const desc = f.replace(/(?:PNG\s*)?K\s*[\d,]+(?:\.\d+)?.*$/i, "").replace(/^\s*•\s*/, "").trim();
        if (amt === null) { pending = (pending + " " + desc).trim(); continue; }
        const d = (pending ? pending + " " + desc : desc).replace(/\s*:\s*$/, "").replace(/\s+/g, " ").trim();
        feeLines.push({ description: d, amount: amt });
        pending = "";
      }
    }
    if (!scope) warnings.push("Couldn't read the scope of works.");

    // Commencement / termination
    let commencementBasis = "On Payment of Invoice", commencementDate = "", terminationDate = "Not Required";
    const cm = /Commencement\s+(.+?)\s{2,}Termination Date\s+(.+)$/im.exec(t);
    if (cm) {
      const c = cm[1].trim();
      if (/payment of invoice/i.test(c)) commencementBasis = "On Payment of Invoice";
      else { commencementBasis = "Fixed Date"; commencementDate = parseDate(c) || c; }
      const term = cm[2].trim();
      terminationDate = parseDate(term) || term;
    }

    const fb = /FEE BASIS\s*:\s*(.+)/i.exec(t);
    const feeBasis = fb && /(T\s*&\s*M|Time\s*(&|and)\s*Materials)/i.test(fb[1]) ? "Time & Materials" : "Fixed Price";

    // Payment milestones from the fee table
    const milestones = [];
    let totalFee = null, totalGst = null;
    const hdrIdx = findIdx(/Applicable Fee/i);
    const totIdx = findIdx(/^\s*Total\b.*K\s*[\d,]+/i, hdrIdx + 1);
    if (hdrIdx >= 0) {
      const hdr = lines[hdrIdx];
      const feeCol = hdr.search(/Applicable Fee/i);
      const gstCol = hdr.search(/\bGST\b/);
      const stop = totIdx > hdrIdx ? totIdx : findIdx(/Executed by/i, hdrIdx + 1);
      let cur = null, lastDelivery = null;
      for (let i = hdrIdx + 1; i < (stop > 0 ? stop : lines.length); i++) {
        const segs = segments(lines[i]);
        for (const s of segs) {
          const money = parseMoney(s.text);
          const isMoney = money !== null && /^(PNG\s*)?K\s*[\d,]/i.test(s.text);
          if (isMoney) {
            if (!cur) continue;
            if (s.start < gstCol - 2 && cur.fee === null) cur.fee = money;
            else cur.gst = money;
            continue;
          }
          if (s.end <= feeCol + 2) {                       // description column
            if (/^Payment schedule\s*:?$/i.test(s.text) || /^Fixed Price Quote and$/i.test(s.text)) continue;
            if (/^\d+(\.\d+)?\s*%/.test(s.text) || !cur) {
              cur = { description: s.text, percentage: parseFloat(s.text) || 0, fee: null, gst: null, deliveryDate: "" };
              milestones.push(cur);
            } else cur.description += " " + s.text;
          } else if (s.start > gstCol) {                   // delivery column
            if (/^Delivery Date|^\(Estimate\)$/i.test(s.text)) continue;
            const target = /^[a-z]/.test(s.text) && lastDelivery ? lastDelivery : cur;
            if (!target) continue;
            target.deliveryDate = (target.deliveryDate ? target.deliveryDate + " " : "") + s.text;
            lastDelivery = target;
          }
        }
      }
      if (totIdx > 0) {
        const amounts = segments(lines[totIdx]).map(s => parseMoney(s.text)).filter(v => v !== null);
        if (amounts.length) totalFee = amounts[0];
        if (amounts.length > 1) totalGst = amounts[1];
      }
    } else warnings.push("Couldn't find the fee / payment table.");

    milestones.forEach(m => { m.fee = m.fee || 0; m.gst = m.gst || 0; m.description = m.description.replace(/\s+/g, " ").trim(); m.deliveryDate = m.deliveryDate.replace(/\s+/g, " ").trim(); });
    if (totalFee === null) totalFee = feeLines.reduce((s, l) => s + l.amount, 0) || milestones.reduce((s, m) => s + m.fee, 0);
    if (totalGst === null) totalGst = milestones.length ? milestones.reduce((s, m) => s + m.gst, 0) : Math.round(totalFee * 10) / 100;
    const gstRate = totalFee && totalGst === 0 ? "0% (S65)" : "10%";

    const lineSum = feeLines.reduce((s, l) => s + l.amount, 0);
    if (feeLines.length && Math.abs(lineSum - totalFee) > 0.5)
      warnings.push(`Fee lines add up to K${lineSum.toLocaleString()} but the total is K${totalFee.toLocaleString()}.`);
    const msSum = milestones.reduce((s, m) => s + m.fee, 0);
    if (milestones.length && Math.abs(msSum - totalFee) > 0.5)
      warnings.push(`Milestones add up to K${msSum.toLocaleString()} but the total is K${totalFee.toLocaleString()}.`);

    // Signatories: "Name:  CLIENT  Name:  BHC" etc.
    const execIdx = findIdx(/Executed by/i);
    const sigText = execIdx >= 0 ? lines.slice(execIdx).join("\n") : t;
    const sig = label => {
      const m = new RegExp("^\\s*" + label + "\\s*:(.*)$", "m").exec(sigText);
      if (!m) return ["", ""];
      const parts = m[1].split(new RegExp("\\b" + label + "\\s*:"));
      return [(parts[0] || "").trim(), (parts[1] || "").trim()];
    };
    const [cName, bName] = sig("Name");
    const [cTitle, bTitle] = sig("Title");
    const [cDate, bDate] = sig("Date");
    const clientSignedDate = parseDate(cDate), bhcSignedDate = parseDate(bDate);
    const status = clientSignedDate && bhcSignedDate ? "Signed" : (clientSignedDate || bhcSignedDate) ? "Sent for Signature" : "Draft";

    // Short label for the Time Tracker dropdown, e.g. "NPA CEO Remuneration Review"
    let description = noParts.length > 1 ? tidyCase(noParts.slice(1).join(" "), [client.abbr]) : (projectName || "");
    if (description.length > 60) description = description.slice(0, 60).trim();

    return {
      type: "csa",
      originalNumber,
      client: Object.assign(client, { address: clientAddress, contact: tidyCase(cName) }),
      fields: {
        ProjectName: projectName,
        Description: description,
        Status: status,
        FeeBasis: feeBasis,
        GSTRate: gstRate,
        CommencementBasis: commencementBasis,
        CommencementDate: commencementDate,
        TerminationDate: terminationDate,
        TotalFeeExclGST: totalFee || 0,
        GSTAmount: totalGst || 0,
        TotalInclGST: Math.round(((totalFee || 0) + (totalGst || 0)) * 100) / 100,
        ScopeOfWorks: scope,
        Deliverables: "",
        ClientSignatoryName: tidyCase(cName),
        ClientSignatoryTitle: tidyCase(cTitle, [client.abbr]),
        ClientSignedDate: clientSignedDate,
        BHCSignatoryName: tidyCase(bName),
        BHCSignatoryTitle: tidyCase(bTitle),
        BHCSignedDate: bhcSignedDate
      },
      feeLines,
      milestones: milestones.map(m => Object.assign(m, { status: "Pending" })),
      warnings
    };
  }

  /* ---------------- Proposal ---------------- */
  function parseProposal(text) {
    const t = clean(text);
    const pages = t.split("\f");
    const repeated = repeatedLines(pages);
    const isNoise = l => {
      const s = l.trim();
      return repeated.has(s) || /^Page( Number)?\s*:?\s*\d+/i.test(s) || /^Page \d+ of \d+$/i.test(s)
        || /\b(Ph|Tel|Fax|Email)\s*:/i.test(s) || /^Lot \d+,?\s*Section/i.test(s) || /^PNG Agents? for\b/i.test(s)
        || /^PO Box \d+,\s*Waigani$/i.test(s);
    };
    const lines = t.split("\n").filter(l => l !== "\f" && !isNoise(l)).map(l => l.replace(/\f/g, ""));
    const warnings = [];
    const findIdx = (re, from = 0) => { for (let i = from; i < lines.length; i++) if (re.test(lines[i])) return i; return -1; };

    // Date: first "10 July 2026"-style line
    const dateIdx = findIdx(/^\s*\d{1,2}(st|nd|rd|th)?\s+[A-Za-z]+\s*,?\s+\d{4}\s*$/);
    const proposalDate = dateIdx >= 0 ? parseDate(lines[dateIdx]) : "";
    if (!proposalDate) warnings.push("Couldn't read the proposal date.");

    // Recipient block: lines between the date and "Dear ..."
    const dearIdx = findIdx(/^\s*Dear\s+/i);
    const block = dateIdx >= 0 && dearIdx > dateIdx
      ? lines.slice(dateIdx + 1, dearIdx).map(l => l.trim()).filter(Boolean) : [];
    const recipientName = block[0] || "";
    const recipientTitle = block[1] || "";
    const orgLine = block[2] || "";
    const addressLines = block.slice(3).filter(l => !/^Papua New Guinea$/i.test(l));
    if (!recipientName) warnings.push("Couldn't read the recipient.");

    // Subject: may wrap onto following lines until a blank line
    const subjIdx = findIdx(/^\s*Subject\s*:/i);
    let subject = "";
    if (subjIdx >= 0) {
      const parts = [lines[subjIdx].replace(/^\s*Subject\s*:/i, "")];
      for (let i = subjIdx + 1; i < lines.length && lines[i].trim(); i++) parts.push(lines[i]);
      subject = parts.join(" ").replace(/\s+/g, " ").trim();
    }
    const abbrM = /\(([A-Z][A-Z0-9&]{1,10})\)\s*$/.exec(subject) || /\(([A-Z][A-Z0-9&]{1,10})\)/.exec(subject);
    const abbr = abbrM ? abbrM[1] : "";
    const orgName = tidyCase(orgLine, [abbr]);
    const client = {
      name: orgName, abbr,
      suggestedTitle: orgName ? (abbr ? `${orgName} (${abbr})` : orgName) : "",
      address: addressLines.join(", "),
      contact: recipientName.replace(/^(Mr|Mrs|Ms|Miss|Dr|Prof)\.?\s+/i, "")
    };

    // Project name: the running header repeated on later pages, else the subject
    let projectName = "";
    const headerCandidates = [...repeated].filter(l =>
      l.length >= 5 && l.length <= 60 && !/Brian Hart|PNG Agents|Page|Ph:|Fax|Email|Avenue|PO Box|Waigani|Korn|Hay Group|Limited|www\./i.test(l));
    if (headerCandidates.length) projectName = headerCandidates.sort((a, b) => b.length - a.length)[0];
    if (!projectName && subject) projectName = subject.replace(/^.*?(Proposal|Quotation)\s+(to|for)\s+/i, "").slice(0, 120);

    // Introduction: after the greeting, skipping the opening "Thank you" paragraph,
    // up to "Scope of Works" (or the line introducing it).
    const sowIdx = findIdx(/^\s*Scope of Works\s*:?\s*$/i);
    let introduction = "";
    if (dearIdx >= 0) {
      let end = sowIdx > dearIdx ? sowIdx : findIdx(/Professional Fees/i, dearIdx);
      if (end < 0) end = lines.length;
      // Start after the subject block if the letter puts it after the greeting
      let start = dearIdx + 1;
      if (subjIdx > dearIdx && subjIdx < end) { start = subjIdx + 1; while (start < end && lines[start].trim()) start++; }
      let body = lines.slice(start, end);
      while (body.length && /would undertake the following|following Scope of Works/i.test(body[body.length - 1] || "") === false && !body[body.length - 1].trim()) body.pop();
      if (body.length && /would undertake the following Scope of Works/i.test(body[body.length - 1])) body.pop();
      // drop the first paragraph if it is the standard thank-you
      const firstText = body.findIndex(l => l.trim());
      if (firstText >= 0 && /^\s*Thank you/i.test(body[firstText])) {
        let j = firstText; while (j < body.length && body[j].trim()) j++;
        body = body.slice(j);
      }
      introduction = reflow(body);
    }

    // Scope of works: from the heading to "Our Professional Fees"
    const feesIdx = findIdx(/Professional Fees/i);
    let scope = "";
    if (sowIdx >= 0) scope = reflow(lines.slice(sowIdx + 1, feesIdx > sowIdx ? feesIdx : lines.length));
    else warnings.push("Couldn't find a 'Scope of Works' heading.");

    // Fees: bullet description, then "Our Fees - K13,500 plus GST"
    const feeLines = [];
    let total = null;
    if (feesIdx >= 0) {
      let pending = "";
      for (let i = feesIdx + 1; i < lines.length; i++) {
        const l = lines[i].trim();
        if (!l) continue;
        if (/^(Project\s+)?Total\b/i.test(l)) { total = parseMoney(l); break; }
        if (/payment of|will require payment/i.test(l)) break;
        if (/^Our professional fees/i.test(l)) continue;
        const amt = parseMoney(l);
        if (amt === null) {
          if (/^•/.test(l)) pending = l.replace(/^•\s*/, "");
          else if (pending) pending += " " + l;
          continue;
        }
        const own = l.replace(/^•\s*/, "").replace(/(?:Our Fees?\s*[-:]?\s*)?(?:PNG\s*)?K\s*[\d,]+(?:\.\d+)?.*$/i, "").replace(/[:\-\s]+$/, "").trim();
        feeLines.push({ description: (own || pending).replace(/\s+/g, " ").trim(), amount: amt });
        pending = "";
      }
    } else warnings.push("Couldn't find the 'Professional Fees' section.");
    if (total === null) total = feeLines.reduce((s, l) => s + l.amount, 0);
    const lineSum = feeLines.reduce((s, l) => s + l.amount, 0);
    if (feeLines.length && Math.abs(lineSum - total) > 0.5)
      warnings.push(`Fee lines add up to K${lineSum.toLocaleString()} but the total is K${total.toLocaleString()}.`);

    // Payment terms: "payment of 70% of the total fee upon acceptance ... and the remaining 30% when ..."
    const paymentTerms = [];
    const flat = lines.join(" ").replace(/\s+/g, " ");
    const payM = /(?:require|requires|request)\s+payment of\s+(.+?%.+?)(?:\.\s|\.$)/i.exec(flat);
    if (payM) {
      const parts = payM[1].split(/(?:,\s*|\s+)(?:and\s+)?(?=(?:the\s+remaining\s+|a\s+further\s+|the\s+balance\s+of\s+)?\d+(?:\.\d+)?%)/i);
      for (const p of parts) {
        const pm = /(\d+(?:\.\d+)?)%\s*(.*)$/.exec(p);
        if (!pm) continue;
        const pct = parseFloat(pm[1]);
        const what = pm[2].replace(/^of the (total )?fee\s*/i, "").replace(/^(is\s+)?(payable\s+)?/i, "").replace(/[,;]\s*$/, "").trim();
        const fee = Math.round(total * pct) / 100;
        paymentTerms.push({ description: `${pct}% ${what}`.trim(), percentage: pct, fee, gst: Math.round(fee * 10) / 100, deliveryDate: "" });
      }
    }
    if (!paymentTerms.length) warnings.push("Couldn't read the payment terms.");

    const feeBasis = /time\s*(&|and)\s*materials|daily rate|hourly rate/i.test(flat) && !/fixed/i.test(flat) ? "Time & Materials" : "Fixed Price";
    const gst = Math.round(total * 10) / 100;
    let description = projectName.length > 60 ? projectName.slice(0, 60).trim() : projectName;

    return {
      type: "proposal",
      client,
      fields: {
        ProjectName: projectName,
        Description: description,
        DocumentType: "BHC Proposal",
        Status: "Sent",
        RecipientName: recipientName,
        RecipientTitle: recipientTitle,
        ProposalDate: proposalDate,
        FeeBasis: feeBasis,
        GSTRate: "10%",
        TotalFeeExclGST: total || 0,
        GSTAmount: gst,
        TotalInclGST: Math.round(((total || 0) + gst) * 100) / 100,
        Introduction: introduction,
        ScopeOfWorks: scope
      },
      feeLines,
      paymentTerms,
      warnings
    };
  }

  /* ---------------- client matching ---------------- */
  const normName = s => String(s || "").toLowerCase().replace(/\([^)]*\)/g, " ").replace(/\b(limited|ltd|inc|the|of|png|papua new guinea)\b/g, " ").replace(/[^a-z0-9]+/g, " ").trim();
  const abbrOf = s => { const m = /\(([^)]+)\)\s*$/.exec(String(s || "")); return m ? m[1].replace(/\s+/g, "").toUpperCase() : ""; };

  // clients: [{ id, fields: { Title } }] -> matching client or null
  function matchClient(parsedClient, clients) {
    if (!parsedClient) return null;
    const n = normName(parsedClient.name), a = (parsedClient.abbr || "").toUpperCase();
    return (clients || []).find(c => n && normName(c.fields.Title) === n)
        || (a ? (clients || []).find(c => abbrOf(c.fields.Title) === a) : null)
        || null;
  }

  function parse(text, forceType) {
    const type = forceType || detectType(text);
    if (type === "csa") return parseCsa(text);
    if (type === "proposal") return parseProposal(text);
    return null;
  }

  const api = { textFromPdf, parse, parseCsa, parseProposal, detectType, matchClient, parseDate, parseMoney, titleCase };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else global.CSAImport = api;
})(typeof window !== "undefined" ? window : globalThis);
