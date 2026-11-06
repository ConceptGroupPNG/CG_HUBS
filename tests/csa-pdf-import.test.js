// Tests for CSA HUB "Import from PDF" (csa/pdf-import.js) - run with:  node --test tests/
// The fixtures are the text pdf.js extracts from a BHC Proposal letter and a BHC CSA,
// with the client, people and amounts replaced by made-up ones (this repo is public).
const { test } = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const path = require("path");
const I = require("../csa/pdf-import.js");

const fixture = name => fs.readFileSync(path.join(__dirname, "fixtures", name), "utf8");
const CSA = fixture("csa-import-csa.txt");
const PROPOSAL = fixture("csa-import-proposal.txt");

test("helpers: dates, money, title case", () => {
  assert.equal(I.parseDate("8 September 2026"), "2026-09-08");
  assert.equal(I.parseDate("12 January 2026"), "2026-01-12");
  assert.equal(I.parseDate("no date here"), "");
  assert.equal(I.parseMoney("PNGK 21,350"), 21350);
  assert.equal(I.parseMoney("Our Fees - K13,500 plus GST"), 13500);
  assert.equal(I.parseMoney("none"), null);
  assert.equal(I.titleCase("PACIFIC HARBOUR BOARD CEO REMUNERATION REVIEW"), "Pacific Harbour Board CEO Remuneration Review");
});

test("detects which kind of document it is", () => {
  assert.equal(I.detectType(CSA), "csa");
  assert.equal(I.detectType(PROPOSAL), "proposal");
  assert.equal(I.detectType("just some text"), "");
});

test("CSA: header, client, scope, fees, milestones, signatories", () => {
  const r = I.parse(CSA);
  assert.equal(r.type, "csa");
  assert.deepEqual(r.warnings, []);
  assert.equal(r.originalNumber, "BCSA 2026-0101");
  assert.equal(r.client.suggestedTitle, "Pacific Harbour Board (PHB)");
  assert.equal(r.client.address, "PO Box 1234 Port Moresby, NCD");

  const f = r.fields;
  assert.equal(f.ProjectName, "Pacific Harbour Board CEO Remuneration Review");
  assert.equal(f.Description, "PHB CEO Remuneration Review");
  assert.equal(f.FeeBasis, "Fixed Price");
  assert.equal(f.CommencementBasis, "On Payment of Invoice");
  assert.equal(f.TerminationDate, "Not Required");
  assert.equal(f.GSTRate, "10%");
  assert.equal(f.TotalFeeExclGST, 26000);
  assert.equal(f.GSTAmount, 2600);
  assert.equal(f.TotalInclGST, 28600);
  assert.match(f.ScopeOfWorks, /^BHC will provide a Salary Benchmarking Report/);
  assert.match(f.ScopeOfWorks, /\n7\. +Prepare a Remuneration Submission .*consideration\nBased on Job Evaluation Results/);
  assert.doesNotMatch(f.ScopeOfWorks, /Professional Fees|K 12,000/);

  assert.deepEqual(r.feeLines, [
    { description: "Job Evaluation & Remuneration Review for CEO Role", amount: 12000 },
    { description: "Board Report detailing Findings & Recommendations", amount: 6000 },
    { description: "Prepare an SRC / SCMC Submission, if required", amount: 8000 }
  ]);
  assert.equal(r.milestones.length, 2);
  assert.deepEqual(r.milestones.map(m => [m.percentage, m.fee, m.gst]), [[70, 18200, 1820], [30, 7800, 780]]);
  assert.equal(r.milestones[0].description, "70% on acceptance of CSA for mobilisation and preparation of Draft JE & REM Review Report");
  assert.equal(r.milestones[0].deliveryDate, "Draft Report within 10 days of receipt of mobilization payment and required information");
  assert.equal(r.milestones[1].description, "30% at submission of Draft Final Report");
  assert.equal(r.milestones[1].deliveryDate, "");

  assert.equal(f.ClientSignatoryName, "Jane Kapi");
  assert.equal(f.ClientSignatoryTitle, "Acting Board Secretary");
  assert.equal(f.ClientSignedDate, "");
  assert.equal(f.BHCSignatoryName, "Brendan Coombs");
  assert.equal(f.BHCSignatoryTitle, "Managing Director, BHC");
  assert.equal(f.BHCSignedDate, "2026-03-03");
  assert.equal(f.Status, "Sent for Signature");
});

test("Proposal: date, recipient, client, intro, scope, fees, payment terms", () => {
  const r = I.parse(PROPOSAL);
  assert.equal(r.type, "proposal");
  assert.deepEqual(r.warnings, []);
  const f = r.fields;
  assert.equal(f.ProposalDate, "2026-01-12");
  assert.equal(f.RecipientName, "Mr Jane Kapi");
  assert.equal(f.RecipientTitle, "Acting Board Secretary");
  assert.equal(f.ProjectName, "PHB CEO REM Review");
  assert.equal(r.client.suggestedTitle, "Pacific Harbour Board (PHB)");
  assert.equal(r.client.address, "PO Box 1234, Port Moresby");
  assert.equal(r.client.contact, "Jane Kapi");

  // Intro skips the subject and the standard thank-you, and leaves out the letterhead
  assert.match(f.Introduction, /^Our response consists of:/);
  assert.doesNotMatch(f.Introduction, /Thank you|Subject|Ph:|Airvos|^PNG Agents for$/m);
  assert.match(f.Introduction, /including:\n\n• AAA\n• BBBB\n• CCCC\n/);

  assert.match(f.ScopeOfWorks, /^Brian Hart Consulting Limited will:\n1\. +Conduct a review/);
  assert.match(f.ScopeOfWorks, /4\. +Prepare a Submission to the Board recommending a competitive remuneration package\.$/);
  assert.doesNotMatch(f.ScopeOfWorks, /Page Number|Professional Fees/);

  assert.deepEqual(r.feeLines, [
    { description: "Job Evaluation and Remuneration Review for the CEO role", amount: 12000 },
    { description: "Produce a Report to the Board detailing our Findings and Recommendations", amount: 6000 },
    { description: "Prepare and SRC/SCMC Submission", amount: 8000 }
  ]);
  assert.equal(f.TotalFeeExclGST, 26000);
  assert.equal(f.GSTAmount, 2600);
  assert.equal(f.TotalInclGST, 28600);
  assert.deepEqual(r.paymentTerms.map(t => [t.description, t.percentage, t.fee, t.gst]), [
    ["70% upon acceptance of the proposal", 70, 18200, 1820],
    ["30% when we supply the Consultation Draft Final Report", 30, 7800, 780]
  ]);
});

test("matching the client to CSA HUB's client list", () => {
  const clients = [
    { id: "1", fields: { Title: "Bank South Pacific (BSP)" } },
    { id: "2", fields: { Title: "Pacific Harbour Board (PHB)" } },
    { id: "3", fields: { Title: "Other Harbour Ltd (PHB2)" } }
  ];
  assert.equal(I.matchClient({ name: "Pacific Harbour Board", abbr: "PHB" }, clients).id, "2");
  assert.equal(I.matchClient({ name: "PACIFIC HARBOUR BOARD LIMITED", abbr: "" }, clients.map(c => ({ id: c.id, fields: { Title: c.fields.Title.toUpperCase() } }))).id, "2");
  assert.equal(I.matchClient({ name: "Ports Authority", abbr: "BSP" }, clients).id, "1");   // by short name
  assert.equal(I.matchClient({ name: "Unknown Org", abbr: "UO" }, clients), null);
});

test("a CSA where the money columns are missing reports it instead of guessing", () => {
  const r = I.parse(CSA.split("FEE BASIS")[0], "csa");
  assert.ok(r.warnings.some(w => /fee \/ payment table/.test(w)));
  assert.equal(r.milestones.length, 0);
});
