const OUTPUT_FOLDER_NAME = 'Debt Agent Output';

const PLAN = {
  tz: 'America/Chicago',
  deadlines: {
    oppose_msj: '2025-10-09',
    mov_bureaus_623: '2025-10-10',
    fdcpas: '2025-10-13',
    rebuild_day45: '2025-11-22',
  },
  court: {
    cause_no: '45D03-2504-CC-003807',
    court_name: 'Lake Superior Court, Civil Division 3',
    jurisdiction: 'Lake County, Indiana',
    efile_service_for_plaintiff: 'Blitt & Gaines, P.C., 775 Corporate Woods Pkwy, Vernon Hills, IL 60061',
  },
  party: {
    defendant_name: 'Alyssa C. Heeter',
    defendant_addr: '[Your PO Box / safe address]',
    email: '[optional]',
    phone: '[optional]',
  },
  facts: {
    chase_balance_4155: 7774.65,
    chase_last_payment_memo: { date: '2024-08-16', amount: 226.0 },
    chase_last_payment_affidavit: { date: '2024-05-16', amount: 230.0 },
    jpmcb_4978_balance: 2754,
    jpmcb_4978_dofd: '2024-07-19',
    jpmcb_4978_cl: 2300,
    jpmcb_4978_high: 13065,
    chase_4155_cl: 6300,
    amex_balance: 3013,
    bofa_balance: 6653,
    scs_comed_balance: 3477,
    credit_collections_prog_balance: 119,
    nordstrom_td_open: true,
  },
  mailing: {
    plaintiff_counsel: 'Blitt & Gaines, P.C., 775 Corporate Woods Pkwy, Vernon Hills, IL 60061',
    chase_furnisher: 'JPMorgan Chase Bank, N.A., PO Box 15369, Wilmington, DE 19850',
    bureaus: [
      'Equifax, PO Box 740241, Atlanta, GA 30374',
      'TransUnion, PO Box 2000, Chester, PA 19016',
      'Experian, PO Box 4500, Allen, TX 75013',
    ],
    scs: ['Southwest Credit Systems, 4120 International Pkwy, Carrollton, TX 75007'],
    credit_collections: ['Credit Collection Services, 725 Canton St, Norwood, MA 02062'],
  },
  offers: {
    chase_4155: [30, 37, 45],
    jpmcb_4978: [20, 30, 40],
    amex: [25, 33, 45],
    bofa: [20, 30, 42],
    scs_comed: [25, 35, 45],
    credit_collections_prog: [50, 100],
  },
  out_dir: 'out',
};

const TEMPLATES = {
  opposition_msj: `IN THE {{ court.jurisdiction | upper }} {{ court.court_name | upper }}
CAUSE NO. {{ court.cause_no }}

JPMORGAN CHASE BANK, N.A.                )
      Plaintiff,                          )
v.                                        )
{{ party.defendant_name | upper }},         )
      Defendant.                          )

DEFENDANT’S OPPOSITION TO PLAINTIFF’S MOTION FOR SUMMARY JUDGMENT

Defendant opposes Summary Judgment. A genuine dispute of material fact exists because
Plaintiff’s own designated materials conflict on the “last payment” (Memo asserts ${{ facts.chase_last_payment_memo.amount }}
on {{ facts.chase_last_payment_memo.date }}; Affidavit swears ${{ facts.chase_last_payment_affidavit.amount }} on {{ facts.chase_last_payment_affidavit.date }}),
defeating account-stated as a matter of law. Plaintiff also failed to itemize the balance and to produce the full run of
statements despite designating “Periodic Billing Statements (pp. 1–49).”

Requested relief: deny Summary Judgment; alternatively, defer under Rule 56 to compel complete statements and
itemization; set for trial.

Dated: {{ now }}

{{ party.defendant_name }}
{{ party.defendant_addr }}`,
  motion_strike_affidavit: `… MOTION TO STRIKE OR DISREGARD AFFIDAVIT PORTIONS …

Grounds: (1) internal contradiction on last-payment date/amount; (2) hearsay and lack of proper foundation for business
records; (3) no itemization of fees/interest from charge-off to present; (4) absence of complete monthly statements.

Relief: strike contradictory paragraphs; disregard balance assertions lacking itemization; order production of pp. 1–49 statements.`,
  settlement_offer: `{{ date }}

Via Email & Certified Mail
{{ to_address }}

Re: Settlement Offer (No admission) – {{ name }} – Balance ${{ balance }}

I can tender {{ pct }}% (${{ amount }}) as a lump-sum for (i) dismissal with prejudice (if in suit),
(ii) each side bears fees/costs, (iii) no post–charge-off interest or fees, (iv) furnish “settled” within 30 days,
(v) no resale of remainder; written agreement required before any payment.

Hardship summary attached.

Sincerely,
{{ party.defendant_name }}`,
  frca_611_mov: `Re: FCRA §611 Dispute & Method of Verification Request – Duplicate/Fragment JPMCB *4978 vs. Chase 4155

I dispute JPMCB *4978 as inaccurate/duplicative. *4978 shows CL ${{ facts.jpmcb_4978_cl }}, high ${{ facts.jpmcb_4978_high }},
DOFD {{ facts.jpmcb_4978_dofd }}, while the Chase/4155 lawsuit reflects a different product with CL ${{ facts.chase_4155_cl }}.
Delete or consolidate and provide MOV identifying the furnisher, contact, and specific records relied upon.`,
  frca_623_direct: `Re: Direct Dispute under FCRA §623 – Reconcile / Delete Duplicative Tradeline

Identify the correct tradeline (4155 vs. *4978). If 4155 is the live account of record, delete *4978 and correct limits/high credit/DOFD.`,
  fdCPA_1692g_pfd: `Re: Validation Request under 15 U.S.C. §1692g and Pay-for-Delete Negotiation

Please provide: (1) original contract/authority, (2) full itemization, (3) proof of assignment, (4) DOFD reported to CRAs, (5) collector license.
Cease phone contact. Upon validation, I’m prepared to resolve on a written Pay-for-Delete basis at {{ pct }}% (${{ amount }}).`,
  hardship_declaration: `Brief Declaration of Hardship

I experienced job loss (Feb 2024), caregiving for a parent with cancer, and unexpected housing remediation costs (Sept 2025).
This declaration is submitted solely to explain current inability to pay in full and to support reasonable settlement.`,
};

function makeFilings() {
  const ctx = buildContext();
  const folder = getOutputFolder();
  const cause = PLAN.court.cause_no;

  writeDocToFolder(`${cause}_Opposition_to_MSJ.docx`, renderTemplate('opposition_msj', ctx), folder);
  writeDocToFolder(`${cause}_Motion_to_Strike_Affidavit.docx`, renderTemplate('motion_strike_affidavit', ctx), folder);
  writeDocToFolder('Hardship_Declaration.docx', renderTemplate('hardship_declaration', ctx), folder);
}

function makeLetters() {
  const ctx = buildContext();
  const folder = getOutputFolder();
  const today = ctx.today;
  const facts = PLAN.facts;

  PLAN.offers.chase_4155.forEach((pct) => {
    const letterCtx = Object.assign({}, ctx, {
      date: today,
      name: 'Chase 4155',
      balance: formatMoney(facts.chase_balance_4155),
      pct: pct,
      amount: formatMoney(computeAmount(facts.chase_balance_4155, pct)),
      to_address: PLAN.mailing.plaintiff_counsel,
    });
    writeDocToFolder(
      `Chase_4155_Settlement_Offer_${pctLabel(pct)}pct_${today}.docx`,
      renderTemplate('settlement_offer', letterCtx),
      folder
    );
  });

  const movBody = renderTemplate('frca_611_mov', Object.assign({}, ctx, { date: today }));
  PLAN.mailing.bureaus.forEach((bureau) => {
    const shortName = bureau.split(',')[0].replace(/\s+/g, '');
    writeDocToFolder(
      `Bureaus_§611_MOV_JPMCB4978_${today}_${shortName}.docx`,
      movBody + `\n\nAddressed to: ${bureau}`,
      folder
    );
  });

  writeDocToFolder(
    `Chase_§623_DirectDispute_DuplicateTradeline_${today}.docx`,
    renderTemplate('frca_623_direct', Object.assign({}, ctx, { date: today })),
    folder
  );

  PLAN.offers.scs_comed.forEach((pct) => {
    const letterCtx = Object.assign({}, ctx, {
      date: today,
      pct: pct,
      amount: formatMoney(computeAmount(facts.scs_comed_balance, pct)),
      to_address: PLAN.mailing.scs[0],
      account_reference: `Southwest Credit Systems – ComEd Balance $${formatMoney(facts.scs_comed_balance)}`,
    });
    writeDocToFolder(
      `SCS_ComEd_Validation_PFD_${pctLabel(pct)}pct_${today}.docx`,
      renderTemplate('fdCPA_1692g_pfd', letterCtx),
      folder
    );
  });

  PLAN.offers.credit_collections_prog.forEach((pct) => {
    const letterCtx = Object.assign({}, ctx, {
      date: today,
      pct: pct,
      amount: formatMoney(computeAmount(facts.credit_collections_prog_balance, pct)),
      to_address: PLAN.mailing.credit_collections[0],
      account_reference: `Credit Collection Services – Progressive Balance $${formatMoney(facts.credit_collections_prog_balance)}`,
    });
    writeDocToFolder(
      `CreditCollections_Progressive_Validation_PFD_${pctLabel(pct)}pct_${today}.docx`,
      renderTemplate('fdCPA_1692g_pfd', letterCtx),
      folder
    );
  });

  [
    { name: 'Amex_15443', key: 'amex_balance', ladder: PLAN.offers.amex },
    { name: 'BofA', key: 'bofa_balance', ladder: PLAN.offers.bofa },
  ].forEach(({ name, key, ladder }) => {
    ladder.forEach((pct) => {
      const letterCtx = Object.assign({}, ctx, {
        date: today,
        name: name,
        balance: formatMoney(facts[key]),
        pct: pct,
        amount: formatMoney(computeAmount(facts[key], pct)),
        to_address: '[Creditor address here]',
      });
      writeDocToFolder(
        `${name}_Settlement_Offer_${pctLabel(pct)}pct_${today}.docx`,
        renderTemplate('settlement_offer', letterCtx),
        folder
      );
    });
  });
}

function makeCalendar() {
  const folder = getOutputFolder();
  const nowStamp = Utilities.formatDate(new Date(), 'UTC', "yyyyMMdd'T'HHmmss'Z'");
  const events = [
    buildIcsEvent('opp_msj', PLAN.deadlines.oppose_msj, 'File Opposition to MSJ + Motion to Strike', nowStamp),
    buildIcsEvent('mov_623', PLAN.deadlines.mov_bureaus_623, 'Mail §611 MOV to EQ/TU/EX + §623 to Chase', nowStamp),
    buildIcsEvent('fdcpas', PLAN.deadlines.fdcpas, 'Mail FDCPA validations to SCS and Credit Collections', nowStamp),
    buildIcsEvent('rebuild', PLAN.deadlines.rebuild_day45, 'Day 45: Rebuild plan kickoff', nowStamp),
  ];
  const ics = ['BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//Debt Agent//Apps Script//EN', `X-WR-TIMEZONE:${PLAN.tz}`]
    .concat(events)
    .concat(['END:VCALENDAR'])
    .join('\n');
  upsertFile('deadlines.ics', ics, MimeType.PLAIN_TEXT, folder);
}

function scaffold() {
  const folder = getOutputFolder();
  const today = formatDate(new Date());
  const cause = PLAN.court.cause_no;
  const targets = [
    `${cause}_Opposition_to_MSJ.docx`,
    `${cause}_Motion_to_Strike_Affidavit.docx`,
    `Chase_4155_Settlement_Offer_${pctLabel(PLAN.offers.chase_4155[0])}pct_${today}.docx`,
    `Bureaus_§611_MOV_JPMCB4978_${today}.docx`,
    `Chase_§623_DirectDispute_DuplicateTradeline_${today}.docx`,
    `SCS_ComEd_Validation_PFD_${pctLabel(PLAN.offers.scs_comed[0])}pct_${today}.docx`,
    `CreditCollections_Progressive_Validation_PFD_${pctLabel(PLAN.offers.credit_collections_prog[0])}pct_${today}.docx`,
    `Amex_15443_Settlement_Offer_${pctLabel(PLAN.offers.amex[0])}pct_${today}.docx`,
    `BofA_Settlement_Offer_${pctLabel(PLAN.offers.bofa[0])}pct_${today}.docx`,
  ];
  targets.forEach((name) => {
    ensurePlaceholderDoc(name, folder);
  });
}

function ensurePlaceholderDoc(name, folder) {
  const existing = folder.getFilesByName(name);
  if (existing.hasNext()) {
    return;
  }
  const doc = DocumentApp.create(name);
  const file = DriveApp.getFileById(doc.getId());
  folder.addFile(file);
  try {
    DriveApp.getRootFolder().removeFile(file);
  } catch (err) {
    // Ignore if removal is not permitted.
  }
}

function buildIcsEvent(uid, dateStr, summary, stamp) {
  const start = parsePlanDate(dateStr);
  const end = new Date(start.getTime() + 60 * 60 * 1000);
  const startLocal = Utilities.formatDate(start, PLAN.tz, "yyyyMMdd'T'HHmmss");
  const endLocal = Utilities.formatDate(end, PLAN.tz, "yyyyMMdd'T'HHmmss");
  return [
    'BEGIN:VEVENT',
    `UID:${uid}@debt-agent`,
    `DTSTAMP:${stamp}`,
    `DTSTART;TZID=${PLAN.tz}:${startLocal}`,
    `DTEND;TZID=${PLAN.tz}:${endLocal}`,
    `SUMMARY:${summary}`,
    'END:VEVENT',
  ].join('\n');
}

function writeDocToFolder(name, content, folder) {
  const existing = folder.getFilesByName(name);
  let doc;
  let file;
  if (existing.hasNext()) {
    file = existing.next();
    doc = DocumentApp.openById(file.getId());
  } else {
    doc = DocumentApp.create(name);
    file = DriveApp.getFileById(doc.getId());
    folder.addFile(file);
    try {
      DriveApp.getRootFolder().removeFile(file);
    } catch (err) {
      // Ignore if already removed or permission denied.
    }
  }
  const body = doc.getBody();
  body.clear();
  const parts = content.split('\n');
  body.appendParagraph(parts[0] || '');
  parts.slice(1).forEach((line) => body.appendParagraph(line));
  doc.saveAndClose();
}

function upsertFile(name, content, mimeType, folder) {
  const files = folder.getFilesByName(name);
  if (files.hasNext()) {
    const file = files.next();
    file.setContent(content);
    while (files.hasNext()) {
      folder.removeFile(files.next());
    }
    return file;
  }
  return folder.createFile(name, content, mimeType);
}

function getOutputFolder() {
  const folders = DriveApp.getFoldersByName(OUTPUT_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(OUTPUT_FOLDER_NAME);
}

function buildContext() {
  const today = formatDate(new Date());
  return Object.assign({}, PLAN, {
    now: today,
    today: today,
  });
}

function renderTemplate(name, ctx) {
  const template = TEMPLATES[name];
  if (!template) {
    throw new Error(`Missing template: ${name}`);
  }
  return template.replace(/{{\s*([^}]+)\s*}}/g, (match, expr) => {
    return resolveExpression(expr.trim(), ctx);
  });
}

function resolveExpression(expr, ctx) {
  const [path, ...filters] = expr.split('|').map((part) => part.trim());
  let value = resolvePath(path, ctx);
  filters.forEach((filter) => {
    value = applyFilter(value, filter);
  });
  return value == null ? '' : value;
}

function resolvePath(path, ctx) {
  if (path === 'now') {
    return formatDate(new Date());
  }
  return path.split('.').reduce((acc, key) => {
    if (acc == null) {
      return acc;
    }
    return acc[key];
  }, ctx);
}

function applyFilter(value, filter) {
  if (value == null) {
    return value;
  }
  switch (filter) {
    case 'upper':
      return String(value).toUpperCase();
    case 'money':
      return formatMoney(value);
    default:
      return value;
  }
}

function formatMoney(amount) {
  return Utilities.formatString('%,.2f', Number(amount));
}

function computeAmount(balance, pct) {
  return Math.round(balance * (pct / 100) * 100) / 100;
}

function pctLabel(pct) {
  return String(pct).replace('.', '_');
}

function parsePlanDate(dateStr) {
  const [year, month, day] = dateStr.split('-').map((part) => parseInt(part, 10));
  const date = new Date(year, month - 1, day, 0, 0, 0, 0);
  return date;
}

function formatDate(date) {
  return Utilities.formatDate(date, PLAN.tz, 'yyyy-MM-dd');
}
