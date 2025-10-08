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

const DEFAULT_FONT_FAMILY = 'Calibri';
const DEFAULT_FONT_SIZE = 12;

const ALIGNMENT_MAP = (function () {
  if (typeof DocumentApp !== 'undefined' && DocumentApp.ParagraphAlignment) {
    return {
      LEFT: DocumentApp.ParagraphAlignment.LEFT,
      RIGHT: DocumentApp.ParagraphAlignment.RIGHT,
      CENTER: DocumentApp.ParagraphAlignment.CENTER,
      JUSTIFY: DocumentApp.ParagraphAlignment.JUSTIFY,
    };
  }
  return {};
})();

const TEMPLATES = {
  opposition_msj: (ctx) => ({
    heading: 'DEFENDANT’S OPPOSITION TO PLAINTIFF’S MOTION FOR SUMMARY JUDGMENT',
    headingStyle: { fontSize: 16, alignment: 'CENTER', bold: true },
    caption: {
      lines: [
        `IN THE ${ctx.court.jurisdiction.toUpperCase()} ${ctx.court.court_name.toUpperCase()}`,
        `CAUSE NO. ${ctx.court.cause_no}`,
        '',
        'JPMORGAN CHASE BANK, N.A., Plaintiff,',
        'v.',
        `${ctx.party.defendant_name.toUpperCase()}, Defendant.`,
      ],
      styles: { alignment: 'CENTER', fontSize: 11 },
    },
    sections: [
      {
        type: 'paragraph',
        text:
          'Comes now the Defendant and opposes summary judgment. A genuine dispute of material fact exists because the Plaintiff’s own designated materials conflict on the alleged “last payment,” defeating account-stated as a matter of law.',
        styles: { fontSize: 12 },
      },
      {
        type: 'paragraph',
        text:
          `Plaintiff’s memorandum asserts a payment of $${ctx.facts.chase_last_payment_memo.amount} on ${ctx.facts.chase_last_payment_memo.date}, while the affiant swears to $${ctx.facts.chase_last_payment_affidavit.amount} on ${ctx.facts.chase_last_payment_affidavit.date}. The contradiction undermines the reliability of the business records offered in support of summary judgment.`,
        styles: { fontSize: 12, bold: true },
      },
      {
        type: 'paragraph',
        text:
          'In addition, Plaintiff has not itemized the balance or produced the complete run of statements despite designating “Periodic Billing Statements (pp. 1–49).” Rule 56 relief is therefore improper until full documentation is produced.',
        styles: { fontSize: 12 },
      },
      {
        type: 'paragraph',
        text:
          'Defendant respectfully requests denial of the motion for summary judgment or, in the alternative, a deferral under Rule 56 to compel complete statements, an itemization of fees and interest, and a prompt trial setting.',
        styles: { fontSize: 12, italic: true },
      },
      {
        type: 'paragraph',
        text: `Dated: ${ctx.now}`,
        styles: { fontSize: 11, bold: true },
      },
    ],
    closing: {
      lines: ['Respectfully submitted,', ctx.party.defendant_name, ctx.party.defendant_addr],
      styles: { fontSize: 11 },
    },
  }),
  motion_strike_affidavit: (ctx) => ({
    heading: 'MOTION TO STRIKE OR DISREGARD AFFIDAVIT PORTIONS',
    headingStyle: { fontSize: 15, alignment: 'CENTER', bold: true },
    caption: {
      lines: [
        `IN THE ${ctx.court.jurisdiction.toUpperCase()} ${ctx.court.court_name.toUpperCase()}`,
        `CAUSE NO. ${ctx.court.cause_no}`,
        '',
        'JPMORGAN CHASE BANK, N.A., Plaintiff,',
        'v.',
        `${ctx.party.defendant_name.toUpperCase()}, Defendant.`,
      ],
      styles: { alignment: 'CENTER', fontSize: 11 },
    },
    sections: [
      {
        type: 'paragraph',
        text: 'Defendant moves to strike or disregard the affidavit portions relied upon by Plaintiff in support of summary judgment.',
        styles: { fontSize: 12 },
      },
      {
        type: 'bullets',
        items: [
          'Internal contradiction on the purported last-payment date and amount renders the affidavit unreliable.',
          'The affiant lacks personal knowledge and fails to lay a proper hearsay exception foundation for third-party business records.',
          'No itemization of fees or interest from charge-off to present is supplied.',
          'Plaintiff has not produced the complete monthly statement series it references (pp. 1–49).',
        ],
        styles: { fontSize: 11 },
      },
      {
        type: 'paragraph',
        text: 'Defendant requests that the contradictory paragraphs be stricken, that unsupported balance assertions be disregarded, and that production of the complete statement set be ordered forthwith.',
        styles: { fontSize: 12, italic: true },
      },
      {
        type: 'paragraph',
        text: `Dated: ${ctx.now}`,
        styles: { fontSize: 11, bold: true },
      },
    ],
    closing: {
      lines: ['Respectfully submitted,', ctx.party.defendant_name, ctx.party.defendant_addr],
      styles: { fontSize: 11 },
    },
  }),
  settlement_offer: (ctx) => ({
    heading: ctx.party.defendant_name,
    headingStyle: { fontSize: 16, alignment: 'CENTER', bold: true },
    subject: `Subject: Immediate Resolution Proposal – ${ctx.name}`,
    subjectStyle: { fontSize: 13, bold: true },
    topLines: [
      { text: ctx.date, styles: { italic: true, fontSize: 11 } },
      { text: 'Via Email & Certified Mail', styles: { bold: true, fontSize: 11 } },
      { text: ctx.to_address, styles: { fontSize: 11 } },
    ],
    sections: [
      {
        type: 'paragraph',
        text: `Re: ${ctx.name} – Current Balance $${ctx.balance}`,
        styles: { fontSize: 12, bold: true },
      },
      {
        type: 'paragraph',
        text: 'Thank you for reviewing this proposal. I am committed to closing this matter on businesslike terms that deliver immediate value.',
        styles: { fontSize: 12 },
      },
      {
        type: 'paragraph',
        text: `I am authorized to remit ${ctx.pct}% ($${ctx.amount}) as a guaranteed settlement figure, with funds available within five business days of written acceptance.`,
        styles: { fontSize: 12, bold: true },
      },
      {
        type: 'bullets',
        items: [
          'Written confirmation of dismissal with prejudice (if litigation is pending) and each side bearing its own fees and costs.',
          'No post–charge-off interest or fees assessed following receipt of settlement funds.',
          'Update to consumer reporting agencies within 30 days to reflect the account as settled in full.',
          'No resale, reassignment, or continued collection on any residual balance.',
        ],
        styles: { fontSize: 11 },
      },
      {
        type: 'paragraph',
        text: 'This proposal recognizes the financial hardship outlined in the attached declaration while securing a swift, mutually beneficial resolution.',
        styles: { fontSize: 11, italic: true },
      },
      {
        type: 'paragraph',
        text: 'Please confirm approval in writing and provide the designated payee and remittance instructions so I can finalize payment immediately.',
        styles: { fontSize: 12 },
      },
    ],
    closing: {
      lines: ['Respectfully,', ctx.party.defendant_name, ctx.party.defendant_addr],
      styles: { fontSize: 11 },
    },
  }),
  frca_611_mov: (ctx) => ({
    heading: ctx.party.defendant_name,
    headingStyle: { fontSize: 16, alignment: 'CENTER', bold: true },
    subject: 'Subject: FCRA §611 Investigation Demand – JPMCB *4978 Duplicate vs. Chase 4155',
    subjectStyle: { fontSize: 13, bold: true },
    topLines: [
      { text: ctx.date, styles: { italic: true, fontSize: 11 } },
      { text: 'Via Certified Mail', styles: { bold: true, fontSize: 11 } },
      { text: ctx.to_address, styles: { fontSize: 11 } },
    ],
    sections: [
      {
        type: 'paragraph',
        text: 'Re: Dispute of JPMCB *4978 Tradeline – Duplicate/Fragment of Chase 4155 Lawsuit Account',
        styles: { fontSize: 12, bold: true },
      },
      {
        type: 'paragraph',
        text:
          `I dispute the completeness and accuracy of the JPMCB *4978 tradeline. Your file reflects a credit limit of $${ctx.facts.jpmcb_4978_cl}, a reported high balance of $${ctx.facts.jpmcb_4978_high}, and a DOFD of ${ctx.facts.jpmcb_4978_dofd}. Those metrics are incompatible with the Chase/4155 account presently in litigation, which carries a credit limit of $${ctx.facts.chase_4155_cl}.`,
        styles: { fontSize: 12 },
      },
      {
        type: 'paragraph',
        text: 'The duplicate reporting is materially misleading and depresses my credit profile. Pursuant to FCRA §611, please conduct a reinvestigation and supply a detailed method-of-verification response.',
        styles: { fontSize: 12, italic: true },
      },
      {
        type: 'bullets',
        items: [
          'Identify the furnisher representative, contact information, and records relied upon.',
          'Clarify whether *4978 and 4155 are the same obligation and reconcile all credit limit, high balance, and DOFD fields.',
          'Delete, consolidate, or otherwise correct the tradeline so only accurate data remains.',
        ],
        styles: { fontSize: 11 },
      },
      {
        type: 'paragraph',
        text: 'Please forward the written results of the investigation within 30 days to the mailing address below.',
        styles: { fontSize: 12 },
      },
    ],
    closing: {
      lines: ['Thank you for your prompt attention.', ctx.party.defendant_name, ctx.party.defendant_addr],
      styles: { fontSize: 11 },
    },
  }),
  frca_623_direct: (ctx) => ({
    heading: ctx.party.defendant_name,
    headingStyle: { fontSize: 16, alignment: 'CENTER', bold: true },
    subject: 'Subject: Direct Furnisher Dispute under FCRA §623 – Reconcile / Delete Duplicate Tradeline',
    subjectStyle: { fontSize: 13, bold: true },
    topLines: [
      { text: ctx.date, styles: { italic: true, fontSize: 11 } },
      { text: 'Via Certified Mail', styles: { bold: true, fontSize: 11 } },
      { text: ctx.to_address, styles: { fontSize: 11 } },
    ],
    sections: [
      {
        type: 'paragraph',
        text: 'Re: Duplicate Reporting of Chase 4155 vs. JPMCB *4978',
        styles: { fontSize: 12, bold: true },
      },
      {
        type: 'paragraph',
        text:
          'I am invoking my rights under FCRA §623 to dispute the accuracy of the Chase 4155 and JPMCB *4978 tradelines. Only the correct, current obligation should be reported.',
        styles: { fontSize: 12 },
      },
      {
        type: 'bullets',
        items: [
          'Identify which tradeline (4155 or *4978) reflects the live account of record.',
          'Delete the duplicate tradeline and update all credit limit, high balance, and DOFD data points for the surviving account.',
          'Provide written confirmation of the correction, including the date the update is furnished to the nationwide consumer reporting agencies.',
        ],
        styles: { fontSize: 11 },
      },
      {
        type: 'paragraph',
        text: 'Please respond within 30 days. Continued reporting of inaccurate data will be treated as willful noncompliance.',
        styles: { fontSize: 12, italic: true },
      },
    ],
    closing: {
      lines: ['Sincerely,', ctx.party.defendant_name, ctx.party.defendant_addr],
      styles: { fontSize: 11 },
    },
  }),
  fdCPA_1692g_pfd: (ctx) => ({
    heading: ctx.party.defendant_name,
    headingStyle: { fontSize: 16, alignment: 'CENTER', bold: true },
    subject: `Subject: Validation & Resolution Request – ${ctx.account_reference || 'Outstanding Account'}`,
    subjectStyle: { fontSize: 13, bold: true },
    topLines: [
      { text: ctx.date, styles: { italic: true, fontSize: 11 } },
      { text: 'Via Certified Mail', styles: { bold: true, fontSize: 11 } },
      { text: ctx.to_address, styles: { fontSize: 11 } },
    ],
    sections: [
      {
        type: 'paragraph',
        text: `Re: ${ctx.account_reference || 'Account'} Validation and Pay-for-Delete Proposal`,
        styles: { fontSize: 12, bold: true },
      },
      {
        type: 'paragraph',
        text: 'This letter is a timely request under 15 U.S.C. §1692g. Until the account is validated, collection activity must cease.',
        styles: { fontSize: 12 },
      },
      {
        type: 'bullets',
        items: [
          'Original contract or authorization establishing the obligation.',
          'Complete itemization from charge-off to present, including interest and fees.',
          'Proof of assignment and collector licensing details for my state.',
          'Date of first delinquency furnished to the consumer reporting agencies.',
        ],
        styles: { fontSize: 11 },
      },
      {
        type: 'paragraph',
        text: `Upon validation, I am prepared to resolve the matter on a written pay-for-delete basis at ${ctx.pct}% ($${ctx.amount}), with payment remitted immediately upon agreement.`,
        styles: { fontSize: 12, bold: true },
      },
      {
        type: 'paragraph',
        text: 'All phone contact is revoked; please communicate exclusively in writing. I look forward to confirming a professional, mutually beneficial resolution.',
        styles: { fontSize: 11, italic: true },
      },
    ],
    closing: {
      lines: ['Respectfully,', ctx.party.defendant_name, ctx.party.defendant_addr],
      styles: { fontSize: 11 },
    },
  }),
  hardship_declaration: (ctx) => ({
    heading: 'DECLARATION OF FINANCIAL HARDSHIP',
    headingStyle: { fontSize: 16, alignment: 'CENTER', bold: true },
    sections: [
      {
        type: 'paragraph',
        text: `I, ${ctx.party.defendant_name}, submit this declaration to explain current financial hardship impacting my ability to pay outstanding obligations in full.`,
        styles: { fontSize: 12 },
      },
      {
        type: 'bullets',
        items: [
          'Job loss in February 2024 reduced household income significantly.',
          'Ongoing caregiving responsibilities for a parent undergoing cancer treatment.',
          'Unexpected housing remediation expenses incurred in September 2025.',
        ],
        styles: { fontSize: 11 },
      },
      {
        type: 'paragraph',
        text: 'These circumstances are offered solely to contextualize my negotiation posture and to support reasonable settlement efforts.',
        styles: { fontSize: 12, italic: true },
      },
    ],
    closing: {
      lines: ['Dated: ' + ctx.now, ctx.party.defendant_name],
      styles: { fontSize: 11 },
    },
  }),
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

  PLAN.mailing.bureaus.forEach((bureau) => {
    const letterCtx = Object.assign({}, ctx, {
      date: today,
      to_address: bureau,
    });
    const shortName = bureau.split(',')[0].replace(/\s+/g, '');
    writeDocToFolder(
      `Bureaus_§611_MOV_JPMCB4978_${today}_${shortName}.docx`,
      renderTemplate('frca_611_mov', letterCtx),
      folder
    );
  });

  writeDocToFolder(
    `Chase_§623_DirectDispute_DuplicateTradeline_${today}.docx`,
    renderTemplate(
      'frca_623_direct',
      Object.assign({}, ctx, { date: today, to_address: PLAN.mailing.chase_furnisher })
    ),
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

function writeDocToFolder(name, docSpec, folder) {
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
  body.setAttributes({
    [DocumentApp.Attribute.FONT_FAMILY]: DEFAULT_FONT_FAMILY,
    [DocumentApp.Attribute.FONT_SIZE]: DEFAULT_FONT_SIZE,
  });

  if (!docSpec || typeof docSpec !== 'object') {
    appendParagraphWithStyles(body, docSpec || '', {});
    doc.saveAndClose();
    return;
  }

  if (docSpec.heading) {
    appendParagraphWithStyles(
      body,
      docSpec.heading,
      Object.assign({ fontSize: 16, bold: true, alignment: 'CENTER', spacingAfter: 6 }, docSpec.headingStyle || {})
    );
  }

  if (docSpec.subject) {
    appendParagraphWithStyles(
      body,
      docSpec.subject,
      Object.assign({ fontSize: 13, bold: true, spacingAfter: 4 }, docSpec.subjectStyle || {})
    );
  }

  if (docSpec.topLines && docSpec.topLines.length) {
    docSpec.topLines.forEach((line) => {
      appendParagraphWithStyles(body, line.text, Object.assign({ fontSize: 11 }, line.styles || {}));
    });
    body.appendParagraph('');
  }

  if (docSpec.caption && docSpec.caption.lines) {
    docSpec.caption.lines.forEach((line) => {
      appendParagraphWithStyles(
        body,
        line,
        Object.assign({ alignment: 'CENTER', fontSize: 11 }, docSpec.caption.styles || {})
      );
    });
    body.appendParagraph('');
  }

  if (docSpec.sections && docSpec.sections.length) {
    docSpec.sections.forEach((section) => appendSection(body, section));
  }

  if (docSpec.closing && docSpec.closing.lines && docSpec.closing.lines.length) {
    body.appendParagraph('');
    docSpec.closing.lines.forEach((line) => {
      appendParagraphWithStyles(body, line, docSpec.closing.styles || { fontSize: 11 });
    });
  }

  doc.saveAndClose();
}

function appendSection(body, section) {
  const type = section.type || 'paragraph';
  if (type === 'bullets') {
    appendListItems(body, section.items || [], section.styles || {}, DocumentApp.GlyphType.BULLET);
  } else if (type === 'numbered') {
    appendListItems(body, section.items || [], section.styles || {}, DocumentApp.GlyphType.NUMBER);
  } else {
    appendParagraphWithStyles(body, section.text || '', section.styles || {});
  }
}

function appendListItems(body, items, styles, glyphType) {
  if (!items.length) {
    return;
  }
  items.forEach((item) => {
    const listItem = body.appendListItem(item);
    listItem.setGlyphType(glyphType);
    applyParagraphStyles(listItem, styles);
  });
}

function appendParagraphWithStyles(body, text, styles) {
  const paragraph = body.appendParagraph(text || '');
  applyParagraphStyles(paragraph, styles || {});
  return paragraph;
}

function applyParagraphStyles(paragraph, styles) {
  const fontFamily = styles.font || DEFAULT_FONT_FAMILY;
  const fontSize = styles.fontSize || DEFAULT_FONT_SIZE;
  paragraph.setFontFamily(fontFamily);
  paragraph.setFontSize(fontSize);
  if (styles.bold !== undefined) {
    paragraph.setBold(styles.bold);
  }
  if (styles.italic !== undefined) {
    paragraph.setItalic(styles.italic);
  }
  if (styles.underline !== undefined) {
    paragraph.setUnderline(styles.underline);
  }
  if (styles.alignment) {
    const alignmentKey = String(styles.alignment).toUpperCase();
    const alignment = ALIGNMENT_MAP[alignmentKey];
    if (alignment) {
      paragraph.setAlignment(alignment);
    }
  }
  if (styles.spacingBefore !== undefined) {
    paragraph.setSpacingBefore(styles.spacingBefore);
  }
  if (styles.spacingAfter !== undefined) {
    paragraph.setSpacingAfter(styles.spacingAfter);
  }
  if (styles.lineSpacing !== undefined) {
    paragraph.setLineSpacing(styles.lineSpacing);
  }
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
  return template(ctx);
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
