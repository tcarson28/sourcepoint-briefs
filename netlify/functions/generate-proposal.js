const Anthropic = require("@anthropic-ai/sdk");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  TabStopType, TabStopPosition, PageNumber, Header, Footer
} = require("docx");

const DEEP_BLUE = "0056A8";
const DARK_GRAY = "333333";
const MID_GRAY = "888888";
const WHITE = "FFFFFF";
const CONTENT_WIDTH = 9360;

const PROPOSAL_SYS = `You are SourcePoint's senior sourcing operator. Generate a complete client-ready sourcing proposal.

SourcePoint: U.S.-founded, China-based sourcing consultancy. Founder: Taylor Carson. Website: sourcepointco.com. Email: hello@sourcepointco.com.

Standard Sourcing Package includes: Supplier search, vetting of 3 suppliers including 1 site visit within Guangdong Province, price benchmarking and negotiation in Mandarin, bilingual spec sheet and product rendering, sample coordination for 2 suppliers, formal supplier recommendation report.

CURRENT FEE SCHEDULE:
Founder standard package: 1200. Returning client package: 950. Corporate package: 3500. Rush supplement: 25 percent added to base package fee for engagements under 21 days.
Additional category pricing per category beyond the first: Founder standard 500, moderate 700, complex 900. Returning standard 400, moderate 550, complex 700. Corporate standard 1000, moderate 1500, complex 2000.
Coordination fee: 10 percent of confirmed order value plus shipping. Minimum 500 total. Split evenly across Stage 3 and Stage 4. Product cost and freight are pass-through payments and do not appear in sourcing fee or coordination fee calculations.

PAYMENT STRUCTURE (4 stages):
Stage 1: 50 percent of total sourcing fee due at proposal acceptance.
Stage 2: Remaining 50 percent of sourcing fee due at sample approval and sourcing phase close.
Stage 3: 50 percent of coordination fee (5 percent of confirmed order value plus shipping) due at production coordination start.
Stage 4: Remaining 50 percent of coordination fee (5 percent of confirmed order value plus shipping) due prior to full order shipping.

Generate the proposal as clean plain text with numbered section headers in ALL CAPS. No markdown symbols. No dashes of any kind anywhere including em dashes, en dashes, or hyphens used as punctuation. Use commas or rewrite instead.

Sections: 1. INTRODUCTION 2. GOALS AND PRIORITIES 3. PROPOSED SCOPE OF WORK TIMELINE AND DELIVERABLES 4. PRODUCT SPECIFICATIONS 5. COST STRUCTURE 6. PAYMENT TERMS 7. SCOPE NOTES 8. WHY SOURCEPOINT 9. ACCEPTANCE AND NEXT STEPS

Rules: Operator tone. Calm and direct. No hype. Fill every field with real data from the brief. Payment Terms section must describe all four stages clearly. Cost Structure must show coordination fee as 10 percent of order value plus shipping with minimum 500, collected in two equal installments at Stages 3 and 4. End with: SourcePoint | U.S.-Founded. China-Based. Real Factory Access.`;

function makeHeaderCell(text, width) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: { top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
    shading: { fill: DEEP_BLUE, type: ShadingType.CLEAR },
    margins: { top: 100, bottom: 100, left: 120, right: 120 },
    children: [new Paragraph({ children: [new TextRun({ text, font: "Open Sans", size: 18, bold: true, color: WHITE })] })]
  });
}

function makeCell(text, width, shade) {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    borders: { top: { style: BorderStyle.SINGLE, size: 1, color: "E2E8F0" }, bottom: { style: BorderStyle.SINGLE, size: 1, color: "E2E8F0" }, left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE } },
    shading: { fill: shade ? "F0F5FC" : WHITE, type: ShadingType.CLEAR },
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.TOP,
    children: [new Paragraph({ children: [new TextRun({ text: text || "", font: "Open Sans", size: 19, color: DARK_GRAY })] })]
  });
}

function sectionHeading(text) {
  return new Paragraph({
    spacing: { before: 360, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: DEEP_BLUE, space: 4 } },
    children: [new TextRun({ text, font: "Cinzel", size: 22, bold: true, color: DEEP_BLUE })]
  });
}

function bodyPara(text) {
  return new Paragraph({
    spacing: { before: 60, after: 80 },
    children: [new TextRun({ text: text || "", font: "Open Sans", size: 20, color: DARK_GRAY })]
  });
}

function spacer() {
  return new Paragraph({ spacing: { before: 120, after: 0 }, children: [new TextRun("")] });
}

function parseProposalSections(text) {
  const sections = [];
  const lines = text.split("\n");
  let current = null;

  for (const line of lines) {
    const trimmed = line.trim();
    const match = trimmed.match(/^(\d+)\.\s+([A-Z][A-Z\s&]+)$/);
    if (match) {
      if (current) sections.push(current);
      current = { heading: trimmed, lines: [] };
    } else if (current && trimmed) {
      current.lines.push(trimmed);
    }
  }
  if (current) sections.push(current);
  return sections;
}

exports.handler = async (event) => {
  if (event.httpMethod !== "POST") {
    return { statusCode: 405, body: "Method not allowed" };
  }

  let body;
  try {
    body = JSON.parse(event.body);
  } catch {
    return { statusCode: 400, body: JSON.stringify({ error: "Invalid JSON" }) };
  }

  const { brief, pricing } = body;
  if (!brief) {
    return { statusCode: 400, body: JSON.stringify({ error: "Missing brief data" }) };
  }

  const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });

  const relMap = { cold: "Cold / New", referral: "Referral", returning: "Returning Client" };
  const pkgPrice = brief.tier === "corporate" ? 3500 : brief.relationshipType === "returning" ? 950 : 1200;

  const catTable = (brief.productCategories || [])
    .flatMap((c) =>
      (c.skus || []).map((s, si) =>
        (si === 0 ? c.category : "") + " | " + (s.sku || "") + " | " + (s.quantity || "TBD") + " | " + (s.phase || "Phase 1") + " | " + (s.notes || "")
      )
    )
    .join("\n");

  const briefData = [
    "CLIENT NAME: " + (brief.clientName || "Not confirmed"),
    "COMPANY: " + (brief.company || "N/A"),
    "PRODUCT: " + (brief.product || "Not confirmed"),
    "QUANTITY: " + (brief.quantity || "Not confirmed"),
    "SPECIFICATIONS: " + (brief.specifications || "Not confirmed"),
    "BUDGET PER UNIT: " + (brief.budgetPerUnit || "Not confirmed"),
    "DESTINATION: " + (brief.destination || "Not confirmed"),
    "TIMELINE: " + (brief.timeline || "Not confirmed"),
    "CLIENT RELATIONSHIP: " + (relMap[brief.relationshipType] || "New"),
    "TIER: " + (brief.tier === "corporate" ? "Corporate" : "Founder"),
    "PACKAGE FEE: $" + pkgPrice,
    "TOTAL FEE: $" + (pricing ? pricing.total : pkgPrice),
    "STAGE 1 DEPOSIT: $" + (pricing ? pricing.stage1 : Math.round(pkgPrice / 2)),
    "STAGE 2 BALANCE: $" + (pricing ? pricing.stage2 : pkgPrice - Math.round(pkgPrice / 2)),
    "STATED GOALS: " + ((Array.isArray(brief.statedGoals) ? brief.statedGoals.map((g) => g.text).join("; ") : brief.statedGoals) || ""),
    "PRODUCT CATEGORIES TABLE:\nCategory | SKU | Quantity | Phase | Notes\n" + (catTable || "Not specified"),
    "CURRENT SITUATION: " + (brief.currentSituation || ""),
    "ANALYST NOTES: " + (brief.analystNotes || ""),
  ].join("\n");

  try {
    const response = await client.messages.create({
      model: "claude-sonnet-4-20250514",
      max_tokens: 3000,
      system: PROPOSAL_SYS,
      messages: [{ role: "user", content: "Generate a complete sourcing proposal using this confirmed brief data:\n\n" + briefData }],
    });

    const proposalText = response.content[0].text;
    const sections = parseProposalSections(proposalText);

    const children = [
      // Cover block
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [new TextRun({ text: "SOURCEPOINT", font: "Cinzel", size: 44, bold: true, color: DEEP_BLUE })] }),
      new Paragraph({ spacing: { before: 0, after: 60 }, children: [new TextRun({ text: "Sourcing Proposal", font: "Cinzel", size: 30, color: DARK_GRAY })] }),
      new Paragraph({
        spacing: { before: 0, after: 40 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: DEEP_BLUE, space: 6 } },
        children: [new TextRun({ text: (brief.clientName || "[Client]") + (brief.company ? " | " + brief.company : ""), font: "Open Sans", size: 22, color: MID_GRAY, italics: true })]
      }),
      new Paragraph({ spacing: { before: 120, after: 0 }, children: [new TextRun({ text: new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" }), font: "Open Sans", size: 18, color: MID_GRAY })] }),
      spacer(),
      spacer(),
    ];

    // Metadata row
    children.push(new Table({
      width: { size: CONTENT_WIDTH, type: WidthType.DXA },
      columnWidths: [3120, 3120, 3120],
      rows: [
        new TableRow({ children: [makeHeaderCell("Proposal For", 3120), makeHeaderCell("Prepared By", 3120), makeHeaderCell("Date", 3120)] }),
        new TableRow({ children: [makeCell(brief.clientName + (brief.company ? "\n" + brief.company : ""), 3120, false), makeCell("Taylor Carson\nSourcepoint", 3120, true), makeCell(new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" }), 3120, false)] }),
      ]
    }));

    children.push(spacer());
    children.push(spacer());

    // Add each parsed section
    for (const section of sections) {
      children.push(sectionHeading(section.heading));
      for (const line of section.lines) {
        children.push(bodyPara(line));
      }
    }

    // If parsing failed, fall back to plain text
    if (sections.length === 0) {
      for (const line of proposalText.split("\n")) {
        const t = line.trim();
        if (!t) { children.push(spacer()); continue; }
        const isHeading = /^\d+\.\s+[A-Z]/.test(t);
        if (isHeading) {
          children.push(sectionHeading(t));
        } else {
          children.push(bodyPara(t));
        }
      }
    }

    const doc = new Document({
      styles: { default: { document: { run: { font: "Open Sans", size: 20, color: DARK_GRAY } } } },
      sections: [{
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
          }
        },
        headers: {
          default: new Header({
            children: [new Paragraph({
              tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
              border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: DEEP_BLUE, space: 4 } },
              spacing: { after: 0 },
              children: [
                new TextRun({ text: "SOURCEPOINT", font: "Cinzel", size: 18, bold: true, color: DEEP_BLUE }),
                new TextRun({ text: "\tSourcepoint Sourcing Proposal | " + (brief.clientName || ""), font: "Open Sans", size: 16, color: MID_GRAY }),
              ]
            })]
          })
        },
        footers: {
          default: new Footer({
            children: [new Paragraph({
              tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
              border: { top: { style: BorderStyle.SINGLE, size: 4, color: DEEP_BLUE, space: 4 } },
              spacing: { before: 0 },
              children: [
                new TextRun({ text: "sourcepointco.com | hello@sourcepointco.com", font: "Open Sans", size: 16, color: MID_GRAY }),
                new TextRun({ text: "\tPage ", font: "Open Sans", size: 16, color: MID_GRAY }),
                new TextRun({ children: [PageNumber.CURRENT], font: "Open Sans", size: 16, color: MID_GRAY }),
              ]
            })]
          })
        },
        children,
      }]
    });

    const buffer = await Packer.toBuffer(doc);
    return {
      statusCode: 200,
      headers: {
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": `attachment; filename="SP_Proposal_${(brief.clientName || "Client").replace(/[^a-zA-Z0-9]/g, "_")}_${new Date().toISOString().split("T")[0]}.docx"`,
      },
      body: buffer.toString("base64"),
      isBase64Encoded: true,
    };
  } catch (err) {
    console.error("Proposal generation error:", err);
    return { statusCode: 500, body: JSON.stringify({ message: err.message || "Generation failed" }) };
  }
};
