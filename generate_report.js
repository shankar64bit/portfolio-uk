const {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  AlignmentType,
  BorderStyle,
  WidthType,
  ShadingType,
  LevelFormat,
  Header,
  Footer,
  PageNumber,
} = require("docx");
const fs = require("fs");

const INK = "111827";
const NAVY = "1A3557";
const TEAL = "1D7A6A";
const SLATE = "64748B";
const RULE = "CBD5E1";
const ROW_ALT = "F8FAFC";
const WHITE = "FFFFFF";
const NAVY_BG = "1A3557";
const PURPLE = "5B3FA6";
const AMBER = "B45309";

const CW = 9866;

const nb = { style: BorderStyle.NONE, size: 0, color: WHITE };
const NO = { top: nb, bottom: nb, left: nb, right: nb };
const thin = (c) => ({ style: BorderStyle.SINGLE, size: 1, color: c });
const ALL_RULE = {
  top: thin(RULE),
  bottom: thin(RULE),
  left: thin(RULE),
  right: thin(RULE),
};
const BOT_RULE = (c, sz = 4) => ({
  top: nb,
  bottom: { style: BorderStyle.SINGLE, size: sz, color: c, space: 1 },
  left: nb,
  right: nb,
});

const run = (text, opts = {}) =>
  new TextRun({ text, font: "Georgia", size: 20, color: INK, ...opts });

const p = (children, opts = {}) =>
  new Paragraph({
    spacing: { before: 0, after: 100, line: 276 },
    children: Array.isArray(children) ? children : [children],
    ...opts,
  });

const pH1 = (text) =>
  new Paragraph({
    spacing: { before: 240, after: 60 },
    border: BOT_RULE(NAVY, 6),
    children: [
      new TextRun({ text, font: "Arial", size: 28, bold: true, color: NAVY }),
    ],
  });

const pH2 = (text) =>
  new Paragraph({
    spacing: { before: 160, after: 40 },
    children: [
      new TextRun({ text, font: "Arial", size: 22, bold: true, color: NAVY }),
    ],
  });

const pSub = (text) =>
  p([
    new TextRun({
      text,
      font: "Georgia",
      size: 19,
      color: SLATE,
      italics: true,
    }),
  ]);

const pBullet = (label, detail = "") =>
  new Paragraph({
    numbering: { reference: "bullets", level: 0 },
    spacing: { before: 0, after: 80, line: 268 },
    children: detail
      ? [
          new TextRun({
            text: label,
            font: "Arial",
            size: 19,
            bold: true,
            color: INK,
          }),
          new TextRun({
            text: "  " + detail,
            font: "Georgia",
            size: 19,
            color: SLATE,
          }),
        ]
      : [new TextRun({ text: label, font: "Georgia", size: 19, color: INK })],
  });

const rule = () =>
  new Paragraph({
    spacing: { before: 180, after: 0 },
    border: BOT_RULE(RULE, 2),
    children: [run("")],
  });

const gap = (after = 80) =>
  new Paragraph({ spacing: { before: 0, after }, children: [run("")] });

function callout(title, tag, lines, tagColor = TEAL) {
  const accentW = 80;
  const contentW = CW - accentW - 200;
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: [accentW, contentW],
    margins: { top: 60, bottom: 60 },
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: NO,
            shading: { fill: tagColor, type: ShadingType.CLEAR },
            width: { size: accentW, type: WidthType.DXA },
            margins: { top: 0, bottom: 0, left: 0, right: 0 },
            children: [p([run("")])],
          }),
          new TableCell({
            borders: {
              top: nb,
              bottom: nb,
              left: nb,
              right: nb,
              insideH: nb,
              insideV: nb,
            },
            shading: { fill: ROW_ALT, type: ShadingType.CLEAR },
            width: { size: contentW, type: WidthType.DXA },
            margins: { top: 100, bottom: 100, left: 160, right: 120 },
            children: [
              new Paragraph({
                spacing: { before: 0, after: 60 },
                children: [
                  new TextRun({
                    text: title,
                    font: "Arial",
                    size: 20,
                    bold: true,
                    color: NAVY,
                  }),
                  new TextRun({ text: "  " }),
                  new TextRun({
                    text: tag,
                    font: "Arial",
                    size: 16,
                    bold: true,
                    color: WHITE,
                    shading: { fill: tagColor, type: ShadingType.CLEAR },
                  }),
                ],
              }),
              ...lines.map(
                (line) =>
                  new Paragraph({
                    spacing: { before: 0, after: 60, line: 268 },
                    children: [
                      new TextRun({
                        text: "→  " + line,
                        font: "Georgia",
                        size: 19,
                        color: SLATE,
                      }),
                    ],
                  }),
              ),
            ],
          }),
        ],
      }),
    ],
  });
}

function dataTable(headers, rows, colWidths) {
  const makeRow = (cells, isHeader = false, rowIdx = 0) =>
    new TableRow({
      children: cells.map(
        (text, ci) =>
          new TableCell({
            borders: ALL_RULE,
            shading: {
              fill: isHeader ? NAVY_BG : rowIdx % 2 === 0 ? WHITE : ROW_ALT,
              type: ShadingType.CLEAR,
            },
            margins: { top: 80, bottom: 80, left: 120, right: 80 },
            width: { size: colWidths[ci], type: WidthType.DXA },
            children: [
              new Paragraph({
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text,
                    font: isHeader ? "Arial" : "Georgia",
                    size: 18,
                    bold: isHeader,
                    color: isHeader ? WHITE : INK,
                  }),
                ],
              }),
            ],
          }),
      ),
    });
  return new Table({
    width: { size: CW, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [makeRow(headers, true), ...rows.map((r, i) => makeRow(r, false, i))],
  });
}

// ─────────────────────────────────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      {
        reference: "bullets",
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "–",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 480, hanging: 300 } } },
          },
        ],
      },
    ],
  },
  styles: {
    default: { document: { run: { font: "Georgia", size: 20, color: INK } } },
    paragraphStyles: [
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { font: "Arial", size: 28, bold: true, color: NAVY },
        paragraph: { spacing: { before: 240, after: 60 }, outlineLevel: 0 },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        quickFormat: true,
        run: { font: "Arial", size: 22, bold: true, color: NAVY },
        paragraph: { spacing: { before: 160, after: 40 }, outlineLevel: 1 },
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1020, right: 1020, bottom: 1020, left: 1020 },
        },
      },

      headers: {
        default: new Header({
          children: [
            new Table({
              width: { size: CW, type: WidthType.DXA },
              columnWidths: [6000, 3866],
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      borders: { ...NO, bottom: thin(RULE) },
                      margins: { top: 0, bottom: 60, left: 0, right: 0 },
                      width: { size: 6000, type: WidthType.DXA },
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: "Nomad Mediation",
                              font: "Arial",
                              size: 18,
                              bold: true,
                              color: NAVY,
                            }),
                            new TextRun({
                              text: "  ·  Website Review & Ideas",
                              font: "Arial",
                              size: 18,
                              color: SLATE,
                            }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      borders: { ...NO, bottom: thin(RULE) },
                      margins: { top: 0, bottom: 60, left: 0, right: 0 },
                      width: { size: 3866, type: WidthType.DXA },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.RIGHT,
                          children: [
                            new TextRun({
                              text: "Shankar Mallesh  ·  May 2026",
                              font: "Arial",
                              size: 16,
                              color: SLATE,
                            }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      },

      footers: {
        default: new Footer({
          children: [
            new Table({
              width: { size: CW, type: WidthType.DXA },
              columnWidths: [6000, 3866],
              rows: [
                new TableRow({
                  children: [
                    new TableCell({
                      borders: { ...NO, top: thin(RULE) },
                      margins: { top: 60, bottom: 0, left: 0, right: 0 },
                      width: { size: 6000, type: WidthType.DXA },
                      children: [
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: "Confidential — prepared for Amy Foster",
                              font: "Arial",
                              size: 16,
                              color: SLATE,
                            }),
                          ],
                        }),
                      ],
                    }),
                    new TableCell({
                      borders: { ...NO, top: thin(RULE) },
                      margins: { top: 60, bottom: 0, left: 0, right: 0 },
                      width: { size: 3866, type: WidthType.DXA },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.RIGHT,
                          children: [
                            new TextRun({
                              text: "Page ",
                              font: "Arial",
                              size: 16,
                              color: SLATE,
                            }),
                            new TextRun({
                              children: [PageNumber.CURRENT],
                              font: "Arial",
                              size: 16,
                              color: SLATE,
                            }),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      },

      children: [
        // ══════════════════════════════════════════════════════
        // COVER
        // ══════════════════════════════════════════════════════
        gap(200),
        new Paragraph({
          spacing: { before: 0, after: 40 },
          children: [
            new TextRun({
              text: "NOMAD MEDIATION",
              font: "Arial",
              size: 52,
              bold: true,
              color: NAVY,
            }),
          ],
        }),
        new Paragraph({
          spacing: { before: 0, after: 100 },
          border: BOT_RULE(TEAL, 8),
          children: [
            new TextRun({
              text: "Website Review & Creative Ideas Report",
              font: "Georgia",
              size: 28,
              color: TEAL,
              italics: true,
            }),
          ],
        }),
        gap(60),

        new Table({
          width: { size: CW, type: WidthType.DXA },
          columnWidths: [3288, 3289, 3289],
          rows: [
            new TableRow({
              children: [
                ["Prepared by", "Shankar Mallesh"],
                ["Date", "May 2026"],
                ["For", "Amy Foster, Nomad Mediation"],
              ].map(
                ([lbl, val]) =>
                  new TableCell({
                    borders: NO,
                    margins: { top: 0, bottom: 0, left: 0, right: 40 },
                    width: { size: 3289, type: WidthType.DXA },
                    children: [
                      new Paragraph({
                        spacing: { before: 0, after: 20 },
                        children: [
                          new TextRun({
                            text: lbl.toUpperCase(),
                            font: "Arial",
                            size: 14,
                            bold: true,
                            color: SLATE,
                            characterSpacing: 40,
                          }),
                        ],
                      }),
                      new Paragraph({
                        spacing: { before: 0, after: 0 },
                        children: [
                          new TextRun({
                            text: val,
                            font: "Arial",
                            size: 18,
                            color: INK,
                          }),
                        ],
                      }),
                    ],
                  }),
              ),
            }),
          ],
        }),
        gap(120),

        new Table({
          width: { size: CW, type: WidthType.DXA },
          columnWidths: [120, CW - 120],
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders: {
                    ...NO,
                    left: { style: BorderStyle.SINGLE, size: 10, color: TEAL },
                  },
                  width: { size: 120, type: WidthType.DXA },
                  margins: { top: 0, bottom: 0, left: 0, right: 0 },
                  children: [p([run("")])],
                }),
                new TableCell({
                  borders: NO,
                  width: { size: CW - 120, type: WidthType.DXA },
                  margins: { top: 0, bottom: 0, left: 160, right: 0 },
                  children: [
                    new Paragraph({
                      spacing: { before: 0, after: 60, line: 276 },
                      children: [
                        new TextRun({
                          text: "Nomad Mediation launched in July 2025 — still in its first year. The site is already in good shape: clean, professional, and the content is strong. This report is about taking what's working and helping more of the right people find it, trust it, and reach out. Nine sections, each with specific, actionable suggestions and forward-thinking ideas to grow the business significantly.",
                          font: "Georgia",
                          size: 20,
                          color: INK,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),

        gap(60),
        rule(),

        // ══════════════════════════════════════════════════════
        // 1 — QUICK FIXES
        // ══════════════════════════════════════════════════════
        pH1("1 — Quick Fixes"),
        pSub(
          "Small things that are easy to sort and make an immediate difference.",
        ),
        gap(60),

        callout(
          "Fix the About Page Meta Description",
          "Urgent",
          [
            "Confirmed live on the site right now: the About page meta description still reads 'Homepage For A Professional Mediation Company, Focused On Resolving Disputes And Promoting Peaceful Solutions.' — template placeholder copy that appears in Google before anyone clicks.",
            "Suggested replacement: 'Meet Amy Foster — CMC-registered workplace mediator with 10+ years HR experience and a Queen's Award background, helping teams across the UK resolve conflict confidently.' One afternoon to fix, permanent improvement.",
          ],
          AMBER,
        ),
        gap(80),

        callout("Fix the Copyright Year", "Quick Win", [
          "The footer currently reads © 2025. It is now 2026. Small detail, but visitors notice — and it signals the site hasn't been touched.",
          "In Framer: open the footer component, find the text element, update to © 2026, publish. Two minutes.",
        ]),
        gap(80),

        callout("Make the Header Sticky on Scroll", "Quick Win", [
          "The navigation bar disappears the moment a visitor scrolls down — so the phone number, 'Book a Free Consultation' button, and all nav links vanish entirely.",
          "In Framer: select the Nav component → in the right panel under 'Position', change to 'Sticky' (top: 0). Done in under a minute. This is one of the highest-value single changes on the list — it keeps the booking action available at every scroll position.",
        ]),
        gap(80),

        callout("Add 'How It Works' to the Navigation", "Quick Win", [
          "The homepage links to /how-it-works but it's absent from the nav. Someone landing on About or Pricing will never find it.",
          "This page probably does the best job of reassuring nervous first-timers — it should be reachable from anywhere on the site.",
        ]),
        gap(80),

        callout("Check Mobile on a Real Device", "Quick Win", [
          "Framer handles responsive layouts well but worth a physical check — specifically the pricing table, the three-step process, and the testimonials carousel.",
          "These are the sections most likely to clip or stack oddly on a small screen. Also check the sticky header once added — ensure it doesn't eat into content on mobile.",
        ]),

        // ══════════════════════════════════════════════════════
        // 2 — NEW PAGES
        // ══════════════════════════════════════════════════════
        pH1("2 — New Pages Worth Adding"),
        pSub(
          "The site covers workplace mediation only. These additions open new audiences and give Google more to index.",
        ),
        gap(60),

        callout(
          "Community & Housing Mediation Page",
          "High Impact",
          [
            "Nomad also offers housing and community mediation — but there's no page for it. Landlords, families, and housing officers searching for help won't find the site at all.",
            "A focused page covering who it's for, what the process looks like, and pricing would open an entirely new stream of enquiries with almost no extra work.",
          ],
          TEAL,
        ),
        gap(80),

        callout(
          "Resources / Blog Section",
          "High Impact",
          [
            "The single biggest SEO opportunity on the site. HR managers actively Google questions the FAQ already answers — the problem is those answers aren't indexed as standalone pages.",
            "Even 3–4 short articles would start bringing in organic traffic. Framer's built-in CMS makes this straightforward — add a 'Blog' collection, publish articles without touching code. Specific topics listed in Section 5.",
          ],
          TEAL,
        ),
        gap(80),

        callout("Case Studies Page", "Medium Impact", [
          "The testimonials are strong but brief. A case study gives people a story: what the situation was, what happened, what changed.",
          "Two or three anonymised case studies (labelled by sector — NHS, Charity, Engineering) would significantly increase conversion for people who are almost ready to book but not quite convinced.",
        ]),
        gap(80),

        callout("Online / Remote Mediation Page", "Medium Impact", [
          "The site makes no mention of remote mediation, yet many organisations — especially those with distributed teams or multi-site operations — would prefer or require it.",
          "A dedicated page ('Mediation via Zoom — how it works, what to expect, pricing') opens up a national client base beyond the West Midlands.",
        ]),

        // ══════════════════════════════════════════════════════
        // 3 — UX & INTERACTION
        // ══════════════════════════════════════════════════════
        pH1("3 — User Experience & Interaction"),
        pSub(
          "The tone is calm and trustworthy — exactly right for mediation. These ideas reduce friction without changing the feel.",
        ),
        gap(60),

        callout(
          "Expand Contact Options — Add WhatsApp & Live Chat",
          "High Impact",
          [
            "The current contact form asks 'How should we reach you? No preference / Email / Phone.' That's it. No WhatsApp, no chat widget, no messaging option.",
            "WhatsApp Business (free): add a click-to-chat button — many HR managers prefer messaging over calling, especially when it's a sensitive situation. A live chat widget (Tidio or Crisp — both have free tiers) gives visitors instant access without committing to a call.",
            "The contact form options should be extended to: Email / Phone / WhatsApp / Live Chat. This alone could meaningfully increase enquiry volume from people who want a lower-barrier first contact.",
          ],
          TEAL,
        ),
        gap(80),

        callout(
          "Embed a Calendly Booking Widget",
          "High Impact",
          [
            "The 'Book a Free Consultation' button goes to a custom /book-a-consultation page — great. But there's no live calendar showing Amy's actual availability.",
            "Embedding Calendly (free plan) means visitors can book a specific slot immediately, without a back-and-forth email exchange. Reduces drop-off at the highest-intent moment on the site.",
          ],
          TEAL,
        ),
        gap(80),

        callout(
          "Sticky 'Book a Free Consultation' CTA",
          "High Impact",
          [
            "Even after the header is made sticky, the primary CTA button should remain persistently visible on long-scroll pages — especially FAQ and Pricing.",
            "In Framer: a fixed-position button at the bottom-right of the screen (mobile: full-width banner at the bottom). This is industry-standard for service businesses and keeps the booking action one tap away at all times.",
          ],
          TEAL,
        ),
        gap(80),

        callout("'Is Mediation Right for Me?' Quiz", "Creative Add-on", [
          "A short 4–5 question quiz (Typeform embed or built in Framer) ending with a personalised recommendation.",
          "Sample questions: How long has the issue been ongoing? Have formal HR processes been started? How many people are involved? Is everyone willing to participate?",
          "It helps undecided visitors feel heard before they book — and makes Nomad stand out from every other mediator's static website. The quiz results page should include a direct booking link.",
        ]),
        gap(80),

        callout("Add Sector Labels to Testimonials", "Small Tweak", [
          "Testimonials are currently text-only. Adding a visible label — 'Head of People — NHS', 'CEO — Charity Sector', 'COO — Engineering' — lets people scan and self-identify instantly.",
          "In Framer's CMS: add a 'sector' field to the testimonials collection and display it as a badge. Twenty minutes of work with lasting conversion impact.",
        ]),
        gap(80),

        callout("Set Up Google Analytics 4 + Search Console", "Quick Win", [
          "There's currently no visibility into who's visiting, what they're reading, or where they drop off. Both tools are free and take roughly an hour to connect via Framer's integrations panel.",
          "Once live: track which pages drive the most enquiries, what search terms bring people in, and which sections of the page people scroll past without reading.",
        ]),

        // ══════════════════════════════════════════════════════
        // 4 — FRAMER-SPECIFIC
        // ══════════════════════════════════════════════════════
        pH1("4 — Framer-Specific Improvements"),
        pSub(
          "The site is built on Framer — these improvements are native to the platform and don't require a developer.",
        ),
        gap(60),

        dataTable(
          ["What to do", "Where in Framer", "Impact"],
          [
            [
              "Make nav sticky",
              "Select Nav → Position → Sticky, top: 0",
              "High — keeps booking CTA visible on all pages",
            ],
            [
              "Fix copyright year to 2026",
              "Footer component → text element",
              "Low — but signals a maintained, active site",
            ],
            [
              "Fix About page meta description",
              "Page settings → SEO → Meta Description",
              "High — this is the first thing Google shows",
            ],
            [
              "Add blog via Framer CMS",
              "CMS tab → New Collection → 'Blog Posts'",
              "High — enables indexed, searchable articles",
            ],
            [
              "Add sector field to testimonials",
              "CMS → Testimonials collection → Add field",
              "Medium — faster to scan, better for trust",
            ],
            [
              "Connect GA4 via Framer integrations",
              "Site Settings → Integrations → Google Analytics",
              "High — foundation for all data decisions",
            ],
            [
              "Add WhatsApp button component",
              "Insert → Link → tel: replaced with wa.me link",
              "High — lowers barrier to first contact",
            ],
            [
              "Embed Calendly on booking page",
              "Insert → Embed → paste Calendly widget URL",
              "High — removes email back-and-forth",
            ],
            [
              "Add How It Works to nav",
              "Nav component → add link to /how-it-works",
              "Medium — improves discoverability of key page",
            ],
          ],
          [3600, 3200, 3066],
        ),

        // ══════════════════════════════════════════════════════
        // 5 — SEO
        // ══════════════════════════════════════════════════════
        pH1("5 — SEO — Getting Found"),
        pSub(
          "Not about gaming algorithms. About making sure the right words are in the right places so stressed HR managers find Nomad first.",
        ),
        gap(80),

        pH2("Technical fixes"),
        gap(40),
        dataTable(
          ["What to fix", "Why it matters"],
          [
            [
              "About page meta description",
              "Confirmed showing template placeholder — first impression in Google before anyone clicks",
            ],
            [
              "Page titles on Pricing / Contact",
              "Could be more specific — 'Mediation Pricing UK | Nomad Mediation' outperforms generic site name",
            ],
            [
              "Image alt text",
              "Framer images likely have no alt text — both an SEO and accessibility gap",
            ],
            [
              "Internal linking",
              "Pages don't cross-link — blog posts should link to Pricing, Pricing should link to How It Works",
            ],
            [
              "Local keywords: Coventry / West Midlands / Nuneaton",
              "No local area mention anywhere — missing a significant slice of local search traffic",
            ],
            [
              "Schema markup (LocalBusiness)",
              "Helps Google show Nomad in local map results — can be added via Framer's custom code panel",
            ],
          ],
          [3600, 6266],
        ),
        gap(100),

        pH2("Blog articles that would actually rank"),
        gap(40),
        pBullet(
          "When Should You Use Mediation Instead of a Grievance Procedure?",
          "targets HR managers mid-process — high volume, high intent",
        ),
        pBullet(
          "What Happens on Mediation Day? A Step-by-Step Guide",
          "reduces anxiety for first-timers, highly shareable",
        ),
        pBullet(
          "How Much Does Workplace Mediation Cost in the UK?",
          "high-intent commercial search — Nomad's transparent pricing is a competitive advantage here",
        ),
        pBullet(
          "Can Mediation Work if One Person Doesn't Want to Participate?",
          "very common concern, very searchable — FAQ answer already exists, needs its own page",
        ),
        pBullet(
          "Mediation for Landlords and Tenants: What You Need to Know",
          "opens up the housing audience with almost no extra effort",
        ),
        pBullet(
          "Remote Workplace Mediation: How It Works via Zoom",
          "growing search term, no page exists for it currently",
        ),

        gap(80),
        pH2("Google Business Profile — optimise what's already there"),
        gap(40),
        p([
          new TextRun({
            text: "A Google Business Profile already exists (linked from the footer) but it's likely under-optimised. Specific improvements:",
            font: "Georgia",
            size: 19,
            color: SLATE,
          }),
        ]),
        gap(40),
        pBullet(
          "Add 'Workplace Mediator' and 'Conflict Resolution Service' as business categories",
        ),
        pBullet(
          "Upload 5–10 professional photos — office, Amy, any team training shots",
        ),
        pBullet(
          "Enable Google Messaging — lets enquiries come in directly from search results",
        ),
        pBullet(
          "Add opening hours, service areas (Coventry, West Midlands, Nuneaton, remote UK-wide)",
        ),
        pBullet(
          "Route the /feedback page to prompt a Google review — this is the single fastest way to build social proof in search",
        ),

        // ══════════════════════════════════════════════════════
        // 6 — CONTENT TWEAKS
        // ══════════════════════════════════════════════════════
        pH1("6 — Content Tweaks"),
        pSub(
          "The writing is already warm and clear. A few targeted changes would make it convert better.",
        ),
        gap(60),

        callout("Surface Amy's Full Story on the Homepage", "Medium Impact", [
          "The About page is excellent — international experience in New Zealand and the UAE, 10 years as Vice-Chair of the Positive Youth Foundation, the Queen's Award for Voluntary Service, Buckingham Palace. None of this appears on the homepage.",
          "A short 'About Amy' strip on the homepage with 3–4 of these headline credentials would immediately increase trust for cold visitors who never click through to About.",
        ]),
        gap(80),

        callout("Homepage Hero — Add a Specific Outcome", "Small Tweak", [
          "Current headline: 'Build a happier, more productive workplace.' is solid but general.",
          "Adding one concrete fact — 'Most mediations are resolved in a single day' — gives people a specific reason to believe it's worth trying. Specificity builds trust faster than warmth alone.",
        ]),
        gap(80),

        callout(
          "Mention Coventry / West Midlands Explicitly",
          "Medium Impact",
          [
            "The site reads as national but Amy is based in Nuneaton — a significant local SEO opportunity being missed.",
            "'Serving businesses across Coventry, the West Midlands, and remotely across the UK' — one line, added to the homepage and footer, would help Nomad appear for local searches like 'workplace mediator Coventry' and 'HR mediation West Midlands'.",
          ],
        ),
        gap(80),

        callout(
          "Make the CMC Accreditation and Queen's Award Unmissable",
          "Medium Impact",
          [
            "Both credentials are mentioned on the About page but easy to miss. A trust strip near the top of the homepage — 'CMC Registered · Queen's Award for Voluntary Service · 10+ Years HR Experience · International Experience (NZ & UAE)' — signals credibility immediately.",
            "For someone arriving cold, credentials are often what tips them from 'maybe' to 'let me book that call'.",
          ],
        ),
        gap(80),

        callout("Leverage the /feedback Page for Google Reviews", "Quick Win", [
          "A feedback page already exists on the site. Right now it's likely capturing responses privately. A simple addition: after submitting feedback, show a prompt — 'Would you be willing to share this on Google?' with a direct link to the Business Profile review page.",
          "Google reviews directly affect local search ranking. Even 10–15 reviews would meaningfully improve visibility.",
        ]),

        // ══════════════════════════════════════════════════════
        // 7 — AI & TECHNOLOGY INTEGRATION
        // ══════════════════════════════════════════════════════
        pH1("7 — AI & Technology Integration"),
        pSub(
          "Practical AI tools that could handle enquiry triage, after-hours support, and first contact — freeing Amy's time while improving client experience.",
        ),
        gap(60),

        callout(
          "AI Chatbot for Pre-Qualification",
          "High Impact",
          [
            "Tools like Tidio AI, Intercom, or a custom GPT-powered widget can handle first-contact questions 24/7 — 'What does mediation cost?', 'How long does it take?', 'Is it right for my situation?'",
            "The chatbot doesn't replace Amy — it qualifies the enquiry, answers FAQs instantly, and routes serious prospects to book a consultation. This means Amy only picks up calls from people who are already informed and ready.",
            "Framer supports embed codes: the chat widget drops in as a one-line embed. Most tools have a free tier sufficient for Nomad's current traffic volume.",
          ],
          PURPLE,
        ),
        gap(80),

        callout(
          "AI-Powered Intake Form Before Consultations",
          "High Impact",
          [
            "Before a free consultation call, send an automated intake form with 5–8 smart questions: number of parties, duration of the issue, sector, what's already been tried.",
            "An AI model (or Typeform's AI logic) can tailor the questions based on earlier answers. Amy arrives at every consultation already briefed — calls become more efficient and more likely to convert.",
            "Tools: Typeform (AI follow-up logic, free tier), or a custom Claude-powered form embedded on the booking confirmation page.",
          ],
          PURPLE,
        ),
        gap(80),

        callout(
          "Automated Email Sequences for Enquiries",
          "Medium Impact",
          [
            "When someone submits a consultation request, they currently likely receive a manual reply. An automated sequence could: immediately confirm receipt and set expectations (within minutes), send relevant resources 24 hours later ('What to expect on mediation day'), follow up gently if no booking is confirmed after 48 hours.",
            "Tools: Mailchimp (free up to 500 contacts), ConvertKit, or HubSpot free CRM. All integrate with Framer forms via Zapier.",
          ],
          PURPLE,
        ),
        gap(80),

        callout(
          "AI-Generated Blog Content Assistant",
          "Creative Add-on",
          [
            "Amy has the expertise; the bottleneck is time to write. Using Claude or ChatGPT with a consistent prompt template — 'Write a 600-word blog post in Amy's voice about X, targeting UK HR managers, in plain English with no jargon' — she could produce a publishable first draft in 10 minutes.",
            "The output should always be reviewed and personalised before publishing — the AI handles structure and volume, Amy handles accuracy and tone. Six posts a year, consistently published, would meaningfully improve organic search ranking within 6–12 months.",
          ],
          PURPLE,
        ),
        gap(80),

        callout(
          "WhatsApp Business API for Appointment Reminders",
          "Medium Impact",
          [
            "Once a consultation is booked via Calendly, an automated WhatsApp message can confirm the booking, send a reminder 24 hours before, and share any prep information.",
            "WhatsApp's open rate is approximately 95% vs 20% for email. For a business built on personal trust, a WhatsApp touchpoint feels appropriately human while being fully automated behind the scenes.",
          ],
          PURPLE,
        ),

        // ══════════════════════════════════════════════════════
        // 8 — BUSINESS GROWTH & NEXT LEVEL
        // ══════════════════════════════════════════════════════
        pH1("8 — Business Growth: Taking Nomad to the Next Level"),
        pSub(
          "Nomad Mediation is 10 months old. These are the moves that could significantly accelerate growth in year two and beyond.",
        ),
        gap(60),

        callout(
          "Activate the Nomad HR & Recruitment Client Base",
          "Highest Impact",
          [
            "Amy has run Nomad HR and Recruitment Ltd for 10 years, specialising in manufacturing and engineering. That existing client base — people who already know and trust Amy — is the warmest possible audience for mediation services.",
            "A simple email to existing HR/recruitment clients introducing Nomad Mediation, written personally from Amy, would likely generate the first several bookings at zero marketing cost. This should happen before anything else on this list.",
          ],
          AMBER,
        ),
        gap(80),

        callout(
          "LinkedIn B2B Content Strategy",
          "High Impact",
          [
            "HR Directors and People Managers — Nomad's primary buyers — spend significant time on LinkedIn. Amy's profile and the Nomad Mediation company page are both present but underused as growth levers.",
            "A consistent posting rhythm of 2–3 posts per week (short reflections on workplace conflict, anonymised case insights, links to blog posts) would build an audience of the exact people most likely to book. LinkedIn's organic reach for personal profiles is still remarkably strong for B2B service businesses.",
            "Key content types: 'A conflict situation I see regularly, and what usually helps'; 'Three signs a grievance process won't solve the underlying problem'; 'What good mediation actually looks like on the day'. Amy's voice and experience are the differentiator — this content can't be faked.",
          ],
          TEAL,
        ),
        gap(80),

        callout(
          "HR Consultancy Referral Partnerships",
          "High Impact",
          [
            "Many small HR consultancies encounter workplace conflicts they're not equipped to mediate. A formal referral arrangement — 'refer a client to Nomad, receive a finder's fee or reciprocal referral' — could become a consistent source of warm leads.",
            "Target: independent HR consultants in the West Midlands and Midlands region. A personal LinkedIn message from Amy to 20 relevant contacts, with a clear and simple referral proposition, is a realistic starting point.",
          ],
          TEAL,
        ),
        gap(80),

        callout(
          "Conflict Resolution Training Workshops for HR Teams",
          "Medium Impact",
          [
            "A half-day or full-day workshop for HR teams — 'How to manage conflict early, before it escalates to mediation' — creates a new revenue stream, positions Amy as an expert, and keeps Nomad top-of-mind when mediation is eventually needed.",
            "Priced at £500–£1,200 per workshop (group of up to 12), this could run alongside mediation without significant extra overhead. It also generates case study material and testimonials.",
          ],
        ),
        gap(80),

        callout(
          "Offer a Retainer Package for Corporate Clients",
          "Medium Impact",
          [
            "For larger organisations (NHS trusts, engineering firms, charities), a retainer model — 'X mediations per year, priority access, dedicated HR support line' — provides predictable recurring revenue for Nomad and cost certainty for the client.",
            "Even two or three retainer clients would significantly stabilise monthly income. The NHS Head of People testimonial suggests this relationship style is already forming naturally — a formal offering would make it easier to say yes to.",
          ],
        ),
        gap(80),

        callout("Speaking, Awards & Industry Visibility", "Longer Term", [
          "Amy's story — Queen's Award, international experience, CMC accreditation, launching a mediation firm from an HR background — is genuinely interesting. Speaking at HR conferences, CIPD events, or regional business breakfasts would build awareness and credibility simultaneously.",
          "Entering Nomad for regional SME awards (Coventry & Warwickshire Chamber, FSB) generates press coverage, backlinks, and trust signals at low cost. Award-winning organisations feature prominently on their homepage for a reason.",
        ]),

        // ══════════════════════════════════════════════════════
        // 9 — ROADMAP
        // ══════════════════════════════════════════════════════
        pH1("9 — Suggested Order of Work"),
        pSub(
          "Prioritised by effort vs. impact — quick wins and foundations first, growth levers second.",
        ),
        gap(80),

        dataTable(
          ["Week", "What", "Why first"],
          [
            [
              "1",
              "Email Nomad HR client base · Fix meta description · Make header sticky · Fix copyright year",
              "Highest ROI, near-zero cost — warm leads and immediate site fixes",
            ],
            [
              "2",
              "Set up Analytics & Search Console · Add WhatsApp Business button · Fix nav (add How It Works)",
              "Data foundation + contact friction reduced before anything else",
            ],
            [
              "3",
              "Add Community & Housing Mediation page · Add Remote Mediation page",
              "Opens two new audiences currently invisible to the site",
            ],
            [
              "4",
              "Write first 2 blog posts · Add sector labels to testimonials · Surface Amy's credentials on homepage",
              "SEO compounds over time — earlier is always better",
            ],
            [
              "5",
              "Embed Calendly on booking page · Set up AI chatbot (Tidio free) · Automated intake form",
              "Removes friction at the conversion moment",
            ],
            [
              "6–7",
              "Case Studies page · LinkedIn content plan · Optimise Google Business Profile · Feedback-to-review funnel",
              "Converts near-ready visitors, builds organic authority",
            ],
            [
              "8–12",
              "HR consultancy referral outreach · Training workshop offering · Mediation quiz · Sticky CTA · Local SEO pass · Retainer package",
              "Business development and advanced UX — builds on solid foundations",
            ],
          ],
          [720, 4720, 4426],
        ),

        // ══════════════════════════════════════════════════════
        // CLOSING
        // ══════════════════════════════════════════════════════
        rule(),
        gap(80),

        new Table({
          width: { size: CW, type: WidthType.DXA },
          columnWidths: [6800, 3066],
          rows: [
            new TableRow({
              children: [
                new TableCell({
                  borders: NO,
                  width: { size: 6800, type: WidthType.DXA },
                  margins: { top: 0, bottom: 0, left: 0, right: 200 },
                  children: [
                    new Paragraph({
                      spacing: { before: 0, after: 60, line: 276 },
                      children: [
                        new TextRun({
                          text: "Nomad Mediation is ten months old with a strong foundation, real testimonials, genuine credentials, and an existing client base from Nomad HR that most new businesses would dream of. The opportunities here are significant — and none of them require starting from scratch.",
                          font: "Georgia",
                          size: 19,
                          color: SLATE,
                          italics: true,
                        }),
                      ],
                    }),
                    gap(40),
                    new Paragraph({
                      spacing: { before: 0, after: 40 },
                      children: [
                        new TextRun({
                          text: "Happy to walk through any section in more detail, prioritise differently, or help with specific implementation.",
                          font: "Georgia",
                          size: 19,
                          color: INK,
                        }),
                      ],
                    }),
                  ],
                }),
                new TableCell({
                  borders: { ...NO, left: thin(RULE) },
                  width: { size: 3066, type: WidthType.DXA },
                  margins: { top: 0, bottom: 0, left: 200, right: 0 },
                  children: [
                    new Paragraph({
                      spacing: { before: 0, after: 40 },
                      children: [
                        new TextRun({
                          text: "Shankar Mallesh",
                          font: "Arial",
                          size: 20,
                          bold: true,
                          color: NAVY,
                        }),
                      ],
                    }),
                    new Paragraph({
                      spacing: { before: 0, after: 40 },
                      children: [
                        new TextRun({
                          text: "MSc Data Science & AI",
                          font: "Arial",
                          size: 17,
                          color: SLATE,
                        }),
                      ],
                    }),
                    new Paragraph({
                      spacing: { before: 0, after: 40 },
                      children: [
                        new TextRun({
                          text: "Aston University, Birmingham",
                          font: "Arial",
                          size: 17,
                          color: SLATE,
                        }),
                      ],
                    }),
                    new Paragraph({
                      spacing: { before: 0, after: 40 },
                      children: [
                        new TextRun({
                          text: "shankar64bitcomputing@gmail.com",
                          font: "Arial",
                          size: 17,
                          color: TEAL,
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
        gap(40),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buffer) => {
  fs.writeFileSync("Nomad_Mediation_Website_Report_v2.docx", buffer);
  console.log("Done! Report generated: Nomad_Mediation_Website_Report_v2.docx");
});
