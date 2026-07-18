const {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, ShadingType, BorderStyle,
  LevelFormat, convertInchesToTwip
} = require("docx");

const NAVY = "1F3864";
const ACCENT = "C00000";
const GREY = "595959";

function title(text) {
  return new Paragraph({
    text,
    heading: HeadingLevel.TITLE,
    alignment: AlignmentType.CENTER,
    spacing: { after: 100 },
  });
}

function subtitle(text) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { after: 300 },
    children: [new TextRun({ text, color: GREY, size: 22, italics: true })],
  });
}

function layerHeading(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 400, after: 150 },
    border: {
      bottom: { color: NAVY, space: 4, style: BorderStyle.SINGLE, size: 12 },
    },
    children: [new TextRun({ text, color: NAVY, bold: true })],
  });
}

function sectionHeading(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 100 },
    children: [new TextRun({ text, color: ACCENT, bold: true })],
  });
}

function bullet(text, opts = {}) {
  return new Paragraph({
    numbering: { reference: "digest-bullets", level: 0 },
    spacing: { after: 80 },
    children: [new TextRun({ text, bold: !!opts.bold })],
  });
}

function sourceLine(text) {
  return new Paragraph({
    spacing: { after: 150 },
    children: [new TextRun({ text, italics: true, size: 18, color: GREY })],
  });
}

function para(text) {
  return new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text })] });
}

const doc = new Document({
  numbering: {
    config: [
      {
        reference: "digest-bullets",
        levels: [
          { level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: convertInchesToTwip(0.35), hanging: convertInchesToTwip(0.2) } } } },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 }, // A4
          margin: { top: 900, bottom: 900, left: 1000, right: 1000 },
        },
      },
      children: [
        title("NBR Group – Daily Morning Real Estate Digest"),
        subtitle("Compiled for Senior Management  |  Thursday, 18 July 2026  |  For internal circulation"),

        para(
          "This digest tracks residential-led news and updates across the real estate sector, with additional coverage of commercial, industrial and leasing activity, structured across three layers: (1) Bengaluru & Karnataka, (2) South India, and (3) National. RERA and CREDAI updates are reported within each respective layer."
        ),

        // LAYER 1
        layerHeading("Layer 1 — Bengaluru & Karnataka"),

        sectionHeading("Residential"),
        bullet("Bengaluru recorded 12,544 residential unit launches in Q2 2026 (+4% YoY); H1 2026 launches crossed 25,000 units, marking a record first half for the city."),
        bullet("Submarket share of new launches: South Bengaluru led with 38% (Electronic City, Bannerghatta Road, Kanakapura Road corridors), North 28%, East 26% (led by Whitefield)."),
        bullet("Average residential prices rose from ~₹7,120/sq ft in 2023 to ~₹9,260/sq ft in 2025; price growth has cooled to ~6% in 2025 (from a sharp 17% in 2024), signalling a return to fundamentals with more disciplined pricing and better negotiating room for buyers."),
        bullet("Bengaluru sold 128 ultra-luxury homes (₹10 crore+) in FY26 — smaller than Hyderabad (625 units) and larger than Chennai (58 units) in South India's luxury segment, with ₹10 crore fetching ~3,900 sq ft in Bengaluru against 6,200 sq ft in Hyderabad."),
        sourceLine("Sources: Cushman & Wakefield Bengaluru MarketBeat; Modern Spaaces market analysis; IBEF/luxury housing report."),

        sectionHeading("Commercial, Industrial & Leasing"),
        bullet("Bengaluru held a 23% share of India's Q1 2026 office gross leasing (21.9 million sq ft across top 8 cities, +13% YoY) — second only to Mumbai (30%)."),
        bullet("Demand continues to be led by IT-BPM occupiers and expanding Global Capability Centres (GCCs), alongside sustained flex-space operator uptake."),
        sourceLine("Source: Cushman & Wakefield India Office Market Report, Q1 2026."),

        sectionHeading("Karnataka RERA & CREDAI Updates"),
        bullet("K-RERA ruling: handing over a flat's keys does not constitute complete possession — a Bengaluru developer has been directed to finish all promised amenities and transfer khata, corpus fund and maintenance accounts to the owners' association within 90 days; the case involved a 2017 possession where water, power-meter and DG back-up commitments remained unfulfilled for over six years."),
        bullet("Earlier K-RERA order: a Bengaluru developer was directed to hand over a delayed villa and pay ₹70 lakh interest to the affected buyers (April 2026)."),
        bullet("CREDAI has sought a 3–6 month extension on RERA project-completion timelines nationally, citing construction and supply-chain challenges — directly relevant to Karnataka/CREDAI Bengaluru members managing delayed projects."),
        sourceLine("Sources: Deccan Herald; Business Standard; Outlook Money; CREDAI Bengaluru/Karnataka."),

        // LAYER 2
        layerHeading("Layer 2 — South India"),

        sectionHeading("Residential"),
        bullet("Chennai recorded ~4,350 residential launches in Q2 2026 (+17% QoQ), with Suburban West leading at a 44% share; prices are projected to grow 7–9% in 2025-26, with the city continuing to be positioned as a value-driven, lower-entry-price alternative to Bengaluru and Hyderabad."),
        bullet("Hyderabad remains South India's ultra-luxury leader — 625 homes priced above ₹10 crore were sold in FY26 worth ₹8,562 crore, versus Bengaluru (128 units) and Chennai (58 units); a ₹10 crore ticket buys ~6,200 sq ft in Hyderabad against 3,900 sq ft in Bengaluru and 4,300 sq ft in Chennai."),
        bullet("Collectively, Bengaluru, Chennai and Hyderabad sold 811 ultra-luxury units worth ₹11,246 crore in FY26, underscoring South India's growing share of India's high-end residential demand."),
        sourceLine("Sources: Cushman & Wakefield Chennai MarketBeat; The Hans India; IBEF; Tradebrains."),

        sectionHeading("Commercial, Industrial & Leasing"),
        bullet("Chennai's office market posted gross leasing of ~1.66 million sq ft in Q1 2026, with Global Capability Centres accounting for 55% of quarterly leasing — reinforcing the city's rise as a GCC hub alongside its traditional manufacturing and export-led economy."),
        bullet("Hyderabad held a 14% share of India's Q1 2026 office leasing (third-largest after Mumbai and Bengaluru), continuing to draw biotech and technology occupiers."),
        sourceLine("Sources: Cushman & Wakefield Chennai MarketBeat; The Hans India."),

        sectionHeading("South India RERA & CREDAI Updates"),
        bullet("Telangana RERA project registrations rose from 3,630 (November 2021) to 8,227 (November 2023), reflecting steady formalisation of the state's real estate market."),
        bullet("CREDAI Hyderabad Property Show 2026, held at HITEX in February 2026, drew a strong visitor and exhibitor response."),
        bullet("Tamil Nadu RERA issued show-cause notices to five real-estate companies for norm violations, directing responses within 15 days."),
        bullet("CREDAI's national ask for a 3–6 month RERA completion-timeline extension also covers Telangana and Tamil Nadu developer members facing supply-chain-linked delays."),
        sourceLine("Sources: Telangana Today; TN-RERA; RERA.news; CREDAI Hyderabad."),

        // LAYER 3
        layerHeading("Layer 3 — National"),

        sectionHeading("Residential"),
        bullet("Housing price growth across India's top 8 cities cooled to ~6% in 2025 (from a sharp 17% in 2024); affordability improved in H1 2026 on the back of RBI's cumulative 125 bps repo rate cuts, though MMR and NCR remain above the affordability threshold."),
        bullet("Developer momentum stayed strong even as prices cooled: Puravankara reported a 28% rise in Q1 FY27 sales bookings to ₹1,439 crore; Oberoi Realty's 'Three Sixty North' project in NCR logged gross bookings of ~₹8,109 crore on 13.52 lakh sq ft of RERA carpet area sold."),
        bullet("Magicbricks continues to expand its 'MagicHomes' new-launch discovery platform — now covering Delhi-NCR, MMR, Bengaluru, Hyderabad, Kolkata, Pune, Ahmedabad and Chennai — with project data cross-verified against RERA filings."),
        sourceLine("Sources: Business Standard; company results/exchange filings; Magicbricks."),

        sectionHeading("Commercial, Industrial & Leasing"),
        bullet("Office: Q1 2026 gross leasing reached 21.9 million sq ft across the top 8 cities (+13% YoY); Mumbai led with 30% share, followed by Bengaluru (23%) and Hyderabad (14%); IT-BPM (23%) and BFSI (21%) were the top demand drivers, with flex operators (18%) and engineering & manufacturing (15%) close behind."),
        bullet("Industrial & Warehousing: H1 2026 net absorption reached 34.8 million sq ft (+2.4% YoY), with gross leasing across the top 8 cities up 12% YoY to ~22 million sq ft; manufacturing (30%) and 3PL operators (23%) led demand, followed by FMCG/FMCD (18%) and e-commerce (10%)."),
        bullet("Delhi-NCR remained India's leading logistics micro-market with ~20% of national absorption, ahead of the Pune (17%) and Mumbai (16%) industrial corridors; Grade-A asset absorption rose to 59% in H1 2026 (from 55% in H1 2025) on ESG and global-compliance-led flight to quality."),
        sourceLine("Sources: Cushman & Wakefield India Office Market Report & Outlook 2026; BizzBuzz; India Tribune; Newkerala."),

        sectionHeading("National CREDAI & RERA Updates"),
        bullet("CREDAI's Budget 2026 recommendations to the Centre: raise the affordable-housing price cap from ₹45 lakh to ₹90 lakh, and restore tax incentives for developers building affordable housing."),
        bullet("CREDAI has formally sought a 3–6 month extension to RERA project-completion timelines nationwide, citing ongoing construction-cost and supply-chain pressures on builders."),
        bullet("CREDAI's 'Real Estate @2047 — Fostering Equity and Fuelling Economic Growth' roadmap projects the sector reaching USD 10 trillion by 2047 (~one-fifth of India's GDP), citing urbanisation, infrastructure build-out, affordable housing demand and digital/GenAI adoption as the key growth drivers; CREDAI currently estimates sector-wide growth at a 10–12% CAGR."),
        sourceLine("Sources: CREDAI press releases & knowledge-center reports; SKJ Landbase; Business News This Week."),

        new Paragraph({
          spacing: { before: 400 },
          border: { top: { color: "BFBFBF", space: 4, style: BorderStyle.SINGLE, size: 6 } },
          children: [new TextRun({
            text: "Methodology note: This digest is compiled through web-based monitoring of real estate trade portals (ET Realty, Moneycontrol Real Estate/Property, Housing.com, Magicbricks, RealtyNXT, Realty Plus) and other business media, RERA authority orders/notices, and CREDAI communications. Some publisher sites restrict direct automated access; where that occurs, coverage is corroborated through secondary reporting and cited accordingly. Figures are the latest publicly available as of the digest date and should be independently verified before use in decision-making.",
            italics: true, size: 16, color: GREY,
          })],
        }),
      ],
    },
  ],
});

Packer.toBuffer(doc).then((buf) => {
  require("fs").writeFileSync(__dirname + "/NBR_Daily_Morning_Digest_2026-07-18.docx", buf);
  console.log("written");
});
