// build_pptx.js — Data Cleaning & Visualisation Presentation
// Run: node build_pptx.js

const pptxgen = require("pptxgenjs");
const path = require("path");

const CHARTS = path.join(__dirname, "charts");
const OUT    = path.join(__dirname, "Data_Cleaning_Visualization.pptx");

// ── Palette ────────────────────────────────────────────────────────────────
const C = {
  navy   : "1F4E79",
  navyDk : "16375A",
  blue   : "2E75B6",
  blueLt : "DEEAF1",
  teal   : "2A9D8F",
  white  : "FFFFFF",
  off    : "F7FAFC",
  grey   : "64748B",
  greyLt : "E2E8F0",
  dark   : "1E293B",
  green  : "375623",
  greenBg: "E2EFDA",
  amber  : "C55A11",
  red    : "C00000",
};

const W = 10, H = 5.625;  // slide dimensions

// ── Helpers ────────────────────────────────────────────────────────────────
function header(slide, title) {
  // Navy header bar
  slide.addShape("rect", { x: 0, y: 0, w: W, h: 0.75, fill: { color: C.navy }, line: { color: C.navy } });
  // White title text
  slide.addText(title, {
    x: 0.35, y: 0, w: W - 0.7, h: 0.75,
    fontSize: 22, bold: true, color: C.white,
    fontFace: "Calibri", align: "left", valign: "middle", margin: 0
  });
}

function footerBar(slide, text) {
  slide.addShape("rect", { x: 0, y: H - 0.32, w: W, h: 0.32, fill: { color: C.navy }, line: { color: C.navy } });
  slide.addText(text, {
    x: 0.3, y: H - 0.32, w: W - 0.6, h: 0.32,
    fontSize: 9, color: C.blueLt, fontFace: "Calibri", align: "left", valign: "middle", margin: 0
  });
}

function sectionLabel(slide, text, x, y, w, bgColor) {
  const bg = bgColor || C.blue;
  slide.addShape("rect", { x, y, w, h: 0.3, fill: { color: bg }, line: { color: bg } });
  slide.addText(text, {
    x, y, w, h: 0.3,
    fontSize: 9, bold: true, color: C.white, fontFace: "Calibri",
    align: "center", valign: "middle", margin: 0
  });
}

function makeCard(slide, x, y, w, h) {
  slide.addShape("rect", {
    x, y, w, h,
    fill: { color: C.white },
    line: { color: C.greyLt, width: 1 },
    shadow: { type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.08 }
  });
}

// ── Presentation ───────────────────────────────────────────────────────────
(async () => {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author  = "Pathuma Hiruni";
  pres.title   = "Data Cleaning & Visualisation — MSE-803";

  // ════════════════════════════════════════════════════════════════════════
  // SLIDE 1 — Title
  // ════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };

    // Decorative accent bar at top
    s.addShape("rect", { x: 0, y: 0, w: W, h: 0.12, fill: { color: C.blue }, line: { color: C.blue } });

    // Centre card
    makeCard(s, 1.0, 1.0, 8.0, 3.5);
    // Left accent
    s.addShape("rect", { x: 1.0, y: 1.0, w: 0.12, h: 3.5, fill: { color: C.teal }, line: { color: C.teal } });

    s.addText("Messy Data: Cleaning,\nCorrelation & Visualisation", {
      x: 1.3, y: 1.2, w: 7.2, h: 1.6,
      fontSize: 32, bold: true, color: C.navy, fontFace: "Calibri",
      align: "left", valign: "top"
    });

    s.addText([
      { text: "MSE-803 Data Analytics  |  Week 3 — Activity 2", options: { breakLine: true } },
    ], {
      x: 1.3, y: 2.85, w: 7.2, h: 0.35,
      fontSize: 13, color: C.blue, fontFace: "Calibri", align: "left", italic: true
    });

    s.addText([
      { text: "Pathuma Hiruni", options: { breakLine: true } },
      { text: "1 May 2026" }
    ], {
      x: 1.3, y: 3.25, w: 7.2, h: 0.6,
      fontSize: 12, color: C.grey, fontFace: "Calibri", align: "left"
    });

    // Bottom bar
    s.addShape("rect", { x: 0, y: H - 0.12, w: W, h: 0.12, fill: { color: C.teal }, line: { color: C.teal } });

    // Course tag bottom-right
    s.addText("MSE-803 Data Analytics  |  Yoobee College", {
      x: 0, y: H - 0.45, w: W - 0.3, h: 0.32,
      fontSize: 9, color: C.blueLt, fontFace: "Calibri", align: "right", italic: true
    });
  }

  // ════════════════════════════════════════════════════════════════════════
  // SLIDE 2 — About This Task
  // ════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.off };
    header(s, "About This Task");

    const cols = [
      { icon: "📁", label: "Dataset", val: "messy_dataset_Mukesh.csv\n10 rows, multiple quality issues" },
      { icon: "🎯", label: "Goal",    val: "Clean messy data & find if\nAge is connected to Salary" },
      { icon: "🛠", label: "Tools",   val: "Python · Pandas · NumPy\nSeaborn · Matplotlib · OpenPyXL" },
      { icon: "📊", label: "Output",  val: "Cleaned data + 6 charts\n5-sheet Excel workbook" },
    ];

    cols.forEach((c, i) => {
      const cx = 0.3 + i * 2.4;
      makeCard(s, cx, 0.9, 2.2, 2.6);
      s.addShape("rect", { x: cx, y: 0.9, w: 2.2, h: 0.38, fill: { color: C.navy }, line: { color: C.navy } });
      s.addText(c.label, { x: cx, y: 0.9, w: 2.2, h: 0.38, fontSize: 10, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
      s.addText(c.icon, { x: cx + 0.7, y: 1.38, w: 0.8, h: 0.6, fontSize: 22, align: "center", valign: "middle" });
      s.addText(c.val, { x: cx + 0.1, y: 2.0, w: 2.0, h: 1.3, fontSize: 10, color: C.dark, fontFace: "Calibri", align: "center", valign: "top" });
    });

    // Key question
    makeCard(s, 0.3, 3.65, 9.4, 1.6);
    s.addShape("rect", { x: 0.3, y: 3.65, w: 9.4, h: 0.35, fill: { color: C.teal }, line: { color: C.teal } });
    s.addText("Key Question", { x: 0.3, y: 3.65, w: 9.4, h: 0.35, fontSize: 10, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
    s.addText('"Is Age connected to Salary? If older, do employees earn more?"', {
      x: 0.5, y: 4.1, w: 9.0, h: 0.95,
      fontSize: 14, color: C.navy, fontFace: "Calibri", italic: true, align: "center", valign: "middle"
    });

    footerBar(s, "MSE-803 Data Analytics  |  Week 3 Activity 2");
  }

  // ════════════════════════════════════════════════════════════════════════
  // SLIDE 3 — Problems Found
  // ════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.off };
    header(s, "Problems Found in Raw Data");

    const hdr = [
      { text: "#",         options: { bold: true, fill: { color: C.navy }, color: C.white, fontSize: 9, align: "center" } },
      { text: "Problem",   options: { bold: true, fill: { color: C.navy }, color: C.white, fontSize: 9, align: "center" } },
      { text: "Example",   options: { bold: true, fill: { color: C.navy }, color: C.white, fontSize: 9, align: "center" } },
      { text: "How Fixed", options: { bold: true, fill: { color: C.navy }, color: C.white, fontSize: 9, align: "center" } },
    ];

    const rows = [
      ["1",  "Duplicate row",           "Bob appears 2 times",         "Second Bob removed"],
      ["2",  "Missing Age (×2)",         "Empty cells",                 "Filled with median: 29"],
      ["3",  "Missing Salary (×2)",      "Empty cells",                 "Filled with median: $62,000"],
      ["4",  "Age written as word",      '"thirty-eight"',              "Converted to 38"],
      ["5",  "Salary written as word",   '"sixty five thousand"',       "Converted to 65,000"],
      ["6",  "Broken date",              "2019-13-01 (month = 13!)",    "Marked as Unknown"],
      ["7",  "Missing Join Date",        "Empty cell",                  "Marked as Unknown"],
      ["8",  "Missing Name",             "Empty cell",                  "Filled as Unknown"],
      ["9",  "Missing ID",               "Eve has no ID",               "Assigned ID: 6"],
      ["10", "Missing Country",          "Empty cell",                  "Marked as Unknown"],
    ];

    const tableData = [hdr, ...rows.map((r, i) => {
      const bg = i % 2 === 0 ? C.blueLt : C.white;
      return r.map((cell, ci) => ({
        text: cell,
        options: { fill: { color: bg }, color: C.dark, fontSize: 8.5,
                   align: ci === 0 ? "center" : "left" }
      }));
    })];

    s.addTable(tableData, {
      x: 0.25, y: 0.85, w: 9.5, h: 4.35,
      colW: [0.35, 2.2, 2.6, 2.8],
      border: { pt: 0.5, color: C.greyLt }
    });

    // Footer note
    s.addShape("rect", { x: 0, y: H - 0.32, w: W, h: 0.32, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText("Before: 10 rows with many problems    →    After: 9 clean rows ✓", {
      x: 0.3, y: H - 0.32, w: W - 0.6, h: 0.32,
      fontSize: 9, color: C.blueLt, fontFace: "Calibri", align: "center", valign: "middle", margin: 0
    });
  }

  // ════════════════════════════════════════════════════════════════════════
  // SLIDE 4 — Code Structure
  // ════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.off };
    header(s, "Code Structure");

    const sections = [
      { num: "1", label: "Raw Data",      desc: "Read CSV file from disk" },
      { num: "2", label: "Issues Log",    desc: "Document every problem found" },
      { num: "3", label: "Data Cleaning", desc: "Fix all messy problems step by step" },
      { num: "4", label: "Statistics",    desc: "Pearson correlation + IQR outlier detection" },
      { num: "5", label: "Charts",        desc: "Build and save 6 PNG charts" },
      { num: "6", label: "Excel Workbook",desc: "Write everything into 5 formatted sheets" },
    ];

    const bw = 1.52, bh = 1.0, startX = 0.28, startY = 0.9;
    sections.forEach((sec, i) => {
      const bx = startX + i * (bw + 0.06);
      makeCard(s, bx, startY, bw, bh);
      s.addShape("rect", { x: bx, y: startY, w: bw, h: 0.3, fill: { color: C.navy }, line: { color: C.navy } });
      s.addText(`Section ${sec.num}`, { x: bx, y: startY, w: bw, h: 0.3, fontSize: 8, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
      s.addText(sec.label, { x: bx + 0.05, y: startY + 0.32, w: bw - 0.1, h: 0.3, fontSize: 10, bold: true, color: C.blue, fontFace: "Calibri", align: "center", valign: "middle" });
      s.addText(sec.desc, { x: bx + 0.05, y: startY + 0.62, w: bw - 0.1, h: 0.35, fontSize: 8.5, color: C.dark, fontFace: "Calibri", align: "center", valign: "top" });
    });

    // Arrow connectors (simple right-arrows)
    for (let i = 0; i < 5; i++) {
      const ax = startX + (i + 1) * (bw + 0.06) - 0.06;
    }

    // Helper functions box
    makeCard(s, 0.28, 2.1, 9.44, 1.35);
    s.addShape("rect", { x: 0.28, y: 2.1, w: 9.44, h: 0.3, fill: { color: C.teal }, line: { color: C.teal } });
    s.addText("Helper Functions", { x: 0.28, y: 2.1, w: 9.44, h: 0.3, fontSize: 9, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
    s.addText("apply_header()     style_cell()     set_col_widths()     iqr_outliers()", {
      x: 0.5, y: 2.45, w: 9.0, h: 0.85,
      fontSize: 11, color: C.navy, fontFace: "Consolas", align: "center", valign: "middle", bold: true
    });

    // Pipeline label
    makeCard(s, 0.28, 3.6, 9.44, 1.7);
    s.addShape("rect", { x: 0.28, y: 3.6, w: 9.44, h: 0.3, fill: { color: C.blue }, line: { color: C.blue } });
    s.addText("Pipeline Flow", { x: 0.28, y: 3.6, w: 9.44, h: 0.3, fontSize: 9, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
    s.addText("CSV Input  →  Clean Data  →  Pearson Correlation  →  IQR Outliers  →  6 Charts  →  5-Sheet Excel Workbook", {
      x: 0.4, y: 3.95, w: 9.2, h: 1.2,
      fontSize: 10.5, color: C.navy, fontFace: "Calibri", align: "center", valign: "middle"
    });

    footerBar(s, "MSE-803 Data Analytics  |  Week 3 Activity 2");
  }

  // ════════════════════════════════════════════════════════════════════════
  // SLIDE 5 — Data Cleaning Steps
  // ════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.off };
    header(s, "How the Data Was Cleaned");

    const steps = [
      {
        title: "Step 1 — Fix Text-as-Number",
        code:  "df['Age']    = df['Age'].replace('thirty-eight', '38')\ndf['Salary'] = df['Salary'].replace('sixty five thousand', '65000')"
      },
      {
        title: "Step 2 — Remove Duplicate (Bob)",
        code:  "df.loc[1, 'Age'] = 30\ndf = df.drop(index=2).reset_index(drop=True)"
      },
      {
        title: "Step 3 — Fix Missing / Broken Values",
        code:  "df['Name']      → filled with 'Unknown'\ndf['Join Date'] 2019-13-01 → 'Unknown'\ndf['Country']   empty cells → 'Unknown'"
      },
      {
        title: "Step 4 — Fill Missing Numbers with Median",
        code:  "Age median    = 29 years\nSalary median = $62,000"
      },
    ];

    const sw = 4.5, sh = 1.9;
    const positions = [
      { x: 0.25, y: 0.85 }, { x: 5.02, y: 0.85 },
      { x: 0.25, y: 2.9  }, { x: 5.02, y: 2.9  },
    ];

    steps.forEach((step, i) => {
      const { x, y } = positions[i];
      makeCard(s, x, y, sw, sh);
      s.addShape("rect", { x, y, w: sw, h: 0.3, fill: { color: C.navy }, line: { color: C.navy } });
      s.addText(step.title, { x, y, w: sw, h: 0.3, fontSize: 9, bold: true, color: C.white, fontFace: "Calibri", align: "left", valign: "middle", margin: [0, 0, 0, 8] });

      // Code block background
      s.addShape("rect", { x: x + 0.08, y: y + 0.35, w: sw - 0.16, h: sh - 0.45, fill: { color: "1E293B" }, line: { color: "1E293B" } });
      s.addText(step.code, {
        x: x + 0.12, y: y + 0.38, w: sw - 0.24, h: sh - 0.52,
        fontSize: 9, color: "A8D8EA", fontFace: "Consolas", align: "left", valign: "top"
      });
    });

    footerBar(s, "MSE-803 Data Analytics  |  Week 3 Activity 2");
  }

  // ════════════════════════════════════════════════════════════════════════
  // SLIDE 6 — Pearson Correlation
  // ════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.off };
    header(s, "Pearson Correlation — Age vs Salary");

    // Left: heatmap image
    makeCard(s, 0.25, 0.85, 4.2, 4.4);
    s.addImage({ path: path.join(CHARTS, "heatmap.png"), x: 0.3, y: 0.9, w: 4.1, h: 4.2 });

    // Right: interpretation
    makeCard(s, 4.65, 0.85, 5.1, 4.4);

    // Big r value
    s.addText("r = 0.629", { x: 4.85, y: 0.98, w: 4.7, h: 0.7, fontSize: 30, bold: true, color: C.teal, fontFace: "Calibri", align: "center" });
    s.addText("Moderate Positive Correlation", { x: 4.85, y: 1.65, w: 4.7, h: 0.35, fontSize: 12, color: C.blue, fontFace: "Calibri", align: "center", italic: true });

    // Scale bar
    const scale = [
      { label: "+1.0   Perfect positive",  bg: "1A5276" },
      { label: "+0.7 to +1.0   Strong positive",      bg: "1F618D" },
      { label: "+0.4 to +0.7   Moderate ← Our result (r = 0.629)", bg: "2471A3", bold: true },
      { label: "0.0    No connection",     bg: "7F8C8D" },
      { label: "−1.0   Perfect negative",  bg: "922B21" },
    ];
    scale.forEach((row, i) => {
      s.addShape("rect", { x: 4.75, y: 2.1 + i * 0.38, w: 4.9, h: 0.34, fill: { color: row.bg }, line: { color: row.bg } });
      s.addText(row.label, {
        x: 4.85, y: 2.1 + i * 0.38, w: 4.7, h: 0.34,
        fontSize: row.bold ? 9 : 8.5, bold: row.bold || false,
        color: C.white, fontFace: "Calibri", align: "left", valign: "middle"
      });
    });

    s.addText("Older employees tend to earn higher salaries", {
      x: 4.75, y: 4.05, w: 4.9, h: 0.65,
      fontSize: 11, color: C.navy, fontFace: "Calibri", align: "center", valign: "middle", italic: true, bold: true
    });

    footerBar(s, "MSE-803 Data Analytics  |  Week 3 Activity 2");
  }

  // ════════════════════════════════════════════════════════════════════════
  // SLIDE 7 — Scatter Plot
  // ════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.off };
    header(s, "Age vs Salary — Scatter & Regression");

    // Main image — large
    makeCard(s, 0.25, 0.85, 6.5, 4.4);
    s.addImage({ path: path.join(CHARTS, "scatter_age_salary.png"), x: 0.3, y: 0.9, w: 6.4, h: 4.25 });

    // Right panel — key points
    makeCard(s, 7.0, 0.85, 2.75, 4.4);
    s.addShape("rect", { x: 7.0, y: 0.85, w: 2.75, h: 0.34, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText("Key Points", { x: 7.0, y: 0.85, w: 2.75, h: 0.34, fontSize: 10, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });

    const points = [
      { icon: "●", color: "2A9D8F", text: "Each dot = one employee" },
      { icon: "╱", color: "E76F51", text: "Red line = regression trend" },
      { icon: "↗", color: C.blue,   text: "Line goes bottom-left to top-right" },
      { icon: "r", color: C.navy,   text: "r = 0.64 shown in legend" },
      { icon: "✓", color: C.green,  text: "Confirms: older → higher salary" },
    ];

    points.forEach((p, i) => {
      s.addShape("rect", { x: 7.08, y: 1.3 + i * 0.62, w: 0.3, h: 0.3, fill: { color: p.color }, line: { color: p.color } });
      s.addText(p.icon, { x: 7.08, y: 1.3 + i * 0.62, w: 0.3, h: 0.3, fontSize: 10, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0, bold: true });
      s.addText(p.text, { x: 7.42, y: 1.3 + i * 0.62, w: 2.25, h: 0.3, fontSize: 9, color: C.dark, fontFace: "Calibri", align: "left", valign: "middle" });
    });

    footerBar(s, "MSE-803 Data Analytics  |  Week 3 Activity 2");
  }

  // ════════════════════════════════════════════════════════════════════════
  // SLIDE 8 — Outlier Detection
  // ════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.off };
    header(s, "Outlier Detection — IQR Method");

    // Chart image
    makeCard(s, 0.25, 0.85, 5.8, 4.4);
    s.addImage({ path: path.join(CHARTS, "boxplots.png"), x: 0.3, y: 0.9, w: 5.7, h: 4.25 });

    // Right panel
    makeCard(s, 6.3, 0.85, 3.45, 4.4);
    s.addShape("rect", { x: 6.3, y: 0.85, w: 3.45, h: 0.34, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText("IQR Method", { x: 6.3, y: 0.85, w: 3.45, h: 0.34, fontSize: 10, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
    s.addText("IQR = Q3 − Q1\nFence = Q1 − 1.5×IQR  to  Q3 + 1.5×IQR\nValues outside fence = Outliers", {
      x: 6.4, y: 1.25, w: 3.25, h: 0.9,
      fontSize: 8.5, color: C.dark, fontFace: "Calibri", align: "left", valign: "top", italic: true
    });

    // IQR table
    const iqrHdr = [
      { text: "Column",       options: { bold: true, fill: { color: C.navy }, color: C.white, fontSize: 9, align: "center" } },
      { text: "Lower Fence",  options: { bold: true, fill: { color: C.navy }, color: C.white, fontSize: 9, align: "center" } },
      { text: "Upper Fence",  options: { bold: true, fill: { color: C.navy }, color: C.white, fontSize: 9, align: "center" } },
      { text: "Outliers",     options: { bold: true, fill: { color: C.navy }, color: C.white, fontSize: 9, align: "center" } },
    ];
    const iqrRows = [
      [
        { text: "Age",    options: { fill: { color: C.blueLt }, color: C.dark, fontSize: 9, align: "center" } },
        { text: "15.0 yrs", options: { fill: { color: C.blueLt }, color: C.dark, fontSize: 9, align: "center" } },
        { text: "47.0 yrs", options: { fill: { color: C.blueLt }, color: C.dark, fontSize: 9, align: "center" } },
        { text: "None ✓",   options: { fill: { color: C.greenBg }, color: C.green, fontSize: 9, align: "center", bold: true } },
      ],
      [
        { text: "Salary",    options: { fill: { color: C.white }, color: C.dark, fontSize: 9, align: "center" } },
        { text: "$50,000",   options: { fill: { color: C.white }, color: C.dark, fontSize: 9, align: "center" } },
        { text: "$74,000",   options: { fill: { color: C.white }, color: C.dark, fontSize: 9, align: "center" } },
        { text: "None ✓",    options: { fill: { color: C.greenBg }, color: C.green, fontSize: 9, align: "center", bold: true } },
      ],
    ];
    s.addTable([iqrHdr, ...iqrRows], {
      x: 6.3, y: 2.25, w: 3.45, h: 0.85,
      border: { pt: 0.5, color: C.greyLt }
    });

    // Result callout
    s.addShape("rect", { x: 6.3, y: 3.2, w: 3.45, h: 1.0, fill: { color: C.greenBg }, line: { color: "375623", width: 1 } });
    s.addText("✓ ZERO outliers found\nin Age or Salary", {
      x: 6.35, y: 3.25, w: 3.35, h: 0.9,
      fontSize: 16, bold: true, color: C.green, fontFace: "Calibri", align: "center", valign: "middle"
    });

    s.addText("All values fall within\nnormal IQR fences", {
      x: 6.35, y: 4.25, w: 3.35, h: 0.85,
      fontSize: 9, color: C.grey, fontFace: "Calibri", align: "center", valign: "top", italic: true
    });

    footerBar(s, "MSE-803 Data Analytics  |  Week 3 Activity 2");
  }

  // ════════════════════════════════════════════════════════════════════════
  // SLIDE 9 — Salary by Country
  // ════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.off };
    header(s, "Average Salary by Country");

    makeCard(s, 0.25, 0.85, 6.5, 4.4);
    s.addImage({ path: path.join(CHARTS, "salary_by_country.png"), x: 0.3, y: 0.9, w: 6.4, h: 4.25 });

    // Findings
    makeCard(s, 7.0, 0.85, 2.75, 4.4);
    s.addShape("rect", { x: 7.0, y: 0.85, w: 2.75, h: 0.34, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText("Findings", { x: 7.0, y: 0.85, w: 2.75, h: 0.34, fontSize: 10, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });

    const findings = [
      { rank: "🥇", country: "AUS",     note: "Highest average salary" },
      { rank: "🥈", country: "NZ",      note: "Middle range salary" },
      { rank: "🥉", country: "AU",      note: "Lowest among known" },
      { rank: "✕",  country: "Unknown", note: "Excluded from chart" },
    ];

    findings.forEach((f, i) => {
      const fy = 1.3 + i * 0.7;
      makeCard(s, 7.08, fy, 2.6, 0.58);
      s.addText(f.rank, { x: 7.08, y: fy + 0.04, w: 0.42, h: 0.5, fontSize: 16, align: "center", valign: "middle" });
      s.addText(f.country, { x: 7.52, y: fy + 0.04, w: 1.0, h: 0.25, fontSize: 11, bold: true, color: C.navy, fontFace: "Calibri", align: "left", valign: "top" });
      s.addText(f.note,    { x: 7.52, y: fy + 0.29, w: 2.0, h: 0.25, fontSize: 8.5, color: C.grey, fontFace: "Calibri", align: "left", valign: "top" });
    });

    footerBar(s, "MSE-803 Data Analytics  |  Week 3 Activity 2");
  }

  // ════════════════════════════════════════════════════════════════════════
  // SLIDE 10 — Missing Value Map
  // ════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.off };
    header(s, "Missing Value Map — Before vs After Cleaning");

    makeCard(s, 0.25, 0.85, 9.5, 3.8);
    s.addImage({ path: path.join(CHARTS, "missing_map.png"), x: 0.3, y: 0.9, w: 9.4, h: 3.65 });

    // Legend bar
    makeCard(s, 0.25, 4.75, 9.5, 0.6);
    const legend = [
      { color: "FF6B6B", label: "Red = Missing value" },
      { color: "DEEAF1", label: "Light blue = Data present" },
      { color: "2A9D8F", label: "After cleaning: almost no red cells" },
      { color: C.amber,  label: "One Unknown Join Date remains (unfixable)" },
    ];
    legend.forEach((l, i) => {
      const lx = 0.4 + i * 2.35;
      s.addShape("rect", { x: lx, y: 4.82, w: 0.18, h: 0.18, fill: { color: l.color }, line: { color: l.color } });
      s.addText(l.label, { x: lx + 0.22, y: 4.8, w: 2.1, h: 0.22, fontSize: 8, color: C.dark, fontFace: "Calibri", align: "left", valign: "top" });
    });

    footerBar(s, "MSE-803 Data Analytics  |  Week 3 Activity 2");
  }

  // ════════════════════════════════════════════════════════════════════════
  // SLIDE 11 — Excel Workbook
  // ════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.off };
    header(s, "Excel Workbook — 5 Professional Sheets");

    const sheets = [
      { num: "1", name: "Raw Data",              desc: "Original CSV with red cells highlighting problems",  color: "C00000" },
      { num: "2", name: "Issues Log",            desc: "Every problem documented with colour coding",        color: C.amber  },
      { num: "3", name: "Cleaned Data",          desc: "Clean 9-row dataset + summary statistics",          color: C.teal   },
      { num: "4", name: "Correlation & Outliers",desc: "Pearson matrix + IQR fence table",                  color: C.blue   },
      { num: "5", name: "Visualizations",        desc: "All 6 charts embedded side by side",                color: C.navy   },
    ];

    const bw = 1.82;
    sheets.forEach((sh, i) => {
      const bx = 0.25 + i * (bw + 0.07);
      makeCard(s, bx, 0.85, bw, 2.3);
      s.addShape("rect", { x: bx, y: 0.85, w: bw, h: 0.36, fill: { color: sh.color }, line: { color: sh.color } });
      s.addText(`Sheet ${sh.num}`, { x: bx, y: 0.85, w: bw, h: 0.36, fontSize: 10, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
      s.addText(sh.name, { x: bx + 0.05, y: 1.26, w: bw - 0.1, h: 0.42, fontSize: 10, bold: true, color: C.dark, fontFace: "Calibri", align: "center", valign: "middle" });
      s.addText(sh.desc, { x: bx + 0.05, y: 1.7, w: bw - 0.1, h: 1.15, fontSize: 8.5, color: C.grey, fontFace: "Calibri", align: "center", valign: "top" });
    });

    // Colour legend
    makeCard(s, 0.25, 3.3, 9.5, 1.95);
    s.addShape("rect", { x: 0.25, y: 3.3, w: 9.5, h: 0.3, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText("Cell Colour Legend", { x: 0.25, y: 3.3, w: 9.5, h: 0.3, fontSize: 9, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });

    const legend = [
      { color: "FFFF00", text: "Yellow\nMissing value" },
      { color: "FFB3C1", text: "Pink\nDuplicate row" },
      { color: "FF9999", text: "Red\nNon-numeric text" },
      { color: "FFD580", text: "Orange\nInvalid date" },
      { color: C.greenBg, text: "Green\nClean data" },
    ];
    legend.forEach((l, i) => {
      const lx = 0.6 + i * 1.85;
      s.addShape("rect", { x: lx, y: 3.75, w: 1.2, h: 0.45, fill: { color: l.color }, line: { color: "AAAAAA", width: 0.5 } });
      s.addText(l.text, { x: lx, y: 4.25, w: 1.2, h: 0.75, fontSize: 8.5, color: C.dark, fontFace: "Calibri", align: "center", valign: "top" });
    });

    footerBar(s, "MSE-803 Data Analytics  |  Week 3 Activity 2");
  }

  // ════════════════════════════════════════════════════════════════════════
  // SLIDE 12 — Key Findings
  // ════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.off };
    header(s, "Key Findings");

    const findings = [
      { icon: "✓", color: C.teal,  text: "Data was messy but fixable — 10 problems found and all fixed with Python" },
      { icon: "✓", color: C.blue,  text: "Age and Salary have MODERATE POSITIVE correlation: r = 0.629" },
      { icon: "✓", color: C.teal,  text: "Older employees tend to earn more salary" },
      { icon: "✓", color: C.blue,  text: "ZERO outliers found — IQR method confirms all values are normal" },
      { icon: "✓", color: C.teal,  text: "Australia (AUS) pays the highest average salary" },
      { icon: "✓", color: C.blue,  text: "Professional Excel output: 5 sheets, colour coding, all charts embedded" },
      { icon: "⚠", color: C.amber, text: "Warning: Only 9 rows of data — larger dataset needed for stronger conclusions" },
    ];

    findings.forEach((f, i) => {
      const fy = 0.9 + i * 0.62;
      s.addShape("rect", { x: 0.3, y: fy + 0.06, w: 0.35, h: 0.35, fill: { color: f.color }, line: { color: f.color } });
      s.addText(f.icon, { x: 0.3, y: fy + 0.06, w: 0.35, h: 0.35, fontSize: 11, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });
      makeCard(s, 0.72, fy, 9.0, 0.48);
      s.addText(f.text, { x: 0.85, y: fy, w: 8.75, h: 0.48, fontSize: 11, color: C.dark, fontFace: "Calibri", align: "left", valign: "middle" });
    });

    footerBar(s, "MSE-803 Data Analytics  |  Week 3 Activity 2");
  }

  // ════════════════════════════════════════════════════════════════════════
  // SLIDE 13 — Conclusion
  // ════════════════════════════════════════════════════════════════════════
  {
    const s = pres.addSlide();
    s.background = { color: C.navy };

    s.addShape("rect", { x: 0, y: 0, w: W, h: 0.12, fill: { color: C.teal }, line: { color: C.teal } });

    // White content card
    makeCard(s, 0.5, 0.5, 9.0, 3.9);
    s.addShape("rect", { x: 0.5, y: 0.5, w: 9.0, h: 0.38, fill: { color: C.navy }, line: { color: C.navy } });
    s.addText("Conclusion", { x: 0.5, y: 0.5, w: 9.0, h: 0.38, fontSize: 14, bold: true, color: C.white, fontFace: "Calibri", align: "center", valign: "middle", margin: 0 });

    const points = [
      "10 data quality problems were identified and all fixed using Python",
      "Pearson r = 0.629 shows moderate positive relationship between Age and Salary",
      "IQR outlier detection confirms no extreme values in the dataset",
      "6 professional charts were created and embedded in a formatted Excel workbook",
      "The code is clean and modular — each section has one clear responsibility",
      "Lesson: Clean data + correct method + adequate sample size = reliable analysis",
    ];

    s.addText(points.map((p, i) => ({
      text: p, options: { bullet: true, breakLine: i < points.length - 1 }
    })), {
      x: 0.7, y: 1.0, w: 8.5, h: 3.2,
      fontSize: 11, color: C.dark, fontFace: "Calibri", align: "left", valign: "top",
      paraSpaceAfter: 4
    });

    // Thank you / questions
    s.addText("Questions?", {
      x: 1.5, y: 4.45, w: 7.0, h: 0.5,
      fontSize: 28, bold: true, color: C.white, fontFace: "Calibri", align: "center"
    });
    s.addText("Pathuma Hiruni  |  MSE-803 Data Analytics  |  Yoobee College  |  1 May 2026", {
      x: 0.3, y: 5.05, w: 9.4, h: 0.3,
      fontSize: 9, color: C.blueLt, fontFace: "Calibri", align: "center", italic: true
    });

    s.addShape("rect", { x: 0, y: H - 0.12, w: W, h: 0.12, fill: { color: C.teal }, line: { color: C.teal } });
  }

  // ── Write file ──────────────────────────────────────────────────────────
  await pres.writeFile({ fileName: OUT });
  console.log("✓ Saved:", OUT);
})();
