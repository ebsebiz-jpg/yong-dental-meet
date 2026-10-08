import type { Cell, Table } from "@/lib/dentweb";

/* 회의 자료를 파워포인트(.pptx)로 만든다. 표마다 슬라이드 한 장, 병원 로고를 모든 장에 넣는다. */

const NAVY = "17457A";
const LIGHT = "EAF0F7";
const GRID = "C9D3DF";
const FONT = "맑은 고딕";

export type TrendChart = { title: string; labels: string[]; values: number[]; format: string };

export type SlideItem =
  | { kind: "table"; table: Table; summary?: string | null; comment?: string; commentLabel?: string }
  | { kind: "charts"; title: string; summary?: string | null; charts: TrendChart[]; comment?: string }
  | { kind: "text"; title: string; body: string };

export type PptxInput = {
  logo: string; // data URL
  title: string;
  lines: string[];
  items: SlideItem[];
  fileName: string;
};

const fmt = (c: Cell): string =>
  c === null || c === "" ? "–" : typeof c === "number" ? c.toLocaleString("ko-KR") : c;

/** 한글은 영문보다 넓게 쳐서 열 너비 비율을 정한다. */
function widthOf(s: string): number {
  let w = 0;
  for (const ch of s) w += /[가-힣]/.test(ch) ? 1.8 : 1;
  return w;
}

function colWidths(t: Table, total: number): number[] {
  const weights = t.headers.map((h, i) => {
    const longest = Math.max(widthOf(h) * 0.7, ...t.rows.map((r) => widthOf(fmt(r[i] ?? null))));
    return Math.min(30, Math.max(6, longest));
  });
  const sum = weights.reduce((a, b) => a + b, 0);
  const min = 0.7;
  const raw = weights.map((w) => (w / sum) * total);
  const widths = raw.map((w) => Math.max(min, w));
  const scale = total / widths.reduce((a, b) => a + b, 0);
  return widths.map((w) => w * scale);
}

function bodyLines(body: string): string[] {
  return body
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean)
    .map((l) => l.replace(/^- \[ \]\s*/, "☐ ").replace(/^- /, "• "));
}

export async function savePptx(input: PptxInput): Promise<void> {
  const PptxGenJS = (await import("pptxgenjs")).default;
  const pres = new PptxGenJS();
  pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5 in
  pres.title = input.title;

  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: "FFFFFF" },
    objects: [
      { image: { x: 10.2, y: -0.12, w: 2.8, h: 1.4, data: input.logo } },
      { rect: { x: 0.5, y: 1.08, w: 12.33, h: 0.03, fill: { color: NAVY } } },
    ],
    slideNumber: { x: 12.3, y: 7.08, w: 0.6, h: 0.3, fontFace: FONT, fontSize: 10, color: "888888", align: "right" },
  });
  pres.defineSlideMaster({ title: "COVER", background: { color: "FFFFFF" } });

  // 표지
  const cover = pres.addSlide({ masterName: "COVER" });
  cover.addImage({ data: input.logo, x: 3.87, y: 0.25, w: 5.6, h: 2.8 });
  cover.addText(input.title, { x: 0.8, y: 2.9, w: 11.73, h: 0.9, align: "center", fontFace: FONT, fontSize: 34, bold: true, color: NAVY });
  cover.addShape(pres.ShapeType.rect, { x: 5.67, y: 3.95, w: 2, h: 0.04, fill: { color: NAVY }, line: { color: NAVY, width: 0 } });
  cover.addText(input.lines.filter(Boolean).join("\n"), {
    x: 1.5, y: 4.2, w: 10.33, h: 1.8, align: "center", valign: "top", fontFace: FONT, fontSize: 18, color: "444444", lineSpacingMultiple: 1.3,
  });

  for (const item of input.items) {
    const slide = pres.addSlide({ masterName: "CONTENT" });
    if (item.kind === "text") {
      slide.addText(item.title, { x: 0.5, y: 0.3, w: 9.7, h: 0.7, fontFace: FONT, fontSize: 24, bold: true, color: NAVY, valign: "middle" });
      const lines = bodyLines(item.body);
      const size = lines.length > 12 ? 14 : lines.length > 8 ? 16 : 18;
      slide.addText(
        lines.map((l) => ({ text: l, options: { breakLine: true } })),
        { x: 0.7, y: 1.4, w: 11.9, h: 5.4, fontFace: FONT, fontSize: size, color: "222222", valign: "top", paraSpaceAfter: 8 },
      );
      continue;
    }

    if (item.kind === "charts") {
      slide.addText(item.title, { x: 0.5, y: 0.3, w: 9.7, h: 0.7, fontFace: FONT, fontSize: 22, bold: true, color: NAVY, valign: "middle" });
      let cy = 1.25;
      if (item.summary) {
        slide.addText(item.summary, { x: 0.5, y: cy, w: 12.33, h: 0.5, fontFace: FONT, fontSize: 15, bold: true, color: "222222", valign: "middle" });
        cy += 0.6;
      }
      const hasC = !!item.comment?.trim();
      const bottomC = hasC ? 6.2 : 6.95;
      const gap = 0.3;
      const w = (12.33 - gap * (item.charts.length - 1)) / item.charts.length;
      item.charts.forEach((c, i) => {
        slide.addChart(
          pres.ChartType.bar,
          [{ name: c.title, labels: c.labels, values: c.values }],
          {
            x: 0.5 + i * (w + gap), y: cy, w, h: bottomC - cy,
            barDir: "col", chartColors: [NAVY], barGapWidthPct: 70,
            showTitle: true, title: c.title, titleFontFace: FONT, titleFontSize: 14, titleColor: "222222",
            showLegend: false, showValue: true, dataLabelFormatCode: c.format, dataLabelFontFace: FONT,
            dataLabelFontSize: 12, dataLabelColor: "222222", dataLabelPosition: "outEnd",
            catAxisLabelFontFace: FONT, catAxisLabelFontSize: 12, catAxisLabelColor: "444444",
            valAxisHidden: true, valAxisMinVal: 0, valGridLine: { style: "none" },
          },
        );
      });
      if (hasC) {
        slide.addText(
          [
            { text: "분석 코멘트  ", options: { bold: true, color: NAVY } },
            { text: item.comment!.trim().replace(/\n/g, "  "), options: { color: "222222" } },
          ],
          { x: 0.5, y: 6.3, w: 12.33, h: 0.75, fontFace: FONT, fontSize: 12, fill: { color: LIGHT }, valign: "middle", margin: 8 },
        );
      }
      continue;
    }

    const { table: t, comment } = item;
    slide.addText(t.title, { x: 0.5, y: 0.3, w: 9.7, h: 0.7, fontFace: FONT, fontSize: 22, bold: true, color: NAVY, valign: "middle" });
    let y = 1.25;
    if (item.summary) {
      slide.addText(item.summary, { x: 0.5, y, w: 12.33, h: 0.5, fontFace: FONT, fontSize: 15, bold: true, color: "222222", valign: "middle" });
      y += 0.6;
    }
    if (t.note) {
      slide.addText(t.note, { x: 0.5, y, w: 12.33, h: 0.45, fontFace: FONT, fontSize: 10, color: "666666", valign: "top" });
      y += 0.5;
    }
    const hasComment = !!comment?.trim();
    const bottom = hasComment ? 6.2 : 6.9;
    const nCols = t.headers.length;
    const nRows = t.rows.length + 1;
    const fontSize = Math.max(8, (nCols >= 10 ? 9 : nCols >= 7 ? 10 : 12) - (nRows >= 14 ? 1 : 0));
    const rowH = Math.min(0.42, (bottom - y) / nRows);
    const diffCol = t.headers.map((h) => /증감|변화/.test(h));

    const rows = [
      t.headers.map((h, j) => ({
        text: h,
        options: { bold: true, color: "FFFFFF", fill: { color: NAVY }, align: (j === 0 ? "left" : "center") as "left" | "center" },
      })),
      ...t.rows.map((r, i) =>
        t.headers.map((_, j) => {
          const c = r[j] ?? null;
          return {
            text: fmt(c),
            options: {
              align: (j === 0 || (typeof c === "string" && c !== "") ? "left" : "right") as "left" | "right",
              fill: { color: i % 2 ? LIGHT : "FFFFFF" },
              color: diffCol[j] && typeof c === "number" && c < 0 ? "C0392B" : "222222",
              bold: j === 0 && /^(합계|총 )/.test(String(r[0] ?? "")),
            },
          };
        }),
      ),
    ];
    slide.addTable(rows, {
      x: 0.5, y, w: 12.33, colW: colWidths(t, 12.33), rowH,
      fontFace: FONT, fontSize, valign: "middle",
      border: { type: "solid", pt: 0.5, color: GRID },
      margin: [0.02, 0.08, 0.02, 0.08],
    });

    if (hasComment) {
      const text = comment!.trim();
      slide.addText(
        [
          { text: `${item.commentLabel ?? "분석 코멘트"}  `, options: { bold: true, color: NAVY } },
          { text: text.replace(/\n/g, "  "), options: { color: "222222" } },
        ],
        { x: 0.5, y: 6.3, w: 12.33, h: 0.75, fontFace: FONT, fontSize: text.length > 200 ? 10 : 12, fill: { color: LIGHT }, valign: "middle", margin: 8 },
      );
    }
  }

  await pres.writeFile({ fileName: input.fileName });
}
