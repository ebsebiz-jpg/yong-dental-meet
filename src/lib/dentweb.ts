import * as XLSX from "xlsx";
import JSZip from "jszip";

/*
 * 덴트웹에서 내보낸 월간 리포트 5종을 읽어 집계한다.
 * 환자 이름·연락처가 들어 있는 시트는 집계에 필요한 값(구 단위 주소, 소개자 성명)만 꺼내고
 * 나머지는 메모리에 남기지 않는다.
 */

export type Kind = "cost" | "route" | "consult" | "income" | "lab";

export const KINDS: { kind: Kind; label: string; hint: string }[] = [
  { kind: "cost", label: "기간별 진료비", hint: "진료비 통계" },
  { kind: "route", label: "내원경로 분포", hint: "내원환자 내원경로 분포" },
  { kind: "consult", label: "상담자별 상담", hint: "상담자별 상담통계" },
  { kind: "income", label: "수입 통계", hint: "수입 통계" },
  { kind: "lab", label: "기공 의뢰", hint: "기공의뢰통계" },
];

export type CostData = {
  days: number;
  total: number;
  claim: number; // 공단 청구액
  copay: number; // 본인부담금
  nonIns: number; // 비급여
  received: number; // 총수납액
  card: number;
  cash: number;
  online: number;
  cashReceipt: number;
  discount: number;
};
export type RouteRow = { name: string; patients: number; ret: number; fresh: number; visits: number; avg: number; total: number };
export type RouteData = {
  rows: RouteRow[];
  totals: { patients: number; ret: number; fresh: number; visits: number; total: number };
  regions: Map<string, number>; // 신환 거주 구
  referrers: Map<string, { name: string; count: number }>; // 소개한 환자 차트번호 -> 성명, 소개 인원
};
export type ConsultRow = {
  name: string;
  plans: number;
  confirmed: number;
  patients: number;
  confirmedPatients: number;
  amount: number;
  discount: number;
};
export type IncomeData = {
  total: number;
  lines: { text: string; amount: number }[];
};
export type LabRow = { group: string; type: string; requests: number; patients: number; fixes: number; remakes: number; teeth: number };
export type LabData = { rows: LabRow[]; vendors: { name: string; count: number }[] };

export type Month = {
  ym: string;
  cost?: CostData;
  route?: RouteData;
  consult?: ConsultRow[];
  income?: IncomeData;
  lab?: LabData;
};

const num = (v: unknown): number =>
  typeof v === "number" ? v : parseFloat(String(v ?? "").replace(/[^0-9.\-]/g, "")) || 0;
const text = (v: unknown): string => String(v ?? "").trim();
const grid = (ws: XLSX.WorkSheet): unknown[][] =>
  XLSX.utils.sheet_to_json<unknown[]>(ws, { header: 1, defval: "", blankrows: false });

/** "2,139만원", "1억 2,000만원" -> 원 */
function won(v: unknown): number {
  const s = text(v).replace(/,/g, "");
  const eok = s.match(/(\d+(?:\.\d+)?)억/);
  const man = s.match(/(\d+(?:\.\d+)?)만/);
  if (!eok && !man) return num(s);
  return (eok ? parseFloat(eok[1]) * 1e8 : 0) + (man ? parseFloat(man[1]) * 1e4 : 0);
}

function colIndex(header: unknown[], name: string): number {
  return header.findIndex((h) => text(h).replace(/\s/g, "") === name.replace(/\s/g, ""));
}

// ---------- 파일 풀기 ----------

export async function expandFiles(files: File[]): Promise<{ name: string; data: ArrayBuffer }[]> {
  const out: { name: string; data: ArrayBuffer }[] = [];
  for (const f of files) {
    if (/\.zip$/i.test(f.name)) {
      const zip = await JSZip.loadAsync(await f.arrayBuffer());
      for (const e of Object.values(zip.files)) {
        if (e.dir) continue;
        const base = e.name.split("/").pop() ?? e.name;
        if (base.startsWith("~$") || !/\.xlsx?$/i.test(base)) continue;
        out.push({ name: base, data: await e.async("arraybuffer") });
      }
    } else if (/\.xlsx?$/i.test(f.name)) {
      out.push({ name: f.name, data: await f.arrayBuffer() });
    }
  }
  return out;
}

function detectKind(wb: XLSX.WorkBook): Kind | null {
  const head = grid(wb.Sheets[wb.SheetNames[0]])
    .slice(0, 3)
    .flat()
    .map((c) => text(c).replace(/\s/g, ""));
  const has = (s: string) => head.includes(s);
  if (has("총진료비") && has("청구액")) return "cost";
  if (has("내원경로") && has("신환수")) return "route";
  if (has("상담자명")) return "consult";
  if (has("수입항목")) return "income";
  if (has("기공물종류")) return "lab";
  return null;
}

function monthFromName(name: string): string | null {
  const m = name.match(/(\d{4})\s*년\s*(\d{1,2})\s*월/);
  return m ? `${m[1]}-${m[2].padStart(2, "0")}` : null;
}

// ---------- 리포트별 파서 ----------

function parseCost(wb: XLSX.WorkBook): { data: CostData; ym: string | null } {
  const g = grid(wb.Sheets[wb.SheetNames[0]]);
  const h = g[0] ?? [];
  const ix = (n: string) => colIndex(h, n);
  const c = {
    total: ix("총 진료비"), claim: ix("청구액"), copay: ix("본인부담금"), nonIns: ix("비급여"),
    received: ix("총수납액"), card: ix("카드수납액"), cash: ix("현금수납액"), online: ix("기타(온라인)"),
    cashReceipt: ix("현영발행액"), discount: ix("할인금액"),
  };
  const days = g.filter((r) => /^\d{4}-\d{2}-\d{2}$/.test(text(r[0])));
  const sum = (i: number) => (i < 0 ? 0 : days.reduce((a, r) => a + num(r[i]), 0));
  const first = days[0] ? text(days[0][0]) : "";
  return {
    ym: first ? first.slice(0, 7) : null,
    data: {
      days: days.length,
      total: sum(c.total), claim: sum(c.claim), copay: sum(c.copay), nonIns: sum(c.nonIns),
      received: sum(c.received), card: sum(c.card), cash: sum(c.cash), online: sum(c.online),
      cashReceipt: sum(c.cashReceipt), discount: sum(c.discount),
    },
  };
}

/** 주소에서 구·군 단위만 꺼낸다. */
function regionOf(addr: string): string {
  const m = addr.match(/([가-힣]+[구군])(?=\s|$|[0-9])/);
  return m ? m[1] : "기타·미기재";
}

function parseRoute(wb: XLSX.WorkBook): RouteData {
  const g = grid(wb.Sheets[wb.SheetNames[0]]);
  const hi = g.findIndex((r) => text(r[0]) === "내원경로");
  const h = g[hi] ?? [];
  const ix = (n: string) => colIndex(h, n);
  const c = { patients: ix("내원환자수"), ret: ix("구환수"), fresh: ix("신환수"), visits: ix("총내원횟수"), avg: ix("평균진료비"), total: ix("총 진료비") };
  const rows: RouteRow[] = [];
  let totals = { patients: 0, ret: 0, fresh: 0, visits: 0, total: 0 };
  let hasTotal = false;
  for (const r of g.slice(hi + 1)) {
    const name = text(r[0]);
    if (!name) continue;
    const row = {
      patients: num(r[c.patients]), ret: num(r[c.ret]), fresh: num(r[c.fresh]),
      visits: num(r[c.visits]), total: num(r[c.total]),
    };
    if (name === "합계") {
      totals = row;
      hasTotal = true;
    } else if (name !== "평균") {
      rows.push({ name, ...row, avg: num(r[c.avg]) });
    }
  }
  if (!hasTotal) {
    totals = rows.reduce(
      (a, r) => ({ patients: a.patients + r.patients, ret: a.ret + r.ret, fresh: a.fresh + r.fresh, visits: a.visits + r.visits, total: a.total + r.total }),
      totals,
    );
  }

  // 신환 시트: 주소(구 단위)와 소개자(메모에 적힌 성명·차트번호)만 집계하고, 신환 본인의 이름·연락처는 읽지 않는다.
  const regions = new Map<string, number>();
  const referrers = new Map<string, { name: string; count: number }>();
  for (const name of wb.SheetNames.slice(1)) {
    if (!name.startsWith("신환")) continue;
    const sg = grid(wb.Sheets[name]);
    const sh = sg[0] ?? [];
    const addr = colIndex(sh, "주소");
    const memo = colIndex(sh, "메모");
    for (const r of sg.slice(1)) {
      if (!text(r[0])) continue;
      const k = addr >= 0 ? regionOf(text(r[addr])) : "기타·미기재";
      regions.set(k, (regions.get(k) ?? 0) + 1);
      if (name.includes("소개") && memo >= 0) {
        // 메모 예: "홍길동(61060301)님 자녀" -> 성명 홍길동, 차트번호 61060301
        const m = text(r[memo]).match(/([가-힣A-Za-z]+)\s*\((\d{3,9})\)/);
        if (m) {
          const prev = referrers.get(m[2]);
          referrers.set(m[2], { name: prev?.name ?? m[1], count: (prev?.count ?? 0) + 1 });
        }
      }
    }
  }
  return { rows, totals, regions, referrers };
}

function parseConsult(wb: XLSX.WorkBook): ConsultRow[] {
  const g = grid(wb.Sheets[wb.SheetNames[0]]);
  const hi = g.findIndex((r) => text(r[0]) === "상담자명");
  const h = g[hi] ?? [];
  const ix = (n: string) => colIndex(h, n);
  const c = {
    plans: ix("계획 수"), confirmed: ix("확정건수"), patients: ix("상담환자 수"),
    cp: ix("확정환자수"), amount: ix("확정금액"), discount: ix("평균할인률"),
  };
  return g
    .slice(hi + 1)
    .filter((r) => text(r[0]))
    .map((r) => ({
      name: text(r[0]),
      plans: num(r[c.plans]), confirmed: num(r[c.confirmed]), patients: num(r[c.patients]),
      confirmedPatients: num(r[c.cp]), amount: won(r[c.amount]), discount: num(r[c.discount]),
    }));
}

function parseIncome(wb: XLSX.WorkBook): { data: IncomeData; ym: string | null } {
  const ws = wb.Sheets["진료 수입"];
  const g = ws ? grid(ws) : [];
  const h = g[0] ?? [];
  const di = colIndex(h, "날짜");
  const ti = colIndex(h, "수입내용");
  const ai = colIndex(h, "금액");
  const lines = g.slice(1).map((r) => ({ text: text(r[ti]), amount: num(r[ai]) }));
  const first = g[1] && di >= 0 ? text(g[1][di]) : "";
  return {
    ym: /^\d{4}-\d{2}/.test(first) ? first.slice(0, 7) : null,
    data: { total: lines.reduce((a, l) => a + l.amount, 0), lines },
  };
}

function parseLab(wb: XLSX.WorkBook): LabData {
  const g = grid(wb.Sheets[wb.SheetNames[0]]);
  const hi = g.findIndex((r) => text(r[0]) === "구분");
  const h = g[hi] ?? [];
  const ix = (n: string) => colIndex(h, n);
  const c = { type: ix("기공물 종류"), req: ix("의뢰횟수"), pat: ix("환자수"), fix: ix("수정횟수"), re: ix("재제작횟수"), teeth: ix("치아수") };
  let group = "";
  const rows: LabRow[] = [];
  for (const r of g.slice(hi + 1)) {
    const gcell = text(r[0]).replace(/"/g, "").trim();
    if (gcell) group = gcell;
    if (!text(r[c.type])) continue;
    rows.push({
      group, type: text(r[c.type]), requests: num(r[c.req]), patients: num(r[c.pat]),
      fixes: num(r[c.fix]), remakes: num(r[c.re]), teeth: num(r[c.teeth]),
    });
  }
  const vendors = wb.SheetNames.slice(1)
    .map((n) => {
      const m = n.match(/^(.*?)\s*\((\d+)\)\s*$/);
      return m ? { name: m[1], count: Number(m[2]) } : null;
    })
    .filter((v): v is { name: string; count: number } => v !== null);
  return { rows, vendors };
}

// ---------- 월별 묶기 ----------

export type LoadResult = { months: Map<string, Month>; unknown: string[] };

export function ingest(files: { name: string; data: ArrayBuffer }[], into?: Map<string, Month>): LoadResult {
  const months = new Map(into ?? []);
  const unknown: string[] = [];
  for (const f of files) {
    let wb: XLSX.WorkBook;
    try {
      wb = XLSX.read(f.data, { type: "array" });
    } catch {
      unknown.push(f.name);
      continue;
    }
    const kind = detectKind(wb);
    if (!kind) {
      unknown.push(f.name);
      continue;
    }
    let ym = monthFromName(f.name);
    let patch: Partial<Month> = {};
    if (kind === "cost") {
      const p = parseCost(wb);
      ym = ym ?? p.ym;
      patch = { cost: p.data };
    } else if (kind === "income") {
      const p = parseIncome(wb);
      ym = ym ?? p.ym;
      patch = { income: p.data };
    } else if (kind === "route") patch = { route: parseRoute(wb) };
    else if (kind === "consult") patch = { consult: parseConsult(wb) };
    else patch = { lab: parseLab(wb) };
    if (!ym) {
      unknown.push(f.name);
      continue;
    }
    months.set(ym, { ...(months.get(ym) ?? { ym }), ym, ...patch });
  }
  return { months, unknown };
}

export function prevYm(ym: string): string {
  const [y, m] = ym.split("-").map(Number);
  return m === 1 ? `${y - 1}-12` : `${y}-${String(m - 1).padStart(2, "0")}`;
}
export const ymLabel = (ym: string) => `${ym.slice(0, 4)}년 ${Number(ym.slice(5))}월`;

// ---------- 치료 분류 ----------

/** 위에서부터 먼저 걸리는 분류로 묶는다. 수납 한 줄에 여러 치료가 섞여 있으면 위쪽 분류로 잡힌다. */
export const DEFAULT_RULES = `임플란트: 임플란트, 식립, 고정체, 스캔바디, 어버트먼트
보철: 지르코니아, 인레이, 온레이, 크라운, 브릿지, 보철, 틀니, 의치, 임시치아, 영구접착, 셋팅, 라미네이트
교정: 교정, 리테이너, 브라켓
신경치료: 근관, 신경, 발수, 근충, 치근단 충전, MTA
발치·수술: 발치, 수술, 사랑니, 치조골, 절개, 봉합
치주: 치근활택, 치석, 치주, 스케일링, 잇몸, 미노클린, 클린팁
보존: 레진, 충전, 치경부, GI, 홈메우기, 실란트
진단·검사: 파노라마, 교익촬영, 치근단, 진단, 치료계획, 방사선`;

export function parseRules(t: string): [string, string[]][] {
  return t
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean)
    .map((l) => {
      const i = l.indexOf(":");
      if (i < 0) return null;
      const name = l.slice(0, i).trim();
      const kws = l.slice(i + 1).split(",").map((s) => s.trim().toLowerCase()).filter(Boolean);
      return name && kws.length ? ([name, kws] as [string, string[]]) : null;
    })
    .filter((x): x is [string, string[]] => x !== null);
}

export function treatmentMix(inc: IncomeData, rules: [string, string[]][]): Map<string, number> {
  const out = new Map<string, number>();
  for (const l of inc.lines) {
    const t = l.text.toLowerCase();
    const cat = rules.find(([, kws]) => kws.some((k) => t.includes(k)))?.[0] ?? "기타";
    out.set(cat, (out.get(cat) ?? 0) + l.amount);
  }
  return out;
}

// ---------- 지표 ----------

export type Metrics = {
  total: number | null;
  claim: number | null;
  copay: number | null;
  insured: number | null;
  nonIns: number | null;
  nonInsShare: number | null;
  received: number | null;
  card: number | null;
  cash: number | null;
  online: number | null;
  cashReceipt: number | null;
  days: number | null;
  perDay: number | null;
  patients: number | null;
  fresh: number | null;
  ret: number | null;
  visits: number | null;
  visitsPer: number | null;
  perDayPatients: number | null;
  perDayFresh: number | null;
  perPatient: number | null;
  freshShare: number | null;
  c: { plans: number; confirmed: number; patients: number; confirmedPatients: number; amount: number; discount: number } | null;
  lab: { requests: number; remakes: number; patients: number } | null;
};

const r1 = (v: number) => Math.round(v * 10) / 10;

export function metrics(m: Month | undefined): Metrics {
  const c = m?.cost;
  const r = m?.route;
  const total = c?.total ?? r?.totals.total ?? null;
  const patients = r?.totals.patients ?? null;
  const k = m?.consult;
  const consult = k && k.length
    ? k.reduce(
        (a, x) => ({
          plans: a.plans + x.plans, confirmed: a.confirmed + x.confirmed, patients: a.patients + x.patients,
          confirmedPatients: a.confirmedPatients + x.confirmedPatients, amount: a.amount + x.amount,
          discount: a.discount + x.discount,
        }),
        { plans: 0, confirmed: 0, patients: 0, confirmedPatients: 0, amount: 0, discount: 0 },
      )
    : null;
  const lab = m?.lab
    ? m.lab.rows.reduce(
        (a, x) => ({ requests: a.requests + x.requests, remakes: a.remakes + x.remakes, patients: a.patients + x.patients }),
        { requests: 0, remakes: 0, patients: 0 },
      )
    : null;
  return {
    total,
    claim: c?.claim ?? null,
    copay: c?.copay ?? null,
    insured: c ? c.claim + c.copay : null,
    nonIns: c?.nonIns ?? null,
    nonInsShare: c && c.total ? r1((c.nonIns / c.total) * 100) : null,
    received: c?.received ?? null,
    card: c?.card ?? null,
    cash: c?.cash ?? null,
    online: c?.online ?? null,
    cashReceipt: c?.cashReceipt ?? null,
    days: c?.days ?? null,
    perDay: c && c.days ? c.total / c.days : null,
    patients,
    fresh: r?.totals.fresh ?? null,
    ret: r?.totals.ret ?? null,
    visits: r?.totals.visits ?? null,
    visitsPer: r && r.totals.patients ? r1(r.totals.visits / r.totals.patients) : null,
    perDayPatients: r && c?.days ? r1(r.totals.patients / c.days) : null,
    perDayFresh: r && c?.days ? r1(r.totals.fresh / c.days) : null,
    perPatient: total !== null && patients ? total / patients : null,
    freshShare: r && r.totals.patients ? r1((r.totals.fresh / r.totals.patients) * 100) : null,
    c: consult,
    lab,
  };
}

// ---------- 표 ----------

export type Cell = string | number | null;
export type Table = { id: string; title: string; note?: string; headers: string[]; rows: Cell[][] };

type V = number | null | undefined;
type Kd = "n" | "d" | "p"; // n: 원·명(증감률 포함), d: 소수 한 자리, p: % (증감은 %p)

function row(label: string, a: V, b: V, kind: Kd = "n"): Cell[] {
  const A = a ?? null;
  const B = b ?? null;
  const fmt = (v: number | null) => (v === null ? null : kind === "n" ? Math.round(v) : r1(v));
  const diff = A === null || B === null ? null : kind === "n" ? Math.round(A - B) : r1(A - B);
  const pc = kind !== "n" || A === null || B === null || B === 0 ? null : r1(((A - B) / B) * 100);
  return [label, fmt(A), fmt(B), diff, pc];
}
const share = (a: number, tot: number): number | null => (tot ? r1((a / tot) * 100) : null);
const CMP = ["구분", "당월", "전월", "증감", "증감률(%)"];

export function buildTables(
  months: Map<string, Month>,
  cur: string,
  prev: string | null,
  rules: [string, string[]][],
): Table[] {
  const C = months.get(cur);
  const P = prev ? months.get(prev) : undefined;
  const mc = metrics(C);
  const mp = metrics(P);
  const out: Table[] = [];

  if (C?.cost) {
    out.push({
      id: "t31",
      title: "3-1. 총 진료비 구성 당월·전월 비교 (원)",
      note: "급여 진료비 = 공단 청구액 + 환자 본인부담금. 총 수납액은 환자가 실제 낸 금액이라 공단 입금분은 포함되지 않습니다.",
      headers: CMP,
      rows: [
        row("총 진료비", mc.total, mp.total),
        row("급여 진료비", mc.insured, mp.insured),
        row("  · 공단 청구액", mc.claim, mp.claim),
        row("  · 환자 본인부담금", mc.copay, mp.copay),
        row("비급여 진료비", mc.nonIns, mp.nonIns),
        row("비급여 비중(%)", mc.nonInsShare, mp.nonInsShare, "p"),
        row("총 수납액", mc.received, mp.received),
        row("  · 카드", mc.card, mp.card),
        row("  · 현금", mc.cash, mp.cash),
        row("  · 기타(온라인)", mc.online, mp.online),
        row("현금영수증 발행액", mc.cashReceipt, mp.cashReceipt),
        row("영업일수(일)", mc.days, mp.days),
        row("일평균 진료비", mc.perDay, mp.perDay),
      ],
    });
  }

  if (C?.route) {
    out.push({
      id: "t32",
      title: "3-2. 내원 환자 수 당월·전월 비교",
      headers: CMP,
      rows: [
        row("총 내원 환자 수(명)", mc.patients, mp.patients),
        row("신환(명)", mc.fresh, mp.fresh),
        row("구환(명)", mc.ret, mp.ret),
        row("신환 비율(%)", mc.freshShare, mp.freshShare, "p"),
        row("총 내원 횟수(회)", mc.visits, mp.visits),
        row("환자 1인당 내원 횟수(회)", mc.visitsPer, mp.visitsPer, "d"),
        row("일평균 내원 환자(명)", mc.perDayPatients, mp.perDayPatients, "d"),
        row("일평균 신환(명)", mc.perDayFresh, mp.perDayFresh, "d"),
        row("환자 1인당 평균 진료비(원)", mc.perPatient, mp.perPatient),
      ],
    });
  }

  if (C?.income) {
    const a = treatmentMix(C.income, rules);
    const b = P?.income ? treatmentMix(P.income, rules) : null;
    const ta = [...a.values()].reduce((x, y) => x + y, 0);
    const tb = b ? [...b.values()].reduce((x, y) => x + y, 0) : 0;
    const cats = new Set([...a.keys(), ...(b?.keys() ?? [])]);
    const rows = [...cats]
      .map((k) => ({ k, x: a.get(k) ?? 0, y: b ? (b.get(k) ?? 0) : null }))
      .sort((p, q) => q.x - p.x)
      .map(({ k, x, y }): Cell[] => {
        const sa = share(x, ta);
        const sb = y === null ? null : share(y, tb);
        return [k, Math.round(x), sa, y === null ? null : Math.round(y), sb, sa !== null && sb !== null ? r1(sa - sb) : null];
      });
    out.push({
      id: "t33",
      title: "3-3. 주요 치료별 수납액 비중 비교",
      note: "수입 통계의 진료 수입(환자 수납액)을 치료 내용 문구로 분류했습니다. 한 건에 여러 치료가 섞여 있으면 분류 규칙의 위쪽 항목으로 묶이므로 근사치입니다.",
      headers: ["치료 항목", "당월 수납액(원)", "당월 비중(%)", "전월 수납액(원)", "전월 비중(%)", "비중 변화(%p)"],
      rows,
    });
  }

  if (C?.route) {
    const R = C.route;
    const pm = new Map((P?.route?.rows ?? []).map((x) => [x.name, x]));
    out.push({
      id: "t34",
      title: "3-4. 내원 경로별 현황",
      headers: ["내원 경로", "내원 환자(명)", "비중(%)", "신환(명)", "구환(명)", "내원 횟수(회)", "평균 진료비(원)", "총 진료비(원)", "전월 환자(명)", "환자 증감(명)", "전월 신환(명)", "신환 증감(명)"],
      rows: [...R.rows]
        .sort((x, y) => y.patients - x.patients)
        .map((x): Cell[] => {
          const p = P?.route ? pm.get(x.name) : undefined;
          const known = !!P?.route;
          return [
            x.name, x.patients, share(x.patients, R.totals.patients), x.fresh, x.ret, x.visits,
            Math.round(x.avg), x.total,
            known ? (p?.patients ?? 0) : null, known ? x.patients - (p?.patients ?? 0) : null,
            known ? (p?.fresh ?? 0) : null, known ? x.fresh - (p?.fresh ?? 0) : null,
          ];
        }),
    });

    if (R.regions.size) {
      const tot = [...R.regions.values()].reduce((x, y) => x + y, 0);
      const keys = new Set([...R.regions.keys(), ...(P?.route?.regions.keys() ?? [])]);
      out.push({
        id: "t34b",
        title: "3-4. 신환 거주 지역(구)별",
        note: "신환 등록 주소에서 구·군 단위만 집계했습니다.",
        headers: ["지역", "당월 신환(명)", "비중(%)", "전월 신환(명)", "증감(명)"],
        rows: [...keys]
          .map((k) => ({ k, n: R.regions.get(k) ?? 0, p: P?.route ? (P.route.regions.get(k) ?? 0) : null }))
          .filter(({ n, p }) => n > 0 || (p ?? 0) > 0)
          .sort((x, y) => y.n - x.n)
          .map(({ k, n, p }): Cell[] => [k, n, share(n, tot), p, p === null ? null : n - p]),
      });
    }

    const sg = (rt: RouteData | undefined) => rt?.rows.find((x) => x.name === "소개");
    const a = sg(R);
    const b = sg(P?.route);
    if (a) {
      const m = (x: RouteRow | undefined, rt: RouteData | undefined) => ({
        patients: x?.patients ?? null,
        share: x && rt ? share(x.patients, rt.totals.patients) : null,
        fresh: x?.fresh ?? null,
        freshShare: x && rt ? share(x.fresh, rt.totals.fresh) : null,
        total: x?.total ?? null,
        avg: x?.avg ?? null,
      });
      const A = m(a, R);
      const B = P?.route ? m(b, P.route) : m(undefined, undefined);
      out.push({
        id: "t35",
        title: "3-5. 소개환자 분석",
        note: "내원 경로가 '소개'로 등록된 환자 기준입니다. 원장님·직원 지인은 3-4 표에서 따로 봅니다.",
        headers: CMP,
        rows: [
          row("소개 내원 환자(명)", A.patients, B.patients),
          row("전체 내원 환자 중 비중(%)", A.share, B.share, "p"),
          row("소개 신환(명)", A.fresh, B.fresh),
          row("전체 신환 중 소개 비중(%)", A.freshShare, B.freshShare, "p"),
          row("소개환자 총 진료비(원)", A.total, B.total),
          row("소개환자 평균 진료비(원)", A.avg, B.avg),
        ],
      });
      if (R.referrers.size) {
        out.push({
          id: "t35b",
          title: "3-5. 소개를 많이 해 준 환자 (당월 신환 기준)",
          note: "신환 메모에 적힌 소개자 성명을 집계했습니다. 감사 인사·혜택 안내 대상 확인용이며, 성명이 들어 있으므로 인쇄·공유 전에 확인해 주세요.",
          headers: ["소개자 성명", "소개 신환(명)"],
          rows: [...R.referrers.values()].sort((x, y) => y.count - x.count).slice(0, 10).map((v) => [v.name, v.count]),
        });
      }
    }
  }

  if (C?.consult && mc.c) {
    const a = mc.c;
    const b = mp.c;
    const rate = (x: number, y: number) => (y ? r1((x / y) * 100) : null);
    out.push({
      id: "t36",
      title: "3-6. 상담·치료 확정 현황 (전체 상담자 합계)",
      note: "상담환자 수는 상담자별 수를 더한 값이라 같은 환자가 여러 상담자에게 상담받았다면 중복됩니다.",
      headers: CMP,
      rows: [
        row("치료 계획 수(건)", a.plans, b?.plans),
        row("확정 건수(건)", a.confirmed, b?.confirmed),
        row("확정 비율(%)", rate(a.confirmed, a.plans), b ? rate(b.confirmed, b.plans) : null, "p"),
        row("상담 환자 수(명)", a.patients, b?.patients),
        row("확정 환자 수(명)", a.confirmedPatients, b?.confirmedPatients),
        row("확정 환자 비율(%)", rate(a.confirmedPatients, a.patients), b ? rate(b.confirmedPatients, b.patients) : null, "p"),
        row("확정 금액(원)", a.amount, b?.amount),
        row("확정 환자 1인당 금액(원)", a.confirmedPatients ? a.amount / a.confirmedPatients : null, b && b.confirmedPatients ? b.amount / b.confirmedPatients : null),
      ],
    });
    const pr = new Map((P?.consult ?? []).map((x) => [x.name, x]));
    out.push({
      id: "t36b",
      title: "3-6. 상담자별 확정 현황",
      headers: ["상담자", "계획(건)", "확정(건)", "확정 비율(%)", "상담 환자(명)", "확정 환자(명)", "확정 환자 비율(%)", "확정 금액(원)", "평균 할인률(%)", "전월 확정 비율(%)", "비율 변화(%p)"],
      rows: [...C.consult]
        .sort((x, y) => y.amount - x.amount)
        .map((x): Cell[] => {
          const cr = rate(x.confirmed, x.plans);
          const p = pr.get(x.name);
          const pr1 = p ? rate(p.confirmed, p.plans) : null;
          return [
            x.name, x.plans, x.confirmed, cr, x.patients, x.confirmedPatients,
            rate(x.confirmedPatients, x.patients), Math.round(x.amount), x.discount, pr1,
            cr !== null && pr1 !== null ? r1(cr - pr1) : null,
          ];
        }),
    });
  }

  if (C?.lab && mc.lab) {
    const grp = (l: LabData | undefined) => {
      const o = new Map<string, { req: number; re: number }>();
      for (const x of l?.rows ?? []) {
        const g = o.get(x.group) ?? { req: 0, re: 0 };
        g.req += x.requests;
        g.re += x.remakes;
        o.set(x.group, g);
      }
      return o;
    };
    const ga = grp(C.lab);
    const gb = P?.lab ? grp(P.lab) : null;
    const rows: Cell[][] = [...ga.entries()]
      .sort((x, y) => y[1].req - x[1].req)
      .map(([k, v]) => [
        k, v.req, gb ? (gb.get(k)?.req ?? 0) : null, gb ? v.req - (gb.get(k)?.req ?? 0) : null,
        v.re, share(v.re, v.req),
      ]);
    rows.push([
      "합계", mc.lab.requests, P?.lab ? mp.lab!.requests : null, P?.lab ? mc.lab.requests - mp.lab!.requests : null,
      mc.lab.remakes, share(mc.lab.remakes, mc.lab.requests),
    ]);
    out.push({
      id: "t37",
      title: "3-7. 기공 의뢰 현황",
      note: "재제작률 = 재제작 횟수 ÷ 의뢰 횟수. 업체별 의뢰: " + C.lab.vendors.map((v) => `${v.name} ${v.count}건`).join(", "),
      headers: ["구분", "당월 의뢰(건)", "전월 의뢰(건)", "증감(건)", "당월 재제작(건)", "재제작률(%)"],
      rows,
    });
  }

  // 월별 추이 (선택한 당월까지 최근 6개월)
  const list = [...months.keys()].filter((k) => k <= cur).sort().slice(-6);
  if (list.length > 1) {
    out.push({
      id: "t38",
      title: "3-8. 월별 추이",
      headers: ["월", "총 진료비(원)", "비급여 비중(%)", "내원 환자(명)", "신환(명)", "환자 1인당 진료비(원)", "상담 확정 금액(원)", "상담 확정 비율(%)", "기공 의뢰(건)"],
      rows: list.map((k): Cell[] => {
        const x = metrics(months.get(k));
        return [
          ymLabel(k), x.total === null ? null : Math.round(x.total), x.nonInsShare, x.patients, x.fresh,
          x.perPatient === null ? null : Math.round(x.perPatient), x.c ? Math.round(x.c.amount) : null,
          x.c ? share(x.c.confirmed, x.c.plans) : null, x.lab?.requests ?? null,
        ];
      }),
    });
  }

  if (mc.total !== null && mp.total !== null && mc.patients && mp.patients && mc.perPatient !== null && mp.perPatient !== null) {
    const vol = (mc.patients - mp.patients) * mp.perPatient;
    const prc = (mc.perPatient - mp.perPatient) * mc.patients;
    const d = mc.total - mp.total;
    out.push({
      id: "t39",
      title: "3-9. 총 진료비 증감 요인 분해 (원)",
      note: "총 진료비 = 내원 환자 수 × 환자 1인당 진료비. 환자 수 효과 = (당월−전월 환자 수) × 전월 1인당 진료비, 단가 효과 = (당월−전월 1인당 진료비) × 당월 환자 수.",
      headers: ["요인", "총 진료비 증감(원)", "총 증감 대비(%)"],
      rows: [
        ["내원 환자 수 변화", Math.round(vol), share(vol, d)],
        ["환자 1인당 진료비 변화", Math.round(prc), share(prc, d)],
        ["합계", Math.round(d), d ? 100 : null],
      ],
    });
  }
  return out;
}

export type Growth = { label: string; unit: string; def: number; estimate: (m: Metrics, x: number) => number | null };

/** 개선 시뮬레이션: 입력한 만큼 좋아졌을 때 월 진료비 증가 추정치(원) */
export const GROWTH: Growth[] = [
  {
    label: "신환이 늘어난다면",
    unit: "%",
    def: 10,
    estimate: (m, x) => (m.fresh !== null && m.perPatient !== null ? Math.round(m.fresh * (x / 100) * m.perPatient) : null),
  },
  {
    label: "상담 확정 비율이 오른다면",
    unit: "%p",
    def: 5,
    estimate: (m, x) =>
      m.c && m.c.plans && m.c.confirmed ? Math.round(m.c.plans * (x / 100) * (m.c.amount / m.c.confirmed)) : null,
  },
  {
    label: "환자 1인당 진료비가 오른다면",
    unit: "%",
    def: 5,
    estimate: (m, x) => (m.total !== null ? Math.round(m.total * (x / 100)) : null),
  },
];
