import * as XLSX from "xlsx";

export type FieldKey =
  | "date"
  | "chartNo"
  | "treatment"
  | "amount"
  | "insurance"
  | "patientType"
  | "source"
  | "region"
  | "referrer"
  | "apptStatus";

export const FIELDS: {
  key: FieldKey;
  label: string;
  hints: string[];
  needed: string;
}[] = [
  { key: "amount", label: "진료비(금액)", hints: ["진료비", "금액", "수납", "청구", "매출"], needed: "총 진료비, 치료별 비중" },
  { key: "chartNo", label: "차트번호(환자 식별)", hints: ["차트", "환자번호", "환자no", "등록번호"], needed: "내원 환자 수" },
  { key: "date", label: "진료일", hints: ["진료일", "내원일", "일자", "날짜", "수납일"], needed: "영업일수, 일평균" },
  { key: "treatment", label: "치료 항목", hints: ["치료", "진료항목", "수가명", "처치", "항목"], needed: "치료별 비중" },
  { key: "insurance", label: "보험/비보험 구분", hints: ["보험", "급여", "구분"], needed: "보험·비보험 분리" },
  { key: "patientType", label: "신환/구환 구분", hints: ["초재진", "신환", "초진", "구분"], needed: "신환·구환" },
  { key: "source", label: "내원 경로", hints: ["내원경로", "경로", "유입", "알게된"], needed: "내원 경로, 소개환자" },
  { key: "region", label: "거주 지역(동·구)", hints: ["지역", "주소", "동", "구"], needed: "지역별 현황" },
  { key: "referrer", label: "소개자", hints: ["소개자", "소개한", "추천인"], needed: "소개해 준 환자 순위" },
  { key: "apptStatus", label: "예약 상태", hints: ["예약상태", "내원여부", "상태", "이행"], needed: "예약 이행률" },
];

export type Sheet = { headers: string[]; rows: Record<string, unknown>[] };
export type Mapping = Partial<Record<FieldKey, string>>;

/** 엑셀/CSV 파일을 브라우저에서 읽는다. 서버로 전송하지 않는다. */
export async function readSheet(file: File): Promise<Sheet> {
  const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const grid = XLSX.utils.sheet_to_json<unknown[]>(ws, {
    header: 1,
    defval: "",
    blankrows: false,
  });
  // 제목 줄이 있는 경우를 위해, 앞 10행 중 문자 셀이 가장 많은 행을 머리글로 본다.
  let headerIdx = 0;
  let best = -1;
  grid.slice(0, 10).forEach((row, i) => {
    const score = row.filter(
      (c) => typeof c === "string" && c.trim() !== "",
    ).length;
    if (score > best) {
      best = score;
      headerIdx = i;
    }
  });
  const seen = new Map<string, number>();
  const headers = (grid[headerIdx] ?? []).map((h, i) => {
    const base = String(h ?? "").trim() || `열${i + 1}`;
    const n = (seen.get(base) ?? 0) + 1;
    seen.set(base, n);
    return n === 1 ? base : `${base}(${n})`;
  });
  const rows = grid.slice(headerIdx + 1).map((r) => {
    const o: Record<string, unknown> = {};
    headers.forEach((h, i) => (o[h] = r[i] ?? ""));
    return o;
  });
  return { headers, rows };
}

export function guessMapping(headers: string[]): Mapping {
  const m: Mapping = {};
  const used = new Set<string>();
  for (const f of FIELDS) {
    const hit = headers.find(
      (h) => !used.has(h) && f.hints.some((k) => h.replace(/\s/g, "").includes(k)),
    );
    if (hit) {
      m[f.key] = hit;
      used.add(hit);
    }
  }
  return m;
}

// ---------- 치료 분류 ----------

export const DEFAULT_RULES = `임플란트: 임플란트, 임플, 픽스처, 어버트먼트, implant
교정: 교정, 인비절, 브라켓, 투명교정
보철: 크라운, 브릿지, 틀니, 의치, 보철, 인레이, 온레이
심미: 라미네이트, 미백, 심미
신경치료: 신경, 근관, 근충
치주: 치주, 스케일링, 잇몸, 치석, 판막
보존: 레진, 충전, 치아우식, 글라스
소아: 소아, 유치, 실란트, 불소
발치·수술: 발치, 사랑니, 수술, 절제`;

export function parseRules(text: string): [string, string[]][] {
  return text
    .split("\n")
    .map((l) => l.trim())
    .filter(Boolean)
    .map((l) => {
      const i = l.indexOf(":");
      if (i < 0) return null;
      const name = l.slice(0, i).trim();
      const kws = l
        .slice(i + 1)
        .split(",")
        .map((s) => s.trim().toLowerCase())
        .filter(Boolean);
      return name && kws.length ? ([name, kws] as [string, string[]]) : null;
    })
    .filter((x): x is [string, string[]] => x !== null);
}

function categorize(name: string, rules: [string, string[]][]): string {
  const t = name.toLowerCase();
  for (const [cat, kws] of rules) if (kws.some((k) => t.includes(k))) return cat;
  return "기타";
}

// ---------- 월별 집계 ----------

const num = (v: unknown): number =>
  typeof v === "number" ? v : parseFloat(String(v ?? "").replace(/[^0-9.\-]/g, "")) || 0;
const str = (v: unknown): string => String(v ?? "").trim();

export type Stats = {
  rowCount: number;
  total: number | null;
  insured: number | null;
  nonInsured: number | null;
  visitors: number | null;
  newPatients: number | null;
  returning: number | null;
  days: number | null;
  perPatient: number | null;
  treatment: Map<string, number> | null;
  appt: { total: number; done: number; cancel: number; noshow: number } | null;
  sources: Map<string, { patients: number; amount: number }> | null;
  regions: Map<string, { patients: number; amount: number }> | null;
  referral: {
    patients: number;
    amount: number;
    referrers: Map<string, number>;
  } | null;
};

export function computeStats(
  sheet: Sheet,
  m: Mapping,
  rules: [string, string[]][],
): Stats {
  const g = (r: Record<string, unknown>, k: FieldKey) => (m[k] ? r[m[k]!] : undefined);

  // 환자 단위 첫 행(신환 판정, 경로, 지역, 소개자)
  const patientFirst = new Map<string, Record<string, unknown>>();
  const patientAmount = new Map<string, number>();
  const newSet = new Set<string>();
  const dates = new Set<string>();
  let total = 0;
  let insured = 0;
  let nonInsured = 0;
  const treatment = new Map<string, number>();
  const appt = { total: 0, done: 0, cancel: 0, noshow: 0 };

  for (const r of sheet.rows) {
    const chart = m.chartNo ? str(g(r, "chartNo")) : "";
    const amt = m.amount ? num(g(r, "amount")) : 0;
    total += amt;
    if (chart) {
      if (!patientFirst.has(chart)) patientFirst.set(chart, r);
      patientAmount.set(chart, (patientAmount.get(chart) ?? 0) + amt);
      if (m.patientType && /신|초진/.test(str(g(r, "patientType")))) newSet.add(chart);
    }
    if (m.date) {
      const d = str(g(r, "date"));
      if (d) dates.add(d);
    }
    if (m.insurance) {
      const v = str(g(r, "insurance"));
      if (/비보험|비급여/.test(v)) nonInsured += amt;
      else if (/보험|급여/.test(v)) insured += amt;
    }
    if (m.treatment) {
      const c = categorize(str(g(r, "treatment")), rules);
      treatment.set(c, (treatment.get(c) ?? 0) + amt);
    }
    if (m.apptStatus) {
      const v = str(g(r, "apptStatus"));
      if (v) {
        appt.total++;
        if (/노쇼|부도|미내원|불참/.test(v)) appt.noshow++;
        else if (/취소/.test(v)) appt.cancel++;
        else appt.done++;
      }
    }
  }

  const hasChart = !!m.chartNo;
  const visitors = hasChart ? patientFirst.size : null;
  const newPatients: number | null = hasChart && m.patientType ? newSet.size : null;

  const group = (key: FieldKey, onlyNew: boolean) => {
    if (!m[key] || !hasChart) return null;
    const out = new Map<string, { patients: number; amount: number }>();
    for (const [chart, r] of patientFirst) {
      if (onlyNew && m.patientType && !newSet.has(chart)) continue;
      const k = str(g(r, key)) || "미기재";
      const cur = out.get(k) ?? { patients: 0, amount: 0 };
      cur.patients++;
      cur.amount += patientAmount.get(chart) ?? 0;
      out.set(k, cur);
    }
    return out;
  };

  let referral: Stats["referral"] = null;
  if (hasChart && (m.source || m.referrer)) {
    const referrers = new Map<string, number>();
    let patients = 0;
    let amount = 0;
    for (const [chart, r] of patientFirst) {
      if (m.patientType && !newSet.has(chart)) continue;
      const viaSource = m.source ? /소개|지인/.test(str(g(r, "source"))) : false;
      const who = m.referrer ? str(g(r, "referrer")) : "";
      if (viaSource || who) {
        patients++;
        amount += patientAmount.get(chart) ?? 0;
        if (who) referrers.set(who, (referrers.get(who) ?? 0) + 1);
      }
    }
    referral = { patients, amount, referrers };
  }

  return {
    rowCount: sheet.rows.length,
    total: m.amount ? total : null,
    insured: m.insurance ? insured : null,
    nonInsured: m.insurance ? nonInsured : null,
    visitors,
    newPatients,
    returning: visitors !== null && newPatients !== null ? visitors - newPatients : null,
    days: m.date ? dates.size : null,
    perPatient: m.amount && visitors ? total / visitors : null,
    treatment: m.treatment && m.amount ? treatment : null,
    appt: m.apptStatus ? appt : null,
    sources: group("source", true),
    regions: group("region", false),
    referral,
  };
}

// ---------- 보고서 표 ----------

export type Cell = string | number | null;
export type Table = { id: string; title: string; note?: string; headers: string[]; rows: Cell[][] };

const won = (v: number | null): Cell => (v === null ? null : Math.round(v));
const diff = (a: number | null, b: number | null): Cell => (a === null || b === null ? null : a - b);
const pct = (a: number | null, b: number | null): Cell =>
  a === null || b === null || b === 0 ? null : Math.round(((a - b) / b) * 1000) / 10;
const ratio = (a: number, b: number): number | null => (b === 0 ? null : Math.round((a / b) * 1000) / 10);

function compareRows(label: string, cur: number | null, prev: number | null, isWon = true): Cell[] {
  const c = isWon ? won(cur) : cur;
  const p = isWon ? won(prev) : prev;
  const d = diff(cur, prev) as number | null;
  return [label, c, p, d === null ? null : isWon ? Math.round(d) : Math.round(d * 10) / 10, pct(cur, prev)];
}

export function buildTables(cur: Stats, prev: Stats | null): Table[] {
  const P = prev;
  const out: Table[] = [];
  const cmpHead = ["구분", "당월", "전월", "증감", "증감률(%)"];
  const p = <K extends keyof Stats>(k: K) => (P ? (P[k] as number | null) : null);

  out.push({
    id: "t31",
    title: "3-1. 총 진료비 당월·전월 비교 (원)",
    headers: cmpHead,
    rows: [
      compareRows("보험 진료비", cur.insured, p("insured")),
      compareRows("비보험 진료비", cur.nonInsured, p("nonInsured")),
      compareRows("총 진료비", cur.total, p("total")),
    ],
  });

  out.push({
    id: "t32",
    title: "3-2. 내원 환자 수 당월·전월 비교",
    headers: cmpHead,
    rows: [
      compareRows("총 내원 환자 수(명)", cur.visitors, p("visitors"), false),
      compareRows("신환(명)", cur.newPatients, p("newPatients"), false),
      compareRows("구환(명)", cur.returning, p("returning"), false),
      compareRows("영업일수(일)", cur.days, p("days"), false),
      compareRows(
        "일평균 내원 환자 수(명)",
        cur.visitors && cur.days ? Math.round((cur.visitors / cur.days) * 10) / 10 : null,
        P && P.visitors && P.days ? Math.round((P.visitors / P.days) * 10) / 10 : null,
        false,
      ),
      compareRows("환자 1인당 평균 진료비(원)", cur.perPatient, p("perPatient")),
    ],
  });

  if (cur.treatment) {
    const cats = new Set([...cur.treatment.keys(), ...(P?.treatment?.keys() ?? [])]);
    const totalC = [...cur.treatment.values()].reduce((a, b) => a + b, 0);
    const totalP = P?.treatment ? [...P.treatment.values()].reduce((a, b) => a + b, 0) : 0;
    const rows = [...cats]
      .map((c) => {
        const a = cur.treatment!.get(c) ?? 0;
        const b = P?.treatment ? (P.treatment.get(c) ?? 0) : null;
        const sa = ratio(a, totalC);
        const sb = b === null ? null : ratio(b, totalP);
        return { c, a, b, sa, sb };
      })
      .sort((x, y) => y.a - x.a)
      .map(({ c, a, b, sa, sb }): Cell[] => [
        c,
        Math.round(a),
        sa,
        b === null ? null : Math.round(b),
        sb,
        sa !== null && sb !== null ? Math.round((sa - sb) * 10) / 10 : null,
      ]);
    out.push({
      id: "t33",
      title: "3-3. 주요 치료별 매출 비중 비교",
      note: "치료 분류는 화면의 '치료 분류 규칙'에 따라 나뉩니다.",
      headers: ["치료 항목", "당월 매출(원)", "당월 비중(%)", "전월 매출(원)", "전월 비중(%)", "비중 변화(%p)"],
      rows,
    });
  }

  if (cur.appt) {
    const rate = (s: Stats["appt"], k: "done" | "cancel" | "noshow") =>
      s ? ratio(s[k], s.total) : null;
    const pp = (a: number | null, b: number | null): Cell =>
      a === null || b === null ? null : Math.round((a - b) * 10) / 10;
    out.push({
      id: "t34",
      title: "3-4. 예약 이행률 분석",
      note: "이행률 = 내원 건수 ÷ 전체 예약 건수 × 100",
      headers: ["구분", "당월", "전월", "증감(%p 또는 건)"],
      rows: [
        ["전체 예약 건수", cur.appt.total, P?.appt?.total ?? null, diff(cur.appt.total, P?.appt?.total ?? null)],
        ["예약 이행률(%)", rate(cur.appt, "done"), rate(P?.appt ?? null, "done"), pp(rate(cur.appt, "done"), rate(P?.appt ?? null, "done"))],
        ["사전 취소(%)", rate(cur.appt, "cancel"), rate(P?.appt ?? null, "cancel"), pp(rate(cur.appt, "cancel"), rate(P?.appt ?? null, "cancel"))],
        ["노쇼(%)", rate(cur.appt, "noshow"), rate(P?.appt ?? null, "noshow"), pp(rate(cur.appt, "noshow"), rate(P?.appt ?? null, "noshow"))],
      ],
    });
  }

  const groupTable = (
    id: string,
    title: string,
    label: string,
    a: Stats["sources"],
    b: Stats["sources"] | null,
    note?: string,
  ) => {
    if (!a) return;
    const totalA = [...a.values()].reduce((s, v) => s + v.patients, 0);
    const rows = [...a.entries()]
      .sort((x, y) => y[1].patients - x[1].patients)
      .map(([k, v]): Cell[] => {
        const prevN = b ? (b.get(k)?.patients ?? 0) : null;
        return [
          k,
          v.patients,
          ratio(v.patients, totalA),
          prevN,
          prevN === null ? null : v.patients - prevN,
          v.patients ? Math.round(v.amount / v.patients) : null,
        ];
      });
    out.push({
      id,
      title,
      note,
      headers: [label, "당월 환자(명)", "비중(%)", "전월 환자(명)", "증감(명)", "환자 1인당 매출(원)"],
      rows,
    });
  };
  groupTable("t35", "3-5. 내원 경로별 신환", "내원 경로", cur.sources, P?.sources ?? null, "신환 기준으로 집계합니다.");
  groupTable("t35b", "3-5. 거주 지역별 내원 환자", "지역", cur.regions, P?.regions ?? null);

  if (cur.referral) {
    const r = cur.referral;
    const pr = P?.referral ?? null;
    const perOf = (x: { patients: number; amount: number } | null) =>
      x && x.patients ? x.amount / x.patients : null;
    out.push({
      id: "t36",
      title: "3-6. 소개환자 분석",
      headers: cmpHead,
      rows: [
        compareRows("소개환자 수(명)", r.patients, pr ? pr.patients : null, false),
        [
          "신환 대비 소개 비율(%)",
          cur.newPatients ? ratio(r.patients, cur.newPatients) : null,
          pr && P?.newPatients ? ratio(pr.patients, P.newPatients) : null,
          null,
          null,
        ],
        compareRows("소개환자 총 매출(원)", r.amount, pr ? pr.amount : null),
        compareRows("소개환자 1인당 매출(원)", perOf(r), perOf(pr)),
      ],
    });
    if (r.referrers.size) {
      out.push({
        id: "t36b",
        title: "3-6. 소개를 많이 해 준 환자 (당월)",
        note: "감사 인사·혜택 안내 대상 확인용입니다. 환자 식별정보는 회의 자료 배포 전에 가려 주세요.",
        headers: ["소개자", "소개 인원(명)"],
        rows: [...r.referrers.entries()].sort((a, b) => b[1] - a[1]).slice(0, 10).map(([k, v]) => [k, v]),
      });
    }
  }

  // 3-7 매출 증감 분해
  if (cur.total !== null && prev && prev.total !== null && cur.visitors && prev.visitors) {
    const v0 = prev.visitors;
    const v1 = cur.visitors;
    const a0 = prev.total / v0;
    const a1 = cur.total / v1;
    const volume = (v1 - v0) * a0;
    const price = (a1 - a0) * v1;
    out.push({
      id: "t37",
      title: "3-7. 매출 증감 요인 분해 (원)",
      note: "매출 = 내원 환자 수 × 1인당 진료비. 환자 수 효과 = (당월−전월 환자 수) × 전월 1인당 진료비, 단가 효과 = (당월−전월 1인당 진료비) × 당월 환자 수.",
      headers: ["요인", "매출 증감(원)", "총 증감 대비(%)"],
      rows: [
        ["내원 환자 수 변화", Math.round(volume), ratio(volume, cur.total - prev.total)],
        ["환자 1인당 진료비 변화", Math.round(price), ratio(price, cur.total - prev.total)],
        ["합계", Math.round(cur.total - prev.total), 100],
      ],
    });
  }

  return out;
}

export type Growth = { label: string; unit: string; def: number; estimate: (c: Stats, x: number) => number | null };

/** 개선 시뮬레이션: 입력값(%, %p)만큼 좋아졌을 때 월 매출 증가 추정치 */
export const GROWTH: Growth[] = [
  {
    label: "신환이 늘어난다면",
    unit: "%",
    def: 10,
    estimate: (c, x) =>
      c.newPatients && c.total && c.visitors ? Math.round(c.newPatients * (x / 100) * (c.total / c.visitors)) : null,
  },
  {
    label: "노쇼·취소율이 줄어든다면",
    unit: "%p",
    def: 3,
    estimate: (c, x) =>
      c.appt && c.appt.done && c.total ? Math.round(c.appt.total * (x / 100) * (c.total / c.appt.done)) : null,
  },
  {
    label: "환자 1인당 진료비가 오른다면",
    unit: "%",
    def: 5,
    estimate: (c, x) => (c.total ? Math.round(c.total * (x / 100)) : null),
  },
];
