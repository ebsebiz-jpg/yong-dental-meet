"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import {
  deleteMonth,
  purgeFromTrash,
  restoreFromTrash,
  saveMonth,
  saveReport,
  saveRules,
  type Loaded,
} from "@/app/data";
import { savePptx, type SlideItem } from "@/lib/pptx";
import { seriesSummary, summarize } from "@/lib/summary";
import { deserializeMonth, emptyReport, normalizeReport, serializeMonth, type Report } from "@/lib/store";
import {
  DEFAULT_RULES,
  GROWTH,
  KINDS,
  buildTables,
  expandFiles,
  ingest,
  metrics,
  parseRules,
  prevYm,
  ymLabel,
  type Cell,
  type Month,
  type Table,
} from "@/lib/dentweb";

const fmt = (c: Cell) =>
  c === null || c === "" ? "–" : typeof c === "number" ? c.toLocaleString("ko-KR") : c;

const box = "rounded-xl border border-slate-200 bg-white p-5 print:border-0 print:p-0";
const input =
  "w-full rounded-md border border-slate-300 px-2 py-1.5 text-sm outline-none focus:border-slate-900";

function DataTable({
  t,
  summary,
  comment,
  onComment,
}: {
  t: Table;
  summary: string | null;
  comment: string;
  onComment: (v: string) => void;
}) {
  return (
    <section className="break-inside-avoid">
      <h3 className="text-base font-semibold text-slate-900">{t.title}</h3>
      {summary && <p className="mt-1 text-sm font-medium text-slate-800">{summary}</p>}
      {t.note && <p className="mt-1 text-xs text-slate-500">{t.note}</p>}
      <div className="mt-2 overflow-x-auto">
        <table className="w-full border-collapse text-sm">
          <thead>
            <tr className="bg-slate-100">
              {t.headers.map((h) => (
                <th key={h} className="border border-slate-200 px-2 py-1.5 text-left font-medium whitespace-nowrap">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {t.rows.map((r, i) => (
              <tr key={i}>
                {r.map((c, j) => (
                  <td
                    key={j}
                    className={`border border-slate-200 px-2 py-1.5 whitespace-pre ${j > 0 ? "text-right tabular-nums" : ""}`}
                  >
                    {fmt(c)}
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <textarea
        value={comment}
        onChange={(e) => onComment(e.target.value)}
        placeholder="분석 코멘트 (증감 원인, 특이사항)"
        rows={2}
        className={`${input} mt-2`}
      />
    </section>
  );
}

const EXAMPLE_META = {
  title: "2026년 10월 월간 경영회의",
  date: "2026. 10. 05 (월) 19:00 ~ 20:30",
  place: "상담실",
  attendees: "원장, 부원장, 실장, 치과위생사 팀장, 상담 담당",
};

// 문장 속 ○○, ○는 직접 채워 넣는 자리입니다. 숫자는 예시일 뿐 실제 자료와 무관합니다.
const EXAMPLE_COMMENTS: Record<string, string> = {
  actions: [
    "- 상담 확정률 개선: 담당 ○○, 진행 중 (○월 ○일까지)",
    "- 신환 유입 채널 점검: 담당 ○○, 완료",
    "- 기공 재제작 원인 정리: 담당 ○○, 지연 (업체 일정 조율 중)",
  ].join("\n"),
  t31: "총 진료비가 전월 대비 ○% 줄었습니다. 비급여 진료 감소가 주된 원인이고, 영업일수는 전월과 같습니다.",
  t32: "신환은 늘었지만 구환이 줄어 총 내원 환자는 소폭 감소했습니다. 환자 1인당 진료비도 줄었습니다.",
  t33: "보철 비중이 ○%p 줄고 발치·수술 비중이 늘었습니다. 전월 임플란트 식립 환자의 보철 진행 시기를 확인합니다.",
  t34: "소개 환자 비중이 가장 높고 인터넷 유입 신환이 늘었습니다. 미입력 경로가 ○%라 접수 시 경로 기록을 철저히 합니다.",
  t34b: "신환의 대부분이 ○○구 거주입니다. ○○구 외 지역 유입은 ○명으로 미미합니다.",
  t35: "소개 신환 비중이 전월보다 ○%p 낮아졌습니다. 소개해 주신 환자께 감사 인사를 보내고 소개 프로그램을 점검합니다.",
  t35b: "소개를 많이 해 준 환자 상위 분께 감사 문자와 혜택 안내를 보냅니다.",
  t36: "상담 확정 비율이 전월보다 ○%p 하락했습니다. 계획 수는 늘었지만 확정이 따라오지 못했습니다.",
  t36b: "○○ 상담자의 확정 비율이 크게 낮아졌습니다. 미확정 환자 후속 상담과 상담 방식 공유가 필요합니다.",
  t37: "임플란트 의뢰가 줄고 틀니 의뢰가 늘었습니다. Cr./Br. 재제작률이 높아 원인을 업체와 확인합니다.",
  t38: "최근 3개월 총 진료비가 계속 감소했습니다. 내원 환자보다 환자 1인당 진료비 하락 폭이 큽니다.",
  t39: "총 진료비 감소의 대부분은 환자 1인당 진료비 하락 때문입니다. 고단가 치료 상담과 확정에 집중합니다.",
  staff: "- 치과위생사 1명 입사(○월 ○일), 조무사 1명 퇴사 예정\n- 보수교육 이수 현황 점검\n- 연차·근무시간 조정 건 논의",
  ops: "- 유닛체어 ○번 유압 점검 필요\n- 멸균 기록 누락 없이 관리 중\n- 임플란트 재료 재고 부족 예상, 발주 예정\n- 기공소 납기 지연 1건, 업체와 협의",
  risk: "- 의료분쟁: 임플란트 환자 불만 1건, 경위서 작성 및 사후 관리 중\n- 보험 심사: 삭감 예상 청구 건 사전 점검\n- 인력 이탈: 핵심 직원 면담 필요",
  decide: "1. 신규 장비 도입 여부 → 견적 비교 후 다음 회의에서 결정\n2. 비보험 상담 프로세스 개편 → 상담실장 주도로 시범 운영\n3. 마케팅 예산 조정 → 효율 낮은 채널 축소",
  next: "- 총 진료비 ○원, 신환 ○명, 상담 확정 비율 ○% 목표\n- [ ] 미확정 환자 후속 상담: 담당 ○○, ○월 ○일까지\n- [ ] 소개환자 감사 프로그램 시행: 담당 ○○, ○월 ○일까지\n- 다음 회의: ○월 ○일 ○요일 19:00",
  focus: "1. 상담 확정 비율 회복(미확정 환자 후속 상담)\n2. 환자 1인당 진료비 회복(고단가 치료 상담 강화)\n3. 신환 유입 유지(인터넷·소개 채널 관리)",
};

const MANUAL = [
  ["actions", "전월 액션아이템 점검"],
  ["staff", "인력 현황 (입·퇴사, 교육, 근무 관련 안건)"],
  ["ops", "진료·운영 현황 (장비, 감염관리, 재료 재고, 기공소)"],
  ["risk", "리스크 및 이슈"],
  ["decide", "논의 안건 및 의사결정"],
  ["next", "차월 목표 및 액션아이템"],
] as const;

type SaveState = "idle" | "saving" | "saved" | "error";
const nowLabel = () => new Date().toLocaleTimeString("ko-KR", { hour: "2-digit", minute: "2-digit" });

export default function MeetingApp({ logout, initial }: { logout: () => Promise<void>; initial: Loaded }) {
  const [months, setMonths] = useState<Map<string, Month>>(
    () => new Map(Object.entries(initial.months).map(([k, v]) => [k, deserializeMonth(v)])),
  );
  const [reports, setReports] = useState<Record<string, Report>>(() =>
    Object.fromEntries(Object.entries(initial.reports).map(([k, v]) => [k, normalizeReport(v)])),
  );
  const [unknown, setUnknown] = useState<string[]>([]);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [curYm, setCurYm] = useState("");
  const [prevSel, setPrevSel] = useState<string | null>(null);
  const [rulesText, setRulesText] = useState(initial.rules ?? DEFAULT_RULES);
  const [pptBusy, setPptBusy] = useState(false);
  const [delBusy, setDelBusy] = useState(false);
  const [undo, setUndo] = useState<{ ym: string; month?: Month; report?: Report | null } | null>(null);
  const [trash, setTrash] = useState(initial.trash);
  const [saveState, setSaveState] = useState<SaveState>("idle");
  const [savedAt, setSavedAt] = useState("");
  const dirtyReports = useRef(new Set<string>());
  const rulesDirty = useRef(false);
  const reportsRef = useRef(reports);
  const rulesRef = useRef(rulesText);

  const rules = useMemo(() => parseRules(rulesText), [rulesText]);
  const keys = useMemo(() => [...months.keys()].sort(), [months]);
  const cur = keys.includes(curYm) ? curYm : (keys[keys.length - 1] ?? "");
  const defPrev = keys.includes(prevYm(cur)) ? prevYm(cur) : (keys.filter((k) => k < cur).pop() ?? null);
  const prev = prevSel !== null && (prevSel === "" || keys.includes(prevSel)) ? (prevSel || null) : defPrev;

  // 회의 개요·코멘트·시뮬레이션 입력은 기준월별로 따로 보관한다.
  const blank = useMemo(() => emptyReport(), []);
  const report = reports[cur] ?? blank;
  const meta = report.meta;
  const comments = report.comments;
  const growthIn = report.growthIn;
  const patch = (fn: (r: Report) => Report) => {
    if (!cur) return;
    dirtyReports.current.add(cur);
    setReports((p) => ({ ...p, [cur]: fn(p[cur] ?? emptyReport()) }));
  };
  const setMeta = (v: Report["meta"] | ((m: Report["meta"]) => Report["meta"])) =>
    patch((r) => ({ ...r, meta: typeof v === "function" ? v(r.meta) : v }));
  const setComments = (v: Record<string, string> | ((c: Record<string, string>) => Record<string, string>)) =>
    patch((r) => ({ ...r, comments: typeof v === "function" ? v(r.comments) : v }));
  const setGrowthIn = (v: number[]) => patch((r) => ({ ...r, growthIn: v }));

  useEffect(() => {
    reportsRef.current = reports;
  }, [reports]);
  useEffect(() => {
    rulesRef.current = rulesText;
  }, [rulesText]);

  // 입력을 멈추고 1초 뒤에 자동 저장한다.
  useEffect(() => {
    if (!initial.dbReady) return;
    if (!dirtyReports.current.size && !rulesDirty.current) return;
    const t = setTimeout(async () => {
      const yms = [...dirtyReports.current];
      const withRules = rulesDirty.current;
      dirtyReports.current.clear();
      rulesDirty.current = false;
      setSaveState("saving");
      const results = await Promise.all([
        ...yms.map((ym) => saveReport(ym, reportsRef.current[ym])),
        ...(withRules ? [saveRules(rulesRef.current)] : []),
      ]);
      const ok = results.every((r) => r.ok);
      if (!ok) {
        yms.forEach((ym) => dirtyReports.current.add(ym));
        rulesDirty.current = withRules;
      } else {
        setSavedAt(nowLabel());
      }
      setSaveState(ok ? "saved" : "error");
    }, 1000);
    return () => clearTimeout(t);
  }, [reports, rulesText, initial.dbReady]);

  const tables = useMemo(() => (cur ? buildTables(months, cur, prev, rules) : []), [months, cur, prev, rules]);
  const mc = useMemo(() => metrics(months.get(cur)), [months, cur]);

  const onFiles = async (list: FileList | null) => {
    if (!list || !list.length) return;
    setBusy(true);
    setErr("");
    try {
      const files = await expandFiles([...list]);
      if (!files.length) {
        setErr("엑셀(.xlsx, .xls) 또는 zip 파일을 올려 주세요.");
        return;
      }
      const res = ingest(files, months);
      setMonths(res.months);
      setUnknown(res.unknown);
      if (initial.dbReady) {
        // 새로 올린 달은 저장하고, 이미 있던 달은 같은 리포트를 덮어써 최신 자료로 갱신한다.
        setSaveState("saving");
        const touched = [...res.months].filter(([k, m]) => months.get(k) !== m);
        const results = await Promise.all(touched.map(([k, m]) => saveMonth(k, serializeMonth(m))));
        const ok = results.every((r) => r.ok);
        if (ok) setSavedAt(nowLabel());
        setSaveState(ok ? "saved" : "error");
      }
    } catch {
      setErr("파일을 읽지 못했습니다. 덴트웹에서 내보낸 원본 파일인지 확인해 주세요.");
    } finally {
      setBusy(false);
    }
  };


  // 이미 입력한 칸은 건드리지 않고 빈 칸만 예시로 채운다.
  const fillExample = () => {
    setMeta((m) => ({
      title: m.title || EXAMPLE_META.title,
      date: m.date || EXAMPLE_META.date,
      place: m.place || EXAMPLE_META.place,
      attendees: m.attendees || EXAMPLE_META.attendees,
    }));
    setComments((c) => {
      const next = { ...c };
      for (const [k, v] of Object.entries(EXAMPLE_COMMENTS)) if (!next[k]) next[k] = v;
      return next;
    });
  };

  // 삭제하면 서버의 휴지통으로 옮겨 trashDays일 동안 보관한다(새로고침해도 복구 가능).
  // 저장소가 없는 환경에서는 이 화면을 닫기 전까지만 되돌릴 수 있다.
  const removeMonth = async (ym: string) => {
    const month = months.get(ym);
    if (!month || delBusy) return;
    setDelBusy(true);
    setErr("");
    try {
      if (initial.dbReady) {
        const r = await deleteMonth(ym);
        if (!r.ok) {
          setErr(r.error ?? "삭제에 실패했습니다.");
          return;
        }
        setTrash((p) => [{ ym, deletedAt: new Date().toISOString() }, ...p.filter((x) => x.ym !== ym)]);
        setUndo({ ym });
      } else {
        setUndo({ ym, month, report: reports[ym] ?? null });
      }
      setMonths((p) => new Map([...p].filter(([k]) => k !== ym)));
      setReports((p) => Object.fromEntries(Object.entries(p).filter(([k]) => k !== ym)));
      dirtyReports.current.delete(ym);
    } finally {
      setDelBusy(false);
    }
  };

  /** 휴지통(또는 방금 지운 로컬 백업)에서 한 달을 되살린다. */
  const restoreMonth = async (ym: string) => {
    if (delBusy) return;
    setDelBusy(true);
    setErr("");
    try {
      if (initial.dbReady) {
        const r = await restoreFromTrash(ym);
        if (!r.ok || !r.month) {
          setErr(r.error ?? "복구에 실패했습니다.");
          return;
        }
        setMonths((p) => new Map(p).set(ym, deserializeMonth(r.month)));
        if (r.report) setReports((p) => ({ ...p, [ym]: normalizeReport(r.report) }));
        setTrash((p) => p.filter((x) => x.ym !== ym));
      } else if (undo?.ym === ym && undo.month) {
        const { month, report: rep } = undo;
        setMonths((p) => new Map(p).set(ym, month));
        if (rep) setReports((p) => ({ ...p, [ym]: rep }));
      }
      setUndo((u) => (u?.ym === ym ? null : u));
    } finally {
      setDelBusy(false);
    }
  };

  const purgeMonth = async (ym: string) => {
    if (delBusy) return;
    if (!window.confirm(`${ymLabel(ym)} 자료를 휴지통에서 영구 삭제할까요? 이 작업은 되돌릴 수 없습니다.`)) return;
    setDelBusy(true);
    setErr("");
    try {
      const r = await purgeFromTrash(ym);
      if (!r.ok) {
        setErr(r.error ?? "영구 삭제에 실패했습니다.");
        return;
      }
      setTrash((p) => p.filter((x) => x.ym !== ym));
      setUndo((u) => (u?.ym === ym ? null : u));
    } finally {
      setDelBusy(false);
    }
  };

  const exportPptx = async () => {
    setPptBusy(true);
    setErr("");
    try {
      const blob = await (await fetch("/logo.png")).blob();
      const logo = await new Promise<string>((resolve, reject) => {
        const fr = new FileReader();
        fr.onload = () => resolve(String(fr.result));
        fr.onerror = () => reject(fr.error);
        fr.readAsDataURL(blob);
      });
      const items: SlideItem[] = [];
      if (comments.actions?.trim()) items.push({ kind: "text", title: "전월 액션아이템 점검", body: comments.actions });
      for (const t of tables) {
        items.push({ kind: "table", table: t, summary: summarize(t), comment: comments[t.id] });
        if (t.id === "t38") {
          // 월별 추이 표 바로 뒤에 추이 그래프 두 장을 넣는다(최근 6개월, 막대 0에서 시작).
          const ks = keys.filter((k) => k <= cur).slice(-6);
          const ms = ks.map((k) => metrics(months.get(k)));
          const labels = ks.map(ymLabel);
          const vals = (f: (m: ReturnType<typeof metrics>) => number | null) => ms.map((m) => f(m) ?? 0);
          if (ks.length >= 2) {
            const total = vals((m) => m.total);
            const patients = vals((m) => m.patients);
            const fresh = vals((m) => m.fresh);
            items.push({
              kind: "charts",
              title: "월별 추이 ① 총 진료비·환자 1인당 진료비",
              summary: seriesSummary("총 진료비", labels, total, "원"),
              charts: [
                { title: "총 진료비(백만원)", labels, values: total.map((v) => Math.round(v / 1e5) / 10), format: "#,##0.0" },
                { title: "환자 1인당 진료비(원)", labels, values: vals((m) => m.perPatient).map(Math.round), format: "#,##0" },
              ],
            });
            const last = ks.length - 1;
            items.push({
              kind: "charts",
              title: "월별 추이 ② 내원 환자·신환",
              summary: `내원 환자는 ${labels[0]} ${patients[0].toLocaleString("ko-KR")}명에서 ${labels[last]} ${patients[last].toLocaleString("ko-KR")}명으로 ${patients[last] > patients[0] ? "늘" : patients[last] < patients[0] ? "줄" : "같"}${patients[last] === patients[0] ? "습니다" : "었고"}, 신환은 ${fresh[0].toLocaleString("ko-KR")}명에서 ${fresh[last].toLocaleString("ko-KR")}명으로 ${fresh[last] > fresh[0] ? "늘었습니다" : fresh[last] < fresh[0] ? "줄었습니다" : "같습니다"}.`,
              charts: [
                { title: "내원 환자(명)", labels, values: patients, format: "#,##0" },
                { title: "신환(명)", labels, values: fresh, format: "#,##0" },
              ],
            });
          }
        }
      }
      items.push({
        kind: "table",
        table: {
          id: "growth",
          title: "4. 진료비 성장 시뮬레이션",
          note: "당월 실적을 기준으로, 항목이 개선되면 월 총 진료비가 얼마나 늘 수 있는지 단순 추정한 값입니다.",
          headers: ["개선 항목", "개선폭", "월 진료비 증가 추정(원)"],
          rows: GROWTH.map((g, i) => [g.label, `${growthIn[i]}${g.unit}`, g.estimate(mc, growthIn[i])]),
        },
        comment: comments.focus?.trim() ? comments.focus : undefined,
        commentLabel: "이번 달 집중 과제",
      });
      for (const [k, label] of MANUAL.slice(1)) {
        if (comments[k]?.trim()) items.push({ kind: "text", title: label, body: comments[k] });
      }
      await savePptx({
        logo,
        title: meta.title || `${ymLabel(cur)} 월간 경영회의`,
        lines: [
          meta.date && `일시  ${meta.date}`,
          meta.place && `장소  ${meta.place}`,
          meta.attendees && `참석자  ${meta.attendees}`,
          `기준 ${ymLabel(cur)}${prev ? ` · 비교 ${ymLabel(prev)}` : ""}`,
        ].filter((x): x is string => !!x),
        items,
        fileName: `월간경영회의_${cur}.pptx`,
      });
    } catch {
      setErr("PPT를 만들지 못했습니다. 잠시 뒤 다시 시도해 주세요.");
    } finally {
      setPptBusy(false);
    }
  };

  const exportXlsx = () => {
    const wb = XLSX.utils.book_new();
    const overview = [
      ["회의명", meta.title],
      ["일시", meta.date],
      ["장소", meta.place],
      ["참석자", meta.attendees],
      ["기준월", cur ? ymLabel(cur) : ""],
      ["비교월", prev ? ymLabel(prev) : ""],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(overview), "회의 개요");
    for (const t of tables) {
      const aoa: Cell[][] = [t.headers, ...t.rows];
      const c = comments[t.id];
      if (c) aoa.push([], ["분석 코멘트", c]);
      const name = t.title.replace(/[\\/?*[\]:]/g, "").slice(0, 31);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name);
    }
    const manual = [...MANUAL.map(([k, label]) => [label, comments[k] ?? ""]), ["이번 달 집중 과제", comments.focus ?? ""]];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(manual), "직접 입력 항목");
    XLSX.writeFile(wb, `월간경영회의_${cur || "자료"}.xlsx`);
  };

  return (
    <div className="mx-auto max-w-5xl space-y-6 px-4 py-8">
      <header className="flex items-center justify-between print:hidden">
        <h1 className="text-xl font-bold text-slate-900">월간 경영회의 자료 만들기</h1>
        <form action={logout}>
          <button className="text-sm text-slate-500 underline">로그아웃</button>
        </form>
      </header>

      <div className="rounded-lg bg-amber-50 p-3 text-xs text-amber-900 print:hidden">
        엑셀 파일 자체는 서버로 보내지 않고 이 브라우저에서 읽어 집계합니다. 집계 결과(소개자·소개받은 신환의 성명, 신환
        주소의 구 단위 포함)와 입력한 내용은 서버 데이터베이스에 저장되어 어느 PC에서든 이어서 볼 수 있습니다. 연락처와
        생년월일 등은 읽지 않습니다. 인쇄·공유 전에 식별정보가 없는지 확인해 주세요.
        {!initial.dbReady && (
          <span className="mt-1 block font-medium text-red-700">
            저장소(DATABASE_URL)가 연결되지 않아 지금은 저장되지 않습니다. 새로고침하면 입력한 내용이 사라집니다.
            {initial.dbProblem && <span className="mt-1 block">원인: {initial.dbProblem}</span>}
          </span>
        )}
      </div>

      <div className={`${box} print:hidden`}>
        <h2 className="font-semibold text-slate-900">1. 덴트웹 자료 올리기</h2>
        <p className="mt-1 text-xs text-slate-500">
          월별 리포트 5종(기간별 진료비, 내원경로 분포, 상담자별 상담, 수입 통계, 기공 의뢰)을 zip 그대로 또는 엑셀 파일로
          올립니다. 여러 달을 한꺼번에 올려도 됩니다. 이미 올린 달의 같은 리포트를 다시 올리면 최신 자료로 바뀌고, 한 번 올린
          달은 다시 올리지 않아도 저장되어 있습니다.
        </p>
        <input
          type="file"
          multiple
          accept=".xlsx,.xls,.zip"
          disabled={busy}
          onChange={(e) => {
            onFiles(e.target.files);
            e.target.value = "";
          }}
          className="mt-3 block w-full rounded-lg border border-dashed border-slate-300 p-4 text-sm"
        />
        {busy && <p className="mt-2 text-xs text-slate-500">파일을 읽는 중…</p>}
        {err && <p className="mt-2 text-xs text-red-600">{err}</p>}
        {unknown.length > 0 && (
          <p className="mt-2 text-xs text-red-600">인식하지 못한 파일: {unknown.join(", ")}</p>
        )}
        {undo && (
          <div className="mt-3 flex flex-wrap items-center gap-3 rounded-lg border border-amber-200 bg-amber-50 px-3 py-2 text-sm text-amber-900">
            <span>
              {initial.dbReady
                ? `${ymLabel(undo.ym)} 자료를 휴지통으로 옮겼습니다. 아래 휴지통에서 ${initial.trashDays}일 안에 복구할 수 있습니다.`
                : `${ymLabel(undo.ym)} 자료를 삭제했습니다. 저장소가 없어 새로고침하면 되돌릴 수 없으니 필요하면 지금 되돌리세요.`}
            </span>
            <button
              onClick={() => restoreMonth(undo.ym)}
              disabled={delBusy}
              className="rounded bg-slate-900 px-3 py-1 text-xs font-medium text-white disabled:opacity-50"
            >
              되돌리기
            </button>
            <button onClick={() => setUndo(null)} className="text-xs text-amber-800 underline">
              닫기
            </button>
          </div>
        )}
        {keys.length > 0 && (
          <div className="mt-4 overflow-x-auto">
            <table className="w-full border-collapse text-sm">
              <thead>
                <tr className="bg-slate-100">
                  <th className="border border-slate-200 px-2 py-1.5 text-left font-medium">월</th>
                  {KINDS.map((k) => (
                    <th key={k.kind} className="border border-slate-200 px-2 py-1.5 text-center font-medium">
                      {k.label}
                    </th>
                  ))}
                  <th className="border border-slate-200 px-2 py-1.5 text-center font-medium">삭제</th>
                </tr>
              </thead>
              <tbody>
                {keys.map((k) => (
                  <tr key={k}>
                    <td className="border border-slate-200 px-2 py-1.5">{ymLabel(k)}</td>
                    {KINDS.map((x) => (
                      <td key={x.kind} className="border border-slate-200 px-2 py-1.5 text-center">
                        {months.get(k)?.[x.kind] ? <span className="text-emerald-700">✓</span> : <span className="text-slate-300">–</span>}
                      </td>
                    ))}
                    <td className="border border-slate-200 px-2 py-1.5 text-center">
                      <button
                        onClick={() => removeMonth(k)}
                        disabled={delBusy}
                        title={`${ymLabel(k)} 자료와 입력 내용을 삭제합니다. 삭제 직후에는 되돌릴 수 있습니다.`}
                        className="rounded border border-red-300 px-2 py-0.5 text-xs font-medium text-red-600 hover:bg-red-50 disabled:cursor-not-allowed disabled:opacity-50"
                      >
                        삭제
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <p className="mt-1 text-xs text-slate-500">빠진 리포트가 있으면 그 항목의 표는 만들어지지 않습니다.</p>
          </div>
        )}
        {initial.dbReady && trash.length > 0 && (
          <details className="mt-4 rounded-lg border border-slate-200 p-3" open={!!undo}>
            <summary className="cursor-pointer text-sm font-medium text-slate-800">
              휴지통 ({trash.length}) — 삭제 후 {initial.trashDays}일이 지나면 자동으로 영구 삭제됩니다
            </summary>
            <p className="mt-2 text-xs text-slate-500">
              지운 달의 집계 자료와 입력한 코멘트가 그대로 보관되어 있습니다. 소개자·신환 이름이 들어 있으니 필요 없으면 영구 삭제하세요.
            </p>
            <ul className="mt-2 divide-y divide-slate-100 text-sm">
              {trash.map((x) => (
                <li key={x.ym} className="flex flex-wrap items-center justify-between gap-2 py-2">
                  <span>
                    {ymLabel(x.ym)}
                    <span className="ml-2 text-xs text-slate-500">
                      삭제 {new Date(x.deletedAt).toLocaleString("ko-KR", { dateStyle: "medium", timeStyle: "short" })}
                    </span>
                  </span>
                  <span className="flex gap-2">
                    <button
                      onClick={() => restoreMonth(x.ym)}
                      disabled={delBusy}
                      className="rounded bg-slate-900 px-3 py-1 text-xs font-medium text-white disabled:opacity-50"
                    >
                      복구
                    </button>
                    <button
                      onClick={() => purgeMonth(x.ym)}
                      disabled={delBusy}
                      className="rounded border border-red-300 px-3 py-1 text-xs font-medium text-red-600 hover:bg-red-50 disabled:opacity-50"
                    >
                      영구 삭제
                    </button>
                  </span>
                </li>
              ))}
            </ul>
          </details>
        )}
      </div>

      {keys.length > 0 && (
        <div className={`${box} print:hidden`}>
          <h2 className="font-semibold text-slate-900">2. 비교할 달 고르기</h2>
          <div className="mt-3 grid gap-3 sm:grid-cols-2">
            <label className="text-sm">
              <span className="text-slate-700">당월(기준)</span>
              <select className={`${input} mt-1`} value={cur} onChange={(e) => { setCurYm(e.target.value); setPrevSel(null); }}>
                {keys.map((k) => (
                  <option key={k} value={k}>{ymLabel(k)}</option>
                ))}
              </select>
            </label>
            <label className="text-sm">
              <span className="text-slate-700">전월(비교)</span>
              <select className={`${input} mt-1`} value={prev ?? ""} onChange={(e) => setPrevSel(e.target.value)}>
                <option value="">(비교 안 함)</option>
                {keys.filter((k) => k !== cur).map((k) => (
                  <option key={k} value={k}>{ymLabel(k)}</option>
                ))}
              </select>
            </label>
          </div>
          <details className="mt-4">
            <summary className="cursor-pointer text-sm text-slate-700">치료 분류 규칙 편집</summary>
            <p className="mt-2 text-xs text-slate-500">
              한 줄에 하나, &quot;분류명: 키워드, 키워드&quot; 형식입니다. 수납 내역의 치료 문구에 키워드가 들어 있으면 그 분류로
              집계하고, 위쪽 줄이 먼저 적용됩니다. 어디에도 안 걸리면 &apos;기타&apos;로 잡힙니다.
            </p>
            <textarea
              rows={9}
              value={rulesText}
              className={`${input} mt-2 font-mono`}
              onChange={(e) => {
                rulesDirty.current = true;
                setRulesText(e.target.value);
              }}
            />
          </details>
        </div>
      )}

      {cur && (
        <div className={`${box} space-y-8`}>
          <div className="flex flex-wrap items-center justify-between gap-2 print:hidden">
            <h2 className="font-semibold text-slate-900">
              3. 회의 자료{" "}
              <span
                className={`ml-2 text-xs font-normal ${saveState === "error" ? "text-red-600" : "text-slate-500"}`}
                aria-live="polite"
              >
                {!initial.dbReady
                  ? "저장소 미연결(저장 안 됨)"
                  : saveState === "saving"
                    ? "저장 중…"
                    : saveState === "error"
                      ? "저장 실패 — 입력을 이어가면 다시 시도합니다"
                      : saveState === "saved"
                        ? `자동 저장됨 ${savedAt}`
                        : "입력하면 자동 저장됩니다"}
              </span>
            </h2>
            <div className="flex gap-2">
              <button
                onClick={fillExample}
                title="비어 있는 칸에만 예시 문장을 넣습니다. 이미 입력한 내용은 유지됩니다."
                className="rounded-md border border-slate-300 px-3 py-1.5 text-sm"
              >
                예시 채우기
              </button>
              <button onClick={exportXlsx} className="rounded-md border border-slate-300 px-3 py-1.5 text-sm">
                엑셀로 저장
              </button>
              <button
                onClick={exportPptx}
                disabled={pptBusy}
                className="rounded-md border border-slate-300 px-3 py-1.5 text-sm disabled:opacity-50"
              >
                {pptBusy ? "PPT 만드는 중…" : "PPT로 저장"}
              </button>
              <button onClick={() => window.print()} className="rounded-md bg-slate-900 px-3 py-1.5 text-sm text-white">
                인쇄 / PDF
              </button>
            </div>
          </div>

          <section className="space-y-2">
            <h3 className="text-base font-semibold">1. 회의 개요</h3>
            <p className="text-xs text-slate-500">
              기준월 {ymLabel(cur)}
              {prev ? `, 비교월 ${ymLabel(prev)}` : ", 비교월 없음"}
            </p>
            <div className="grid gap-2 sm:grid-cols-2">
              <input className={input} placeholder="회의명 (예: 2026년 10월 월간 경영회의)" value={meta.title} onChange={(e) => setMeta({ ...meta, title: e.target.value })} />
              <input className={input} placeholder="일시" value={meta.date} onChange={(e) => setMeta({ ...meta, date: e.target.value })} />
              <input className={input} placeholder="장소" value={meta.place} onChange={(e) => setMeta({ ...meta, place: e.target.value })} />
              <input className={input} placeholder="참석자" value={meta.attendees} onChange={(e) => setMeta({ ...meta, attendees: e.target.value })} />
            </div>
          </section>

          <section className="space-y-2">
            <h3 className="text-base font-semibold">{MANUAL[0][1]}</h3>
            <textarea rows={3} className={input} value={comments.actions ?? ""} onChange={(e) => setComments({ ...comments, actions: e.target.value })} />
          </section>

          {!prev && (
            <p className="text-xs text-slate-500 print:hidden">비교할 달을 고르면 증감과 비교 표가 채워집니다.</p>
          )}

          {tables.map((t) => (
            <DataTable
              key={t.id}
              t={t}
              summary={summarize(t)}
              comment={comments[t.id] ?? ""}
              onComment={(v) => setComments((c) => ({ ...c, [t.id]: v }))}
            />
          ))}

          <section className="break-inside-avoid space-y-2">
            <h3 className="text-base font-semibold">4. 진료비 성장 시뮬레이션</h3>
            <p className="text-xs text-slate-500">
              당월 실적을 기준으로, 아래 항목이 개선되면 월 총 진료비가 얼마나 늘 수 있는지 단순 추정합니다. 숫자를 바꿔 보세요.
            </p>
            <table className="w-full border-collapse text-sm">
              <thead>
                <tr className="bg-slate-100">
                  {["개선 항목", "개선폭", "월 진료비 증가 추정(원)"].map((h) => (
                    <th key={h} className="border border-slate-200 px-2 py-1.5 text-left font-medium">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {GROWTH.map((g, i) => (
                  <tr key={g.label}>
                    <td className="border border-slate-200 px-2 py-1.5">{g.label}</td>
                    <td className="border border-slate-200 px-2 py-1">
                      <span className="inline-flex items-center gap-1">
                        <input
                          type="number"
                          className={`${input} w-20`}
                          value={growthIn[i]}
                          onChange={(e) => setGrowthIn(growthIn.map((v, j) => (j === i ? Number(e.target.value) : v)))}
                        />
                        {g.unit}
                      </span>
                    </td>
                    <td className="border border-slate-200 px-2 py-1.5 text-right tabular-nums">
                      {fmt(g.estimate(mc, growthIn[i]))}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
            <textarea
              rows={3}
              className={input}
              placeholder="이번 달 집중 과제 3가지"
              value={comments.focus ?? ""}
              onChange={(e) => setComments({ ...comments, focus: e.target.value })}
            />
          </section>

          {MANUAL.slice(1).map(([k, label]) => (
            <section key={k} className="break-inside-avoid space-y-2">
              <h3 className="text-base font-semibold">{label}</h3>
              <textarea rows={3} className={input} value={comments[k] ?? ""} onChange={(e) => setComments({ ...comments, [k]: e.target.value })} />
            </section>
          ))}
        </div>
      )}
    </div>
  );
}
