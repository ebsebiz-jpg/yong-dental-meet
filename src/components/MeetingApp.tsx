"use client";

import { useMemo, useState } from "react";
import * as XLSX from "xlsx";
import {
  DEFAULT_RULES,
  FIELDS,
  GROWTH,
  buildTables,
  computeStats,
  guessMapping,
  parseRules,
  readSheet,
  type Cell,
  type FieldKey,
  type Mapping,
  type Sheet,
  type Table,
} from "@/lib/analysis";

const RULES_KEY = "dm_rules";
const fmt = (c: Cell) =>
  c === null || c === "" ? "–" : typeof c === "number" ? c.toLocaleString("ko-KR") : c;

const box = "rounded-xl border border-slate-200 bg-white p-5 print:border-0 print:p-0";
const input =
  "w-full rounded-md border border-slate-300 px-2 py-1.5 text-sm outline-none focus:border-slate-900";

function FileSlot({
  label,
  sheet,
  onLoad,
  optional,
}: {
  label: string;
  sheet: Sheet | null;
  onLoad: (s: Sheet) => void;
  optional?: boolean;
}) {
  const [err, setErr] = useState("");
  return (
    <label className="block rounded-lg border border-dashed border-slate-300 p-4 text-sm">
      <span className="font-medium text-slate-800">
        {label} {optional && <span className="text-slate-400">(선택)</span>}
      </span>
      <input
        type="file"
        accept=".xlsx,.xls,.csv"
        className="mt-2 block w-full text-xs"
        onChange={async (e) => {
          const f = e.target.files?.[0];
          if (!f) return;
          try {
            setErr("");
            onLoad(await readSheet(f));
          } catch {
            setErr("파일을 읽지 못했습니다. 엑셀(.xlsx, .xls) 또는 CSV인지 확인해 주세요.");
          }
        }}
      />
      {sheet && (
        <span className="mt-2 block text-xs text-emerald-700">
          {sheet.rows.length.toLocaleString()}행, {sheet.headers.length}열 읽음
        </span>
      )}
      {err && <span className="mt-2 block text-xs text-red-600">{err}</span>}
    </label>
  );
}

function DataTable({ t, comment, onComment }: { t: Table; comment: string; onComment: (v: string) => void }) {
  return (
    <section className="break-inside-avoid">
      <h3 className="text-base font-semibold text-slate-900">{t.title}</h3>
      {t.note && <p className="mt-1 text-xs text-slate-500">{t.note}</p>}
      <div className="mt-2 overflow-x-auto">
        <table className="w-full border-collapse text-sm">
          <thead>
            <tr className="bg-slate-100">
              {t.headers.map((h) => (
                <th key={h} className="border border-slate-200 px-2 py-1.5 text-left font-medium">
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
                    className={`border border-slate-200 px-2 py-1.5 ${j > 0 ? "text-right tabular-nums" : ""}`}
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

const MANUAL = [
  ["actions", "전월 액션아이템 점검"],
  ["staff", "인력 현황 (입·퇴사, 교육, 근무 관련 안건)"],
  ["ops", "진료·운영 현황 (장비, 감염관리, 재료 재고, 기공소)"],
  ["risk", "리스크 및 이슈"],
  ["decide", "논의 안건 및 의사결정"],
  ["next", "차월 목표 및 액션아이템"],
] as const;

const EXAMPLE_META = {
  title: "2026년 10월 월간 경영회의",
  date: "2026. 10. 05 (월) 19:00 ~ 20:30",
  place: "상담실",
  attendees: "원장, 부원장, 실장, 치과위생사 팀장, 코디네이터",
};

// 문장 속 ○○, ○는 직접 채워 넣는 자리입니다. 숫자는 예시일 뿐 실제 자료와 무관합니다.
const EXAMPLE_COMMENTS: Record<string, string> = {
  actions:
    "- 노쇼 환자 리마인드 문자 발송 체계 정비: 담당 ○○, 완료\n- 임플란트 상담 자료 개선: 담당 ○○, 진행 중 (○월 ○일까지)\n- 소독실 장비 점검: 담당 ○○, 지연 (업체 일정 조율 중)",
  t31: "총 진료비가 전월 대비 ○% 증가했습니다. 비보험 진료(임플란트·교정) 증가가 주된 원인이고, 영업일수는 전월과 같습니다.",
  t32: "신환이 전월보다 ○명 늘었으나 환자 1인당 진료비는 줄었습니다. 저단가 진료(스케일링 등) 환자 비중이 늘어난 영향으로 보입니다.",
  t33: "임플란트 비중이 ○%p 감소하고 보철 비중이 증가했습니다. 전월 식립 환자의 보철 진행 시기가 겹친 영향으로 판단합니다.",
  t34: "예약 이행률이 전월보다 ○%p 낮아졌습니다. 주로 취소가 늘었고, 오후 시간대와 토요일에 집중되어 있습니다. 전일 리마인드 문자 발송을 강화합니다.",
  t35: "네이버 검색·온라인 광고 유입이 늘었고 간판 유입은 줄었습니다. ○○동 거주 환자 비중이 높아 해당 지역 광고 집행을 유지합니다.",
  t35b: "○○구 환자가 가장 많고 평균 진료비도 높습니다. ○○동은 환자가 줄어 원인 확인이 필요합니다.",
  t36: "소개환자가 신환의 ○%를 차지하고, 소개환자 1인당 매출은 전체 평균보다 높습니다. 소개해 주신 환자께 감사 인사를 드립니다.",
  t36b: "소개를 많이 해 준 환자 상위 분께 감사 문자와 혜택 안내를 보냅니다.",
  staff: "- 치과위생사 1명 입사(○월 ○일), 조무사 1명 퇴사 예정\n- 보수교육 이수 현황 점검\n- 연차·근무시간 조정 건 논의",
  ops: "- 유닛체어 ○번 유압 점검 필요\n- 멸균 기록 누락 없이 관리 중\n- 임플란트 재료 재고 부족 예상, 발주 예정\n- 기공소 납기 지연 1건, 업체와 협의",
  risk: "- 의료분쟁: 임플란트 환자 불만 1건, 경위서 작성 및 사후 관리 중\n- 보험 심사: 삭감 예상 청구 건 사전 점검\n- 인력 이탈: 핵심 직원 면담 필요",
  decide: "1. 신규 CT 장비 도입 여부 → 견적 비교 후 다음 회의에서 결정\n2. 비보험 상담 프로세스 개편 → 상담실장 주도로 시범 운영\n3. 마케팅 예산 조정 → 효율 낮은 채널 축소",
  next: "- 총 진료비 ○원, 신환 ○명, 예약 이행률 ○% 목표\n- [ ] 리마인드 문자 개선: 담당 ○○, ○월 ○일까지\n- [ ] 소개환자 감사 프로그램 시행: 담당 ○○, ○월 ○일까지\n- 다음 회의: ○월 ○일 ○요일 19:00",
  focus: "1. 노쇼·취소 감소(리마인드 강화)\n2. 상담 후 치료 동의율 향상(상담 자료·스크립트 개선)\n3. 소개환자 확대(감사 프로그램 운영)",
};

const COSTS = ["재료비", "인건비", "임대료·관리비", "마케팅비", "기타 운영비"];

export default function MeetingApp({ logout }: { logout: () => Promise<void> }) {
  const [cur, setCur] = useState<Sheet | null>(null);
  const [prev, setPrev] = useState<Sheet | null>(null);
  const [mapping, setMapping] = useState<Mapping>({});
  const [rulesText, setRulesText] = useState(DEFAULT_RULES);
  const [meta, setMeta] = useState({ title: "", date: "", place: "", attendees: "" });
  const [comments, setComments] = useState<Record<string, string>>({});
  const [costCur, setCostCur] = useState<Record<string, string>>({});
  const [costPrev, setCostPrev] = useState<Record<string, string>>({});
  const [growthIn, setGrowthIn] = useState(GROWTH.map((g) => g.def));

  const rules = useMemo(() => parseRules(rulesText), [rulesText]);

  const prevMapping = useMemo<Mapping>(() => {
    if (!prev) return {};
    const g = guessMapping(prev.headers);
    const out: Mapping = {};
    for (const f of FIELDS) {
      const c = mapping[f.key];
      out[f.key] = c && prev.headers.includes(c) ? c : g[f.key];
    }
    return out;
  }, [prev, mapping]);

  const curStats = useMemo(() => (cur ? computeStats(cur, mapping, rules) : null), [cur, mapping, rules]);
  const prevStats = useMemo(
    () => (prev ? computeStats(prev, prevMapping, rules) : null),
    [prev, prevMapping, rules],
  );
  const tables = useMemo(() => (curStats ? buildTables(curStats, prevStats) : []), [curStats, prevStats]);

  const sumCost = (c: Record<string, string>) =>
    COSTS.reduce((s, k) => s + (parseFloat((c[k] ?? "").replace(/,/g, "")) || 0), 0);
  const costC = sumCost(costCur);
  const costP = sumCost(costPrev);

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

  const exportXlsx = () => {
    const wb = XLSX.utils.book_new();
    const overview = [
      ["회의명", meta.title],
      ["일시", meta.date],
      ["장소", meta.place],
      ["참석자", meta.attendees],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(overview), "회의 개요");
    for (const t of tables) {
      const aoa: Cell[][] = [t.headers, ...t.rows];
      const c = comments[t.id];
      if (c) aoa.push([], ["분석 코멘트", c]);
      const name = t.title.replace(/[\\/?*[\]:]/g, "").slice(0, 31);
      XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name);
    }
    const manual = MANUAL.map(([k, label]) => [label, comments[k] ?? ""]);
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(manual), "직접 입력 항목");
    XLSX.writeFile(wb, `월간경영회의_${meta.date || "자료"}.xlsx`);
  };

  const missing = FIELDS.filter((f) => !mapping[f.key]);

  return (
    <div className="mx-auto max-w-5xl space-y-6 px-4 py-8">
      <header className="flex items-center justify-between print:hidden">
        <h1 className="text-xl font-bold text-slate-900">월간 경영회의 자료 만들기</h1>
        <form action={logout}>
          <button className="text-sm text-slate-500 underline">로그아웃</button>
        </form>
      </header>

      <div className="rounded-lg bg-amber-50 p-3 text-xs text-amber-900 print:hidden">
        업로드한 엑셀은 이 브라우저 안에서만 읽고 계산하며 서버로 전송하거나 저장하지 않습니다.
        회의 자료를 인쇄·공유하기 전에 환자 이름 등 식별정보가 없는지 확인해 주세요.
      </div>

      <div className={`${box} print:hidden`}>
        <h2 className="font-semibold text-slate-900">1. 자료 업로드</h2>
        <p className="mt-1 text-xs text-slate-500">
          덴트웹에서 내보낸 진료 내역(엑셀·CSV)을 올립니다. 환자 1건 진료가 1행인 자료가 가장 잘 맞습니다.
        </p>
        <div className="mt-3 grid gap-3 sm:grid-cols-2">
          <FileSlot
            label="당월 자료"
            sheet={cur}
            onLoad={(s) => {
              setCur(s);
              setMapping(guessMapping(s.headers));
              try {
                const saved = localStorage.getItem(RULES_KEY);
                if (saved) setRulesText(saved);
              } catch {}
            }}
          />
          <FileSlot label="전월 자료" sheet={prev} onLoad={setPrev} optional />
        </div>
      </div>

      {cur && (
        <div className={`${box} print:hidden`}>
          <h2 className="font-semibold text-slate-900">2. 열 연결 확인</h2>
          <p className="mt-1 text-xs text-slate-500">
            자동으로 찾은 결과입니다. 틀리면 바꾸고, 자료에 없는 항목은 비워 두세요. 비운 항목의 분석 표는 만들어지지 않습니다.
          </p>
          <div className="mt-3 grid gap-3 sm:grid-cols-2">
            {FIELDS.map((f) => (
              <label key={f.key} className="text-sm">
                <span className="text-slate-700">{f.label}</span>
                <select
                  className={`${input} mt-1`}
                  value={mapping[f.key] ?? ""}
                  onChange={(e) =>
                    setMapping((m) => ({ ...m, [f.key as FieldKey]: e.target.value || undefined }))
                  }
                >
                  <option value="">(없음)</option>
                  {cur.headers.map((h) => (
                    <option key={h} value={h}>
                      {h}
                    </option>
                  ))}
                </select>
              </label>
            ))}
          </div>
          {missing.length > 0 && (
            <p className="mt-3 text-xs text-slate-500">
              연결되지 않은 항목: {missing.map((f) => `${f.label}(${f.needed})`).join(", ")}
            </p>
          )}
          <details className="mt-4">
            <summary className="cursor-pointer text-sm text-slate-700">치료 분류 규칙 편집</summary>
            <p className="mt-2 text-xs text-slate-500">
              한 줄에 하나, &quot;분류명: 키워드, 키워드&quot; 형식입니다. 치료 항목명에 키워드가 들어 있으면 그 분류로 집계합니다.
            </p>
            <textarea
              rows={8}
              value={rulesText}
              className={`${input} mt-2 font-mono`}
              onChange={(e) => {
                setRulesText(e.target.value);
                try {
                  localStorage.setItem(RULES_KEY, e.target.value);
                } catch {}
              }}
            />
          </details>
        </div>
      )}

      {curStats && (
        <div className={`${box} space-y-8`}>
          <div className="flex flex-wrap items-center justify-between gap-2 print:hidden">
            <h2 className="font-semibold text-slate-900">3. 회의 자료</h2>
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
              <button onClick={() => window.print()} className="rounded-md bg-slate-900 px-3 py-1.5 text-sm text-white">
                인쇄 / PDF
              </button>
            </div>
          </div>

          <section className="space-y-2">
            <h3 className="text-base font-semibold">1. 회의 개요</h3>
            <div className="grid gap-2 sm:grid-cols-2">
              <input className={input} placeholder="회의명 (예: 2026년 9월 월간 경영회의)" value={meta.title} onChange={(e) => setMeta({ ...meta, title: e.target.value })} />
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
            <p className="text-xs text-slate-500 print:hidden">전월 자료를 올리면 증감과 비교 표가 채워집니다.</p>
          )}

          {tables.map((t) => (
            <DataTable
              key={t.id}
              t={t}
              comment={comments[t.id] ?? ""}
              onComment={(v) => setComments((c) => ({ ...c, [t.id]: v }))}
            />
          ))}

          <section className="break-inside-avoid space-y-2">
            <h3 className="text-base font-semibold">4. 비용 및 영업이익 (직접 입력, 원)</h3>
            <table className="w-full border-collapse text-sm">
              <thead>
                <tr className="bg-slate-100">
                  {["항목", "당월", "전월"].map((h) => (
                    <th key={h} className="border border-slate-200 px-2 py-1.5 text-left font-medium">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {COSTS.map((k) => (
                  <tr key={k}>
                    <td className="border border-slate-200 px-2 py-1">{k}</td>
                    <td className="border border-slate-200 px-1 py-1"><input inputMode="numeric" className={input} value={costCur[k] ?? ""} onChange={(e) => setCostCur({ ...costCur, [k]: e.target.value })} /></td>
                    <td className="border border-slate-200 px-1 py-1"><input inputMode="numeric" className={input} value={costPrev[k] ?? ""} onChange={(e) => setCostPrev({ ...costPrev, [k]: e.target.value })} /></td>
                  </tr>
                ))}
                <tr className="font-medium">
                  <td className="border border-slate-200 px-2 py-1.5">비용 합계</td>
                  <td className="border border-slate-200 px-2 py-1.5 text-right tabular-nums">{costC.toLocaleString("ko-KR")}</td>
                  <td className="border border-slate-200 px-2 py-1.5 text-right tabular-nums">{costP.toLocaleString("ko-KR")}</td>
                </tr>
                <tr className="font-medium">
                  <td className="border border-slate-200 px-2 py-1.5">영업이익 (총 진료비 − 비용)</td>
                  <td className="border border-slate-200 px-2 py-1.5 text-right tabular-nums">{curStats.total !== null ? Math.round(curStats.total - costC).toLocaleString("ko-KR") : "–"}</td>
                  <td className="border border-slate-200 px-2 py-1.5 text-right tabular-nums">{prevStats?.total != null ? Math.round(prevStats.total - costP).toLocaleString("ko-KR") : "–"}</td>
                </tr>
              </tbody>
            </table>
            <p className="text-xs text-slate-500">
              진료비는 수납·청구 기준이라 실제 입금액과 다를 수 있고, 공단 삭감·환수는 반영되지 않습니다.
            </p>
          </section>

          <section className="break-inside-avoid space-y-2">
            <h3 className="text-base font-semibold">5. 매출 성장 시뮬레이션</h3>
            <p className="text-xs text-slate-500">
              당월 실적을 기준으로, 아래 항목이 개선되면 월 매출이 얼마나 늘 수 있는지 단순 추정합니다. 숫자를 바꿔 보세요.
            </p>
            <table className="w-full border-collapse text-sm">
              <thead>
                <tr className="bg-slate-100">
                  {["개선 항목", "개선폭", "월 매출 증가 추정(원)"].map((h) => (
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
                      {fmt(g.estimate(curStats, growthIn[i]))}
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
