import type { Cell, Table } from "@/lib/dentweb";

/* 표의 숫자만 읽어 슬라이드·화면에 넣을 핵심 한 줄 요약을 만든다. 표에 없는 사실은 쓰지 않는다. */

const num = (c: Cell | undefined): number | null => (typeof c === "number" ? c : null);
const n0 = (v: number) => Math.round(v).toLocaleString("ko-KR");
const r1 = (v: number) => Math.round(v * 10) / 10;
const sign = (v: number, unit: string, dec = false) => {
  const a = dec ? r1(Math.abs(v)) : Math.round(Math.abs(v));
  return `${v > 0 ? "+" : v < 0 ? "-" : ""}${a.toLocaleString("ko-KR")}${unit}`;
};
const rowOf = (t: Table, label: string) => t.rows.find((r) => String(r[0]).trim() === label);

/** 전월 대비 증감 문구. 증감이 없으면 null */
function change(diff: number | null, unit: string, dec = false): string | null {
  if (diff === null) return null;
  if (diff === 0) return "전월과 같음";
  return `전월 대비 ${sign(diff, unit, dec)}`;
}

function summarizeUnsafe(t: Table): string | null {
  switch (t.id) {
    case "t31": {
      const tot = rowOf(t, "총 진료비");
      const non = rowOf(t, "비급여 비중(%)");
      const cur = num(tot?.[1]);
      if (cur === null) return null;
      const pct = num(tot?.[4]);
      let s = `총 진료비는 ${n0(cur)}원`;
      s += pct === null ? "입니다." : pct === 0 ? "으로 전월과 같습니다." : `으로 전월 대비 ${r1(Math.abs(pct))}% ${pct > 0 ? "증가" : "감소"}했습니다.`;
      const share = num(non?.[1]);
      if (share !== null) {
        const d = num(non?.[3]);
        s += ` 비급여 비중은 ${share}%${d !== null && d !== 0 ? `(${sign(d, "%p", true)})` : ""}입니다.`;
      }
      return s;
    }
    case "t32": {
      const p = rowOf(t, "총 내원 환자 수(명)");
      const f = rowOf(t, "신환(명)");
      const a = rowOf(t, "환자 1인당 평균 진료비(원)");
      const pc = num(p?.[1]);
      if (pc === null) return null;
      const pd = change(num(p?.[3]), "명");
      let s = `내원 환자는 ${n0(pc)}명${pd ? `(${pd})` : ""}`;
      const fc = num(f?.[1]);
      if (fc !== null) {
        const fd = change(num(f?.[3]), "명");
        s += `, 신환은 ${n0(fc)}명${fd ? `(${fd})` : ""}`;
      }
      const ac = num(a?.[1]);
      if (ac !== null) {
        const ap = num(a?.[4]);
        s += `, 환자 1인당 진료비는 ${n0(ac)}원${ap !== null && ap !== 0 ? `(${sign(ap, "%", true)})` : ""}`;
      }
      return `${s}입니다.`;
    }
    case "t33": {
      const rows = t.rows.filter((r) => r[0] !== "기타");
      if (!rows.length) return null;
      const top = rows[0];
      let s = `수납액 비중 1위는 ${top[0]}(${num(top[2]) ?? "-"}%)입니다.`;
      const moved = rows
        .filter((r) => num(r[5]) !== null)
        .sort((x, y) => Math.abs(num(y[5]) ?? 0) - Math.abs(num(x[5]) ?? 0))[0];
      if (moved && num(moved[5]) !== 0) s += ` 전월 대비 비중이 가장 크게 변한 항목은 ${moved[0]}(${sign(num(moved[5]) ?? 0, "%p", true)})입니다.`;
      return s;
    }
    case "t34": {
      const top = t.rows[0];
      if (!top) return null;
      let s = `내원 경로 1위는 ${top[0]}(${num(top[2]) ?? "-"}%)입니다.`;
      const withDiff = t.rows.filter((r) => num(r[11]) !== null);
      const up = [...withDiff].sort((x, y) => (num(y[11]) ?? 0) - (num(x[11]) ?? 0))[0];
      const down = [...withDiff].sort((x, y) => (num(x[11]) ?? 0) - (num(y[11]) ?? 0))[0];
      const parts: string[] = [];
      if (up && (num(up[11]) ?? 0) > 0) parts.push(`${up[0]}에서 ${sign(num(up[11]) ?? 0, "명")}`);
      if (down && (num(down[11]) ?? 0) < 0) parts.push(`${down[0]}에서 ${sign(num(down[11]) ?? 0, "명")}`);
      if (parts.length) s += ` 신환 변화가 큰 경로는 ${parts.join(", ")}입니다.`;
      return s;
    }
    case "t34b": {
      const top = t.rows.find((r) => r[0] !== "기타·미기재") ?? t.rows[0];
      if (!top) return null;
      return `신환의 ${num(top[2]) ?? "-"}%가 ${top[0]}에 거주합니다.`;
    }
    case "t35": {
      const p = rowOf(t, "소개 내원 환자(명)");
      const share = rowOf(t, "전체 내원 환자 중 비중(%)");
      const f = rowOf(t, "소개 신환(명)");
      const pc = num(p?.[1]);
      if (pc === null) return null;
      let s = `소개 내원 환자는 ${n0(pc)}명`;
      const sh = num(share?.[1]);
      if (sh !== null) s += `(전체의 ${sh}%)`;
      const fc = num(f?.[1]);
      if (fc !== null) {
        const fd = change(num(f?.[3]), "명");
        s += `, 소개 신환은 ${n0(fc)}명${fd ? `(${fd})` : ""}`;
      }
      return `${s}입니다.`;
    }
    case "t35b": {
      const top = t.rows[0];
      if (!top) return null;
      const counts = t.rows.map((r) => num(r[1]) ?? 0);
      if (Math.min(...counts) === Math.max(...counts)) {
        return `표에 나온 소개자(상위 ${t.rows.length}명)는 모두 신환을 ${counts[0]}명씩 소개했습니다.`;
      }
      return `가장 많이 소개한 환자는 ${top[0]}(${num(top[1]) ?? "-"}명)입니다.`;
    }
    case "t36": {
      const rate = rowOf(t, "확정 비율(%)");
      const amount = rowOf(t, "확정 금액(원)");
      const rc = num(rate?.[1]);
      if (rc === null) return null;
      const rd = num(rate?.[3]);
      let s = `상담 확정 비율은 ${rc}%`;
      if (rd !== null && rd !== 0) s += `로 전월 대비 ${r1(Math.abs(rd))}%p ${rd > 0 ? "상승" : "하락"}`;
      const ac = num(amount?.[1]);
      if (ac !== null) {
        const ap = num(amount?.[4]);
        s += `${rd !== null && rd !== 0 ? "했고" : "이고"}, 확정 금액은 ${n0(ac)}원${ap !== null && ap !== 0 ? `(${sign(ap, "%", true)})` : ""}`;
      }
      return `${s}입니다.`;
    }
    case "t36b": {
      // 계획 5건 미만인 상담자는 표본이 작아 비교에서 뺀다.
      const rows = t.rows.filter((r) => (num(r[1]) ?? 0) >= 5 && num(r[3]) !== null);
      if (!rows.length) return null;
      const low = [...rows].sort((x, y) => (num(x[3]) ?? 0) - (num(y[3]) ?? 0))[0];
      let s = `확정 비율이 가장 낮은 상담자는 ${low[0]}(${num(low[3])}%)`;
      const drop = [...rows.filter((r) => num(r[10]) !== null)].sort((x, y) => (num(x[10]) ?? 0) - (num(y[10]) ?? 0))[0];
      if (drop && (num(drop[10]) ?? 0) < 0) s += `이고, 전월 대비 가장 크게 떨어진 상담자는 ${drop[0]}(${sign(num(drop[10]) ?? 0, "%p", true)})`;
      return `${s}입니다. (계획 5건 이상 기준)`;
    }
    case "t37": {
      const total = t.rows.find((r) => r[0] === "합계");
      const cur = num(total?.[1]);
      if (cur === null) return null;
      const d = change(num(total?.[3]), "건");
      const rate = num(total?.[5]);
      return `기공 의뢰는 ${n0(cur)}건${d ? `(${d})` : ""}${rate !== null ? `, 재제작률은 ${rate}%` : ""}입니다.`;
    }
    case "t38": {
      if (t.rows.length < 2) return null;
      const a = t.rows[0];
      const b = t.rows[t.rows.length - 1];
      const t0 = num(a[1]);
      const t1 = num(b[1]);
      const p0 = num(a[3]);
      const p1 = num(b[3]);
      if (t0 === null || t1 === null) return null;
      const w = (x: number, y: number) => (y > x ? "늘었" : y < x ? "줄었" : "같았");
      let s = `${a[0]}부터 ${b[0]}까지 총 진료비는 ${n0(t0)}원에서 ${n0(t1)}원으로 ${w(t0, t1)}습니다.`;
      if (p0 !== null && p1 !== null) s = s.replace(/습니다\.$/, `고, 내원 환자는 ${n0(p0)}명에서 ${n0(p1)}명으로 ${w(p0, p1)}습니다.`);
      return s;
    }
    case "t39": {
      const vol = rowOf(t, "내원 환자 수 변화");
      const prc = rowOf(t, "환자 1인당 진료비 변화");
      const tot = t.rows.find((r) => r[0] === "합계");
      const v = num(vol?.[1]);
      const p = num(prc?.[1]);
      const d = num(tot?.[1]);
      if (v === null || p === null || d === null) return null;
      const big = Math.abs(v) >= Math.abs(p) ? { label: "내원 환자 수 변화", value: v } : { label: "환자 1인당 진료비 변화", value: p };
      return `총 진료비 ${sign(d, "원")} 변화는 ${big.label}(${sign(big.value, "원")})의 영향이 더 큽니다.`;
    }
    default:
      return null;
  }
}

export function summarize(t: Table): string | null {
  try {
    return summarizeUnsafe(t);
  } catch {
    return null;
  }
}

/** 추이 그래프 슬라이드용 요약 */
export function seriesSummary(label: string, labels: string[], values: number[], unit: string): string | null {
  if (values.length < 2) return null;
  const first = values[0];
  const last = values[values.length - 1];
  const w = last > first ? "늘었습니다" : last < first ? "줄었습니다" : "같습니다";
  return `${label}은 ${labels[0]} ${n0(first)}${unit}에서 ${labels[labels.length - 1]} ${n0(last)}${unit}으로 ${w}.`;
}
