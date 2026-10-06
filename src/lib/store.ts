import { GROWTH, type Month, type RouteData } from "@/lib/dentweb";

/* 서버 데이터베이스에 JSON으로 저장하기 위한 형태와 변환 함수. */

export type Report = {
  meta: { title: string; date: string; place: string; attendees: string };
  comments: Record<string, string>;
  growthIn: number[];
};

export const emptyReport = (): Report => ({
  meta: { title: "", date: "", place: "", attendees: "" },
  comments: {},
  growthIn: GROWTH.map((g) => g.def),
});

/** 저장된 값이 일부 비어 있어도 화면이 깨지지 않게 기본값과 합친다. */
export function normalizeReport(raw: unknown): Report {
  const base = emptyReport();
  const r = (raw ?? {}) as Partial<Report>;
  return {
    meta: { ...base.meta, ...(r.meta ?? {}) },
    comments: { ...(r.comments ?? {}) },
    growthIn: base.growthIn.map((d, i) => (typeof r.growthIn?.[i] === "number" ? r.growthIn[i] : d)),
  };
}

type StoredRoute = Omit<RouteData, "regions" | "referrers"> & {
  regions: [string, number][];
  referrers: [string, { name: string; count: number; referred: string[] }][];
};
type StoredMonth = Omit<Month, "route"> & { route?: StoredRoute };

export function serializeMonth(m: Month): StoredMonth {
  const { route, ...rest } = m;
  if (!route) return rest;
  return { ...rest, route: { ...route, regions: [...route.regions], referrers: [...route.referrers] } };
}

export function deserializeMonth(raw: unknown): Month {
  const s = raw as StoredMonth;
  const { route, ...rest } = s;
  if (!route) return rest;
  return { ...rest, route: { ...route, regions: new Map(route.regions), referrers: new Map(route.referrers) } };
}
