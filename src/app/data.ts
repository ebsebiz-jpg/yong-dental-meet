"use server";

import postgres from "postgres";
import { db, hasDb } from "@/lib/db";
import { isAuthed } from "@/lib/session";

/*
 * 입력한 내용과 올린 집계 자료를 저장한다. 모든 함수는 로그인한 사용자만 쓸 수 있다.
 * key 규칙: month:YYYY-MM (집계 자료), report:YYYY-MM (회의 개요·코멘트·시뮬레이션 입력), rules (치료 분류 규칙)
 */

export type Loaded = {
  dbReady: boolean;
  months: Record<string, unknown>;
  reports: Record<string, unknown>;
  rules: string | null;
};
export type SaveResult = { ok: boolean; error?: string };

const YM = /^\d{4}-\d{2}$/;
const MAX_BYTES = 3_000_000;

async function put(key: string, value: unknown): Promise<SaveResult> {
  if (!(await isAuthed())) return { ok: false, error: "로그인이 필요합니다." };
  if (!hasDb()) return { ok: false, error: "저장소가 설정되지 않았습니다." };
  const json = JSON.stringify(value);
  if (json.length > MAX_BYTES) return { ok: false, error: "저장할 내용이 너무 큽니다." };
  try {
    const sql = await db();
    await sql`
      insert into meeting_state (key, value, updated_at)
      values (${key}, ${sql.json(JSON.parse(json) as postgres.JSONValue)}, now())
      on conflict (key) do update set value = excluded.value, updated_at = now()`;
    return { ok: true };
  } catch {
    return { ok: false, error: "저장에 실패했습니다." };
  }
}

export async function loadAll(): Promise<Loaded> {
  const empty: Loaded = { dbReady: false, months: {}, reports: {}, rules: null };
  if (!(await isAuthed()) || !hasDb()) return empty;
  try {
    const sql = await db();
    const rows = await sql<{ key: string; value: unknown }[]>`select key, value from meeting_state`;
    const out: Loaded = { dbReady: true, months: {}, reports: {}, rules: null };
    for (const r of rows) {
      if (r.key.startsWith("month:")) out.months[r.key.slice(6)] = r.value;
      else if (r.key.startsWith("report:")) out.reports[r.key.slice(7)] = r.value;
      else if (r.key === "rules" && typeof r.value === "string") out.rules = r.value;
    }
    return out;
  } catch {
    return empty;
  }
}

export async function saveMonth(ym: string, data: unknown): Promise<SaveResult> {
  if (!YM.test(ym)) return { ok: false, error: "월 형식이 올바르지 않습니다." };
  return put(`month:${ym}`, data);
}

export async function saveReport(ym: string, data: unknown): Promise<SaveResult> {
  if (!YM.test(ym)) return { ok: false, error: "월 형식이 올바르지 않습니다." };
  return put(`report:${ym}`, data);
}

export async function saveRules(text: string): Promise<SaveResult> {
  return put("rules", String(text).slice(0, 5000));
}

export async function deleteMonth(ym: string): Promise<SaveResult> {
  if (!YM.test(ym)) return { ok: false, error: "월 형식이 올바르지 않습니다." };
  if (!(await isAuthed())) return { ok: false, error: "로그인이 필요합니다." };
  if (!hasDb()) return { ok: true };
  try {
    const sql = await db();
    await sql`delete from meeting_state where key in (${"month:" + ym}, ${"report:" + ym})`;
    return { ok: true };
  } catch {
    return { ok: false, error: "삭제에 실패했습니다." };
  }
}
