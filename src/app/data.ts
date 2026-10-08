"use server";

import postgres from "postgres";
import { db, hasDb, inspectDbUrl } from "@/lib/db";
import { isAuthed } from "@/lib/session";

/*
 * 입력한 내용과 올린 집계 자료를 저장한다. 모든 함수는 로그인한 사용자만 쓸 수 있다.
 * key 규칙: month:YYYY-MM (집계 자료), report:YYYY-MM (회의 개요·코멘트·시뮬레이션 입력), rules (치료 분류 규칙)
 */

export type Loaded = {
  dbReady: boolean;
  /** 저장소를 못 쓰는 이유(비밀 값은 포함하지 않음). 정상이면 null */
  dbProblem: string | null;
  months: Record<string, unknown>;
  reports: Record<string, unknown>;
  rules: string | null;
  /** 휴지통에 있는 달과 삭제 시각(ISO). 내용은 서버에만 두고 보내지 않는다 */
  trash: { ym: string; deletedAt: string }[];
  trashDays: number;
};
export type SaveResult = { ok: boolean; error?: string };

const TRASH_DAYS = 30;
const YM = /^\d{4}-\d{2}$/;
const MAX_BYTES = 3_000_000;

/** 연결 오류를 사용자가 고칠 수 있는 말로 바꾼다. 주소나 비밀번호는 내보내지 않는다. */
function describeDbError(e: unknown): string {
  const msg = e instanceof Error ? e.message : String(e);
  const code = (e as { code?: string } | null)?.code ?? "";
  console.error("[db]", code || msg.split("\n")[0].slice(0, 120));
  if (/password authentication failed/i.test(msg)) return "비밀번호가 맞지 않습니다. 연결 주소의 데이터베이스 비밀번호를 확인해 주세요.";
  if (/tenant or user not found/i.test(msg)) return "사용자·프로젝트 식별자가 맞지 않습니다. Supabase의 Transaction pooler 주소를 그대로 복사했는지 확인해 주세요.";
  if (code === "ENOTFOUND" || /getaddrinfo/i.test(msg)) return "서버 주소를 찾지 못했습니다. 주소 중간(호스트 이름)이 잘렸거나 틀렸는지 확인해 주세요.";
  if (code === "ETIMEDOUT" || code === "CONNECT_TIMEOUT" || /timeout/i.test(msg)) return "데이터베이스에 연결하는 시간이 초과됐습니다. 포트 6543의 Transaction pooler 주소인지 확인해 주세요.";
  if (e instanceof TypeError || /invalid url|invalid connection/i.test(msg)) return `연결 주소 형식이 올바르지 않습니다. 점검 결과: ${inspectDbUrl()}`;
  return `데이터베이스에 연결하지 못했습니다(${(code || msg.split("\n")[0]).slice(0, 80)}). 주소 점검: ${inspectDbUrl()}`;
}

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
  const empty: Loaded = { dbReady: false, dbProblem: null, months: {}, reports: {}, rules: null, trash: [], trashDays: TRASH_DAYS };
  if (!(await isAuthed())) return empty;
  if (!hasDb()) {
    return { ...empty, dbProblem: "서버가 DATABASE_URL을 받지 못했습니다. Vercel 환경변수에 DATABASE_URL이 있는지, Production에 체크되어 있는지 확인하고 Redeploy 해 주세요." };
  }
  try {
    const sql = await db();
    // 보관 기간이 지난 휴지통 항목은 여기서 영구 삭제한다.
    await sql`delete from meeting_state
              where key like 'trash:%' and updated_at < now() - make_interval(days => ${TRASH_DAYS}::int)`;
    // 휴지통 항목은 내용(환자 이름 포함)을 불러오지 않고 삭제 시각만 읽는다.
    const rows = await sql<{ key: string; value: unknown; updated_at: Date }[]>`
      select key, case when key like 'trash:%' then null else value end as value, updated_at from meeting_state`;
    const out: Loaded = { dbReady: true, dbProblem: null, months: {}, reports: {}, rules: null, trash: [], trashDays: TRASH_DAYS };
    for (const r of rows) {
      if (r.key.startsWith("month:")) out.months[r.key.slice(6)] = r.value;
      else if (r.key.startsWith("report:")) out.reports[r.key.slice(7)] = r.value;
      else if (r.key === "rules" && typeof r.value === "string") out.rules = r.value;
      else if (r.key.startsWith("trash:month:")) out.trash.push({ ym: r.key.slice(12), deletedAt: new Date(r.updated_at).toISOString() });
    }
    out.trash.sort((a, b) => b.deletedAt.localeCompare(a.deletedAt));
    return out;
  } catch (e) {
    return { ...empty, dbProblem: describeDbError(e) };
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

/** 삭제는 휴지통으로 옮기는 것이다. 키 앞에 'trash:'를 붙여 두고 TRASH_DAYS일 뒤 자동으로 지운다. */
export async function deleteMonth(ym: string): Promise<SaveResult> {
  if (!YM.test(ym)) return { ok: false, error: "월 형식이 올바르지 않습니다." };
  if (!(await isAuthed())) return { ok: false, error: "로그인이 필요합니다." };
  if (!hasDb()) return { ok: true };
  try {
    const sql = await db();
    await sql.begin(async (tx) => {
      // 같은 달을 전에 지웠다면 휴지통의 이전 기록은 새 기록으로 바뀐다.
      await tx`delete from meeting_state where key in (${"trash:month:" + ym}, ${"trash:report:" + ym})`;
      await tx`update meeting_state set key = 'trash:' || key, updated_at = now()
               where key in (${"month:" + ym}, ${"report:" + ym})`;
    });
    return { ok: true };
  } catch {
    return { ok: false, error: "삭제에 실패했습니다." };
  }
}

export type RestoreResult = SaveResult & { month?: unknown; report?: unknown };

export async function restoreFromTrash(ym: string): Promise<RestoreResult> {
  if (!YM.test(ym)) return { ok: false, error: "월 형식이 올바르지 않습니다." };
  if (!(await isAuthed())) return { ok: false, error: "로그인이 필요합니다." };
  if (!hasDb()) return { ok: false, error: "저장소가 설정되지 않았습니다." };
  try {
    const sql = await db();
    const res = await sql.begin(async (tx) => {
      const exists = await tx`select 1 from meeting_state where key = ${"month:" + ym}`;
      if (exists.length) return "exists" as const;
      const trashed = await tx`select 1 from meeting_state where key = ${"trash:month:" + ym}`;
      if (!trashed.length) return "missing" as const;
      await tx`delete from meeting_state
               where key = ${"report:" + ym}
                 and exists (select 1 from meeting_state where key = ${"trash:report:" + ym})`;
      await tx`update meeting_state set key = substr(key, 7), updated_at = now()
               where key in (${"trash:month:" + ym}, ${"trash:report:" + ym})`;
      return await tx<{ key: string; value: unknown }[]>`
        select key, value from meeting_state where key in (${"month:" + ym}, ${"report:" + ym})`;
    });
    if (res === "exists") return { ok: false, error: "같은 달 자료가 이미 있어 복구할 수 없습니다. 현재 자료를 먼저 삭제한 뒤 복구해 주세요." };
    if (res === "missing") return { ok: false, error: "휴지통에서 찾을 수 없습니다(보관 기간이 지났을 수 있습니다)." };
    return {
      ok: true,
      month: res.find((r) => r.key.startsWith("month:"))?.value,
      report: res.find((r) => r.key.startsWith("report:"))?.value,
    };
  } catch {
    return { ok: false, error: "복구에 실패했습니다." };
  }
}

export async function purgeFromTrash(ym: string): Promise<SaveResult> {
  if (!YM.test(ym)) return { ok: false, error: "월 형식이 올바르지 않습니다." };
  if (!(await isAuthed())) return { ok: false, error: "로그인이 필요합니다." };
  if (!hasDb()) return { ok: true };
  try {
    const sql = await db();
    await sql`delete from meeting_state where key in (${"trash:month:" + ym}, ${"trash:report:" + ym})`;
    return { ok: true };
  } catch {
    return { ok: false, error: "영구 삭제에 실패했습니다." };
  }
}
