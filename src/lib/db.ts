import postgres from "postgres";

/* 서버 전용. DATABASE_URL이 없으면 저장 기능만 꺼지고 앱은 그대로 동작한다. */

let sql: ReturnType<typeof postgres> | null = null;
let ready: Promise<void> | null = null;

export const hasDb = (): boolean => !!process.env.DATABASE_URL;

/** 붙여넣다가 생기기 쉬운 앞뒤 공백·따옴표와 끝의 쿼리(?pgbouncer=true 등)를 걷어낸다. */
function cleanUrl(raw: string): string {
  const u = raw.trim().replace(/^["']+|["']+$/g, "").trim();
  const q = u.indexOf("?");
  return q >= 0 ? u.slice(0, q) : u;
}

export async function db(): Promise<ReturnType<typeof postgres>> {
  const url = process.env.DATABASE_URL ? cleanUrl(process.env.DATABASE_URL) : "";
  if (!url) throw new Error("NO_DB");
  sql ??= postgres(url, { prepare: false, max: 1, idle_timeout: 20, connect_timeout: 15 });
  const conn = sql;
  ready ??= conn
    .unsafe(
      `create table if not exists meeting_state (
         key text primary key,
         value jsonb not null,
         updated_at timestamptz not null default now()
       )`,
    )
    .then(() => undefined)
    .catch((e) => {
      ready = null;
      throw e;
    });
  await ready;
  return conn;
}
