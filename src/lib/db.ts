import postgres from "postgres";

/* 서버 전용. DATABASE_URL이 없으면 저장 기능만 꺼지고 앱은 그대로 동작한다. */

let sql: ReturnType<typeof postgres> | null = null;
let ready: Promise<void> | null = null;

export const hasDb = (): boolean => !!process.env.DATABASE_URL;

export async function db(): Promise<ReturnType<typeof postgres>> {
  const url = process.env.DATABASE_URL;
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
