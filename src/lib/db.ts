import postgres from "postgres";

/* 서버 전용. DATABASE_URL이 없으면 저장 기능만 꺼지고 앱은 그대로 동작한다. */

let sql: ReturnType<typeof postgres> | null = null;
let ready: Promise<void> | null = null;

export const hasDb = (): boolean => !!process.env.DATABASE_URL;

/** 붙여넣다가 생기기 쉬운 앞뒤 공백·따옴표와 끝의 쿼리(?pgbouncer=true 등)를 걷어낸다. */
function cleanUrl(raw: string): string {
  const u = raw.trim().replace(/^["']+|["']+$/g, "").trim();
  const q = u.indexOf("?", u.lastIndexOf("/"));
  return q >= 0 ? u.slice(0, q) : u;
}

function decodeSafe(s: string): string {
  if (!/%[0-9A-Fa-f]{2}/.test(s)) return s;
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
}

/**
 * 비밀번호에 @ # / ? 같은 문자가 있어도 읽을 수 있게 주소를 직접 나눈다.
 * 비밀번호는 "마지막 @" 앞까지로 본다. Supabase 자리 표시처럼 [비밀번호]로 감싸져 있으면 대괄호를 벗긴다.
 */
function parseUrl(u: string) {
  const m = u.match(/^postgres(?:ql)?:\/\/([^:/@]+):(.*)@([^@/?#]+?)(?::(\d+))?\/([^?#]+)$/);
  if (!m) return null;
  let password = decodeSafe(m[2]);
  if (/^\[.*\]$/.test(password)) password = password.slice(1, -1);
  return { user: decodeSafe(m[1]), password, host: m[3], port: m[4] ? Number(m[4]) : 5432, database: m[5] };
}

/** 연결 주소의 모양만 점검한다. 비밀번호와 사용자 이름은 절대 내보내지 않는다. */
export function inspectDbUrl(): string {
  const u = cleanUrl(process.env.DATABASE_URL ?? "");
  const hints: string[] = [];
  if (!/^postgres(ql)?:\/\//.test(u)) hints.push("주소가 postgresql:// 로 시작하지 않음");
  if (/\s/.test(u)) hints.push("주소 안에 공백이나 줄바꿈이 있음");
  if (/["']/.test(u)) hints.push("따옴표가 남아 있음");
  if (/\[YOUR-PASSWORD\]/i.test(u)) hints.push("[YOUR-PASSWORD] 자리 표시가 그대로 남아 있음");
  const p = parseUrl(u);
  if (!p) {
    hints.push("user:password@host:port/database 모양으로 나눌 수 없음");
  } else {
    hints.push(`호스트 ${p.host}, 포트 ${p.port}, 데이터베이스 ${p.database}`);
    if (!/pooler\.supabase\.com$/.test(p.host)) hints.push("호스트가 pooler.supabase.com 으로 끝나지 않음(Transaction pooler 주소가 아닐 수 있음)");
    if (p.port !== 6543) hints.push("포트가 6543이 아님");
  }
  return hints.join(" / ");
}

export async function db(): Promise<ReturnType<typeof postgres>> {
  const url = process.env.DATABASE_URL ? cleanUrl(process.env.DATABASE_URL) : "";
  if (!url) throw new Error("NO_DB");
  const opts = { prepare: false, max: 1, idle_timeout: 20, connect_timeout: 15 } as const;
  const cfg = parseUrl(url);
  sql ??= cfg
    ? postgres({ host: cfg.host, port: cfg.port, user: cfg.user, password: cfg.password, database: cfg.database, ssl: "require", ...opts })
    : postgres(url, opts);
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
      sql = null;
      throw e;
    });
  await ready;
  return conn;
}
