import { createHash, createHmac, timingSafeEqual } from "node:crypto";
import { cookies } from "next/headers";
import { redirect } from "next/navigation";

const COOKIE = "dm_session";
const TTL_SECONDS = 60 * 60 * 8;

function secret(): string {
  const s = process.env.SESSION_SECRET;
  if (!s || s.length < 16) {
    throw new Error("SESSION_SECRET 환경변수(16자 이상)가 필요합니다.");
  }
  return s;
}

function sign(payload: string): string {
  return createHmac("sha256", secret()).update(payload).digest("base64url");
}

function safeEqual(a: string, b: string): boolean {
  const ba = Buffer.from(a);
  const bb = Buffer.from(b);
  return ba.length === bb.length && timingSafeEqual(ba, bb);
}

export function checkPassword(input: string): boolean {
  const expected = process.env.ADMIN_PASSWORD;
  if (!expected) throw new Error("ADMIN_PASSWORD 환경변수가 필요합니다.");
  const h = (v: string) => createHash("sha256").update(v).digest("hex");
  return safeEqual(h(input), h(expected));
}

function verifyToken(token: string | undefined): boolean {
  if (!token) return false;
  const [exp, sig] = token.split(".");
  if (!exp || !sig || !safeEqual(sig, sign(exp))) return false;
  return Number(exp) > Date.now() / 1000;
}

export async function isAuthed(): Promise<boolean> {
  const store = await cookies();
  return verifyToken(store.get(COOKIE)?.value);
}

export async function requireAuth(): Promise<void> {
  if (!(await isAuthed())) redirect("/login");
}

export async function startSession(): Promise<void> {
  const exp = String(Math.floor(Date.now() / 1000) + TTL_SECONDS);
  const store = await cookies();
  store.set(COOKIE, `${exp}.${sign(exp)}`, {
    httpOnly: true,
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production",
    path: "/",
    maxAge: TTL_SECONDS,
  });
}

export async function endSession(): Promise<void> {
  const store = await cookies();
  store.delete(COOKIE);
}
