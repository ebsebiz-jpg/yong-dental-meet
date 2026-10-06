"use server";

import { headers } from "next/headers";
import { redirect } from "next/navigation";
import { checkPassword, endSession, startSession } from "@/lib/session";

export type LoginState = { error?: string } | undefined;

// 접속 주소(IP)별 실패 횟수. 서버 인스턴스 메모리에만 있어 완벽한 차단은 아니지만 무차별 대입을 크게 늦춘다.
const MAX_FAILS = 5;
const WINDOW_MS = 15 * 60 * 1000;
const fails = new Map<string, { count: number; first: number }>();

async function clientKey(): Promise<string> {
  const h = await headers();
  return h.get("x-forwarded-for")?.split(",")[0]?.trim() || h.get("x-real-ip") || "local";
}

export async function login(
  _prev: LoginState,
  formData: FormData,
): Promise<LoginState> {
  const key = await clientKey();
  const now = Date.now();
  const rec = fails.get(key);
  if (rec && now - rec.first > WINDOW_MS) fails.delete(key);
  const cur = fails.get(key);
  if (cur && cur.count >= MAX_FAILS) {
    return { error: "시도가 너무 많습니다. 15분 뒤에 다시 시도해 주세요." };
  }

  const password = String(formData.get("password") ?? "");
  if (!checkPassword(password)) {
    fails.set(key, { count: (cur?.count ?? 0) + 1, first: cur?.first ?? now });
    // 무차별 대입을 늦추기 위한 고정 지연
    await new Promise((r) => setTimeout(r, 800));
    return { error: "비밀번호가 올바르지 않습니다." };
  }
  fails.delete(key);
  await startSession();
  redirect("/");
}

export async function logout(): Promise<void> {
  await endSession();
  redirect("/login");
}
