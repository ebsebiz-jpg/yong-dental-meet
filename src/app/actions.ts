"use server";

import { redirect } from "next/navigation";
import { checkPassword, endSession, startSession } from "@/lib/session";

export type LoginState = { error?: string } | undefined;

export async function login(
  _prev: LoginState,
  formData: FormData,
): Promise<LoginState> {
  const password = String(formData.get("password") ?? "");
  if (!checkPassword(password)) {
    // 무차별 대입을 늦추기 위한 고정 지연
    await new Promise((r) => setTimeout(r, 800));
    return { error: "비밀번호가 올바르지 않습니다." };
  }
  await startSession();
  redirect("/");
}

export async function logout(): Promise<void> {
  await endSession();
  redirect("/login");
}
