"use client";

import { useActionState } from "react";
import { login } from "@/app/actions";

export default function LoginForm() {
  const [state, action, pending] = useActionState(login, undefined);
  return (
    <form action={action} className="mt-6 space-y-4">
      <input
        type="password"
        name="password"
        required
        autoFocus
        autoComplete="current-password"
        placeholder="비밀번호"
        className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm outline-none focus:border-slate-900"
      />
      {state?.error && <p className="text-sm text-red-600">{state.error}</p>}
      <button
        type="submit"
        disabled={pending}
        className="w-full rounded-lg bg-slate-900 py-2 text-sm font-medium text-white disabled:opacity-50"
      >
        {pending ? "확인 중…" : "로그인"}
      </button>
    </form>
  );
}
