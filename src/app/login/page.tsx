import { redirect } from "next/navigation";
import { isAuthed } from "@/lib/session";
import LoginForm from "./LoginForm";

export default async function LoginPage() {
  if (await isAuthed()) redirect("/");
  return (
    <main className="flex min-h-screen items-center justify-center bg-slate-50 px-4">
      <div className="w-full max-w-sm rounded-xl border border-slate-200 bg-white p-8 shadow-sm">
        <h1 className="text-xl font-bold text-slate-900">월간 경영회의 자료</h1>
        <p className="mt-1 text-sm text-slate-500">
          관리자 비밀번호를 입력해 주세요.
        </p>
        <LoginForm />
      </div>
    </main>
  );
}
