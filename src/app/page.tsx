import { logout } from "@/app/actions";
import { loadAll } from "@/app/data";
import MeetingApp from "@/components/MeetingApp";
import { requireAuth } from "@/lib/session";

export default async function Home() {
  await requireAuth();
  const initial = await loadAll();
  return <MeetingApp logout={logout} initial={initial} />;
}
