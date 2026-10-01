import { logout } from "@/app/actions";
import MeetingApp from "@/components/MeetingApp";
import { requireAuth } from "@/lib/session";

export default async function Home() {
  await requireAuth();
  return <MeetingApp logout={logout} />;
}
