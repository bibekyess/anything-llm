import Chat from "./Chat";
import Orb from "./Orb";

export default function App() {
  const view = (window as any).assistant?.view || "chat";
  return view === "orb" ? <Orb /> : <Chat />;
}
