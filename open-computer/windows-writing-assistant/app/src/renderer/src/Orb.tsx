/** Floating orb shown while the chat is hidden: draggable ring, clickable core. */
export default function Orb() {
  const api = (window as any).assistant;
  return (
    <div className="orb-wrap">
      <button
        className="orb"
        title="Open Writing Assistant (Alt+Space)"
        onClick={() => api.win.restore()}
      >
        ✍️
      </button>
    </div>
  );
}
