// Placeholder shown while a reopened conversation's messages load, so the panel
// reads as "this thread is coming" rather than flashing the fresh-chat starters.
// Mirrors the message rhythm — a short user bubble answered by a taller assistant
// block — with shimmer bars standing in for the text.
export default function ChatSkeleton() {
  return (
    <div className="chat-skeleton" aria-hidden>
      {[0, 1].map((i) => (
        <div key={i} className="sk-turn">
          <div className="sk-bubble sk-user">
            <div className="sk-line" style={{ width: "60%" }} />
          </div>
          <div className="sk-bubble sk-assistant">
            <div className="sk-line" style={{ width: "92%" }} />
            <div className="sk-line" style={{ width: "98%" }} />
            <div className="sk-line" style={{ width: "74%" }} />
          </div>
        </div>
      ))}
    </div>
  )
}
