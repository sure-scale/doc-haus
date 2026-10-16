// A lawyer's message as typed: plain text, with `code` spans (often file names)
// shown as code rather than as literal backticks.
export default function UserText({ text }: { text: string }) {
  return <>{text.split(/`([^`\n]+)`/).map((part, index) => (index % 2 ? <code key={index}>{part}</code> : part))}</>
}
