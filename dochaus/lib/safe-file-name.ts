// A file name every platform can hold, for a name that comes from outside (an
// email subject or attachment, a zip entry, an upload, an agent's artifact), so
// a matter folder moves between macOS, Linux and Windows unchanged. Windows
// refuses <>:"/\|?* and control characters (a colon names an NTFS alternate
// stream, so "RE: Offer.msg" would hide the content in a file named "RE"), drops
// trailing dots and spaces, and opens a device for CON, PRN, AUX, NUL, COM0-9
// and LPT0-9 (and COM/LPT with ¹²³), with or without an extension ("nul.docx").
// Leading dots, which hide a file on macOS and Linux, go too. A reserved
// character becomes a space (nothing before a dot or the end of the name) and a
// device name gets a "_" after it. A safe name comes back unchanged; an empty
// result is the caller's to replace.
const RESERVED = /\s*[<>:"/\\|?*\u0000-\u001f][\s<>:"/\\|?*\u0000-\u001f]*/g
const RESERVED_BEFORE_DOT = /\s*[<>:"/\\|?*\u0000-\u001f][\s<>:"/\\|?*\u0000-\u001f]*(?=\.|$)/g
const DEVICE = /^(con|prn|aux|nul|com[0-9¹²³]|lpt[0-9¹²³])(?=\s*(?:\.|$))/i

export function safeFileName(name: string) {
  return name
    .replace(RESERVED_BEFORE_DOT, "")
    .replace(RESERVED, " ")
    .replace(/^[.\s]+|[.\s]+$/g, "")
    .replace(DEVICE, "$1_")
}
