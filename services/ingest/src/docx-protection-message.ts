/** User-facing wording for shared writer preservation refusals. Never expose
 * the amendment identifier embedded in an internal error. */
export function docxProtectionMessage(error: unknown): string | undefined {
  if (!(error instanceof Error)) return
  if (error.message === "docx_protected_inline_content_changed")
    return "Word content or its anchors could not be confirmed as preserved. Review the passage in Word and prepare a fresh change."
  if (/^docx_protected_inline_content(?::[^\r\n]*)?$/.test(error.message))
    return "This change touches protected Word content, such as a control or annotation. Review the passage in Word and prepare a fresh change."
}
