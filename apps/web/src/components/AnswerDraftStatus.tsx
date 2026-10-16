export default function AnswerDraftStatus({
  loaded,
  saving,
  error,
  retry,
}: {
  loaded: boolean
  saving: boolean
  error: string
  retry: () => unknown
}) {
  if (error)
    return (
      <p className="msg-error" role="alert">
        {error}{" "}
        <button type="button" onClick={() => retry()}>
          Retry
        </button>
      </p>
    )
  if (!loaded || saving)
    return (
      <p className="run-note" role="status">
        {loaded ? "Saving answers…" : "Loading saved answers…"}
      </p>
    )
  return null
}
