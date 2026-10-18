import type { LeanQuestion } from "../../api/ingest/lean"
import QuestionDockShell from "../chat/QuestionDockShell"
import Questions from "./Questions"

// The run's questions for the lawyer. It stays while a question is left to the
// drafter, so the lawyer can still answer it instead.
export default function QuestionDock(props: {
  matterId: string
  ownerSessionID: string
  questions: LeanQuestion[]
  answered: number
  handedBack: LeanQuestion[]
  titles: Map<string, string>
  busy: boolean
  error: string
  errorFor?: string
  onSend: (question: LeanQuestion, answer: string) => Promise<boolean>
  onDecide: (question: LeanQuestion) => Promise<boolean>
}) {
  const left = props.handedBack.length
  return (
    <QuestionDockShell
      count={props.questions.length}
      kept={left}
      keptTitle={left === 1 ? "1 question left to the drafter" : `${left} questions left to the drafter`}
      preview={(props.questions[0] ?? props.handedBack[0])?.question ?? ""}
      subject="The drafting run"
      renderQuestions={(titleId, onClose) => <Questions {...props} titleId={titleId} onClose={onClose} />}
    />
  )
}
