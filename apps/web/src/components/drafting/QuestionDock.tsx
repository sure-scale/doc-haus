import type { LeanQuestion } from "../../api/ingest/lean"
import QuestionDockShell from "../chat/QuestionDockShell"
import Questions from "./Questions"

// The run's questions for the lawyer; with nothing asked, it shows nothing.
export default function QuestionDock(props: {
  matterId: string
  ownerSessionID: string
  questions: LeanQuestion[]
  answered: number
  answeredQuestions?: LeanQuestion[]
  queuedAnswerIDs?: string[]
  titles: Map<string, string>
  busy: boolean
  error: string
  errorFor?: string
  onSend: (question: LeanQuestion, answer: string) => Promise<boolean>
  onChoose: (question: LeanQuestion, optionID: string) => Promise<boolean>
  onDecide: (question: LeanQuestion) => Promise<boolean>
}) {
  return (
    <QuestionDockShell
      count={props.questions.length}
      answeredCount={props.answeredQuestions?.length ?? 0}
      preview={props.questions[0]?.question ?? ""}
      subject="The drafting run"
      renderQuestions={(titleId, onClose) => <Questions {...props} titleId={titleId} onClose={onClose} />}
    />
  )
}
