import type { LeanQuestion } from "../../api/ingest/lean"
import QuestionDockShell from "../chat/QuestionDockShell"
import Questions from "./Questions"

export default function QuestionDock(props: {
  matterId: string
  ownerSessionID: string
  questions: LeanQuestion[]
  answered: number
  handedBack: number
  titles: Map<string, string>
  busy: boolean
  error: string
  onSend: (question: LeanQuestion, answer: string) => Promise<boolean>
  onDecide: (question: LeanQuestion) => Promise<boolean>
}) {
  return (
    <QuestionDockShell
      count={props.questions.length}
      preview={props.questions[0]?.question ?? ""}
      subject="The drafting run"
      renderQuestions={(titleId, onClose) => <Questions {...props} titleId={titleId} onClose={onClose} />}
    />
  )
}
