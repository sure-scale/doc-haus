// Which questions the lawyer can leave to the drafter's judgment. One rule for
// the lean workflow, which takes Decide for me and a delegating answer in the
// lawyer's own words (dochaus/workflows/lean/state.ts), and for the web, which
// offers Decide for me only where the workflow takes it. Browser-safe.

type Option = { id: string; text: string }

type DelegationQuestion = {
  kind?: string
  mixed?: true
  options?: Option[]
  advice?: { options?: Option[]; recommendation: { option: string } } | { options?: Option[]; unavailable: unknown }
}

/** Whether leaving the question to the drafter's judgment settles it: a
 * judgment question, an instruction question (the delegation is the
 * instruction it asks for) or a scope question that asks for nothing else. A
 * fact, an approval (mixed scope) or which source governs (conflict) needs the
 * lawyer's own answer. */
export function delegationSettles(question: { kind?: string; mixed?: true }) {
  return question.kind === "judgment" || question.kind === "instruction" || (question.kind === "scope" && !question.mixed)
}

/** The option delegating the question adopts: the offered option its advice
 * recommends, when delegation settles the question. */
export function delegatedOption(question: DelegationQuestion): Option | undefined {
  const advice = question.advice
  if (!delegationSettles(question) || !advice || !("recommendation" in advice)) return undefined
  return (question.options ?? advice.options ?? []).find(option => option.id === advice.recommendation.option)
}
