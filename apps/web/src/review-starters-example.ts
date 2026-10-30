// Static worked example for the "Compare agreements" review-table starter.
// Fictional, synthetic content: no model runs. Each excerpt is text.slice(charStart, charEnd).

export type StarterExampleCitation = { section: string; excerpt: string; charStart: number; charEnd: number }

export type StarterExampleAnswer = {
  document: string
  question: number
  answer: string
  citation?: StarterExampleCitation
  comparison?: "consistent" | "differs"
  reason?: string
}

export type StarterExample = {
  note: string
  documents: readonly { name: string; text: string }[]
  answers: readonly StarterExampleAnswer[]
}

export const COMPARE_AGREEMENTS_EXAMPLE: StarterExample = {
  note: "Fictional example. These agreements and answers were written to show how the table works.",
  documents: [
    {
      name: "Quillbarrow packaging supply agreement.docx",
      text: [
        "PACKAGING SUPPLY AGREEMENT",
        "Parties",
        "This agreement is made on 14 January 2026 between Thistlemere Provisions Limited (the Customer) and Quillbarrow Packaging Limited (the Supplier).",
        "1. Term",
        "This agreement starts on 1 February 2026 and continues for an initial term of 12 months.",
        "2. Renewal",
        "This agreement then renews automatically for further periods of 12 months unless either party gives the other at least 60 days' written notice before the end of the current term that it does not wish to renew.",
        "3. Supply",
        "The Supplier will supply the cardboard boxes, trays and labels set out in each purchase order, to the agreed specification.",
        "4. Price and payment",
        "The prices are set out in the Schedule. The Customer will pay each undisputed invoice within 30 days of receiving it.",
        "5. Termination",
        "Either party may terminate this agreement by giving the other at least 3 months' written notice. Either party may also terminate this agreement immediately by written notice if the other party commits a material breach and does not remedy it within 30 days of being asked to do so, or becomes insolvent.",
        "6. Liability",
        "Each party's total liability arising under or in connection with this agreement is limited to the amounts paid or payable by the Customer under this agreement in the 12 months before the event giving rise to the claim. This limit does not apply to the Customer's obligation to pay the prices. Neither party is liable for loss of profit or for any indirect or consequential loss. Nothing in this agreement limits or excludes liability for death or personal injury caused by negligence, for fraud or fraudulent misrepresentation, or for any other liability that cannot be limited or excluded by law.",
        "7. Confidentiality",
        "Each party will keep the other's confidential information secret and use it only to perform this agreement.",
        "8. Governing law and disputes",
        "This agreement, and any dispute arising from it, is governed by the law of England and Wales. The parties will first try to settle any dispute through discussion between senior managers. If the dispute is not settled within 30 days, either party may refer it to the courts of England and Wales, which have exclusive jurisdiction.",
      ].join("\n"),
    },
    {
      name: "Hollowmarsh logistics services agreement.docx",
      text: [
        "LOGISTICS SERVICES AGREEMENT",
        "Parties",
        "This agreement is made on 2 February 2026 between Thistlemere Provisions Limited (the Customer) and Hollowmarsh Freight Limited (the Supplier).",
        "1. Term",
        "This agreement starts on 1 March 2026 and continues for an initial term of 12 months.",
        "2. Renewal",
        "This agreement renews automatically for successive periods of 12 months unless either party gives the other at least 60 days' written notice before the end of the current term that it does not wish to renew.",
        "3. Services",
        "The Supplier will collect finished goods from the Customer's bakery and deliver them to the Customer's shops on the agreed weekly schedule.",
        "4. Charges and payment",
        "The charges are set out in the Schedule. The Customer will pay each undisputed invoice within 30 days of receiving it.",
        "5. Termination",
        "Either party may end this agreement by giving at least 3 months' written notice. Either party may also end it immediately by written notice if the other party commits a material breach that is not remedied within 30 days of a written request, or becomes insolvent.",
        "6. Liability",
        "Each party's total liability arising under or in connection with this agreement is limited to £250,000 in aggregate. This limit does not apply to the Customer's obligation to pay the charges. Neither party is liable for loss of profit or for any indirect or consequential loss. Nothing in this agreement limits or excludes liability for death or personal injury caused by negligence, for fraud or fraudulent misrepresentation, or for any other liability that cannot be limited or excluded by law.",
        "7. Confidentiality",
        "Each party will keep the other's confidential information secret and use it only to perform this agreement.",
        "8. Governing law and disputes",
        "This agreement, and any dispute arising from it, is governed by the law of England and Wales. Any dispute will first be referred to senior managers of each party for discussion. If it is not settled within 30 days, either party may bring proceedings in the courts of England and Wales, which have exclusive jurisdiction.",
      ].join("\n"),
    },
    {
      name: "Sorrelgate cleaning services agreement.docx",
      text: [
        "CLEANING SERVICES AGREEMENT",
        "Parties",
        "This agreement is made on 9 March 2026 between Thistlemere Provisions Limited (the Customer) and Sorrelgate Cleaning Limited (the Supplier).",
        "1. Term",
        "This agreement starts on 1 April 2026 and continues for a term of 12 months.",
        "2. Services",
        "The Supplier will clean the Customer's bakery and offices each weekday evening to the standard set out in the Schedule.",
        "3. Site rules",
        "The Supplier's staff will follow the Customer's food hygiene and site safety rules while on site.",
        "4. Charges and payment",
        "The monthly charge is set out in the Schedule. The Customer will pay each undisputed invoice within 30 days of receiving it.",
        "5. Termination",
        "Either party may terminate this agreement on at least 3 months' written notice to the other. Either party may also terminate it immediately by written notice if the other party commits a material breach and fails to remedy it within 30 days of being asked to do so, or becomes insolvent.",
        "6. Liability",
        "Each party's total liability arising under or in connection with this agreement is limited to the amounts paid or payable by the Customer under this agreement in the 12 months before the event giving rise to the claim. This limit does not apply to the Customer's obligation to pay the charges. Neither party is liable for loss of profit or for any indirect or consequential loss. Nothing in this agreement limits or excludes liability for death or personal injury caused by negligence, for fraud or fraudulent misrepresentation, or for any other liability that cannot be limited or excluded by law.",
        "7. Confidentiality",
        "Each party will keep the other's confidential information secret and use it only to perform this agreement.",
        "8. Governing law and disputes",
        "This agreement, and any dispute arising from it, is governed by the law of England and Wales. The parties will first try to settle any dispute through discussion between senior managers. If the dispute is not settled within 30 days, either party may refer it to the courts of England and Wales, which have exclusive jurisdiction.",
      ].join("\n"),
    },
  ],
  answers: [
    {
      document: "Quillbarrow packaging supply agreement.docx",
      question: 0,
      answer:
        "Thistlemere Provisions Limited (Customer) and Quillbarrow Packaging Limited (Supplier). Made on 14 January 2026.",
      citation: {
        section: "Parties",
        excerpt:
          "This agreement is made on 14 January 2026 between Thistlemere Provisions Limited (the Customer) and Quillbarrow Packaging Limited (the Supplier).",
        charStart: 35,
        charEnd: 180,
      },
      comparison: "consistent",
    },
    {
      document: "Quillbarrow packaging supply agreement.docx",
      question: 1,
      answer:
        "Yes. It renews for 12 months at a time unless either party gives at least 60 days' written notice before the current term ends.",
      citation: {
        section: "2 Renewal",
        excerpt:
          "This agreement then renews automatically for further periods of 12 months unless either party gives the other at least 60 days' written notice before the end of the current term that it does not wish to renew.",
        charStart: 289,
        charEnd: 498,
      },
      comparison: "consistent",
    },
    {
      document: "Quillbarrow packaging supply agreement.docx",
      question: 2,
      answer:
        "Either party may terminate on 3 months' written notice, or immediately by notice if the other party commits a material breach not remedied within 30 days or becomes insolvent.",
      citation: {
        section: "5 Termination",
        excerpt:
          "Either party may terminate this agreement by giving the other at least 3 months' written notice. Either party may also terminate this agreement immediately by written notice if the other party commits a material breach and does not remedy it within 30 days of being asked to do so, or becomes insolvent.",
        charStart: 787,
        charEnd: 1090,
      },
      comparison: "consistent",
    },
    {
      document: "Quillbarrow packaging supply agreement.docx",
      question: 3,
      answer:
        "Each party's liability is capped at the amounts paid or payable in the 12 months before the claim event, except the Customer's duty to pay; loss of profit and indirect or consequential loss are excluded. Death or personal injury caused by negligence, fraud, and liability that cannot be limited by law are not limited.",
      citation: {
        section: "6 Liability",
        excerpt:
          "Each party's total liability arising under or in connection with this agreement is limited to the amounts paid or payable by the Customer under this agreement in the 12 months before the event giving rise to the claim. This limit does not apply to the Customer's obligation to pay the prices. Neither party is liable for loss of profit or for any indirect or consequential loss. Nothing in this agreement limits or excludes liability for death or personal injury caused by negligence, for fraud or fraudulent misrepresentation, or for any other liability that cannot be limited or excluded by law.",
        charStart: 1104,
        charEnd: 1701,
      },
      comparison: "consistent",
    },
    {
      document: "Quillbarrow packaging supply agreement.docx",
      question: 4,
      answer:
        "Law of England and Wales. Disputes go to senior managers first, then to the courts of England and Wales if not settled within 30 days.",
      citation: {
        section: "8 Governing law and disputes",
        excerpt:
          "This agreement, and any dispute arising from it, is governed by the law of England and Wales. The parties will first try to settle any dispute through discussion between senior managers. If the dispute is not settled within 30 days, either party may refer it to the courts of England and Wales, which have exclusive jurisdiction.",
        charStart: 1859,
        charEnd: 2188,
      },
      comparison: "consistent",
    },
    {
      document: "Hollowmarsh logistics services agreement.docx",
      question: 0,
      answer:
        "Thistlemere Provisions Limited (Customer) and Hollowmarsh Freight Limited (Supplier). Made on 2 February 2026.",
      citation: {
        section: "Parties",
        excerpt:
          "This agreement is made on 2 February 2026 between Thistlemere Provisions Limited (the Customer) and Hollowmarsh Freight Limited (the Supplier).",
        charStart: 37,
        charEnd: 180,
      },
      comparison: "consistent",
    },
    {
      document: "Hollowmarsh logistics services agreement.docx",
      question: 1,
      answer:
        "Yes. It renews for 12 months at a time unless either party gives at least 60 days' written notice before the current term ends.",
      citation: {
        section: "2 Renewal",
        excerpt:
          "This agreement renews automatically for successive periods of 12 months unless either party gives the other at least 60 days' written notice before the end of the current term that it does not wish to renew.",
        charStart: 286,
        charEnd: 493,
      },
      comparison: "consistent",
    },
    {
      document: "Hollowmarsh logistics services agreement.docx",
      question: 2,
      answer:
        "Either party may end it on 3 months' written notice, or immediately by notice if the other party commits a material breach not remedied within 30 days or becomes insolvent.",
      citation: {
        section: "5 Termination",
        excerpt:
          "Either party may end this agreement by giving at least 3 months' written notice. Either party may also end it immediately by written notice if the other party commits a material breach that is not remedied within 30 days of a written request, or becomes insolvent.",
        charStart: 803,
        charEnd: 1067,
      },
      comparison: "consistent",
    },
    {
      document: "Hollowmarsh logistics services agreement.docx",
      question: 3,
      answer:
        "Each party's liability is capped at £250,000 in aggregate, except the Customer's duty to pay; loss of profit and indirect or consequential loss are excluded. Death or personal injury caused by negligence, fraud, and liability that cannot be limited by law are not limited.",
      citation: {
        section: "6 Liability",
        excerpt:
          "Each party's total liability arising under or in connection with this agreement is limited to £250,000 in aggregate. This limit does not apply to the Customer's obligation to pay the charges. Neither party is liable for loss of profit or for any indirect or consequential loss. Nothing in this agreement limits or excludes liability for death or personal injury caused by negligence, for fraud or fraudulent misrepresentation, or for any other liability that cannot be limited or excluded by law.",
        charStart: 1081,
        charEnd: 1577,
      },
      comparison: "differs",
      reason: "Fixed cap of £250,000 in aggregate.",
    },
    {
      document: "Hollowmarsh logistics services agreement.docx",
      question: 4,
      answer:
        "Law of England and Wales. Disputes go to senior managers first, then to the courts of England and Wales if not settled within 30 days.",
      citation: {
        section: "8 Governing law and disputes",
        excerpt:
          "This agreement, and any dispute arising from it, is governed by the law of England and Wales. Any dispute will first be referred to senior managers of each party for discussion. If it is not settled within 30 days, either party may bring proceedings in the courts of England and Wales, which have exclusive jurisdiction.",
        charStart: 1735,
        charEnd: 2055,
      },
      comparison: "consistent",
    },
    {
      document: "Sorrelgate cleaning services agreement.docx",
      question: 0,
      answer:
        "Thistlemere Provisions Limited (Customer) and Sorrelgate Cleaning Limited (Supplier). Made on 9 March 2026.",
      citation: {
        section: "Parties",
        excerpt:
          "This agreement is made on 9 March 2026 between Thistlemere Provisions Limited (the Customer) and Sorrelgate Cleaning Limited (the Supplier).",
        charStart: 36,
        charEnd: 176,
      },
      comparison: "consistent",
    },
    {
      document: "Sorrelgate cleaning services agreement.docx",
      question: 1,
      answer: "Not addressed",
    },
    {
      document: "Sorrelgate cleaning services agreement.docx",
      question: 2,
      answer:
        "Either party may terminate on 3 months' written notice, or immediately by notice if the other party commits a material breach not remedied within 30 days or becomes insolvent.",
      citation: {
        section: "5 Termination",
        excerpt:
          "Either party may terminate this agreement on at least 3 months' written notice to the other. Either party may also terminate it immediately by written notice if the other party commits a material breach and fails to remedy it within 30 days of being asked to do so, or becomes insolvent.",
        charStart: 669,
        charEnd: 956,
      },
      comparison: "consistent",
    },
    {
      document: "Sorrelgate cleaning services agreement.docx",
      question: 3,
      answer:
        "Each party's liability is capped at the amounts paid or payable in the 12 months before the claim event, except the Customer's duty to pay; loss of profit and indirect or consequential loss are excluded. Death or personal injury caused by negligence, fraud, and liability that cannot be limited by law are not limited.",
      citation: {
        section: "6 Liability",
        excerpt:
          "Each party's total liability arising under or in connection with this agreement is limited to the amounts paid or payable by the Customer under this agreement in the 12 months before the event giving rise to the claim. This limit does not apply to the Customer's obligation to pay the charges. Neither party is liable for loss of profit or for any indirect or consequential loss. Nothing in this agreement limits or excludes liability for death or personal injury caused by negligence, for fraud or fraudulent misrepresentation, or for any other liability that cannot be limited or excluded by law.",
        charStart: 970,
        charEnd: 1568,
      },
      comparison: "consistent",
    },
    {
      document: "Sorrelgate cleaning services agreement.docx",
      question: 4,
      answer:
        "Law of England and Wales. Disputes go to senior managers first, then to the courts of England and Wales if not settled within 30 days.",
      citation: {
        section: "8 Governing law and disputes",
        excerpt:
          "This agreement, and any dispute arising from it, is governed by the law of England and Wales. The parties will first try to settle any dispute through discussion between senior managers. If the dispute is not settled within 30 days, either party may refer it to the courts of England and Wales, which have exclusive jurisdiction.",
        charStart: 1726,
        charEnd: 2055,
      },
      comparison: "consistent",
    },
  ],
}
