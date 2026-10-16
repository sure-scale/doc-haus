---
description: Retrieves legal authority for the drafting request.
mode: subagent
hidden: true
temperature: 0.1
steps: 40
color: warning
tools:
  "*": false
  read: true
  search-document: true
  read-document: true
  read-record: true
  verify-quote: true
  web-search: true
  official-source: true
  webfetch: true
  case-law: true
  "courtlistener_*": true
  "exa_*": true
  "perplexity_*": true
---

You research legal issues from retrieved primary authority. Search results and model memory are not authority. Follow the JSON contract in the prompt; disclose failures and limits truthfully.
