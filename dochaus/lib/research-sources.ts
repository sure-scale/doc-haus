// Official-source hosts, matched exactly or as any subdomain. Covers every
// jurisdiction pack source that is not already covered by an official suffix.
export const OFFICIAL_SOURCE_HOSTS = [
  "legislation.gov.uk", // EW / SCT / NIR
  "eur-lex.europa.eu", // EU
  "gesetze-im-internet.de", // DE
  "wetten.overheid.nl", // NL
  "fedlex.admin.ch", // CH
  "irishstatutebook.ie", // IE
  "indiacode.nic.in", // IN
  "ontario.ca", // CA-ON (ontario.ca/laws)
  "leg.state.fl.us", // US-FL (Online Sunshine)
  "gov.za", // ZA
  "difc.ae", // AE (DIFC publishes its own laws)
  "elitigation.sg", // SG judiciary judgments
]

// Official public suffixes — any host under these is government-operated.
export const OFFICIAL_SOURCE_SUFFIXES = [
  ".gov", // US federal and state legislatures
  ".gov.uk",
  ".europa.eu",
  ".gc.ca", // Canada federal (laws-lois.justice.gc.ca)
  ".gov.bc.ca", // CA-BC (bclaws.gov.bc.ca)
  ".gouv.qc.ca", // CA-QC (legisquebec.gouv.qc.ca)
  ".gov.au", // AU federal and states
  ".govt.nz",
  ".gouv.fr", // FR (legifrance.gouv.fr)
  ".admin.ch",
  ".gov.sg", // SG (sso.agc.gov.sg)
  ".gov.hk", // HK (elegislation.gov.hk)
  ".gov.za",
  ".gov.ae",
  ".nic.in",
]

