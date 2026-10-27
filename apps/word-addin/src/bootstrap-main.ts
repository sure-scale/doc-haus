import { initializeBootstrapOffice, showBootstrap } from "./bootstrap"

const initialization = initializeBootstrapOffice(document,
  () => (window as unknown as { Office?: { onReady(): Promise<unknown> } }).Office)
void showBootstrap(document, window.location.href, url => window.location.replace(url), undefined, initialization)
