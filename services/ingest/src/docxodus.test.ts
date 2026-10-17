import { expect, test } from "bun:test"

// Child processes report JSON on stdout, so a cold engine start and its first
// compare must add nothing there while ordinary console.log output still passes.
test("a cold engine start in a new Bun process keeps stdout clean", async () => {
  const child = Bun.spawn([process.execPath, "-e", `
    import {docxodus} from './docxodus';
    const dx = await docxodus();
    const blank = dx.createBlankDocx();
    await dx.compareDocuments(blank, blank);
    console.log("ready");
  `], { cwd: import.meta.dirname, stdout: "pipe", stderr: "pipe" })
  const stdout = await new Response(child.stdout).text()
  expect(await child.exited, await new Response(child.stderr).text()).toBe(0)
  expect(stdout).toBe("ready\n")
})
