import { expect, test, type Page } from '@playwright/test'
import fs from 'node:fs'
import path from 'node:path'
import { JSDOM } from 'jsdom'
import { importDocx } from '../../src/docx/import'
import { validatePackage } from '../../src/docx/validate'

const out = path.resolve('tests/output')
// the DOCX re-import check runs in the Node test runner, which needs a DOM XML parser
const dom = new JSDOM('')
Object.assign(globalThis, { DOMParser: dom.window.DOMParser, XMLSerializer: dom.window.XMLSerializer })

async function signUp(page: Page, email: string) {
  await page.goto('/#/signin')
  await page.getByRole('button', { name: 'Create an account' }).click()
  await page.getByLabel('Name').fill(email.split('@')[0])
  await page.getByLabel('Email').fill(email)
  await page.getByLabel('Password').fill('correct horse battery')
  await page.getByRole('button', { name: 'Create account' }).click()
  await expect(page.getByRole('heading', { name: 'Start a document' })).toBeVisible()
}

const body = (page: Page) => page.getByRole('textbox', { name: 'Document body' })

test('opens straight into a blank document with no sign-in, and reopens it next time', async ({ page }) => {
  await page.goto('/')
  await expect(body(page)).toBeVisible({ timeout: 20000 })
  await expect(page.getByRole('heading', { name: 'Sign in' })).toHaveCount(0)
  await expect(page.getByLabel('Document title')).toHaveValue('Document')
  await body(page).click()
  await page.keyboard.type('Typed without an account')
  await expect(page.getByRole('status').filter({ hasText: 'All changes saved' })).toBeVisible({ timeout: 15000 })
  await page.waitForTimeout(4500)
  const url = page.url()
  await page.goto('/')
  await expect(body(page)).toContainText('Typed without an account', { timeout: 20000 })
  expect(page.url()).toBe(url)
  // File → New gives a fresh blank document
  await page.getByRole('button', { name: 'File menu' }).click()
  await page.getByRole('menuitem', { name: /New blank document/ }).click()
  await expect.poll(() => page.url()).not.toBe(url)
  await expect(body(page)).toBeVisible()
  await expect(body(page)).not.toContainText('Typed without an account')
  // upgrading the guest keeps every document
  await page.getByRole('button', { name: 'File menu' }).click()
  await page.getByRole('menuitem', { name: /Save to an account/ }).click()
  const email = `upgraded${Date.now()}@example.com`
  await page.getByLabel('Name').fill('Upgraded User')
  await page.getByLabel('Email').fill(email)
  await page.getByLabel('Password').fill('correct horse battery')
  await page.getByRole('button', { name: 'Create account' }).click()
  await expect(page.getByText('Account created')).toBeVisible()
  await page.goto('/#/documents')
  await expect(page.locator('.sc-doclist li')).toHaveCount(2)
})

test('create, format, paginate, autosave, reopen, export DOCX/PDF', async ({ page }) => {
  test.setTimeout(60_000)
  const errors: string[] = []
  page.on('pageerror', (e) => errors.push(e.message))
  await signUp(page, `writer${Date.now()}@example.com`)
  await page.getByRole('button', { name: /Blank document/ }).click()
  await expect(body(page)).toBeVisible()
  await page.getByLabel('Document title').fill('E2E Report')

  // typing + formatting via keyboard and ribbon
  await body(page).click()
  await page.keyboard.type('Quarterly report')
  await page.keyboard.press('ControlOrMeta+Alt+1')
  await page.keyboard.press('Enter')
  await page.keyboard.press('ControlOrMeta+b')
  await page.keyboard.type('Bold start')
  await page.keyboard.press('ControlOrMeta+b')
  await page.keyboard.type(' and normal text.')
  await expect(page.locator('.sc-doc h1')).toHaveText('Quarterly report')
  await expect(page.locator('.sc-doc strong')).toHaveText('Bold start')

  // enough content for several pages
  await page.keyboard.press('Enter')
  for (let i = 0; i < 70; i++) {
    await page.keyboard.insertText(`Paragraph ${i + 1}: The quick brown fox jumps over the lazy dog, repeatedly, to fill this page with text.`)
    await page.keyboard.press('Enter')
  }
  await expect.poll(async () => page.locator('.sc-page').count(), { timeout: 15000 }).toBeGreaterThanOrEqual(3)
  // page break creates a new page
  const before = await page.locator('.sc-page').count()
  await page.keyboard.press('ControlOrMeta+Enter')
  await page.keyboard.type('After the break')
  await expect.poll(async () => page.locator('.sc-page').count(), { timeout: 15000 }).toBe(before + 1)

  // pagination: no text line crosses a page boundary
  const overlaps = await page.evaluate(() => {
    const pages = [...document.querySelectorAll<HTMLElement>('.sc-page')].map((p) => p.getBoundingClientRect())
    const bad: string[] = []
    // measure every line box of real text (page spacers are excluded: they are not text nodes)
    const root = document.querySelector('.ProseMirror.sc-doc')!
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT)
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      if (!n.textContent!.trim()) continue
      const range = document.createRange()
      range.selectNodeContents(n)
      for (const r of range.getClientRects()) {
        if (r.height === 0) continue
        const inside = pages.some((pg) => r.top >= pg.top - 1 && r.bottom <= pg.bottom + 1)
        if (!inside) bad.push(`${n.textContent!.slice(0, 24)} line[${Math.round(r.top)}-${Math.round(r.bottom)}]`)
      }
    }
    return bad
  })
  expect(overlaps).toEqual([])

  // page setup: A4 via Layout tab
  await page.getByRole('tab', { name: 'Layout' }).click()
  await page.getByRole('button', { name: 'A4' }).click()
  await expect.poll(async () => Math.round((await page.locator('.sc-page').first().boundingBox())!.width)).toBe(794)

  // footer with page numbers
  await page.getByRole('tab', { name: 'Insert' }).click()
  await page.getByRole('button', { name: 'Page numbers' }).click()
  await expect(page.locator('.sc-page').nth(1).locator('.sc-page-footer')).toHaveText('2')

  // undo / redo
  await body(page).click()
  await page.keyboard.press('ControlOrMeta+End')
  await page.keyboard.type(' UNDO-ME')
  await expect(body(page)).toContainText('UNDO-ME')
  await page.keyboard.press('ControlOrMeta+z')
  await expect(body(page)).not.toContainText('UNDO-ME')
  await page.keyboard.press('ControlOrMeta+Shift+z')
  await expect(body(page)).toContainText('UNDO-ME')

  // find & replace
  await page.keyboard.press('ControlOrMeta+h')
  await page.getByLabel('Find', { exact: true }).fill('lazy dog')
  await expect(page.locator('.sc-find-count')).toHaveText(/of 70/)
  await page.getByLabel('Replace with').fill('sleepy cat')
  await page.getByRole('button', { name: 'Replace all' }).click()
  await expect(body(page)).not.toContainText('lazy dog')
  await page.getByLabel('Close find').click()

  // autosave + reopen
  await expect(page.getByRole('status').filter({ hasText: 'All changes saved' })).toBeVisible({ timeout: 15000 })
  await page.waitForTimeout(5000)
  await page.reload()
  await expect(page.locator('.sc-doc h1')).toHaveText('Quarterly report', { timeout: 15000 })
  await expect(body(page)).toContainText('sleepy cat')
  await expect(page.getByLabel('Document title')).toHaveValue('E2E Report')

  // DOCX export
  const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('button', { name: 'Download as DOCX' }).click()])
  const file = path.join(out, 'e2e-export.docx')
  await dl.saveAs(file)
  const bytes = new Uint8Array(fs.readFileSync(file))
  expect(await validatePackage(bytes)).toEqual([])
  const re = await importDocx(bytes)
  const json = JSON.stringify(re.doc.content)
  expect(json).toContain('Quarterly report')
  expect(json).toContain('sleepy cat')
  expect(re.doc.meta.sections[0].page.width).toBe(11906) // A4 survived
  expect(JSON.stringify(re.doc.meta.headerFooters)).toContain('PAGE')

  // PDF rendering through the print pipeline
  await page.evaluate(() => document.documentElement.classList.add('is-printing'))
  await page.emulateMedia({ media: 'print' })
  const pdf = await page.pdf({ preferCSSPageSize: true, printBackground: true })
  fs.writeFileSync(path.join(out, 'e2e-export.pdf'), pdf)
  expect(pdf.subarray(0, 5).toString()).toBe('%PDF-')
  expect(errors).toEqual([])
})

test('import DOCX shows the compatibility report and renders tables, images, comments', async ({ page }) => {
  await signUp(page, `importer${Date.now()}@example.com`)
  await page.locator('input[type=file]').setInputFiles(path.resolve('tests/fixtures/complex.docx'))
  await expect(page.getByRole('dialog', { name: 'Import report' })).toBeVisible({ timeout: 20000 })
  await expect(page.getByRole('dialog')).toContainText('Footnotes')
  await page.getByRole('button', { name: 'Open document' }).click()
  await expect(page.locator('.sc-doc table').first()).toBeVisible({ timeout: 20000 })
  // ProseMirror adds invisible <img class="ProseMirror-separator"> cursor helpers; count real images only
  await expect(page.locator('.sc-doc img:not(.ProseMirror-separator)')).toHaveCount(1)
  await expect(page.locator('.sc-doc .sc-comment')).toHaveCount(1)
  await expect(page.locator('.sc-doc ins.sc-ins')).toContainText('inserted words')
  await expect(page.locator('.sc-page-header').first()).toContainText('Complex Fixture')
  await expect(page.locator('.sc-page-footer').first()).toContainText('Page 1 of')
  // landscape section renders a wider page
  const widths = await page.locator('.sc-page').evaluateAll((els) => els.map((e) => Math.round(e.getBoundingClientRect().width)))
  expect(Math.max(...widths)).toBeGreaterThan(Math.min(...widths))
  // comments panel lists the imported comment
  await page.getByRole('button', { name: 'Comments' }).click()
  await expect(page.locator('.sc-comments')).toContainText('Please check this sentence.')
  // accept all tracked changes
  await page.getByRole('tab', { name: 'Review' }).click()
  await page.getByRole('button', { name: 'Accept all' }).click()
  await expect(page.locator('.sc-doc ins.sc-ins')).toHaveCount(0)
  await expect(body(page)).not.toContainText('deleted words')
})

test('two users collaborate in real time; viewer cannot edit', async ({ browser }) => {
  const a = await browser.newContext()
  const b = await browser.newContext()
  const pa = await a.newPage()
  const pb = await b.newPage()
  const emailB = `collab${Date.now()}@example.com`
  await signUp(pb, emailB)
  await signUp(pa, `owner${Date.now()}@example.com`)
  await pa.getByRole('button', { name: /Blank document/ }).click()
  await body(pa).click()
  await pa.keyboard.type('Hello from A')
  await pa.getByRole('button', { name: 'Share', exact: true }).click()
  await pa.getByLabel('Email').fill(emailB)
  await pa.getByLabel('Access level').selectOption('viewer')
  await pa.getByRole('button', { name: 'Share', exact: true }).last().click()
  await expect(pa.locator('.sc-share-list')).toContainText(emailB)
  await pa.getByRole('button', { name: 'Done' }).click()
  await pb.goto(pa.url())
  await expect(body(pb)).toContainText('Hello from A', { timeout: 15000 })
  await expect(pb.getByText('View only')).toBeVisible()
  await expect(body(pb)).toHaveAttribute('contenteditable', 'false')
  await body(pa).click()
  await pa.keyboard.press('ControlOrMeta+End')
  await pa.keyboard.type(' — live update')
  await expect(body(pb)).toContainText('live update', { timeout: 10000 })
  await a.close()
  await b.close()
})

test('security: imported HTML is sanitized', async ({ page }) => {
  await signUp(page, `sec${Date.now()}@example.com`)
  const evil = path.join(out, 'evil.html')
  fs.writeFileSync(evil, '<p>Safe text</p><img src=x onerror="window.__pwned=1"><script>window.__pwned=1</script><a href="javascript:window.__pwned=1">click</a>')
  await page.locator('input[type=file]').setInputFiles(evil)
  await expect(body(page)).toContainText('Safe text', { timeout: 15000 })
  expect(await page.evaluate(() => (window as any).__pwned)).toBeUndefined()
  expect(await page.locator('.sc-doc a[href^="javascript"]').count()).toBe(0)
})

test('Editing tools: format painter, change case, grow font, zoom slider', async ({ page }) => {
  await page.goto('/')
  await expect(body(page)).toBeVisible({ timeout: 20000 })
  await body(page).click()
  await page.keyboard.press('ControlOrMeta+b')
  await page.keyboard.type('Source')
  await page.keyboard.press('ControlOrMeta+b')
  await page.keyboard.type(' plain words here')
  // copy bold from "Source" to "plain"
  await page.locator('.sc-doc strong').first().click()
  await page.getByRole('button', { name: 'Format painter' }).click()
  await page.locator('.sc-doc p').first().dblclick({ position: { x: 90, y: 8 } })
  await expect(page.locator('.sc-doc strong')).toHaveCount(2)
  // change case
  await page.locator('.sc-doc p').first().click()
  await page.keyboard.press('ControlOrMeta+a')
  await page.getByRole('button', { name: 'Change case' }).click()
  await page.getByRole('menuitem', { name: 'UPPERCASE' }).click()
  await expect(body(page)).toContainText('SOURCE PLAIN WORDS HERE')
  // grow font twice from 11pt → 14pt
  await page.keyboard.press('ControlOrMeta+a')
  await page.getByRole('button', { name: 'Grow font' }).click()
  await page.getByRole('button', { name: 'Grow font' }).click()
  await expect(page.getByLabel('Font size')).toHaveValue('14')
  // zoom slider
  await page.getByLabel('Zoom level').fill('150')
  await expect(page.locator('.sc-zoom-pct')).toHaveText('150%')
})

test('shapes, equations, citations, bibliography, proofing and mail merge work end to end', async ({ page }) => {
  await page.goto('/')
  await expect(body(page)).toBeVisible({ timeout: 20000 })
  const before = page.url()
  await page.getByRole('button', { name: 'File menu' }).click()
  await page.getByRole('menuitem', { name: /New blank document/ }).click()
  // wait until the new document has actually opened before typing into it
  await expect(page).not.toHaveURL(before, { timeout: 20000 })
  await expect(body(page)).toBeVisible({ timeout: 20000 })
  await expect(page.getByLabel('Document title')).toHaveValue('Document', { timeout: 15000 })
  await body(page).click()
  await page.keyboard.type('Dear ')

  // merge field from a CSV
  await page.getByRole('tab', { name: 'Mailings' }).click()
  await page.getByRole('toolbar').getByRole('button', { name: 'Mail merge' }).click()
  await page.locator('.sc-steps input[type=file]').setInputFiles({ name: 'people.csv', mimeType: 'text/csv', buffer: Buffer.from('Name,City\nAda,London\nGrace,Arlington\n') })
  await expect(page.getByText('2 recipients', { exact: false }).first()).toBeVisible()
  await body(page).click()
  await page.keyboard.press('End')
  await page.getByRole('button', { name: '«Name»' }).click()
  await page.keyboard.type(' wrote this this sentence. ')

  // shape and equation
  await page.getByRole('tab', { name: 'Insert' }).click()
  await page.getByRole('button', { name: 'Shapes' }).click()
  await page.getByRole('menuitem', { name: 'Oval' }).click()
  await page.locator('img.sc-shape').dblclick()
  await page.getByLabel('Text in shape').fill('Hi')
  await page.getByRole('button', { name: 'Save' }).click()
  await page.getByRole('button', { name: 'Equation' }).click()
  await page.getByRole('dialog').getByRole('textbox').fill('E = mc^2')
  await expect(page.getByRole('dialog').locator('math msup')).toBeVisible()
  await page.getByRole('button', { name: 'Insert', exact: true }).click()
  await expect(page.locator('.sc-doc .sc-equation math')).toBeVisible()
  await page.getByRole('button', { name: 'Chart', exact: true }).click()
  await page.getByRole('radio', { name: 'Pie' }).click()
  await page.getByRole('textbox', { name: 'Chart title' }).fill('Budget split')
  await page.getByRole('dialog').getByRole('button', { name: 'Save', exact: true }).click()
  await expect(page.locator('.sc-doc img.sc-chart')).toHaveAttribute('alt', 'Budget split')

  // citation and bibliography
  await page.getByRole('tab', { name: 'References' }).click()
  await page.getByRole('button', { name: 'Insert citation' }).click()
  await page.getByRole('button', { name: 'Add a source' }).click()
  await page.getByLabel('Author 1 last name').fill('Smith')
  await page.getByLabel('Author 1 first name').fill('Jane')
  await page.getByRole('dialog').getByLabel('Title', { exact: true }).fill('Writing Well')
  await page.getByLabel('Year').fill('2020')
  await page.getByRole('button', { name: 'Save sources' }).click()
  await page.locator('.sc-pick-list input').first().check()
  await page.getByRole('button', { name: 'Insert', exact: true }).click()
  await expect(body(page)).toContainText('(Smith, 2020)')
  await page.getByRole('button', { name: 'Bibliography' }).click()
  await expect(page.locator('.sc-bibliography')).toContainText('Smith, J. (2020). Writing Well.')
  await page.getByLabel('Citation style').selectOption('MLA')
  await expect(page.locator('.sc-bibliography')).toContainText('Works Cited')
  await expect(body(page)).toContainText('(Smith)')

  // footnote: typed in the notes panel, shown at the bottom of the page
  await body(page).getByText('wrote this this sentence').click()
  await page.keyboard.press('End')
  await page.getByRole('tab', { name: 'References' }).click()
  await page.getByRole('button', { name: 'Insert footnote' }).click()
  await expect(page.getByLabel('Footnote 1 text')).toBeFocused()
  await page.keyboard.type('A footnote typed in SCRIBE.')
  await expect(page.locator('.sc-page .sc-page-notes')).toContainText('A footnote typed in SCRIBE.')

  // a table taller than a page continues on the next page
  await body(page).locator(':scope > p').last().click()
  await page.getByRole('tab', { name: 'Insert' }).click()
  await page.getByRole('button', { name: 'Table', exact: true }).click()
  for (let i = 0; i < 150; i++) {
    await page.keyboard.type(`c${i}`)
    await page.keyboard.press('Tab')
  }
  await expect(page.locator('tr.sc-page-spacer-row').first()).toBeAttached()

  // proofing flags the repeated word and fixes it
  await page.getByRole('navigation').getByRole('button', { name: 'Proofing' }).click()
  await page.getByLabel('Check grammar and style').check()
  await expect(page.locator('.sc-doc .sc-proof').first()).toBeVisible()
  await page.getByRole('button', { name: 'Change to “this”' }).click()
  await expect(body(page)).toContainText('wrote this sentence')

  // export validates and keeps everything
  await expect(page.getByRole('status').filter({ hasText: 'All changes saved' })).toBeVisible({ timeout: 15000 })
  const [dl] = await Promise.all([page.waitForEvent('download'), page.getByRole('button', { name: 'Download as DOCX' }).click()])
  const file = path.join(out, 'e2e-features.docx')
  await dl.saveAs(file)
  const bytes = new Uint8Array(fs.readFileSync(file))
  expect(await validatePackage(bytes)).toEqual([])
  const re = await importDocx(bytes)
  const json = JSON.stringify(re.doc.content)
  expect(json).toContain('"type":"shape"')
  expect(json).toContain('MERGEFIELD Name')
  expect(json).toContain('CITATION')
  expect(json).toContain('"type":"bibliography"')
  expect(json).toContain('"kind":"equation"')
  expect(json).toContain('"title":"Budget split"')
  expect(re.doc.meta.sources?.[0].title).toBe('Writing Well')
  expect(re.doc.meta.citationStyle).toBe('MLA')
  expect(re.doc.meta.footnotes.map((n) => n.text)).toContain('A footnote typed in SCRIBE.')

  // merge to DOCX: one page per recipient
  await page.getByRole('navigation').getByRole('button', { name: 'Mail merge' }).click()
  const [merged] = await Promise.all([page.waitForEvent('download'), page.locator('.sc-steps').getByRole('button', { name: 'DOCX' }).click()])
  const mfile = path.join(out, 'e2e-merged.docx')
  await merged.saveAs(mfile)
  const mbytes = new Uint8Array(fs.readFileSync(mfile))
  expect(await validatePackage(mbytes)).toEqual([])
  const mjson = JSON.stringify((await importDocx(mbytes)).doc.content)
  expect(mjson).toContain('Dear Ada')
  expect(mjson).toContain('Dear Grace')
})

test('two-column sections are laid out side by side and stay editable', async ({ page }) => {
  await page.goto('/')
  await expect(body(page)).toBeVisible({ timeout: 20000 })
  const before = page.url()
  await page.getByRole('button', { name: 'File menu' }).click()
  await page.getByRole('menuitem', { name: /New blank document/ }).click()
  await expect(page).not.toHaveURL(before, { timeout: 20000 })
  await body(page).click()
  for (let i = 1; i <= 34; i++) {
    await page.keyboard.type(`Paragraph ${i}. Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor.`)
    await page.keyboard.press('Enter')
  }
  await page.getByRole('tab', { name: 'Layout' }).click()
  await page.getByRole('button', { name: 'Columns' }).click()
  await page.getByRole('menuitem', { name: 'Two columns' }).click()
  await expect(page.getByText('2 columns', { exact: true })).toBeVisible()
  const first = body(page).locator(':scope > p', { hasText: 'Paragraph 1.' })
  // the last paragraphs move into the second column, beside the first ones
  await expect
    .poll(async () => {
      const a = await first.boundingBox()
      const lefts = await body(page).locator(':scope > p').evaluateAll((ps) => ps.map((p) => Math.round(p.getBoundingClientRect().left)))
      return a ? new Set(lefts).size : 0
    })
    .toBe(2)
  const p1 = (await first.boundingBox())!
  const later = body(page).locator(':scope > p').filter({ hasNot: page.locator('xx') })
  const boxes = await later.evaluateAll((ps) => ps.map((p) => [p.getBoundingClientRect().left, p.getBoundingClientRect().top]))
  const col2 = boxes.find(([l]) => l > p1.x + 50)!
  expect(Math.abs(col2[1] - p1.y)).toBeLessThan(4) // column 2 starts level with column 1
  // typing in the second column edits the right paragraph
  const p30 = body(page).locator(':scope > p', { hasText: 'Paragraph 30.' })
  await p30.scrollIntoViewIfNeeded()
  const b30 = (await p30.boundingBox())!
  await page.mouse.click(b30.x + 2, b30.y + 6)
  await page.keyboard.press('Home')
  await page.keyboard.type('Edited ')
  await expect(p30).toContainText('Edited Paragraph 30.')
})

test('table of contents lists headings with page numbers, jumps to them, and symbols insert', async ({ page }) => {
  await page.goto('/')
  await expect(body(page)).toBeVisible({ timeout: 20000 })
  const before = page.url()
  await page.getByRole('button', { name: 'File menu' }).click()
  await page.getByRole('menuitem', { name: /New blank document/ }).click()
  await expect(page).not.toHaveURL(before, { timeout: 20000 })
  await body(page).click()
  await page.keyboard.press('Enter')
  for (const [i, title] of ['Introduction', 'Method', 'Results'].entries()) {
    await page.keyboard.type(title)
    await page.keyboard.press('ControlOrMeta+Alt+1')
    await page.keyboard.press('Enter')
    for (let k = 0; k < (i === 0 ? 40 : 4); k++) {
      await page.keyboard.type(`Body text ${k}.`)
      await page.keyboard.press('Enter')
    }
  }
  // insert the table at the top
  await body(page).locator(':scope > p').first().click()
  await page.keyboard.press('ControlOrMeta+Home')
  await page.getByRole('tab', { name: 'References' }).click()
  await page.getByRole('button', { name: 'Table of contents' }).click()
  const toc = page.locator('.sc-doc .sc-toc')
  await expect(toc.locator('.sc-toc-text')).toHaveText(['Introduction', 'Method', 'Results'])
  // Introduction is on page 1; the long body text pushes Method to page 2
  await expect.poll(async () => toc.locator('.sc-toc-page').allTextContents()).toEqual(['1', '2', '2'])
  await toc.locator('.sc-toc-entry', { hasText: 'Results' }).click()
  await expect.poll(async () => page.evaluate(() => window.getSelection()?.anchorNode?.parentElement?.closest('h1')?.textContent)).toBe('Results')
  // symbol picker: inserts at the cursor (the start of the heading after the jump)
  await page.getByRole('tab', { name: 'Insert' }).click()
  await page.getByRole('button', { name: 'Symbol', exact: true }).click()
  await page.getByRole('tab', { name: 'Currency' }).click()
  await page.getByRole('button', { name: /^€ \(U\+20AC\)$/ }).click()
  await page.getByRole('button', { name: 'Done' }).click()
  await expect(body(page).locator('h1', { hasText: '€Results' })).toBeVisible()
})

test('command search finds and runs actions', async ({ page }) => {
  await page.goto('/')
  await expect(body(page)).toBeVisible({ timeout: 20000 })
  await body(page).click()
  await page.keyboard.press('ControlOrMeta+Shift+p')
  const input = page.getByRole('combobox', { name: 'Search commands' })
  await input.fill('symbol')
  await page.keyboard.press('Enter')
  await expect(page.getByRole('dialog', { name: 'Insert symbol' })).toBeVisible()
  await page.keyboard.press('Escape')
  await page.getByRole('button', { name: 'Search commands' }).click()
  await input.fill('two col')
  await expect(page.getByRole('option', { name: /Two columns/ })).toBeVisible()
  await page.keyboard.press('Escape')
  await expect(input).toHaveCount(0)
})

test('links written as paths open the right screen, and missing files are real 404s', async ({ page, request }) => {
  await page.goto('/documents')
  await expect(page.getByRole('heading', { name: 'Start a document' })).toBeVisible({ timeout: 20000 })
  expect(new URL(page.url()).pathname).toBe('/')
  expect(new URL(page.url()).hash).toBe('#/documents')
  await page.goto('/some/old/link')
  await expect(body(page)).toBeVisible({ timeout: 20000 })
  expect(new URL(page.url()).pathname).toBe('/')
  expect((await request.get('/assets/missing-file.js')).status()).toBe(404)
  expect((await request.get('/api/missing')).status()).toBe(404)
  expect((await request.get('/favicon.svg')).status()).toBe(200)
})
