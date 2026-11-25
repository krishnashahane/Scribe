import type { PMNode, ScribeDocument } from './types'
import { defaultPage, emptyDocument } from './types'

const p = (text = '', attrs?: Record<string, any>, marks?: any[]): PMNode => ({
  type: 'paragraph',
  ...(attrs ? { attrs } : {}),
  ...(text ? { content: [{ type: 'text', text, ...(marks ? { marks } : {}) }] } : {}),
})
const h = (level: number, text: string): PMNode => ({ type: 'heading', attrs: { level }, content: [{ type: 'text', text }] })
const bullets = (items: string[]): PMNode => ({ type: 'bulletList', content: items.map((t) => ({ type: 'listItem', content: [p(t)] })) })
const bold = [{ type: 'bold' }]

export interface Template {
  id: string
  name: string
  description: string
  build: () => ScribeDocument
}

const withContent = (content: PMNode[], tweak?: (d: ScribeDocument) => void): ScribeDocument => {
  const d = emptyDocument()
  d.content = { type: 'doc', content }
  tweak?.(d)
  return d
}

const pageNumberFooter = (d: ScribeDocument) => {
  d.meta.headerFooters['new:footer-1'] = {
    type: 'doc',
    content: [{ type: 'paragraph', attrs: { textAlign: 'center' }, content: [{ type: 'text', text: 'Page ' }, { type: 'field', attrs: { instr: 'PAGE', result: '1' } }, { type: 'text', text: ' of ' }, { type: 'field', attrs: { instr: 'NUMPAGES', result: '1' } }] }],
  }
  d.meta.sections[0].footerDefault = 'new:footer-1'
}

export const TEMPLATES: Template[] = [
  { id: 'blank', name: 'Blank document', description: 'Letter size, 1″ margins', build: () => emptyDocument() },
  {
    id: 'blank-a4',
    name: 'Blank A4',
    description: 'A4 size, 2.54 cm margins',
    build: () => withContent([p()], (d) => (d.meta.sections[0].page = defaultPage('A4'))),
  },
  {
    id: 'letter',
    name: 'Formal letter',
    description: 'Business letter layout',
    build: () =>
      withContent([
        p('Your Name', undefined, bold),
        p('Street Address'),
        p('City, State ZIP'),
        p(),
        p(new Date().toLocaleDateString(undefined, { year: 'numeric', month: 'long', day: 'numeric' })),
        p(),
        p('Recipient Name'),
        p('Company'),
        p('Address'),
        p(),
        p('Dear Recipient,'),
        p('Write the opening paragraph of your letter here. State the purpose of the letter clearly.'),
        p('Use the body to give supporting details. Keep paragraphs focused and concise.'),
        p('Sincerely,'),
        p(),
        p('Your Name'),
      ]),
  },
  {
    id: 'report',
    name: 'Report',
    description: 'Title, sections and page numbers',
    build: () =>
      withContent(
        [
          { type: 'heading', attrs: { level: 1, textAlign: 'center' }, content: [{ type: 'text', text: 'Report Title' }] },
          p('Prepared by · Date', { textAlign: 'center' }),
          h(2, 'Executive summary'),
          p('Summarize the purpose, findings and recommendations of this report.'),
          h(2, 'Background'),
          p('Describe the context and the problem being addressed.'),
          h(2, 'Findings'),
          bullets(['First finding', 'Second finding', 'Third finding']),
          h(2, 'Recommendations'),
          { type: 'orderedList', attrs: { start: 1 }, content: ['First recommendation', 'Second recommendation'].map((t) => ({ type: 'listItem', content: [p(t)] })) },
          h(2, 'Appendix'),
          {
            type: 'table',
            content: [
              { type: 'tableRow', content: ['Metric', 'Value', 'Notes'].map((t) => ({ type: 'tableHeader', attrs: { background: '#e8eef8' }, content: [p(t, undefined, bold)] })) },
              { type: 'tableRow', content: ['', '', ''].map(() => ({ type: 'tableCell', content: [p()] })) },
            ],
          },
        ],
        pageNumberFooter,
      ),
  },
  {
    id: 'meeting',
    name: 'Meeting notes',
    description: 'Agenda, notes and action items',
    build: () =>
      withContent([
        h(1, 'Meeting notes'),
        p('Date: · Attendees:'),
        h(2, 'Agenda'),
        bullets(['Topic one', 'Topic two']),
        h(2, 'Notes'),
        p(),
        h(2, 'Action items'),
        bullets(['Owner — action — due date']),
      ]),
  },
  {
    id: 'resume',
    name: 'Résumé',
    description: 'Clean one-page résumé',
    build: () =>
      withContent(
        [
          { type: 'heading', attrs: { level: 1 }, content: [{ type: 'text', text: 'Your Name' }] },
          p('email@example.com · +1 555 0100 · City'),
          h(2, 'Experience'),
          p('Job Title — Company', undefined, bold),
          p('2022 – Present'),
          bullets(['Accomplishment with a measurable result', 'Another accomplishment']),
          h(2, 'Education'),
          p('Degree — University', undefined, bold),
          h(2, 'Skills'),
          p('Skill · Skill · Skill'),
        ],
        (d) => {
          d.meta.sections[0].page = { ...defaultPage('Letter'), marginTop: 1080, marginBottom: 1080, marginLeft: 1080, marginRight: 1080 }
        },
      ),
  },
]
