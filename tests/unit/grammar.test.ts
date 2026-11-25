import { describe, expect, it } from 'vitest'
import { checkText } from '../../src/features/grammar'

const rules = (t: string) => checkText(t).map((i) => `${i.rule}:${t.slice(i.start, i.end)}→${i.fix ?? ''}`)

describe('grammar checker', () => {
  it('finds common mistakes with fixes', () => {
    expect(rules('This is the the best plan.')).toEqual(['repeated:the the→the'])
    expect(rules('We  met yesterday.')).toEqual(['spaces:  → '])
    expect(rules('Hello , world.')).toEqual(['spaceBefore: ,→,'])
    expect(rules('Apples,pears and plums.')).toEqual(['spaceAfter:,→, '])
    expect(rules('It works. then it stops.')).toEqual(['capital:t→T'])
    expect(rules('Yesterday i went home.')).toEqual(['pronounI:i→I'])
    expect(rules('She ate a apple and an banana.')).toEqual(['article:a→an', 'article:an→a'])
    expect(rules('I could of done it.')).toEqual(['couldOf:could of→could have'])
    expect(rules('This is better then that.')).toEqual(['confused:better then→better than'])
    expect(rules('We left early in order to rest.')).toEqual(['wordy:in order to→to'])
    expect(rules('Wait.. what?')).toEqual(['doublePunct:..→.'])
  })
  it('stays quiet on correct text', () => {
    for (const ok of [
      'An hour later, a university student met an honest man at 10:30.',
      'The value is 1,000.5 and the ratio 3:2 holds.',
      'He had had enough. That that is, is.',
      'See e.g. the appendix, i.e. section 2.',
      'Wait... and then it happened.',
      'Visit https://example.com/a,b for details.',
      'Dr. smith arrived.',
      'A one-time offer for a user and a European.',
      'Use (i) and (ii) below.',
      'An MBA and an FBI agent met a UN delegate.',
      'short line',
    ])
      expect(checkText(ok), ok).toEqual([])
  })
})
