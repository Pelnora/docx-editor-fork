/**
 * vanillaText — mapVanillaOffsetToPm (Pelnora c11f offset addressing) +
 * findTextInPmParagraph regression. Offset addressing lets a specific
 * occurrence of a repeated phrase in ONE paragraph be targeted (search is
 * ambiguous there). Insertion-marked runs are excluded from the vanilla
 * coordinate space; deletion-marked runs are kept.
 */

import { describe, expect, it } from 'bun:test';
import { Schema } from 'prosemirror-model';

import { findTextInPmParagraph, mapVanillaOffsetToPm } from './vanillaText';

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { group: 'block', content: 'inline*' },
    text: { group: 'inline' },
  },
  marks: { insertion: {}, deletion: {} },
});

/** Single-paragraph doc from runs (text + optional insertion/deletion mark). */
function docFromRuns(runs: Array<{ text: string; ins?: boolean; del?: boolean }>) {
  const nodes = runs.map((r) => {
    const marks = [];
    if (r.ins) marks.push(schema.marks.insertion!.create());
    if (r.del) marks.push(schema.marks.deletion!.create());
    return schema.text(r.text, marks);
  });
  const doc = schema.node('doc', null, [schema.node('paragraph', null, nodes)]);
  return { doc, from: 0, to: doc.child(0).nodeSize };
}

describe('mapVanillaOffsetToPm', () => {
  it('maps each occurrence of a repeated phrase to a DISTINCT PM span (c11f)', () => {
    // "Hello RODO and RODO" — two "RODO" at offsets 6 and 15.
    const { doc, from, to } = docFromRuns([{ text: 'Hello RODO and RODO' }]);
    expect(mapVanillaOffsetToPm(doc, from, to, 6, 4)).toEqual({ from: 7, to: 11 });
    expect(mapVanillaOffsetToPm(doc, from, to, 15, 4)).toEqual({ from: 16, to: 20 });
    // The two spans are different — search ("RODO") could never disambiguate them.
  });

  it('EXCLUDES insertion-marked runs from the offset space (deletions kept)', () => {
    // vanilla = "AB" + "CD" = "ABCD"; the inserted "X" is not in vanilla but its
    // PM node still consumes a position, so offset 2 ("C") maps PAST it.
    const { doc, from, to } = docFromRuns([
      { text: 'AB' },
      { text: 'X', ins: true },
      { text: 'CD' },
    ]);
    expect(mapVanillaOffsetToPm(doc, from, to, 2, 1)).toEqual({ from: 4, to: 5 });
    // a deletion-marked run stays in the vanilla space
    const d = docFromRuns([{ text: 'AB' }, { text: 'CD', del: true }]);
    expect(mapVanillaOffsetToPm(d.doc, d.from, d.to, 2, 2)).toEqual({ from: 3, to: 5 });
  });

  it('returns null for an empty span or one outside the paragraph text', () => {
    const { doc, from, to } = docFromRuns([{ text: 'short' }]);
    expect(mapVanillaOffsetToPm(doc, from, to, 0, 0)).toBeNull();
    expect(mapVanillaOffsetToPm(doc, from, to, -1, 2)).toBeNull();
    expect(mapVanillaOffsetToPm(doc, from, to, 3, 99)).toBeNull(); // overruns
  });
});

describe('findTextInPmParagraph (regression after the refactor)', () => {
  it('finds a UNIQUE phrase', () => {
    const { doc, from, to } = docFromRuns([{ text: 'Termin wynosi 14 dni.' }]);
    expect(findTextInPmParagraph(doc, from, to, '14 dni')).toEqual({ from: 15, to: 21 });
  });

  it('rejects an AMBIGUOUS phrase (2+ occurrences) — null, the caller disambiguates', () => {
    const { doc, from, to } = docFromRuns([{ text: 'Hello RODO and RODO' }]);
    expect(findTextInPmParagraph(doc, from, to, 'RODO')).toBeNull();
  });

  it('is null for an absent phrase or empty search', () => {
    const { doc, from, to } = docFromRuns([{ text: 'abc' }]);
    expect(findTextInPmParagraph(doc, from, to, 'xyz')).toBeNull();
    expect(findTextInPmParagraph(doc, from, to, '')).toBeNull();
  });
});
