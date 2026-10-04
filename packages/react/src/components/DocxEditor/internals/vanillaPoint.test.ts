/**
 * vanillaText, Pelnora .9 — the point mapper behind a pure insertion at an
 * offset (proposeChangeAt with length 0) and the formatting marks inserted
 * text inherits.
 */

import { describe, expect, it } from 'bun:test';
import { Schema } from 'prosemirror-model';

import { inheritedFormattingMarks, mapVanillaOffsetToPmPoint } from './vanillaText';

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: { group: 'block', content: 'inline*' },
    text: { group: 'inline' },
  },
  marks: {
    insertion: {},
    deletion: {},
    comment: {},
    hyperlink: {},
    bold: {},
    fontFamily: { attrs: { ascii: { default: null } } },
  },
});

type Run = {
  text: string;
  ins?: boolean;
  del?: boolean;
  bold?: boolean;
  font?: string;
  comment?: boolean;
};

function docFromRuns(runs: Run[]) {
  const nodes = runs.map((r) => {
    const marks = [];
    if (r.ins) marks.push(schema.marks.insertion!.create());
    if (r.del) marks.push(schema.marks.deletion!.create());
    if (r.bold) marks.push(schema.marks.bold!.create());
    if (r.font) marks.push(schema.marks.fontFamily!.create({ ascii: r.font }));
    if (r.comment) marks.push(schema.marks.comment!.create());
    return schema.text(r.text, marks);
  });
  const doc = schema.node('doc', null, [schema.node('paragraph', null, nodes)]);
  return { doc, from: 0, to: doc.child(0).nodeSize };
}

describe('mapVanillaOffsetToPmPoint', () => {
  it('maps every point 0 … length, the end of the paragraph included', () => {
    // "14 dni": paragraph content starts at PM 1.
    const { doc, from, to } = docFromRuns([{ text: '14 dni' }]);
    expect(mapVanillaOffsetToPmPoint(doc, from, to, 0)).toBe(1);
    expect(mapVanillaOffsetToPmPoint(doc, from, to, 3)).toBe(4); // "14 |dni"
    expect(mapVanillaOffsetToPmPoint(doc, from, to, 6)).toBe(7); // end
  });

  it('at a run boundary the point stays in the run before it', () => {
    // "AB" + "CD": offset 2 is the end of "AB" (PM 3), not the start of "CD".
    const { doc, from, to } = docFromRuns([{ text: 'AB' }, { text: 'CD', bold: true }]);
    expect(mapVanillaOffsetToPmPoint(doc, from, to, 2)).toBe(3);
  });

  it('skips insertion-marked runs, like the span mapper', () => {
    const { doc, from, to } = docFromRuns([
      { text: 'AB' },
      { text: 'X', ins: true },
      { text: 'CD' },
    ]);
    // offset 3 ("C|D") lies past the inserted "X".
    expect(mapVanillaOffsetToPmPoint(doc, from, to, 3)).toBe(5);
  });

  it('returns null outside the text or for an empty paragraph', () => {
    const { doc, from, to } = docFromRuns([{ text: 'abc' }]);
    expect(mapVanillaOffsetToPmPoint(doc, from, to, -1)).toBeNull();
    expect(mapVanillaOffsetToPmPoint(doc, from, to, 4)).toBeNull();
    const empty = schema.node('doc', null, [schema.node('paragraph')]);
    expect(mapVanillaOffsetToPmPoint(empty, 0, empty.child(0).nodeSize, 0)).toBeNull();
  });
});

describe('inheritedFormattingMarks', () => {
  const names = (marks: readonly { type: { name: string } }[]) =>
    marks.map((m) => m.type.name).sort();

  it('takes the formatting of the text it lands in, never change, comment or link marks', () => {
    const { doc } = docFromRuns([
      { text: 'Ab', bold: true, font: 'Times New Roman', comment: true },
      { text: 'cd', del: true, font: 'Arial' },
    ]);
    // after PM 1: the first character of "Ab".
    expect(names(inheritedFormattingMarks(doc, 1, 'after'))).toEqual(['bold', 'fontFamily']);
    expect(
      inheritedFormattingMarks(doc, 1, 'after').find((m) => m.type.name === 'fontFamily')!.attrs
        .ascii
    ).toBe('Times New Roman');
    // before PM 5: the struck "cd" → its font only, not the deletion mark.
    expect(names(inheritedFormattingMarks(doc, 5, 'before'))).toEqual(['fontFamily']);
  });

  it('falls back to the other side at a paragraph edge', () => {
    const { doc } = docFromRuns([{ text: 'ab', bold: true }]);
    expect(names(inheritedFormattingMarks(doc, 1, 'before'))).toEqual(['bold']); // start: after
    expect(names(inheritedFormattingMarks(doc, 3, 'after'))).toEqual(['bold']); // end: before
  });
});
