/**
 * trackedInsert — Pelnora H0-3 (b): `insertTracked` builds ONE transaction
 * that inserts block nodes as a tracked insertion: every text node gets the
 * `insertion` mark (author/date, one revision id per paragraph) and every
 * paragraph gets its own paragraph-mark revision (`paragraphMarkChange`), so
 * the core serializes `w:ins` around the runs and `w:pPr/w:rPr/w:ins` on the
 * paragraph mark — what Word writes for a paragraph typed with Track Changes.
 */

import { describe, expect, it } from 'bun:test';
import { Schema, type Node as PMNode } from 'prosemirror-model';
import { EditorState } from 'prosemirror-state';

import { buildTrackedInsertTransaction } from './trackedInsert';

const schema = new Schema({
  nodes: {
    doc: { content: 'block+' },
    paragraph: {
      group: 'block',
      content: 'inline*',
      attrs: {
        paraId: { default: null },
        styleId: { default: null },
        paragraphMarkChange: { default: null },
      },
    },
    text: { group: 'inline' },
  },
  marks: {
    bold: {},
    insertion: {
      attrs: { revisionId: { default: 0 }, author: { default: '' }, date: { default: null } },
      inclusive: false,
    },
    deletion: {
      attrs: { revisionId: { default: 0 }, author: { default: '' }, date: { default: null } },
      inclusive: false,
    },
  },
});

const AUTHOR = 'Asystent Pelnory';
const DATE = '2026-09-02T10:00:00Z';

function baseState(): EditorState {
  const doc = schema.node('doc', null, [
    schema.node('paragraph', { paraId: 'P1' }, [schema.text('Pierwszy')]),
    schema.node('paragraph', { paraId: 'P2' }, [schema.text('Drugi')]),
  ]);
  return EditorState.create({ doc, schema });
}

function newBlocks(): PMNode[] {
  return [
    schema.node('paragraph', { styleId: 'Heading2' }, [
      schema.text('Nowy '),
      schema.text('nagłówek', [schema.marks.bold.create()]),
    ]),
    schema.node('paragraph', null, [schema.text('Treść akapitu.')]),
  ];
}

function makeAllocator(start = 100): { next: () => number; calls: number } {
  const box = { calls: 0, next: () => 0 };
  let id = start;
  box.next = () => {
    box.calls++;
    return id++;
  };
  return box;
}

function paraTexts(doc: PMNode): string[] {
  const out: string[] = [];
  doc.forEach((n) => out.push(n.textContent));
  return out;
}

describe('buildTrackedInsertTransaction', () => {
  it('inserts after the anchor paragraph, marks runs and paragraph marks with author/date', () => {
    const state = baseState();
    const alloc = makeAllocator();
    const tr = buildTrackedInsertTransaction(
      state,
      { paraId: 'P1', position: 'after', nodes: newBlocks(), author: AUTHOR, date: DATE },
      alloc.next
    );
    expect(tr).not.toBeNull();
    const doc = state.apply(tr!).doc;

    expect(paraTexts(doc)).toEqual(['Pierwszy', 'Nowy nagłówek', 'Treść akapitu.', 'Drugi']);

    const heading = doc.child(1);
    const bodyPara = doc.child(2);
    expect(heading.attrs.styleId).toBe('Heading2');

    // Every text node of the inserted paragraphs carries the insertion mark.
    const insertionAttrs: Array<Record<string, unknown>> = [];
    for (const p of [heading, bodyPara]) {
      p.forEach((n) => {
        const m = n.marks.find((mark) => mark.type.name === 'insertion');
        expect(m).toBeDefined();
        insertionAttrs.push(m!.attrs);
      });
    }
    for (const a of insertionAttrs) {
      expect(a.author).toBe(AUTHOR);
      expect(a.date).toBe(DATE);
    }
    // One revision id per paragraph for its runs.
    expect(insertionAttrs[0]!.revisionId).toBe(insertionAttrs[1]!.revisionId);
    expect(insertionAttrs[2]!.revisionId).not.toBe(insertionAttrs[0]!.revisionId);

    // Bold survives next to the insertion mark.
    expect(heading.child(1).marks.some((m) => m.type.name === 'bold')).toBe(true);

    // Paragraph-mark revision, own id, same author/date.
    for (const p of [heading, bodyPara]) {
      const change = p.attrs.paragraphMarkChange as {
        type: string;
        info: { id: number; author: string; date: string };
      };
      expect(change.type).toBe('insertion');
      expect(change.info.author).toBe(AUTHOR);
      expect(change.info.date).toBe(DATE);
    }

    // All allocated ids are distinct: 2 paragraphs × (runs + paragraph mark).
    const ids = new Set<number>([
      ...insertionAttrs.map((a) => a.revisionId as number),
      (heading.attrs.paragraphMarkChange as { info: { id: number } }).info.id,
      (bodyPara.attrs.paragraphMarkChange as { info: { id: number } }).info.id,
    ]);
    expect(ids.size).toBe(4);
    expect(alloc.calls).toBe(4);

    // Existing paragraphs are untouched (no marks, same ids).
    expect(doc.child(0).attrs.paraId).toBe('P1');
    expect(doc.child(3).attrs.paraId).toBe('P2');
    doc.child(0).forEach((n) => expect(n.marks).toHaveLength(0));
  });

  it("position 'before' inserts in front of the anchor paragraph", () => {
    const state = baseState();
    const tr = buildTrackedInsertTransaction(
      state,
      { paraId: 'P2', position: 'before', nodes: newBlocks(), author: AUTHOR, date: DATE },
      makeAllocator().next
    );
    expect(paraTexts(state.apply(tr!).doc)).toEqual([
      'Pierwszy',
      'Nowy nagłówek',
      'Treść akapitu.',
      'Drugi',
    ]);
  });

  it('defaults to after', () => {
    const state = baseState();
    const tr = buildTrackedInsertTransaction(
      state,
      { paraId: 'P2', nodes: newBlocks(), author: AUTHOR, date: DATE },
      makeAllocator().next
    );
    expect(paraTexts(state.apply(tr!).doc)).toEqual([
      'Pierwszy',
      'Drugi',
      'Nowy nagłówek',
      'Treść akapitu.',
    ]);
  });

  it('paraId null appends at the end of the document', () => {
    const state = baseState();
    const tr = buildTrackedInsertTransaction(
      state,
      { paraId: null, nodes: newBlocks(), author: AUTHOR, date: DATE },
      makeAllocator().next
    );
    expect(paraTexts(state.apply(tr!).doc)).toEqual([
      'Pierwszy',
      'Drugi',
      'Nowy nagłówek',
      'Treść akapitu.',
    ]);
  });

  it('returns null for an unknown paraId or no nodes, without touching the doc', () => {
    const state = baseState();
    expect(
      buildTrackedInsertTransaction(
        state,
        { paraId: 'NOPE', nodes: newBlocks(), author: AUTHOR, date: DATE },
        makeAllocator().next
      )
    ).toBeNull();
    expect(
      buildTrackedInsertTransaction(
        state,
        { paraId: 'P1', nodes: [], author: AUTHOR, date: DATE },
        makeAllocator().next
      )
    ).toBeNull();
  });

  it('fills in the date when the caller omits it (ISO 8601)', () => {
    const state = baseState();
    const tr = buildTrackedInsertTransaction(
      state,
      { paraId: 'P1', nodes: newBlocks(), author: AUTHOR },
      makeAllocator().next
    );
    const inserted = state.apply(tr!).doc.child(1);
    const mark = inserted.child(0).marks.find((m) => m.type.name === 'insertion')!;
    expect(typeof mark.attrs.date).toBe('string');
    expect(() => new Date(mark.attrs.date as string).toISOString()).not.toThrow();
    expect((inserted.attrs.paragraphMarkChange as { info: { date: string } }).info.date).toBe(
      mark.attrs.date
    );
  });

  it('does not add the mark to text that already carries a deletion (defensive)', () => {
    const state = baseState();
    const del = schema.marks.deletion.create({ revisionId: 1, author: 'X', date: DATE });
    const nodes = [
      schema.node('paragraph', null, [schema.text('gone', [del]), schema.text(' kept')]),
    ];
    const tr = buildTrackedInsertTransaction(
      state,
      { paraId: 'P1', nodes, author: AUTHOR, date: DATE },
      makeAllocator().next
    );
    const p = state.apply(tr!).doc.child(1);
    expect(p.child(0).marks.map((m) => m.type.name)).toEqual(['deletion']);
    expect(p.child(1).marks.map((m) => m.type.name)).toEqual(['insertion']);
  });
});
