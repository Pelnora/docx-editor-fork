/**
 * Tracked block insertion (Pelnora H0-3 b).
 *
 * Builds ONE transaction that inserts block nodes as a Word-style tracked
 * insertion: every text node of an inserted paragraph gets the `insertion`
 * mark (author / date / one revision id per paragraph) and the paragraph gets
 * its own paragraph-mark revision (`paragraphMarkChange`, own id). The core
 * serializes those as `<w:ins>` around the runs and `w:pPr/w:rPr/w:ins` on the
 * paragraph mark — what Word writes for a paragraph typed with Track Changes.
 *
 * Pure: takes the editor state and returns the transaction (or null); the
 * caller dispatches. Revision ids come from `allocateId` so the shared
 * comment/revision id space stays consistent.
 */

import { Fragment, type Node as PMNode } from 'prosemirror-model';
import type { EditorState, Transaction } from 'prosemirror-state';
import { findParaIdRange } from './pmAnchors';

export interface TrackedInsertOptions {
  /** Anchor paragraph. `null` appends at the end of the document. */
  paraId: string | null;
  /** Where relative to the anchor. Default `'after'`. */
  position?: 'after' | 'before';
  /** Block nodes built with the editor schema (paragraphs, tables, …). */
  nodes: PMNode[];
  /** Revision author (`w:author`). */
  author: string;
  /** Revision date (`w:date`, ISO 8601). Defaults to now. */
  date?: string;
}

interface RevisionInfo {
  id: number;
  author: string;
  date: string;
}

function markInline(node: PMNode, state: EditorState, info: RevisionInfo): PMNode {
  if (!node.isText) return node;
  // Defensive: never re-mark text that already carries a tracked change.
  if (node.marks.some((m) => m.type.name === 'insertion' || m.type.name === 'deletion')) {
    return node;
  }
  const mark = state.schema.marks.insertion.create({
    revisionId: info.id,
    author: info.author,
    date: info.date,
  });
  return node.mark(mark.addToSet(node.marks));
}

function trackBlock(
  node: PMNode,
  state: EditorState,
  author: string,
  date: string,
  allocateId: () => number
): PMNode {
  if (node.type.name === 'paragraph') {
    const runsInfo: RevisionInfo = { id: allocateId(), author, date };
    const markInfo: RevisionInfo = { id: allocateId(), author, date };
    const children: PMNode[] = [];
    node.forEach((child) => children.push(markInline(child, state, runsInfo)));
    const attrs =
      node.type.spec.attrs && 'paragraphMarkChange' in node.type.spec.attrs
        ? { ...node.attrs, paragraphMarkChange: { type: 'insertion', info: markInfo } }
        : node.attrs;
    return node.type.create(attrs, children, node.marks);
  }
  if (node.isLeaf || node.childCount === 0) return node;
  // Containers (tables, rows, cells, …): recurse so nested paragraphs are tracked.
  const children: PMNode[] = [];
  node.forEach((child) => children.push(trackBlock(child, state, author, date, allocateId)));
  return node.copy(Fragment.from(children));
}

/**
 * Build the tracked-insert transaction. Returns `null` (and touches nothing)
 * when the anchor paragraph is unknown, `nodes` is empty, the schema has no
 * `insertion` mark, or the insertion would be invalid at that position.
 */
export function buildTrackedInsertTransaction(
  state: EditorState,
  options: TrackedInsertOptions,
  allocateId: () => number
): Transaction | null {
  if (!state.schema.marks.insertion) return null;
  if (!options.nodes || options.nodes.length === 0) return null;

  let pos: number;
  if (options.paraId === null) {
    pos = state.doc.content.size;
  } else {
    const range = findParaIdRange(state.doc, options.paraId);
    if (!range) return null;
    pos = (options.position ?? 'after') === 'before' ? range.from : range.to;
  }

  const date = options.date ?? new Date().toISOString();
  const tracked = options.nodes.map((n) => trackBlock(n, state, options.author, date, allocateId));

  try {
    const tr = state.tr.insert(pos, tracked);
    return tr.docChanged ? tr : null;
  } catch {
    return null;
  }
}
