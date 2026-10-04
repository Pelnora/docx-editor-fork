/**
 * "Vanilla view" text extraction for DocxEditor.
 *
 * Skips text inside `insertion` marks (tracked-change additions that
 * aren't accepted yet) so the agent's view of the document matches what
 * `add_comment` / `suggest_change` can anchor to. Tracked deletions stay
 * included — they're still in the doc until accepted.
 */

import type { Mark, Node as PMNode } from 'prosemirror-model';

/** Text of a single PM node (typically a paragraph), vanilla view. */
export function getVanillaNodeText(node: PMNode): string {
  const parts: string[] = [];
  node.descendants((child) => {
    if (!child.isText || !child.text) return true;
    if (child.marks.some((m) => m.type.name === 'insertion')) return false;
    parts.push(child.text);
    return true;
  });
  return parts.join('');
}

/** Text between two doc positions, vanilla view. */
export function getVanillaTextBetween(doc: PMNode, from: number, to: number): string {
  if (from >= to) return '';
  const parts: string[] = [];
  doc.nodesBetween(from, to, (child, pos) => {
    if (!child.isText || !child.text) return;
    if (child.marks.some((m) => m.type.name === 'insertion')) return;
    const start = Math.max(from, pos);
    const end = Math.min(to, pos + child.text.length);
    if (start < end) parts.push(child.text.slice(start - pos, end - pos));
  });
  return parts.join('');
}

/**
 * Find `searchText` within a PM paragraph range and return its position.
 *
 * Returns null if:
 *   - searchText is empty
 *   - searchText is not found
 *   - searchText appears more than once (ambiguous; caller disambiguates)
 *
 * The fullText is built from PM text nodes only and matches the vanilla
 * view the agent reads via `read_document`: tracked insertions are
 * excluded (not in the doc yet), tracked deletions are included (still
 * in the doc until accepted), and comment markers are stripped.
 */
/** Vanilla text of a paragraph range + the PM position of each kept text run,
 *  built identically to the agent's read view (insertion-marked runs excluded,
 *  deletions kept). Shared by findTextInPmParagraph (search) and
 *  mapVanillaOffsetToPm (offset). */
function vanillaTextAndPositions(
  doc: PMNode,
  paragraphFrom: number,
  paragraphTo: number
): { fullText: string; textPositions: { pos: number; len: number }[] } {
  let fullText = '';
  const textPositions: { pos: number; len: number }[] = [];
  doc.nodesBetween(paragraphFrom, paragraphTo, (node, pos) => {
    if (!node.isText || !node.text) return;
    if (node.marks.some((m) => m.type.name === 'insertion')) return;
    textPositions.push({ pos, len: node.text.length });
    fullText += node.text;
  });
  return { fullText, textPositions };
}

/** Map a [start, start + length) span in the vanilla string back to PM
 *  positions, walking the kept runs. Returns null if the span isn't fully
 *  covered by the paragraph's vanilla text. */
function mapCharSpanToPm(
  textPositions: { pos: number; len: number }[],
  start: number,
  length: number
): { from: number; to: number } | null {
  const end = start + length;
  let charOffset = 0;
  let fromPos: number | null = null;
  let toPos: number | null = null;
  for (const tp of textPositions) {
    const segEnd = charOffset + tp.len;
    if (fromPos === null && charOffset <= start && start < segEnd) {
      fromPos = tp.pos + (start - charOffset);
    }
    if (charOffset <= end && end <= segEnd) {
      toPos = tp.pos + (end - charOffset);
      break;
    }
    charOffset = segEnd;
  }
  return fromPos !== null && toPos !== null ? { from: fromPos, to: toPos } : null;
}

/**
 * Map a POINT in a paragraph's vanilla text (0 … length) to a PM position, for
 * a pure insertion at an offset (Pelnora .9). At a boundary between two runs
 * the point belongs to the run BEFORE it, so inserted text continues that
 * run, as typing does in Word; offset 0 is the start of the first run. Null
 * when the offset is outside the vanilla text or the paragraph has none.
 */
export function mapVanillaOffsetToPmPoint(
  doc: PMNode,
  paragraphFrom: number,
  paragraphTo: number,
  offset: number
): number | null {
  const { fullText, textPositions } = vanillaTextAndPositions(doc, paragraphFrom, paragraphTo);
  if (offset < 0 || offset > fullText.length) return null;
  // An empty paragraph (a blank line of a template): its content start.
  if (textPositions.length === 0) return offset === 0 ? paragraphFrom + 1 : null;
  let charOffset = 0;
  for (const tp of textPositions) {
    const segEnd = charOffset + tp.len;
    if (offset <= segEnd) return tp.pos + (offset - charOffset);
    charOffset = segEnd;
  }
  return null;
}

export function findTextInPmParagraph(
  doc: PMNode,
  paragraphFrom: number,
  paragraphTo: number,
  searchText: string
): { from: number; to: number } | null {
  if (!searchText) return null;
  const { fullText, textPositions } = vanillaTextAndPositions(doc, paragraphFrom, paragraphTo);
  const firstMatch = fullText.indexOf(searchText);
  if (firstMatch === -1) return null;
  // Reject ambiguous searches — the LLM gets a clearer error than a silent mistarget.
  const secondMatch = fullText.indexOf(searchText, firstMatch + 1);
  if (secondMatch !== -1) return null;
  return mapCharSpanToPm(textPositions, firstMatch, searchText.length);
}

/**
 * Map a KNOWN character span [offset, offset + length) in a paragraph's vanilla
 * text to PM positions. Offset-based sibling of findTextInPmParagraph WITHOUT
 * the indexOf search and WITHOUT the ambiguity rejection — the caller already
 * knows the exact occurrence offset, so multiple identical phrases in one
 * paragraph can each be addressed individually (Pelnora c11f: "replace all
 * occurrences" inside one paragraph). Returns null if the span is empty or
 * falls outside the paragraph's vanilla text.
 */
export function mapVanillaOffsetToPm(
  doc: PMNode,
  paragraphFrom: number,
  paragraphTo: number,
  offset: number,
  length: number
): { from: number; to: number } | null {
  if (offset < 0 || length <= 0) return null;
  const { fullText, textPositions } = vanillaTextAndPositions(doc, paragraphFrom, paragraphTo);
  if (offset + length > fullText.length) return null;
  return mapCharSpanToPm(textPositions, offset, length);
}

/** Marks inserted text never inherits: the change marks themselves, comment
 *  ranges (the app repairs those on save), links and footnote references. */
const NOT_INHERITED = new Set(['insertion', 'deletion', 'comment', 'hyperlink', 'footnoteRef']);

/** A run whose look is not plain text: a link (its colour and underline come
 *  from the link style), a footnote reference (superscript) or hidden text.
 *  Inserted text never takes its formatting from such a run. */
const NOT_A_SOURCE = new Set(['hyperlink', 'footnoteRef', 'hidden']);

function formattingSource(node: PMNode | null | undefined): PMNode | null {
  if (!node || !node.isText) return null;
  return node.marks.some((m) => NOT_A_SOURCE.has(m.type.name)) ? null : node;
}

/**
 * The formatting marks (font, size, bold, colour …) inserted text takes from
 * the text it lands in (Pelnora .9): without them a tracked insertion came out
 * in the document default font, unlike the words around it. `prefer` picks the
 * run after the position (a replacement: the first struck character) or before
 * it (an insertion continues the run it follows); the other side is the
 * fallback when the preferred one is missing, not text (a tab, an image), or a
 * link, a footnote reference or hidden text. Neither: no formatting.
 */
export function inheritedFormattingMarks(
  doc: PMNode,
  pos: number,
  prefer: 'after' | 'before'
): readonly Mark[] {
  const $pos = doc.resolve(pos);
  const [first, second] =
    prefer === 'after' ? [$pos.nodeAfter, $pos.nodeBefore] : [$pos.nodeBefore, $pos.nodeAfter];
  const source = formattingSource(first) ?? formattingSource(second);
  if (!source) return [];
  return source.marks.filter((m) => !NOT_INHERITED.has(m.type.name));
}
