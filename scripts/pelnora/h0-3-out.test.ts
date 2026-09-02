/**
 * Pelnora H0-3 — produce the three DOCX files the orchestrator opens in Word
 * (no Word in this session), through the exact code the app uses:
 *
 *   01_po_zamianie.docx          fixture 01 + one proposeChange-style redline
 *                                (deletion + insertion, distinct ids, author)
 *   01_po_tracked_insert.docx    fixture 01 + insertTracked (heading, bold run,
 *                                two list-like paragraphs) after one paragraph
 *   05_roundtrip_bez_edycji.docx fixture 05 (Word revisions + comments), no edit
 *
 * Per file: xmllint --noout on every XML part, a structural summary on stdout
 * (w:ins / w:del / w:delText / moveFrom / w:cols / paragraph-mark w:ins /
 * duplicate revision ids) and assertions on the H0-3 contract plus
 * `w14:paraId` / `w14:textId` validity (ST_LongHexNumber, unique).
 *
 * Usage:
 *   bun test ./scripts/pelnora/h0-3-out.test.ts
 *   H0_3_OUT=<dir> bun test ./scripts/pelnora/h0-3-out.test.ts
 *                                        (default /tmp/pelnora-audit/fork-out)
 *
 * Why a test file: toProseDoc needs a DOM and happy-dom 20.x registers
 * reliably only under `bun test` (`bun run` hit the happy-dom ESM interop
 * error "Missing 'default' export" on one Bun setup and floods stderr with
 * stylesheet-loading errors on another). The file lives outside bunfig
 * `root = ./packages`, so the regular suite never picks it up — it writes to
 * /tmp and needs xmllint. The leading `./` matters: without it bun treats the
 * argument as a name filter and finds nothing.
 */

import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import * as fs from 'node:fs';
import * as path from 'node:path';
import JSZip from 'jszip';
import { EditorState } from '../../packages/core/node_modules/prosemirror-state';
import type { Node as PMNode } from '../../packages/core/node_modules/prosemirror-model';
import { parseDocx } from '../../packages/core/src/docx/parser';
import { repackDocx } from '../../packages/core/src/docx/rezip';
import { toProseDoc } from '../../packages/core/src/prosemirror/conversion/toProseDoc';
import { fromProseDoc } from '../../packages/core/src/prosemirror/conversion/fromProseDoc';
import { schema } from '../../packages/core/src/prosemirror/schema';
import type { Document } from '../../packages/core/src/types/document';
import { generateHexId, MAX_HEX_ID_EXCLUSIVE } from '../../packages/core/src/utils/hexId';
import { buildTrackedInsertTransaction } from '../../packages/react/src/components/DocxEditor/internals/trackedInsert';
import { findTextInPmParagraph } from '../../packages/react/src/components/DocxEditor/internals/vanillaText';
import { findParaIdRange } from '../../packages/react/src/components/DocxEditor/internals/pmAnchors';

const ROOT = path.resolve(import.meta.dir, '../..');
const FIXTURES = path.join(ROOT, 'e2e/fixtures/pelnora');
const OUT = path.resolve(process.env.H0_3_OUT ?? '/tmp/pelnora-audit/fork-out');
const AUTHOR = 'Asystent Pelnory';
const TIMEOUT_MS = 60_000;
/** ST_LongHexNumber as Word writes it: 8 uppercase hex digits. */
const HEX8 = /^[0-9A-F]{8}$/;

beforeAll(() => {
  fs.mkdirSync(OUT, { recursive: true });
  GlobalRegistrator.register({
    settings: {
      disableCSSFileLoading: true,
      disableJavaScriptFileLoading: true,
      // The core's font loader appends <link> tags for document fonts; with
      // loading disabled happy-dom would log one error per link otherwise.
      handleDisabledFileLoadingAsSuccess: true,
    },
  });
});
afterAll(() => GlobalRegistrator.unregister());

function loadFixture(name: string): ArrayBuffer {
  const b = fs.readFileSync(path.join(FIXTURES, name));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
}

/**
 * Synthetic fixtures carry no w14:paraId. The app allocates them through
 * ParaIdAllocatorExtension (a plugin this script bypasses by building
 * EditorState without plugins), so mirror it: generateHexId() — 8 uppercase
 * hex digits below MAX_HEX_ID_EXCLUSIVE, the strictest ST_LongHexNumber bound
 * — unique in the document. textId gets the same value (the serializer writes
 * both). An earlier revision wrote `P0001…`, which is not a valid
 * ST_LongHexNumber and could make Word report unreadable content.
 */
function withParaIds(doc: PMNode): PMNode {
  const seen = new Set<string>();
  doc.descendants((node) => {
    if (node.type.name === 'paragraph' && node.attrs.paraId) seen.add(node.attrs.paraId as string);
    return true;
  });
  const tr = EditorState.create({ doc, schema }).tr;
  doc.descendants((node, pos) => {
    if (node.type.name === 'paragraph' && !node.attrs.paraId) {
      let id = generateHexId();
      while (seen.has(id)) id = generateHexId();
      seen.add(id);
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, paraId: id, textId: id });
    }
    return true;
  });
  return tr.doc;
}

function paraIdContaining(doc: PMNode, needle: string): string {
  let found: string | null = null;
  doc.descendants((node) => {
    if (found) return false;
    if (node.type.name === 'paragraph' && node.textContent.includes(needle)) {
      found = node.attrs.paraId as string;
      return false;
    }
    return true;
  });
  if (!found) throw new Error(`paragraph containing ${JSON.stringify(needle)} not found`);
  return found;
}

function maxRevisionId(doc: PMNode): number {
  let max = 0;
  doc.descendants((node) => {
    for (const m of node.marks)
      if (m.attrs.revisionId != null) max = Math.max(max, Number(m.attrs.revisionId));
    return true;
  });
  return max;
}

async function open(name: string): Promise<{ parsed: Document; state: EditorState }> {
  const parsed = await parseDocx(loadFixture(name));
  const doc = withParaIds(toProseDoc(parsed));
  return { parsed, state: EditorState.create({ doc, schema }) };
}

async function save(state: EditorState, parsed: Document, name: string): Promise<string> {
  const back = fromProseDoc(state.doc, parsed);
  const buffer = await repackDocx(back);
  const file = path.join(OUT, name);
  fs.writeFileSync(file, Buffer.from(buffer));
  return file;
}

// --- 1. proposeChange-style replacement (mirrors useDocxEditorRefApi.proposeChange) ---
async function makeReplacement(): Promise<string> {
  const { parsed, state } = await open('01_umowa_B2B_uslugi.docx');
  let nextId = maxRevisionId(state.doc) + 1;
  const paraId = paraIdContaining(state.doc, '18 000 zł netto miesięcznie');
  const range = findParaIdRange(state.doc, paraId)!;
  const span = findTextInPmParagraph(state.doc, range.from, range.to, '18 000 zł');
  if (!span) throw new Error('search span not found');
  const date = new Date().toISOString();
  const deletionMark = schema.marks.deletion.create({ revisionId: nextId++, author: AUTHOR, date });
  const insertionMark = schema.marks.insertion.create({
    revisionId: nextId++,
    author: AUTHOR,
    date,
  });
  const tr = state.tr
    .addMark(span.from, span.to, deletionMark)
    .insert(span.to, schema.text('20 000 zł', [insertionMark]));
  return save(state.apply(tr), parsed, '01_po_zamianie.docx');
}

// --- 2. insertTracked (the real helper behind DocxEditorRef.insertTracked) ---
async function makeTrackedInsert(): Promise<string> {
  const { parsed, state } = await open('01_umowa_B2B_uslugi.docx');
  let nextId = maxRevisionId(state.doc) + 1;
  const paraId = paraIdContaining(state.doc, '3. Harmonogram płatności');
  const bold = schema.marks.bold.create();
  const nodes = [
    schema.node('paragraph', { styleId: 'Heading2' }, [schema.text('Zastrzeżenie dodatkowe')]),
    schema.node('paragraph', null, [
      schema.text('Uwaga: ', [bold]),
      schema.text(
        'wynagrodzenie za miesiąc rozpoczęty ustala się proporcjonalnie do liczby dni świadczenia usług.'
      ),
    ]),
    schema.node('paragraph', { styleId: 'ListParagraph' }, [
      schema.text('1) faktura wystawiana jest ostatniego dnia miesiąca;'),
    ]),
    schema.node('paragraph', { styleId: 'ListParagraph' }, [
      schema.text('2) termin płatności liczy się od dnia doręczenia faktury.'),
    ]),
  ];
  const tr = buildTrackedInsertTransaction(
    state,
    { paraId, position: 'after', nodes, author: AUTHOR },
    () => nextId++
  );
  if (!tr) throw new Error('insertTracked refused');
  return save(state.apply(tr), parsed, '01_po_tracked_insert.docx');
}

// --- 3. no-edit round-trip of the revisions + comments fixture ---
async function makeRoundTrip(): Promise<string> {
  const parsed = await parseDocx(loadFixture('05_rewizje_i_komentarze.docx'));
  const state = EditorState.create({ doc: toProseDoc(parsed), schema });
  return save(state, parsed, '05_roundtrip_bez_edycji.docx');
}

// --- verification without Word ---
interface Summary {
  parts: number;
  lintFailures: number;
  ins: number;
  del: number;
  delText: number;
  moves: number;
  cols: number;
  paragraphMarkIns: number;
  author: number;
  delIds: string[];
  insIds: string[];
  duplicateRevisionIds: string[];
  paraIds: string[];
  textIds: string[];
  comments: number | null;
}

async function verify(file: string): Promise<Summary> {
  const xmllint = Bun.which('xmllint');
  if (!xmllint) throw new Error('xmllint not found on PATH (libxml2) — needed for the XML check');
  const zip = await JSZip.loadAsync(fs.readFileSync(file));
  const dir = file.replace(/\.docx$/, '');
  fs.rmSync(dir, { recursive: true, force: true });
  fs.mkdirSync(dir, { recursive: true });
  const parts = Object.keys(zip.files).filter((n) => !zip.files[n]!.dir);
  let lintFailures = 0;
  for (const name of parts) {
    const target = path.join(dir, name);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, Buffer.from(await zip.file(name)!.async('arraybuffer')));
    if (/\.(xml|rels)$/.test(name)) {
      const r = Bun.spawnSync([xmllint, '--noout', target]);
      if (r.exitCode !== 0) {
        lintFailures++;
        console.log(`  xmllint FAIL ${name}: ${new TextDecoder().decode(r.stderr).trim()}`);
      }
    }
  }
  const xml = await zip.file('word/document.xml')!.async('text');
  const count = (re: RegExp) => (xml.match(re) ?? []).length;
  const attrValues = (re: RegExp) => [...xml.matchAll(re)].map((m) => m[1]!);
  const ids = attrValues(/<w:(?:ins|del|moveFrom|moveTo)\b[^>]*\bw:id="(\d+)"/g);
  const comments = zip.file('word/comments.xml');
  const s: Summary = {
    parts: parts.length,
    lintFailures,
    ins: count(/<w:ins\b/g),
    del: count(/<w:del\b/g),
    delText: count(/<w:delText\b/g),
    moves: count(/<w:move(?:From|To)\b/g),
    cols: count(/<w:cols\b/g),
    paragraphMarkIns: count(/<w:rPr><w:ins\b/g),
    author: count(new RegExp(`w:author="${AUTHOR}"`, 'g')),
    delIds: attrValues(/<w:del\b[^>]*\bw:id="(\d+)"/g),
    insIds: attrValues(/<w:ins\b[^>]*\bw:id="(\d+)"/g),
    duplicateRevisionIds: ids.filter((id, i) => ids.indexOf(id) !== i),
    paraIds: attrValues(/\bw14:paraId="([^"]*)"/g),
    textIds: attrValues(/\bw14:textId="([^"]*)"/g),
    comments: comments
      ? ((await comments.async('text')).match(/<w:comment\b/g) ?? []).length
      : null,
  };
  console.log(`${path.basename(file)}`);
  console.log(`  parts: ${s.parts}, xmllint failures: ${s.lintFailures}`);
  console.log(
    `  w:ins ${s.ins}  w:del ${s.del}  w:delText ${s.delText}  moveFrom/moveTo ${s.moves}  ` +
      `w:cols ${s.cols}  paragraph-mark w:ins ${s.paragraphMarkIns}  author "${AUTHOR}" ${s.author}  ` +
      `duplicate revision ids: ${s.duplicateRevisionIds.length ? s.duplicateRevisionIds.join(',') : 'none'}`
  );
  console.log(
    `  w14:paraId ${s.paraIds.length} (${new Set(s.paraIds).size} distinct)  ` +
      `w14:textId ${s.textIds.length} (${new Set(s.textIds).size} distinct)`
  );
  if (s.comments !== null) console.log(`  comments.xml: ${s.comments} comments`);
  return s;
}

/** Every id is ST_LongHexNumber (8 hex, < MAX_HEX_ID_EXCLUSIVE) and unique. */
function expectValidHexIds(ids: string[]): void {
  const invalid = ids.filter((id) => !HEX8.test(id) || parseInt(id, 16) >= MAX_HEX_ID_EXCLUSIVE);
  expect(invalid).toEqual([]);
  expect(ids.filter((id, i) => ids.indexOf(id) !== i)).toEqual([]);
}

function expectCommon(s: Summary): void {
  expect(s.lintFailures).toBe(0);
  expect(s.moves).toBe(0);
  expect(s.cols).toBeGreaterThan(0);
  expect(s.duplicateRevisionIds).toEqual([]);
  expectValidHexIds(s.paraIds);
  expectValidHexIds(s.textIds);
}

const written: string[] = [];

describe('H0-3 Word-check files', () => {
  test(
    '01_po_zamianie.docx — proposeChange redline is w:del + w:ins with distinct ids',
    async () => {
      const file = await makeReplacement();
      written.push(file);
      const s = await verify(file);
      expectCommon(s);
      expect(s.paraIds.length).toBeGreaterThan(0);
      expect(s.textIds).toEqual(s.paraIds);
      expect(s.delIds).toEqual(['1']);
      expect(s.insIds).toEqual(['2']);
      expect(s.delText).toBe(1);
      expect(s.author).toBe(2);
    },
    TIMEOUT_MS
  );

  test(
    '01_po_tracked_insert.docx — four paragraphs as w:ins runs + paragraph-mark w:ins',
    async () => {
      const file = await makeTrackedInsert();
      written.push(file);
      const s = await verify(file);
      expectCommon(s);
      expect(s.paraIds.length).toBeGreaterThan(0);
      expect(s.textIds).toEqual(s.paraIds);
      expect(s.ins).toBe(8);
      expect(s.del).toBe(0);
      expect(s.paragraphMarkIns).toBe(4);
      expect(s.author).toBe(8);
    },
    TIMEOUT_MS
  );

  test(
    '05_roundtrip_bez_edycji.docx — Word revisions and comments survive an untouched save',
    async () => {
      const file = await makeRoundTrip();
      written.push(file);
      const s = await verify(file);
      expectCommon(s);
      expect(s.ins).toBe(3);
      expect(s.del).toBe(2);
      expect(s.comments).toBe(2);
    },
    TIMEOUT_MS
  );

  afterAll(() => {
    console.log(`written to ${OUT}: ${written.map((f) => path.basename(f)).join(', ')}`);
    const soffice =
      Bun.which('soffice') ??
      (fs.existsSync('/Applications/LibreOffice.app/Contents/MacOS/soffice')
        ? '/Applications/LibreOffice.app/Contents/MacOS/soffice'
        : null);
    if (!soffice) {
      console.log('LibreOffice niedostępne (soffice not found) — PDF conversion skipped.');
      return;
    }
    for (const f of written) {
      const r = Bun.spawnSync([soffice, '--headless', '--convert-to', 'pdf', '--outdir', OUT, f]);
      console.log(`  soffice → pdf ${path.basename(f)}: exit ${r.exitCode}`);
    }
  });
});
