#!/usr/bin/env bun
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
 * Then: xmllint --noout on every XML part and a structural summary
 * (w:ins / w:del / w:delText / moveFrom / w:cols / paragraph-mark w:ins /
 * duplicate revision ids).
 *
 * Usage: bun run scripts/pelnora/h0-3-out.ts [outDir]   (default /tmp/pelnora-audit/fork-out)
 */

import { GlobalRegistrator } from '@happy-dom/global-registrator';
GlobalRegistrator.register({
  settings: { disableCSSFileLoading: true, disableJavaScriptFileLoading: true },
} as never);

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
import { buildTrackedInsertTransaction } from '../../packages/react/src/components/DocxEditor/internals/trackedInsert';
import { findTextInPmParagraph } from '../../packages/react/src/components/DocxEditor/internals/vanillaText';
import { findParaIdRange } from '../../packages/react/src/components/DocxEditor/internals/pmAnchors';

const ROOT = path.resolve(import.meta.dir, '../..');
const FIXTURES = path.join(ROOT, 'e2e/fixtures/pelnora');
const OUT = path.resolve(process.argv[2] ?? '/tmp/pelnora-audit/fork-out');
const AUTHOR = 'Asystent Pelnory';

fs.mkdirSync(OUT, { recursive: true });

function loadFixture(name: string): ArrayBuffer {
  const b = fs.readFileSync(path.join(FIXTURES, name));
  return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
}

/** Synthetic fixtures carry no w14:paraId; give every paragraph one (P0001…). */
function withParaIds(doc: PMNode): PMNode {
  let n = 0;
  const tr = EditorState.create({ doc, schema }).tr;
  doc.descendants((node, pos) => {
    if (node.type.name === 'paragraph' && !node.attrs.paraId) {
      n++;
      tr.setNodeMarkup(pos, undefined, { ...node.attrs, paraId: `P${String(n).padStart(4, '0')}` });
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
async function verify(file: string): Promise<void> {
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
      const r = Bun.spawnSync(['xmllint', '--noout', target]);
      if (r.exitCode !== 0) {
        lintFailures++;
        console.log(`  xmllint FAIL ${name}: ${new TextDecoder().decode(r.stderr).trim()}`);
      }
    }
  }
  const xml = await zip.file('word/document.xml')!.async('text');
  const count = (re: RegExp) => (xml.match(re) ?? []).length;
  const ids = [...xml.matchAll(/<w:(?:ins|del|moveFrom|moveTo)\b[^>]*\bw:id="(\d+)"/g)].map(
    (m) => m[1]!
  );
  const dupes = ids.filter((id, i) => ids.indexOf(id) !== i);
  console.log(`${path.basename(file)}`);
  console.log(`  parts: ${parts.length}, xmllint failures: ${lintFailures}`);
  console.log(
    `  w:ins ${count(/<w:ins\b/g)}  w:del ${count(/<w:del\b/g)}  w:delText ${count(/<w:delText\b/g)}  ` +
      `moveFrom/moveTo ${count(/<w:move(?:From|To)\b/g)}  w:cols ${count(/<w:cols\b/g)}  ` +
      `paragraph-mark w:ins ${count(/<w:rPr><w:ins\b/g)}  author "${AUTHOR}" ${count(new RegExp(`w:author="${AUTHOR}"`, 'g'))}  ` +
      `duplicate revision ids: ${dupes.length ? dupes.join(',') : 'none'}`
  );
  const comments = zip.file('word/comments.xml');
  if (comments)
    console.log(
      `  comments.xml: ${((await comments.async('text')).match(/<w:comment\b/g) ?? []).length} comments`
    );
}

const files = [await makeReplacement(), await makeTrackedInsert(), await makeRoundTrip()];
console.log(`written to ${OUT}:`);
for (const f of files) await verify(f);
const soffice =
  Bun.which('soffice') ??
  (fs.existsSync('/Applications/LibreOffice.app/Contents/MacOS/soffice')
    ? '/Applications/LibreOffice.app/Contents/MacOS/soffice'
    : null);
if (soffice) {
  for (const f of files) {
    const r = Bun.spawnSync([soffice, '--headless', '--convert-to', 'pdf', '--outdir', OUT, f]);
    console.log(`  soffice → pdf ${path.basename(f)}: exit ${r.exitCode}`);
  }
} else {
  console.log('LibreOffice niedostępne (soffice not found) — PDF conversion skipped.');
}
