/**
 * Pelnora H0-3 (a)/(b): PM tracked-change marks → OOXML.
 *
 * - A deletion + insertion pair (the fork's proposeChange redline) must
 *   serialize as `w:del` + `w:ins`, never as `w:moveFrom`/`w:moveTo`
 *   (audit C5-1: the old same-revisionId heuristic turned "b"→"f" into a move).
 * - Real Word moves (distinct ids + range markers) stay faithful at Document
 *   level; through PM they degrade to del/ins exactly as before this change.
 * - Consecutive nodes sharing one revision coalesce into a single wrapper
 *   (Word wraps all runs of one insertion in one `<w:ins>`).
 * - The paragraph-mark revision travels through PM as a paragraph attr.
 */

import { afterAll, beforeAll, describe, test, expect } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import JSZip from 'jszip';
import type { Node as PMNode, Mark } from 'prosemirror-model';
import { schema } from '../../schema';
import { toProseDoc } from '../toProseDoc';
import { fromProseDoc } from '../fromProseDoc';
import { serializeParagraph } from '../../../docx/serializer/paragraphSerializer';
import { serializeDocument } from '../../../docx/serializer/documentSerializer';
import { parseDocx } from '../../../docx/parser';
import type { Document, Paragraph } from '../../../types/document';

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const AUTHOR = 'Asystent Pelnory';
const DATE = '2026-09-02T10:00:00Z';

const ins = (revisionId: number, author = AUTHOR, date = DATE): Mark =>
  schema.marks.insertion.create({ revisionId, author, date });
const del = (revisionId: number, author = AUTHOR, date = DATE): Mark =>
  schema.marks.deletion.create({ revisionId, author, date });
const text = (t: string, marks: Mark[] = []): PMNode => schema.text(t, marks);
const para = (children: PMNode[], attrs: Record<string, unknown> | null = null): PMNode =>
  schema.node('paragraph', attrs, children);
const doc = (...paras: PMNode[]): PMNode => schema.node('doc', null, paras);

function firstParagraph(pmDoc: PMNode): Paragraph {
  const back = fromProseDoc(pmDoc);
  const block = back.package.document.content[0];
  if (!block || block.type !== 'paragraph') throw new Error('expected paragraph');
  return block;
}

describe('deletion + insertion pair (proposeChange redline)', () => {
  test('distinct revision ids → w:del + w:ins, never a move', () => {
    const pm = doc(para([text('a'), text('b', [del(1)]), text('f', [ins(2)]), text('c')]));
    const paragraph = firstParagraph(pm);
    expect(paragraph.content.map((c) => c.type)).toEqual(['run', 'deletion', 'insertion', 'run']);

    const xml = serializeParagraph(paragraph);
    expect(xml).toContain(
      `<w:del w:id="1" w:author="${AUTHOR}" w:date="${DATE}"><w:r><w:delText>b</w:delText></w:r></w:del>`
    );
    expect(xml).toContain(
      `<w:ins w:id="2" w:author="${AUTHOR}" w:date="${DATE}"><w:r><w:t>f</w:t></w:r></w:ins>`
    );
    expect(xml).not.toContain('moveFrom');
    expect(xml).not.toContain('moveTo');
  });

  test('a pair that still shares one revision id (legacy .7 redline) is also w:del + w:ins', () => {
    const pm = doc(para([text('a'), text('b', [del(107)]), text('f', [ins(107)]), text('c')]));
    const xml = serializeParagraph(firstParagraph(pm));
    expect(xml).toContain('<w:del w:id="107"');
    expect(xml).toContain('<w:ins w:id="107"');
    expect(xml).not.toContain('moveFrom');
    expect(xml).not.toContain('moveTo');
  });

  test('a deletion in one paragraph and an insertion with the same id in another are not paired into a move', () => {
    const pm = doc(para([text('x', [del(5)])]), para([text('y', [ins(5)])]));
    const back = fromProseDoc(pm);
    const xml = serializeDocument(back);
    expect(xml).toContain('<w:del w:id="5"');
    expect(xml).toContain('<w:ins w:id="5"');
    expect(xml).not.toContain('moveFrom');
    expect(xml).not.toContain('moveTo');
  });

  test('author and date land on the wrapper as w:author / w:date', () => {
    const pm = doc(para([text('b', [del(1, 'Jan Testowy', '2026-08-30T10:15:00Z')])]));
    expect(serializeParagraph(firstParagraph(pm))).toContain(
      '<w:del w:id="1" w:author="Jan Testowy" w:date="2026-08-30T10:15:00Z">'
    );
  });
});

describe('coalescing consecutive nodes of one revision', () => {
  test('three text nodes (one bold) with the same insertion → one w:ins with three runs', () => {
    const bold = schema.marks.bold.create();
    const pm = doc(
      para([text('Nowy ', [ins(5)]), text('tytuł', [bold, ins(5)]), text(' koniec', [ins(5)])])
    );
    const paragraph = firstParagraph(pm);
    expect(paragraph.content).toHaveLength(1);
    expect(paragraph.content[0]!.type).toBe('insertion');
    const xml = serializeParagraph(paragraph);
    expect(xml.match(/<w:ins\b/g)).toHaveLength(1);
    expect(xml).toContain('<w:r><w:t xml:space="preserve">Nowy </w:t></w:r>');
    expect(xml).toContain('<w:b/>');
    expect(xml).toContain('<w:t>tytuł</w:t>');
    expect(xml).toContain('<w:t xml:space="preserve"> koniec</w:t>');
  });

  test('adjacent insertions with different ids stay separate wrappers', () => {
    const pm = doc(para([text('a', [ins(1)]), text('b', [ins(2)])]));
    const xml = serializeParagraph(firstParagraph(pm));
    expect(xml.match(/<w:ins\b/g)).toHaveLength(2);
  });

  test('deletion followed by insertion never coalesces across types', () => {
    const pm = doc(para([text('a', [del(1)]), text('b', [ins(1)])]));
    const paragraph = firstParagraph(pm);
    expect(paragraph.content.map((c) => c.type)).toEqual(['deletion', 'insertion']);
  });
});

describe('paragraph-mark revision through PM', () => {
  const change = { type: 'insertion' as const, info: { id: 7, author: AUTHOR, date: DATE } };

  test('Document → PM → Document keeps formatting.paragraphMarkChange and serializes pPr/rPr/ins', () => {
    const source: Document = {
      package: {
        document: {
          content: [
            {
              type: 'paragraph',
              formatting: { styleId: 'Heading1', paragraphMarkChange: change },
              content: [
                {
                  type: 'insertion',
                  info: { id: 8, author: AUTHOR, date: DATE },
                  content: [{ type: 'run', content: [{ type: 'text', text: 'Nagłówek' }] }],
                },
              ],
            },
          ],
        },
      },
    };
    const pm = toProseDoc(source);
    expect(pm.firstChild?.attrs.paragraphMarkChange).toEqual(change);

    const back = fromProseDoc(pm, source);
    const paragraph = back.package.document.content[0] as Paragraph;
    expect(paragraph.formatting?.paragraphMarkChange).toEqual(change);
    expect(paragraph.formatting?.styleId).toBe('Heading1');

    const xml = serializeParagraph(paragraph);
    expect(xml).toContain(
      `<w:pPr><w:pStyle w:val="Heading1"/><w:rPr><w:ins w:id="7" w:author="${AUTHOR}" w:date="${DATE}"/></w:rPr></w:pPr>`
    );
    expect(xml).toContain(
      `<w:ins w:id="8" w:author="${AUTHOR}" w:date="${DATE}"><w:r><w:t>Nagłówek</w:t></w:r></w:ins>`
    );
  });

  test('a fresh PM paragraph (no _originalFormatting) with the attr serializes the revision', () => {
    const pm = doc(
      para([text('Nowy', [ins(8)])], { styleId: 'Heading2', paragraphMarkChange: change })
    );
    const paragraph = firstParagraph(pm);
    expect(paragraph.formatting?.paragraphMarkChange).toEqual(change);
    expect(paragraph.formatting?.styleId).toBe('Heading2');
    expect(serializeParagraph(paragraph)).toContain(
      `<w:pPr><w:pStyle w:val="Heading2"/><w:rPr><w:ins w:id="7" w:author="${AUTHOR}" w:date="${DATE}"/></w:rPr></w:pPr>`
    );
  });

  test('a paragraph without the attr serializes no paragraph-mark revision', () => {
    const pm = doc(para([text('plain')], { styleId: 'Heading2' }));
    expect(serializeParagraph(firstParagraph(pm))).not.toContain('<w:rPr><w:ins');
  });
});

describe('real Word move (distinct ids + range markers)', () => {
  const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const MOVE_DOC =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${W}><w:body>` +
    '<w:p><w:r><w:t xml:space="preserve">Alpha </w:t></w:r>' +
    '<w:moveFromRangeStart w:id="10" w:name="move1"/>' +
    '<w:moveFrom w:id="11" w:author="Jan Testowy" w:date="2026-08-30T10:15:00Z"><w:r><w:delText>beta</w:delText></w:r></w:moveFrom>' +
    '<w:moveFromRangeEnd w:id="10"/><w:r><w:t xml:space="preserve"> gamma</w:t></w:r></w:p>' +
    '<w:p><w:moveToRangeStart w:id="12" w:name="move1"/>' +
    '<w:moveTo w:id="13" w:author="Jan Testowy" w:date="2026-08-30T10:15:00Z"><w:r><w:t>beta</w:t></w:r></w:moveTo>' +
    '<w:moveToRangeEnd w:id="12"/><w:r><w:t xml:space="preserve"> delta</w:t></w:r></w:p>' +
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:cols w:space="708"/></w:sectPr></w:body></w:document>';

  async function buildMoveDocx(): Promise<ArrayBuffer> {
    const zip = new JSZip();
    zip.file(
      '[Content_Types].xml',
      '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
        '<Default Extension="xml" ContentType="application/xml"/>' +
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    );
    zip.file(
      '_rels/.rels',
      '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    );
    zip.file('word/document.xml', MOVE_DOC);
    return zip.generateAsync({ type: 'arraybuffer' });
  }

  test('Document-level parse → serialize keeps moveFrom/moveTo and the range markers', async () => {
    const parsed = await parseDocx(await buildMoveDocx());
    const xml = serializeDocument(parsed);
    expect(xml).toContain('<w:moveFromRangeStart w:id="10" w:name="move1"/>');
    expect(xml).toContain(
      '<w:moveFrom w:id="11" w:author="Jan Testowy" w:date="2026-08-30T10:15:00Z">'
    );
    expect(xml).toContain('<w:delText>beta</w:delText>');
    expect(xml).toContain('<w:moveFromRangeEnd w:id="10"/>');
    expect(xml).toContain('<w:moveToRangeStart w:id="12" w:name="move1"/>');
    expect(xml).toContain(
      '<w:moveTo w:id="13" w:author="Jan Testowy" w:date="2026-08-30T10:15:00Z">'
    );
    expect(xml).toContain('<w:moveToRangeEnd w:id="12"/>');
    expect(xml).toContain('<w:cols w:space="708"/>');
  });

  test('through PM the move degrades to w:del + w:ins with the original ids (unchanged behaviour), never a bare move', async () => {
    const parsed = await parseDocx(await buildMoveDocx());
    const back = fromProseDoc(toProseDoc(parsed), parsed);
    const xml = serializeDocument(back);
    expect(xml).toContain('<w:del w:id="11" w:author="Jan Testowy" w:date="2026-08-30T10:15:00Z">');
    expect(xml).toContain('<w:ins w:id="13" w:author="Jan Testowy" w:date="2026-08-30T10:15:00Z">');
    expect(xml).not.toContain('moveFrom');
    expect(xml).not.toContain('moveTo');
    expect(xml).toContain('<w:cols w:space="708"/>');
  });
});
