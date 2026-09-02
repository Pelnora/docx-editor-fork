/**
 * Pelnora H0-3 (b)/(d): paragraph-mark revisions and rFonts fidelity.
 *
 * ECMA-376 §17.13.5.15 / CT_ParaRPr: a tracked-inserted paragraph carries the
 * revision on its paragraph mark as `w:pPr/w:rPr/w:ins` (first child of the
 * paragraph-mark rPr). The core needs a model for it so `insertTracked`
 * paragraphs come out as real Word revisions and so Word-authored ones
 * round-trip.
 *
 */

import { describe, test, expect } from 'bun:test';
import { parseXml, type XmlElement } from '../xmlParser';
import { parseParagraphProperties } from '../paragraphParser/properties';
import {
  serializeParagraph,
  serializeParagraphFormatting,
} from '../serializer/paragraphSerializer';
import type { Paragraph, ParagraphFormatting } from '../../types/document';

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const DATE = '2026-09-02T10:00:00Z';
const AUTHOR = 'Asystent Pelnory';

function parsePPr(inner: string): XmlElement {
  const doc = parseXml(`<w:pPr ${W_NS}>${inner}</w:pPr>`);
  return (doc.elements as XmlElement[])[0];
}

function pPrRoundTrip(inner: string): { formatting: ParagraphFormatting | undefined; xml: string } {
  const formatting = parseParagraphProperties(parsePPr(inner), null);
  return { formatting, xml: serializeParagraphFormatting(formatting) };
}

describe('paragraph-mark revision (w:pPr/w:rPr/w:ins)', () => {
  test('parses w:ins on the paragraph mark into formatting.paragraphMarkChange', () => {
    const { formatting } = pPrRoundTrip(
      `<w:rPr><w:ins w:id="7" w:author="${AUTHOR}" w:date="${DATE}"/><w:rFonts w:hint="eastAsia"/></w:rPr>`
    );
    expect(formatting?.paragraphMarkChange).toEqual({
      type: 'insertion',
      info: { id: 7, author: AUTHOR, date: DATE },
    });
    expect(formatting?.runProperties?.fontFamily?.hint).toBe('eastAsia');
  });

  test('serializes the paragraph-mark w:ins first inside w:rPr, before run properties', () => {
    const { xml } = pPrRoundTrip(
      `<w:rPr><w:ins w:id="7" w:author="${AUTHOR}" w:date="${DATE}"/><w:rFonts w:hint="eastAsia"/></w:rPr>`
    );
    expect(xml).toBe(
      `<w:pPr><w:rPr><w:ins w:id="7" w:author="${AUTHOR}" w:date="${DATE}"/><w:rFonts w:hint="eastAsia"/></w:rPr></w:pPr>`
    );
  });

  test('paragraph mark with only the revision (no run properties) still gets a w:rPr', () => {
    const { xml } = pPrRoundTrip(
      `<w:rPr><w:ins w:id="7" w:author="${AUTHOR}" w:date="${DATE}"/></w:rPr>`
    );
    expect(xml).toBe(
      `<w:pPr><w:rPr><w:ins w:id="7" w:author="${AUTHOR}" w:date="${DATE}"/></w:rPr></w:pPr>`
    );
  });

  test('pStyle stays first, paragraph-mark rPr stays last', () => {
    const { xml } = pPrRoundTrip(
      `<w:pStyle w:val="Heading1"/><w:rPr><w:ins w:id="7" w:author="${AUTHOR}" w:date="${DATE}"/></w:rPr>`
    );
    expect(xml).toBe(
      `<w:pPr><w:pStyle w:val="Heading1"/><w:rPr><w:ins w:id="7" w:author="${AUTHOR}" w:date="${DATE}"/></w:rPr></w:pPr>`
    );
  });

  test.each([
    ['del', 'deletion'],
    ['moveFrom', 'moveFrom'],
    ['moveTo', 'moveTo'],
  ] as const)('w:%s on the paragraph mark round-trips', (tag, type) => {
    const { formatting, xml } = pPrRoundTrip(
      `<w:rPr><w:${tag} w:id="9" w:author="Jan Testowy" w:date="${DATE}"/></w:rPr>`
    );
    expect(formatting?.paragraphMarkChange?.type).toBe(type);
    expect(xml).toContain(
      `<w:rPr><w:${tag} w:id="9" w:author="Jan Testowy" w:date="${DATE}"/></w:rPr>`
    );
  });

  test('a paragraph mark without a revision serializes no tracked-change element', () => {
    const { formatting, xml } = pPrRoundTrip('<w:rPr><w:b/></w:rPr>');
    expect(formatting?.paragraphMarkChange).toBeUndefined();
    expect(xml).toBe('<w:pPr><w:rPr><w:b/></w:rPr></w:pPr>');
  });

  test('a tracked-inserted heading serializes like Word: pPr/rPr/ins + w:ins around the runs', () => {
    const paragraph: Paragraph = {
      type: 'paragraph',
      formatting: {
        styleId: 'Heading1',
        paragraphMarkChange: { type: 'insertion', info: { id: 7, author: AUTHOR, date: DATE } },
      },
      content: [
        {
          type: 'insertion',
          info: { id: 8, author: AUTHOR, date: DATE },
          content: [{ type: 'run', content: [{ type: 'text', text: 'Nowy nagłówek' }] }],
        },
      ],
    };
    expect(serializeParagraph(paragraph)).toBe(
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/><w:rPr><w:ins w:id="7" w:author="${AUTHOR}" w:date="${DATE}"/></w:rPr></w:pPr>` +
        `<w:ins w:id="8" w:author="${AUTHOR}" w:date="${DATE}"><w:r><w:t>Nowy nagłówek</w:t></w:r></w:ins></w:p>`
    );
  });
});
