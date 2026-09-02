/**
 * Pelnora H0-3 (d): rFonts fidelity.
 *
 * `w:hint` used to be neither parsed nor written (an rFonts carrying only
 * w:hint — common on Word paragraph marks — vanished on save) and the
 * serializer wrote `w:csTheme` although the schema attribute the parser reads
 * is `w:cstheme`.
 */

import { describe, test, expect } from 'bun:test';
import { parseXml, type XmlElement } from '../xmlParser';
import { parseRunProperties } from '../runParser';
import { serializeTextFormatting } from '../serializer/runSerializer';

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function parseRPr(inner: string): XmlElement {
  const doc = parseXml(`<w:rPr ${W_NS}>${inner}</w:rPr>`);
  return (doc.elements as XmlElement[])[0];
}

describe('rFonts fidelity (Pelnora H0-3 d)', () => {
  test('hint-only rFonts is not dropped', () => {
    const formatting = parseRunProperties(parseRPr('<w:rFonts w:hint="eastAsia"/>'), null);
    expect(formatting?.fontFamily?.hint).toBe('eastAsia');
    expect(serializeTextFormatting(formatting)).toBe(
      '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr>'
    );
  });

  test('ascii/hAnsi/hint round-trip together', () => {
    const formatting = parseRunProperties(
      parseRPr('<w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:hint="eastAsia"/>'),
      null
    );
    expect(serializeTextFormatting(formatting)).toBe(
      '<w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:hint="eastAsia"/></w:rPr>'
    );
  });

  test('complex-script theme font is written as w:cstheme (schema spelling), not w:csTheme', () => {
    const formatting = parseRunProperties(
      parseRPr('<w:rFonts w:ascii="Aptos" w:hAnsi="Aptos" w:cstheme="minorBidi"/>'),
      null
    );
    const xml = serializeTextFormatting(formatting);
    expect(xml).toContain('w:cstheme="minorBidi"');
    expect(xml).not.toContain('csTheme');
  });
});
