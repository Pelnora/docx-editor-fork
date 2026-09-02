/**
 * Pelnora H0-3 (c): `<w:cols>` in sectPr must survive parse → serialize.
 *
 * Word writes the single-column default as `<w:cols w:space="708"/>` (no
 * `w:num`, no `w:col` children). The serializer used to emit nothing unless a
 * column count or explicit columns were present, so every save dropped the
 * element (audit C3-1 / C-03).
 */

import { describe, test, expect } from 'bun:test';
import { parseXml, type XmlElement } from '../xmlParser';
import { parseSectionProperties } from '../sectionParser';
import { serializeSectionProperties } from '../serializer/documentSerializer';

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

function parseSectPr(inner: string): XmlElement {
  const doc = parseXml(`<w:sectPr ${W_NS}>${inner}</w:sectPr>`);
  return (doc.elements as XmlElement[])[0];
}

function roundTrip(inner: string): string {
  const props = parseSectionProperties(parseSectPr(inner));
  return serializeSectionProperties(props);
}

describe('sectPr w:cols round-trip (Pelnora H0-3 c)', () => {
  test('Word default single column `<w:cols w:space="708"/>` is preserved', () => {
    const xml = roundTrip('<w:pgSz w:w="11906" w:h="16838"/><w:cols w:space="708"/>');
    expect(xml).toContain('<w:cols w:space="708"/>');
  });

  test('space-only cols keeps columnSpace on the model', () => {
    const props = parseSectionProperties(parseSectPr('<w:cols w:space="708"/>'));
    expect(props.columnSpace).toBe(708);
    expect(props.columnCount).toBeUndefined();
  });

  test('multi-column definition round-trips num/space/equalWidth/col children', () => {
    const xml = roundTrip(
      '<w:cols w:num="2" w:space="708" w:equalWidth="0"><w:col w:w="4000" w:space="500"/><w:col w:w="5000"/></w:cols>'
    );
    expect(xml).toContain('<w:cols w:num="2" w:space="708" w:equalWidth="0">');
    expect(xml).toContain('<w:col w:w="4000" w:space="500"/>');
    expect(xml).toContain('<w:col w:w="5000"/>');
    expect(xml).toContain('</w:cols>');
  });

  test('separator-only cols is preserved', () => {
    const xml = roundTrip('<w:cols w:sep="1"/>');
    expect(xml).toContain('<w:cols w:sep="1"/>');
  });

  test('no cols in input → no cols in output', () => {
    const xml = roundTrip('<w:pgSz w:w="11906" w:h="16838"/>');
    expect(xml).not.toContain('<w:cols');
  });
});
