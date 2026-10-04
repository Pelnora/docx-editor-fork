# pelnora-docx-editor-react 1.0.3-pelnora.9

Core unchanged (`pelnora-docx-editor-core` 1.0.3-pelnora.1). PK decisions of
2026-10-03/04 (Pelnora app, tracked changes like Word).

- `DocxEditor` prop `showTrackedChangeCards` (default `true`). `false`: tracked
  changes get no cards in the comments sidebar, comments keep theirs, and the
  sidebar (and the page shift that comes with it) appears only when there are
  comments. The host provides accept / reject (core `acceptChange`,
  `rejectChange`, `acceptAllChanges`, `rejectAllChanges`).
- `proposeChangeAt` with `length: 0` and a non-empty `replaceWith` inserts at
  the offset without striking anything (`mapVanillaOffsetToPmPoint`; at a run
  boundary the point stays in the run before it). Refuses a point inside an
  existing tracked change.
- Inserted text (`proposeChange`, `proposeChangeAt`) keeps the formatting marks
  of the text it replaces or follows (`inheritedFormattingMarks`): font, size,
  bold, colour and the rest, never the change, comment, link or footnote
  marks. Before, it came out in the document default font.
- `DocxEditor` prop `zoomFit` (`ZoomFitOption`: `label`, `active`, `onFit`): an
  entry above the zoom levels; the host computes and sets the fit.
- `PageIndicator` carries the class `docx-page-indicator` for host styling.

Tests: `vanillaPoint.test.ts` (point mapper, inherited marks). `bun test`
938 + 6 pass. Pack: `bun run build` in `packages/react`, then
`node scripts/pelnora/pack.mjs <outDir>`; only the react tarball changes.
