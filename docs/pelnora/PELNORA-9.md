# pelnora-docx-editor-react 1.0.3-pelnora.9

Core unchanged (`pelnora-docx-editor-core` 1.0.3-pelnora.1). PK decisions of
2026-10-03/04 (Pelnora app, tracked changes like Word).

- `DocxEditor` prop `showTrackedChangeCards` (default `true`). `false`: tracked
  changes get no cards in the comments sidebar, comments keep theirs, and the
  sidebar (and the page shift that comes with it) appears only when there are
  comments. A proposed or inserted change and a document loaded with changes
  no longer open the sidebar (`openSidebarForChanges`). The host provides
  accept / reject (core `acceptChange`, `rejectChange`, `acceptAllChanges`,
  `rejectAllChanges`; paragraph-mark revisions are the host's to resolve).
- `proposeChangeAt` with `length: 0` and a non-empty `replaceWith` inserts at
  the offset without striking anything (`mapVanillaOffsetToPmPoint`; at a run
  boundary the point stays in the run before it; an empty paragraph takes it
  at its start). Refuses a point strictly inside an existing tracked run.
- Inserted text (`proposeChange`, `proposeChangeAt`) keeps the formatting marks
  of the text it replaces or follows (`inheritedFormattingMarks`): font, size,
  bold, colour and the rest, never the change, comment, link or footnote
  marks, and never from a link, a footnote reference or hidden text (the other
  side then, or none). Before, it came out in the document default font.
- `DocxEditor` prop `zoomFit` (`ZoomFitOption`: `label`, `active`, `onFit`,
  `onLevel`): an entry above the zoom levels; the host computes and sets the
  fit, and hears every level picked (also one equal to the current zoom).
- `PageIndicator` carries the class `docx-page-indicator` for host styling and
  is visual only (`aria-hidden`; as a live status it was read on every scroll).

Tests: `vanillaPoint.test.ts` (point mapper, inherited marks). `bun test`
938 + 7 pass. Pack: `bun run build` in `packages/react`, then
`node scripts/pelnora/pack.mjs <outDir>`; only the react tarball changes.
