# H0-3 — fork `pelnora-docx-editor-react` → 1.0.3-pelnora.8

Recon + plan for the Pelnora H0-3 mission (decision 4A: fix now, release `.8`).
Branch `pelnora-8` (from `pelnora-fluent-icons`). Nothing is published or pushed
from this branch by the implementing agent; publication is the orchestrator's step.

## 1. How the pieces work today (recon, with file:line)

### 1.1 `proposeChange` / `proposeChangeAt` (react)

`packages/react/src/components/DocxEditor/hooks/useDocxEditorRefApi.ts:330-403` and
`:411-463`. Both allocate **one** `revisionId = getNextCommentId()` and create a
`deletion` mark and an `insertion` mark with that same id (`:376-390`, `:446-460`),
then `tr.addMark(textFrom, textTo, deletionMark)` + `tr.insert(textTo, schema.text(replaceWith, [insertionMark]))`.
`getNextCommentId` lives in `commentFactories.ts` (shared comment/revision id
counter). The counter is bumped above ids already in the document only when the
document has comments (`useDocumentLoader.ts:118-141`: the bump sits inside
`if (bodyComments && bodyComments.length > 0)`), so a document with Word
revisions but no comments can get colliding `w:id`s.

### 1.2 Where the serializer decides move vs del/ins (core)

`packages/core/src/prosemirror/conversion/fromProseDoc/paragraph.ts:285-335`:
for every inline node with an `insertion`/`deletion` mark it looks the mark's
`revisionId` up in the document-wide counters built by
`buildDocumentTrackedChangeCounts` (`fromProseDoc/marks.ts:25-48`) and sets
`isMovePair = hasInsertionForId && hasDeletionForId`. A pair sharing one id is
emitted as `moveTo` + `moveFrom` (`:322-333`), which
`docx/serializer/paragraphSerializer/content.ts:290-320` writes as
`<w:moveTo>` / `<w:moveFrom>`. **This is exactly the fork's proposeChange pair**
→ root cause of C5-1 (`moveFrom/moveTo w:id="107"` for "b"→"f").

Real Word moves: the parser keeps `moveFrom`/`moveTo` **and** the
`moveFromRangeStart/End` / `moveToRangeStart/End` markers
(`docx/paragraphParser/content.ts:745-800`), and the Document-level serializer
writes them back verbatim (verified with a synthetic Word-style move: distinct
ids 11/13, range name `move1` — byte-identical body). `toProseDoc/paragraph.ts:107-128`
maps `moveFrom` → `deletion` mark and `moveTo` → `insertion` mark with their
own (distinct) ids and drops the range markers. Because Word never gives the two
halves the same id, the same-id heuristic **never** fires for Word moves; after a
PM round-trip they already come out as `w:del` + `w:ins` (verified). The
heuristic therefore only ever affects the fork's own pairs.

### 1.3 Where `<w:cols>` disappears

`docx/sectionParser.ts:296-345` stores `columnCount` only when `w:num` is
present and `columns` only when `w:col` children exist. Word's default
`<w:cols w:space="708"/>` yields `{ columnSpace: 708 }` only, and
`docx/serializer/documentSerializer.ts:303-304`
`if (!props.columnCount && !props.columns?.length) return '';` drops it.
Reproduced on all three fixtures and on the audit's real document.

### 1.4 `rFonts` on the paragraph mark

Paragraph-mark `w:pPr/w:rPr` is parsed into `ParagraphFormatting.runProperties`
(`paragraphParser/properties.ts:458-460`), carried through PM in
`attrs._originalFormatting` (`toProseDoc/paragraph.ts:227`,
`fromProseDoc/paragraph.ts:169-190`) and serialized by
`serializeParagraphFormatting` (`paragraphSerializer.ts:147-152`).
Re-measured on the audit's real "Wezwanie" before/after copies
(`/tmp/pelnora-audit/c3/wezwanie-noedit`): paragraph-mark `rFonts` 32 → 32
(no loss); the total 56 → 54 comes from three runs split by `<w:proofErr>`
being coalesced into one run (paragraph 23). So "3 z 20 rFonts" was run
coalescing, not a paragraph-mark loss. Two cheap fidelity gaps remain and are
fixed here: `w:hint` is neither parsed nor written (an `rFonts` carrying only
`w:hint` is dropped entirely), and the serializer writes `w:csTheme` although
the schema/Word attribute (and the parser) use `w:cstheme`.

### 1.5 Tracked insertion of blocks

Runs: the `insertion` mark on text nodes already serializes to `<w:ins>` with
`w:author`/`w:date`. Paragraph mark: ECMA-376 `CT_ParaRPr` allows
`w:pPr/w:rPr/w:ins` (schema `EG_ParaRPrTrackChanges`, first in the sequence) —
the core has no model for it (parser reads only run properties from `pPr/rPr`).
Also `fromProseDoc` emits one `<w:ins>` per text node even when consecutive
nodes share one revision (Word wraps all runs of one insertion in a single
`<w:ins>`, so a Word doc with a bold word inside an insertion currently round-trips
into two `<w:ins>` with the same id).

### 1.6 Tests / build / pack

- Tests: `bun test` (bunfig `root = ./packages`), baseline **898 pass / 0 fail, 98 files**.
  DOCX fixtures live in `e2e/fixtures/`; PM round-trip tests register happy-dom
  (`toProseDoc.test.ts`). Pre-commit: `bun run typecheck` (all workspaces),
  `bun run check:parity` (passes on HEAD), lint-staged (prettier + eslint).
- Build: `bun run build` in `packages/core` (tsup + copy-assets + inject-package-doc)
  and `packages/react` (tsup + tailwind → `dist/styles.css`). Bare `tsup` in react
  drops `dist/styles.css` (project memory).
- App save path: `editorRef.getDocument()` (= `fromProseDoc(pm, original)`) →
  `repackDocx` from `@eigenpal/docx-editor-core` (App.tsx:173-186). The fork's
  own `save()` (selective) is not used by the app, so tests target
  `parseDocx → toProseDoc → fromProseDoc → repackDocx`.

## 2. Plan of changes

### Core (`@eigenpal/docx-editor-core` fork, published as `pelnora-docx-editor-core@1.0.3-pelnora.1`)

(a) `fromProseDoc/paragraph.ts`: drop the same-id move heuristic — an `insertion`
mark always serializes as `insertion`, a `deletion` mark as `deletion`.
Keep `buildDocumentTrackedChangeCounts` / the `documentCounts` parameter for API
compatibility (no longer consulted). Real Word moves keep today's behaviour:
faithful at Document level, degraded to del/ins through PM (unchanged).
Coalesce consecutive nodes with the same tracked-change type + id + author into
one wrapper (one `<w:ins>` with several runs, like Word).
(b) Paragraph-mark revision: `ParagraphFormatting.paragraphMarkChange?: { type: 'insertion'|'deletion'|'moveFrom'|'moveTo'; info }`
parsed from `pPr/rPr/{ins,del,moveFrom,moveTo}`, serialized first inside
`<w:rPr>` (schema order), carried through PM as a dedicated paragraph attr
`paragraphMarkChange` (ParagraphExtension + ParagraphAttrs + toProseDoc/fromProseDoc).
(c) `serializeColumns`: emit `<w:cols>` whenever any column property is set.
(d) `rFonts`: parse + serialize `w:hint`; write `w:cstheme`.

### React (`pelnora-docx-editor-react@1.0.3-pelnora.8`)

- `proposeChange` / `proposeChangeAt`: two ids (deletion, insertion) — Word gives
  each annotation its own id; the sidebar already pairs adjacent del+ins by
  author/date (`extractTrackedChanges.ts:96-121`).
- `insertTracked({ paraId, position, nodes, author, date? })` on `DocxEditorRef`:
  pure helper `internals/trackedInsert.ts` builds one transaction that inserts the
  given block nodes after/before the paragraph (`paraId: null` → document end),
  adds an `insertion` mark (one id per paragraph) to every text node and sets
  `paragraphMarkChange` (own id) on each paragraph. `markdown` input is not
  implemented (the app already builds PM nodes from its own markdown blocks).
- `useDocumentLoader`: bump the id counter above existing revision ids even when
  the document has no comments.

### Publication variant for the core

Keep the source package name `@eigenpal/docx-editor-core` (workspace links,
imports in react/vue/examples, api-extractor/inject-package-doc scripts all key
on it) and rename **at pack time**: `scripts/pelnora/pack.mjs` stages
`packages/core` as `pelnora-docx-editor-core` and `packages/react` unchanged,
runs `npm pack` in the staging dir, writes tarballs to `/tmp/pelnora-audit/fork-pack/`.
The app maps both its direct dependency and the react package's dependency with
one pnpm override: `"@eigenpal/docx-editor-core": "npm:pelnora-docx-editor-core@1.0.3-pelnora.1"`.
Why not an `npm:` alias inside react's package.json: the app also depends on
`@eigenpal/docx-editor-core` directly (App.tsx imports `repackDocx`), so without the
app-side override it would get two copies of core (two PM schemas) — the override
is required either way, and the alias would break the monorepo workspace link on
the next `bun install`. Less risk for the round-trip: zero import churn.

## 3. Tests (test-first)

- `packages/core/src/docx/__tests__/pelnora-cols-roundtrip.test.ts`
- `packages/core/src/docx/__tests__/pelnora-paragraph-mark-revision.test.ts`
- `packages/core/src/prosemirror/conversion/__tests__/pelnora-tracked-changes-roundtrip.test.ts`
- `packages/core/src/docx/__tests__/pelnora-fixtures-roundtrip.test.ts` (3 fixtures in
  `e2e/fixtures/pelnora/`; normalized XML diff; counts of `w:ins`/`w:del`/`w:comment`/`w:cols`/`w:rFonts`)
- `packages/react/src/components/DocxEditor/internals/trackedInsert.test.ts`

## 4. Risks

- Word rendering of `w:pPr/w:rPr/w:ins` from our serializer is unverified until the
  orchestrator opens `/tmp/pelnora-audit/fork-out/*.docx` in Word.
- Upstream normalizations remain in the no-edit round-trip: `xml:space="preserve"`
  only where needed, `<w:bCs/>` added next to `<w:b/>`, header/footer reference
  order. They predate this work and are outside 4A; the fixture test lists them as
  the accepted diff classes.
- API extractor snapshots (`docs/api/*`) are not regenerated (same as `.6`/`.7`).

## 5. Word-check files (`scripts/pelnora/h0-3-out.test.ts`)

The three DOCX files the orchestrator opens in Word are produced by the app's
own path (`parseDocx → toProseDoc → [edit] → fromProseDoc → repackDocx`, the
`proposeChange` marks and `buildTrackedInsertTransaction`):

```
bun test ./scripts/pelnora/h0-3-out.test.ts            # → /tmp/pelnora-audit/fork-out/
H0_3_OUT=<dir> bun test ./scripts/pelnora/h0-3-out.test.ts
```

- `01_po_zamianie.docx` — fixture 01 + one redline (`w:del w:id="1"` + `w:ins w:id="2"`).
- `01_po_tracked_insert.docx` — fixture 01 + `insertTracked` of four paragraphs
  (8× `w:ins`, 4× paragraph-mark `w:pPr/w:rPr/w:ins`).
- `05_roundtrip_bez_edycji.docx` — fixture 05 saved untouched (revisions + 2 comments).

Each file is unzipped next to itself, every XML part goes through
`xmllint --noout`, and the test asserts the contract above plus `w:cols`
present, no `moveFrom`/`moveTo`, no duplicate revision ids.

Why `bun test` and not `bun run`: `toProseDoc` needs a DOM and happy-dom 20.x
registers reliably only under `bun test` (`bun run` failed with the happy-dom
ESM interop error `Missing 'default' export … Element.js` on one Bun setup and
floods stderr with stylesheet-loading errors on another). The file lives
outside bunfig `root = ./packages`, so `bun test` (whole suite) never runs it —
it writes to `/tmp` and needs `xmllint`. The leading `./` is required; without
it bun treats the argument as a name filter and finds nothing.

`w14:paraId` / `w14:textId`: the synthetic fixtures carry no paragraph ids and
the harness builds `EditorState` without plugins, so it allocates them the way
`ParaIdAllocatorExtension` does — `generateHexId()` from
`packages/core/src/utils/hexId.ts` (8 uppercase hex digits below
`0x7FFFFFFF`, the strictest `ST_LongHexNumber` bound), unique per document,
`textId` = `paraId`. An earlier revision wrote `P0001…`, which is not a valid
`ST_LongHexNumber`; Word could have reported unreadable content for a reason
unrelated to the H0-3 serializer work. The test asserts the format, the bound
and uniqueness on every generated file.
