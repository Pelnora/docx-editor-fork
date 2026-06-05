import { useImperativeHandle, useRef } from 'react';
import { TextSelection } from 'prosemirror-state';
import type { Mark } from 'prosemirror-model';
import type { Document } from '@eigenpal/docx-editor-core/types/document';
import type { Comment } from '@eigenpal/docx-editor-core/types/content';
import {
  DocumentAgent,
  createAgentFromDocument,
  // type imports re-added per feature as Phase 2.7 brings them in (B3 image,
  // B4 table, ...). Unused imports trip the fork's `noUnusedParameters` /
  // `noUnusedLocals` so we only keep what each shipped function references.
} from '@eigenpal/docx-editor-core/agent';
import {
  applyStyle,
  // Pelnora 6.2w Phase 2.7 — PM commands for the Wstawiaj cluster.
  // Imported from the subpath (not the package root) because the subpath
  // ships pre-built `.d.ts`, so the tsup dts pass never recompiles core
  // sources. Each command dispatches its own transaction through the
  // active editor view, so changes flow through the standard PM history
  // pipeline (handleDocumentChange fires automatically).
  insertPageBreak as pmInsertPageBreak,
  insertTable as pmInsertTable,
  generateTOC as pmGenerateTOC,
} from '@eigenpal/docx-editor-core/prosemirror/commands';
import { createStyleResolver, type SelectionState } from '@eigenpal/docx-editor-core/prosemirror';
import type { DocxInput } from '@eigenpal/docx-editor-core/utils';
import type { DocxEditorRef } from '../../DocxEditor';
import type { PagedEditorRef } from '../PagedEditor';
import { findParaIdRange } from '../internals/pmAnchors';
import {
  getVanillaNodeText,
  getVanillaTextBetween,
  findTextInPmParagraph,
  mapVanillaOffsetToPm,
} from '../internals/vanillaText';
import { mapHexToHighlightName } from '../../toolbarUtils';
import { pointsToHalfPoints } from '../../ui/FontSizePicker';
import { getNextCommentId, createComment } from '../commentFactories';

/**
 * Owns the `useImperativeHandle` that exposes the public `DocxEditorRef`
 * surface to consumers. Hand-rolled to preserve the exact dep array the
 * editor-contract gate enforces.
 *
 * The shape MUST match `DocxEditorRef` byte-for-byte —
 * `scripts/check-editor-contract.mjs` will fail otherwise.
 */
export function useDocxEditorRefApi({
  ref,
  agentRef,
  document,
  historyStateRef,
  pagedEditorRef,
  handleSave,
  handleDirectPrint,
  zoom,
  setZoom,
  scrollPageInfo,
  loadParsedDocument,
  loadBuffer,
  comments,
  setComments,
  setShowCommentsSidebar,
  contentChangeSubscribersRef,
  selectionChangeSubscribersRef,
  getCachedStyleResolver,
  handleDocumentChange,
}: {
  ref: React.ForwardedRef<DocxEditorRef>;
  agentRef: React.RefObject<DocumentAgent | null>;
  document: Document | null;
  historyStateRef: React.RefObject<Document | null>;
  pagedEditorRef: React.RefObject<PagedEditorRef | null>;
  handleSave: (options?: { selective?: boolean }) => Promise<ArrayBuffer | null>;
  handleDirectPrint: () => void;
  zoom: number;
  setZoom: (zoom: number) => void;
  scrollPageInfo: { currentPage: number; totalPages: number; visible: boolean };
  loadParsedDocument: (doc: Document) => void;
  loadBuffer: (buffer: DocxInput) => Promise<void>;
  comments: Comment[];
  setComments: React.Dispatch<React.SetStateAction<Comment[]>>;
  setShowCommentsSidebar: React.Dispatch<React.SetStateAction<boolean>>;
  contentChangeSubscribersRef: React.RefObject<Set<(doc: Document) => void>>;
  selectionChangeSubscribersRef: React.RefObject<Set<(state: SelectionState | null) => void>>;
  getCachedStyleResolver: (
    styles: Parameters<typeof createStyleResolver>[0]
  ) => ReturnType<typeof createStyleResolver>;
  /** Pelnora 6.2w Phase 2.7 — entry point into the internal document-change
   *  pipeline (history, layout, contentChange subscribers). Used by ref
   *  methods that mutate the Document model via the agent (insertImage,
   *  insertTable, page setup, TOC). PM commands like insertPageBreak
   *  dispatch through the view and don't need this. */
  handleDocumentChange: (doc: Document) => void;
}) {
  // Pelnora 6.2w Phase 2.7 B2 — Malarz formatów (format painter) store.
  // Holds raw PM Marks captured from the source selection; pasteFormat
  // dispatches a single transaction that strips matching mark types from
  // the target range and re-adds the stored marks. Storing PM-native marks
  // (rather than a translated snapshot shape) avoids any round-trip loss
  // for attrs (underline style, color rgb/themeColor, fontSize halfPoints,
  // fontFamily ascii/hAnsi, highlight colour name).
  const formatPainterStoreRef = useRef<readonly Mark[] | null>(null);

  useImperativeHandle(
    ref,
    () => ({
      getAgent: () => agentRef.current,
      getDocument: () => document,
      getEditorRef: () => pagedEditorRef.current,
      save: handleSave,
      setZoom,
      getZoom: () => zoom,
      focus: () => {
        pagedEditorRef.current?.focus();
      },
      getCurrentPage: () => scrollPageInfo.currentPage,
      getTotalPages: () => scrollPageInfo.totalPages,
      scrollToPage: (pageNumber: number) => {
        pagedEditorRef.current?.scrollToPage(pageNumber);
      },
      scrollToPosition: (pmPos: number) => {
        pagedEditorRef.current?.scrollToPosition(pmPos);
      },
      openPrintPreview: handleDirectPrint,
      print: handleDirectPrint,
      loadDocument: loadParsedDocument,
      loadDocumentBuffer: loadBuffer,

      addComment: (options) => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return null;
        const { schema } = view.state;
        if (!schema.marks.comment) return null;

        const range = findParaIdRange(view.state.doc, options.paraId);
        if (!range) return null;

        let from = range.from;
        let to = range.to;

        if (options.search) {
          const textRange = findTextInPmParagraph(
            view.state.doc,
            range.from,
            range.to,
            options.search
          );
          if (!textRange) return null;
          from = textRange.from;
          to = textRange.to;
        }

        const comment = createComment(options.text, options.author);
        const commentMark = schema.marks.comment.create({ commentId: comment.id });
        view.dispatch(view.state.tr.addMark(from, to, commentMark));
        setComments((prev) => [...prev, comment]);
        setShowCommentsSidebar(true);
        return comment.id;
      },

      replyToComment: (commentId, text, authorName) => {
        if (!comments.some((c) => c.id === commentId)) return null;
        const reply = createComment(text, authorName, commentId);
        setComments((prev) => [...prev, reply]);
        return reply.id;
      },

      resolveComment: (commentId) => {
        setComments((prev) => prev.map((c) => (c.id === commentId ? { ...c, done: true } : c)));
      },

      // Pelnora 6.2w Phase 2.4 — drive the same internal showCommentsSidebar
      // state the built-in CommentsSidebarToggle button drives. Consumers
      // can now mount a custom toggle button anywhere (e.g. inside their own
      // toolbar extras cluster) and have it stay in sync with the built-in.
      toggleCommentsSidebar: () => {
        setShowCommentsSidebar((v) => !v);
      },

      // Pelnora 6.2w Phase 2.7 B1 — page break at cursor. Pure PM command
      // dispatch; the core's `insertPageBreak` ensures a paragraph follows
      // the break and places the cursor in it.
      insertPageBreak: () => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return false;
        return pmInsertPageBreak(view.state, view.dispatch);
      },

      // Pelnora 6.2w Phase 2.7 B2 — Malarz formatów (format painter): copy.
      // Captures the marks active at the selection's start position into
      // an internal ref. Works for both collapsed cursor (stored marks +
      // marks at that position) and a non-empty range (marks at $from).
      // Returns true when at least one mark was captured, so the caller
      // can switch the trigger into its armed UX state.
      copyFormat: () => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return false;
        const marks = view.state.selection.$from.marks();
        formatPainterStoreRef.current = marks;
        return marks.length > 0;
      },

      // Pelnora 6.2w Phase 2.7 B2 — Malarz formatów: paste. Applies the
      // stored marks to the current selection range. Requires a non-empty
      // range (a bare cursor has nothing to paint onto). Strips existing
      // marks of the same TYPE first so colour/size/family values override
      // cleanly rather than layering. Clears the store afterwards so the
      // single-click flow returns to idle; double-click multi-paste is a
      // possible later enhancement.
      pasteFormat: () => {
        const view = pagedEditorRef.current?.getView();
        const stored = formatPainterStoreRef.current;
        if (!view || !stored || stored.length === 0) return false;
        const { selection } = view.state;
        const { from, to } = selection;
        if (from === to) return false;

        let tr = view.state.tr;
        for (const mark of stored) {
          tr = tr.removeMark(from, to, mark.type);
        }
        for (const mark of stored) {
          tr = tr.addMark(from, to, mark);
        }
        view.dispatch(tr);
        formatPainterStoreRef.current = null;
        return true;
      },

      // Pelnora 6.2w Phase 2.7 B2 — Malarz formatów: clear armed state
      // without applying (Esc / off-click cancel from consumer side).
      clearFormatPainter: () => {
        formatPainterStoreRef.current = null;
      },

      // Pelnora 6.2w Phase 2.7 B3 — insert image at the current cursor.
      // Delegates to DocumentAgent.insertImage at the Position derived from
      // the PM cursor (paraId → paragraphIndex via body walk; parentOffset
      // for character offset). Creates a fresh agent from the current
      // document so the mutation lands cleanly through handleDocumentChange.
      insertImage: (src, options) => {
        const view = pagedEditorRef.current?.getView();
        const doc = historyStateRef.current;
        if (!view || !doc) return false;

        const $from = view.state.selection.$from;
        let depth = $from.depth;
        while (depth > 0 && !$from.node(depth).isTextblock) depth--;
        const para = depth > 0 ? $from.node(depth) : null;
        const paraId = (para?.attrs?.paraId as string | undefined) ?? null;
        if (!paraId) return false;

        // `DocumentBody.content` holds the BlockContent array (paragraphs +
        // tables). Walk to map paraId → paragraphIndex per the Position
        // contract; only paragraphs carry a paraId so tables interleaved
        // in the body don't confuse the search.
        const content = doc.package.document?.content;
        if (!content) return false;
        let paragraphIndex = -1;
        for (let i = 0; i < content.length; i++) {
          const block = content[i] as { type?: string; paraId?: string };
          if (block.type === 'paragraph' && block.paraId === paraId) {
            paragraphIndex = i;
            break;
          }
        }
        if (paragraphIndex === -1) return false;

        const offset = $from.parentOffset;
        const agent = createAgentFromDocument(doc);
        agent.insertImage({ paragraphIndex, offset }, src, options);
        handleDocumentChange(agent.getDocument());
        return true;
      },

      // Pelnora 6.2w Phase 2.7 B4 — insert rows × cols table at the cursor.
      // Uses the PM command from the core's prosemirror/commands subpath
      // (same path the menubar's Table > Insert binds to), so the table
      // lands as a real PM tableNode with the editor's existing tableSchema
      // and the active selection inside the first cell.
      insertTable: (rows, cols) => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return false;
        if (!Number.isInteger(rows) || !Number.isInteger(cols) || rows < 1 || cols < 1) {
          return false;
        }
        return pmInsertTable(rows, cols)(view.state, view.dispatch);
      },

      // Pelnora 6.2w Phase 2.7 B6 — insert generated Spis treści at the
      // cursor. Dispatches the core's PM `generateTOC` command, which
      // walks the document's headings and inserts a Word-style TOC. The
      // shape of the inserted content (page numbers, hierarchy, hyperlink
      // anchors) is whatever the core ships today — we treat this as the
      // MVP. A custom enhancement that adds page numbers / dotted leaders
      // would land here if empirical review shows gaps.
      generateTOC: () => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return false;
        return pmGenerateTOC(view.state, view.dispatch);
      },

      // Pelnora 6.2w Phase 2.7 B5 — set document `finalSectionProperties`.
      // Mirrors the menubar's Page Setup > Apply flow (which uses the
      // identically-shaped `handlePageSetupApply` in usePageSetupControls)
      // so size / orientation / margin changes land in undo/redo through
      // the same handleDocumentChange pipeline. Partial — only the keys
      // present on `props` override; other keys keep their previous value.
      setPageSetup: (props) => {
        const doc = historyStateRef.current;
        if (!doc) return false;
        const newDoc = {
          ...doc,
          package: {
            ...doc.package,
            document: {
              ...doc.package.document,
              finalSectionProperties: {
                ...doc.package.document?.finalSectionProperties,
                ...props,
              },
            },
          },
        };
        handleDocumentChange(newDoc);
        return true;
      },

      proposeChange: (options) => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return false;
        const { schema } = view.state;
        if (!schema.marks.deletion || !schema.marks.insertion) return false;

        const range = findParaIdRange(view.state.doc, options.paraId);
        if (!range) return false;

        const isInsertion = options.search === '';
        const isDeletion = options.replaceWith === '';

        let textFrom: number;
        let textTo: number;

        if (isInsertion) {
          // Insert at end of paragraph (just before closing token).
          textFrom = range.to - 1;
          textTo = range.to - 1;
        } else {
          const textRange = findTextInPmParagraph(
            view.state.doc,
            range.from,
            range.to,
            options.search
          );
          if (!textRange) return false;
          textFrom = textRange.from;
          textTo = textRange.to;
        }

        // Refuse to layer onto an existing tracked change.
        let overlapsTrackedChange = false;
        if (textFrom < textTo) {
          view.state.doc.nodesBetween(textFrom, textTo, (node) => {
            for (const m of node.marks) {
              if (m.type === schema.marks.insertion || m.type === schema.marks.deletion) {
                overlapsTrackedChange = true;
                return false;
              }
            }
            return true;
          });
          if (overlapsTrackedChange) return false;
        }

        const revisionId = getNextCommentId();
        const date = new Date().toISOString();

        const deletionMark = schema.marks.deletion.create({
          revisionId,
          author: options.author,
          date,
        });
        const insertionMark = schema.marks.insertion.create({
          revisionId,
          author: options.author,
          date,
        });

        let tr = view.state.tr;
        if (!isInsertion) {
          tr = tr.addMark(textFrom, textTo, deletionMark);
        }
        if (!isDeletion) {
          const insertedNode = schema.text(options.replaceWith, [insertionMark]);
          tr = tr.insert(textTo, insertedNode);
        }

        if (isInsertion && isDeletion) return false; // nothing to do
        view.dispatch(tr);

        setShowCommentsSidebar(true);
        return true;
      },

      // Offset-addressed sibling of proposeChange. The caller passes the exact
      // [offset, offset + length) span in the paragraph's vanilla text, so a
      // specific occurrence can be targeted even when the same phrase repeats in
      // one paragraph (search would be ambiguous → proposeChange refuses).
      // Pelnora c11f. `replaceWith: ''` deletes the span.
      proposeChangeAt: (options) => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return false;
        const { schema } = view.state;
        if (!schema.marks.deletion || !schema.marks.insertion) return false;

        const range = findParaIdRange(view.state.doc, options.paraId);
        if (!range) return false;

        const span = mapVanillaOffsetToPm(
          view.state.doc,
          range.from,
          range.to,
          options.offset,
          options.length
        );
        if (!span) return false;
        const { from: textFrom, to: textTo } = span;
        if (textFrom >= textTo) return false; // offset addressing always strikes a span

        // Refuse to layer onto an existing tracked change (same guard as
        // proposeChange). A redline from the SAME multi-occurrence apply lives at
        // a different span, so it never trips this — only a PRE-existing one does.
        let overlapsTrackedChange = false;
        view.state.doc.nodesBetween(textFrom, textTo, (node) => {
          for (const m of node.marks) {
            if (m.type === schema.marks.insertion || m.type === schema.marks.deletion) {
              overlapsTrackedChange = true;
              return false;
            }
          }
          return true;
        });
        if (overlapsTrackedChange) return false;

        const revisionId = getNextCommentId();
        const date = new Date().toISOString();
        const deletionMark = schema.marks.deletion.create({
          revisionId,
          author: options.author,
          date,
        });
        let tr = view.state.tr.addMark(textFrom, textTo, deletionMark);
        if (options.replaceWith !== '') {
          const insertionMark = schema.marks.insertion.create({
            revisionId,
            author: options.author,
            date,
          });
          tr = tr.insert(textTo, schema.text(options.replaceWith, [insertionMark]));
        }
        view.dispatch(tr);

        setShowCommentsSidebar(true);
        return true;
      },

      applyFormatting: (options) => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return false;
        const { schema } = view.state;

        const range = findParaIdRange(view.state.doc, options.paraId);
        if (!range) return false;

        // Default range: the paragraph's text content (skip open/close tokens).
        let from = range.from + 1;
        let to = range.to - 1;

        if (options.search) {
          const textRange = findTextInPmParagraph(
            view.state.doc,
            range.from,
            range.to,
            options.search
          );
          if (!textRange) return false;
          from = textRange.from;
          to = textRange.to;
        }

        if (from >= to) return true;

        let tr = view.state.tr;
        const m = options.marks;

        if (m.bold !== undefined && schema.marks.bold) {
          tr = m.bold
            ? tr.addMark(from, to, schema.marks.bold.create())
            : tr.removeMark(from, to, schema.marks.bold);
        }
        if (m.italic !== undefined && schema.marks.italic) {
          tr = m.italic
            ? tr.addMark(from, to, schema.marks.italic.create())
            : tr.removeMark(from, to, schema.marks.italic);
        }
        if (m.underline !== undefined && schema.marks.underline) {
          if (m.underline) {
            const style = typeof m.underline === 'object' ? m.underline.style : undefined;
            tr = tr.addMark(from, to, schema.marks.underline.create({ style: style ?? 'single' }));
          } else {
            tr = tr.removeMark(from, to, schema.marks.underline);
          }
        }
        if (m.strike !== undefined && schema.marks.strike) {
          tr = m.strike
            ? tr.addMark(from, to, schema.marks.strike.create())
            : tr.removeMark(from, to, schema.marks.strike);
        }
        if (m.color !== undefined && schema.marks.textColor) {
          if (m.color && (m.color.rgb || m.color.themeColor)) {
            tr = tr.addMark(
              from,
              to,
              schema.marks.textColor.create({
                rgb: m.color.rgb ?? null,
                themeColor: m.color.themeColor ?? null,
              })
            );
          } else {
            tr = tr.removeMark(from, to, schema.marks.textColor);
          }
        }
        if (m.highlight !== undefined && schema.marks.highlight) {
          if (m.highlight) {
            const name = mapHexToHighlightName(m.highlight);
            tr = tr.addMark(
              from,
              to,
              schema.marks.highlight.create({ color: name || m.highlight })
            );
          } else {
            tr = tr.removeMark(from, to, schema.marks.highlight);
          }
        }
        if (m.fontSize !== undefined && schema.marks.fontSize) {
          if (m.fontSize > 0) {
            tr = tr.addMark(
              from,
              to,
              schema.marks.fontSize.create({ size: pointsToHalfPoints(m.fontSize) })
            );
          } else {
            tr = tr.removeMark(from, to, schema.marks.fontSize);
          }
        }
        if (m.fontFamily !== undefined && schema.marks.fontFamily) {
          if (m.fontFamily && (m.fontFamily.ascii || m.fontFamily.hAnsi)) {
            tr = tr.addMark(
              from,
              to,
              schema.marks.fontFamily.create({
                ascii: m.fontFamily.ascii ?? null,
                hAnsi: m.fontFamily.hAnsi ?? m.fontFamily.ascii ?? null,
              })
            );
          } else {
            tr = tr.removeMark(from, to, schema.marks.fontFamily);
          }
        }

        view.dispatch(tr);
        return true;
      },

      setParagraphStyle: (options) => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return false;

        const range = findParaIdRange(view.state.doc, options.paraId);
        if (!range) return false;

        const currentDoc = historyStateRef.current;
        const styleResolver = currentDoc?.package?.styles
          ? getCachedStyleResolver(currentDoc.package.styles)
          : null;

        // Refuse unknown styleIds so the agent gets a clear error instead of
        // silently writing `<w:pStyle w:val="NoSuchStyle"/>`. Without a
        // resolver we can't know which styles are defined, so fall through.
        if (styleResolver && !styleResolver.hasParagraphStyle(options.styleId)) {
          return false;
        }

        // Build a synthetic state with selection inside the target paragraph
        // so applyStyle's cursor-driven walk lands on it. Restore the original
        // selection on the dispatched transaction.
        const $from = view.state.doc.resolve(range.from + 1);
        const $to = view.state.doc.resolve(range.to - 1);
        const paraSelection = TextSelection.between($from, $to);
        const stateWithSel = view.state.apply(view.state.tr.setSelection(paraSelection));

        const cmd = styleResolver
          ? (() => {
              const r = styleResolver.resolveParagraphStyle(options.styleId);
              return applyStyle(options.styleId, {
                paragraphFormatting: r.paragraphFormatting,
                runFormatting: r.runFormatting,
              });
            })()
          : applyStyle(options.styleId);

        let didApply = false;
        cmd(stateWithSel, (newTr) => {
          didApply = true;
          newTr.setSelection(view.state.selection.map(newTr.doc, newTr.mapping));
          view.dispatch(newTr);
        });

        return didApply;
      },

      getPageContent: (pageNumber) => {
        const layout = pagedEditorRef.current?.getLayout();
        if (!layout) return null;
        const page = layout.pages[pageNumber - 1];
        if (!page) return null;
        const view = pagedEditorRef.current?.getView();
        if (!view) return null;
        const doc = view.state.doc;

        const seen = new Set<string>();
        const paragraphs: Array<{ paraId: string; text: string; styleId?: string }> = [];

        for (const frag of page.fragments) {
          if (frag.kind !== 'paragraph') continue;
          // `pmStart` is the position immediately before the paragraph node;
          // `doc.nodeAt(pmStart)` resolves to the paragraph itself.
          const pmStart = frag.pmStart;
          if (pmStart == null) continue;
          const node = doc.nodeAt(pmStart);
          if (!node || !node.isTextblock) continue;

          const paraId = node.attrs?.paraId as string | undefined;
          if (!paraId || seen.has(paraId)) continue;
          seen.add(paraId);
          paragraphs.push({
            paraId,
            text: getVanillaNodeText(node),
            styleId: (node.attrs?.styleId as string | undefined) ?? undefined,
          });
        }

        const text = paragraphs.map((p) => `[${p.paraId}] ${p.text}`).join('\n');
        return { pageNumber, text, paragraphs };
      },

      scrollToParaId: (paraId) => pagedEditorRef.current?.scrollToParaId(paraId) ?? false,

      findInDocument: (query, opts) => {
        const view = pagedEditorRef.current?.getView();
        if (!view || !query) return [];
        const caseSensitive = opts?.caseSensitive ?? false;
        const limit = opts?.limit ?? 20;
        const needle = caseSensitive ? query : query.toLowerCase();
        const results: Array<{
          paraId: string;
          match: string;
          before: string;
          after: string;
        }> = [];

        view.state.doc.descendants((node) => {
          if (results.length >= limit) return false;
          if (!node.isTextblock) return true;
          const paraId = node.attrs?.paraId as string | undefined;
          if (!paraId) return false;
          const text = getVanillaNodeText(node);
          const haystack = caseSensitive ? text : text.toLowerCase();
          const at = haystack.indexOf(needle);
          if (at === -1) return false;

          // Reject ambiguous matches in the same paragraph — agent should narrow query.
          if (haystack.indexOf(needle, at + 1) !== -1) return false;

          const match = text.slice(at, at + query.length);
          const CONTEXT = 40;
          results.push({
            paraId,
            match,
            before: text.slice(Math.max(0, at - CONTEXT), at),
            after: text.slice(at + query.length, at + query.length + CONTEXT),
          });
          return false;
        });

        return results;
      },

      getSelectionInfo: () => {
        const view = pagedEditorRef.current?.getView();
        if (!view) return null;
        const { selection, doc } = view.state;
        const $from = selection.$from;
        let depth = $from.depth;
        while (depth > 0 && !$from.node(depth).isTextblock) depth--;
        const para = depth > 0 ? $from.node(depth) : null;
        if (!para) return null;
        const paraId = (para.attrs?.paraId as string | undefined) ?? null;
        const paraStart = $from.start(depth);
        const paraEnd = paraStart + para.content.size;
        // Vanilla view: build before/selectedText/after from the doc so the
        // result matches what the agent reads via read_document and can anchor
        // via add_comment. Insertion-marked text never appears.
        const before = getVanillaTextBetween(doc, paraStart, selection.from);
        const selectedText = getVanillaTextBetween(doc, selection.from, selection.to);
        const after = getVanillaTextBetween(doc, selection.to, paraEnd);
        return {
          paraId,
          selectedText,
          paragraphText: before + selectedText + after,
          before,
          after,
        };
      },

      getComments: () => comments,

      onContentChange: (listener) => {
        contentChangeSubscribersRef.current.add(listener);
        return () => {
          contentChangeSubscribersRef.current.delete(listener);
        };
      },

      onSelectionChange: (listener) => {
        selectionChangeSubscribersRef.current.add(listener);
        return () => {
          selectionChangeSubscribersRef.current.delete(listener);
        };
      },
    }),
    // Dep array preserved byte-for-byte from the original site so the editor-
    // contract parity gate stays green and consumers see the same ref-identity
    // semantics they had pre-extraction.
    [
      document,
      zoom,
      scrollPageInfo,
      handleSave,
      handleDirectPrint,
      loadParsedDocument,
      loadBuffer,
      comments,
      handleDocumentChange,
    ]
  );
}
