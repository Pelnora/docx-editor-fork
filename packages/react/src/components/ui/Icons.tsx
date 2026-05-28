/**
 * Inline SVG Icons - Material Symbols
 *
 * Official Material Symbols from Google Fonts, bundled as inline SVGs.
 * Source: https://fonts.google.com/icons
 */

import type { CSSProperties } from 'react';
import {
  ArrowDownload20Regular,
  ArrowRedo20Regular,
  ArrowUndo20Regular,
  ArrowUpload20Regular,
  ChevronDown20Regular,
  ChevronLeft20Regular,
  ChevronRight20Regular,
  ChevronUp20Regular,
  ColorFill20Regular,
  Comment20Regular,
  CommentAdd20Regular,
  CommentEdit20Regular,
  DocumentPageBreak20Regular,
  Edit20Regular,
  Eye20Regular,
  Highlight20Regular,
  Image20Regular,
  Link20Regular,
  MoreVertical20Regular,
  Print20Regular,
  Table20Regular,
  TextAlignCenter20Regular,
  TextAlignJustify20Regular,
  TextAlignLeft20Regular,
  TextAlignRight20Regular,
  TextBold20Regular,
  TextBulletListLtr20Regular,
  TextClearFormatting20Regular,
  TextColor20Regular,
  TextIndentDecreaseLtr20Regular,
  TextIndentIncreaseLtr20Regular,
  TextItalic20Regular,
  TextLineSpacing20Regular,
  TextNumberListLtr20Regular,
  TextStrikethrough20Regular,
  TextSubscript20Regular,
  TextSuperscript20Regular,
  TextUnderline20Regular,
} from '@fluentui/react-icons';

export interface IconProps {
  size?: number;
  className?: string;
  style?: CSSProperties;
}

const defaultSize = 20;

// SVG wrapper for Material Symbols (viewBox 0 -960 960 960)
function SvgIcon({
  size = defaultSize,
  className = '',
  style,
  children,
}: IconProps & { children: React.ReactNode }) {
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      width={size}
      height={size}
      viewBox="0 -960 960 960"
      fill="currentColor"
      className={className}
      style={{ display: 'inline-flex', flexShrink: 0, ...style }}
      aria-hidden="true"
    >
      {children}
    </svg>
  );
}

// ============================================================================
// TOOLBAR ICONS
// ============================================================================

export function IconUndo({ size = defaultSize, className, style }: IconProps) {
  return <ArrowUndo20Regular fontSize={size} className={className} style={style} />;
}

export function IconRedo({ size = defaultSize, className, style }: IconProps) {
  return <ArrowRedo20Regular fontSize={size} className={className} style={style} />;
}

export function IconPrint({ size = defaultSize, className, style }: IconProps) {
  return <Print20Regular fontSize={size} className={className} style={style} />;
}

export function IconFileDownload({ size = defaultSize, className, style }: IconProps) {
  return <ArrowDownload20Regular fontSize={size} className={className} style={style} />;
}

export function IconFileUpload({ size = defaultSize, className, style }: IconProps) {
  return <ArrowUpload20Regular fontSize={size} className={className} style={style} />;
}

export function IconBold({ size = defaultSize, className, style }: IconProps) {
  return <TextBold20Regular fontSize={size} className={className} style={style} />;
}

export function IconItalic({ size = defaultSize, className, style }: IconProps) {
  return <TextItalic20Regular fontSize={size} className={className} style={style} />;
}

export function IconUnderline({ size = defaultSize, className, style }: IconProps) {
  return <TextUnderline20Regular fontSize={size} className={className} style={style} />;
}

export function IconStrikethrough({ size = defaultSize, className, style }: IconProps) {
  return <TextStrikethrough20Regular fontSize={size} className={className} style={style} />;
}

export function IconSuperscript({ size = defaultSize, className, style }: IconProps) {
  return <TextSuperscript20Regular fontSize={size} className={className} style={style} />;
}

export function IconSubscript({ size = defaultSize, className, style }: IconProps) {
  return <TextSubscript20Regular fontSize={size} className={className} style={style} />;
}

export function IconLink({ size = defaultSize, className, style }: IconProps) {
  return <Link20Regular fontSize={size} className={className} style={style} />;
}

export function IconFormatClear({ size = defaultSize, className, style }: IconProps) {
  return <TextClearFormatting20Regular fontSize={size} className={className} style={style} />;
}

export function IconAlignLeft({ size = defaultSize, className, style }: IconProps) {
  return <TextAlignLeft20Regular fontSize={size} className={className} style={style} />;
}

export function IconAlignCenter({ size = defaultSize, className, style }: IconProps) {
  return <TextAlignCenter20Regular fontSize={size} className={className} style={style} />;
}

export function IconAlignRight({ size = defaultSize, className, style }: IconProps) {
  return <TextAlignRight20Regular fontSize={size} className={className} style={style} />;
}

export function IconAlignJustify({ size = defaultSize, className, style }: IconProps) {
  return <TextAlignJustify20Regular fontSize={size} className={className} style={style} />;
}

export function IconLineSpacing({ size = defaultSize, className, style }: IconProps) {
  return <TextLineSpacing20Regular fontSize={size} className={className} style={style} />;
}

export function IconListBulleted({ size = defaultSize, className, style }: IconProps) {
  // Pelnora D4 colour cue: subtle blue on list controls. primaryFill overrides
  // the inherited granat (currentColor), so this stays blue regardless of the
  // toolbar's granat recolor.
  return <TextBulletListLtr20Regular fontSize={size} className={className} style={style} primaryFill="#3f6ea3" />;
}

export function IconListNumbered({ size = defaultSize, className, style }: IconProps) {
  // Pelnora D4 colour cue: subtle blue on list controls (see IconListBulleted).
  return <TextNumberListLtr20Regular fontSize={size} className={className} style={style} primaryFill="#3f6ea3" />;
}

export function IconIndentIncrease({ size = defaultSize, className, style }: IconProps) {
  return <TextIndentIncreaseLtr20Regular fontSize={size} className={className} style={style} />;
}

export function IconIndentDecrease({ size = defaultSize, className, style }: IconProps) {
  return <TextIndentDecreaseLtr20Regular fontSize={size} className={className} style={style} />;
}

export function IconTextColor({ size = defaultSize, className, style }: IconProps) {
  return <TextColor20Regular fontSize={size} className={className} style={style} />;
}

export function IconHighlight({ size = defaultSize, className, style }: IconProps) {
  return <Highlight20Regular fontSize={size} className={className} style={style} />;
}

export function IconColorReset(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M800-436q0 36-8 69t-22 63l-62-60q6-17 9-34.5t3-37.5q0-47-17.5-89T650-600L480-768l-88 86-56-56 144-142 226 222q44 42 69 99.5T800-436Zm-8 380L668-180q-41 29-88 44.5T480-120q-133 0-226.5-92.5T160-436q0-51 16-98t48-90L56-792l56-56 736 736-56 56ZM480-200q36 0 68.5-10t61.5-28L280-566q-21 32-30.5 64t-9.5 66q0 98 70 167t170 69Zm-37-204Zm110-116Z" />
    </SvgIcon>
  );
}

export function IconDropdown({ size = defaultSize, className, style }: IconProps) {
  return <ChevronDown20Regular fontSize={size} className={className} style={style} />;
}

// ============================================================================
// TABLE ICONS
// ============================================================================

export function IconTable({ size = defaultSize, className, style }: IconProps) {
  return <Table20Regular fontSize={size} className={className} style={style} />;
}

export function IconTableChart(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M760-120H200q-33 0-56.5-23.5T120-200v-560q0-33 23.5-56.5T200-840h560q33 0 56.5 23.5T840-760v560q0 33-23.5 56.5T760-120ZM200-640h560v-120H200v120Zm100 80H200v360h100v-360Zm360 0v360h100v-360H660Zm-80 0H380v360h200v-360Z" />
    </SvgIcon>
  );
}

export function IconGridOn(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M200-120q-33 0-56.5-23.5T120-200v-560q0-33 23.5-56.5T200-840h560q33 0 56.5 23.5T840-760v560q0 33-23.5 56.5T760-120H200Zm0-80h133v-133H200v133Zm213 0h134v-133H413v133Zm214 0h133v-133H627v133ZM200-413h133v-134H200v134Zm213 0h134v-134H413v134Zm214 0h133v-134H627v134ZM200-627h133v-133H200v133Zm213 0h134v-133H413v133Zm214 0h133v-133H627v133Z" />
    </SvgIcon>
  );
}

export function IconTableRows(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M760-200v-120H200v120h560Zm0-200v-160H200v160h560Zm0-240v-120H200v120h560ZM200-120q-33 0-56.5-23.5T120-200v-560q0-33 23.5-56.5T200-840h560q33 0 56.5 23.5T840-760v560q0 33-23.5 56.5T760-120H200Z" />
    </SvgIcon>
  );
}

export function IconViewColumn(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M121-280v-400q0-33 23.5-56.5T201-760h559q33 0 56.5 23.5T840-680v400q0 33-23.5 56.5T760-200H201q-33 0-56.5-23.5T121-280Zm79 0h133v-400H200v400Zm213 0h133v-400H413v400Zm213 0h133v-400H626v400Z" />
    </SvgIcon>
  );
}

export function IconBorderAll(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M120-120v-720h720v720H120Zm640-80v-240H520v240h240Zm0-560H520v240h240v-240Zm-560 0v240h240v-240H200Zm0 560h240v-240H200v240Z" />
    </SvgIcon>
  );
}

export function IconBorderOuter(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M200-200h560v-560H200v560Zm-80 80v-720h720v720H120Zm160-320v-80h80v80h-80Zm160 160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm160 160v-80h80v80h-80Z" />
    </SvgIcon>
  );
}

export function IconBorderInner(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M120-120v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-320v-80h80v80h-80Zm0-160v-80h80v80h-80Zm160 640v-80h80v80h-80Zm0-640v-80h80v80h-80Zm320 640v-80h80v80h-80Zm160 0v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-320v-80h80v80h-80Zm0-160v-80h80v80h-80Zm-160 0v-80h80v80h-80ZM440-120v-320H120v-80h320v-320h80v320h320v80H520v320h-80Z" />
    </SvgIcon>
  );
}

export function IconBorderClear(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M120-120v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm160 640v-80h80v80h-80Zm0-320v-80h80v80h-80Zm0-320v-80h80v80h-80Zm160 640v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm160 640v-80h80v80h-80Zm0-320v-80h80v80h-80Zm0-320v-80h80v80h-80Zm160 640v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Z" />
    </SvgIcon>
  );
}

export function IconAdd(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M440-440H200v-80h240v-240h80v240h240v80H520v240h-80v-240Z" />
    </SvgIcon>
  );
}

export function IconRemove(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M200-440v-80h560v80H200Z" />
    </SvgIcon>
  );
}

export function IconDelete(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M280-120q-33 0-56.5-23.5T200-200v-520h-40v-80h200v-40h240v40h200v80h-40v520q0 33-23.5 56.5T680-120H280Zm400-600H280v520h400v-520ZM360-280h80v-360h-80v360Zm160 0h80v-360h-80v360ZM280-720v520-520Z" />
    </SvgIcon>
  );
}

export function IconDeleteSweep(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M600-240v-80h160v80H600Zm0-320v-80h280v80H600Zm0 160v-80h240v80H600ZM120-640H80v-80h160v-60h160v60h160v80h-40v360q0 33-23.5 56.5T440-200H200q-33 0-56.5-23.5T120-280v-360Zm80 0v360h240v-360H200Zm0 0v360-360Z" />
    </SvgIcon>
  );
}

export function IconMerge(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="m296-160-56-56 200-200v-269L337-582l-57-57 200-200 201 201-57 57-104-104v301L296-160Zm368 1L536-286l57-57 127 128-56 56Z" />
    </SvgIcon>
  );
}

export function IconSplit(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M440-160v-304L240-664v104h-80v-240h240v80H296l224 224v336h-80Zm154-376-58-58 128-126H560v-80h240v240h-80v-104L594-536Z" />
    </SvgIcon>
  );
}

export function IconDragIndicator(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M360-160q-33 0-56.5-23.5T280-240q0-33 23.5-56.5T360-320q33 0 56.5 23.5T440-240q0 33-23.5 56.5T360-160Zm240 0q-33 0-56.5-23.5T520-240q0-33 23.5-56.5T600-320q33 0 56.5 23.5T680-240q0 33-23.5 56.5T600-160ZM360-400q-33 0-56.5-23.5T280-480q0-33 23.5-56.5T360-560q33 0 56.5 23.5T440-480q0 33-23.5 56.5T360-400Zm240 0q-33 0-56.5-23.5T520-480q0-33 23.5-56.5T600-560q33 0 56.5 23.5T680-480q0 33-23.5 56.5T600-400ZM360-640q-33 0-56.5-23.5T280-720q0-33 23.5-56.5T360-800q33 0 56.5 23.5T440-720q0 33-23.5 56.5T360-640Zm240 0q-33 0-56.5-23.5T520-720q0-33 23.5-56.5T600-800q33 0 56.5 23.5T680-720q0 33-23.5 56.5T600-640Z" />
    </SvgIcon>
  );
}

// ============================================================================
// IMAGE TOOLBAR ICONS
// ============================================================================

export function IconImage({ size = defaultSize, className, style }: IconProps) {
  return <Image20Regular fontSize={size} className={className} style={style} />;
}

export function IconFormatImageLeft(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M120-280v-400h400v400H120Zm80-80h240v-240H200v240Zm-80-400v-80h720v80H120Zm480 160v-80h240v80H600Zm0 160v-80h240v80H600Zm0 160v-80h240v80H600ZM120-120v-80h720v80H120Z" />
    </SvgIcon>
  );
}

export function IconFormatImageRight(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M440-280v-400h400v400H440Zm80-80h240v-240H520v240ZM120-120v-80h720v80H120Zm0-160v-80h240v80H120Zm0-160v-80h240v80H120Zm0-160v-80h240v80H120Zm0-160v-80h720v80H120Z" />
    </SvgIcon>
  );
}

export function IconHorizontalRule(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M160-440v-80h640v80H160Z" />
    </SvgIcon>
  );
}

export function IconFlipToBack(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M200-120q-33 0-56.5-23.5T120-200v-480h80v480h480v80H200Zm160-240v80q-33 0-56.5-23.5T280-360h80Zm-80-80v-80h80v80h-80Zm0-160v-80h80v80h-80Zm80-160h-80q0-33 23.5-56.5T360-840v80Zm80 480v-80h80v80h-80Zm0-480v-80h80v80h-80Zm160 0v-80h80v80h-80Zm0 480v-80h80v80h-80Zm160-480v-80q33 0 56.5 23.5T840-760h-80Zm0 400h80q0 33-23.5 56.5T760-280v-80Zm0-80v-80h80v80h-80Zm0-160v-80h80v80h-80Z" />
    </SvgIcon>
  );
}

export function IconFlipToFront(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M360-280q-33 0-56.5-23.5T280-360v-400q0-33 23.5-56.5T360-840h400q33 0 56.5 23.5T840-760v400q0 33-23.5 56.5T760-280H360Zm0-80h400v-400H360v400ZM200-200v80q-33 0-56.5-23.5T120-200h80Zm-80-80v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm160 480v-80h80v80h-80Zm160 0v-80h80v80h-80Zm160 0v-80h80v80h-80Z" />
    </SvgIcon>
  );
}

export function IconOpenWith(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M480-80 310-250l57-57 73 73v-166h80v165l72-73 58 58L480-80ZM250-310 80-480l169-169 57 57-72 72h166v80H235l73 72-58 58Zm460 0-57-57 73-73H560v-80h165l-73-72 58-58 170 170-170 170ZM440-560v-166l-73 73-57-57 170-170 170 170-57 57-73-73v166h-80Z" />
    </SvgIcon>
  );
}

export function IconTune(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M440-120v-240h80v80h320v80H520v80h-80Zm-320-80v-80h240v80H120Zm160-160v-80H120v-80h160v-80h80v240h-80Zm160-80v-80h400v80H440Zm160-160v-240h80v80h160v80H680v80h-80Zm-480-80v-80h400v80H120Z" />
    </SvgIcon>
  );
}

export function IconRotateRight(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M522-80v-82q34-5 66.5-18t61.5-34l56 58q-42 32-88 51.5T522-80Zm-80 0Q304-98 213-199.5T122-438q0-75 28.5-140.5t77-114q48.5-48.5 114-77T482-798h6l-62-62 56-58 160 160-160 160-56-56 64-64h-8q-117 0-198.5 81.5T202-438q0 104 68 182.5T442-162v82Zm322-134-58-56q21-29 34-61.5t18-66.5h82q-5 50-24.5 96T764-214Zm76-264h-82q-5-34-18-66.5T706-606l58-56q32 39 51 86t25 98Z" />
    </SvgIcon>
  );
}

export function IconRotateLeft(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M440-80q-50-5-96-24.5T256-156l56-58q29 21 61.5 34t66.5 18v82Zm80 0v-82q104-15 172-93.5T760-438q0-117-81.5-198.5T480-718h-8l64 64-56 56-160-160 160-160 56 58-62 62h6q75 0 140.5 28.5t114 77q48.5 48.5 77 114T840-438q0 137-91 238.5T520-80ZM198-214q-32-42-51.5-88T122-398h82q5 34 18 66.5t34 61.5l-58 56Zm-76-264q6-51 25-98t51-86l58 56q-21 29-34 61.5T204-478h-82Z" />
    </SvgIcon>
  );
}

export function IconSwapHoriz(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M280-160 80-360l200-200 56 57-103 103h287v80H233l103 103-56 57Zm400-240-56-57 103-103H440v-80h287L624-743l56-57 200 200-200 200Z" />
    </SvgIcon>
  );
}

export function IconSwapVert(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M320-440v-287L217-624l-57-56 200-200 200 200-57 56-103-103v287h-80ZM600-80 400-280l57-56 103 103v-287h80v287l103-103 57 56L600-80Z" />
    </SvgIcon>
  );
}

export function IconShapes(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M600-360ZM320-242q10 1 19.5 1.5t20.5.5q11 0 20.5-.5T400-242v82h400v-400h-82q1-10 1.5-19.5t.5-20.5q0-11-.5-20.5T718-640h82q33 0 56.5 23.5T880-560v400q0 33-23.5 56.5T800-80H400q-33 0-56.5-23.5T320-160v-82Zm40-78q-117 0-198.5-81.5T80-600q0-117 81.5-198.5T360-880q117 0 198.5 81.5T640-600q0 117-81.5 198.5T360-320Zm0-80q83 0 141.5-58.5T560-600q0-83-58.5-141.5T360-800q-83 0-141.5 58.5T160-600q0 83 58.5 141.5T360-400Zm0-200Z" />
    </SvgIcon>
  );
}

// ============================================================================
// TABLE DROPDOWN ICONS
// ============================================================================

export function IconFormatPaint(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M440-80q-33 0-56.5-23.5T360-160v-160H240q-33 0-56.5-23.5T160-400v-280q0-66 47-113t113-47h480v440q0 33-23.5 56.5T720-320H600v160q0 33-23.5 56.5T520-80h-80ZM240-560h480v-200h-40v160h-80v-160h-40v80h-80v-80H320q-33 0-56.5 23.5T240-680v120Zm0 160h480v-80H240v80Zm0 0v-80 80Z" />
    </SvgIcon>
  );
}

export function IconExpandMore({ size = defaultSize, className, style }: IconProps) {
  return <ChevronDown20Regular fontSize={size} className={className} style={style} />;
}

export function IconExpandLess({ size = defaultSize, className, style }: IconProps) {
  return <ChevronUp20Regular fontSize={size} className={className} style={style} />;
}

export function IconBorderTop(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M120-120v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h720v80H120Zm160 640v-80h80v80h-80Zm0-320v-80h80v80h-80Zm160 320v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm160 480v-80h80v80h-80Zm0-320v-80h80v80h-80Zm160 320v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Z" />
    </SvgIcon>
  );
}

export function IconBorderBottom(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M120-120v-80h720v80H120Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm160 320v-80h80v80h-80Zm0-320v-80h80v80h-80Zm160 480v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm160 320v-80h80v80h-80Zm0-320v-80h80v80h-80Zm160 480v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Z" />
    </SvgIcon>
  );
}

export function IconBorderLeft(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M120-120v-720h80v720h-80Zm160 0v-80h80v80h-80Zm0-320v-80h80v80h-80Zm0-320v-80h80v80h-80Zm160 640v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm160 640v-80h80v80h-80Zm0-320v-80h80v80h-80Zm0-320v-80h80v80h-80Zm160 640v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Z" />
    </SvgIcon>
  );
}

export function IconBorderRight(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M120-120v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm160 640v-80h80v80h-80Zm0-320v-80h80v80h-80Zm0-320v-80h80v80h-80Zm160 640v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm0-160v-80h80v80h-80Zm160 640v-80h80v80h-80Zm0-320v-80h80v80h-80Zm0-320v-80h80v80h-80Zm160 640v-720h80v720h-80Z" />
    </SvgIcon>
  );
}

export function IconBorderHorizontal(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M120-120v-60h60v60h-60Zm0-165v-60h60v60h-60Zm0-165v-60h720v60H120Zm0-165v-60h60v60h-60Zm0-165v-60h60v60h-60Zm165 660v-60h60v60h-60Zm0-660v-60h60v60h-60Zm165 660v-60h60v60h-60Zm0-165v-60h60v60h-60Zm0-330v-60h60v60h-60Zm0-165v-60h60v60h-60Zm165 660v-60h60v60h-60Zm0-660v-60h60v60h-60Zm165 660v-60h60v60h-60Zm0-165v-60h60v60h-60Zm0-330v-60h60v60h-60Zm0-165v-60h60v60h-60Z" />
    </SvgIcon>
  );
}

export function IconBorderVertical(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M120-120v-60h60v60h-60Zm0-165v-60h60v60h-60Zm0-165v-60h60v60h-60Zm0-165v-60h60v60h-60Zm0-165v-60h60v60h-60Zm165 660v-60h60v60h-60Zm0-330v-60h60v60h-60Zm0-330v-60h60v60h-60Zm165 660v-720h60v720h-60Zm165 0v-60h60v60h-60Zm0-330v-60h60v60h-60Zm0-330v-60h60v60h-60Zm165 660v-60h60v60h-60Zm0-165v-60h60v60h-60Zm0-165v-60h60v60h-60Zm0-165v-60h60v60h-60Zm0-165v-60h60v60h-60Z" />
    </SvgIcon>
  );
}

export function IconPadding(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M320-600q17 0 28.5-11.5T360-640q0-17-11.5-28.5T320-680q-17 0-28.5 11.5T280-640q0 17 11.5 28.5T320-600Zm160 0q17 0 28.5-11.5T520-640q0-17-11.5-28.5T480-680q-17 0-28.5 11.5T440-640q0 17 11.5 28.5T480-600Zm160 0q17 0 28.5-11.5T680-640q0-17-11.5-28.5T640-680q-17 0-28.5 11.5T600-640q0 17 11.5 28.5T640-600ZM200-120q-33 0-56.5-23.5T120-200v-560q0-33 23.5-56.5T200-840h560q33 0 56.5 23.5T840-760v560q0 33-23.5 56.5T760-120H200Zm0-80h560v-560H200v560Zm0-560v560-560Z" />
    </SvgIcon>
  );
}

export function IconTextRotationNone(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M160-200v-80h528l-42-42 56-56 138 138-138 138-56-56 42-42H160Zm116-200 164-440h80l164 440h-76l-38-112H392l-40 112h-76Zm138-176h132l-64-182h-4l-64 182Z" />
    </SvgIcon>
  );
}

export function IconWrapText(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M588-132 440-280l148-148 56 58-50 50h96q29 0 49.5-20.5T760-390q0-29-20.5-49.5T690-460H160v-80h530q63 0 106.5 43.5T840-390q0 63-43.5 106.5T690-240h-96l50 50-56 58ZM160-240v-80h200v80H160Zm0-440v-80h640v80H160Z" />
    </SvgIcon>
  );
}

export function IconHeight(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M480-120 320-280l56-56 64 63v-414l-64 63-56-56 160-160 160 160-56 57-64-64v414l64-63 56 56-160 160Z" />
    </SvgIcon>
  );
}

export function IconFitWidth(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M120-120v-720h80v720h-80Zm640 0v-720h80v720h-80ZM280-440v-80h80v80h-80Zm160 0v-80h80v80h-80Zm160 0v-80h80v80h-80Z" />
    </SvgIcon>
  );
}

export function IconSettings(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="m370-80-16-128q-13-5-24.5-12T307-235l-119 50L78-375l103-78q-1-7-1-13.5v-27q0-6.5 1-13.5L78-585l110-190 119 50q11-8 23-15t24-12l16-128h220l16 128q13 5 24.5 12t22.5 15l119-50 110 190-103 78q1 7 1 13.5v27q0 6.5-2 13.5l103 78-110 190-118-50q-11 8-23 15t-24 12L590-80H370Zm70-80h79l14-106q31-8 57.5-23.5T639-327l99 41 39-68-86-65q5-14 7-29.5t2-31.5q0-16-2-31.5t-7-29.5l86-65-39-68-99 42q-22-23-48.5-38.5T533-694l-13-106h-79l-14 106q-31 8-57.5 23.5T321-633l-99-41-39 68 86 64q-5 15-7 30t-2 32q0 16 2 31t7 30l-86 65 39 68 99-42q22 23 48.5 38.5T427-266l13 106Zm42-180q58 0 99-41t41-99q0-58-41-99t-99-41q-59 0-99.5 41T342-480q0 58 40.5 99t99.5 41Zm-2-140Z" />
    </SvgIcon>
  );
}

export function IconBorderColor(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M80 0v-160h800V0H80Zm160-320h56l312-311-29-29-28-28-311 312v56Zm-80 80v-170l448-447q11-11 25.5-17t30.5-6q16 0 31 6t27 18l55 56q12 11 17.5 26t5.5 31q0 15-5.5 29.5T777-687L330-240H160Zm560-504-56-56 56 56ZM608-631l-29-29-28-28 57 57Z" />
    </SvgIcon>
  );
}

export function IconFormatColorFill({ size = defaultSize, className, style }: IconProps) {
  return <ColorFill20Regular fontSize={size} className={className} style={style} />;
}

export function IconVerticalAlignTop(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M160-760v-80h640v80H160Zm280 640v-408L336-424l-56-56 200-200 200 200-56 56-104-104v408h-80Z" />
    </SvgIcon>
  );
}

export function IconVerticalAlignCenter(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M440-80v-168l-64 64-56-56 160-160 160 160-56 56-64-64v168h-80ZM160-440v-80h640v80H160Zm320-120L320-720l56-56 64 64v-168h80v168l64-64 56 56-160 160Z" />
    </SvgIcon>
  );
}

export function IconVerticalAlignBottom(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M160-120v-80h640v80H160Zm320-160L280-480l56-56 104 104v-408h80v408l104-104 56 56-200 200Z" />
    </SvgIcon>
  );
}

// Table toolbar icons
export function IconLineWeight(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M120-160v-40h720v40H120Zm0-120v-80h720v80H120Zm0-160v-120h720v120H120Zm0-200v-160h720v160H120Z" />
    </SvgIcon>
  );
}

export function IconKeyboardArrowUp({ size = defaultSize, className, style }: IconProps) {
  return <ChevronUp20Regular fontSize={size} className={className} style={style} />;
}

export function IconKeyboardArrowDown({ size = defaultSize, className, style }: IconProps) {
  return <ChevronDown20Regular fontSize={size} className={className} style={style} />;
}

export function IconKeyboardArrowLeft({ size = defaultSize, className, style }: IconProps) {
  return <ChevronLeft20Regular fontSize={size} className={className} style={style} />;
}

export function IconKeyboardArrowRight({ size = defaultSize, className, style }: IconProps) {
  return <ChevronRight20Regular fontSize={size} className={className} style={style} />;
}

export function IconMoreVert({ size = defaultSize, className, style }: IconProps) {
  return <MoreVertical20Regular fontSize={size} className={className} style={style} />;
}

export function IconPageBreak({ size = defaultSize, className, style }: IconProps) {
  return <DocumentPageBreak20Regular fontSize={size} className={className} style={style} />;
}

export function IconArrowBack(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M313-440l224 224-57 56-320-320 320-320 57 56-224 224h487v80H313Z" />
    </SvgIcon>
  );
}

export function IconDoneAll(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M268-240 42-466l57-56 170 170 56 56-57 56Zm226 0L268-466l56-57 170 170 368-368 57 57-425 424Zm0-226-57-56 198-198 57 56-198 198Z" />
    </SvgIcon>
  );
}

export function IconCheckCircle(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="m424-296 282-282-56-56-226 226-114-114-56 56 170 170Zm56 216q-83 0-156-31.5T197-197q-54-54-85.5-127T80-480q0-83 31.5-156T197-763q54-54 127-85.5T480-880q83 0 156 31.5T763-763q54 54 85.5 127T880-480q0 83-31.5 156T763-197q-54 54-127 85.5T480-80Z" />
    </SvgIcon>
  );
}

/** Plain speech bubble outline (no lines inside) */
export function IconChatBubbleOutline(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M80-80v-720q0-33 23.5-56.5T160-880h640q33 0 56.5 23.5T880-800v480q0 33-23.5 56.5T800-240H240L80-80Zm126-240h594v-480H160v525l46-45Zm-46 0v-480 480Z" />
    </SvgIcon>
  );
}

/** Speech bubble with green checkmark (bubble inherits color, check is green) */
export function IconChatBubbleCheck(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M80-80v-720q0-33 23.5-56.5T160-880h640q33 0 56.5 23.5T880-800v480q0 33-23.5 56.5T800-240H240L80-80Zm126-240h594v-480H160v525l46-45Zm-46 0v-480 480Z" />
      <path fill="#188038" d="m421-380 227-227-45-45-182 182-92-91-45 45 137 136Z" />
    </SvgIcon>
  );
}

export function IconCheck(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M382-240 154-468l57-57 171 171 367-367 57 57-424 424Z" />
    </SvgIcon>
  );
}

export function IconClose(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="m256-200-56-56 224-224-224-224 56-56 224 224 224-224 56 56-224 224 224 224-56 56-224-224-224 224Z" />
    </SvgIcon>
  );
}

export function IconAddComment({ size = defaultSize, className, style }: IconProps) {
  return <CommentAdd20Regular fontSize={size} className={className} style={style} />;
}

export function IconComment({ size = defaultSize, className, style }: IconProps) {
  return <Comment20Regular fontSize={size} className={className} style={style} />;
}

export function IconEditNote({ size = defaultSize, className, style }: IconProps) {
  return <Edit20Regular fontSize={size} className={className} style={style} />;
}

export function IconRateReview({ size = defaultSize, className, style }: IconProps) {
  return <CommentEdit20Regular fontSize={size} className={className} style={style} />;
}

export function IconVisibility({ size = defaultSize, className, style }: IconProps) {
  return <Eye20Regular fontSize={size} className={className} style={style} />;
}

// Text direction icons
export function IconTextDirectionLtr(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M360-360v-200q-66 0-113-47t-47-113q0-66 47-113t113-47h320v80h-80v440h-80v-440h-80v440h-80Zm0-280v-160q-33 0-56.5 23.5T280-720q0 33 23.5 56.5T360-640Zm0-80ZM680-80l-56-56 64-64H120v-80h568l-64-64 56-56 160 160L680-80Z" />
    </SvgIcon>
  );
}

export function IconTextDirectionRtl(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="M360-360v-200q-66 0-113-47t-47-113q0-66 47-113t113-47h320v80h-80v440h-80v-440h-80v440h-80Zm-88 160 64 64-56 56-160-160 160-160 56 56-64 64h568v80H272Zm88-440v-160q-33 0-56.5 23.5T280-720q0 33 23.5 56.5T360-640Zm0-80Z" />
    </SvgIcon>
  );
}

// Material Symbol "auto_awesome" — official path from Google Fonts.
// Source: https://fonts.gstatic.com/s/i/short-term/release/materialsymbolsoutlined/auto_awesome/default/24px.svg
export function IconAgentSparkle(props: IconProps) {
  return (
    <SvgIcon {...props}>
      <path d="m760-600-50-110-110-50 110-50 50-110 50 110 110 50-110 50-50 110Zm0 560-50-110-110-50 110-50 50-110 50 110 110 50-110 50-50 110ZM360-160 260-380 40-480l220-100 100-220 100 220 220 100-220 100-100 220Zm0-194 40-86 86-40-86-40-40-86-40 86-86 40 86 40 40 86Zm0-126Z" />
    </SvgIcon>
  );
}

// ============================================================================
// ICON MAP - for MaterialSymbol compatibility
// ============================================================================

const iconMap: Record<string, React.ComponentType<IconProps>> = {
  undo: IconUndo,
  redo: IconRedo,
  print: IconPrint,
  file_download: IconFileDownload,
  file_upload: IconFileUpload,
  format_bold: IconBold,
  format_italic: IconItalic,
  format_underlined: IconUnderline,
  strikethrough_s: IconStrikethrough,
  superscript: IconSuperscript,
  subscript: IconSubscript,
  link: IconLink,
  format_clear: IconFormatClear,
  format_align_left: IconAlignLeft,
  format_align_center: IconAlignCenter,
  format_align_right: IconAlignRight,
  format_align_justify: IconAlignJustify,
  format_line_spacing: IconLineSpacing,
  format_list_bulleted: IconListBulleted,
  format_list_numbered: IconListNumbered,
  format_indent_increase: IconIndentIncrease,
  format_indent_decrease: IconIndentDecrease,
  format_color_text: IconTextColor,
  ink_highlighter: IconHighlight,
  format_color_reset: IconColorReset,
  arrow_drop_down: IconDropdown,
  table: IconTable,
  table_chart: IconTableChart,
  grid_on: IconGridOn,
  table_rows: IconTableRows,
  view_column: IconViewColumn,
  border_all: IconBorderAll,
  border_outer: IconBorderOuter,
  border_inner: IconBorderInner,
  border_clear: IconBorderClear,
  add: IconAdd,
  remove: IconRemove,
  delete: IconDelete,
  delete_sweep: IconDeleteSweep,
  call_merge: IconMerge,
  call_split: IconSplit,
  drag_indicator: IconDragIndicator,
  // Image toolbar
  image: IconImage,
  format_image_left: IconFormatImageLeft,
  format_image_right: IconFormatImageRight,
  horizontal_rule: IconHorizontalRule,
  flip_to_back: IconFlipToBack,
  flip_to_front: IconFlipToFront,
  open_with: IconOpenWith,
  tune: IconTune,
  rotate_right: IconRotateRight,
  rotate_left: IconRotateLeft,
  swap_horiz: IconSwapHoriz,
  swap_vert: IconSwapVert,
  // Shape gallery
  shapes: IconShapes,
  // Table dropdown
  format_paint: IconFormatPaint,
  expand_more: IconExpandMore,
  expand_less: IconExpandLess,
  border_top: IconBorderTop,
  border_bottom: IconBorderBottom,
  border_left: IconBorderLeft,
  border_right: IconBorderRight,
  border_horizontal: IconBorderHorizontal,
  border_vertical: IconBorderVertical,
  padding: IconPadding,
  text_rotation_none: IconTextRotationNone,
  wrap_text: IconWrapText,
  height: IconHeight,
  fit_width: IconFitWidth,
  settings: IconSettings,
  border_color: IconBorderColor,
  format_color_fill: IconFormatColorFill,
  vertical_align_top: IconVerticalAlignTop,
  vertical_align_center: IconVerticalAlignCenter,
  vertical_align_bottom: IconVerticalAlignBottom,
  // Table toolbar new icons
  line_weight: IconLineWeight,
  keyboard_arrow_up: IconKeyboardArrowUp,
  keyboard_arrow_down: IconKeyboardArrowDown,
  keyboard_arrow_left: IconKeyboardArrowLeft,
  keyboard_arrow_right: IconKeyboardArrowRight,
  more_vert: IconMoreVert,
  // Page break
  page_break: IconPageBreak,
  // Navigation
  arrow_back: IconArrowBack,
  // Comments sidebar
  done_all: IconDoneAll,
  check_circle: IconCheckCircle,
  chat_bubble_outline: IconChatBubbleOutline,
  chat_bubble_check: IconChatBubbleCheck,
  check: IconCheck,
  close: IconClose,
  add_comment: IconAddComment,
  comment: IconComment,
  edit_note: IconEditNote,
  rate_review: IconRateReview,
  visibility: IconVisibility,
  // Text direction
  format_textdirection_l_to_r: IconTextDirectionLtr,
  format_textdirection_r_to_l: IconTextDirectionRtl,
  // Agent
  'agent-sparkle': IconAgentSparkle,
};

/**
 * MaterialSymbol-compatible component using inline SVGs
 */
export function MaterialSymbol({
  name,
  size = 20,
  className = '',
  style,
}: {
  name: string;
  size?: number;
  filled?: boolean;
  weight?: number;
  className?: string;
  style?: CSSProperties;
}) {
  const IconComponent = iconMap[name];

  if (!IconComponent) {
    // Fallback: render the name as text (for debugging)
    console.warn(`Icon not found: ${name}`);
    return (
      <span className={className} style={{ fontSize: size, width: size, height: size, ...style }}>
        {name}
      </span>
    );
  }

  return <IconComponent size={size} className={className} style={style} />;
}

export default MaterialSymbol;
