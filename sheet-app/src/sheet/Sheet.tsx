import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { CSSProperties } from 'react';
import { Tooltip } from 'react-tooltip';
import 'react-tooltip/dist/react-tooltip.css';
import './fonts.css';
import styles from './Sheet.module.css';
import { MAX_COLS, MAX_ROWS } from './engine/SpreadsheetEngine';
import { MAX_SHEETS, Workbook } from './engine/Workbook';
import { colName } from './engine/cellref';
import { format } from './engine/format';
import { selectionToRange } from './engine/range';
import type { BorderMode, MergeRegion, NumberFormat, Range, Selection } from './engine/types';

const TOOLTIP_ID = 'excel-tip';
/** Spread onto any element to give it a react-tooltip. */
const tip = (content: string) => ({ 'data-tooltip-id': TOOLTIP_ID, 'data-tooltip-content': content });

/** Default custom border color (theme-adaptive via a CSS variable). */
const BORDER_COLOR = 'var(--ex-cell-border)';
import type { StorageAdapter } from './storage/StorageAdapter';
import { useWorkbook } from './hooks/useWorkbook';
import { useTheme } from './hooks/useTheme';
import { dangerButton, dangerGhost, mix } from './ui/styles';

export interface SheetProps {
  /**
   * Persistence backend (injected by the consuming app). The package depends
   * only on this port — see the StorageAdapter interface.
   */
  storage: StorageAdapter;
  rows?: number;
  cols?: number;
  accent?: string;
  defaultTitle?: string;
}

/** Grid geometry (column widths are per-column via the engine; see cw/colX). */
const ROWH = 30;
const HDRH = 34;
const GUTW = 52;
const LINE = 'var(--ex-gridline)';

type EditSource = 'grid' | 'bar' | null;

interface ViewState {
  sel: Selection;
  editing: boolean;
  editText: string;
  editSource: EditSource;
}

/** Where a pointed-in reference will be spliced into the formula being edited. */
interface PointAnchor {
  prefix: string;
  suffix: string;
}

const TRAILING_REF = /[A-Za-z]{1,2}\d{1,5}(?::[A-Za-z]{1,2}\d{1,5})?$/;

/** 4px dot separator between status-bar stats. */
function Dot() {
  return <span style={{ width: 4, height: 4, borderRadius: '50%', background: 'var(--ex-dot)', flex: '0 0 4px' }} />;
}

/** Vertical divider between toolbar groups. */
function Divider() {
  return <span style={{ width: 1, height: 20, background: 'var(--ex-border-2)', margin: '0 5px', flex: '0 0 1px' }} />;
}

function PlusIcon({ size = 13 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 14 14" fill="none">
      <path d="M7 2.5v9M2.5 7h9" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" />
    </svg>
  );
}

function CloseIcon({ size = 13 }: { size?: number }) {
  return (
    <svg width={size} height={size} viewBox="0 0 14 14" fill="none">
      <path d="M3.5 3.5l7 7M10.5 3.5l-7 7" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" />
    </svg>
  );
}

/** Formula functions supported by the engine, grouped for the help modal. */
const FUNCTION_HELP: Array<{ group: string; items: Array<{ sig: string; desc: string }> }> = [
  {
    group: 'Aggregate',
    items: [
      { sig: 'SUM(range…)', desc: 'Add all numbers in the arguments.' },
      { sig: 'AVERAGE(range…)', desc: 'Arithmetic mean of the numbers (alias: AVG).' },
      { sig: 'MIN(range…)', desc: 'Smallest number.' },
      { sig: 'MAX(range…)', desc: 'Largest number.' },
      { sig: 'COUNT(range…)', desc: 'How many values are numeric.' },
      { sig: 'COUNTA(range…)', desc: 'How many cells are non-empty.' },
      { sig: 'MEDIAN(range…)', desc: 'Middle value of the numbers.' },
      { sig: 'PRODUCT(range…)', desc: 'Multiply all numbers together.' },
    ],
  },
  {
    group: 'Math',
    items: [
      { sig: 'ROUND(n, digits?)', desc: 'Round n to the given decimal places.' },
      { sig: 'ABS(n)', desc: 'Absolute value.' },
      { sig: 'SQRT(n)', desc: 'Square root.' },
      { sig: 'POWER(base, exp)', desc: 'base raised to exp (same as base ^ exp).' },
      { sig: 'MOD(a, b)', desc: 'Remainder of a divided by b.' },
      { sig: 'INT(n)', desc: 'Round down to an integer.' },
      { sig: 'PI()', desc: 'The constant π.' },
    ],
  },
  {
    group: 'Logic',
    items: [
      { sig: 'IF(test, a, b?)', desc: 'a when test is true, otherwise b.' },
      { sig: 'AND(range…)', desc: 'True when every argument is truthy.' },
      { sig: 'OR(range…)', desc: 'True when any argument is truthy.' },
      { sig: 'NOT(x)', desc: 'Invert a boolean.' },
    ],
  },
  {
    group: 'Text',
    items: [
      { sig: 'CONCAT(range…)', desc: 'Join values into one string (or use a & b).' },
      { sig: 'LEN(text)', desc: 'Number of characters.' },
      { sig: 'UPPER(text)', desc: 'Uppercase.' },
      { sig: 'LOWER(text)', desc: 'Lowercase.' },
      { sig: 'TRIM(text)', desc: 'Remove leading/trailing whitespace.' },
    ],
  },
  {
    group: 'Operators',
    items: [
      { sig: '+  -  *  /  ^  %', desc: 'Arithmetic; ^ is power, % is modulo.' },
      { sig: '&', desc: 'Concatenate two values as text.' },
      { sig: '=  <>  <  >  <=  >=', desc: 'Comparisons; return TRUE / FALSE.' },
      { sig: 'A1  ·  A1:B9', desc: 'A single cell, or a rectangular range.' },
    ],
  },
];

/** Preset cell fills — 20 soft, absolute tints (kept as data, so theme-independent). */
const FILL_PRESETS: Array<{ label: string; color: string }> = [
  { label: 'Rose', color: '#F7D9D5' },
  { label: 'Peach', color: '#F9E0CE' },
  { label: 'Apricot', color: '#F6E7CC' },
  { label: 'Butter', color: '#FBF3C0' },
  { label: 'Lemon', color: '#F3F0BE' },
  { label: 'Lime', color: '#E7F0C8' },
  { label: 'Sage', color: '#DCEBD0' },
  { label: 'Mint', color: '#D3EBDE' },
  { label: 'Seafoam', color: '#CFEAE4' },
  { label: 'Aqua', color: '#D2EAEC' },
  { label: 'Sky', color: '#D6E6F2' },
  { label: 'Periwinkle', color: '#DCE0F3' },
  { label: 'Lavender', color: '#E4DEF3' },
  { label: 'Lilac', color: '#EBDCF1' },
  { label: 'Mauve', color: '#F1DCEC' },
  { label: 'Pink', color: '#F6DCE8' },
  { label: 'Sand', color: '#EEE7D6' },
  { label: 'Oat', color: '#E9E3D2' },
  { label: 'Stone', color: '#E7E4DB' },
  { label: 'Cloud', color: '#E2E4E7' },
];

/** Preset text colors — 20 inks. Neutral defaults adapt to the theme. */
const TEXT_PRESETS: Array<{ label: string; color: string }> = [
  { label: 'Ink', color: 'var(--ex-text)' },
  { label: 'Slate', color: 'var(--ex-text-muted)' },
  { label: 'Gray', color: 'var(--ex-muted-2)' },
  { label: 'Charcoal', color: '#374151' },
  { label: 'Steel', color: '#6B7280' },
  { label: 'Rust', color: '#B4451F' },
  { label: 'Orange', color: '#C2410C' },
  { label: 'Amber', color: '#B45309' },
  { label: 'Gold', color: '#A16207' },
  { label: 'Olive', color: '#4D7C0F' },
  { label: 'Green', color: '#2F6F5E' },
  { label: 'Emerald', color: '#059669' },
  { label: 'Teal', color: '#0F766E' },
  { label: 'Cyan', color: '#0E7490' },
  { label: 'Blue', color: '#1F4FD8' },
  { label: 'Indigo', color: '#3730A3' },
  { label: 'Purple', color: '#6B21A8' },
  { label: 'Fuchsia', color: '#A21CAF' },
  { label: 'Pink', color: '#BE185D' },
  { label: 'Crimson', color: '#9F1239' },
];

/**
 * Guaranteed fonts (Mono is the default, stored as ''). The app's own faces are
 * bundled and the generic ones always resolve; everything else is detected at
 * runtime so the list only ever shows fonts that are actually available.
 */
const DEFAULT_FONT = 'var(--ex-font-mono)';
const FONTS: Array<{ label: string; value: string }> = [
  { label: 'Mono', value: DEFAULT_FONT },
  { label: 'Sans', value: 'var(--ex-font-sans)' },
  { label: 'Serif', value: 'var(--ex-font-serif)' },
  { label: 'System', value: 'system-ui, sans-serif' },
];

/** Common fonts across Windows/macOS/Linux to probe for (no permission needed). */
const FONT_CANDIDATES: string[] = [
  'Arial', 'Georgia', 'Times New Roman', 'Courier New', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Times', 'Courier',
  'Andale Mono', 'Arial Black', 'Arial Narrow', 'Avenir', 'Avenir Next', 'Baskerville', 'Big Caslon',
  'Bodoni 72', 'Book Antiqua', 'Bookman Old Style', 'Brush Script MT', 'Calibri', 'Cambria', 'Candara',
  'Cantarell', 'Century Gothic', 'Chalkboard', 'Chalkduster', 'Cochin', 'Comic Sans MS', 'Consolas',
  'Constantia', 'Copperplate', 'Corbel', 'DejaVu Sans', 'DejaVu Sans Mono', 'DejaVu Serif', 'Didot',
  'Fira Code', 'Fira Sans', 'Franklin Gothic Medium', 'Futura', 'Garamond', 'Geneva', 'Gill Sans',
  'Helvetica', 'Helvetica Neue', 'Hoefler Text', 'IBM Plex Mono', 'IBM Plex Sans', 'Impact', 'Inconsolata',
  'Iowan Old Style', 'Lato', 'Liberation Mono', 'Liberation Sans', 'Liberation Serif', 'Lucida Console',
  'Lucida Grande', 'Lucida Sans Unicode', 'Marker Felt', 'Menlo', 'Monaco', 'Montserrat', 'Noto Sans',
  'Noto Serif', 'Open Sans', 'Optima', 'Palatino', 'Palatino Linotype', 'Papyrus', 'PT Sans', 'PT Serif',
  'Roboto', 'Roboto Mono', 'Rockwell', 'SF Mono', 'SF Pro', 'Segoe UI', 'Source Code Pro', 'Source Sans Pro',
  'Source Serif Pro', 'Tahoma', 'Trebuchet MS', 'Ubuntu', 'Ubuntu Mono', 'Zapfino',
];

/**
 * Detect which candidate fonts are installed by comparing rendered text width
 * against generic fallbacks — an installed font measures differently. Permission
 * free (no queryLocalFonts prompt) and works in every browser.
 */
function detectAvailableFonts(candidates: string[]): string[] {
  if (typeof document === 'undefined') return [];
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');
  if (!ctx) return [];
  const test = 'mmmmmmmmmmlliWQ019';
  const size = '72px';
  const bases = ['monospace', 'serif', 'sans-serif'];
  const baseW: Record<string, number> = {};
  for (const b of bases) {
    ctx.font = `${size} ${b}`;
    baseW[b] = ctx.measureText(test).width;
  }
  const out: string[] = [];
  for (const f of candidates) {
    for (const b of bases) {
      ctx.font = `${size} "${f}", ${b}`;
      if (ctx.measureText(test).width !== baseW[b]) {
        out.push(f);
        break;
      }
    }
  }
  return out.sort((a, b) => a.localeCompare(b));
}

/** Number-format options for the toolbar dropdown. */
const FORMAT_OPTIONS: Array<{ label: string; value: string }> = [
  { label: 'Automatic', value: 'auto' },
  { label: 'Integer', value: 'int' },
  { label: 'Decimal · 0.00', value: 'n2' },
  { label: 'Thousands · 1,000', value: 'k' },
  { label: 'Currency · $', value: 'cur' },
  { label: 'Percent · %', value: 'pct' },
];

const ICON = 'var(--ex-icon)';

/** Flat icon-button base (Google-docs style: no chrome until hover/active). */
function iconBtnStyle(active: boolean, accent: string): CSSProperties {
  return {
    display: 'inline-flex',
    alignItems: 'center',
    justifyContent: 'center',
    height: 28,
    minWidth: 28,
    padding: '0 4px',
    gap: 3,
    borderRadius: 5,
    border: '1px solid transparent',
    background: active ? mix(accent, 13, 'var(--ex-surface)') : 'transparent',
    color: active ? accent : ICON,
    cursor: 'pointer',
    outline: 'none',
  };
}

const caret = (
  <svg width="8" height="8" viewBox="0 0 8 8" fill="none" style={{ marginLeft: 1 }}>
    <path d="M1.5 3l2.5 2.5L6.5 3" stroke="currentColor" strokeWidth="1.1" strokeLinecap="round" strokeLinejoin="round" />
  </svg>
);

function BorderGlyph({ mode, size = 15 }: { mode: BorderMode; size?: number }) {
  const L = 'var(--ex-glyph-faint)';
  const D = ICON;
  const e =
    mode === 'all' || mode === 'outer'
      ? { t: 1, r: 1, b: 1, l: 1 }
      : mode === 'top'
      ? { t: 1 }
      : mode === 'bottom'
      ? { b: 1 }
      : mode === 'left'
      ? { l: 1 }
      : mode === 'right'
      ? { r: 1 }
      : {};
  const t = (e as any).t,
    r = (e as any).r,
    b = (e as any).b,
    l = (e as any).l;
  return (
    <svg width={size} height={size} viewBox="0 0 15 15" fill="none">
      <rect x="2.5" y="2.5" width="10" height="10" stroke={L} strokeWidth="1" />
      {mode === 'all' && (
        <g stroke={D} strokeWidth="1">
          <line x1="7.5" y1="2.5" x2="7.5" y2="12.5" />
          <line x1="2.5" y1="7.5" x2="12.5" y2="7.5" />
        </g>
      )}
      {t && <line x1="2.5" y1="2.5" x2="12.5" y2="2.5" stroke={D} strokeWidth="1.7" />}
      {b && <line x1="2.5" y1="12.5" x2="12.5" y2="12.5" stroke={D} strokeWidth="1.7" />}
      {l && <line x1="2.5" y1="2.5" x2="2.5" y2="12.5" stroke={D} strokeWidth="1.7" />}
      {r && <line x1="12.5" y1="2.5" x2="12.5" y2="12.5" stroke={D} strokeWidth="1.7" />}
      {mode === 'none' && <line x1="3.5" y1="11.5" x2="11.5" y2="3.5" stroke={D} strokeWidth="1.2" />}
    </svg>
  );
}

const POPOVER: CSSProperties = {
  position: 'absolute',
  top: 'calc(100% + 6px)',
  left: 0,
  zIndex: 41,
  padding: 10,
  background: 'var(--ex-popover)',
  border: '1px solid var(--ex-border-2)',
  borderRadius: 8,
  boxShadow: '0 10px 28px rgba(26,26,24,0.14)',
};

/**
 * Close a popover on outside click / Escape, via a wrapper ref instead of a
 * full-screen backdrop — the backdrop otherwise intercepts clicks meant for the
 * popover (e.g. the native color input), dismissing it mid-interaction.
 */
function useDismiss(open: boolean, close: () => void) {
  const ref = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) close();
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') close();
    };
    document.addEventListener('mousedown', onDown);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('mousedown', onDown);
      document.removeEventListener('keydown', onKey);
    };
  }, [open, close]);
  return ref;
}

/** A trigger button + swatch popover, shared by the Text-color and Fill controls. */
function ColorControl({
  kind,
  current,
  presets,
  accent,
  tooltip,
  onPick,
  onClear,
  clearLabel,
}: {
  kind: 'text' | 'fill';
  current?: string;
  presets: Array<{ label: string; color: string }>;
  accent: string;
  tooltip: string;
  onPick: (color: string) => void;
  onClear: () => void;
  clearLabel: string;
}) {
  const [open, setOpen] = useState(false);
  const ref = useDismiss(open, () => setOpen(false));
  const colorInputRef = useRef<HTMLInputElement>(null);
  const bar = current || (kind === 'text' ? ICON : 'var(--ex-glyph-faint)');

  // Open the native picker imperatively. The input lives OUTSIDE the popover so
  // closing the popover never tears the OS picker down mid-drag.
  const openCustom = () => {
    const el = colorInputRef.current;
    if (!el) return;
    el.value = current && current[0] === '#' ? current : kind === 'text' ? '#1a1a18' : '#fbf3c0';
    el.click();
  };

  return (
    <div ref={ref} style={{ position: 'relative' }}>
      <input
        ref={colorInputRef}
        type="color"
        onChange={(e) => onPick(e.target.value)}
        style={{ position: 'absolute', left: 8, bottom: 0, width: 1, height: 1, opacity: 0, pointerEvents: 'none' }}
        tabIndex={-1}
        aria-hidden
      />
      <button {...tip(tooltip)} className={styles.tbtn} onClick={() => setOpen((v) => !v)} style={iconBtnStyle(open, accent)}>
        {kind === 'text' ? (
          <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', lineHeight: 1 }}>
            <span style={{ font: "700 13px/1 var(--ex-font-sans)" }}>A</span>
            <span style={{ width: 15, height: 3, marginTop: 1, borderRadius: 1, background: bar }} />
          </span>
        ) : (
          <span style={{ display: 'inline-flex', flexDirection: 'column', alignItems: 'center', lineHeight: 1 }}>
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
              <path
                d="M6.4 2.2l6 6c.4.4.4 1 0 1.4l-4.2 4.2c-.4.4-1 .4-1.4 0L1.5 8.1c-.4-.4-.4-1 0-1.4l3.5-3.5"
                stroke="currentColor"
                strokeWidth="1.2"
                strokeLinejoin="round"
              />
              <path d="M4.6 4l2.9 2.9M2 8.4h9" stroke="currentColor" strokeWidth="1.2" strokeLinecap="round" />
            </svg>
            <span style={{ width: 15, height: 3, marginTop: 1, borderRadius: 1, background: bar }} />
          </span>
        )}
      </button>
      {open && (
        <div style={POPOVER}>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(4, 22px)', gap: 6 }}>
              {presets.map((p) => (
                <button
                  key={p.color}
                  {...tip(p.label)}
                  onClick={() => {
                    onPick(p.color);
                    setOpen(false);
                  }}
                  style={{
                    width: 22,
                    height: 22,
                    borderRadius: 5,
                    border: current === p.color ? `2px solid ${accent}` : '1px solid var(--ex-swatch-border)',
                    background: p.color,
                    cursor: 'pointer',
                    padding: 0,
                    outline: 'none',
                  }}
                />
              ))}
            </div>
            {/* Custom color via the native picker (opened imperatively) */}
            <button
              {...tip('Pick a custom color')}
              onClick={openCustom}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 8,
                marginTop: 8,
                width: '100%',
                padding: '4px 6px',
                borderRadius: 5,
                border: '1px solid var(--ex-border-2)',
                background: 'var(--ex-surface)',
                color: 'var(--ex-text-muted)',
                font: "500 12px/1 var(--ex-font-sans)",
                cursor: 'pointer',
                outline: 'none',
              }}
            >
              <span
                style={{
                  width: 20,
                  height: 20,
                  borderRadius: 4,
                  border: '1px solid var(--ex-swatch-border)',
                  background: current && current[0] === '#' ? current : 'var(--ex-popover)',
                  flex: '0 0 20px',
                }}
              />
              Custom…
            </button>
            <button
              onClick={() => {
                onClear();
                setOpen(false);
              }}
              style={{
                marginTop: 6,
                width: '100%',
                padding: '5px 0',
                borderRadius: 5,
                border: '1px solid var(--ex-border-2)',
                background: 'var(--ex-surface-3)',
                color: 'var(--ex-text-muted)',
                font: "500 12px/1 var(--ex-font-sans)",
                cursor: 'pointer',
                outline: 'none',
              }}
            >
              {clearLabel}
            </button>
        </div>
      )}
    </div>
  );
}

/** Borders trigger + popover of border modes. */
function BordersControl({ accent, onApply }: { accent: string; onApply: (mode: BorderMode) => void }) {
  const [open, setOpen] = useState(false);
  const ref = useDismiss(open, () => setOpen(false));
  const modes: Array<{ mode: BorderMode; label: string }> = [
    { mode: 'all', label: 'All borders' },
    { mode: 'outer', label: 'Outer border' },
    { mode: 'top', label: 'Top border' },
    { mode: 'bottom', label: 'Bottom border' },
    { mode: 'left', label: 'Left border' },
    { mode: 'right', label: 'Right border' },
  ];
  return (
    <div ref={ref} style={{ position: 'relative' }}>
      <button {...tip('Borders')} className={styles.tbtn} onClick={() => setOpen((v) => !v)} style={iconBtnStyle(open, accent)}>
        <BorderGlyph mode="all" size={16} />
      </button>
      {open && (
        <div style={POPOVER}>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(3, 30px)', gap: 4 }}>
              {modes.map((m) => (
                <button
                  key={m.mode}
                  {...tip(m.label)}
                  onClick={() => {
                    onApply(m.mode);
                    setOpen(false);
                  }}
                  className={styles.tbtn}
                  style={{ ...iconBtnStyle(false, accent), width: 30, height: 28 }}
                >
                  <BorderGlyph mode={m.mode} />
                </button>
              ))}
            </div>
            <button
              onClick={() => {
                onApply('none');
                setOpen(false);
              }}
              style={{
                marginTop: 8,
                width: '100%',
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'center',
                gap: 6,
                padding: '5px 0',
                borderRadius: 5,
                border: '1px solid var(--ex-border-2)',
                background: 'var(--ex-surface-3)',
                color: 'var(--ex-text-muted)',
                font: "500 12px/1 var(--ex-font-sans)",
                cursor: 'pointer',
                outline: 'none',
              }}
            >
              <BorderGlyph mode="none" /> Clear borders
            </button>
        </div>
      )}
    </div>
  );
}

/** Alignment trigger (shows current) + caret dropdown. */
function AlignControl({ value, accent, onPick }: { value: 'left' | 'center' | 'right'; accent: string; onPick: (a: 'left' | 'center' | 'right') => void }) {
  const [open, setOpen] = useState(false);
  const ref = useDismiss(open, () => setOpen(false));
  const glyph = (a: 'left' | 'center' | 'right') => {
    const d = a === 'left' ? 'M2.5 4h10M2.5 7.5h6M2.5 11h8.5' : a === 'center' ? 'M2.5 4h10M4.5 7.5h6M3.5 11h8' : 'M2.5 4h10M6.5 7.5h6M4 11h8.5';
    return (
      <svg width="15" height="15" viewBox="0 0 15 15" fill="none">
        <path d={d} stroke="currentColor" strokeWidth="1.35" strokeLinecap="round" />
      </svg>
    );
  };
  const items: Array<'left' | 'center' | 'right'> = ['left', 'center', 'right'];
  return (
    <div ref={ref} style={{ position: 'relative' }}>
      <button {...tip('Horizontal align')} className={styles.tbtn} onClick={() => setOpen((v) => !v)} style={iconBtnStyle(open, accent)}>
        {glyph(value)}
        {caret}
      </button>
      {open && (
        <div style={{ ...POPOVER, padding: 4, display: 'flex', gap: 2 }}>
          {items.map((a) => (
            <button
              key={a}
              {...tip(`Align ${a}`)}
              onClick={() => {
                onPick(a);
                setOpen(false);
              }}
              className={styles.tbtn}
              style={iconBtnStyle(value === a, accent)}
            >
              {glyph(a)}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

/** A compact custom dropdown (trigger + chevron + option list), matching FontControl. */
function MenuSelect({
  value,
  options,
  accent,
  tooltip,
  width = 160,
  onPick,
}: {
  value: string;
  options: Array<{ label: string; value: string }>;
  accent: string;
  tooltip?: string;
  width?: number;
  onPick: (value: string) => void;
}) {
  const [open, setOpen] = useState(false);
  const ref = useDismiss(open, () => setOpen(false));
  const current = options.find((o) => o.value === value) ?? options[0];
  return (
    <div ref={ref} style={{ position: 'relative' }}>
      <button
        {...(tooltip ? tip(tooltip) : {})}
        className={styles.tbtn}
        onClick={() => setOpen((v) => !v)}
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: 6,
          height: 23,
          width,
          padding: '0 8px',
          borderRadius: 5,
          border: `1px solid ${open ? accent : 'var(--ex-border-2)'}`,
          background: 'var(--ex-surface)',
          color: 'var(--ex-select-text)',
          cursor: 'pointer',
          outline: 'none',
        }}
      >
        <span style={{ flex: 1, textAlign: 'left', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', font: "400 12.5px/1 var(--ex-font-sans)" }}>
          {current?.label ?? ''}
        </span>
        {caret}
      </button>
      {open && (
        <div style={{ ...POPOVER, minWidth: width, padding: 4 }}>
          {options.map((o) => {
            const selected = o.value === value;
            return (
              <button
                key={o.value}
                onClick={() => {
                  onPick(o.value);
                  setOpen(false);
                }}
                className={styles.tbtn}
                style={{
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                  gap: 8,
                  width: '100%',
                  padding: '7px 9px',
                  border: 0,
                  borderRadius: 5,
                  background: selected ? mix(accent, 12, 'var(--ex-popover)') : 'transparent',
                  color: selected ? accent : 'var(--ex-text)',
                  cursor: 'pointer',
                  outline: 'none',
                  textAlign: 'left',
                  font: "400 12.5px/1 var(--ex-font-sans)",
                }}
              >
                <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' }}>{o.label}</span>
                {selected && (
                  <svg width="13" height="13" viewBox="0 0 14 14" fill="none" style={{ flex: '0 0 13px' }}>
                    <path d="M2.5 7.5l3 3 6-6.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
                  </svg>
                )}
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}

/** Freeze rows/columns — trigger + popover with per-axis options. */
function FreezeControl({
  accent,
  frozenRows,
  frozenCols,
  selR2,
  selC2,
  onFreeze,
}: {
  accent: string;
  frozenRows: number;
  frozenCols: number;
  selR2: number;
  selC2: number;
  onFreeze: (rows: number, cols: number) => void;
}) {
  const [open, setOpen] = useState(false);
  const ref = useDismiss(open, () => setOpen(false));
  const active = frozenRows > 0 || frozenCols > 0;

  const heading: CSSProperties = {
    padding: '4px 9px 3px',
    font: "500 10.5px/1 var(--ex-font-mono)",
    letterSpacing: '0.12em',
    textTransform: 'uppercase',
    color: 'var(--ex-text-faint)',
  };
  const opt = (selected: boolean): CSSProperties => ({
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'space-between',
    gap: 8,
    width: '100%',
    padding: '7px 9px',
    border: 0,
    borderRadius: 5,
    background: selected ? mix(accent, 12, 'var(--ex-popover)') : 'transparent',
    color: selected ? accent : 'var(--ex-text)',
    cursor: 'pointer',
    outline: 'none',
    textAlign: 'left',
    font: "400 12.5px/1 var(--ex-font-sans)",
  });
  const check = (
    <svg width="13" height="13" viewBox="0 0 14 14" fill="none" style={{ flex: '0 0 13px' }}>
      <path d="M2.5 7.5l3 3 6-6.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
  const rowOpts = [
    { label: 'No frozen rows', val: 0 },
    { label: 'Freeze 1 row', val: 1 },
    ...(selR2 + 1 > 1 ? [{ label: `Freeze up to row ${selR2 + 1}`, val: selR2 + 1 }] : []),
  ];
  const colOpts = [
    { label: 'No frozen columns', val: 0 },
    { label: 'Freeze 1 column', val: 1 },
    ...(selC2 + 1 > 1 ? [{ label: `Freeze up to column ${colName(selC2)}`, val: selC2 + 1 }] : []),
  ];

  return (
    <div ref={ref} style={{ position: 'relative' }}>
      <button {...tip('Freeze rows / columns')} className={styles.tbtn} onClick={() => setOpen((v) => !v)} style={iconBtnStyle(active, accent)}>
        <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
          <rect x="2" y="2" width="12" height="12" rx="1" stroke="currentColor" strokeWidth="1" />
          <path d="M2 6h12M6 2v12" stroke="currentColor" strokeWidth="1.6" />
        </svg>
      </button>
      {open && (
        <div style={{ ...POPOVER, width: 200, padding: 4 }}>
          <div style={heading}>Rows</div>
          {rowOpts.map((o) => (
            <button
              key={o.val}
              onClick={() => {
                onFreeze(o.val, frozenCols);
                setOpen(false);
              }}
              style={opt(frozenRows === o.val)}
            >
              {o.label}
              {frozenRows === o.val && check}
            </button>
          ))}
          <div style={{ ...heading, marginTop: 4 }}>Columns</div>
          {colOpts.map((o) => (
            <button
              key={o.val}
              onClick={() => {
                onFreeze(frozenRows, o.val);
                setOpen(false);
              }}
              style={opt(frozenCols === o.val)}
            >
              {o.label}
              {frozenCols === o.val && check}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

/** Font-family trigger + searchable, previewed popover. */
function FontControl({
  value,
  accent,
  fonts,
  systemFonts,
  onPick,
}: {
  value: string;
  accent: string;
  fonts: Array<{ label: string; value: string }>;
  systemFonts: string[];
  onPick: (value: string) => void;
}) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState('');
  const ref = useDismiss(open, () => setOpen(false));

  const all = [...fonts, ...systemFonts.map((f) => ({ label: f, value: `"${f}"` }))];
  const currentLabel = all.find((o) => o.value === value)?.label ?? (value[0] === '"' ? value.slice(1, -1) : 'Mono');
  const q = query.trim().toLowerCase();
  const filtered = q ? all.filter((o) => o.label.toLowerCase().includes(q)) : all;

  return (
    <div ref={ref} style={{ position: 'relative' }}>
      <button
        {...tip('Font')}
        className={styles.tbtn}
        onClick={() => setOpen((v) => !v)}
        style={{
          display: 'flex',
          alignItems: 'center',
          gap: 6,
          height: 23,
          width: 118,
          padding: '0 8px',
          borderRadius: 5,
          border: `1px solid ${open ? accent : 'var(--ex-border-2)'}`,
          background: 'var(--ex-surface)',
          color: 'var(--ex-select-text)',
          cursor: 'pointer',
          outline: 'none',
        }}
      >
        <span style={{ flex: 1, minWidth: 0, textAlign: 'left', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', font: `400 12.5px/1 ${value}` }}>{currentLabel}</span>
        {caret}
      </button>
      {open && (
        <div style={{ ...POPOVER, width: 230, padding: 0, overflow: 'hidden' }}>
          <div style={{ padding: 8, borderBottom: '1px solid var(--ex-border)' }}>
            <input
              autoFocus
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder="Search fonts…"
              style={{
                width: '100%',
                padding: '6px 9px',
                borderRadius: 5,
                border: '1px solid var(--ex-border-2)',
                background: 'var(--ex-surface)',
                color: 'var(--ex-text)',
                font: "400 12.5px/1 var(--ex-font-sans)",
                outline: 'none',
              }}
            />
          </div>
          <div style={{ maxHeight: 300, overflowY: 'auto', padding: 4 }}>
            {filtered.map((o) => {
              const selected = o.value === value;
              return (
                <button
                  key={o.value}
                  onClick={() => {
                    onPick(o.value);
                    setOpen(false);
                  }}
                  className={styles.tbtn}
                  style={{
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'space-between',
                    gap: 8,
                    width: '100%',
                    padding: '7px 9px',
                    border: 0,
                    borderRadius: 5,
                    background: selected ? mix(accent, 12, 'var(--ex-popover)') : 'transparent',
                    color: selected ? accent : 'var(--ex-text)',
                    cursor: 'pointer',
                    outline: 'none',
                    textAlign: 'left',
                  }}
                >
                  <span style={{ overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', font: `400 13px/1.2 ${o.value}` }}>{o.label}</span>
                  {selected && (
                    <svg width="13" height="13" viewBox="0 0 14 14" fill="none" style={{ flex: '0 0 13px' }}>
                      <path d="M2.5 7.5l3 3 6-6.5" stroke="currentColor" strokeWidth="1.4" strokeLinecap="round" strokeLinejoin="round" />
                    </svg>
                  )}
                </button>
              );
            })}
            {filtered.length === 0 && (
              <div style={{ padding: '10px 9px', color: 'var(--ex-text-faint)', font: "400 12px/1 var(--ex-font-sans)" }}>No fonts match “{query}”.</div>
            )}
          </div>
        </div>
      )}
    </div>
  );
}

/**
 * Decide whether the caret sits somewhere a cell reference can be inserted
 * ("point mode"). Returns the text before/after the reference slot, or null.
 *
 * A reference is insertable when the formula text so far ends with '=', an
 * operator, '(' or ',' — optionally with a just-placed reference right before
 * the caret, which a fresh click replaces (drag-to-extend, click-to-repick).
 */
function pointContext(text: string, caret: number): PointAnchor | null {
  if (!text.startsWith('=')) return null;
  const before = text.slice(0, caret);
  const suffix = text.slice(caret);
  const m = TRAILING_REF.exec(before);
  const prefix = m ? before.slice(0, m.index) : before;
  const trimmed = prefix.replace(/\s+$/, '');
  const canInsert = trimmed === '=' || '=(,+-*/^%&<>:'.indexOf(trimmed.slice(-1)) >= 0;
  return canInsert ? { prefix, suffix } : null;
}

/**
 * Reusable spreadsheet component. Owns a SpreadsheetEngine + StorageAdapter and
 * renders the full Ledger UI (title, toolbar, formula bar, grid, status bar).
 * All spreadsheet logic lives in the engine; this file is presentation + input.
 */
export function Sheet(props: SheetProps) {
  const accent = props.accent ?? '#2F6F5E';

  const storage = props.storage;
  const workbook = useMemo(
    () => new Workbook({ rows: props.rows, cols: props.cols, defaultName: props.defaultTitle }),
    [props.rows, props.cols, props.defaultTitle],
  );
  const { savedLabel } = useWorkbook(workbook, storage);
  const engine = workbook.getActive();
  const { theme, toggle: toggleTheme } = useTheme();

  // Detect installed fonts once (permission-free width probing).
  useEffect(() => {
    setSystemFonts(detectAvailableFonts(FONT_CANDIDATES));
  }, []);

  // Reflect the workbook name in the browser tab: "<name> - Sheet".
  const workbookName = workbook.getName();
  useEffect(() => {
    document.title = `${workbookName?.trim() || 'Untitled'} — Sheet`;
  }, [workbookName]);

  // Track the scroll viewport width so the sticky bottom bars span exactly it.
  useLayoutEffect(() => {
    const el = scrollRef.current;
    if (!el) return;
    const update = () => setViewportW((w) => (w !== el.clientWidth ? el.clientWidth : w));
    update();
    const ro = new ResizeObserver(update);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  // --- transient UI state (selection / editing) lives in React, not the engine ---
  const [sel, setSel] = useState<Selection>({ r: 0, c: 0, r2: 0, c2: 0 });
  const [editing, setEditing] = useState(false);
  const [editText, setEditText] = useState('');
  const [editSource, setEditSource] = useState<EditSource>(null);

  // Point mode: highlighted range being inserted into the formula via the mouse.
  const [pointRange, setPointRange] = useState<Range | null>(null);
  // Formula-reference modal open state.
  const [helpOpen, setHelpOpen] = useState(false);
  // Reset-all confirmation modal open state.
  const [confirmReset, setConfirmReset] = useState(false);
  // "Insert rows/columns" modal (count + before/after position).
  const [insertAxis, setInsertAxis] = useState<'rows' | 'cols' | null>(null);
  // Inline sheet-tab rename ({ id, text }) and delete confirmation (sheet id).
  const [renaming, setRenaming] = useState<{ id: string; text: string } | null>(null);
  const [confirmDeleteSheet, setConfirmDeleteSheet] = useState<string | null>(null);
  // System fonts detected via width probing.
  const [systemFonts, setSystemFonts] = useState<string[]>([]);
  // Visible width of the scroll viewport — sizes the sticky bottom bars.
  const [viewportW, setViewportW] = useState(0);

  const scrollRef = useRef<HTMLDivElement>(null);
  const editInputRef = useRef<HTMLInputElement | null>(null);
  const barInputRef = useRef<HTMLInputElement | null>(null);
  const pointing = useRef<PointAnchor | null>(null);
  const caret = useRef(0);
  const pendingCaret = useRef(false);

  // Latest view state, readable from document-level listeners without stale closures.
  const view = useRef<ViewState>({ sel, editing, editText, editSource });
  view.current = { sel, editing, editText, editSource };

  /** The input currently being edited (inline overlay or the formula bar). */
  function activeEditInput(): HTMLInputElement | null {
    return view.current.editSource === 'bar' ? barInputRef.current : editInputRef.current;
  }

  // After a pointed reference is spliced in, restore the caret just past it
  // (a controlled input otherwise resets the caret to the end on re-render).
  useLayoutEffect(() => {
    if (!pendingCaret.current) return;
    pendingCaret.current = false;
    const inp = activeEditInput();
    if (inp) {
      inp.focus();
      inp.setSelectionRange(caret.current, caret.current);
    }
  }, [editText]);

  /** Splice the reference for anchor..(r,c) into the formula being edited. */
  function applyPoint(anchor: { r: number; c: number }, r: number, c: number) {
    const p = pointing.current;
    if (!p) return;
    const r1 = Math.min(anchor.r, r);
    const r2 = Math.max(anchor.r, r);
    const c1 = Math.min(anchor.c, c);
    const c2 = Math.max(anchor.c, c);
    const single = r1 === r2 && c1 === c2;
    const ref = single
      ? `${colName(c1)}${r1 + 1}`
      : `${colName(c1)}${r1 + 1}:${colName(c2)}${r2 + 1}`;
    const text = p.prefix + ref + p.suffix;
    caret.current = p.prefix.length + ref.length;
    pendingCaret.current = true;
    setEditText(text);
    setPointRange({ r1, r2, c1, c2 });
  }

  const nRows = engine.rows;
  const nCols = engine.cols;
  // Effective selection expands to fully cover any merged regions it touches.
  const range = engine.expandRange(selectionToRange(sel));

  /* ---------- column geometry (per-column widths) ---------- */

  const cw = (c: number) => engine.getColWidth(c);
  // Absolute left offset (in inner-grid coords) of each column; colX[nCols] = total.
  // Recomputed each render so live column resizing reflects immediately.
  const colX: number[] = new Array(nCols + 1);
  colX[0] = GUTW;
  for (let c = 0; c < nCols; c++) colX[c + 1] = colX[c] + cw(c);

  /* ---------- navigation / selection ---------- */

  function scrollInto(r: number, c: number) {
    const el = scrollRef.current;
    if (!el) return;
    const x = colX[c];
    const y = HDRH + r * ROWH;
    if (x - GUTW < el.scrollLeft) el.scrollLeft = Math.max(0, x - GUTW);
    if (x + cw(c) > el.scrollLeft + el.clientWidth) el.scrollLeft = x + cw(c) - el.clientWidth;
    if (y - HDRH < el.scrollTop) el.scrollTop = Math.max(0, y - HDRH);
    if (y + ROWH > el.scrollTop + el.clientHeight) el.scrollTop = y + ROWH - el.clientHeight;
  }

  function select(r: number, c: number, extend: boolean) {
    const R = Math.max(0, Math.min(nRows - 1, r));
    const C = Math.max(0, Math.min(nCols - 1, c));
    setSel((prev) => (extend ? { r: prev.r, c: prev.c, r2: R, c2: C } : { r: R, c: C, r2: R, c2: C }));
    if (!extend) scrollInto(R, C);
  }

  /* ---------- editing ---------- */

  function startEdit(seed: string | undefined, source: EditSource = 'grid') {
    const s = view.current.sel;
    setEditing(true);
    setEditSource(source);
    setEditText(seed != null ? seed : engine.raw(s.r, s.c));
  }

  function endPointing() {
    pointing.current = null;
    setPointRange(null);
  }

  function commit(move?: [number, number]) {
    const { sel: s, editing: wasEditing, editText: text } = view.current;
    if (wasEditing) engine.setRaw(s.r, s.c, text);
    endPointing();
    setEditing(false);
    setEditSource(null);
    setEditText('');
    if (move) select(s.r + move[0], s.c + move[1], false);
  }

  function cancelEdit() {
    endPointing();
    setEditing(false);
    setEditSource(null);
    setEditText('');
  }

  /* ---------- formatting / actions ---------- */

  function toggle(attr: 'b' | 'i' | 's') {
    const s = view.current.sel;
    const cur = engine.cellAt(s.r, s.c);
    const on = !!(cur && cur[attr]);
    engine.setAttr(selectionToRange(s), attr, on ? 0 : 1);
  }
  function applyAlign(a: 'left' | 'center' | 'right') {
    engine.setAttr(engine.expandRange(selectionToRange(view.current.sel)), 'a', a);
  }
  function applyFormat(f: NumberFormat | 'auto') {
    engine.setAttr(engine.expandRange(selectionToRange(view.current.sel)), 'f', f === 'auto' ? '' : f);
  }
  function applyFont(value: string) {
    engine.setAttr(engine.expandRange(selectionToRange(view.current.sel)), 'ff', value === DEFAULT_FONT ? '' : value);
  }
  function applyFill(color: string) {
    engine.setAttr(engine.expandRange(selectionToRange(view.current.sel)), 'g', color);
  }
  function applyText(color: string) {
    engine.setAttr(engine.expandRange(selectionToRange(view.current.sel)), 'co', color);
  }
  function applyBorders(mode: BorderMode) {
    engine.applyBorders(engine.expandRange(selectionToRange(view.current.sel)), mode, BORDER_COLOR);
  }
  function toggleGridlines() {
    engine.setGridlines(!engine.getGridlines());
  }
  function toggleMerge() {
    const b = engine.expandRange(selectionToRange(view.current.sel));
    if (view.current.editing) commit();
    if (engine.hasMergeIn(b)) engine.unmergeCells(b);
    else engine.mergeCells(b);
    setSel({ r: b.r1, c: b.c1, r2: b.r1, c2: b.c1 });
  }
  function applyFreeze(rows: number, cols: number) {
    engine.setFreeze(rows, cols);
  }
  function clearSelection() {
    engine.clearRange(engine.expandRange(selectionToRange(view.current.sel)));
  }
  function deleteSelectedRows() {
    const b = engine.expandRange(selectionToRange(view.current.sel));
    if (view.current.editing) commit();
    engine.deleteRows(b.r1, b.r2);
    const r = Math.min(b.r1, engine.rows - 1);
    setSel({ r, c: 0, r2: r, c2: engine.cols - 1 });
    setEditing(false);
    setEditText('');
    setPointRange(null);
  }
  function deleteSelectedCols() {
    const b = engine.expandRange(selectionToRange(view.current.sel));
    if (view.current.editing) commit();
    engine.deleteCols(b.c1, b.c2);
    const c = Math.min(b.c1, engine.cols - 1);
    setSel({ r: 0, c, r2: engine.rows - 1, c2: c });
    setEditing(false);
    setEditText('');
    setPointRange(null);
  }
  function doInsert(axis: 'rows' | 'cols', count: number, position: 'before' | 'after') {
    const b = engine.expandRange(selectionToRange(view.current.sel));
    if (view.current.editing) commit();
    if (axis === 'cols') {
      const at = position === 'before' ? b.c1 : b.c2 + 1;
      engine.insertCols(at, count);
      setSel({ r: 0, c: at, r2: engine.rows - 1, c2: Math.min(at + count - 1, engine.cols - 1) });
    } else {
      const at = position === 'before' ? b.r1 : b.r2 + 1;
      engine.insertRows(at, count);
      setSel({ r: at, c: 0, r2: Math.min(at + count - 1, engine.rows - 1), c2: engine.cols - 1 });
    }
    setEditing(false);
    setEditText('');
    setPointRange(null);
    setInsertAxis(null);
  }

  /* ---------- column resize ---------- */

  function startColResize(e: React.MouseEvent, c: number) {
    e.preventDefault();
    e.stopPropagation(); // don't let the grid treat this as a column selection
    const startX = e.clientX;
    const startW = cw(c);
    let checkpointed = false;
    const move = (ev: MouseEvent) => {
      if (!checkpointed) {
        engine.checkpoint(); // one undo entry for the whole drag
        checkpointed = true;
      }
      engine.setColWidth(c, startW + (ev.clientX - startX));
    };
    const up = () => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  }

  function resetView() {
    setSel({ r: 0, c: 0, r2: 0, c2: 0 });
    setEditing(false);
    setEditSource(null);
    setEditText('');
    setPointRange(null);
    pointing.current = null;
    const el = scrollRef.current;
    if (el) {
      el.scrollTop = 0;
      el.scrollLeft = 0;
    }
  }

  function doReset() {
    Promise.resolve(storage.clear?.());
    workbook.reset();
    resetView();
    setConfirmReset(false);
  }

  /* ---------- sheets ---------- */
  function switchSheet(id: string) {
    if (id === workbook.getActiveId()) return;
    if (view.current.editing) commit();
    workbook.setActive(id);
    resetView();
  }
  function addSheet() {
    if (view.current.editing) commit();
    workbook.addSheet();
    resetView();
  }
  function removeSheet(id: string) {
    workbook.removeSheet(id);
    setConfirmDeleteSheet(null);
    resetView();
  }
  function commitRename() {
    if (renaming) workbook.renameSheet(renaming.id, renaming.text);
    setRenaming(null);
  }

  // Stable indirection so document listeners always call the latest closures.
  const api = useRef({
    select,
    startEdit,
    commit,
    cancelEdit,
    toggle,
    clearSelection,
    undo: () => engine.undo(),
    redo: () => engine.redo(),
  });
  api.current = {
    select,
    startEdit,
    commit,
    cancelEdit,
    toggle,
    clearSelection,
    undo: () => engine.undo(),
    redo: () => engine.redo(),
  };

  /* ---------- keyboard + clipboard (document level) ---------- */

  useEffect(() => {
    const isField = (t: EventTarget | null) => {
      const tag = (t as HTMLElement | null)?.tagName || '';
      return tag === 'INPUT' || tag === 'SELECT' || tag === 'TEXTAREA';
    };

    const onKey = (e: KeyboardEvent) => {
      if (isField(e.target)) return;
      const s = view.current.sel;
      const meta = e.metaKey || e.ctrlKey;
      const k = e.key;
      const a = api.current;

      if (meta && k.toLowerCase() === 'z') {
        e.preventDefault();
        return e.shiftKey ? a.redo() : a.undo();
      }
      if (meta && k.toLowerCase() === 'b') {
        e.preventDefault();
        return a.toggle('b');
      }
      if (meta && k.toLowerCase() === 'i') {
        e.preventDefault();
        return a.toggle('i');
      }
      if (meta) return;

      const step: Record<string, [number, number]> = {
        ArrowUp: [-1, 0],
        ArrowDown: [1, 0],
        ArrowLeft: [0, -1],
        ArrowRight: [0, 1],
      };
      if (step[k]) {
        e.preventDefault();
        const base = e.shiftKey ? { r: s.r2, c: s.c2 } : { r: s.r, c: s.c };
        return a.select(base.r + step[k][0], base.c + step[k][1], e.shiftKey);
      }
      if (k === 'Tab') {
        e.preventDefault();
        return a.select(s.r, s.c + (e.shiftKey ? -1 : 1), false);
      }
      if (k === 'Enter' || k === 'F2') {
        // Edit the cell keeping its existing content (like double-click).
        e.preventDefault();
        return a.startEdit(undefined, 'grid');
      }
      if (k === 'Backspace' || k === 'Delete') {
        e.preventDefault();
        return a.clearSelection();
      }
      if (k === 'Home') {
        e.preventDefault();
        return a.select(e.shiftKey ? 0 : s.r, 0, false);
      }
      if (k.length === 1 && !e.altKey) {
        e.preventDefault();
        return a.startEdit(k, 'grid');
      }
    };

    const onCopy = (e: ClipboardEvent, cut: boolean) => {
      if (isField(e.target)) return;
      const rng = engine.expandRange(selectionToRange(view.current.sel));
      const lines: string[] = [];
      for (let r = rng.r1; r <= rng.r2; r++) {
        const row: string[] = [];
        for (let c = rng.c1; c <= rng.c2; c++) row.push(engine.raw(r, c));
        lines.push(row.join('\t'));
      }
      e.preventDefault();
      e.clipboardData?.setData('text/plain', lines.join('\n'));
      if (cut) engine.clearRange(rng);
    };
    const onCopyEvt = (e: ClipboardEvent) => onCopy(e, false);
    const onCutEvt = (e: ClipboardEvent) => onCopy(e, true);
    const onPaste = (e: ClipboardEvent) => {
      if (isField(e.target)) return;
      const text = e.clipboardData?.getData('text/plain');
      if (!text) return;
      e.preventDefault();
      const grid = text.replace(/\r/g, '').split('\n').map((l) => l.split('\t'));
      const s = view.current.sel;
      engine.paste({ r: s.r, c: s.c }, grid);
    };

    document.addEventListener('keydown', onKey);
    document.addEventListener('copy', onCopyEvt);
    document.addEventListener('cut', onCutEvt);
    document.addEventListener('paste', onPaste);
    return () => {
      document.removeEventListener('keydown', onKey);
      document.removeEventListener('copy', onCopyEvt);
      document.removeEventListener('cut', onCutEvt);
      document.removeEventListener('paste', onPaste);
    };
  }, [engine]);

  /* ---------- pointer (delegated on the grid) ---------- */

  function onGridMouseDown(e: React.MouseEvent) {
    const t = (e.target as HTMLElement).closest('[data-cell],[data-col],[data-row]') as HTMLElement | null;
    if (!t) return;

    // Point mode: while editing a formula, click/drag a cell to insert its ref.
    if (view.current.editing) {
      const cellKey = t.dataset.cell;
      if (cellKey != null) {
        const inp = activeEditInput();
        const pos = inp?.selectionStart ?? view.current.editText.length;
        const ctx = pointContext(view.current.editText, pos);
        if (ctx) {
          e.preventDefault(); // keep focus in the editor
          pointing.current = ctx;
          const [pr, pc] = cellKey.split(',').map(Number);
          const anchor = { r: pr, c: pc };
          applyPoint(anchor, pr, pc);
          const move = (ev: MouseEvent) => {
            const el = document.elementFromPoint(ev.clientX, ev.clientY) as HTMLElement | null;
            const cell = el?.closest?.('[data-cell]') as HTMLElement | null;
            if (!cell) return;
            const [qr, qc] = cell.dataset.cell!.split(',').map(Number);
            applyPoint(anchor, qr, qc);
          };
          const up = () => {
            document.removeEventListener('mousemove', move);
            document.removeEventListener('mouseup', up);
            pointing.current = null;
            activeEditInput()?.focus();
          };
          document.addEventListener('mousemove', move);
          document.addEventListener('mouseup', up);
          return;
        }
      }
      // Editing but not an insertable spot (or a header) → commit, then select.
      commit();
    }

    // Column headers: click selects the column, shift-click / drag extends across columns.
    if (t.dataset.col != null) {
      const c = +t.dataset.col;
      if (e.shiftKey) {
        setSel((prev) => ({ r: 0, c: prev.c, r2: nRows - 1, c2: c }));
        return;
      }
      setSel({ r: 0, c, r2: nRows - 1, c2: c });
      const move = (ev: MouseEvent) => {
        const el = document.elementFromPoint(ev.clientX, ev.clientY) as HTMLElement | null;
        const h = el?.closest?.('[data-col]') as HTMLElement | null;
        if (!h) return;
        setSel({ r: 0, c, r2: nRows - 1, c2: +h.dataset.col! });
      };
      const up = () => {
        document.removeEventListener('mousemove', move);
        document.removeEventListener('mouseup', up);
      };
      document.addEventListener('mousemove', move);
      document.addEventListener('mouseup', up);
      return;
    }
    // Row headers: click selects the row, shift-click / drag extends across rows.
    if (t.dataset.row != null) {
      const r = +t.dataset.row;
      if (e.shiftKey) {
        setSel((prev) => ({ r: prev.r, c: 0, r2: r, c2: nCols - 1 }));
        return;
      }
      setSel({ r, c: 0, r2: r, c2: nCols - 1 });
      const move = (ev: MouseEvent) => {
        const el = document.elementFromPoint(ev.clientX, ev.clientY) as HTMLElement | null;
        const h = el?.closest?.('[data-row]') as HTMLElement | null;
        if (!h) return;
        setSel({ r, c: 0, r2: +h.dataset.row!, c2: nCols - 1 });
      };
      const up = () => {
        document.removeEventListener('mousemove', move);
        document.removeEventListener('mouseup', up);
      };
      document.addEventListener('mousemove', move);
      document.addEventListener('mouseup', up);
      return;
    }

    const [pr, pc] = t.dataset.cell!.split(',').map(Number);
    if (e.shiftKey) {
      setSel((prev) => ({ r: prev.r, c: prev.c, r2: pr, c2: pc }));
      return;
    }
    // Clicking a merged cell selects the whole region (anchor = its top-left).
    const m = engine.mergeAt(pr, pc);
    if (m) setSel({ r: m.r1, c: m.c1, r2: m.r2, c2: m.c2 });
    else setSel({ r: pr, c: pc, r2: pr, c2: pc });
    scrollInto(pr, pc);

    // Second click of a double-click: let onDoubleClick open the editor without
    // wiring up a drag-select (which otherwise churns the selection).
    if (e.detail >= 2) return;

    const move = (ev: MouseEvent) => {
      const el = document.elementFromPoint(ev.clientX, ev.clientY) as HTMLElement | null;
      const cell = el?.closest?.('[data-cell]') as HTMLElement | null;
      if (!cell) return;
      const [qr, qc] = cell.dataset.cell!.split(',').map(Number);
      setSel({ r: pr, c: pc, r2: qr, c2: qc });
    };
    const up = () => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
  }

  function onGridDoubleClick(e: React.MouseEvent) {
    const t = (e.target as HTMLElement).closest('[data-cell]');
    if (t) startEdit(undefined, 'grid');
  }

  /* ---------- derived render data ---------- */

  const cols = useMemo(() => Array.from({ length: nCols }, (_, i) => i), [nCols]);
  const rows = useMemo(() => Array.from({ length: nRows }, (_, i) => i), [nRows]);

  const active = engine.cellAt(sel.r, sel.c) || {};
  const activeAlign = active.a || 'left';
  const gridlinesOn = engine.getGridlines();
  const dark = theme === 'dark';
  const frozenRows = engine.getFrozenRows();
  const frozenCols = engine.getFrozenCols();
  const mergedHere = engine.hasMergeIn(range);
  const canMerge = mergedHere || range.r1 !== range.r2 || range.c1 !== range.c2;
  const stats = engine.stats(range);
  const single = range.r1 === range.r2 && range.c1 === range.c2;
  const selLabel = single
    ? colName(sel.c) + (sel.r + 1)
    : `${colName(range.c1)}${range.r1 + 1}:${colName(range.c2)}${range.r2 + 1}  ·  ${(range.r2 - range.r1 + 1) * (range.c2 - range.c1 + 1)} cells`;
  const barValue = editing ? editText : engine.raw(sel.r, sel.c);

  // Offer row/column insert & delete based on the rows/columns the selection spans.
  // A full-width/height header selection, or any multi-row/column cell selection,
  // targets those rows/columns (so a single cell stays uncluttered).
  const rowsN = range.r2 - range.r1 + 1;
  const colsN = range.c2 - range.c1 + 1;
  const rowsSelected = range.c1 === 0 && range.c2 === nCols - 1; // full-width row band
  const colsSelected = range.r1 === 0 && range.r2 === nRows - 1; // full-height column band
  const rowTarget = rowsSelected || rowsN > 1;
  const colTarget = colsSelected || colsN > 1;
  const canDeleteRows = rowTarget && rowsN < nRows;
  const canDeleteCols = colTarget && colsN < nCols;
  const canInsertRows = rowTarget && nRows + rowsN <= MAX_ROWS;
  const canInsertCols = colTarget && nCols + colsN <= MAX_COLS;

  function colHeaderStyle(c: number): CSSProperties {
    const on = c >= range.c1 && c <= range.c2;
    const frozen = c < frozenCols;
    return {
      position: frozen ? 'sticky' : 'relative',
      left: frozen ? colX[c] : undefined,
      zIndex: frozen ? 5 : undefined,
      width: cw(c),
      flex: `0 0 ${cw(c)}px`,
      height: HDRH,
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      font: "500 11px/1 var(--ex-font-mono)",
      letterSpacing: '0.1em',
      color: on ? accent : 'var(--ex-muted-2)',
      background: on ? mix(accent, 9, 'var(--ex-surface-2)') : 'var(--ex-surface-2)',
      borderRight: frozenCols > 0 && c === frozenCols - 1 ? '2px solid var(--ex-border-strong)' : '1px solid var(--ex-border-strong)',
      borderBottom: `1px solid ${on ? accent : 'var(--ex-border-strong)'}`,
      cursor: 'pointer',
    };
  }
  function rowHeaderStyle(r: number): CSSProperties {
    const on = r >= range.r1 && r <= range.r2;
    return {
      width: GUTW,
      flex: `0 0 ${GUTW}px`,
      height: ROWH,
      display: 'flex',
      alignItems: 'center',
      justifyContent: 'center',
      position: 'sticky',
      left: 0,
      // Above the selection ring (z2) so scrolling doesn't paint it over the gutter.
      zIndex: 4,
      font: "500 11px/1 var(--ex-font-mono)",
      color: on ? accent : 'var(--ex-muted-2)',
      background: on ? mix(accent, 9, 'var(--ex-surface-2)') : 'var(--ex-surface-2)',
      borderRight: `1px solid ${on ? accent : 'var(--ex-border-strong)'}`,
      borderBottom: frozenRows > 0 && r === frozenRows - 1 ? '2px solid var(--ex-border-strong)' : '1px solid var(--ex-gridline)',
      cursor: 'pointer',
    };
  }
  function cellView(r: number, c: number): { text: string; style: CSSProperties } {
    const cell = engine.cellAt(r, c) || {};
    const rw = cell.v || '';
    const val = engine.safeValue(r, c);
    const isErr = typeof val === 'string' && val[0] === '#' && val.length > 2;
    const isNum = typeof val === 'number' || typeof val === 'boolean';
    const inSel = r >= range.r1 && r <= range.r2 && c >= range.c1 && c <= range.c2;
    const isAnchor = r === sel.r && c === sel.c;
    const align = cell.a || (isNum ? 'right' : 'left');
    // Base background: the cell's fill if set, else paper. Fills are stored as light
    // absolute colors, so in dark mode blend them toward the dark surface so they read
    // as muted tints rather than glowing chips. Selection tints fill the whole cell.
    const base = cell.g
      ? dark
        ? `color-mix(in oklab, ${cell.g} 15%, var(--ex-surface))`
        : cell.g
      : 'var(--ex-surface)';
    let bg = base;
    if (inSel) bg = isAnchor ? (single ? base : mix(accent, 5, base)) : mix(accent, 11, base);

    // Collapsed borders: each interior edge is drawn once, by the top/left cell as its
    // bottom/right. Custom color wins over the gridline. Top/left are only drawn on the
    // grid's outer frame (row 0 / col 0) where there is no neighbor to own the line.
    const below = engine.cellAt(r + 1, c);
    const rightN = engine.cellAt(r, c + 1);
    const bottomColor = cell.bd?.b || below?.bd?.t;
    const rightColor = cell.bd?.r || rightN?.bd?.l;

    const style: CSSProperties = {
      width: cw(c),
      flex: `0 0 ${cw(c)}px`,
      height: ROWH,
      display: 'flex',
      alignItems: 'center',
      justifyContent: align === 'right' ? 'flex-end' : align === 'center' ? 'center' : 'flex-start',
      padding: '0 9px',
      overflow: 'hidden',
      whiteSpace: 'nowrap',
      cursor: 'cell',
      background: bg,
      borderTop: r === 0 && cell.bd?.t ? `1px solid ${cell.bd.t}` : '0',
      borderLeft: c === 0 && cell.bd?.l ? `1px solid ${cell.bd.l}` : '0',
      borderRight: rightColor ? `1px solid ${rightColor}` : gridlinesOn ? `1px solid ${LINE}` : '0',
      borderBottom: bottomColor ? `1px solid ${bottomColor}` : gridlinesOn ? `1px solid ${LINE}` : '0',
      font: `${cell.i ? 'italic ' : ''}${cell.b ? '600' : '400'} 12.5px/1 ${cell.ff || DEFAULT_FONT}`,
      letterSpacing: '-0.01em',
      textDecoration: cell.s ? 'line-through' : undefined,
      // Custom text colors are stored absolute; in dark mode lift the darker inks toward
      // the light text so they stay legible on a dark background.
      color: isErr
        ? 'var(--ex-danger)'
        : cell.co
        ? dark && cell.co[0] === '#'
          ? `color-mix(in oklab, ${cell.co} 68%, var(--ex-text))`
          : cell.co
        : rw[0] === '='
        ? 'var(--ex-text-strong)'
        : 'var(--ex-text)',
    };
    // Frozen columns: pin the first `frozenCols` cells of every row to the left
    // (frozen rows are handled by making the whole row div sticky — see the markup).
    const frozenCell = r < frozenRows || c < frozenCols;
    if (c < frozenCols) {
      style.position = 'sticky';
      style.left = colX[c];
      style.zIndex = 3;
    }
    if (frozenRows > 0 && r === frozenRows - 1) style.borderBottom = '2px solid var(--ex-border-strong)';
    if (frozenCols > 0 && c === frozenCols - 1) style.borderRight = '2px solid var(--ex-border-strong)';
    // The selection ring is an overlay that frozen (sticky) cells sit above, so draw
    // the ring per-cell on frozen selected cells via inset shadows on boundary edges.
    if (frozenCell && inSel) {
      const seg: string[] = [];
      if (r === range.r1) seg.push(`inset 0 1.6px 0 ${accent}`);
      if (r === range.r2) seg.push(`inset 0 -1.6px 0 ${accent}`);
      if (c === range.c1) seg.push(`inset 1.6px 0 0 ${accent}`);
      if (c === range.c2) seg.push(`inset -1.6px 0 0 ${accent}`);
      if (seg.length) style.boxShadow = seg.join(', ');
    }
    // The selection ring is drawn as an overlay above the cells (see the grid markup)
    // so it covers the gridlines/borders instead of sitting inside them.
    return { text: format(val, cell.f), style };
  }

  /** A merged region, drawn as one spanning overlay above the covered cells. */
  function mergeView(m: MergeRegion): { text: string; style: CSSProperties } {
    const cell = engine.cellAt(m.r1, m.c1) || {};
    const rw = cell.v || '';
    const val = engine.safeValue(m.r1, m.c1);
    const isErr = typeof val === 'string' && val[0] === '#' && val.length > 2;
    const isNum = typeof val === 'number' || typeof val === 'boolean';
    const align = cell.a || (isNum ? 'right' : 'left');
    const fill = cell.g ? (dark ? `color-mix(in oklab, ${cell.g} 15%, var(--ex-surface))` : cell.g) : 'var(--ex-surface)';
    const overlaps = m.r1 <= range.r2 && m.r2 >= range.r1 && m.c1 <= range.c2 && m.c2 >= range.c1;
    const anchorInside = sel.r >= m.r1 && sel.r <= m.r2 && sel.c >= m.c1 && sel.c <= m.c2;
    const bg = overlaps && !anchorInside ? mix(accent, 11, fill) : fill;
    const gl = gridlinesOn ? `1px solid ${LINE}` : '0';
    const style: CSSProperties = {
      position: 'absolute',
      left: colX[m.c1],
      top: HDRH + m.r1 * ROWH,
      width: colX[m.c2 + 1] - colX[m.c1],
      height: (m.r2 - m.r1 + 1) * ROWH,
      display: 'flex',
      alignItems: 'center',
      justifyContent: align === 'right' ? 'flex-end' : align === 'center' ? 'center' : 'flex-start',
      padding: '0 9px',
      overflow: 'hidden',
      whiteSpace: 'nowrap',
      boxSizing: 'border-box',
      background: bg,
      borderTop: cell.bd?.t ? `1px solid ${cell.bd.t}` : '0',
      borderLeft: cell.bd?.l ? `1px solid ${cell.bd.l}` : '0',
      borderRight: cell.bd?.r ? `1px solid ${cell.bd.r}` : gl,
      borderBottom: cell.bd?.b ? `1px solid ${cell.bd.b}` : gl,
      font: `${cell.i ? 'italic ' : ''}${cell.b ? '600' : '400'} 12.5px/1 ${cell.ff || DEFAULT_FONT}`,
      letterSpacing: '-0.01em',
      textDecoration: cell.s ? 'line-through' : undefined,
      color: isErr
        ? 'var(--ex-danger)'
        : cell.co
        ? dark && cell.co[0] === '#'
          ? `color-mix(in oklab, ${cell.co} 68%, var(--ex-text))`
          : cell.co
        : rw[0] === '='
        ? 'var(--ex-text-strong)'
        : 'var(--ex-text)',
      zIndex: 1,
      pointerEvents: 'none',
    };
    return { text: format(val, cell.f), style };
  }

  // When the active cell is merged, the editor spans the whole region.
  const editMerge = engine.mergeAt(sel.r, sel.c);
  const editorStyle: CSSProperties = {
    position: 'absolute',
    left: colX[sel.c],
    top: HDRH + sel.r * ROWH,
    width: editMerge ? colX[editMerge.c2 + 1] - colX[sel.c] : Math.max(cw(sel.c), 168),
    height: editMerge ? (editMerge.r2 - editMerge.r1 + 1) * ROWH : ROWH,
    zIndex: 9,
    padding: '0 8px',
    border: `1.6px solid ${accent}`,
    borderRadius: 2,
    background: 'var(--ex-popover)',
    outline: 'none',
    font: "400 12.5px/1 var(--ex-font-mono)",
    color: 'var(--ex-text)',
    boxShadow: '0 6px 18px rgba(26,26,24,0.10)',
  };

  /* ---------- markup ---------- */

  return (
    <div
      className={styles.root}
      style={{
        height: '100vh',
        display: 'flex',
        flexDirection: 'column',
        background: 'var(--ex-bg)',
        color: 'var(--ex-text)',
        fontFamily: "var(--ex-font-sans)",
        WebkitFontSmoothing: 'antialiased',
        overflow: 'hidden',
      }}
    >
      {/* Title bar */}
      <div style={{ display: 'flex', alignItems: 'baseline', gap: 14, padding: '18px 22px 14px', borderBottom: '1px solid var(--ex-border)' }}>
        <input
          value={workbook.getName()}
          onChange={(e) => workbook.setName(e.target.value)}
          placeholder="Untitled workbook"
          style={{
            flex: 1,
            minWidth: 0,
            border: 0,
            background: 'transparent',
            outline: 'none',
            font: "400 19px/1.2 var(--ex-font-sans)",
            letterSpacing: '-0.01em',
            color: 'var(--ex-text)',
            padding: '2px 0',
          }}
        />
        {savedLabel && (
          <span
            title={savedLabel}
            aria-label={savedLabel}
            style={{ display: 'inline-flex', alignItems: 'center', color: accent, flex: '0 0 auto' }}
          >
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
              <circle cx="8" cy="8" r="6.4" stroke="currentColor" strokeWidth="1.1" opacity="0.55" />
              <path d="M5.2 8.2l1.9 1.9 3.7-4.1" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round" />
            </svg>
          </span>
        )}
      </div>

      {/* Toolbar */}
      <div style={{ display: 'flex', alignItems: 'center', gap: 3, padding: '8px 18px', borderBottom: '1px solid var(--ex-border)', flexWrap: 'wrap' }}>
        {/* Font family */}
        <FontControl value={active.ff || DEFAULT_FONT} accent={accent} fonts={FONTS} systemFonts={systemFonts} onPick={applyFont} />
        <Divider />

        {/* Text formatting */}
        <button
          {...tip('Bold — ⌘B')}
          className={styles.tbtn}
          onClick={() => toggle('b')}
          style={{ ...iconBtnStyle(!!active.b, accent), font: "700 14px/1 var(--ex-font-sans)" }}
        >
          B
        </button>
        <button
          {...tip('Italic — ⌘I')}
          className={styles.tbtn}
          onClick={() => toggle('i')}
          style={{ ...iconBtnStyle(!!active.i, accent), font: "italic 600 14px/1 var(--ex-font-serif)" }}
        >
          I
        </button>
        <button
          {...tip('Strikethrough')}
          className={styles.tbtn}
          onClick={() => toggle('s')}
          style={{ ...iconBtnStyle(!!active.s, accent), font: "600 14px/1 var(--ex-font-sans)", textDecoration: 'line-through' }}
        >
          S
        </button>
        <ColorControl
          kind="text"
          tooltip="Text color"
          current={active.co}
          presets={TEXT_PRESETS}
          accent={accent}
          onPick={applyText}
          onClear={() => applyText('')}
          clearLabel="Default color"
        />

        <Divider />

        {/* Fill · borders · gridlines */}
        <ColorControl
          kind="fill"
          tooltip="Fill color"
          current={active.g}
          presets={FILL_PRESETS}
          accent={accent}
          onPick={applyFill}
          onClear={() => applyFill('')}
          clearLabel="No fill"
        />
        <BordersControl accent={accent} onApply={applyBorders} />
        <button
          {...tip(gridlinesOn ? 'Hide gridlines' : 'Show gridlines')}
          className={styles.tbtn}
          onClick={toggleGridlines}
          style={iconBtnStyle(gridlinesOn, accent)}
        >
          <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
            <rect x="2" y="2" width="12" height="12" stroke="currentColor" strokeWidth="1.1" />
            <path d="M6 2v12M10 2v12M2 6h12M2 10h12" stroke="currentColor" strokeWidth="0.9" />
          </svg>
        </button>

        <Divider />

        {/* Alignment · merge · number format */}
        <AlignControl value={activeAlign} accent={accent} onPick={applyAlign} />
        <button
          {...tip(mergedHere ? 'Unmerge cells' : 'Merge cells')}
          className={styles.tbtn}
          onClick={toggleMerge}
          disabled={!canMerge}
          style={{
            ...iconBtnStyle(false, accent),
            background: mergedHere ? mix(accent, 13, 'var(--ex-surface)') : 'transparent',
            opacity: canMerge ? 1 : 0.4,
            cursor: canMerge ? 'pointer' : 'default',
          }}
        >
          <svg width="15" height="15" viewBox="0 0 1024 1024" fill="currentColor">
            <path d="M482.2 508.4L331.3 389c-3-2.4-7.3-0.2-7.3 3.6V478H184V184h204v128c0 2.2 1.8 4 4 4h60c2.2 0 4-1.8 4-4V144c0-15.5-12.5-28-28-28H144c-15.5 0-28 12.5-28 28v736c0 15.5 12.5 28 28 28h284c15.5 0 28-12.5 28-28V712c0-2.2-1.8-4-4-4h-60c-2.2 0-4 1.8-4 4v128H184V546h140v85.4c0 3.8 4.4 6 7.3 3.6l150.9-119.4c2.4-1.8 2.4-5.4 0-7.2zM880 116H596c-15.5 0-28 12.5-28 28v168c0 2.2 1.8 4 4 4h60c2.2 0 4-1.8 4-4V184h204v294H700v-85.4c0-3.8-4.3-6-7.3-3.6l-151 119.4c-2.3 1.8-2.3 5.3 0 7.1l151 119.5c2.9 2.3 7.3 0.2 7.3-3.6V546h140v294H636V712c0-2.2-1.8-4-4-4h-60c-2.2 0-4 1.8-4 4v168c0 15.5 12.5 28 28 28h284c15.5 0 28-12.5 28-28V144c0-15.5-12.5-28-28-28z" />
          </svg>
        </button>
        <FreezeControl accent={accent} frozenRows={frozenRows} frozenCols={frozenCols} selR2={range.r2} selC2={range.c2} onFreeze={applyFreeze} />
        <MenuSelect
          value={active.f || 'auto'}
          options={FORMAT_OPTIONS}
          accent={accent}
          tooltip="Number format"
          width={116}
          onPick={(v) => applyFormat(v as NumberFormat | 'auto')}
        />

        <Divider />

        {/* History */}
        <button {...tip('Undo — ⌘Z')} className={styles.tbtn} onClick={() => engine.undo()} style={iconBtnStyle(false, accent)}>
          <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
            <path d="M5.6 3.4 2.6 6.4l3 3" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round" />
            <path d="M3 6.4h5.5a3.3 3.3 0 0 1 0 6.6H6" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        </button>
        <button {...tip('Redo — ⇧⌘Z')} className={styles.tbtn} onClick={() => engine.redo()} style={iconBtnStyle(false, accent)}>
          <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
            <path d="M10.4 3.4 13.4 6.4l-3 3" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round" />
            <path d="M13 6.4H7.5a3.3 3.3 0 0 0 0 6.6H10" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        </button>

        <div style={{ flex: 1 }} />
        {canInsertRows && (
          <button {...tip('Insert rows')} className={styles.tbtn} onClick={() => setInsertAxis('rows')} style={iconBtnStyle(false, accent)}>
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
              <rect x="2.5" y="2.5" width="11" height="3.6" rx="1" stroke="currentColor" strokeWidth="1.1" />
              <path d="M8 9v4.5M5.75 11.25h4.5" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" />
            </svg>
          </button>
        )}
        {canDeleteRows && (
          <button
            {...tip(`Delete ${rowsN} row${rowsN > 1 ? 's' : ''}`)}
            className={styles.danger}
            onClick={deleteSelectedRows}
            style={{ ...iconBtnStyle(false, accent), color: 'var(--ex-remove)' }}
          >
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
              <rect x="2.5" y="2.5" width="11" height="3.6" rx="1" stroke="currentColor" strokeWidth="1.1" />
              <path d="M5.75 11.25h4.5" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" />
            </svg>
          </button>
        )}
        {canInsertCols && (
          <button {...tip('Insert columns')} className={styles.tbtn} onClick={() => setInsertAxis('cols')} style={iconBtnStyle(false, accent)}>
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
              <rect x="2.5" y="2.5" width="3.6" height="11" rx="1" stroke="currentColor" strokeWidth="1.1" />
              <path d="M11.25 5.75v4.5M9 8h4.5" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" />
            </svg>
          </button>
        )}
        {canDeleteCols && (
          <button
            {...tip(`Delete ${colsN} column${colsN > 1 ? 's' : ''}`)}
            className={styles.danger}
            onClick={deleteSelectedCols}
            style={{ ...iconBtnStyle(false, accent), color: 'var(--ex-remove)' }}
          >
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
              <rect x="2.5" y="2.5" width="3.6" height="11" rx="1" stroke="currentColor" strokeWidth="1.1" />
              <path d="M9 8h4.5" stroke="currentColor" strokeWidth="1.3" strokeLinecap="round" />
            </svg>
          </button>
        )}
        {(canInsertRows || canInsertCols || canDeleteRows || canDeleteCols) && <Divider />}
        <button
          {...tip(theme === 'dark' ? 'Light theme' : 'Dark theme')}
          className={styles.tbtn}
          onClick={toggleTheme}
          style={iconBtnStyle(false, accent)}
        >
          {theme === 'dark' ? (
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
              <circle cx="8" cy="8" r="3.2" stroke="currentColor" strokeWidth="1.2" />
              <path
                d="M8 1.5v1.6M8 12.9v1.6M1.5 8h1.6M12.9 8h1.6M3.4 3.4l1.1 1.1M11.5 11.5l1.1 1.1M12.6 3.4l-1.1 1.1M4.5 11.5l-1.1 1.1"
                stroke="currentColor"
                strokeWidth="1.2"
                strokeLinecap="round"
              />
            </svg>
          ) : (
            <svg width="16" height="16" viewBox="0 0 16 16" fill="none">
              <path d="M13.5 9.3A5.6 5.6 0 0 1 6.7 2.5a5.6 5.6 0 1 0 6.8 6.8z" stroke="currentColor" strokeWidth="1.2" strokeLinejoin="round" />
            </svg>
          )}
        </button>
        <Divider />
        <button {...tip('Clear selected cells — ⌫')} className={styles.danger} onClick={clearSelection} style={dangerGhost}>
          Clear
        </button>
        <button {...tip('Delete all data and clear storage')} className={styles.danger} onClick={() => setConfirmReset(true)} style={dangerButton}>
          Reset all
        </button>
      </div>

      {/* Formula bar */}
      <div style={{ display: 'flex', alignItems: 'stretch', borderBottom: '1px solid var(--ex-border)', background: 'var(--ex-surface)' }}>
        <div
          style={{
            width: 92,
            flex: '0 0 92px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            font: "500 13px/1 var(--ex-font-mono)",
            color: accent,
            borderRight: '1px solid var(--ex-border)',
          }}
        >
          {colName(sel.c) + (sel.r + 1)}
        </div>
        <div
          style={{
            width: 34,
            flex: '0 0 34px',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
            font: "italic 500 14px/1 var(--ex-font-sans)",
            color: 'var(--ex-text-faintest)',
            borderRight: '1px solid var(--ex-border)',
          }}
        >
          fx
        </div>
        <input
          ref={barInputRef}
          value={barValue}
          onChange={(e) => {
            setEditing(true);
            setEditSource('bar');
            setEditText(e.target.value);
            setPointRange(null);
          }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              commit([1, 0]);
              (e.target as HTMLInputElement).blur();
            } else if (e.key === 'Escape') {
              e.preventDefault();
              cancelEdit();
              (e.target as HTMLInputElement).blur();
            }
          }}
          onBlur={() => {
            if (view.current.editing && view.current.editSource === 'bar') commit();
          }}
          placeholder="Value or =SUM(A1:A9)"
          style={{
            flex: 1,
            minWidth: 0,
            border: 0,
            outline: 'none',
            background: 'transparent',
            padding: '11px 14px',
            font: "400 13px/1 var(--ex-font-mono)",
            letterSpacing: '-0.01em',
            color: 'var(--ex-text)',
          }}
        />
      </div>

      {/* Grid */}
      <div
        ref={scrollRef}
        style={{
          flex: 1,
          overflow: 'scroll',
          background: 'var(--ex-bg)',
          // Column layout + grid flex-grow below keep the bottom bars pinned to the
          // container's bottom edge even when the grid is shorter than the viewport.
          display: 'flex',
          flexDirection: 'column',
        }}
      >
        <div
          onMouseDown={onGridMouseDown}
          onDoubleClick={onGridDoubleClick}
          style={{ position: 'relative', flex: '1 0 auto', width: 'max-content', minWidth: '100%', userSelect: 'none' }}
        >
          {/* Header row */}
          <div style={{ display: 'flex', position: 'sticky', top: 0, zIndex: 6 }}>
            <div
              style={{
                width: GUTW,
                flex: `0 0 ${GUTW}px`,
                height: HDRH,
                background: 'var(--ex-surface-2)',
                borderRight: '1px solid var(--ex-border-strong)',
                borderBottom: '1px solid var(--ex-border-strong)',
                position: 'sticky',
                left: 0,
                zIndex: 4,
              }}
            />
            {cols.map((c) => (
              <div key={c} data-col={c} style={colHeaderStyle(c)}>
                {colName(c)}
                {/* Drag the right edge to resize this column */}
                <div
                  {...tip('Drag to resize column')}
                  onMouseDown={(e) => startColResize(e, c)}
                  onDoubleClick={(e) => e.stopPropagation()}
                  style={{
                    position: 'absolute',
                    top: 0,
                    right: -3,
                    width: 7,
                    height: '100%',
                    cursor: 'col-resize',
                    zIndex: 5,
                  }}
                />
              </div>
            ))}
          </div>

          {/* Body rows (frozen rows are made sticky at the row level) */}
          {rows.map((r) => (
            <div
              key={r}
              style={
                r < frozenRows
                  ? { display: 'flex', position: 'sticky', top: HDRH + r * ROWH, zIndex: 5 }
                  : { display: 'flex' }
              }
            >
              <div data-row={r} style={rowHeaderStyle(r)}>
                {r + 1}
              </div>
              {cols.map((c) => {
                const v = cellView(r, c);
                return (
                  <div key={c} data-cell={`${r},${c}`} style={v.style}>
                    {v.text}
                  </div>
                );
              })}
            </div>
          ))}

          {/* Merged regions — one spanning overlay above the covered cells. */}
          {engine.getMerges().map((m) => {
            const v = mergeView(m);
            return (
              <div key={`m${m.r1}.${m.c1}`} style={v.style}>
                {v.text}
              </div>
            );
          })}

          {/* Selection ring — drawn above the cells so it covers the gridlines/borders.
              Hidden while editing (the editor overlay stands in for it). */}
          {!editing && (
            <div
              style={{
                position: 'absolute',
                // Expand 1px on every side so the 2px ring straddles and hides the
                // cell's own borders (gridlines and custom borders) on all four edges,
                // including the top/left ones owned by the neighbouring cells.
                left: colX[range.c1] - 1,
                top: HDRH + range.r1 * ROWH - 1,
                width: colX[range.c2 + 1] - colX[range.c1] + 2,
                height: (range.r2 - range.r1 + 1) * ROWH + 2,
                border: `2px solid ${accent}`,
                boxSizing: 'border-box',
                pointerEvents: 'none',
                zIndex: 2,
              }}
            />
          )}

          {/* Point-mode highlight: the range being inserted into the formula */}
          {editing && pointRange && (
            <div
              style={{
                position: 'absolute',
                left: colX[pointRange.c1],
                top: HDRH + pointRange.r1 * ROWH,
                width: colX[pointRange.c2 + 1] - colX[pointRange.c1],
                height: (pointRange.r2 - pointRange.r1 + 1) * ROWH,
                border: `1.5px dashed ${accent}`,
                background: mix(accent, 8, 'transparent'),
                pointerEvents: 'none',
                zIndex: 4,
              }}
            />
          )}

          {/* Inline editor overlay (grid edits only) */}
          {editing && editSource === 'grid' && (
            <input
              ref={(el) => {
                editInputRef.current = el;
                if (el && document.activeElement !== el) {
                  el.focus();
                  el.setSelectionRange(el.value.length, el.value.length);
                }
              }}
              value={editText}
              onChange={(e) => {
                setEditText(e.target.value);
                setPointRange(null);
              }}
              onKeyDown={(e) => {
                if (e.key === 'Enter') {
                  e.preventDefault();
                  commit([1, 0]);
                } else if (e.key === 'Tab') {
                  e.preventDefault();
                  commit([0, e.shiftKey ? -1 : 1]);
                } else if (e.key === 'Escape') {
                  e.preventDefault();
                  cancelEdit();
                }
              }}
              onBlur={() => {
                if (view.current.editing) commit();
              }}
              style={editorStyle}
            />
          )}
        </div>

        {/* Sticky bottom bars — tabs + footer stay put during horizontal scroll, so the
            grid's own native horizontal scrollbar sits below them at the page bottom. */}
        <div style={{ position: 'sticky', left: 0, bottom: 0, width: viewportW || '100%', zIndex: 7 }}>
          {/* Sheet tabs */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 2, padding: '0 10px', borderTop: '1px solid var(--ex-border)', background: 'var(--ex-surface-2)', minHeight: 34 }}>
        {workbook.list().map((s) => {
          const active = s.id === workbook.getActiveId();
          if (renaming?.id === s.id) {
            return (
              <input
                key={s.id}
                autoFocus
                value={renaming.text}
                onChange={(e) => setRenaming({ id: s.id, text: e.target.value })}
                onFocus={(e) => e.target.select()}
                onKeyDown={(e) => {
                  if (e.key === 'Enter') {
                    e.preventDefault();
                    commitRename();
                  } else if (e.key === 'Escape') {
                    e.preventDefault();
                    setRenaming(null);
                  }
                }}
                onBlur={commitRename}
                style={{
                  width: 120,
                  height: 26,
                  margin: '4px 0',
                  padding: '0 8px',
                  border: `1px solid ${accent}`,
                  borderRadius: 5,
                  background: 'var(--ex-surface)',
                  font: "600 12.5px/1 var(--ex-font-sans)",
                  color: 'var(--ex-text)',
                  outline: 'none',
                }}
              />
            );
          }
          return (
            <div
              key={s.id}
              className={`${styles.tab}${active ? ' ' + styles.active : ''}`}
              onClick={() => switchSheet(s.id)}
              onDoubleClick={() => setRenaming({ id: s.id, text: s.name })}
              {...tip('Double-click to rename')}
              style={{
                display: 'flex',
                alignItems: 'center',
                gap: 6,
                height: 34,
                padding: '0 12px',
                borderTop: active ? `2px solid ${accent}` : '2px solid transparent',
                background: active ? 'var(--ex-surface)' : 'transparent',
                color: active ? 'var(--ex-text)' : 'var(--ex-text-faint)',
                font: `${active ? '600' : '500'} 12.5px/1 var(--ex-font-sans)`,
                cursor: 'pointer',
                userSelect: 'none',
              }}
            >
              {s.name}
              {/* Reserved on every tab (when deletable) so switching doesn't shift the layout. */}
              {workbook.count() > 1 && (
                <span
                  className={styles.x}
                  onClick={(e) => {
                    e.stopPropagation();
                    setConfirmDeleteSheet(s.id);
                  }}
                  {...tip('Delete sheet')}
                  style={{ display: 'inline-flex', alignItems: 'center', color: 'var(--ex-text-fainter)' }}
                >
                  <CloseIcon size={11} />
                </span>
              )}
            </div>
          );
        })}
        {workbook.count() < MAX_SHEETS && (
          <button
            {...tip('Add sheet')}
            onClick={addSheet}
            className={styles.tbtn}
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              justifyContent: 'center',
              width: 28,
              height: 28,
              margin: '0 4px',
              border: 0,
              background: 'transparent',
              borderRadius: 5,
              color: 'var(--ex-text-muted)',
              cursor: 'pointer',
              outline: 'none',
            }}
          >
            <PlusIcon size={13} />
          </button>
        )}
        </div>

        {/* Status bar */}
      <div
        style={{
          display: 'flex',
          alignItems: 'center',
          padding: '8px 22px',
          borderTop: '1px solid var(--ex-border)',
          background: 'var(--ex-surface)',
          font: "400 12px/1 var(--ex-font-mono)",
          color: 'var(--ex-text-faint)',
        }}
      >
        <button
          onClick={() => setHelpOpen(true)}
          title="Formula reference"
          style={{
            border: 0,
            background: 'transparent',
            padding: 0,
            color: 'var(--ex-text-faint)',
            font: "500 12px/1 var(--ex-font-mono)",
            cursor: 'pointer',
            outline: 'none',
          }}
        >
          Functions
        </button>
        <div style={{ flex: 1 }} />
        <div style={{ display: 'flex', alignItems: 'center', gap: 10 }}>
          <span>{selLabel}</span>
          <Dot />
          <span>Sum {format(stats.count ? stats.sum : '', 'k')}</span>
          <Dot />
          <span>Avg {stats.count ? format(Math.round((stats.sum / stats.count) * 1e4) / 1e4, 'k') : ''}</span>
          <Dot />
          <span>Filled {stats.filled}</span>
        </div>
      </div>
        </div>
      </div>

      {helpOpen && <HelpModal accent={accent} onClose={() => setHelpOpen(false)} />}
      {confirmReset && (
        <ConfirmModal
          accent={accent}
          title="Reset everything?"
          body="This permanently deletes all cells, formatting, and stored data. This can't be undone."
          confirmLabel="Delete everything"
          onCancel={() => setConfirmReset(false)}
          onConfirm={doReset}
        />
      )}
      {confirmDeleteSheet && (
        <ConfirmModal
          accent={accent}
          title="Delete sheet?"
          body={`"${workbook.list().find((s) => s.id === confirmDeleteSheet)?.name ?? 'This sheet'}" and all its data will be permanently removed.`}
          confirmLabel="Delete sheet"
          onCancel={() => setConfirmDeleteSheet(null)}
          onConfirm={() => removeSheet(confirmDeleteSheet)}
        />
      )}
      {insertAxis && (
        <InsertModal
          accent={accent}
          axis={insertAxis}
          defaultCount={insertAxis === 'rows' ? rowsN : colsN}
          max={insertAxis === 'rows' ? MAX_ROWS - nRows : MAX_COLS - nCols}
          onCancel={() => setInsertAxis(null)}
          onConfirm={(count, position) => doInsert(insertAxis, count, position)}
        />
      )}

      <Tooltip
        id={TOOLTIP_ID}
        className="excel-tip"
        place="bottom"
        delayShow={300}
        offset={6}
        style={{
          background: 'var(--ex-tooltip-bg)',
          color: 'var(--ex-tooltip-text)',
          fontSize: 11,
          fontFamily: "var(--ex-font-sans)",
          padding: '3px 7px',
          borderRadius: 5,
          zIndex: 200,
        }}
      />
    </div>
  );
}

/* -------------------------------------------------------------------------- */

/** Modal listing every supported formula function, grouped by kind. */
function HelpModal({ accent, onClose }: { accent: string; onClose: () => void }) {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onClose();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onClose]);

  return (
    <div
      onMouseDown={onClose}
      style={{
        position: 'fixed',
        inset: 0,
        zIndex: 100,
        background: 'rgba(26,26,24,0.32)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        padding: 24,
      }}
    >
      <div
        onMouseDown={(e) => e.stopPropagation()}
        style={{
          width: 'min(720px, 100%)',
          maxHeight: '82vh',
          overflow: 'auto',
          background: 'var(--ex-bg)',
          border: '1px solid var(--ex-border-2)',
          borderRadius: 12,
          boxShadow: '0 24px 64px rgba(26,26,24,0.28)',
        }}
      >
        <div
          style={{
            position: 'sticky',
            top: 0,
            display: 'flex',
            alignItems: 'baseline',
            gap: 12,
            padding: '18px 22px 14px',
            background: 'var(--ex-bg)',
            borderBottom: '1px solid var(--ex-border)',
          }}
        >
          <span style={{ font: "600 16px/1 var(--ex-font-sans)", color: 'var(--ex-text)' }}>Formula reference</span>
          <span style={{ font: "400 12px/1 var(--ex-font-mono)", color: 'var(--ex-text-fainter)' }}>
            start a cell with = · e.g. =SUM(A1:A9)
          </span>
          <div style={{ flex: 1 }} />
          <button
            onClick={onClose}
            {...tip('Close — Esc')}
            style={{
              display: 'inline-flex',
              alignItems: 'center',
              justifyContent: 'center',
              border: '1px solid var(--ex-border-2)',
              background: 'var(--ex-surface)',
              borderRadius: 6,
              width: 28,
              height: 28,
              cursor: 'pointer',
              color: 'var(--ex-text-muted)',
              outline: 'none',
            }}
          >
            <CloseIcon size={13} />
          </button>
        </div>

        <div style={{ padding: '6px 22px 22px' }}>
          {FUNCTION_HELP.map((group) => (
            <div key={group.group} style={{ marginTop: 18 }}>
              <div
                style={{
                  font: "500 11px/1 var(--ex-font-mono)",
                  letterSpacing: '0.14em',
                  textTransform: 'uppercase',
                  color: accent,
                  marginBottom: 8,
                }}
              >
                {group.group}
              </div>
              <div style={{ display: 'grid', gap: 1, background: 'var(--ex-border)', border: '1px solid var(--ex-border)', borderRadius: 8, overflow: 'hidden' }}>
                {group.items.map((fn) => (
                  <div key={fn.sig} style={{ display: 'flex', gap: 14, padding: '9px 12px', background: 'var(--ex-surface)' }}>
                    <code style={{ flex: '0 0 210px', font: "500 12.5px/1.35 var(--ex-font-mono)", color: 'var(--ex-text-strong)' }}>{fn.sig}</code>
                    <span style={{ font: "400 12.5px/1.4 var(--ex-font-sans)", color: 'var(--ex-text-muted)' }}>{fn.desc}</span>
                  </div>
                ))}
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

/** Small centered confirmation modal (replaces window.confirm). */
function ConfirmModal({
  accent,
  title,
  body,
  confirmLabel,
  onConfirm,
  onCancel,
}: {
  accent: string;
  title: string;
  body: string;
  confirmLabel: string;
  onConfirm: () => void;
  onCancel: () => void;
}) {
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onCancel();
      if (e.key === 'Enter') onConfirm();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onCancel, onConfirm]);

  return (
    <div
      onMouseDown={onCancel}
      style={{
        position: 'fixed',
        inset: 0,
        zIndex: 120,
        background: 'rgba(26,26,24,0.32)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        padding: 24,
      }}
    >
      <div
        onMouseDown={(e) => e.stopPropagation()}
        style={{
          width: 'min(400px, 100%)',
          background: 'var(--ex-bg)',
          border: '1px solid var(--ex-border-2)',
          borderRadius: 12,
          boxShadow: '0 24px 64px rgba(26,26,24,0.28)',
          padding: 22,
        }}
      >
        <div style={{ font: "600 16px/1.3 var(--ex-font-sans)", color: 'var(--ex-text)' }}>{title}</div>
        <div style={{ marginTop: 8, font: "400 13.5px/1.5 var(--ex-font-sans)", color: 'var(--ex-text-muted)' }}>{body}</div>
        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 20 }}>
          <button
            onClick={onCancel}
            style={{
              padding: '8px 14px',
              borderRadius: 7,
              border: '1px solid var(--ex-border-2)',
              background: 'var(--ex-surface)',
              color: 'var(--ex-select-text)',
              font: "500 13px/1 var(--ex-font-sans)",
              cursor: 'pointer',
              outline: 'none',
            }}
          >
            Cancel
          </button>
          <button
            autoFocus
            onClick={onConfirm}
            style={{
              padding: '8px 14px',
              borderRadius: 7,
              border: '1px solid var(--ex-danger)',
              background: 'var(--ex-danger)',
              color: 'var(--ex-on-accent)',
              font: "600 13px/1 var(--ex-font-sans)",
              cursor: 'pointer',
              outline: 'none',
            }}
          >
            {confirmLabel}
          </button>
        </div>
      </div>
    </div>
  );
}

/** Insert rows/columns: choose how many and whether before or after the selection. */
function InsertModal({
  accent,
  axis,
  defaultCount,
  max,
  onConfirm,
  onCancel,
}: {
  accent: string;
  axis: 'rows' | 'cols';
  defaultCount: number;
  max: number;
  onConfirm: (count: number, position: 'before' | 'after') => void;
  onCancel: () => void;
}) {
  const [text, setText] = useState(String(Math.max(1, defaultCount)));
  const [position, setPosition] = useState<'before' | 'after'>('before');
  const n = parseInt(text, 10);
  const valid = !Number.isNaN(n) && n >= 1 && n <= max;
  const submit = () => {
    if (valid) onConfirm(n, position);
  };
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') onCancel();
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onCancel]);

  const noun = axis === 'rows' ? 'rows' : 'columns';
  const beforeLabel = axis === 'rows' ? 'Above' : 'Left';
  const afterLabel = axis === 'rows' ? 'Below' : 'Right';
  const seg = (active: boolean): CSSProperties => ({
    flex: 1,
    padding: '7px 0',
    borderRadius: 6,
    border: `1px solid ${active ? accent : 'var(--ex-border-2)'}`,
    background: active ? mix(accent, 12, 'var(--ex-surface)') : 'var(--ex-surface)',
    color: active ? accent : 'var(--ex-text-muted)',
    font: `${active ? '600' : '500'} 12.5px/1 var(--ex-font-sans)`,
    cursor: 'pointer',
    outline: 'none',
  });

  return (
    <div
      onMouseDown={onCancel}
      style={{ position: 'fixed', inset: 0, zIndex: 120, background: 'rgba(26,26,24,0.32)', display: 'flex', alignItems: 'center', justifyContent: 'center', padding: 24 }}
    >
      <div
        onMouseDown={(e) => e.stopPropagation()}
        style={{ width: 'min(380px, 100%)', background: 'var(--ex-bg)', border: '1px solid var(--ex-border-2)', borderRadius: 12, boxShadow: '0 24px 64px rgba(26,26,24,0.28)', padding: 22 }}
      >
        <div style={{ font: "600 16px/1.3 var(--ex-font-sans)", color: 'var(--ex-text)' }}>Insert {noun}</div>

        <label style={{ display: 'block', marginTop: 14, font: "500 12px/1 var(--ex-font-sans)", color: 'var(--ex-text-faint)' }}>How many {noun}?</label>
        <input
          autoFocus
          type="number"
          min={1}
          max={max}
          value={text}
          onChange={(e) => setText(e.target.value)}
          onFocus={(e) => e.target.select()}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              submit();
            }
          }}
          style={{
            marginTop: 6,
            width: '100%',
            padding: '9px 11px',
            borderRadius: 7,
            border: `1px solid ${valid ? 'var(--ex-border-2)' : 'var(--ex-danger-border)'}`,
            background: 'var(--ex-popover)',
            color: 'var(--ex-text)',
            font: "500 14px/1 var(--ex-font-mono)",
            outline: 'none',
          }}
        />

        <label style={{ display: 'block', marginTop: 16, font: "500 12px/1 var(--ex-font-sans)", color: 'var(--ex-text-faint)' }}>Position</label>
        <div style={{ display: 'flex', gap: 8, marginTop: 6 }}>
          <button onClick={() => setPosition('before')} style={seg(position === 'before')}>
            {beforeLabel} selection
          </button>
          <button onClick={() => setPosition('after')} style={seg(position === 'after')}>
            {afterLabel} selection
          </button>
        </div>

        <div style={{ display: 'flex', justifyContent: 'flex-end', gap: 8, marginTop: 20 }}>
          <button
            onClick={onCancel}
            style={{ padding: '8px 14px', borderRadius: 7, border: '1px solid var(--ex-border-2)', background: 'var(--ex-surface)', color: 'var(--ex-text-muted)', font: "500 13px/1 var(--ex-font-sans)", cursor: 'pointer', outline: 'none' }}
          >
            Cancel
          </button>
          <button
            onClick={submit}
            disabled={!valid}
            style={{
              padding: '8px 16px',
              borderRadius: 7,
              border: `1px solid ${accent}`,
              background: valid ? accent : mix(accent, 45, 'var(--ex-surface)'),
              color: 'var(--ex-on-accent)',
              font: "600 13px/1 var(--ex-font-sans)",
              cursor: valid ? 'pointer' : 'not-allowed',
              outline: 'none',
            }}
          >
            Insert
          </button>
        </div>
      </div>
    </div>
  );
}
