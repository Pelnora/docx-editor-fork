/**
 * Zoom Control Component (Radix UI)
 *
 * A dropdown for controlling document zoom level using Radix Select.
 */

import * as React from 'react';
import { Select, SelectContent, SelectItem, SelectTrigger, SelectValue } from './Select';
import { cn } from '../../lib/utils';
import { useTranslation } from '../../i18n';

// ============================================================================
// TYPES
// ============================================================================

export interface ZoomLevel {
  value: number;
  label: string;
}

/** Pelnora .9 — the host's "fit" entry at the top of the menu. */
export interface ZoomFitOption {
  /** Menu label, e.g. "Dopasuj". */
  label: string;
  /** Whether the host's fit is the current mode (the entry shows as chosen). */
  active: boolean;
  /** Choosing the entry: the host computes the zoom and sets it. */
  onFit: () => void;
}

export interface ZoomControlProps {
  value?: number;
  onChange?: (zoom: number) => void;
  /** Pelnora .9 — a "fit" entry above the levels. */
  fit?: ZoomFitOption;
  levels?: ZoomLevel[];
  disabled?: boolean;
  className?: string;
  minZoom?: number;
  maxZoom?: number;
  showButtons?: boolean;
  persistZoom?: boolean;
  storageKey?: string;
  compact?: boolean;
}

// ============================================================================
// CONSTANTS
// ============================================================================

const DEFAULT_ZOOM_LEVELS: ZoomLevel[] = [
  { value: 0.5, label: '50%' },
  { value: 0.75, label: '75%' },
  { value: 1.0, label: '100%' },
  { value: 1.25, label: '125%' },
  { value: 1.5, label: '150%' },
  { value: 2.0, label: '200%' },
];

// ============================================================================
// COMPONENT
// ============================================================================

const FIT_VALUE = 'fit';

export function ZoomControl({
  value = 1.0,
  onChange,
  fit,
  levels = DEFAULT_ZOOM_LEVELS,
  disabled = false,
  className,
  compact = false,
}: ZoomControlProps) {
  const { t } = useTranslation();
  const displayLabel = React.useMemo(() => {
    const matchingLevel = levels.find((level) => Math.abs(level.value - value) < 0.001);
    if (matchingLevel) return matchingLevel.label;
    return `${Math.round(value * 100)}%`;
  }, [levels, value]);

  const handleValueChange = React.useCallback(
    (newValue: string) => {
      if (newValue === FIT_VALUE) {
        fit?.onFit();
        return;
      }
      const zoom = parseFloat(newValue);
      if (!isNaN(zoom)) {
        onChange?.(zoom);
      }
    },
    [onChange, fit]
  );

  return (
    <Select
      value={fit?.active ? FIT_VALUE : value.toString()}
      onValueChange={handleValueChange}
      disabled={disabled}
    >
      <SelectTrigger
        className={cn(compact ? 'h-7 min-w-[55px] text-xs' : 'h-8 min-w-[70px] text-sm', className)}
        aria-label={t('zoom.ariaLabel', { label: displayLabel })}
      >
        <SelectValue placeholder="100%">{displayLabel}</SelectValue>
      </SelectTrigger>
      <SelectContent>
        {fit && <SelectItem value={FIT_VALUE}>{fit.label}</SelectItem>}
        {levels.map((level) => (
          <SelectItem key={level.value} value={level.value.toString()}>
            {level.label}
          </SelectItem>
        ))}
      </SelectContent>
    </Select>
  );
}

// Re-export types for compatibility
export type { ZoomControlProps as ZoomControlPropsType };
