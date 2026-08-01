"use client";

import { useCallback, useId, useMemo, useState } from "react";

import { formatMoney, formatNumber } from "@/lib/format";
import { cn } from "@/lib/utils";

/**
 * Graphiques maison, en SVG pur.
 *
 * Pourquoi pas une bibliothèque : le poids (aucune dépendance supplémentaire),
 * et surtout le contrôle. Les couleurs viennent des variables `--chart-*` du
 * thème, donc les graphiques suivent le mode clair/sombre sans code JS, et le
 * rendu reste dans la même langue visuelle que le reste de l'outil.
 *
 * Règles appliquées partout :
 *  - une seule échelle par graphique (jamais de double axe) ;
 *  - traits fins, grille en retrait, étiquettes sélectives ;
 *  - survol systématique (infobulle) : un graphique HTML est interactif ;
 *  - l'ordre des couleurs est fixe, jamais recyclé.
 */

export const CHART_COLORS = [
  "var(--chart-1)",
  "var(--chart-2)",
  "var(--chart-3)",
  "var(--chart-4)",
  "var(--chart-5)",
  "var(--chart-6)",
] as const;

export interface Point {
  label: string;
  value: number;
}

/**
 * Les graphiques sont des composants client : on ne peut pas leur passer une
 * fonction de formatage depuis un composant serveur (les props traversent une
 * frontière sérialisable). On passe donc un *descripteur* — « montant »,
 * « montant compact », « nombre » — que le graphique applique lui-même.
 */
export type ValueFormat = "money" | "money-compact" | "number";

export function useValueFormatter(format: ValueFormat, currency = "EUR") {
  return useCallback(
    (value: number) => {
      if (format === "number") return formatNumber(value);
      return formatMoney(value, { currency, compact: format === "money-compact" });
    },
    [format, currency],
  );
}

function niceCeil(value: number): number {
  if (value <= 0) return 1;
  const magnitude = 10 ** Math.floor(Math.log10(value));
  const normalized = value / magnitude;
  const step = normalized <= 1 ? 1 : normalized <= 2 ? 2 : normalized <= 5 ? 5 : 10;
  return step * magnitude;
}

function Tooltip({
  x,
  y,
  title,
  value,
  align,
}: {
  x: number;
  y: number;
  title: string;
  value: string;
  align: "left" | "right";
}) {
  return (
    <div
      className="pointer-events-none absolute z-10 -translate-y-1/2 rounded-md border border-line bg-surface-raised px-2.5 py-1.5 shadow-[var(--shadow-md)]"
      style={{
        left: `${x}%`,
        top: `${y}%`,
        transform: `translate(${align === "left" ? "-100%" : "0"}, -50%) translateX(${
          align === "left" ? "-10px" : "10px"
        })`,
      }}
    >
      <div className="text-[10px] font-medium uppercase tracking-wide text-ink-muted">{title}</div>
      <div className="text-xs font-semibold text-ink tabular-nums">{value}</div>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Courbe / aire — évolution dans le temps
// ---------------------------------------------------------------------------

export function AreaChart({
  data,
  format = "money",
  currency,
  height = 180,
  className,
  colorIndex = 0,
}: {
  data: Point[];
  format?: ValueFormat;
  currency?: string;
  height?: number;
  className?: string;
  colorIndex?: number;
}) {
  const formatValue = useValueFormatter(format, currency);
  const gradientId = useId();
  const [hover, setHover] = useState<number | null>(null);

  const geometry = useMemo(() => {
    const max = niceCeil(Math.max(...data.map((d) => d.value), 0));
    const stepX = data.length > 1 ? 100 / (data.length - 1) : 0;
    const points = data.map((point, index) => ({
      ...point,
      x: index * stepX,
      y: max === 0 ? 100 : 100 - (point.value / max) * 100,
    }));
    return { max, points };
  }, [data]);

  if (data.length === 0) return null;

  const { max, points } = geometry;
  const color = CHART_COLORS[colorIndex % CHART_COLORS.length];
  const line = points.map((p) => `${p.x},${p.y}`).join(" ");
  const area = `${line} 100,100 0,100`;
  const active = hover !== null ? points[hover] : null;

  return (
    <div className={cn("relative", className)}>
      <svg
        viewBox="0 0 100 100"
        preserveAspectRatio="none"
        className="w-full overflow-visible"
        style={{ height }}
        role="img"
        aria-label={`Évolution : de ${formatValue(data[0].value)} à ${formatValue(
          data[data.length - 1].value,
        )}`}
      >
        <defs>
          <linearGradient id={gradientId} x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stopColor={color} stopOpacity="0.22" />
            <stop offset="100%" stopColor={color} stopOpacity="0" />
          </linearGradient>
        </defs>

        {/* Grille en retrait : elle situe, elle ne concurrence pas la donnée. */}
        {[0, 25, 50, 75, 100].map((y) => (
          <line
            key={y}
            x1="0"
            y1={y}
            x2="100"
            y2={y}
            stroke="var(--chart-grid)"
            strokeWidth="0.5"
            vectorEffect="non-scaling-stroke"
          />
        ))}

        <polygon points={area} fill={`url(#${gradientId})`} />
        <polyline
          points={line}
          fill="none"
          stroke={color}
          strokeWidth="2"
          strokeLinecap="round"
          strokeLinejoin="round"
          vectorEffect="non-scaling-stroke"
        />

        {active ? (
          <>
            <line
              x1={active.x}
              y1="0"
              x2={active.x}
              y2="100"
              stroke="var(--chart-axis)"
              strokeWidth="1"
              strokeDasharray="3 3"
              vectorEffect="non-scaling-stroke"
            />
            <circle
              cx={active.x}
              cy={active.y}
              r="4"
              fill={color}
              stroke="var(--surface)"
              strokeWidth="2"
              vectorEffect="non-scaling-stroke"
            />
          </>
        ) : null}
      </svg>

      {/* Zones de survol : plus larges que la marque, comme il se doit. */}
      <div className="absolute inset-0 flex" onMouseLeave={() => setHover(null)}>
        {points.map((point, index) => (
          <button
            key={`${point.label}-${index}`}
            type="button"
            className="h-full flex-1 cursor-default"
            onMouseEnter={() => setHover(index)}
            onFocus={() => setHover(index)}
            aria-label={`${point.label} : ${formatValue(point.value)}`}
          />
        ))}
      </div>

      {active ? (
        <Tooltip
          x={active.x}
          y={active.y}
          title={active.label}
          value={formatValue(active.value)}
          align={active.x > 60 ? "left" : "right"}
        />
      ) : null}

      <div className="mt-2 flex justify-between text-[10px] text-ink-muted">
        {points.map((point, index) =>
          index === 0 || index === points.length - 1 || index === Math.floor(points.length / 2) ? (
            <span key={`${point.label}-${index}`}>{point.label}</span>
          ) : null,
        )}
      </div>
      <span className="sr-only">Maximum de l'échelle : {formatValue(max)}</span>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Barres verticales — comparaison de magnitudes
// ---------------------------------------------------------------------------

export function BarChart({
  data,
  format = "money",
  currency,
  height = 160,
  className,
  colorIndex = 0,
  highlightIndex,
}: {
  data: Point[];
  format?: ValueFormat;
  currency?: string;
  height?: number;
  className?: string;
  colorIndex?: number;
  /** Met une barre en avant (le mois courant, par exemple). */
  highlightIndex?: number;
}) {
  const formatValue = useValueFormatter(format, currency);
  const [hover, setHover] = useState<number | null>(null);
  const max = niceCeil(Math.max(...data.map((d) => d.value), 0));
  const color = CHART_COLORS[colorIndex % CHART_COLORS.length];

  if (data.length === 0) return null;

  return (
    <div className={cn("relative", className)}>
      <div className="flex items-end gap-1.5" style={{ height }} onMouseLeave={() => setHover(null)}>
        {data.map((point, index) => {
          const ratio = max === 0 ? 0 : point.value / max;
          const isHighlight = index === highlightIndex;
          const isHovered = hover === index;
          return (
            <button
              key={`${point.label}-${index}`}
              type="button"
              className="group relative flex h-full flex-1 cursor-default flex-col justify-end"
              onMouseEnter={() => setHover(index)}
              onFocus={() => setHover(index)}
              aria-label={`${point.label} : ${formatValue(point.value)}`}
            >
              <div
                className="chart-bar w-full rounded-t-[4px] transition-opacity duration-150"
                style={{
                  height: `${Math.max(ratio * 100, point.value > 0 ? 2 : 0)}%`,
                  backgroundColor: color,
                  opacity: isHovered ? 1 : isHighlight ? 0.95 : 0.62,
                  animationDelay: `${index * 35}ms`,
                }}
              />
              {isHighlight ? (
                <span className="absolute inset-x-0 -bottom-px h-0.5 rounded-full bg-[color:var(--chart-axis)]" />
              ) : null}
            </button>
          );
        })}
      </div>

      <div className="mt-2 flex gap-1.5 text-[10px] text-ink-muted">
        {data.map((point, index) => (
          <span
            key={`${point.label}-${index}`}
            className={cn(
              "flex-1 truncate text-center",
              index === highlightIndex && "font-semibold text-ink-secondary",
            )}
          >
            {point.label}
          </span>
        ))}
      </div>

      {hover !== null ? (
        <div className="pointer-events-none absolute -top-1 left-1/2 z-10 -translate-x-1/2 -translate-y-full rounded-md border border-line bg-surface-raised px-2.5 py-1.5 shadow-[var(--shadow-md)]">
          <div className="text-[10px] font-medium uppercase tracking-wide text-ink-muted">
            {data[hover].label}
          </div>
          <div className="text-xs font-semibold text-ink tabular-nums">
            {formatValue(data[hover].value)}
          </div>
        </div>
      ) : null}
    </div>
  );
}

// ---------------------------------------------------------------------------
// Anneau — répartition d'un total
// ---------------------------------------------------------------------------

export function DonutChart({
  data,
  format = "money-compact",
  currency,
  centerLabel,
  centerValue,
  size = 148,
  className,
}: {
  data: Point[];
  format?: ValueFormat;
  currency?: string;
  centerLabel: string;
  centerValue: string;
  size?: number;
  className?: string;
}) {
  const formatValue = useValueFormatter(format, currency);
  const [hover, setHover] = useState<number | null>(null);
  const total = data.reduce((sum, point) => sum + point.value, 0);

  if (total <= 0) return null;

  const radius = 42;
  const circumference = 2 * Math.PI * radius;
  /* Espace de 2px entre segments : la surface sépare les parts, pas un liseré. */
  const gap = 2;

  let offset = 0;
  const segments = data.map((point, index) => {
    const fraction = point.value / total;
    const length = Math.max(fraction * circumference - gap, 0.6);
    const segment = {
      ...point,
      fraction,
      color: CHART_COLORS[index % CHART_COLORS.length],
      dashArray: `${length} ${circumference - length}`,
      dashOffset: -offset,
    };
    offset += fraction * circumference;
    return segment;
  });

  return (
    <div className={cn("flex flex-wrap items-center gap-5", className)}>
      <div className="relative shrink-0" style={{ width: size, height: size }}>
        <svg viewBox="0 0 100 100" className="-rotate-90" width={size} height={size}>
          {segments.map((segment, index) => (
            <circle
              key={segment.label}
              cx="50"
              cy="50"
              r={radius}
              fill="none"
              stroke={segment.color}
              strokeWidth={hover === index ? 15 : 12}
              strokeDasharray={segment.dashArray}
              strokeDashoffset={segment.dashOffset}
              strokeLinecap="butt"
              className="transition-[stroke-width,opacity] duration-200"
              style={{ opacity: hover === null || hover === index ? 1 : 0.35 }}
              onMouseEnter={() => setHover(index)}
              onMouseLeave={() => setHover(null)}
            />
          ))}
        </svg>
        <div className="pointer-events-none absolute inset-0 flex flex-col items-center justify-center text-center">
          <span className="text-base font-semibold text-ink tabular-nums">
            {hover === null ? centerValue : formatValue(segments[hover].value)}
          </span>
          <span className="max-w-[80%] truncate text-[10px] uppercase tracking-wide text-ink-muted">
            {hover === null ? centerLabel : segments[hover].label}
          </span>
        </div>
      </div>

      {/* Légende toujours présente : l'identité ne repose jamais sur la seule couleur. */}
      <ul className="min-w-0 flex-1 space-y-1.5">
        {segments.map((segment, index) => (
          <li
            key={segment.label}
            className="flex items-center gap-2 text-xs"
            onMouseEnter={() => setHover(index)}
            onMouseLeave={() => setHover(null)}
          >
            <span
              className="h-2 w-2 shrink-0 rounded-full"
              style={{ backgroundColor: segment.color }}
            />
            <span className="min-w-0 flex-1 truncate text-ink-secondary">{segment.label}</span>
            <span className="shrink-0 font-medium text-ink tabular-nums">
              {formatValue(segment.value)}
            </span>
            <span className="w-9 shrink-0 text-right text-ink-muted tabular-nums">
              {Math.round(segment.fraction * 100)}%
            </span>
          </li>
        ))}
      </ul>
    </div>
  );
}

// ---------------------------------------------------------------------------
// Sparkline — tendance intégrée à une tuile de statistique
// ---------------------------------------------------------------------------

export function Sparkline({
  values,
  className,
  colorIndex = 0,
}: {
  values: number[];
  className?: string;
  colorIndex?: number;
}) {
  if (values.length < 2) return null;

  const max = Math.max(...values);
  const min = Math.min(...values);
  const span = max - min || 1;
  const points = values
    .map((value, index) => {
      const x = (index / (values.length - 1)) * 100;
      const y = 100 - ((value - min) / span) * 100;
      return `${x},${y}`;
    })
    .join(" ");

  return (
    <svg
      viewBox="0 0 100 100"
      preserveAspectRatio="none"
      className={cn("h-8 w-full overflow-visible", className)}
      aria-hidden
    >
      <polyline
        points={points}
        fill="none"
        stroke={CHART_COLORS[colorIndex % CHART_COLORS.length]}
        strokeWidth="2"
        strokeLinecap="round"
        strokeLinejoin="round"
        vectorEffect="non-scaling-stroke"
        opacity="0.85"
      />
    </svg>
  );
}

// ---------------------------------------------------------------------------
// Barre de répartition horizontale — compacte, pour les cartes latérales
// ---------------------------------------------------------------------------

export function StackedBar({
  data,
  className,
}: {
  data: (Point & { tone?: string })[];
  className?: string;
}) {
  const total = data.reduce((sum, point) => sum + point.value, 0);
  if (total <= 0) return null;

  return (
    <div className={cn("flex h-2 w-full gap-0.5 overflow-hidden rounded-full", className)}>
      {data.map((point, index) => (
        <div
          key={point.label}
          className="h-full first:rounded-l-full last:rounded-r-full"
          style={{
            width: `${(point.value / total) * 100}%`,
            backgroundColor: point.tone ?? CHART_COLORS[index % CHART_COLORS.length],
          }}
          title={`${point.label} : ${point.value}`}
        />
      ))}
    </div>
  );
}
