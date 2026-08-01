import type { Tone } from "@/lib/constants";
import { cn } from "@/lib/utils";

/** Barre de progression fine (avancement projet, part d'un total). */
export function Progress({
  value,
  tone = "accent",
  className,
  showValue,
  size = "md",
}: {
  value: number;
  tone?: Tone;
  className?: string;
  showValue?: boolean;
  size?: "sm" | "md";
}) {
  const clamped = Math.max(0, Math.min(100, Math.round(value)));

  return (
    <div className={cn("flex items-center gap-2", className)}>
      <div
        data-tone={tone}
        className={cn(
          "relative min-w-0 flex-1 overflow-hidden rounded-full bg-surface-inset",
          size === "sm" ? "h-1" : "h-1.5",
        )}
        role="progressbar"
        aria-valuenow={clamped}
        aria-valuemin={0}
        aria-valuemax={100}
      >
        <div
          className="h-full rounded-full bg-[color:var(--tone-fg)] transition-[width] duration-700 ease-[var(--ease-out-quint)]"
          style={{ width: `${clamped}%` }}
        />
      </div>
      {showValue ? (
        <span className="w-9 shrink-0 text-right text-xs text-ink-muted tabular-nums">
          {clamped}%
        </span>
      ) : null}
    </div>
  );
}
