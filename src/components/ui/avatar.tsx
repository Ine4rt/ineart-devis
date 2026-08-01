import { cn, hueFromString, initials } from "@/lib/utils";

const SIZES = {
  xs: "h-6 w-6 text-[10px]",
  sm: "h-7 w-7 text-[11px]",
  md: "h-9 w-9 text-xs",
  lg: "h-12 w-12 text-sm",
  xl: "h-16 w-16 text-lg",
} as const;

/**
 * Avatar de client. Si aucun logo n'est fourni, on génère des initiales sur un
 * fond dont la teinte est dérivée du nom : chaque client garde durablement la
 * même couleur, ce qui aide à le repérer dans les listes.
 */
export function Avatar({
  name,
  src,
  size = "md",
  className,
}: {
  name: string;
  src?: string | null;
  size?: keyof typeof SIZES;
  className?: string;
}) {
  const base = cn(
    "flex shrink-0 items-center justify-center overflow-hidden rounded-md border border-line font-semibold select-none",
    SIZES[size],
    className,
  );

  if (src) {
    // Logos clients : sources locales ou distantes arbitraires, on reste sur <img>.
    // eslint-disable-next-line @next/next/no-img-element
    return <img src={src} alt={name} className={cn(base, "bg-surface object-contain")} />;
  }

  const hue = hueFromString(name);

  return (
    <span
      className={base}
      // `light-dark()` s'appuie sur le color-scheme posé par le thème : une seule
      // déclaration couvre les deux modes, sans dupliquer la logique en JS.
      style={{
        backgroundColor: `light-dark(oklch(0.93 0.045 ${hue}), oklch(0.28 0.05 ${hue}))`,
        color: `light-dark(oklch(0.42 0.11 ${hue}), oklch(0.84 0.09 ${hue}))`,
        borderColor: `light-dark(oklch(0.86 0.05 ${hue}), oklch(0.36 0.055 ${hue}))`,
      }}
      aria-hidden
    >
      {initials(name)}
    </span>
  );
}
