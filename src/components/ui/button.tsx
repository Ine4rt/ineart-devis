import Link from "next/link";
import type { ComponentProps, ReactNode } from "react";

import { cn } from "@/lib/utils";

export type ButtonVariant = "primary" | "default" | "ghost" | "danger";
export type ButtonSize = "sm" | "md" | "lg" | "icon";

const VARIANTS: Record<ButtonVariant, string> = {
  primary: "btn-primary",
  default: "btn-default",
  ghost: "btn-ghost",
  danger: "btn-danger",
};

const SIZES: Record<ButtonSize, string> = {
  sm: "h-7 px-2.5 text-xs",
  md: "h-8 px-3",
  lg: "h-9.5 px-4 text-sm",
  icon: "h-8 w-8 p-0",
};

export function buttonClasses(
  variant: ButtonVariant = "default",
  size: ButtonSize = "md",
  className?: string,
) {
  return cn("btn", VARIANTS[variant], SIZES[size], className);
}

interface ButtonProps extends ComponentProps<"button"> {
  variant?: ButtonVariant;
  size?: ButtonSize;
}

export function Button({ variant, size, className, ...props }: ButtonProps) {
  return <button className={buttonClasses(variant, size, className)} {...props} />;
}

interface ButtonLinkProps extends ComponentProps<typeof Link> {
  variant?: ButtonVariant;
  size?: ButtonSize;
}

/** Même apparence qu'un bouton, mais navigue réellement (prefetch, clic milieu…). */
export function ButtonLink({ variant, size, className, ...props }: ButtonLinkProps) {
  return <Link className={buttonClasses(variant, size, className)} {...props} />;
}

/** Bouton d'action de barre d'outils : icône + libellé masqué en petit écran. */
export function ToolbarAction({
  icon,
  label,
  ...props
}: ButtonProps & { icon: ReactNode; label: string }) {
  return (
    <Button aria-label={label} {...props}>
      {icon}
      <span className="hidden sm:inline">{label}</span>
    </Button>
  );
}
