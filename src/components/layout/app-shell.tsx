"use client";

import { Menu, PanelLeftClose, PanelLeftOpen, Plus, Search, X } from "lucide-react";
import Link from "next/link";
import { usePathname } from "next/navigation";
import { useEffect, useState, type ReactNode } from "react";

import { useCommandPalette } from "@/components/layout/command-palette";
import { ThemeToggle } from "@/components/layout/theme";
import { UserMenu } from "@/components/layout/user-menu";
import { Button } from "@/components/ui/button";
import type { SessionUser } from "@/lib/auth";
import { NAVIGATION, SETTINGS_LINK, isNavLinkActive } from "@/lib/navigation";
import { cn } from "@/lib/utils";

/**
 * Coquille applicative : rail de navigation à gauche, barre supérieure fine,
 * contenu au centre.
 *
 * Le rail se réduit en icônes (état mémorisé) pour libérer de la largeur sur les
 * écrans denses — tableaux d'abonnements, échéancier — sans perdre le repère de
 * navigation. Sur mobile il devient un tiroir.
 */

const COLLAPSE_KEY = "ineart-sidebar-collapsed";

export function AppShell({
  user,
  children,
  badges,
}: {
  user: SessionUser;
  children: ReactNode;
  /** Compteurs affichés en pastille (impayés, tâches en retard…). */
  badges?: Record<string, number>;
}) {
  const pathname = usePathname();
  const [collapsed, setCollapsed] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);

  useEffect(() => {
    setCollapsed(localStorage.getItem(COLLAPSE_KEY) === "1");
  }, []);

  useEffect(() => {
    setMobileOpen(false);
  }, [pathname]);

  function toggleCollapsed() {
    setCollapsed((value) => {
      localStorage.setItem(COLLAPSE_KEY, value ? "0" : "1");
      return !value;
    });
  }

  return (
    <div className="flex min-h-dvh bg-canvas">
      {/* Voile du tiroir mobile */}
      {mobileOpen ? (
        <div
          className="fixed inset-0 z-30 animate-fade-in bg-[var(--overlay)] lg:hidden"
          onClick={() => setMobileOpen(false)}
        />
      ) : null}

      <aside
        className={cn(
          "fixed inset-y-0 left-0 z-40 flex flex-col border-r border-line bg-canvas-subtle transition-[width,transform] duration-250 ease-[var(--ease-out-quint)] lg:sticky lg:top-0 lg:h-dvh lg:translate-x-0",
          collapsed ? "lg:w-[60px]" : "lg:w-[228px]",
          "w-[248px]",
          mobileOpen ? "translate-x-0" : "-translate-x-full",
        )}
      >
        <div className="flex h-14 shrink-0 items-center gap-2 px-3">
          <Link href="/" className="flex min-w-0 items-center gap-2.5">
            <Logo />
            {!collapsed ? (
              <span className="min-w-0">
                <span className="block truncate text-[13px] font-semibold leading-tight text-ink">
                  IneWeb
                </span>
                <span className="block truncate text-[10px] uppercase tracking-wider text-ink-muted">
                  Console
                </span>
              </span>
            ) : null}
          </Link>
          <button
            type="button"
            onClick={() => setMobileOpen(false)}
            className="ml-auto text-ink-muted lg:hidden"
            aria-label="Fermer le menu"
          >
            <X className="h-4 w-4" />
          </button>
        </div>

        <nav className="min-h-0 flex-1 space-y-4 overflow-y-auto px-2 pb-3">
          {NAVIGATION.map((group) => (
            <div key={group.label}>
              {!collapsed ? (
                <div className="px-2 pb-1 pt-1 section-label">{group.label}</div>
              ) : (
                <div className="mx-2 my-2 h-px bg-line" />
              )}
              <ul className="space-y-0.5">
                {group.links.map((link) => {
                  const Icon = link.icon;
                  const active = isNavLinkActive(link, pathname);
                  const badge = badges?.[link.href];
                  return (
                    <li key={link.href}>
                      <Link
                        href={link.href}
                        data-active={active}
                        title={collapsed ? link.label : undefined}
                        className={cn("nav-item", collapsed && "justify-center px-0")}
                      >
                        <Icon className="h-4 w-4 shrink-0" />
                        {!collapsed ? (
                          <>
                            <span className="min-w-0 flex-1 truncate">{link.label}</span>
                            {badge ? (
                              <span
                                data-tone="danger"
                                className="rounded-full bg-[color:var(--tone-soft)] px-1.5 text-[10px] font-semibold text-[color:var(--tone-fg)] tabular-nums"
                              >
                                {badge}
                              </span>
                            ) : null}
                          </>
                        ) : badge ? (
                          <span
                            data-tone="danger"
                            className="absolute right-1 top-1 h-1.5 w-1.5 rounded-full bg-[color:var(--tone-fg)]"
                          />
                        ) : null}
                      </Link>
                    </li>
                  );
                })}
              </ul>
            </div>
          ))}
        </nav>

        <div className="shrink-0 space-y-1 border-t border-line p-2">
          <Link
            href={SETTINGS_LINK.href}
            data-active={pathname.startsWith(SETTINGS_LINK.href)}
            title={collapsed ? SETTINGS_LINK.label : undefined}
            className={cn("nav-item", collapsed && "justify-center px-0")}
          >
            <SETTINGS_LINK.icon className="h-4 w-4 shrink-0" />
            {!collapsed ? <span className="truncate">{SETTINGS_LINK.label}</span> : null}
          </Link>

          {!collapsed ? (
            <div className="flex items-center justify-between gap-2 px-1 pt-1">
              <ThemeToggle />
              <button
                type="button"
                onClick={toggleCollapsed}
                className="hidden h-6 w-6 items-center justify-center rounded-md text-ink-muted transition-colors hover:bg-surface-inset hover:text-ink lg:flex"
                aria-label="Réduire le menu"
                title="Réduire le menu"
              >
                <PanelLeftClose className="h-3.5 w-3.5" />
              </button>
            </div>
          ) : (
            <button
              type="button"
              onClick={toggleCollapsed}
              className="hidden h-8 w-full items-center justify-center rounded-md text-ink-muted transition-colors hover:bg-surface-inset hover:text-ink lg:flex"
              aria-label="Déplier le menu"
              title="Déplier le menu"
            >
              <PanelLeftOpen className="h-3.5 w-3.5" />
            </button>
          )}
        </div>
      </aside>

      <div className="flex min-w-0 flex-1 flex-col">
        <TopBar user={user} onOpenMobile={() => setMobileOpen(true)} />
        <main className="min-w-0 flex-1 px-4 py-5 sm:px-6 lg:px-8">
          <div className="mx-auto w-full max-w-[1440px]">{children}</div>
        </main>
      </div>
    </div>
  );
}

function TopBar({ user, onOpenMobile }: { user: SessionUser; onOpenMobile: () => void }) {
  const { open } = useCommandPalette();

  return (
    <header className="sticky top-0 z-20 flex h-14 shrink-0 items-center gap-3 border-b border-line bg-canvas/80 px-4 backdrop-blur-md sm:px-6 lg:px-8">
      <button
        type="button"
        onClick={onOpenMobile}
        className="text-ink-secondary lg:hidden"
        aria-label="Ouvrir le menu"
      >
        <Menu className="h-5 w-5" />
      </button>

      {/* Champ factice : il ouvre la palette, qui fait bien plus qu'un input. */}
      <button
        type="button"
        onClick={open}
        className="group flex h-8 w-full max-w-sm items-center gap-2 rounded-md border border-line bg-surface px-2.5 text-left transition-colors hover:border-line-strong"
      >
        <Search className="h-3.5 w-3.5 shrink-0 text-ink-muted" />
        <span className="min-w-0 flex-1 truncate text-xs text-ink-muted">
          Rechercher partout…
        </span>
        <span className="hidden shrink-0 items-center gap-0.5 sm:flex">
          <kbd className="kbd">⌘</kbd>
          <kbd className="kbd">K</kbd>
        </span>
      </button>

      <div className="ml-auto flex items-center gap-2">
        <QuickCreate />
        <UserMenu user={user} />
      </div>
    </header>
  );
}

/** Création rapide depuis n'importe quel écran. */
function QuickCreate() {
  const [open, setOpen] = useState(false);

  const actions = [
    { href: "/clients/nouveau", label: "Client" },
    { href: "/projets/nouveau", label: "Projet" },
    { href: "/domaines/nouveau", label: "Domaine" },
    { href: "/hebergements/nouveau", label: "Hébergement" },
    { href: "/abonnements/nouveau", label: "Abonnement" },
    { href: "/devis/nouveau", label: "Devis" },
  ];

  useEffect(() => {
    if (!open) return;
    const close = () => setOpen(false);
    window.addEventListener("click", close);
    return () => window.removeEventListener("click", close);
  }, [open]);

  return (
    <div className="relative" onClick={(event) => event.stopPropagation()}>
      <Button variant="primary" size="md" onClick={() => setOpen((value) => !value)}>
        <Plus className="h-3.5 w-3.5" />
        <span className="hidden sm:inline">Créer</span>
      </Button>
      {open ? (
        <div className="absolute right-0 top-[calc(100%+6px)] z-40 w-48 animate-scale-in origin-top-right overflow-hidden rounded-lg border border-line bg-surface-raised p-1 shadow-[var(--shadow-lg)]">
          {actions.map((action) => (
            <Link
              key={action.href}
              href={action.href}
              onClick={() => setOpen(false)}
              className="block rounded-md px-2.5 py-1.5 text-xs font-medium text-ink-secondary transition-colors hover:bg-surface-inset hover:text-ink"
            >
              {action.label}
            </Link>
          ))}
        </div>
      ) : null}
    </div>
  );
}

function Logo() {
  return (
    <span className="flex h-7 w-7 shrink-0 items-center justify-center rounded-md bg-[var(--accent)] text-[13px] font-bold text-white shadow-[var(--shadow-xs)]">
      iW
    </span>
  );
}
