"use client";

import {
  ArrowRight,
  Building2,
  CornerDownLeft,
  FileText,
  FolderKanban,
  Globe,
  Loader2,
  Paperclip,
  Search,
  Server,
  StickyNote,
  ListChecks,
  Waypoints,
} from "lucide-react";
import { useRouter } from "next/navigation";
import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from "react";
import { createPortal } from "react-dom";

import { ALL_NAV_LINKS } from "@/lib/navigation";
import type { SearchEntity, SearchResult } from "@/lib/queries/search";
import { cn } from "@/lib/utils";

/**
 * Palette de commandes (⌘K / Ctrl+K).
 *
 * C'est le raccourci central de l'outil : atteindre n'importe quel client,
 * domaine, projet ou écran sans quitter le clavier. Elle combine deux sources —
 * la navigation (instantanée, filtrée localement) et la recherche serveur
 * (débattue à 180 ms, requête précédente annulée).
 */

const ENTITY_META: Record<SearchEntity, { label: string; icon: typeof Building2 }> = {
  client: { label: "Client", icon: Building2 },
  projet: { label: "Projet", icon: FolderKanban },
  domaine: { label: "Domaine", icon: Globe },
  hebergement: { label: "Hébergement", icon: Server },
  abonnement: { label: "Abonnement", icon: Waypoints },
  devis: { label: "Devis", icon: FileText },
  tache: { label: "Tâche", icon: ListChecks },
  note: { label: "Note", icon: StickyNote },
  document: { label: "Document", icon: Paperclip },
};

const CommandPaletteContext = createContext<{ open: () => void } | null>(null);

export function useCommandPalette() {
  const context = useContext(CommandPaletteContext);
  if (!context) throw new Error("useCommandPalette requiert <CommandPaletteProvider>");
  return context;
}

interface Row {
  key: string;
  title: string;
  subtitle?: string;
  group: string;
  href: string;
  icon: typeof Building2;
}

export function CommandPaletteProvider({ children }: { children: ReactNode }) {
  const router = useRouter();
  const [isOpen, setIsOpen] = useState(false);
  const [query, setQuery] = useState("");
  const [results, setResults] = useState<SearchResult[]>([]);
  const [loading, setLoading] = useState(false);
  const [activeIndex, setActiveIndex] = useState(0);
  const listRef = useRef<HTMLDivElement>(null);

  const open = useCallback(() => setIsOpen(true), []);
  const close = useCallback(() => {
    setIsOpen(false);
    setQuery("");
    setResults([]);
    setActiveIndex(0);
  }, []);

  // Raccourci global.
  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === "k") {
        event.preventDefault();
        setIsOpen((value) => !value);
      }
    };
    window.addEventListener("keydown", onKeyDown);
    return () => window.removeEventListener("keydown", onKeyDown);
  }, []);

  // Recherche serveur débattue, avec annulation de la requête précédente.
  useEffect(() => {
    if (!isOpen) return;
    const trimmed = query.trim();
    if (trimmed.length < 2) {
      setResults([]);
      setLoading(false);
      return;
    }

    const controller = new AbortController();
    setLoading(true);
    const timer = window.setTimeout(async () => {
      try {
        const response = await fetch(`/api/search?q=${encodeURIComponent(trimmed)}`, {
          signal: controller.signal,
        });
        const data = (await response.json()) as { results: SearchResult[] };
        setResults(data.results ?? []);
      } catch {
        /* requête annulée : la suivante fait foi */
      } finally {
        setLoading(false);
      }
    }, 180);

    return () => {
      controller.abort();
      window.clearTimeout(timer);
    };
  }, [query, isOpen]);

  const rows = useMemo<Row[]>(() => {
    const normalized = query.trim().toLowerCase();

    const navRows: Row[] = ALL_NAV_LINKS.filter(
      (link) => !normalized || link.label.toLowerCase().includes(normalized),
    ).map((link) => ({
      key: `nav:${link.href}`,
      title: link.label,
      subtitle: link.description,
      group: "Navigation",
      href: link.href,
      icon: link.icon,
    }));

    const searchRows: Row[] = results.map((result) => ({
      key: `${result.entity}:${result.id}`,
      title: result.title,
      subtitle: result.subtitle,
      group: ENTITY_META[result.entity].label,
      href: result.href,
      icon: ENTITY_META[result.entity].icon,
    }));

    // Les résultats de données passent devant : quand on tape, on cherche une
    // fiche précise bien plus souvent qu'un écran.
    return normalized.length >= 2 ? [...searchRows, ...navRows] : navRows;
  }, [query, results]);

  useEffect(() => setActiveIndex(0), [rows.length]);

  const select = useCallback(
    (row: Row | undefined) => {
      if (!row) return;
      close();
      router.push(row.href);
    },
    [close, router],
  );

  const onKeyDown = (event: React.KeyboardEvent) => {
    if (event.key === "ArrowDown") {
      event.preventDefault();
      setActiveIndex((index) => Math.min(index + 1, rows.length - 1));
    } else if (event.key === "ArrowUp") {
      event.preventDefault();
      setActiveIndex((index) => Math.max(index - 1, 0));
    } else if (event.key === "Enter") {
      event.preventDefault();
      select(rows[activeIndex]);
    } else if (event.key === "Escape") {
      close();
    }
  };

  // Garde l'élément actif visible pendant la navigation clavier.
  useEffect(() => {
    listRef.current
      ?.querySelector<HTMLElement>(`[data-index="${activeIndex}"]`)
      ?.scrollIntoView({ block: "nearest" });
  }, [activeIndex]);

  const value = useMemo(() => ({ open }), [open]);

  let lastGroup = "";

  return (
    <CommandPaletteContext.Provider value={value}>
      {children}
      {isOpen
        ? createPortal(
            <div className="fixed inset-0 z-[70] flex items-start justify-center p-4 pt-[12vh]">
              <div
                className="absolute inset-0 animate-fade-in bg-[var(--overlay)] backdrop-blur-[3px]"
                onClick={close}
              />

              <div className="relative z-10 flex max-h-[70vh] w-full max-w-2xl animate-scale-in flex-col overflow-hidden rounded-xl border border-line bg-surface-raised shadow-[var(--shadow-lg)]">
                <div className="flex items-center gap-3 border-b border-line px-4">
                  {loading ? (
                    <Loader2 className="h-4 w-4 shrink-0 animate-spin text-ink-muted" />
                  ) : (
                    <Search className="h-4 w-4 shrink-0 text-ink-muted" />
                  )}
                  <input
                    autoFocus
                    value={query}
                    onChange={(event) => setQuery(event.target.value)}
                    onKeyDown={onKeyDown}
                    placeholder="Rechercher un client, un domaine, un projet…"
                    className="h-12 w-full bg-transparent text-sm text-ink outline-none placeholder:text-ink-muted"
                    aria-label="Recherche globale"
                  />
                  <kbd className="kbd">esc</kbd>
                </div>

                <div ref={listRef} className="min-h-0 flex-1 overflow-y-auto p-1.5">
                  {rows.length === 0 ? (
                    <p className="px-3 py-10 text-center text-xs text-ink-muted">
                      {query.trim().length < 2
                        ? "Saisissez au moins deux caractères."
                        : `Aucun résultat pour « ${query.trim()} ».`}
                    </p>
                  ) : (
                    rows.map((row, index) => {
                      const showGroup = row.group !== lastGroup;
                      lastGroup = row.group;
                      const Icon = row.icon;
                      const active = index === activeIndex;

                      return (
                        <div key={row.key}>
                          {showGroup ? (
                            <div className="px-2.5 pb-1 pt-3 section-label first:pt-1">
                              {row.group}
                            </div>
                          ) : null}
                          <button
                            type="button"
                            data-index={index}
                            onMouseMove={() => setActiveIndex(index)}
                            onClick={() => select(row)}
                            className={cn(
                              "flex w-full items-center gap-3 rounded-md px-2.5 py-2 text-left transition-colors",
                              active ? "bg-accent-soft" : "hover:bg-surface-inset",
                            )}
                          >
                            <Icon
                              className={cn(
                                "h-4 w-4 shrink-0",
                                active ? "text-accent-text" : "text-ink-muted",
                              )}
                            />
                            <span className="min-w-0 flex-1">
                              <span
                                className={cn(
                                  "block truncate text-[13px] font-medium",
                                  active ? "text-accent-text" : "text-ink",
                                )}
                              >
                                {row.title}
                              </span>
                              {row.subtitle ? (
                                <span className="block truncate text-xs text-ink-muted">
                                  {row.subtitle}
                                </span>
                              ) : null}
                            </span>
                            {active ? (
                              <CornerDownLeft className="h-3.5 w-3.5 shrink-0 text-accent-text" />
                            ) : (
                              <ArrowRight className="h-3.5 w-3.5 shrink-0 text-transparent" />
                            )}
                          </button>
                        </div>
                      );
                    })
                  )}
                </div>

                <div className="flex items-center gap-4 border-t border-line bg-surface-subtle px-4 py-2 text-[11px] text-ink-muted">
                  <span className="flex items-center gap-1.5">
                    <kbd className="kbd">↑</kbd>
                    <kbd className="kbd">↓</kbd> naviguer
                  </span>
                  <span className="flex items-center gap-1.5">
                    <kbd className="kbd">↵</kbd> ouvrir
                  </span>
                  <span className="ml-auto flex items-center gap-1.5">
                    <kbd className="kbd">⌘</kbd>
                    <kbd className="kbd">K</kbd>
                  </span>
                </div>
              </div>
            </div>,
            document.body,
          )
        : null}
    </CommandPaletteContext.Provider>
  );
}
