"use client";

import { Search, SlidersHorizontal, X } from "lucide-react";
import { usePathname, useRouter, useSearchParams } from "next/navigation";
import { useEffect, useRef, useState, type ReactNode } from "react";

import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

/**
 * Barre de filtres des écrans de liste.
 *
 * L'état vit dans l'URL, pas dans le composant : un tri ou un filtre reste
 * partageable, rechargeable et présent dans l'historique du navigateur. La
 * saisie texte est débattue à 250 ms pour ne pas relancer une requête serveur
 * à chaque caractère.
 */

import type { FilterOption } from "@/lib/filters";

export interface FilterDefinition {
  name: string;
  label: string;
  options: FilterOption[];
}

export function ListToolbar({
  searchPlaceholder = "Rechercher…",
  filters = [],
  children,
  className,
}: {
  searchPlaceholder?: string;
  filters?: FilterDefinition[];
  children?: ReactNode;
  className?: string;
}) {
  const router = useRouter();
  const pathname = usePathname();
  const searchParams = useSearchParams();

  const [query, setQuery] = useState(searchParams.get("q") ?? "");
  const [showFilters, setShowFilters] = useState(false);
  const firstRender = useRef(true);

  function buildUrl(updates: Record<string, string>) {
    const params = new URLSearchParams(searchParams.toString());
    for (const [key, value] of Object.entries(updates)) {
      if (value) params.set(key, value);
      else params.delete(key);
    }
    const search = params.toString();
    return search ? `${pathname}?${search}` : pathname;
  }

  useEffect(() => {
    // Ne pas rejouer la navigation au montage (l'URL reflète déjà l'état).
    if (firstRender.current) {
      firstRender.current = false;
      return;
    }
    const timer = window.setTimeout(() => {
      router.replace(buildUrl({ q: query }), { scroll: false });
    }, 250);
    return () => window.clearTimeout(timer);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [query]);

  const activeFilters = filters.filter((filter) => searchParams.get(filter.name));
  const hasActive = activeFilters.length > 0 || query.length > 0;

  return (
    <div className={cn("space-y-2", className)}>
      <div className="flex flex-wrap items-center gap-2">
        <div className="relative min-w-0 flex-1 sm:max-w-xs">
          <Search className="pointer-events-none absolute left-2.5 top-1/2 h-3.5 w-3.5 -translate-y-1/2 text-ink-muted" />
          <input
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            placeholder={searchPlaceholder}
            className="field h-8 pl-8"
            aria-label={searchPlaceholder}
          />
          {query ? (
            <button
              type="button"
              onClick={() => setQuery("")}
              className="absolute right-2 top-1/2 -translate-y-1/2 text-ink-muted hover:text-ink"
              aria-label="Effacer la recherche"
            >
              <X className="h-3.5 w-3.5" />
            </button>
          ) : null}
        </div>

        {filters.length > 0 ? (
          <Button
            variant={showFilters || activeFilters.length > 0 ? "default" : "ghost"}
            onClick={() => setShowFilters((value) => !value)}
          >
            <SlidersHorizontal className="h-3.5 w-3.5" />
            Filtres
            {activeFilters.length > 0 ? (
              <span className="rounded-full bg-accent-soft px-1.5 text-[10px] font-semibold text-accent-text tabular-nums">
                {activeFilters.length}
              </span>
            ) : null}
          </Button>
        ) : null}

        {hasActive ? (
          <Button variant="ghost" onClick={() => { setQuery(""); router.replace(pathname, { scroll: false }); }}>
            Réinitialiser
          </Button>
        ) : null}

        <div className="ml-auto flex items-center gap-2">{children}</div>
      </div>

      {showFilters || activeFilters.length > 0 ? (
        <div className="flex animate-rise-in flex-wrap items-center gap-2 rounded-md border border-line bg-surface-subtle p-2">
          {filters.map((filter) => (
            <label key={filter.name} className="flex items-center gap-1.5">
              <span className="text-[11px] font-medium text-ink-muted">{filter.label}</span>
              <select
                className="field h-7 py-0 text-xs"
                value={searchParams.get(filter.name) ?? ""}
                onChange={(event) =>
                  router.replace(buildUrl({ [filter.name]: event.target.value }), { scroll: false })
                }
              >
                <option value="">Tous</option>
                {filter.options.map((option) => (
                  <option key={option.value} value={option.value}>
                    {option.label}
                  </option>
                ))}
              </select>
            </label>
          ))}
        </div>
      ) : null}
    </div>
  );
}
