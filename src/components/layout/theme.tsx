"use client";

import { Monitor, Moon, Sun } from "lucide-react";
import {
  createContext,
  useCallback,
  useContext,
  useEffect,
  useMemo,
  useState,
  type ReactNode,
} from "react";

import { cn } from "@/lib/utils";

/**
 * Thème clair / sombre / système.
 *
 * Le choix est appliqué avant le premier rendu par `themeInitScript` (injecté
 * dans le <head>) : aucun flash de thème clair au chargement, ce qui est la
 * première chose qu'on remarque sur un outil utilisé le soir.
 */

export type ThemePreference = "light" | "dark" | "system";

const STORAGE_KEY = "ineart-theme";

/** Exécuté avant l'hydratation — volontairement minimal et sans dépendances. */
export const themeInitScript = `
(function(){
  try {
    var stored = localStorage.getItem("${STORAGE_KEY}") || "system";
    var dark = stored === "dark" ||
      (stored === "system" && window.matchMedia("(prefers-color-scheme: dark)").matches);
    document.documentElement.classList.toggle("dark", dark);
  } catch (e) {}
})();
`;

const ThemeContext = createContext<{
  preference: ThemePreference;
  setPreference: (value: ThemePreference) => void;
} | null>(null);

export function useTheme() {
  const context = useContext(ThemeContext);
  if (!context) throw new Error("useTheme doit être utilisé à l'intérieur de <ThemeProvider>");
  return context;
}

function applyTheme(preference: ThemePreference) {
  const dark =
    preference === "dark" ||
    (preference === "system" && window.matchMedia("(prefers-color-scheme: dark)").matches);
  document.documentElement.classList.toggle("dark", dark);
}

export function ThemeProvider({ children }: { children: ReactNode }) {
  const [preference, setPreferenceState] = useState<ThemePreference>("system");

  useEffect(() => {
    const stored = (localStorage.getItem(STORAGE_KEY) as ThemePreference | null) ?? "system";
    setPreferenceState(stored);

    // Suit les changements de l'OS tant que l'utilisateur reste en mode « système ».
    const media = window.matchMedia("(prefers-color-scheme: dark)");
    const onChange = () => {
      if ((localStorage.getItem(STORAGE_KEY) ?? "system") === "system") applyTheme("system");
    };
    media.addEventListener("change", onChange);
    return () => media.removeEventListener("change", onChange);
  }, []);

  const setPreference = useCallback((value: ThemePreference) => {
    setPreferenceState(value);
    localStorage.setItem(STORAGE_KEY, value);
    applyTheme(value);
  }, []);

  const value = useMemo(() => ({ preference, setPreference }), [preference, setPreference]);

  return <ThemeContext.Provider value={value}>{children}</ThemeContext.Provider>;
}

const OPTIONS: { value: ThemePreference; icon: typeof Sun; label: string }[] = [
  { value: "light", icon: Sun, label: "Clair" },
  { value: "dark", icon: Moon, label: "Sombre" },
  { value: "system", icon: Monitor, label: "Système" },
];

/** Sélecteur segmenté à trois positions. */
export function ThemeToggle({ className }: { className?: string }) {
  const { preference, setPreference } = useTheme();

  return (
    <div
      className={cn("inline-flex rounded-md border border-line bg-surface-inset p-0.5", className)}
      role="radiogroup"
      aria-label="Thème de l'interface"
    >
      {OPTIONS.map((option) => {
        const Icon = option.icon;
        const active = preference === option.value;
        return (
          <button
            key={option.value}
            type="button"
            role="radio"
            aria-checked={active}
            aria-label={option.label}
            title={option.label}
            onClick={() => setPreference(option.value)}
            className={cn(
              "flex h-6 w-7 items-center justify-center rounded-[5px] transition-all duration-150",
              active
                ? "bg-surface text-ink shadow-[var(--shadow-xs)]"
                : "text-ink-muted hover:text-ink",
            )}
          >
            <Icon className="h-3.5 w-3.5" />
          </button>
        );
      })}
    </div>
  );
}
