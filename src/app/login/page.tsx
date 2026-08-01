import type { Metadata } from "next";
import { redirect } from "next/navigation";

import { LoginForm } from "@/app/login/login-form";
import { getSessionUser, hasAccount } from "@/lib/auth";

export const metadata: Metadata = { title: "Connexion" };

/**
 * Écran d'entrée. Si aucun compte n'existe encore, la page bascule
 * automatiquement en mode installation : première visite = création du compte,
 * sans script à lancer à la main.
 */
export default async function LoginPage() {
  if (await getSessionUser()) redirect("/");
  const accountExists = await hasAccount();
  const mode = accountExists ? "login" : "setup";

  return (
    <div className="relative flex min-h-dvh items-center justify-center overflow-hidden px-4 py-10">
      {/* Halo très diffus : donne de la profondeur sans attirer l'œil. */}
      <div
        aria-hidden
        className="pointer-events-none absolute left-1/2 top-0 h-[420px] w-[720px] -translate-x-1/2 -translate-y-1/3 rounded-full opacity-[0.13] blur-[100px]"
        style={{ background: "var(--accent)" }}
      />

      <div className="relative w-full max-w-[380px] animate-rise-in">
        <div className="mb-7 flex flex-col items-center text-center">
          <span className="mb-4 flex h-11 w-11 items-center justify-center rounded-xl bg-[var(--accent)] text-base font-bold text-white shadow-[var(--shadow-md)]">
            i4
          </span>
          <h1 className="text-lg font-semibold text-ink">
            {mode === "setup" ? "Installation de la console" : "Ine4rt Console"}
          </h1>
          <p className="mt-1.5 max-w-[300px] text-xs leading-relaxed text-ink-muted">
            {mode === "setup"
              ? "Créez le compte administrateur. Cette étape n'apparaît qu'une seule fois."
              : "Espace privé de gestion des clients, sites et abonnements."}
          </p>
        </div>

        <div className="card p-6 shadow-[var(--shadow-md)]">
          <LoginForm mode={mode} />
        </div>

        <p className="mt-5 text-center text-[11px] text-ink-muted">
          Accès strictement personnel — aucune donnée client n'est exposée publiquement.
        </p>
      </div>
    </div>
  );
}
