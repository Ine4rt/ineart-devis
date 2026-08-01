"use client";

import { AlertCircle } from "lucide-react";
import { useActionState } from "react";

import { loginAction, setupAction, type AuthFormState } from "@/app/login/actions";
import { Field, Input } from "@/components/ui/form";
import { SubmitButton } from "@/components/ui/submit-button";

const INITIAL: AuthFormState = {};

export function LoginForm({ mode }: { mode: "login" | "setup" }) {
  const [state, action] = useActionState(
    mode === "setup" ? setupAction : loginAction,
    INITIAL,
  );

  return (
    <form action={action} className="space-y-4">
      {state.error ? (
        <div
          data-tone="danger"
          role="alert"
          className="flex animate-rise-in items-start gap-2.5 rounded-md border border-[color:var(--tone-line)] bg-[color:var(--tone-soft)] px-3 py-2.5 text-xs text-[color:var(--tone-fg)]"
        >
          <AlertCircle className="mt-px h-3.5 w-3.5 shrink-0" />
          <span>{state.error}</span>
        </div>
      ) : null}

      {mode === "setup" ? (
        <Field label="Nom" htmlFor="name" required>
          <Input id="name" name="name" autoComplete="name" required placeholder="Dimitri" />
        </Field>
      ) : null}

      <Field label="Adresse e-mail" htmlFor="email" required>
        <Input
          id="email"
          name="email"
          type="email"
          autoComplete="email"
          required
          data-autofocus
          placeholder="vous@exemple.be"
        />
      </Field>

      <Field
        label="Mot de passe"
        htmlFor="password"
        required
        hint={mode === "setup" ? "10 caractères minimum. Utilisez un gestionnaire." : undefined}
      >
        <Input
          id="password"
          name="password"
          type="password"
          autoComplete={mode === "setup" ? "new-password" : "current-password"}
          required
        />
      </Field>

      {mode === "setup" ? (
        <Field label="Confirmer le mot de passe" htmlFor="confirm" required>
          <Input id="confirm" name="confirm" type="password" autoComplete="new-password" required />
        </Field>
      ) : null}

      <SubmitButton className="w-full" size="lg" pendingLabel="Vérification…">
        {mode === "setup" ? "Créer le compte" : "Se connecter"}
      </SubmitButton>
    </form>
  );
}
