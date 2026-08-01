"use client";

import { AlertCircle, CheckCircle2 } from "lucide-react";
import { useActionState } from "react";

import {
  changePassword,
  updateProfile,
  type SettingsState,
} from "@/lib/actions/settings";
import { Card, CardBody, CardHeader } from "@/components/ui/card";
import { Field, FormGrid, Input } from "@/components/ui/form";
import { SubmitButton } from "@/components/ui/submit-button";
import type { SessionUser } from "@/lib/auth";

const INITIAL: SettingsState = {};

function Feedback({ state }: { state: SettingsState }) {
  if (!state.error && !state.success) return null;

  return (
    <div
      data-tone={state.error ? "danger" : "success"}
      role="status"
      className="flex animate-rise-in items-center gap-2 rounded-md border border-[color:var(--tone-line)] bg-[color:var(--tone-soft)] px-3 py-2 text-xs text-[color:var(--tone-fg)]"
    >
      {state.error ? (
        <AlertCircle className="h-3.5 w-3.5 shrink-0" />
      ) : (
        <CheckCircle2 className="h-3.5 w-3.5 shrink-0" />
      )}
      {state.error ?? state.success}
    </div>
  );
}

export function ProfileForm({ user }: { user: SessionUser }) {
  const [state, action] = useActionState(updateProfile, INITIAL);

  return (
    <Card>
      <CardHeader title="Profil" description="Nom affiché et adresse de connexion." />
      <CardBody>
        <form action={action} className="space-y-4">
          <Feedback state={state} />
          <FormGrid>
            <Field label="Nom" htmlFor="profile-name" required>
              <Input id="profile-name" name="name" defaultValue={user.name} required />
            </Field>
            <Field label="Adresse e-mail" htmlFor="profile-email" required>
              <Input
                id="profile-email"
                name="email"
                type="email"
                defaultValue={user.email}
                required
              />
            </Field>
          </FormGrid>
          <div className="flex justify-end">
            <SubmitButton pendingLabel="Enregistrement…">Enregistrer</SubmitButton>
          </div>
        </form>
      </CardBody>
    </Card>
  );
}

export function PasswordForm() {
  const [state, action] = useActionState(changePassword, INITIAL);

  return (
    <Card>
      <CardHeader
        title="Mot de passe"
        description="10 caractères minimum. Utilisez un gestionnaire de mots de passe."
      />
      <CardBody>
        <form action={action} className="space-y-4">
          <Feedback state={state} />
          <FormGrid>
            <Field label="Mot de passe actuel" htmlFor="currentPassword" required span={2}>
              <Input
                id="currentPassword"
                name="currentPassword"
                type="password"
                autoComplete="current-password"
                required
              />
            </Field>
            <Field label="Nouveau mot de passe" htmlFor="newPassword" required>
              <Input
                id="newPassword"
                name="newPassword"
                type="password"
                autoComplete="new-password"
                required
              />
            </Field>
            <Field label="Confirmation" htmlFor="confirmPassword" required>
              <Input
                id="confirmPassword"
                name="confirmPassword"
                type="password"
                autoComplete="new-password"
                required
              />
            </Field>
          </FormGrid>
          <div className="flex justify-end">
            <SubmitButton pendingLabel="Modification…">Modifier le mot de passe</SubmitButton>
          </div>
        </form>
      </CardBody>
    </Card>
  );
}
