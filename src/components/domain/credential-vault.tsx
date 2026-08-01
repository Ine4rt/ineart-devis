"use client";

import type { CredentialKind } from "@prisma/client";
import { Eye, EyeOff, KeyRound, Loader2, Plus, Trash2 } from "lucide-react";
import { useState, useTransition } from "react";

import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardHeader } from "@/components/ui/card";
import { CopyButton } from "@/components/ui/copy-button";
import { EmptyState } from "@/components/ui/empty-state";
import { EnumSelect, Field, FormGrid, Input, Textarea } from "@/components/ui/form";
import { Modal } from "@/components/ui/modal";
import { SubmitButton } from "@/components/ui/submit-button";
import { useToast } from "@/components/ui/toast";
import { CREDENTIAL_KIND, CREDENTIAL_KIND_LIST } from "@/lib/constants";
import { createCredential, deleteCredential, revealSecret } from "@/lib/actions/credentials";

/**
 * Coffre à identifiants.
 *
 * Les secrets ne sont jamais rendus dans le HTML : la ligne n'affiche que le
 * libellé et l'identifiant. Cliquer sur « Révéler » déclenche un appel serveur
 * qui déchiffre à la volée, et la valeur est effacée de l'état au bout de
 * 30 secondes — pour éviter qu'un mot de passe reste affiché sur un écran
 * qu'on a quitté.
 */

export interface CredentialItem {
  id: string;
  label: string;
  kind: CredentialKind;
  username: string | null;
  url: string | null;
  notes: string | null;
  hasSecret: boolean;
}

export function CredentialVault({
  credentials,
  scope,
  revalidatePath,
}: {
  credentials: CredentialItem[];
  /** Rattachement du nouvel identifiant (client, projet, domaine ou hébergement). */
  scope: { clientId?: string; projectId?: string; domainId?: string; hostingId?: string };
  revalidatePath: string;
}) {
  const [open, setOpen] = useState(false);

  return (
    <Card>
      <CardHeader
        title="Identifiants"
        description="Chiffrés au repos (AES-256-GCM). Révélés uniquement à la demande."
        icon={<KeyRound className="h-4 w-4" />}
        action={
          <Button size="sm" onClick={() => setOpen(true)}>
            <Plus className="h-3.5 w-3.5" />
            Ajouter
          </Button>
        }
      />

      <div className="divide-y divide-line">
        {credentials.length > 0 ? (
          credentials.map((credential) => (
            <CredentialRow
              key={credential.id}
              credential={credential}
              revalidatePath={revalidatePath}
            />
          ))
        ) : (
          <EmptyState
            compact
            title="Coffre vide"
            description="Stockez ici les accès FTP, SSH, CMS ou panneau d'administration."
          />
        )}
      </div>

      <Modal
        open={open}
        onClose={() => setOpen(false)}
        title="Nouvel identifiant"
        description="Le secret est chiffré avant d'être écrit en base."
      >
        <form
          action={async (formData) => {
            await createCredential(formData);
            setOpen(false);
          }}
          className="space-y-4 p-5"
        >
          <input type="hidden" name="revalidate" value={revalidatePath} />
          {scope.clientId ? <input type="hidden" name="clientId" value={scope.clientId} /> : null}
          {scope.projectId ? <input type="hidden" name="projectId" value={scope.projectId} /> : null}
          {scope.domainId ? <input type="hidden" name="domainId" value={scope.domainId} /> : null}
          {scope.hostingId ? <input type="hidden" name="hostingId" value={scope.hostingId} /> : null}

          <FormGrid>
            <Field label="Libellé" htmlFor="cred-label" required>
              <Input id="cred-label" name="label" required data-autofocus placeholder="FTP production" />
            </Field>
            <Field label="Type" htmlFor="cred-kind">
              <EnumSelect id="cred-kind" name="kind" options={CREDENTIAL_KIND_LIST} defaultValue="FTP" />
            </Field>
            <Field label="Identifiant" htmlFor="cred-username">
              <Input id="cred-username" name="username" autoComplete="off" />
            </Field>
            <Field label="Mot de passe / clé" htmlFor="cred-secret">
              <Input id="cred-secret" name="secret" type="password" autoComplete="new-password" />
            </Field>
            <Field label="URL" htmlFor="cred-url" span={2}>
              <Input id="cred-url" name="url" placeholder="https://…" />
            </Field>
            <Field label="Remarques" htmlFor="cred-notes" span={2}>
              <Textarea id="cred-notes" name="notes" rows={2} />
            </Field>
          </FormGrid>

          <div className="flex justify-end gap-2 pt-1">
            <Button type="button" variant="ghost" onClick={() => setOpen(false)}>
              Annuler
            </Button>
            <SubmitButton>Enregistrer</SubmitButton>
          </div>
        </form>
      </Modal>
    </Card>
  );
}

function CredentialRow({
  credential,
  revalidatePath,
}: {
  credential: CredentialItem;
  revalidatePath: string;
}) {
  const [secret, setSecret] = useState<string | null>(null);
  const [loading, setLoading] = useState(false);
  const [pending, startTransition] = useTransition();
  const { notify } = useToast();

  async function toggleReveal() {
    if (secret) {
      setSecret(null);
      return;
    }

    setLoading(true);
    const result = await revealSecret(credential.id);
    setLoading(false);

    if (result.error || result.value === null) {
      notify({ title: "Impossible d'afficher", description: result.error, tone: "danger" });
      return;
    }

    setSecret(result.value);
    // Masquage automatique : un secret ne doit pas rester à l'écran.
    window.setTimeout(() => setSecret(null), 30_000);
  }

  return (
    <div className="flex items-start gap-3 px-4 py-3">
      <div className="min-w-0 flex-1">
        <div className="flex flex-wrap items-center gap-2">
          <span className="text-[13px] font-medium text-ink">{credential.label}</span>
          <Badge tone="muted" dot={false} square>
            {CREDENTIAL_KIND[credential.kind].label}
          </Badge>
        </div>

        {credential.username ? (
          <div className="mt-1 flex items-center gap-1 text-xs text-ink-secondary">
            <span className="font-mono">{credential.username}</span>
            <CopyButton value={credential.username} label="Identifiant" silent />
          </div>
        ) : null}

        {credential.hasSecret ? (
          <div className="mt-1 flex items-center gap-1">
            <code className="rounded bg-surface-inset px-1.5 py-0.5 font-mono text-xs text-ink">
              {secret ?? "••••••••••••"}
            </code>
            {secret ? <CopyButton value={secret} label="Mot de passe" silent /> : null}
          </div>
        ) : (
          <p className="mt-1 text-xs text-ink-muted">Aucun secret enregistré</p>
        )}

        {credential.url ? (
          <a
            href={credential.url}
            target="_blank"
            rel="noreferrer"
            className="mt-1 block truncate text-xs link-subtle"
          >
            {credential.url}
          </a>
        ) : null}

        {credential.notes ? (
          <p className="mt-1 text-xs text-ink-muted">{credential.notes}</p>
        ) : null}
      </div>

      <div className="flex shrink-0 items-center gap-1">
        {credential.hasSecret ? (
          <Button
            variant="ghost"
            size="icon"
            onClick={toggleReveal}
            aria-label={secret ? "Masquer" : "Révéler"}
            title={secret ? "Masquer" : "Révéler"}
          >
            {loading ? (
              <Loader2 className="h-3.5 w-3.5 animate-spin" />
            ) : secret ? (
              <EyeOff className="h-3.5 w-3.5" />
            ) : (
              <Eye className="h-3.5 w-3.5" />
            )}
          </Button>
        ) : null}

        <Button
          variant="ghost"
          size="icon"
          disabled={pending}
          aria-label="Supprimer"
          onClick={() => {
            if (!window.confirm(`Supprimer l'identifiant « ${credential.label} » ?`)) return;
            startTransition(async () => {
              await deleteCredential(credential.id, revalidatePath);
              notify({ title: "Identifiant supprimé", tone: "success" });
            });
          }}
        >
          <Trash2 className="h-3.5 w-3.5" />
        </Button>
      </div>
    </div>
  );
}
