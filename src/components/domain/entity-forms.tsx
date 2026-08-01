import type { Domain, Hosting, Project, Subscription } from "@prisma/client";
import { Banknote, CalendarClock, Globe, Server, Settings2, ShieldCheck, Wrench } from "lucide-react";
import Link from "next/link";

import { Card } from "@/components/ui/card";
import {
  Checkbox,
  DataList,
  EnumSelect,
  Field,
  FormActions,
  FormGrid,
  FormSection,
  Input,
  Select,
  Textarea,
} from "@/components/ui/form";
import { SubmitButton } from "@/components/ui/submit-button";
import {
  BILLING_CYCLE_LIST,
  COMMON_HOSTS,
  COMMON_REGISTRARS,
  COMMON_TECHNOLOGIES,
  PROJECT_STATUS_LIST,
  PROJECT_TYPE_LIST,
  SUBSCRIPTION_STATUS_LIST,
  SUBSCRIPTION_TYPE_LIST,
} from "@/lib/constants";
import { toDateInputValue, toNumber } from "@/lib/format";
import { parseJsonArray } from "@/lib/utils";

/**
 * Formulaires des entités techniques et financières.
 *
 * Regroupés dans un seul module car ils partagent la même anatomie (sections,
 * grille, barre d'actions) : les faire cohabiter rend les incohérences
 * visibles immédiatement.
 */

export interface ClientOption {
  id: string;
  company: string;
}

export interface ProjectOption {
  id: string;
  name: string;
  clientId: string;
}

function ClientField({
  clients,
  defaultValue,
  required,
  hint,
}: {
  clients: ClientOption[];
  defaultValue?: string | null;
  required?: boolean;
  hint?: string;
}) {
  return (
    <Field label="Client" htmlFor="clientId" required={required} hint={hint}>
      <Select id="clientId" name="clientId" defaultValue={defaultValue ?? ""} required={required}>
        <option value="">{required ? "Sélectionner un client" : "Aucun client"}</option>
        {clients.map((client) => (
          <option key={client.id} value={client.id}>
            {client.company}
          </option>
        ))}
      </Select>
    </Field>
  );
}

function ProjectField({
  projects,
  defaultValue,
}: {
  projects: ProjectOption[];
  defaultValue?: string | null;
}) {
  return (
    <Field label="Projet associé" htmlFor="projectId" hint="Facultatif.">
      <Select id="projectId" name="projectId" defaultValue={defaultValue ?? ""}>
        <option value="">Aucun projet</option>
        {projects.map((project) => (
          <option key={project.id} value={project.id}>
            {project.name}
          </option>
        ))}
      </Select>
    </Field>
  );
}

function Actions({ cancelHref, label }: { cancelHref: string; label: string }) {
  return (
    <FormActions>
      <Link href={cancelHref} className="btn btn-ghost h-8 px-3">
        Annuler
      </Link>
      <SubmitButton pendingLabel="Enregistrement…">{label}</SubmitButton>
    </FormActions>
  );
}

// ---------------------------------------------------------------------------
// Projet
// ---------------------------------------------------------------------------

export function ProjectForm({
  project,
  clients,
  action,
  submitLabel,
  cancelHref,
  defaultClientId,
}: {
  project?: Project | null;
  clients: ClientOption[];
  action: (formData: FormData) => Promise<void>;
  submitLabel: string;
  cancelHref: string;
  defaultClientId?: string;
}) {
  return (
    <form action={action}>
      <Card className="overflow-hidden">
        <FormSection
          title="Le site"
          description="Ce que vous construisez, pour qui, et où il vit."
          icon={<Globe className="h-4 w-4" />}
        >
          <FormGrid>
            <Field label="Nom du projet" htmlFor="name" required span={2}>
              <Input
                id="name"
                name="name"
                required
                data-autofocus
                defaultValue={project?.name ?? ""}
                placeholder="Site vitrine — Boulangerie Martin"
              />
            </Field>
            <ClientField
              clients={clients}
              defaultValue={project?.clientId ?? defaultClientId}
              required
            />
            <Field label="Type" htmlFor="type">
              <EnumSelect
                id="type"
                name="type"
                options={PROJECT_TYPE_LIST}
                defaultValue={project?.type ?? "VITRINE"}
              />
            </Field>
            <Field label="URL de production" htmlFor="url">
              <Input id="url" name="url" defaultValue={project?.url ?? ""} placeholder="https://…" />
            </Field>
            <Field label="URL de préproduction" htmlFor="stagingUrl">
              <Input id="stagingUrl" name="stagingUrl" defaultValue={project?.stagingUrl ?? ""} />
            </Field>
            <Field label="Description" htmlFor="description" span={2}>
              <Textarea id="description" name="description" rows={3} defaultValue={project?.description ?? ""} />
            </Field>
          </FormGrid>
        </FormSection>

        <FormSection
          title="Technique"
          description="Sur mesure ou CMS, et la pile utilisée."
          icon={<Wrench className="h-4 w-4" />}
        >
          <FormGrid>
            <Field span={2}>
              <Checkbox
                name="isCustom"
                label="Développement sur mesure"
                description="Décochez si le site repose sur un CMS existant."
                defaultChecked={project?.isCustom ?? true}
              />
            </Field>
            <Field label="CMS" htmlFor="cms" hint="Uniquement si le site n'est pas sur mesure.">
              <Input
                id="cms"
                name="cms"
                list="cms-list"
                defaultValue={project?.cms ?? ""}
                placeholder="WordPress, Shopify…"
              />
              <DataList id="cms-list" options={["WordPress", "Shopify", "Webflow", "Prestashop", "Drupal"]} />
            </Field>
            <Field label="Dépôt Git" htmlFor="repositoryUrl">
              <Input id="repositoryUrl" name="repositoryUrl" defaultValue={project?.repositoryUrl ?? ""} />
            </Field>
            <Field
              label="Technologies"
              htmlFor="techStack"
              span={2}
              hint="Séparées par des virgules."
            >
              <Input
                id="techStack"
                name="techStack"
                list="tech-list"
                defaultValue={parseJsonArray(project?.techStack).join(", ")}
                placeholder="Next.js, Tailwind CSS, Prisma"
              />
              <DataList id="tech-list" options={COMMON_TECHNOLOGIES} />
            </Field>
          </FormGrid>
        </FormSection>

        <FormSection
          title="Avancement"
          description="Le suivi est manuel et assumé : c'est un tableau de bord, pas un outil de tickets."
          icon={<Settings2 className="h-4 w-4" />}
        >
          <FormGrid>
            <Field label="Statut" htmlFor="status">
              <EnumSelect
                id="status"
                name="status"
                options={PROJECT_STATUS_LIST}
                defaultValue={project?.status ?? "DISCOVERY"}
              />
            </Field>
            <Field label="Progression (%)" htmlFor="progress">
              <Input
                id="progress"
                name="progress"
                type="number"
                min={0}
                max={100}
                step={5}
                defaultValue={project?.progress ?? 0}
              />
            </Field>
            <Field label="Démarré le" htmlFor="startedAt">
              <Input
                id="startedAt"
                name="startedAt"
                type="date"
                defaultValue={toDateInputValue(project?.startedAt)}
              />
            </Field>
            <Field label="Mis en ligne le" htmlFor="launchedAt">
              <Input
                id="launchedAt"
                name="launchedAt"
                type="date"
                defaultValue={toDateInputValue(project?.launchedAt)}
              />
            </Field>
            <Field label="Remarques" htmlFor="notes" span={2}>
              <Textarea id="notes" name="notes" rows={4} defaultValue={project?.notes ?? ""} />
            </Field>
          </FormGrid>
        </FormSection>

        <Actions cancelHref={cancelHref} label={submitLabel} />
      </Card>
    </form>
  );
}

// ---------------------------------------------------------------------------
// Domaine
// ---------------------------------------------------------------------------

export function DomainForm({
  domain,
  clients,
  projects,
  action,
  submitLabel,
  cancelHref,
  defaultClientId,
}: {
  domain?: Domain | null;
  clients: ClientOption[];
  projects: ProjectOption[];
  action: (formData: FormData) => Promise<void>;
  submitLabel: string;
  cancelHref: string;
  defaultClientId?: string;
}) {
  return (
    <form action={action}>
      <Card className="overflow-hidden">
        <FormSection
          title="Nom de domaine"
          description="Le domaine et son rattachement."
          icon={<Globe className="h-4 w-4" />}
        >
          <FormGrid>
            <Field label="Domaine" htmlFor="name" required span={2}>
              <Input
                id="name"
                name="name"
                required
                data-autofocus
                defaultValue={domain?.name ?? ""}
                placeholder="boulangerie-martin.be"
              />
            </Field>
            <ClientField clients={clients} defaultValue={domain?.clientId ?? defaultClientId} />
            <ProjectField projects={projects} defaultValue={domain?.projectId} />
          </FormGrid>
        </FormSection>

        <FormSection
          title="Registrar & échéance"
          description="Le prix saisi est celui que VOUS payez : il alimente le calcul de marge."
          icon={<CalendarClock className="h-4 w-4" />}
        >
          <FormGrid>
            <Field label="Registrar" htmlFor="registrar">
              <Input
                id="registrar"
                name="registrar"
                list="registrars"
                defaultValue={domain?.registrar ?? ""}
              />
              <DataList id="registrars" options={COMMON_REGISTRARS} />
            </Field>
            <Field label="Prix de renouvellement (€/an)" htmlFor="renewalPrice">
              <Input
                id="renewalPrice"
                name="renewalPrice"
                inputMode="decimal"
                defaultValue={domain ? toNumber(domain.renewalPrice) : ""}
                placeholder="12"
              />
            </Field>
            <Field label="Acheté le" htmlFor="purchasedAt">
              <Input
                id="purchasedAt"
                name="purchasedAt"
                type="date"
                defaultValue={toDateInputValue(domain?.purchasedAt)}
              />
            </Field>
            <Field label="Expire le" htmlFor="expiresAt" hint="Déclenche les rappels du tableau de bord.">
              <Input
                id="expiresAt"
                name="expiresAt"
                type="date"
                defaultValue={toDateInputValue(domain?.expiresAt)}
              />
            </Field>
            <Field span={2}>
              <Checkbox
                name="autoRenew"
                label="Reconduction automatique activée"
                description="Si décoché, le domaine est signalé comme nécessitant une action manuelle."
                defaultChecked={domain?.autoRenew ?? true}
              />
            </Field>
          </FormGrid>
        </FormSection>

        <FormSection
          title="DNS & certificat"
          description="Le SSL expire souvent avant le domaine : il est suivi séparément."
          icon={<ShieldCheck className="h-4 w-4" />}
        >
          <FormGrid>
            <Field label="Fournisseur DNS" htmlFor="dnsProvider">
              <Input id="dnsProvider" name="dnsProvider" defaultValue={domain?.dnsProvider ?? ""} />
            </Field>
            <Field label="Émetteur SSL" htmlFor="sslProvider">
              <Input
                id="sslProvider"
                name="sslProvider"
                defaultValue={domain?.sslProvider ?? ""}
                placeholder="Let's Encrypt"
              />
            </Field>
            <Field label="Serveurs de noms" htmlFor="nameservers" span={2} hint="Un par ligne.">
              <Textarea
                id="nameservers"
                name="nameservers"
                rows={3}
                defaultValue={domain?.nameservers ?? ""}
                placeholder={"ns1.exemple.com\nns2.exemple.com"}
              />
            </Field>
            <Field label="SSL expire le" htmlFor="sslExpiresAt">
              <Input
                id="sslExpiresAt"
                name="sslExpiresAt"
                type="date"
                defaultValue={toDateInputValue(domain?.sslExpiresAt)}
              />
            </Field>
            <Field>
              <Checkbox
                name="sslAutoRenew"
                label="Renouvellement SSL automatique"
                defaultChecked={domain?.sslAutoRenew ?? true}
              />
            </Field>
            <Field label="Remarques" htmlFor="notes" span={2}>
              <Textarea id="notes" name="notes" rows={3} defaultValue={domain?.notes ?? ""} />
            </Field>
          </FormGrid>
        </FormSection>

        <Actions cancelHref={cancelHref} label={submitLabel} />
      </Card>
    </form>
  );
}

// ---------------------------------------------------------------------------
// Hébergement
// ---------------------------------------------------------------------------

export function HostingForm({
  hosting,
  clients,
  projects,
  action,
  submitLabel,
  cancelHref,
  defaultClientId,
}: {
  hosting?: Hosting | null;
  clients: ClientOption[];
  projects: ProjectOption[];
  action: (formData: FormData) => Promise<void>;
  submitLabel: string;
  cancelHref: string;
  defaultClientId?: string;
}) {
  return (
    <form action={action}>
      <Card className="overflow-hidden">
        <FormSection
          title="Offre"
          description="Le fournisseur, la formule et son coût réel."
          icon={<Server className="h-4 w-4" />}
        >
          <FormGrid>
            <Field label="Fournisseur" htmlFor="provider" required>
              <Input
                id="provider"
                name="provider"
                required
                data-autofocus
                list="hosts"
                defaultValue={hosting?.provider ?? ""}
              />
              <DataList id="hosts" options={COMMON_HOSTS} />
            </Field>
            <Field label="Offre / formule" htmlFor="plan">
              <Input id="plan" name="plan" defaultValue={hosting?.plan ?? ""} placeholder="Perso, Pro…" />
            </Field>
            <ClientField clients={clients} defaultValue={hosting?.clientId ?? defaultClientId} />
            <ProjectField projects={projects} defaultValue={hosting?.projectId} />
            <Field label="Prix" htmlFor="price">
              <Input
                id="price"
                name="price"
                inputMode="decimal"
                defaultValue={hosting ? toNumber(hosting.price) : ""}
              />
            </Field>
            <Field label="Cycle de facturation" htmlFor="billingCycle">
              <EnumSelect
                id="billingCycle"
                name="billingCycle"
                options={BILLING_CYCLE_LIST}
                defaultValue={hosting?.billingCycle ?? "YEARLY"}
              />
            </Field>
            <Field label="Échéance" htmlFor="renewsAt">
              <Input
                id="renewsAt"
                name="renewsAt"
                type="date"
                defaultValue={toDateInputValue(hosting?.renewsAt)}
              />
            </Field>
            <Field>
              <Checkbox
                name="autoRenew"
                label="Reconduction automatique"
                defaultChecked={hosting?.autoRenew ?? true}
              />
            </Field>
          </FormGrid>
        </FormSection>

        <FormSection
          title="Caractéristiques techniques"
          description="Ce qu'il faut retrouver vite un jour d'incident."
          icon={<Wrench className="h-4 w-4" />}
        >
          <FormGrid>
            <Field label="Espace disque (Go)" htmlFor="diskSpaceGb">
              <Input
                id="diskSpaceGb"
                name="diskSpaceGb"
                type="number"
                min={0}
                defaultValue={hosting?.diskSpaceGb ?? ""}
              />
            </Field>
            <Field label="Bande passante" htmlFor="bandwidth">
              <Input id="bandwidth" name="bandwidth" defaultValue={hosting?.bandwidth ?? ""} />
            </Field>
            <Field label="Adresse IP" htmlFor="serverIp">
              <Input id="serverIp" name="serverIp" defaultValue={hosting?.serverIp ?? ""} />
            </Field>
            <Field label="Région / datacenter" htmlFor="region">
              <Input id="region" name="region" defaultValue={hosting?.region ?? ""} />
            </Field>
            <Field label="Version PHP / runtime" htmlFor="phpVersion">
              <Input id="phpVersion" name="phpVersion" defaultValue={hosting?.phpVersion ?? ""} />
            </Field>
            <Field label="Panneau d'administration" htmlFor="controlPanelUrl">
              <Input
                id="controlPanelUrl"
                name="controlPanelUrl"
                defaultValue={hosting?.controlPanelUrl ?? ""}
              />
            </Field>
            <Field label="Notes techniques" htmlFor="technicalNotes" span={2}>
              <Textarea
                id="technicalNotes"
                name="technicalNotes"
                rows={4}
                defaultValue={hosting?.technicalNotes ?? ""}
                placeholder="Cron, certificats, particularités de déploiement…"
              />
            </Field>
            <Field label="Remarques" htmlFor="notes" span={2}>
              <Textarea id="notes" name="notes" rows={3} defaultValue={hosting?.notes ?? ""} />
            </Field>
          </FormGrid>
        </FormSection>

        <Actions cancelHref={cancelHref} label={submitLabel} />
      </Card>
    </form>
  );
}

// ---------------------------------------------------------------------------
// Abonnement
// ---------------------------------------------------------------------------

export function SubscriptionForm({
  subscription,
  clients,
  projects,
  action,
  submitLabel,
  cancelHref,
  defaultClientId,
}: {
  subscription?: Subscription | null;
  clients: ClientOption[];
  projects: ProjectOption[];
  action: (formData: FormData) => Promise<void>;
  submitLabel: string;
  cancelHref: string;
  defaultClientId?: string;
}) {
  const isEdit = Boolean(subscription);

  return (
    <form action={action}>
      <Card className="overflow-hidden">
        <FormSection
          title="Contrat"
          description="Ce que le client paie, et à quelle fréquence."
          icon={<Banknote className="h-4 w-4" />}
        >
          <FormGrid>
            <Field label="Intitulé" htmlFor="name" required span={2}>
              <Input
                id="name"
                name="name"
                required
                data-autofocus
                defaultValue={subscription?.name ?? ""}
                placeholder="Formule complète — site + domaine + maintenance"
              />
            </Field>
            <ClientField
              clients={clients}
              defaultValue={subscription?.clientId ?? defaultClientId}
              required
            />
            <ProjectField projects={projects} defaultValue={subscription?.projectId} />
            <Field label="Type" htmlFor="type">
              <EnumSelect
                id="type"
                name="type"
                options={SUBSCRIPTION_TYPE_LIST}
                defaultValue={subscription?.type ?? "FULL_CARE"}
              />
            </Field>
            <Field label="Statut" htmlFor="status">
              <EnumSelect
                id="status"
                name="status"
                options={SUBSCRIPTION_STATUS_LIST}
                defaultValue={subscription?.status ?? "ACTIVE"}
              />
            </Field>
            <Field label="Montant" htmlFor="amount" required>
              <Input
                id="amount"
                name="amount"
                inputMode="decimal"
                required
                defaultValue={subscription ? toNumber(subscription.amount) : ""}
                placeholder="480"
              />
            </Field>
            <Field label="Cycle" htmlFor="billingCycle">
              <EnumSelect
                id="billingCycle"
                name="billingCycle"
                options={BILLING_CYCLE_LIST}
                defaultValue={subscription?.billingCycle ?? "YEARLY"}
              />
            </Field>
            <Field label="Devise" htmlFor="currency">
              <Select id="currency" name="currency" defaultValue={subscription?.currency ?? "EUR"}>
                <option value="EUR">EUR (€)</option>
                <option value="CHF">CHF</option>
                <option value="USD">USD ($)</option>
              </Select>
            </Field>
            <Field label="Moyen de paiement" htmlFor="paymentMethod">
              <Select
                id="paymentMethod"
                name="paymentMethod"
                defaultValue={subscription?.paymentMethod ?? ""}
              >
                <option value="">Non précisé</option>
                <option value="Virement">Virement</option>
                <option value="Domiciliation">Domiciliation</option>
                <option value="Carte">Carte</option>
                <option value="Espèces">Espèces</option>
              </Select>
            </Field>
          </FormGrid>
        </FormSection>

        <FormSection
          title="Périodes"
          description={
            isEdit
              ? "Ajustez la période couverte et la date du prochain renouvellement."
              : "La première échéance est créée automatiquement à partir de ces dates."
          }
          icon={<CalendarClock className="h-4 w-4" />}
        >
          <FormGrid>
            <Field label="Date de souscription" htmlFor="startedAt" required>
              <Input
                id="startedAt"
                name="startedAt"
                type="date"
                required
                defaultValue={toDateInputValue(subscription?.startedAt ?? new Date())}
              />
            </Field>

            {isEdit ? (
              <>
                <Field label="Période en cours — début" htmlFor="currentPeriodStart">
                  <Input
                    id="currentPeriodStart"
                    name="currentPeriodStart"
                    type="date"
                    defaultValue={toDateInputValue(subscription?.currentPeriodStart)}
                  />
                </Field>
                <Field label="Période en cours — fin" htmlFor="currentPeriodEnd">
                  <Input
                    id="currentPeriodEnd"
                    name="currentPeriodEnd"
                    type="date"
                    defaultValue={toDateInputValue(subscription?.currentPeriodEnd)}
                  />
                </Field>
                <Field label="Prochain renouvellement" htmlFor="nextRenewalAt">
                  <Input
                    id="nextRenewalAt"
                    name="nextRenewalAt"
                    type="date"
                    defaultValue={toDateInputValue(subscription?.nextRenewalAt)}
                  />
                </Field>
              </>
            ) : (
              <>
                <Field
                  label="Échéance de la première facture"
                  htmlFor="firstDueAt"
                  hint="Par défaut : la date de souscription."
                >
                  <Input id="firstDueAt" name="firstDueAt" type="date" />
                </Field>
                <Field span={2}>
                  <Checkbox
                    name="firstPeriodPaid"
                    label="La première période est déjà réglée"
                    description="Cochez si le client a payé au moment de la signature."
                  />
                </Field>
              </>
            )}

            <Field span={2}>
              <Checkbox
                name="autoRenew"
                label="Reconduction tacite"
                description="Le contrat se poursuit automatiquement à chaque échéance."
                defaultChecked={subscription?.autoRenew ?? true}
              />
            </Field>
            <Field label="Remarques" htmlFor="notes" span={2}>
              <Textarea id="notes" name="notes" rows={3} defaultValue={subscription?.notes ?? ""} />
            </Field>
          </FormGrid>
        </FormSection>

        <Actions cancelHref={cancelHref} label={submitLabel} />
      </Card>
    </form>
  );
}
