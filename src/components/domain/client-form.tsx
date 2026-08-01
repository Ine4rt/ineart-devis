import type { Client } from "@prisma/client";
import { Building2, Contact, MapPin, StickyNote, Target } from "lucide-react";
import Link from "next/link";

import { Card } from "@/components/ui/card";
import {
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
import { CLIENT_STATUS_LIST, PIPELINE_STAGE_LIST } from "@/lib/constants";
import { toNumber } from "@/lib/format";

/**
 * Formulaire client, partagé entre création et modification.
 *
 * Un seul composant pour les deux cas : impossible qu'un champ ajouté à la
 * création soit oublié à la modification.
 */
export function ClientForm({
  client,
  action,
  submitLabel,
}: {
  client?: Client | null;
  action: (formData: FormData) => Promise<void>;
  submitLabel: string;
}) {
  return (
    <form action={action}>
      <Card className="overflow-hidden">
        <FormSection
          title="Identité"
          description="La société est le seul champ obligatoire — vous compléterez le reste plus tard."
          icon={<Building2 className="h-4 w-4" />}
        >
          <FormGrid>
            <Field label="Société" htmlFor="company" required span={2}>
              <Input
                id="company"
                name="company"
                required
                defaultValue={client?.company ?? ""}
                placeholder="Boulangerie Martin SRL"
                data-autofocus
              />
            </Field>
            <Field label="Prénom" htmlFor="firstName">
              <Input id="firstName" name="firstName" defaultValue={client?.firstName ?? ""} />
            </Field>
            <Field label="Nom" htmlFor="lastName">
              <Input id="lastName" name="lastName" defaultValue={client?.lastName ?? ""} />
            </Field>
            <Field label="Secteur d'activité" htmlFor="industry">
              <Input
                id="industry"
                name="industry"
                list="industries"
                defaultValue={client?.industry ?? ""}
                placeholder="Restauration, artisanat…"
              />
              <DataList
                id="industries"
                options={[
                  "Restauration",
                  "Artisanat",
                  "Commerce de détail",
                  "Construction",
                  "Santé",
                  "Immobilier",
                  "Services aux entreprises",
                  "Beauté & bien-être",
                ]}
              />
            </Field>
            <Field label="Logo (URL)" htmlFor="logoUrl" hint="Affiché dans les listes et la fiche.">
              <Input
                id="logoUrl"
                name="logoUrl"
                type="url"
                defaultValue={client?.logoUrl ?? ""}
                placeholder="https://…/logo.png"
              />
            </Field>
          </FormGrid>
        </FormSection>

        <FormSection
          title="Contact"
          description="Coordonnées principales de l'interlocuteur."
          icon={<Contact className="h-4 w-4" />}
        >
          <FormGrid>
            <Field label="E-mail" htmlFor="email">
              <Input id="email" name="email" type="email" defaultValue={client?.email ?? ""} />
            </Field>
            <Field label="Téléphone" htmlFor="phone">
              <Input id="phone" name="phone" type="tel" defaultValue={client?.phone ?? ""} />
            </Field>
            <Field label="Mobile" htmlFor="mobile">
              <Input id="mobile" name="mobile" type="tel" defaultValue={client?.mobile ?? ""} />
            </Field>
            <Field label="Site web actuel" htmlFor="website" hint="Laissez vide s'il n'en a pas.">
              <Input id="website" name="website" defaultValue={client?.website ?? ""} />
            </Field>
          </FormGrid>
        </FormSection>

        <FormSection
          title="Adresse & facturation"
          description="Utilisé sur les devis et les documents contractuels."
          icon={<MapPin className="h-4 w-4" />}
        >
          <FormGrid>
            <Field label="Adresse" htmlFor="addressLine1" span={2}>
              <Input id="addressLine1" name="addressLine1" defaultValue={client?.addressLine1 ?? ""} />
            </Field>
            <Field label="Complément" htmlFor="addressLine2" span={2}>
              <Input id="addressLine2" name="addressLine2" defaultValue={client?.addressLine2 ?? ""} />
            </Field>
            <Field label="Code postal" htmlFor="postalCode">
              <Input id="postalCode" name="postalCode" defaultValue={client?.postalCode ?? ""} />
            </Field>
            <Field label="Ville" htmlFor="city">
              <Input id="city" name="city" defaultValue={client?.city ?? ""} />
            </Field>
            <Field label="Pays" htmlFor="country">
              <Input id="country" name="country" defaultValue={client?.country ?? "Belgique"} />
            </Field>
            <Field label="Numéro de TVA" htmlFor="vatNumber">
              <Input
                id="vatNumber"
                name="vatNumber"
                defaultValue={client?.vatNumber ?? ""}
                placeholder="BE0123.456.789"
              />
            </Field>
            <Field label="Numéro d'entreprise" htmlFor="companyNumber">
              <Input
                id="companyNumber"
                name="companyNumber"
                defaultValue={client?.companyNumber ?? ""}
              />
            </Field>
          </FormGrid>
        </FormSection>

        <FormSection
          title="Suivi commercial"
          description="Un prospect suit le pipeline ; le passer en « Actif » le convertit en client."
          icon={<Target className="h-4 w-4" />}
        >
          <FormGrid>
            <Field label="Statut" htmlFor="status">
              <EnumSelect
                id="status"
                name="status"
                options={CLIENT_STATUS_LIST}
                defaultValue={client?.status ?? "PROSPECT"}
              />
            </Field>
            <Field
              label="Étape du pipeline"
              htmlFor="pipelineStage"
              hint="Ignoré si le statut n'est pas « Prospect »."
            >
              <EnumSelect
                id="pipelineStage"
                name="pipelineStage"
                options={PIPELINE_STAGE_LIST}
                defaultValue={client?.pipelineStage ?? "IDENTIFIED"}
              />
            </Field>
            <Field
              label="Potentiel annuel (€)"
              htmlFor="potentialValue"
              hint="Estimation, utilisée pour pondérer le pipeline."
            >
              <Input
                id="potentialValue"
                name="potentialValue"
                inputMode="decimal"
                defaultValue={client?.potentialValue ? toNumber(client.potentialValue) : ""}
                placeholder="480"
              />
            </Field>
            <Field label="Origine" htmlFor="source">
              <Select id="source" name="source" defaultValue={client?.source ?? ""}>
                <option value="">Non précisée</option>
                <option value="Prospection à froid">Prospection à froid</option>
                <option value="Recommandation">Recommandation</option>
                <option value="Réseaux sociaux">Réseaux sociaux</option>
                <option value="Bouche-à-oreille">Bouche-à-oreille</option>
                <option value="Site web">Site web</option>
                <option value="Événement">Événement</option>
                <option value="Autre">Autre</option>
              </Select>
            </Field>
          </FormGrid>
        </FormSection>

        <FormSection
          title="Notes privées"
          description="Visibles uniquement par vous. Jamais reprises sur un devis."
          icon={<StickyNote className="h-4 w-4" />}
        >
          <Field htmlFor="notes">
            <Textarea
              id="notes"
              name="notes"
              rows={5}
              defaultValue={client?.notes ?? ""}
              placeholder="Contexte, historique de la relation, préférences, points d'attention…"
            />
          </Field>
        </FormSection>

        <FormActions>
          <Link
            href={client ? `/clients/${client.id}` : "/clients"}
            className="btn btn-ghost h-8 px-3"
          >
            Annuler
          </Link>
          <SubmitButton pendingLabel="Enregistrement…">{submitLabel}</SubmitButton>
        </FormActions>
      </Card>
    </form>
  );
}
