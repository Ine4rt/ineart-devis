"use client";

import { Printer } from "lucide-react";

import { Button } from "@/components/ui/button";

/**
 * Déclenche l'impression du navigateur. Combiné aux règles `@media print`,
 * « Enregistrer au format PDF » produit un devis propre — sans embarquer de
 * moteur de génération PDF côté serveur.
 */
export function PrintButton({ label = "Imprimer / PDF" }: { label?: string }) {
  return (
    <Button variant="primary" size="sm" onClick={() => window.print()}>
      <Printer className="h-3.5 w-3.5" />
      {label}
    </Button>
  );
}
