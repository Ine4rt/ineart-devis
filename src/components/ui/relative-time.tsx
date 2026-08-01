"use client";

import { useEffect, useState } from "react";

import { formatDate, formatRelative } from "@/lib/format";

/**
 * Horodatage relatif (« il y a 7 minutes ») utilisable dans un composant client.
 *
 * Une valeur relative dépend de l'instant du rendu : calculée sur le serveur
 * puis réhydratée quelques centaines de millisecondes plus tard, elle produit
 * deux textes différents et donc une erreur d'hydratation. On rend donc une
 * date absolue — stable — au premier rendu, et on bascule sur le libellé
 * relatif une fois monté côté client.
 */
export function RelativeTime({
  value,
  className,
}: {
  value: Date | string;
  className?: string;
}) {
  const [mounted, setMounted] = useState(false);
  useEffect(() => setMounted(true), []);

  return (
    <span className={className} title={formatDate(value, "long")}>
      {mounted ? formatRelative(value) : formatDate(value, "short")}
    </span>
  );
}
