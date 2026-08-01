/**
 * Calculs de devis.
 *
 * Isolé des server actions (un fichier « use server » ne peut exporter que des
 * fonctions asynchrones) et volontairement pur : les mêmes chiffres servent à
 * l'écran de saisie, à l'aperçu imprimable et aux totaux des listes, sans
 * risque de divergence.
 */

export interface QuoteLine {
  quantity: number;
  unitPrice: number;
}

export interface QuoteTotals {
  subtotal: number;
  discount: number;
  taxable: number;
  tax: number;
  total: number;
}

export function computeQuoteTotals(
  items: QuoteLine[],
  discountPct: number,
  taxRate: number,
): QuoteTotals {
  const subtotal = items.reduce((sum, item) => sum + item.quantity * item.unitPrice, 0);
  const discount = subtotal * (discountPct / 100);
  const taxable = subtotal - discount;
  const tax = taxable * (taxRate / 100);
  return { subtotal, discount, taxable, tax, total: taxable + tax };
}
