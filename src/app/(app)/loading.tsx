/**
 * Squelette de chargement.
 *
 * Reprend la silhouette réelle d'un écran (en-tête, tuiles, tableau) plutôt
 * qu'un spinner centré : la page ne « saute » pas au moment où les données
 * arrivent, ce qui donne une impression de rapidité même à durée égale.
 */
export default function Loading() {
  return (
    <div className="space-y-5" aria-busy="true" aria-label="Chargement">
      <div className="space-y-2">
        <div className="skeleton h-3 w-28" />
        <div className="skeleton h-6 w-64" />
      </div>

      <div className="grid gap-3 sm:grid-cols-2 xl:grid-cols-4">
        {Array.from({ length: 4 }).map((_, index) => (
          <div key={index} className="card p-4">
            <div className="skeleton h-3 w-24" />
            <div className="skeleton mt-3 h-7 w-32" />
            <div className="skeleton mt-2 h-3 w-20" />
          </div>
        ))}
      </div>

      <div className="card overflow-hidden">
        <div className="border-b border-line px-4 py-3">
          <div className="skeleton h-4 w-40" />
        </div>
        {Array.from({ length: 6 }).map((_, index) => (
          <div key={index} className="flex items-center gap-4 border-b border-line px-4 py-3 last:border-b-0">
            <div className="skeleton h-7 w-7 rounded-md" />
            <div className="flex-1 space-y-1.5">
              <div className="skeleton h-3.5 w-1/3" />
              <div className="skeleton h-3 w-1/4" />
            </div>
            <div className="skeleton h-5 w-20 rounded-full" />
            <div className="skeleton h-3.5 w-16" />
          </div>
        ))}
      </div>
    </div>
  );
}
