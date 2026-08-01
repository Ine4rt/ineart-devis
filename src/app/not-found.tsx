import Link from "next/link";

export default function NotFound() {
  return (
    <div className="flex min-h-dvh items-center justify-center bg-canvas px-4">
      <div className="max-w-sm animate-rise-in text-center">
        <p className="font-mono text-5xl font-semibold text-ink-muted">404</p>
        <h1 className="mt-3 text-base font-semibold text-ink">Page introuvable</h1>
        <p className="mt-1.5 text-sm leading-relaxed text-ink-muted">
          Cette fiche a peut-être été supprimée, ou le lien est incorrect.
        </p>
        <Link href="/" className="btn btn-primary mt-5 h-8 px-3">
          Retour au tableau de bord
        </Link>
      </div>
    </div>
  );
}
