import { AppShell } from "@/components/layout/app-shell";
import { CommandPaletteProvider } from "@/components/layout/command-palette";
import { requireUser } from "@/lib/auth";
import { db } from "@/lib/db";

/**
 * Toutes les pages de l'outil vivent dans ce groupe : l'authentification est
 * donc vérifiée à un seul endroit, impossible d'oublier une page.
 */
export default async function AppLayout({ children }: { children: React.ReactNode }) {
  const user = await requireUser();
  const now = new Date();

  // Pastilles du menu : ce qui exige une action, pas de simples totaux.
  const [unpaid, lateTasks] = await Promise.all([
    db.subscriptionPeriod.count({
      where: { status: { in: ["DUE", "OVERDUE"] }, dueAt: { lt: now } },
    }),
    db.task.count({
      where: { status: { not: "DONE" }, dueAt: { lt: now } },
    }),
  ]);

  return (
    <CommandPaletteProvider>
      <AppShell user={user} badges={{ "/paiements": unpaid, "/taches": lateTasks }}>
        {children}
      </AppShell>
    </CommandPaletteProvider>
  );
}
