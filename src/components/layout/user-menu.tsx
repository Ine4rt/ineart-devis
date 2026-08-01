"use client";

import { LogOut, Settings, User as UserIcon } from "lucide-react";
import {
  Dropdown,
  DropdownItem,
  DropdownLink,
  DropdownSeparator,
} from "@/components/ui/dropdown";
import type { SessionUser } from "@/lib/auth";
import { initials } from "@/lib/utils";

export function UserMenu({ user }: { user: SessionUser }) {
  return (
    <Dropdown
      trigger={
        <button
          type="button"
          className="flex h-8 w-8 items-center justify-center rounded-md border border-line bg-surface text-[11px] font-semibold text-ink-secondary transition-colors hover:border-line-strong hover:text-ink"
          aria-label="Menu du compte"
        >
          {initials(user.name)}
        </button>
      }
    >
      <div className="border-b border-line px-2.5 pb-2 pt-1.5">
        <p className="truncate text-xs font-semibold text-ink">{user.name}</p>
        <p className="truncate text-[11px] text-ink-muted">{user.email}</p>
      </div>

      <div className="pt-1">
        <DropdownLink href="/reglages" icon={<UserIcon className="h-3.5 w-3.5" />}>
          Mon compte
        </DropdownLink>
        <DropdownLink href="/reglages/donnees" icon={<Settings className="h-3.5 w-3.5" />}>
          Données &amp; export
        </DropdownLink>
      </div>

      <DropdownSeparator />

      <form action="/api/logout" method="post">
        <DropdownItem danger type="submit" icon={<LogOut className="h-3.5 w-3.5" />}>
          Se déconnecter
        </DropdownItem>
      </form>
    </Dropdown>
  );
}
