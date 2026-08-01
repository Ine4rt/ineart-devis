import { NextResponse } from "next/server";

import { getSessionUser } from "@/lib/auth";
import { globalSearch } from "@/lib/queries/search";

/**
 * Alimente la palette de commandes (⌘K). Route plutôt que server action :
 * on veut pouvoir annuler la requête précédente à chaque frappe (AbortController).
 */
export async function GET(request: Request) {
  const user = await getSessionUser();
  if (!user) return NextResponse.json({ results: [] }, { status: 401 });

  const query = new URL(request.url).searchParams.get("q") ?? "";
  const results = await globalSearch(query);

  return NextResponse.json({ results });
}
