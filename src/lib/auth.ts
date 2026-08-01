import "server-only";

import { cookies } from "next/headers";
import { redirect } from "next/navigation";
import { SignJWT, jwtVerify } from "jose";

import { db } from "@/lib/db";

/**
 * Authentification mono-utilisateur.
 *
 * La session est un JWT signé (HS256) stocké dans un cookie httpOnly. Pas de
 * table de sessions : l'outil n'a qu'un utilisateur et la révocation se fait en
 * changeant APP_SESSION_SECRET, ce qui invalide tous les jetons émis.
 */

const COOKIE_NAME = "ineart_session";
const SESSION_DURATION_DAYS = 30;

function sessionSecret(): Uint8Array {
  const secret = process.env.APP_SESSION_SECRET;
  if (!secret || secret.length < 16) {
    throw new Error(
      "APP_SESSION_SECRET manquante ou trop courte. Générez-en une avec : openssl rand -base64 32",
    );
  }
  return new TextEncoder().encode(secret);
}

export interface SessionUser {
  id: string;
  email: string;
  name: string;
  avatarUrl: string | null;
}

export async function createSession(userId: string): Promise<void> {
  const token = await new SignJWT({ sub: userId })
    .setProtectedHeader({ alg: "HS256" })
    .setIssuedAt()
    .setExpirationTime(`${SESSION_DURATION_DAYS}d`)
    .sign(sessionSecret());

  const store = await cookies();
  store.set(COOKIE_NAME, token, {
    httpOnly: true,
    sameSite: "lax",
    secure: process.env.NODE_ENV === "production",
    path: "/",
    maxAge: SESSION_DURATION_DAYS * 24 * 60 * 60,
  });
}

export async function destroySession(): Promise<void> {
  const store = await cookies();
  store.delete(COOKIE_NAME);
}

/** Utilisateur courant, ou `null` si le jeton est absent, expiré ou invalide. */
export async function getSessionUser(): Promise<SessionUser | null> {
  const store = await cookies();
  const token = store.get(COOKIE_NAME)?.value;
  if (!token) return null;

  try {
    const { payload } = await jwtVerify(token, sessionSecret());
    if (!payload.sub) return null;

    const user = await db.user.findUnique({
      where: { id: payload.sub },
      select: { id: true, email: true, name: true, avatarUrl: true },
    });
    return user;
  } catch {
    return null;
  }
}

/** Garde-fou des pages et server actions : redirige vers /login si non connecté. */
export async function requireUser(): Promise<SessionUser> {
  const user = await getSessionUser();
  if (!user) redirect("/login");
  return user;
}

/** Indique si un compte existe déjà — pilote l'écran d'installation initiale. */
export async function hasAccount(): Promise<boolean> {
  return (await db.user.count()) > 0;
}
