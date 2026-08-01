"use server";

import { redirect } from "next/navigation";
import { z } from "zod";

import { createSession, hasAccount } from "@/lib/auth";
import { hashPassword, verifyPassword } from "@/lib/crypto";
import { db } from "@/lib/db";

export interface AuthFormState {
  error?: string;
}

const loginSchema = z.object({
  email: z.string().email("Adresse e-mail invalide"),
  password: z.string().min(1, "Mot de passe requis"),
});

export async function loginAction(
  _previous: AuthFormState,
  formData: FormData,
): Promise<AuthFormState> {
  const parsed = loginSchema.safeParse({
    email: formData.get("email"),
    password: formData.get("password"),
  });

  if (!parsed.success) {
    return { error: parsed.error.issues[0]?.message ?? "Formulaire invalide" };
  }

  const user = await db.user.findUnique({
    where: { email: parsed.data.email.toLowerCase().trim() },
  });

  // Message identique dans les deux cas : on n'indique jamais si l'e-mail existe.
  if (!user || !verifyPassword(parsed.data.password, user.passwordHash)) {
    return { error: "Identifiants incorrects." };
  }

  await createSession(user.id);
  redirect("/");
}

const setupSchema = z
  .object({
    name: z.string().min(2, "Nom trop court"),
    email: z.string().email("Adresse e-mail invalide"),
    password: z.string().min(10, "10 caractères minimum"),
    confirm: z.string(),
  })
  .refine((data) => data.password === data.confirm, {
    message: "Les mots de passe ne correspondent pas",
    path: ["confirm"],
  });

/** Création du compte unique, disponible seulement tant qu'aucun n'existe. */
export async function setupAction(
  _previous: AuthFormState,
  formData: FormData,
): Promise<AuthFormState> {
  if (await hasAccount()) {
    return { error: "Un compte existe déjà sur cette instance." };
  }

  const parsed = setupSchema.safeParse({
    name: formData.get("name"),
    email: formData.get("email"),
    password: formData.get("password"),
    confirm: formData.get("confirm"),
  });

  if (!parsed.success) {
    return { error: parsed.error.issues[0]?.message ?? "Formulaire invalide" };
  }

  const user = await db.user.create({
    data: {
      name: parsed.data.name.trim(),
      email: parsed.data.email.toLowerCase().trim(),
      passwordHash: hashPassword(parsed.data.password),
    },
  });

  await createSession(user.id);
  redirect("/");
}
