"use server";

import { revalidatePath } from "next/cache";

import { requireUser } from "@/lib/auth";
import { hashPassword, verifyPassword } from "@/lib/crypto";
import { db } from "@/lib/db";
import { str } from "@/lib/actions/helpers";

export interface SettingsState {
  error?: string;
  success?: string;
}

export async function updateProfile(
  _previous: SettingsState,
  formData: FormData,
): Promise<SettingsState> {
  const user = await requireUser();

  const name = str(formData, "name");
  const email = str(formData, "email").toLowerCase();

  if (name.length < 2) return { error: "Le nom est trop court." };
  if (!email.includes("@")) return { error: "Adresse e-mail invalide." };

  const existing = await db.user.findUnique({ where: { email } });
  if (existing && existing.id !== user.id) {
    return { error: "Cette adresse est déjà utilisée." };
  }

  await db.user.update({ where: { id: user.id }, data: { name, email } });
  revalidatePath("/reglages");

  return { success: "Profil mis à jour." };
}

export async function changePassword(
  _previous: SettingsState,
  formData: FormData,
): Promise<SettingsState> {
  const user = await requireUser();

  const current = str(formData, "currentPassword");
  const next = str(formData, "newPassword");
  const confirm = str(formData, "confirmPassword");

  if (next.length < 10) return { error: "Le nouveau mot de passe doit faire 10 caractères minimum." };
  if (next !== confirm) return { error: "Les deux saisies ne correspondent pas." };

  const record = await db.user.findUnique({ where: { id: user.id } });
  if (!record || !verifyPassword(current, record.passwordHash)) {
    return { error: "Mot de passe actuel incorrect." };
  }

  await db.user.update({
    where: { id: user.id },
    data: { passwordHash: hashPassword(next) },
  });

  return { success: "Mot de passe modifié." };
}
