import { PrismaClient } from "@prisma/client";

/**
 * Client Prisma partagé. En développement, Next recharge les modules à chaque
 * édition : sans ce cache global on ouvrirait une connexion par rechargement.
 */
const globalForPrisma = globalThis as unknown as { prisma?: PrismaClient };

export const db =
  globalForPrisma.prisma ??
  new PrismaClient({
    log: process.env.NODE_ENV === "development" ? ["warn", "error"] : ["error"],
  });

if (process.env.NODE_ENV !== "production") globalForPrisma.prisma = db;
