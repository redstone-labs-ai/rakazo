-- Additive: the previous release never reads this table or column.
ALTER TABLE "messaging_link_codes" ADD COLUMN "forGroup" BOOLEAN NOT NULL DEFAULT false;

CREATE TABLE "messaging_lines" (
    "id" TEXT NOT NULL,
    "identityId" TEXT NOT NULL,
    "provider" TEXT NOT NULL,
    "threadId" TEXT NOT NULL,
    "ownerAddress" TEXT NOT NULL,
    "name" TEXT,
    "suspendedAt" TIMESTAMP(3),
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "messaging_lines_pkey" PRIMARY KEY ("id")
);

CREATE UNIQUE INDEX "messaging_lines_identityId_key" ON "messaging_lines"("identityId");

CREATE UNIQUE INDEX "messaging_lines_threadId_key" ON "messaging_lines"("threadId");

ALTER TABLE "messaging_lines" ADD CONSTRAINT "messaging_lines_identityId_fkey" FOREIGN KEY ("identityId") REFERENCES "messaging_identities"("id") ON DELETE CASCADE ON UPDATE CASCADE;
