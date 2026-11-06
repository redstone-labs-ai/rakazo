-- Additive: the previous release never reads these tables.
CREATE TABLE "directory_identities" (
    "userId" TEXT NOT NULL,
    "tenantId" TEXT NOT NULL,
    "objectId" TEXT NOT NULL,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "directory_identities_pkey" PRIMARY KEY ("userId")
);

CREATE TABLE "service_roles" (
    "userId" TEXT NOT NULL,
    "issuer" TEXT NOT NULL,
    "role" TEXT NOT NULL,
    "updatedAt" TIMESTAMP(3) NOT NULL,

    CONSTRAINT "service_roles_pkey" PRIMARY KEY ("userId","issuer")
);

CREATE UNIQUE INDEX "directory_identities_tenantId_objectId_key" ON "directory_identities"("tenantId", "objectId");

ALTER TABLE "directory_identities" ADD CONSTRAINT "directory_identities_userId_fkey" FOREIGN KEY ("userId") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE;

ALTER TABLE "service_roles" ADD CONSTRAINT "service_roles_userId_fkey" FOREIGN KEY ("userId") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE;
