-- Additive: the previous release never reads this table.
CREATE TABLE "default_mcp_server_offers" (
    "userId" TEXT NOT NULL,
    "slug" TEXT NOT NULL,
    "createdAt" TIMESTAMP(3) NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "default_mcp_server_offers_pkey" PRIMARY KEY ("userId","slug")
);

ALTER TABLE "default_mcp_server_offers" ADD CONSTRAINT "default_mcp_server_offers_userId_fkey" FOREIGN KEY ("userId") REFERENCES "user"("id") ON DELETE CASCADE ON UPDATE CASCADE;
