/*
  Warnings:

  - The primary key for the `workbook_permissions` table will be changed. If it partially fails, the table could be left without primary key constraint.
  - Added the required column `updated_at` to the `workbook_permissions` table without a default value. This is not possible if the table is not empty.

*/
-- DropIndex
DROP INDEX "workbook_permissions_workbook_id_user_id_idx";

-- AlterTable
ALTER TABLE "workbook_permissions" DROP CONSTRAINT "workbook_permissions_pkey",
ADD COLUMN     "updated_at" TIMESTAMPTZ NOT NULL,
ADD CONSTRAINT "workbook_permissions_pkey" PRIMARY KEY ("workbook_id", "user_id");

-- CreateIndex
CREATE INDEX "workbook_permissions_user_id_idx" ON "workbook_permissions"("user_id");
