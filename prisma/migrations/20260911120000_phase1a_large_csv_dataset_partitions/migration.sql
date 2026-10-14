-- CreateEnum
CREATE TYPE "DatasetStatus" AS ENUM ('UPLOADING', 'UPLOADED', 'QUEUED', 'PARSING', 'PARTITIONING', 'INDEXING', 'VERIFYING', 'READY', 'FAILED', 'CANCELLED', 'RECOVERING');

-- CreateEnum
CREATE TYPE "PartitionStatus" AS ENUM ('PENDING', 'WRITING', 'VALIDATING', 'READY', 'FAILED');

-- CreateTable
CREATE TABLE "workbook_states" (
    "workbook_id" UUID NOT NULL,
    "latest_persisted_rev" BIGINT NOT NULL DEFAULT 0,
    "latest_calculated_rev" BIGINT NOT NULL DEFAULT 0,
    "last_snapshot_rev" BIGINT NOT NULL DEFAULT 0,
    "fencing_token" BIGINT NOT NULL DEFAULT 0,
    "active_worker_id" VARCHAR(128),
    "updated_at" TIMESTAMPTZ NOT NULL,

    CONSTRAINT "workbook_states_pkey" PRIMARY KEY ("workbook_id")
);

-- CreateTable
CREATE TABLE "workbook_snapshots" (
    "id" UUID NOT NULL,
    "workbook_id" UUID NOT NULL,
    "snapshot_rev" BIGINT NOT NULL,
    "binary_data" BYTEA NOT NULL,
    "byte_size" INTEGER NOT NULL,
    "compression_codec" VARCHAR(32) NOT NULL DEFAULT 'zstd',
    "created_at" TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "workbook_snapshots_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "mutation_batches" (
    "id" UUID NOT NULL,
    "workbook_id" UUID NOT NULL,
    "revision_id" BIGINT NOT NULL,
    "user_id" UUID NOT NULL,
    "idempotency_key" VARCHAR(128) NOT NULL,
    "deltas" JSONB NOT NULL,
    "created_at" TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "mutation_batches_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "calculation_results" (
    "id" UUID NOT NULL,
    "workbook_id" UUID NOT NULL,
    "calculated_rev" BIGINT NOT NULL,
    "cell_patches" JSONB NOT NULL,
    "execution_ms" INTEGER NOT NULL,
    "planner_strategy" VARCHAR(64) NOT NULL,
    "created_at" TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "calculation_results_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "datasets" (
    "id" UUID NOT NULL,
    "owner_id" UUID,
    "name" VARCHAR(255) NOT NULL,
    "original_filename" VARCHAR(255) NOT NULL,
    "byte_size" BIGINT NOT NULL,
    "expected_checksum" VARCHAR(64),
    "actual_checksum" VARCHAR(64),
    "row_count" BIGINT NOT NULL DEFAULT 0,
    "column_count" INTEGER NOT NULL DEFAULT 0,
    "delimiter" VARCHAR(10) NOT NULL DEFAULT ',',
    "has_header" BOOLEAN NOT NULL DEFAULT true,
    "header_row" JSONB,
    "status" "DatasetStatus" NOT NULL DEFAULT 'UPLOADING',
    "progress" DOUBLE PRECISION NOT NULL DEFAULT 0.0,
    "current_stage" VARCHAR(50),
    "spool_path" VARCHAR(512),
    "storage_dir" VARCHAR(512),
    "error_message" TEXT,
    "metadata" JSONB,
    "created_at" TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "updated_at" TIMESTAMPTZ NOT NULL,
    "completed_at" TIMESTAMPTZ,

    CONSTRAINT "datasets_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "dataset_partitions" (
    "id" UUID NOT NULL,
    "dataset_id" UUID NOT NULL,
    "partition_number" INTEGER NOT NULL,
    "stable_row_start" BIGINT NOT NULL,
    "stable_row_end" BIGINT NOT NULL,
    "row_count" INTEGER NOT NULL,
    "data_file_path" VARCHAR(512) NOT NULL,
    "index_file_path" VARCHAR(512) NOT NULL,
    "byte_size" BIGINT NOT NULL,
    "checksum" VARCHAR(64),
    "status" "PartitionStatus" NOT NULL DEFAULT 'PENDING',
    "created_at" TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
    "completed_at" TIMESTAMPTZ,

    CONSTRAINT "dataset_partitions_pkey" PRIMARY KEY ("id")
);

-- CreateTable
CREATE TABLE "upload_chunks" (
    "id" UUID NOT NULL,
    "dataset_id" UUID NOT NULL,
    "chunk_index" INTEGER NOT NULL,
    "byte_offset" BIGINT NOT NULL,
    "byte_length" INTEGER NOT NULL,
    "checksum_sha256" VARCHAR(64) NOT NULL,
    "is_committed" BOOLEAN NOT NULL DEFAULT true,
    "created_at" TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,

    CONSTRAINT "upload_chunks_pkey" PRIMARY KEY ("id")
);

-- CreateIndex
CREATE INDEX "workbook_snapshots_workbook_id_created_at_idx" ON "workbook_snapshots"("workbook_id", "created_at" DESC);

-- CreateIndex
CREATE UNIQUE INDEX "workbook_snapshots_workbook_id_snapshot_rev_key" ON "workbook_snapshots"("workbook_id", "snapshot_rev");

-- CreateIndex
CREATE UNIQUE INDEX "mutation_batches_idempotency_key_key" ON "mutation_batches"("idempotency_key");

-- CreateIndex
CREATE INDEX "mutation_batches_workbook_id_revision_id_idx" ON "mutation_batches"("workbook_id", "revision_id");

-- CreateIndex
CREATE UNIQUE INDEX "mutation_batches_workbook_id_revision_id_key" ON "mutation_batches"("workbook_id", "revision_id");

-- CreateIndex
CREATE INDEX "calculation_results_workbook_id_calculated_rev_idx" ON "calculation_results"("workbook_id", "calculated_rev");

-- CreateIndex
CREATE UNIQUE INDEX "calculation_results_workbook_id_calculated_rev_key" ON "calculation_results"("workbook_id", "calculated_rev");

-- CreateIndex
CREATE INDEX "datasets_owner_id_idx" ON "datasets"("owner_id");

-- CreateIndex
CREATE INDEX "datasets_status_idx" ON "datasets"("status");

-- CreateIndex
CREATE INDEX "datasets_created_at_idx" ON "datasets"("created_at");

-- CreateIndex
CREATE INDEX "dataset_partitions_dataset_id_status_idx" ON "dataset_partitions"("dataset_id", "status");

-- CreateIndex
CREATE INDEX "dataset_partitions_dataset_id_stable_row_start_stable_row_e_idx" ON "dataset_partitions"("dataset_id", "stable_row_start", "stable_row_end");

-- CreateIndex
CREATE UNIQUE INDEX "dataset_partitions_dataset_id_partition_number_key" ON "dataset_partitions"("dataset_id", "partition_number");

-- CreateIndex
CREATE INDEX "upload_chunks_dataset_id_idx" ON "upload_chunks"("dataset_id");

-- CreateIndex
CREATE UNIQUE INDEX "upload_chunks_dataset_id_chunk_index_key" ON "upload_chunks"("dataset_id", "chunk_index");

-- CreateIndex
CREATE UNIQUE INDEX "upload_chunks_dataset_id_byte_offset_key" ON "upload_chunks"("dataset_id", "byte_offset");

-- AddForeignKey
ALTER TABLE "workbook_states" ADD CONSTRAINT "workbook_states_workbook_id_fkey" FOREIGN KEY ("workbook_id") REFERENCES "workbooks"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "workbook_snapshots" ADD CONSTRAINT "workbook_snapshots_workbook_id_fkey" FOREIGN KEY ("workbook_id") REFERENCES "workbooks"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "mutation_batches" ADD CONSTRAINT "mutation_batches_workbook_id_fkey" FOREIGN KEY ("workbook_id") REFERENCES "workbooks"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "mutation_batches" ADD CONSTRAINT "mutation_batches_user_id_fkey" FOREIGN KEY ("user_id") REFERENCES "users"("id") ON DELETE RESTRICT ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "calculation_results" ADD CONSTRAINT "calculation_results_workbook_id_fkey" FOREIGN KEY ("workbook_id") REFERENCES "workbooks"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "datasets" ADD CONSTRAINT "datasets_owner_id_fkey" FOREIGN KEY ("owner_id") REFERENCES "users"("id") ON DELETE SET NULL ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "dataset_partitions" ADD CONSTRAINT "dataset_partitions_dataset_id_fkey" FOREIGN KEY ("dataset_id") REFERENCES "datasets"("id") ON DELETE CASCADE ON UPDATE CASCADE;

-- AddForeignKey
ALTER TABLE "upload_chunks" ADD CONSTRAINT "upload_chunks_dataset_id_fkey" FOREIGN KEY ("dataset_id") REFERENCES "datasets"("id") ON DELETE CASCADE ON UPDATE CASCADE;
