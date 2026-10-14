import path from 'path';
import fs from 'fs';

const BASE_DATASETS_DIR = process.env.STORAGE_DATASETS_PATH
  ? path.resolve(process.env.STORAGE_DATASETS_PATH)
  : path.join(process.cwd(), 'storage', 'datasets');

export function getBaseDatasetsDir(): string {
  if (!fs.existsSync(BASE_DATASETS_DIR)) {
    fs.mkdirSync(BASE_DATASETS_DIR, { recursive: true });
  }
  return BASE_DATASETS_DIR;
}

export function getDatasetSpoolDir(datasetId: string): string {
  const dir = path.join(getBaseDatasetsDir(), 'spool', datasetId);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
  return dir;
}

export function getDatasetSpoolFilePath(datasetId: string, isFinal = false): string {
  const dir = getDatasetSpoolDir(datasetId);
  return path.join(dir, isFinal ? 'source.csv' : 'source.csv.spool');
}

export function getDatasetPartitionDir(datasetId: string): string {
  const dir = path.join(getBaseDatasetsDir(), 'partitions', datasetId);
  if (!fs.existsSync(dir)) {
    fs.mkdirSync(dir, { recursive: true });
  }
  return dir;
}

export function getPartitionBinPath(datasetId: string, partitionNumber: number, isTmp = false): string {
  const dir = getDatasetPartitionDir(datasetId);
  const partName = `part_${String(partitionNumber).padStart(3, '0')}.bin${isTmp ? '.tmp' : ''}`;
  return path.join(dir, partName);
}

export function getPartitionIdxPath(datasetId: string, partitionNumber: number, isTmp = false): string {
  const dir = getDatasetPartitionDir(datasetId);
  const partName = `part_${String(partitionNumber).padStart(3, '0')}.idx${isTmp ? '.tmp' : ''}`;
  return path.join(dir, partName);
}
