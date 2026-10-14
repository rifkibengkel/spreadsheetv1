/**
 * Pure RFC 4180 CSV Line Parser
 * Stateless, zero-dependency parser shared across ingestion workers, range readers, and viewport workers.
 */

export function parseCsvLine(line: string): string[] {
  if (!line || line.length === 0) return [''];
  const cells: string[] = [];
  const len = line.length;
  let inQuotes = false;
  let currentCell = '';
  let i = 0;

  while (i < len) {
    const char = line[i];
    if (char === '"') {
      if (inQuotes) {
        if (i + 1 < len && line[i + 1] === '"') {
          currentCell += '"';
          i += 2;
          continue;
        } else {
          inQuotes = false;
          i++;
          continue;
        }
      } else {
        inQuotes = true;
        i++;
        continue;
      }
    }
    if (char === ',' && !inQuotes) {
      cells.push(currentCell);
      currentCell = '';
      i++;
      continue;
    }
    currentCell += char;
    i++;
  }
  cells.push(currentCell);
  return cells;
}
