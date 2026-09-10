/**
 * RFC 4180 Compliant CSV Row Tokenizer & Byte Scanner
 * Preserves raw cell semantics with zero numeric coercion.
 */

/**
 * Tokenizes a single raw CSV row line into an array of cell strings.
 * Preserves raw formula syntax (e.g. ="6283838524206") and long numeric identifiers.
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
        // Check for escaped quote ("")
        if (i + 1 < len && line[i + 1] === '"') {
          currentCell += '"';
          i += 2;
          continue;
        } else {
          // Closing quote
          inQuotes = false;
          i++;
          continue;
        }
      } else {
        // Opening quote
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

/**
 * Streaming Byte-Level Row Boundary Scanner.
 * Tracks quoting state across chunk boundaries to correctly detect row line endings
 * without splitting fields that contain embedded commas, escaped quotes, or embedded newlines.
 */
export class StreamingRowScanner {
  private inQuotes = false;
  private prevChar = 0;

  /**
   * Scans a Uint8Array buffer and returns the indices within the buffer
   * where unquoted line breaks (LF = 0x0A) occur.
   */
  public scanChunk(
    buffer: Uint8Array,
    startOffset: number,
    endOffset: number,
    onRowBoundary: (byteIndexInFile: number) => void,
    baseByteIndex: number
  ): void {
    for (let i = startOffset; i < endOffset; i++) {
      const b = buffer[i];

      if (b === 0x22) { // '"'
        // Flip quote state
        this.inQuotes = !this.inQuotes;
      } else if (b === 0x0a && !this.inQuotes) { // '\n' outside quotes
        onRowBoundary(baseByteIndex + i);
      }

      this.prevChar = b;
    }
  }

  public isInQuotes(): boolean {
    return this.inQuotes;
  }

  public reset(): void {
    this.inQuotes = false;
    this.prevChar = 0;
  }
}
