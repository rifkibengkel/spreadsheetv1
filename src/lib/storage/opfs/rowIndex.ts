/**
 * N+1 Uint64 Row Index Manager
 * Invariant 2: N + 1 uint64 little-endian offsets for N data rows.
 * Section 5: Safe Uint64 handling using BigInt and DataView.
 */

export class RowIndex {
  /**
   * Serializes an array of BigInt offsets into a little-endian Uint8Array buffer.
   */
  public static serialize(offsets: BigUint64Array | bigint[]): Uint8Array {
    const count = offsets.length;
    const buffer = new ArrayBuffer(count * 8);
    const view = new DataView(buffer);

    for (let i = 0; i < count; i++) {
      const val = typeof offsets[i] === 'bigint' ? (offsets[i] as bigint) : BigInt(offsets[i]);
      view.setBigUint64(i * 8, val, true); // little-endian
    }

    return new Uint8Array(buffer);
  }

  /**
   * Wraps an ArrayBuffer as a BigUint64Array.
   */
  public static fromBuffer(buffer: ArrayBuffer): BigUint64Array {
    if (buffer.byteLength % 8 !== 0) {
      throw new Error(`Invalid index buffer size: ${buffer.byteLength} is not a multiple of 8.`);
    }
    return new BigUint64Array(buffer);
  }

  /**
   * Retrieves the byte slice bounds for a specific data row.
   * Row index is 0-based (0 to N-1).
   */
  public static getRowBounds(
    indexArray: BigUint64Array,
    rowIndex: number
  ): { start: number; end: number; length: number } {
    const totalDataRows = indexArray.length - 1;
    if (rowIndex < 0 || rowIndex >= totalDataRows) {
      throw new Error(`Row index ${rowIndex} out of bounds (valid: 0 to ${totalDataRows - 1})`);
    }

    const startBig = indexArray[rowIndex];
    const endBig = indexArray[rowIndex + 1];

    if (startBig > BigInt(Number.MAX_SAFE_INTEGER) || endBig > BigInt(Number.MAX_SAFE_INTEGER)) {
      throw new Error(`Offset exceeds Number.MAX_SAFE_INTEGER`);
    }

    const start = Number(startBig);
    const end = Number(endBig);
    const length = end - start;

    if (length < 0) {
      throw new Error(`Corrupted index: end (${end}) < start (${start}) at row ${rowIndex}`);
    }

    return { start, end, length };
  }

  /**
   * Validates the integrity of an index buffer according to Invariant 2 & Section 11.
   */
  public static validate(
    indexArray: BigUint64Array,
    expectedRowCount: number,
    expectedDataSizeBytes: number
  ): { valid: boolean; error?: string } {
    const expectedLength = expectedRowCount + 1;

    if (indexArray.length !== expectedLength) {
      return {
        valid: false,
        error: `Index count mismatch: expected ${expectedLength} (${expectedRowCount} rows + 1), found ${indexArray.length}`,
      };
    }

    if (indexArray[0] !== BigInt(0)) {
      return {
        valid: false,
        error: `First offset must be 0, found ${indexArray[0]}`,
      };
    }

    const finalOffset = indexArray[expectedRowCount];
    if (finalOffset !== BigInt(expectedDataSizeBytes)) {
      return {
        valid: false,
        error: `Final offset must equal dataSizeBytes (${expectedDataSizeBytes}), found ${finalOffset}`,
      };
    }

    // Monotonicity and range check: 0 <= index[i] <= index[i+1] <= dataSizeBytes
    let prev = BigInt(0);
    const maxDataSize = BigInt(expectedDataSizeBytes);
    for (let i = 0; i < indexArray.length; i++) {
      const curr = indexArray[i];
      if (curr < prev) {
        return {
          valid: false,
          error: `Non-monotonic offset at index ${i}: current ${curr} < previous ${prev}`,
        };
      }
      if (curr > maxDataSize) {
        return {
          valid: false,
          error: `Out-of-range offset at index ${i}: current ${curr} > dataSizeBytes ${expectedDataSizeBytes}`,
        };
      }
      prev = curr;
    }

    return { valid: true };
  }
}
