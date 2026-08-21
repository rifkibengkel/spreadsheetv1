const testFormulas = [
  "SUM('Data All'!B2:B10)",
  "SUM(Data All!B2:B10)",
  "VLOOKUP(B4, Unik Konsumen!A:B, 2, 0)",
  "COUNTIF('Data Valid'!$O$2:$O$100, B4)",
  "Summary Hadiah!C4"
];

function sanitizeFormulaSheetNames(formula) {
  if (!formula) return formula;
  // Match unquoted sheet names containing spaces before an exclamation mark
  // E.g. Data All! -> 'Data All'! (ignoring already quoted sheet names)
  return formula.replace(/(?<!')\b([A-Za-z0-9_]+(?:\s+[A-Za-z0-9_]+)+)\b(?!')!/g, "'$1'!");
}

console.log('=== Formula Sheet Name Sanitizer Test ===');
testFormulas.forEach(f => {
  console.log('Original :', f);
  console.log('Sanitized:', sanitizeFormulaSheetNames(f));
  console.log('---');
});
