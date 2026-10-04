/**
 * Sanitizes Unicode characters into WinAnsi / printable ASCII characters,
 * removing all control characters so pdf-lib's standard fonts never throw
 * "WinAnsi cannot encode" errors.
 *
 * Common non-WinAnsi characters (Turkish letters, smart quotes, dashes,
 * bullets, arrows) are transliterated first so they render as readable
 * text instead of being silently dropped as spaces.
 */
export function sanitizeWinAnsi(text: string): string {
  if (!text) return "";

  const CHAR_MAP: Record<string, string> = {
    "\u0130": "I", // İ
    "\u0131": "i", // ı
    "\u015E": "S", // Ş
    "\u015F": "s", // ş
    "\u011E": "G", // Ğ
    "\u011F": "g", // ğ
    "\u00DC": "U", // Ü
    "\u00FC": "u", // ü
    "\u00D6": "O", // Ö
    "\u00F6": "o", // ö
    "\u00C7": "C", // Ç
    "\u00E7": "c", // ç
    "\u20AC": "EUR",
    "\u2019": "'",
    "\u2018": "'",
    "\u201C": '"',
    "\u201D": '"',
    "\u2014": "--",
    "\u2013": "-",
    "\u2026": "...",
    "\u2022": "*",
  };

  let result = text.replace(
    /[\u0130\u0131\u015E\u015F\u011E\u011F\u20AC\u2019\u2018\u201C\u201D\u2014\u2013\u2026\u2022]/g,
    (ch) => CHAR_MAP[ch] ?? ch
  );

  return result
    .replace(/[\r\n\t\x00-\x1F\x7F]/g, " ")
    .replace(/[→⇒➔➜]/g, "->")
    .replace(/[←⇐]/g, "<-")
    .replace(/[↔⇔]/g, "<->")
    .replace(/[•●▪▫◦◆◇]/g, "*")
    .replace(/[""„]/g, '"')
    .replace(/['''‚`]/g, "'")
    .replace(/[–—―]/g, "-")
    .replace(/…/g, "...")
    .replace(/[✓✔]/g, "[OK]")
    .replace(/[✗✘]/g, "[X]")
    .replace(/[^\x20-\x7E\xA0-\xFF]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
