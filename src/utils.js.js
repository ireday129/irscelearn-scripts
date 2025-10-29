// --- name comparison helper (define once) ---
if (typeof namesMatchFull_ !== 'function') {
  function namesMatchFull_(f1, l1, f2, l2) {
    const norm = s => String(s || '')
      .trim()
      .replace(/\s+/g, ' ')
      .toLowerCase();
    return norm(f1) === norm(f2) && norm(l1) === norm(l2);
  }
}

// --- normalizeCompletionForUpload_ (define once) ---
if (typeof normalizeCompletionForUpload_ !== 'function') {
  /**
   * Normalize a completion date for IRS upload rules:
   * - Parses many date shapes.
   * - If the date is in the future OR more than 4 days in the past,
   *   coerce it to "yesterday" (local tz, no time).
   * - Otherwise return the parsed Date.
   *
   * @param {*} v  A Date, serial number, or string
   * @return {Date|*} normalized Date or original value if unparsable
   */
  function normalizeCompletionForUpload_(v) {
    // Prefer project-wide parser if present
    var d = (typeof parseDate_ === 'function') ? parseDate_(v) : null;

    // Simple fallback parse if parseDate_ is not available
    if (!d) {
      if (v instanceof Date && !isNaN(v)) {
        d = new Date(v.getFullYear(), v.getMonth(), v.getDate());
      } else if (typeof v === 'number' && v > 20000) {
        // Excel serial (1899-12-30 base)
        var base = new Date(1899, 11, 30);
        d = new Date(base.getTime() + v * 24 * 60 * 60 * 1000);
        d = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      } else if (v) {
        var t = new Date(String(v));
        if (!isNaN(t)) d = new Date(t.getFullYear(), t.getMonth(), t.getDate());
      }
    }

    if (!d || isNaN(d)) return v; // leave as-is if we can’t parse

    var today = new Date();
    var todayMid = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    var diffDays = Math.floor((todayMid - d) / 86400000); // 24*60*60*1000

    if (diffDays > 4 || diffDays < 0) {
      var yesterday = new Date(todayMid.getFullYear(), todayMid.getMonth(), todayMid.getDate() - 1);
      return yesterday;
    }
    return d;
  }
}