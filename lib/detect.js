/**
 * Column -> AJEMS field_type detection.
 * Data first, header second — a column called "Contact" full of email
 * addresses is an Email field, not a Phone field.
 *
 * Confirmed AJEMS field_type values so far: Text, Phone, Number, Date,
 * Dropdown, Email, URL, Paragraph, Textarea. The full vocabulary is still
 * unconfirmed by the AJEMS team, so anything uncertain falls back to Text,
 * which always works.
 */

const RE = {
  email: /^[^\s@]+@[^\s@]+\.[a-z]{2,}$/i,
  url:   /^https?:\/\/\S+$/i,
  phone: /^[+()\d][\d\s\-()]{6,19}$/,
  num:   /^-?[\d,]*\.?\d+%?$/,
  date:  /^(\d{4}[-/]\d{1,2}[-/]\d{1,2}|\d{1,2}[-/]\d{1,2}[-/]\d{2,4})$/,
  time:  /^\d{1,2}:\d{2}(:\d{2})?\s*(am|pm)?$/i
};

function nonEmpty(values) {
  return values.filter(v => v !== null && v !== undefined && String(v).trim() !== '');
}

function allMatch(values, re) {
  return values.length > 0 && values.every(v => re.test(String(v).trim()));
}

function detectType(label, values) {
  const vals = nonEmpty(values).slice(0, 200);
  const l = String(label || '').toLowerCase().trim();

  // ── data-driven, highest confidence ──
  if (allMatch(vals, RE.email)) return { field_type: 'Email' };
  if (allMatch(vals, RE.url))   return { field_type: 'URL' };
  if (allMatch(vals, RE.date))  return { field_type: 'Date' };
  if (allMatch(vals, RE.time))  return { field_type: 'Time' };

  // A small set of repeating values reads as a dropdown.
  const uniq = [...new Set(vals.map(v => String(v).trim()))];
  if (vals.length >= 6 && uniq.length > 1 && uniq.length <= 8 &&
      uniq.length <= vals.length / 2 &&
      uniq.every(v => v.length <= 40)) {
    return { field_type: 'Dropdown', options: uniq, multiple: false };
  }

  if (allMatch(vals, RE.phone) && vals.some(v => String(v).replace(/\D/g, '').length >= 8)) {
    return { field_type: 'Phone' };
  }
  if (allMatch(vals, RE.num)) return { field_type: 'Number' };

  // ── header hints, used only when the data was inconclusive ──
  if (/\b(e-?mail)\b/.test(l))                                   return { field_type: 'Email' };
  if (/\b(mobile|phone|contact no|whatsapp|tel)\b/.test(l))       return { field_type: 'Phone' };
  if (/\b(url|website|link)\b/.test(l))                          return { field_type: 'URL' };
  if (/\b(date|dob|birthday)\b/.test(l))                         return { field_type: 'Date' };
  if (/\b(time)\b/.test(l))                                      return { field_type: 'Time' };
  if (/\b(amount|price|cost|rate|qty|quantity|total|salary|balance|count|number|percent)\b/.test(l))
    return { field_type: 'Number' };
  if (/\b(remark|remarks|note|notes|comment|comments|narration)\b/.test(l))
    return { field_type: 'Textarea' };
  if (/\b(description|details|address|scope)\b/.test(l))         return { field_type: 'Paragraph' };

  // Long free text
  const avg = vals.length ? vals.reduce((a, v) => a + String(v).length, 0) / vals.length : 0;
  if (avg > 80) return { field_type: 'Paragraph' };

  return { field_type: 'Text' };
}

const PREFIX = {
  Text: 'text', Number: 'number', Paragraph: 'paragraph', Textarea: 'textarea',
  Dropdown: 'dropdown', Phone: 'phone', Email: 'email', URL: 'url',
  Date: 'date', Time: 'time', Checkbox: 'checkbox'
};

/** AJEMS documents field keys as <type>_<ms>. Step the clock so keys stay unique. */
function makeKey(fieldType, offset) {
  const p = PREFIX[fieldType] || 'text';
  return p + '_' + (Date.now() + (offset || 0));
}

const PLACEHOLDER = {
  Text: 'Enter text', Number: 'Enter number', Paragraph: 'Enter paragraph',
  Textarea: 'Enter notes', Dropdown: 'Select option', Phone: 'Enter phone number',
  Email: 'Enter email address', URL: 'Enter URL', Date: 'Select date', Time: 'Select time'
};

/** Build the AJEMS field objects for a set of columns. */
function buildFields(headers, rows) {
  return headers.map((h, i) => {
    const values = rows.map(r => r[h]);
    const d = detectType(h, values);
    const field = {
      field_type: d.field_type,
      label: h,
      key: makeKey(d.field_type, i),
      required: false,
      isNew: false,
      isExpanded: false,
      isVisible: true,
      is_report_hide: false,
      placeholder: PLACEHOLDER[d.field_type] || 'Enter value'
    };
    if (d.field_type === 'Dropdown') {
      field.options = d.options;
      field.multiple = false;
    }
    return field;
  });
}

/** Preview info for the UI, without generating keys. */
function analyse(headers, rows) {
  return headers.map(h => {
    const values = rows.map(r => r[h]);
    const filled = nonEmpty(values);
    const d = detectType(h, values);
    return {
      column: h,
      field_type: d.field_type,
      options: d.options || null,
      filled: filled.length,
      total: rows.length,
      samples: filled.slice(0, 3).map(v => String(v))
    };
  });
}

module.exports = { detectType, buildFields, analyse, makeKey };
