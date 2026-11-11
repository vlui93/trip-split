/**
 * Trip Split: the parts the app and the Sheet's script must agree on to the cent.
 *
 * One copy, used twice: index.html loads it as a script, and the deploy pushes it
 * to Apps Script next to Code.gs (there, every file shares one global scope). So
 * the Report tab and the app can never drift apart on who owes what, or on what a
 * category is called.
 *
 * Keep it plain: top-level `var` and `function` only, nothing that needs a
 * browser or a Sheet.
 */

/** Categories, in the order they are offered. `ico` names an icon in the app. */
var CATS = [
  { id: 'food',       label: 'Food & drink', ico: 'food' },
  { id: 'transport',  label: 'Transport',    ico: 'transport' },
  { id: 'stay',       label: 'Stay',         ico: 'stay' },
  { id: 'activities', label: 'Activities',   ico: 'activities' },
  { id: 'shopping',   label: 'Shopping',     ico: 'shopping' },
  { id: 'groceries',  label: 'Groceries',    ico: 'groceries' },
  { id: 'bills',      label: 'Bills',        ico: 'bills' },
  { id: 'other',      label: 'Other',        ico: 'other' }
];
var CATEGORIES = CATS.map(function (c) { return c.id; });
var CAT_LABELS = {};
CATS.forEach(function (c) { CAT_LABELS[c.id] = c.label; });

var SPLIT_LABELS = { equal: 'equal', exact: 'exact', percent: '%', itemized: 'itemized' };

/** Split `totalC` cents across `weights` ({ name: weight }) so the parts sum exactly. */
function allocate(totalC, weights) {
  var names = Object.keys(weights).filter(function (n) { return Number(weights[n]) > 0; });
  var out = {};
  if (!names.length) {
    var all = Object.keys(weights);
    if (!all.length) return out;
    var base = Math.floor(totalC / all.length);
    all.forEach(function (n, i) { out[n] = base + (i < totalC - base * all.length ? 1 : 0); });
    return out;
  }
  var sum = names.reduce(function (a, n) { return a + Number(weights[n]); }, 0);
  var used = 0, frac = [];
  names.forEach(function (n) {
    var raw = Number(weights[n]) / sum * totalC, fl = Math.floor(raw);
    out[n] = fl; used += fl; frac.push([n, raw - fl]);
  });
  // The cents left over go to the largest remainders first.
  frac.sort(function (a, b) { return b[1] - a[1]; });
  for (var i = 0, left = totalC - used; left > 0; i++, left--) out[frac[i % frac.length][0]]++;
  return out;
}

/**
 * Each person's share of one expense, in AUD cents. Always sums to amount_aud.
 * `people` is who an expense with nobody named falls to: the project's members.
 */
function sharesFor(ex, people) {
  var totalC = Math.round((Number(ex.amount_aud) || 0) * 100);
  var d = ex.split_detail || {}, weights = {};
  if (ex.split_type === 'exact' || ex.split_type === 'percent') {
    var sh = d.shares || {};
    Object.keys(sh).forEach(function (n) { if (Number(sh[n]) > 0) weights[n] = Number(sh[n]); });
  } else if (ex.split_type === 'itemized') {
    var pack = ex.items || {}, items = pack.items || [], per = {}, itemsTotal = 0;
    items.forEach(function (it) {
      var amt = Number(it.amount) || 0, who = (it.people || []).filter(Boolean);
      if (!amt || !who.length) return;
      itemsTotal += amt;
      who.forEach(function (n) { per[n] = (per[n] || 0) + amt / who.length; });
    });
    // Tax, tip or a service charge follows what each person ordered.
    var extra = Number(pack.extra) || 0;
    if (extra && itemsTotal > 0) {
      Object.keys(per).forEach(function (n) { per[n] += extra * (per[n] / itemsTotal); });
    } else if (extra) {
      var w = Object.keys(per).length ? Object.keys(per) : people;
      w.forEach(function (n) { per[n] = (per[n] || 0) + extra / w.length; });
    }
    weights = per;
  } else {
    var parts = (d.participants && d.participants.length) ? d.participants : people;
    parts.forEach(function (n) { weights[n] = 1; });
  }
  if (!Object.keys(weights).length) people.forEach(function (n) { weights[n] = 1; });
  return allocate(totalC, weights);
}

/** The rate an expense was converted at, as local units per 1 AUD, to four places. */
function expenseRate(ex) {
  var rate = ex.currency === 'AUD' ? 1 : (ex.amount_aud > 0 ? ex.amount_local / ex.amount_aud : 0);
  return Math.round(rate * 1e4) / 1e4;
}
