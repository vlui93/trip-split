// Tests for categories and the spending breakdown. Run: node tests/breakdown.test.mjs
// Pulls the pure functions out of index.html between the cats markers.
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';

const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
const src = html.slice(html.indexOf('/* cats:start'), html.indexOf('/* cats:end */'));
const { guessCategory, breakdown, slices, CATS } =
  new Function(src + '; return { guessCategory, breakdown, slices, CATS };')();

let n = 0;
const t = (name, fn) => { fn(); n++; console.log('ok', name); };

t('guesses from everyday descriptions', () => {
  const cases = {
    'Dim sum for five': 'food', 'Dinner at Tim Ho Wan': 'food', 'Bubble tea': 'food', '海底捞火锅晚餐': 'food',
    'Taxi to hotel': 'transport', 'MTR top-up': 'transport', 'Turbojet to Macau': 'transport', 'Train tickets': 'transport',
    'Hotel deposit': 'stay', 'Airbnb Shanghai': 'stay',
    'Disneyland tickets': 'activities', 'Peak tram': 'transport', 'Museum entry tickets': 'activities',
    'Souvenirs': 'shopping', '7-Eleven snacks': 'groceries', 'Woolworths': 'groceries',
    'Rent': 'bills', 'Electricity': 'bills', 'Water': 'bills', 'Netflix': 'bills',
  };
  for (const [d, c] of Object.entries(cases)) assert.equal(guessCategory(d), c, d);
  assert.equal(guessCategory('Barcelona trip fund'), '');
  assert.equal(guessCategory(''), '');
});

const shares = ex => {                              // equal split in cents
  const who = ex.split_detail.participants, total = Math.round(ex.amount_aud * 100);
  const out = {}; who.forEach((p, i) => out[p] = Math.floor(total / who.length) + (i < total % who.length ? 1 : 0));
  return out;
};
const E = (aud, paid_by, participants, extra) => Object.assign({ amount_aud: aud, paid_by, split_detail: { participants } }, extra);
const exps = [
  E(100, 'Alex', ['Alex', 'Sam'], { category: 'food', city: 'Hong Kong' }),
  E(60, 'Sam', ['Alex', 'Sam', 'Jo'], { category: 'transport', city: 'Macau' }),
  E(30, 'Jo', ['Jo'], { description: 'Egg tarts', city: 'Hong Kong' }),       // no category: guessed food
  E(40, 'Jo', ['Sam', 'Jo'], { description: 'Mystery', city: '' }),            // no category, no guess: other
];
const asObj = r => Object.fromEntries(r.rows.map(x => [x.key, x.c]));

t('by category, everything spent, with guesses counted', () => {
  const r = breakdown(exps, 'category', '', shares);
  assert.deepEqual(asObj(r), { food: 13000, transport: 6000, other: 4000 });
  assert.equal(r.guessed, 2);
});

t('by category, one person\u2019s share', () => {
  assert.deepEqual(asObj(breakdown(exps, 'category', 'Alex', shares)), { food: 5000, transport: 2000 });
});

t('by place, by payer, and by who it was for', () => {
  assert.deepEqual(asObj(breakdown(exps, 'place', '', shares)), { 'Hong Kong': 13000, Macau: 6000, '': 4000 });
  assert.deepEqual(asObj(breakdown(exps, 'payer', '', shares)), { Alex: 10000, Sam: 6000, Jo: 7000 });
  assert.deepEqual(asObj(breakdown(exps, 'person', 'Alex', shares)), { Alex: 7000, Sam: 9000, Jo: 7000 });
});

t('slices keep a fixed order, and fold past six into one', () => {
  const order = CATS.map(c => c.id);
  const few = [{ key: 'bills', c: 5 }, { key: 'food', c: 1 }, { key: 'stay', c: 9 }];
  assert.deepEqual(slices(few, order).map(s => s.key), ['food', 'stay', 'bills']);
  const many = order.map((key, i) => ({ key, c: (i + 1) * 100 }));    // 8 categories, 'other' biggest
  const s = slices(many, order);
  assert.equal(s.length, 6);
  assert.deepEqual(s.map(x => x.key), ['activities', 'shopping', 'groceries', 'bills', 'other', '__rest']);
  assert.equal(s[5].c, 100 + 200 + 300);
});

console.log(`\n${n} passed`);
