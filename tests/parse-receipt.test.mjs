// Tests for the offline receipt parser. Run: node tests/parse-receipt.test.mjs
// Pulls the pure functions out of index.html between the ocr-parse markers.
import { readFileSync } from 'node:fs';
import assert from 'node:assert/strict';

const html = readFileSync(new URL('../index.html', import.meta.url), 'utf8');
const src = html.slice(html.indexOf('/* ocr-parse:start'), html.indexOf('/* ocr-parse:end */'));
const { parseReceiptText, lineAmount } =
  new Function(src + '; return { parseReceiptText, lineAmount };')();

const ctx = { people: ['Alex', 'Sam', 'Jo'], currency: 'HKD', city: 'Hong Kong', defaultPayer: 'Alex' };
let n = 0;
const t = (name, fn) => { fn(); n++; console.log('ok', name); };

t('amount formats', () => {
  assert.equal(lineAmount('Beer 38.00').value, 38);
  assert.equal(lineAmount('Beer 38. 00').value, 38);
  assert.equal(lineAmount('Beer 38,00').value, 38);
  assert.equal(lineAmount('Total HK$1,234.50').value, 1234.5);
  assert.equal(lineAmount('Coke $12 T').value, 12);
  assert.equal(lineAmount('Discount -20.00').value, -20);
  assert.equal(lineAmount('Beer38.00').label, 'Beer');
  assert.equal(lineAmount('Tel 28881234'), null);
  assert.equal(lineAmount('Thank you'), null);
});

t('HK cha chaan teng with 10% service', () => {
  const d = parseReceiptText(`
    翠 華 餐 廳
    Tsui Wah Restaurant
    Date: 18/10/2026 Time 19:42
    Table 12   Pax 3
    1 x 奶茶          32.00
    2 x Fish Ball Noodle   96.00
    Pineapple Bun     28.00
    Sub-Total        156.00
    10% Service Charge  15.60
    Total           171.60
    Cash            200.00
    Change           28.40`, ctx);
  assert.equal(d.split_type, 'itemized');
  assert.equal(d.amount_local, 171.6);
  assert.equal(d.extra, 15.6);
  assert.deepEqual(d.items.map(i => [i.name, i.amount]),
    [['奶茶', 32], ['Fish Ball Noodle', 96], ['Pineapple Bun', 28]]);
  assert.deepEqual(d.items[0].people, ctx.people);
  assert.equal(d.date, '2026-10-18');
  assert.equal(d.description, '翠華餐廳');
});

t('mainland receipt, full-width digits, 合计 / 实付', () => {
  const d = parseReceiptText(`
    海底捞火锅
    2026年10月25日
    肥牛 ２ 份   ９６．００
    虾滑        48.00
    酸梅汤       12.00
    合计       156.00
    优惠       -16.00
    实付       140.00
    支付宝      140.00`, { ...ctx, currency: 'CNY', city: 'Shanghai' });
  assert.equal(d.amount_local, 140);
  assert.equal(d.split_type, 'itemized');
  assert.equal(d.extra, -16);
  assert.equal(d.items.length, 3);
  assert.equal(d.date, '2026-10-25');
  assert.equal(d.currency, 'CNY');
});

t('misread item falls back to equal with a note', () => {
  const d = parseReceiptText(`
    Cafe
    Latte 42.00
    Cake 6.00
    TOTAL 98.00`, ctx);
  assert.equal(d.split_type, 'equal');
  assert.equal(d.amount_local, 98);
  assert.deepEqual(d.participants, ctx.people);
  assert.match(d.note, /split equally/);
});

t('no total line adds up the items', () => {
  const d = parseReceiptText('Egg tart 12.00\nMilk tea 18.00', ctx);
  assert.equal(d.amount_local, 30);
  assert.equal(d.split_type, 'itemized');
  assert.match(d.note, /No total line/);
});

t('garbage gives no total and asks for one', () => {
  const d = parseReceiptText('~~ ## ..\n', ctx);
  assert.equal(d.amount_local, 0);
  assert.equal(d.description, 'Receipt');
  assert.match(d.note, /Couldn't find a total/);
});

t('grand total is the largest total line, not the total qty', () => {
  const d = parseReceiptText(`
    Shop
    Water 10.00
    Chips 15.00
    Total Qty 2
    Total 25.00`, ctx);
  assert.equal(d.amount_local, 25);
  assert.equal(d.split_type, 'itemized');
});

console.log(`\n${n} passed`);
