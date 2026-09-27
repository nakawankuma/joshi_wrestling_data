import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

const html = fs.readFileSync(new URL('./match_card_planner.html', import.meta.url), 'utf8');
const start = html.indexOf('function parseCopiedCardText');
const end = html.indexOf('function updateOutput', start);
if (start < 0 || end < 0) throw new Error('コピー用テキストのパーサーを抽出できません');
const context = {};
vm.runInNewContext(html.slice(start, end), context);
const plain = value => JSON.parse(JSON.stringify(value));

test('コピーした通常対戦と備考をプランナーに復元する', () => {
  const [match] = context.parseCopiedCardText(`第1試合　1/30 タッグマッチ
選手A ＆ 選手B VS 選手C ＆ 選手D
※特別ルール`);

  assert.equal(match.label, '第1試合');
  assert.equal(match.title, '1/30 タッグマッチ');
  assert.deepEqual(plain([match.a, match.b]), [['選手A', '選手B'], ['選手C', '選手D']]);
  assert.equal(match.memo, '特別ルール');
  assert.equal(match.layout, 'versus');
});

test('コピーした3WAY以上の対戦枠を復元する', () => {
  const [match] = context.parseCopiedCardText(`第0試合-1　4WAYマッチ
選手A VS 選手B VS 選手C VS 選手D`);

  assert.deepEqual(plain(match.slotKeys), ['a', 'b', 'c', 'd']);
  assert.deepEqual(plain([match.a, match.b, match.c, match.d]), [['選手A'], ['選手B'], ['選手C'], ['選手D']]);
  assert.equal(match.layout, 'multi');
});

test('コピーしたイリミネーションの入場順を復元する', () => {
  const [match] = context.parseCopiedCardText(`第2試合　イリミネーション
入場順：1. 選手A → 2. 選手B → 3. 選手C`);

  assert.deepEqual(plain(match.entries), ['選手A', '選手B', '選手C']);
  assert.deepEqual(plain(match.slotKeys), ['entries']);
  assert.equal(match.layout, 'elimination');
});
