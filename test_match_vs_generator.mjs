import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

const html = fs.readFileSync(new URL('./match_vs_generator.html', import.meta.url), 'utf8');
const scripts = [...html.matchAll(/<script(?![^>]*src=)[^>]*>([\s\S]*?)<\/script>/gi)];
const applicationScript = scripts.at(-1)?.[1] ?? '';
const parserScript = applicationScript.split('/* ========== 状態管理 ========== */')[0];
const context = {};
vm.runInNewContext(parserScript, context);
const plain = value => JSON.parse(JSON.stringify(value));

test('試合カード組み立てツールの2組カードと備考を読み込む', () => {
  const [match] = context.parseMatches(`第1試合　1/30 タッグマッチ
選手A ＆ 選手B VS 選手C ＆ 選手D
※特別ルール`);

  assert.equal(match.no, '1');
  assert.equal(match.rule, '1/30 タッグマッチ');
  assert.deepEqual(plain(match.groups.map(group => group.members.map(member => member.name))), [
    ['選手A', '選手B'], ['選手C', '選手D'],
  ]);
  assert.deepEqual(plain(match.notes), ['特別ルール']);
  assert.equal(match.parseOk, true);
});

test('3WAY以上の対戦枠を順番どおり読み込む', () => {
  const [match] = context.parseMatches(`第２試合　3WAYマッチ
選手A VS 選手B VS 選手C`);

  assert.equal(match.no, '2');
  assert.deepEqual(plain(match.groups.map(group => group.members[0].name)), ['選手A', '選手B', '選手C']);
  assert.equal(match.parseOk, true);
});

test('イリミネーションの入場順を各選手のグループとして読み込む', () => {
  const [match] = context.parseMatches(`第3試合　イリミネーション（入場順）
入場順：1. 選手A → 2. 選手B → 3. 選手C`);

  assert.deepEqual(plain(match.groups.map(group => group.members[0].name)), ['選手A', '選手B', '選手C']);
  assert.equal(match.parseOk, true);
});

test('編集された試合番号ラベルでも読み込む', () => {
  const [match] = context.parseMatches(`第0試合-1　オープニングマッチ
選手A VS 選手B`);

  assert.equal(match.no, '0');
  assert.equal(match.rule, 'オープニングマッチ');
  assert.equal(match.parseOk, true);
});
