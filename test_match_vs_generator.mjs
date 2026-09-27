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

test('第0-1試合を複合した試合番号として読み込む', () => {
  const [match] = context.parseMatches(`第0-1試合　ランブル
選手A VS 選手B`);

  assert.equal(match.no, '0-1');
  assert.equal(match.rule, 'ランブル');
  assert.equal(match.parseOk, true);
});

test('出力ファイル名の先頭番号は試合番号でなく画面の並び順を使う', () => {
  const items = [
    {type:'match', no:'3'},
    {type:'match', no:'0-1'},
    {type:'section'},
    {type:'match', no:'1'},
  ];

  assert.deepEqual(plain(items.map((item,index)=>context.outputCardFilename(item,index,true))), [
    '!01_match_03.png',
    '!02_match_0-1.png',
    '!03_section.png',
    '!04_match_01.png',
  ]);
  assert.equal(context.outputCardFilename(items[1],1), '02_match_0-1.png');
});

test('勝者が未設定でもZIP出力の検証エラーにしない', () => {
  const [match] = context.parseMatches(`第0-1試合　ランブル
選手A VS 選手B`);

  assert.equal(context.validateMatch(match), '');
  assert.equal(context.matchWarning(match), '勝敗が設定されていません。');

  match.groups[0].members[0].mark = '○';
  match.groups[1].members[0].mark = '●';
  assert.equal(context.matchWarning(match), '');
});
