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

test('ランブルの入場順を専用レイアウトとして読み込む', () => {
  const [match] = context.parseMatches(`第0-1試合　ランブル
入場順：1. 桐生真弥 → 2. アンドレザ・ジャイアントパンダ → 3. 小波 → 4. 琉悪夏 → 5. 清司麗菜 → 6. フキゲンです → 7. ビー・プレストリー → 8. コグマ → 9. 妃南 → 10. 林下詩美 → 11. 舞華 → 12. レディ・Ｃ → 13. 壮麗亜美 → 14. 月山和香 → 15. 虎龍清花 → 16. フワちゃん → 17. 稲葉あずさ
※桐生真弥は、インカム着用`);

  assert.equal(match.no, '0-1');
  assert.equal(match.layout, 'entryOrder');
  assert.deepEqual(plain(context.entryOrderMembers(match).map(member=>member.name)), [
    '桐生真弥', 'アンドレザ・ジャイアントパンダ', '小波', '琉悪夏', '清司麗菜', 'フキゲンです',
    'ビー・プレストリー', 'コグマ', '妃南', '林下詩美', '舞華', 'レディ・Ｃ', '壮麗亜美',
    '月山和香', '虎龍清花', 'フワちゃん', '稲葉あずさ',
  ]);
  assert.deepEqual(plain(match.notes), ['桐生真弥は、インカム着用']);
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
