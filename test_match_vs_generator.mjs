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

test('残り2組の備考から4WAYタッグの全グループを読み込む', () => {
  const [match] = context.parseMatches(`4、4WAYタッグマッチ（30分1本勝負）
○バラモンシュウ＆バラモンケイ（14分8秒、五体不満足）髙木三四郎●＆一般人・澤宗紀
※残り2組はタノムサク鳥羽＆ジムトレーナーKotoka、阿部史典＆Mac松下

選手コメント/寸評`);

  assert.equal(match.no, '4');
  assert.equal(match.time, '14分8秒');
  assert.equal(match.technique, '五体不満足');
  assert.deepEqual(plain(match.groups.map(group=>group.members.map(member=>member.name))), [
    ['バラモンシュウ', 'バラモンケイ'],
    ['髙木三四郎', '一般人・澤宗紀'],
    ['タノムサク鳥羽', 'ジムトレーナーKotoka'],
    ['阿部史典', 'Mac松下'],
  ]);
  assert.deepEqual(plain(match.groups.map(group=>group.members.map(member=>member.mark))), [
    ['○', ''], ['●', ''], ['', ''], ['', ''],
  ]);
  assert.deepEqual(plain(match.notes), []);
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

test('記事の混入を防ぎ、3WAYの第三選手とフィニッシュ技を復元する', () => {
  const [match] = context.parseMatches(`3、3WAYバトル（15分1本勝負）
○舞華（5分35秒、片エビ固め）儛島エマ●
※みちのくドライバーⅡ。もう一人はレディ・C
選手コメント/寸評  ×7
関連記事＆リンク`);
  assert.equal(match.parseOk,true);
  assert.deepEqual(plain(match.groups.map(g=>g.members[0].name)),['舞華','儛島エマ','レディ・C']);
  assert.equal(match.technique,'みちのくドライバーⅡ→片エビ固め');
});

test('時間切れ引き分けの王者と挑戦者を分離する', () => {
  const [match]=context.parseMatches(`7、選手権試合（15分1本勝負）
＜王者＞△八神蘭奈（時間切れ引き分け）稲葉あずさ△＜挑戦者＞
※第14代王者が11度目の防衛に成功
記事見出し`);
  assert.equal(match.parseOk,true);
  assert.deepEqual(plain(match.groups.map(g=>g.members[0])),[{name:'八神蘭奈',mark:'△'},{name:'稲葉あずさ',mark:'△'}]);
  assert.equal(context.validateMatch(match),'');
  assert.equal(context.matchWarning(match),'');
});

test('編集用テキストの末尾の引き分け記号も勝敗として保持する', () => {
  context.document={getElementById:()=>({value:''})};
  const [match]=context.parseMatches(`7、選手権試合
【タイトル区分】通常タイトル戦
【王者グループ】1
【グループ】△八神蘭奈
【グループ】稲葉あずさ△
【決着時間】時間切れ引き分け`);
  assert.equal(match.groups[1].members[0].mark,'△');
  assert.equal(match.groups[1].members[0].name,'稲葉あずさ');
});

test('画像描画へ決着時間・決まり技・引き分け防衛を渡す', () => {
  const drawing=applicationScript.slice(applicationScript.indexOf('function drawMatch('),applicationScript.indexOf('/* ========== エディタUI構築'));
  vm.runInNewContext(drawing,context);
  const labels=[];
  const ctx=new Proxy({measureText:text=>({width:String(text).length*12}),fillText:text=>labels.push(text)},{get:(obj,key)=>key in obj?obj[key]:(()=>{})});
  const canvas={getContext:()=>ctx};
  const [match]=context.parseMatches(`7、選手権試合
＜王者＞△八神蘭奈（時間切れ引き分け）稲葉あずさ△＜挑戦者＞`);
  const settings={size:1080,bgType:'white',textColor:'#111111',accentColor:'#111111',showTimeTech:true,showTitleTag:true,showNotes:false};
  context.drawMatch(canvas,match,settings);
  assert.ok(labels.includes('王座防衛'));
  assert.ok(labels.includes('時間切れ引き分け'));
  labels.length=0;
  const [three]=context.parseMatches(`3、3WAYバトル
○舞華（5分35秒、片エビ固め）儛島エマ●
※みちのくドライバーⅡ。もう一人はレディ・C`);
  context.drawMatch(canvas,three,settings);
  assert.ok(labels.includes('3WAY'));
  assert.ok(labels.includes('レディ・C'));
  assert.ok(labels.includes('みちのくドライバーⅡ→片エビ固め'));
});

const october29=`週刊プロレス・デジタル
10/29 スターダム
STARDOM CRIMSON NIGHTMARE 2026
1、20分1本勝負
○AZM＆天咲光由（8分13秒、片エビ固め）星来芽依＆儛島エマ●
※ダイビング・フットスタンプ
タッグリーグ不参加の王者組・AZM＆天咲はアメリカへ
選手コメント/寸評  ×5
2、3WAYトリオバトル（20分1本勝負）
○伊藤麻希＆古沢稀杏＆梨杏（9分56秒、片エビ固め）山下りな＆青木いつ希＆鉄アキラ●
※ダブルアームDDT。もう1組は朱里＆壮麗亜美＆妃南
選手コメント/寸評  ×10
2・11エディオンアリーナ大阪、3・6広島サンプラザ、シンデレラ・トーナメント2027決定
選手コメント/寸評
3、20分1本勝負
さくらあや＆○玖麗さやか＆安納サオリ＆金屋あんね（11分8秒、片エビ固め）レディ・C＆稲葉ともか＆八神蘭奈＆虎龍清花●
※ファルコンにゃろー
選手コメント/寸評  ×10
4、STARS vs H.A.T.E.イリミネーションマッチ（時間無制限）
羽南＆飯田沙耶＆舞華＆林下詩美＆HANAKO＆月山和香＆姫ゆりあ＆浜辺纏（8-7）刀羅ナツコ＆渡辺桃＆小波＆上谷沙弥＆琉悪夏＆吏南＆ビー・プレストリー＆稲葉あずさ
※1人残りでSTARSの勝利。敗退順は寸評に記載
羽南「STARSが勝ったってことはさ、もうH.A.T.E.一強なんて言わせないよ」
選手コメント/寸評  ×12
5、スペシャルミックスドタッグマッチ（30分1本勝負）
○なつぽい＆鷹木信悟（17分25秒、片エビ固め）水森由菜●＆後藤洋央紀
※MADE IN JAPANっぽい
選手コメント/寸評
鷹木がブーイング受けなつぽいに謝罪
選手コメント/寸評
6、ワールド・オブ・スターダム選手権試合（30分1本勝負）
<王者>○鈴季すず（23分6秒、ジャーマン・スープレックス・ホールド）スターライト・キッド●<挑戦者>
※第22代王者が3度目の防衛に成功
舞華と吏南登場。すずと舞華が先にワールド戦、勝者が12・29両国で吏南を迎撃へ
選手コメント/寸評  ×16
王座記録変遷史
関連リンク
AZM、天咲光由がスターダム10・27後楽園を欠場＆カード変更
ニュース 2026/10/26(Monday)
週刊プロレスTOP`;

test('10/29の6試合を記事から分離し、トリオ3WAY・技・チーム勝敗を復元する',()=>{
  const matches=context.parseMatches(october29);
  assert.equal(matches.length,6);
  for(const match of matches){assert.equal(match.parseOk,true);assert.equal(context.validateMatch(match),'');}
  assert.deepEqual(plain(matches[1].groups.map(g=>g.members.map(m=>m.name))),[
    ['伊藤麻希','古沢稀杏','梨杏'],['山下りな','青木いつ希','鉄アキラ'],['朱里','壮麗亜美','妃南']]);
  assert.equal(matches[1].technique,'ダブルアームDDT→片エビ固め');
  assert.equal(matches[2].technique,'ファルコンにゃろー→片エビ固め');
  assert.equal(matches[4].technique,'MADE IN JAPANっぽい→片エビ固め');
  assert.equal(matches[3].score,'8-7');
  assert.equal(matches[3].time,'');
  assert.equal(matches[3].winnerGroup,0);
  assert.deepEqual(plain(matches[3].groups.map(g=>g.members.length)),[8,8]);
  assert.ok(matches[3].groups.every(g=>g.members.every(m=>m.mark==='')));
  assert.equal(context.matchWarning(matches[3]),'');
  assert.ok(!JSON.stringify(matches).includes('エディオン'));
  assert.ok(!JSON.stringify(matches).includes('謝罪'));
  assert.equal(matches[5].championGroup,0);
});

test('3WAYで2グループしかない編集済みデータは保存時もエラーにする',()=>{
  context.document={getElementById:()=>({value:''})};
  const [match]=context.parseMatches(`2、3WAYトリオバトル
【グループ】○伊藤麻希＆古沢稀杏＆梨杏
【グループ】山下りな＆青木いつ希＆●鉄アキラ`);
  assert.equal(match.parseOk,false);
  assert.match(context.validateMatch(match),/WAY数/);
});

test('チーム勝敗とスコアを編集テキスト経由で往復し画像へ反映する',()=>{
  vm.runInNewContext(applicationScript.slice(applicationScript.indexOf('let matches = []'),applicationScript.indexOf('function syncRawText()')),context);
  context.fixture=october29;
  context.document={getElementById:()=>({value:''})};
  vm.runInNewContext('matches = parseMatches(fixture)',context);
  const text=context.serializeEditableText();
  const parsed=context.parseMatches(text);
  assert.equal(parsed[3].score,'8-7');
  assert.equal(parsed[3].winnerGroup,0);
  assert.equal(parsed[3].groups[0].members[0].mark,'');
  const labels=[];
  const ctx=new Proxy({measureText:text=>({width:String(text).length*12}),fillText:text=>labels.push(text)},{get:(obj,key)=>key in obj?obj[key]:(()=>{})});
  context.drawMatch({getContext:()=>ctx},parsed[3],{size:1080,bgType:'white',showTimeTech:true,showTitleTag:true,showNotes:false});
  assert.ok(labels.includes('結果スコア：8-7'));
  assert.equal(labels.filter(text=>text==='WIN').length,1);
  assert.equal(labels.filter(text=>text==='VS').length,1);
});
