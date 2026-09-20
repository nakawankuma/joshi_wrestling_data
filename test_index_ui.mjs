import assert from 'node:assert/strict';
import fs from 'node:fs';
import test from 'node:test';
import vm from 'node:vm';

const html = fs.readFileSync(new URL('./index.html', import.meta.url), 'utf8');
const scriptMatch = html.match(/<script>([\s\S]*?)<\/script>/i);

if (!scriptMatch) {
  throw new Error('index.htmlのJavaScriptが見つかりません');
}

class FakeClassList {
  constructor() {
    this.values = new Set();
  }

  add(...names) {
    names.forEach(name => this.values.add(name));
  }

  remove(...names) {
    names.forEach(name => this.values.delete(name));
  }

  contains(name) {
    return this.values.has(name);
  }
}

class FakeElement {
  constructor(id = '') {
    this.id = id;
    this.value = '';
    this.hidden = false;
    this.textContent = '';
    this.innerHTML = '';
    this.scrollTop = 0;
    this.scrollHeight = 0;
    this.className = '';
    this.classList = new FakeClassList();
    this.dataset = {};
    this.style = {
      removeProperty() {},
    };
    this.children = [];
    this.listeners = new Map();
  }

  appendChild(child) {
    this.children.push(child);
    return child;
  }

  append(...children) {
    this.children.push(...children);
  }

  replaceChildren(...children) {
    this.children = [...children];
    this.innerHTML = '';
  }

  addEventListener(type, listener) {
    this.listeners.set(type, listener);
  }

  querySelectorAll() {
    return [];
  }
}

function createRuntime(url = 'http://localhost/index.html') {
  const elementIds = [
    'debugInfo',
    'wrestlerSearch',
    'tableBody',
    'totalWrestlers',
    'filteredWrestlers',
    'filteredCard',
    'controls',
    'yearDetailView',
    'tableContainer',
  ];
  const elements = new Map(elementIds.map(id => [id, new FakeElement(id)]));
  const sortable = new FakeElement('sortable');
  const windowListeners = new Map();
  const pushedUrls = [];
  const replacedUrls = [];

  const document = {
    documentElement: new FakeElement('documentElement'),
    getElementById(id) {
      return elements.get(id) ?? null;
    },
    createElement() {
      return new FakeElement();
    },
    createTextNode(text) {
      const node = new FakeElement();
      node.textContent = String(text);
      return node;
    },
    querySelector(selector) {
      return selector === '.sortable' ? sortable : null;
    },
    querySelectorAll(selector) {
      if (selector === '.promotion-header' || selector === '#tableBody tr') {
        return [];
      }
      return [];
    },
  };

  const window = {
    location: new URL(url),
    history: {
      pushState(_state, _unused, nextUrl) {
        pushedUrls.push(String(nextUrl));
        window.location = new URL(nextUrl, window.location);
      },
      replaceState(_state, _unused, nextUrl) {
        replacedUrls.push(String(nextUrl));
        window.location = new URL(nextUrl, window.location);
      },
    },
    addEventListener(type, listener) {
      windowListeners.set(type, listener);
    },
  };

  const context = vm.createContext({
    URL,
    URLSearchParams,
    console: {
      log() {},
      error() {},
    },
    document,
    setTimeout(callback) {
      callback();
      return 1;
    },
    window,
  });

  vm.runInContext(scriptMatch[1], context, { filename: 'index.html' });

  return {
    context,
    elements,
    pushedUrls,
    replacedUrls,
    window,
    windowListeners,
  };
}

function run(context, source) {
  return vm.runInContext(source, context);
}

test('デバッグログは入力文字列をHTMLとして解釈しない', () => {
  const runtime = createRuntime();
  const debugInfo = runtime.elements.get('debugInfo');
  debugInfo.innerHTML = '';

  run(runtime.context, 'debugLog(\'<img src=x onerror="alert(1)">\')');

  assert.doesNotMatch(debugInfo.innerHTML, /<img\b/i);
});

test('フィルター名は操作欄でHTMLとして解釈しない', () => {
  const runtime = createRuntime();
  const controls = runtime.elements.get('controls');

  run(runtime.context, `
    currentFilter = {type: 'promotion', value: '<img src=x onerror="alert(1)">'};
    showControls();
  `);

  assert.doesNotMatch(controls.innerHTML, /<img\b/i);
  assert.equal(controls.children[0].textContent, '← 戻る');
});

test('団体名は詳細見出しでHTMLとして解釈しない', () => {
  const runtime = createRuntime();
  const detailView = runtime.elements.get('yearDetailView');

  run(runtime.context, `
    showPromotionDetail('<img src=x onerror="alert(1)">', [], 1);
  `);

  assert.doesNotMatch(detailView.innerHTML, /<img\b/i);
});

test('正規表現の特殊文字を検索しても例外にならない', () => {
  const runtime = createRuntime();
  runtime.elements.get('wrestlerSearch').value = '(';

  assert.doesNotThrow(() => run(runtime.context, 'searchWrestlers()'));
});

test('検索中の並べ替えは検索結果だけを再描画する', () => {
  const runtime = createRuntime();
  const usedFilteredData = run(runtime.context, `
    isSearching = true;
    const expectedSearchResults = [['検索結果']];
    getCurrentFilteredData = () => expectedSearchResults;
    renderTable = data => { globalThis.__usedFilteredData = data === expectedSearchResults; };
    toggleSort();
    globalThis.__usedFilteredData;
  `);

  assert.equal(usedFilteredData, true);
});

test('検索入力では履歴を追加せず現在のURLを置き換える', () => {
  const runtime = createRuntime();
  run(runtime.context, `
    renderTable = () => {};
    updateSearchResultHeaders = () => {};
    calculateSearchResultWrestlers = () => {};
  `);

  runtime.elements.get('wrestlerSearch').value = '__該当なし1__';
  run(runtime.context, 'searchWrestlers()');
  runtime.elements.get('wrestlerSearch').value = '__該当なし2__';
  run(runtime.context, 'searchWrestlers()');

  assert.equal(runtime.pushedUrls.length, 0);
  assert.equal(runtime.replacedUrls.length, 2);
});

test('ブラウザーの戻る・進む操作で画面状態を復元する', () => {
  const runtime = createRuntime();
  const restoreState = runtime.windowListeners.get('popstate');

  assert.equal(typeof restoreState, 'function');

  runtime.window.location = new URL('http://localhost/index.html?search=朱里');
  restoreState();

  assert.equal(runtime.elements.get('wrestlerSearch').value, '朱里');
  assert.equal(run(runtime.context, 'isSearching'), true);

  runtime.window.location = new URL('http://localhost/index.html');
  restoreState();

  assert.equal(runtime.elements.get('wrestlerSearch').value, '');
  assert.equal(run(runtime.context, 'isSearching'), false);
  assert.equal(runtime.pushedUrls.length, 0);
});

test('検索中に団体フィルターを解除しても検索件数を表示する', () => {
  const runtime = createRuntime('http://localhost/index.html?search=test&filter=promotion&promotion=AWG');
  run(runtime.context, `
    isSearching = true;
    searchWrestlers = () => updateFilteredCount(2);
    resetFilter();
  `);

  assert.equal(runtime.elements.get('filteredCard').hidden, false);
  assert.equal(runtime.elements.get('filteredWrestlers').textContent, 2);
});

test('非表示のフィルター操作欄はCSSで表示されない', () => {
  assert.match(html, /\.controls:not\(\[hidden\]\)\s*\{/);
  assert.doesNotMatch(html, /\.controls\s*\{[^}]*display\s*:\s*flex/si);
});
