/* 브라우저 저장소 — localStorage 를 쓰되, 막혀 있거나 용량이 차면 이번 창의 메모리로만 동작합니다.
   데이터는 이 브라우저에만 있습니다(서버로 보내지 않음). 다른 PC·다른 사람과 나누려면 JSON 백업을 쓰고, 공유 DB 는 2단계입니다. */
(function (root) {
  'use strict';
  var KEY_DB = 'data10-02.db';
  var memory = {};
  var ok = true;
  var lastError = '';
  function get(k) {
    try { var v = root.localStorage.getItem(k); return v == null && memory[k] != null ? memory[k] : v; }
    catch (e) { ok = false; return memory[k] == null ? null : memory[k]; }
  }
  function set(k, v) {
    memory[k] = v;
    try { root.localStorage.setItem(k, v); lastError = ''; return true; }
    catch (e) {
      ok = false;
      lastError = (e && (e.name === 'QuotaExceededError' || e.code === 22)) ? 'quota' : 'blocked';
      return false;
    }
  }
  function del(k) {
    delete memory[k];
    try { root.localStorage.removeItem(k); } catch (e) { ok = false; }
  }
  root.APStore = {
    loadDb: function () {
      var raw = get(KEY_DB);
      if (!raw) return root.APLogic.emptyDb();
      try { return root.APLogic.restoreDb(JSON.parse(raw)); } catch (e) { return root.APLogic.emptyDb(); }
    },
    /* true = 브라우저에 저장됨, false = 이번 창의 메모리에만 있음 */
    saveDb: function (db) { return set(KEY_DB, JSON.stringify(db)); },
    clearDb: function () { del(KEY_DB); },
    lastError: function () { return lastError; },
    sizeKb: function () { var v = get(KEY_DB); return v ? Math.round(v.length * 2 / 1024) : 0; },
    available: function () { get(KEY_DB); return ok; }
  };
})(window);
