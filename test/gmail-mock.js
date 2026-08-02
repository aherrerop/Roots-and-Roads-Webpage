/* ================================================================
   Stateful Gmail mock — lets us integration-test the label / archive /
   Processed flow the same way Apps Script runs it. Overrides the minimal
   stub in mock.js (bundle this AFTER mock.js). Tests drive it via __gmail.
   ================================================================ */
(function () {
  const ALL = [];                 // every thread that exists
  const labels = new Map();       // name -> MockLabel
  let seq = 0;
  const norm = s => String(s || '').toLowerCase().replace(/[\s/]+/g, '-');

  class MockLabel {
    constructor(name) { this.name = name; }
    getName() { return this.name; }
    getThreads(start, max) {
      const list = ALL.filter(t => t.labelSet.has(this.name));
      const s = start || 0;
      return (max == null) ? list.slice(s) : list.slice(s, s + max);
    }
  }
  class MockMsg {
    constructor(subject, body, date) { this._s = subject || ''; this._b = body || ''; this._d = date || new Date(); this._id = 'm' + (++seq); }
    getId() { return this._id; }
    getSubject() { return this._s; }
    getPlainBody() { return this._b; }
    getBody() { return ''; }   // force getBestMessageText_ to use the plain body (newlines intact)
    getDate() { return this._d; }
    markRead() { return this; }
    markUnread() { return this; }
  }
  class MockThread {
    constructor(msgs) { this._id = 't' + (++seq); this._msgs = msgs; this.labelSet = new Set(); this.inInbox = true; this.unread = false; ALL.push(this); }
    getId() { return this._id; }
    getMessages() { return this._msgs; }
    getFirstMessageSubject() { return this._msgs[0] ? this._msgs[0].getSubject() : ''; }
    getLabels() { return [...this.labelSet].map(n => ({ getName: () => n })); }
    addLabel(l) { if (l && l.name) this.labelSet.add(l.name); return this; }
    removeLabel(l) { if (l && l.name) this.labelSet.delete(l.name); return this; }
    moveToArchive() { this.inInbox = false; return this; }
    isInInbox() { return this.inInbox; }
    markRead() { this.unread = false; return this; }
    markUnread() { this.unread = true; return this; }
    isUnread() { return this.unread; }
    hasLabel(name) { return this.labelSet.has(name); }
  }
  const ensureLabel = name => { if (!labels.has(name)) labels.set(name, new MockLabel(name)); return labels.get(name); };

  global.__gmail = {
    reset() { ALL.length = 0; labels.clear(); seq = 0; },
    ensure(names) { (names || []).forEach(ensureLabel); },
    msg(subject, body, date) { return new MockMsg(subject, body, date); },
    add(labelNames, msgs) {
      const t = new MockThread(Array.isArray(msgs) ? msgs : [msgs]);
      (labelNames || []).forEach(n => { ensureLabel(n); t.labelSet.add(n); });
      return t;
    },
    all() { return ALL; }
  };

  global.GmailApp = {
    getUserLabelByName(name) { return labels.has(name) ? labels.get(name) : null; },
    createLabel(name) { return ensureLabel(name); },
    search(query, start, max) {
      const q = String(query || '');
      const excl = (q.match(/-label:\S+/g) || []).map(x => x.slice(7));
      const req = (q.replace(/-label:\S+/g, '').match(/label:\S+/g) || []).map(x => x.slice(6));
      const needs = (q.match(/"([^"]+)"/g) || []).map(x => x.replace(/"/g, ''));
      const res = ALL.filter(t => {
        const ls = new Set([...t.labelSet].map(norm));
        if (!req.every(r => ls.has(norm(r)))) return false;
        if (excl.some(e => ls.has(norm(e)))) return false;
        if (needs.length) {
          const text = t.getMessages().map(m => m.getSubject() + ' ' + m.getPlainBody()).join(' ');
          if (!needs.every(n => text.indexOf(n) !== -1)) return false;
        }
        return true;
      });
      const s = start || 0;
      return (max == null) ? res.slice(s) : res.slice(s, s + max);
    }
  };
})();
