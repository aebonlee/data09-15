"""SQLite 저장소 — 기준표 버전·변경 이력, 점검 기록, 검토자 판정.

파일 하나(data/stdreview.db)에 모두 담깁니다. 폐쇄망 PC 안에만 있습니다.
"""

import datetime
import json
import sqlite3
from pathlib import Path

from . import defaults

SCHEMA = """
CREATE TABLE IF NOT EXISTS criteria_versions (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  created TEXT NOT NULL,
  reason TEXT NOT NULL DEFAULT '',
  source TEXT NOT NULL DEFAULT '',
  config_json TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS criteria_changes (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  version_id INTEGER NOT NULL REFERENCES criteria_versions(id),
  kind TEXT NOT NULL, no TEXT NOT NULL, field TEXT NOT NULL,
  before TEXT NOT NULL DEFAULT '', after TEXT NOT NULL DEFAULT ''
);
CREATE TABLE IF NOT EXISTS reviews (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  created TEXT NOT NULL,
  file TEXT NOT NULL,
  std_type TEXT NOT NULL,
  title TEXT NOT NULL DEFAULT '',
  criteria_version INTEGER NOT NULL,
  is_sample INTEGER NOT NULL DEFAULT 0,
  rows_json TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS feedback (
  review_id INTEGER NOT NULL REFERENCES reviews(id),
  seq INTEGER NOT NULL,
  no TEXT NOT NULL,
  result TEXT NOT NULL DEFAULT '',
  detail TEXT NOT NULL DEFAULT '',
  verdict TEXT NOT NULL DEFAULT '',
  comment TEXT NOT NULL DEFAULT '',
  imported TEXT NOT NULL,
  UNIQUE (review_id, seq)
);
"""


def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class Store:
    def __init__(self, path):
        self.path = str(path)
        if self.path != ":memory:":
            Path(self.path).parent.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(self.path)
        self.db.row_factory = sqlite3.Row
        self.db.executescript(SCHEMA)

    def close(self):
        self.db.close()

    # ----- 기준표 -----
    def current(self):
        """(버전 번호, 설정). DB 에 아무것도 없으면 (0, 기본값)."""
        r = self.db.execute("SELECT id, config_json FROM criteria_versions ORDER BY id DESC LIMIT 1").fetchone()
        if not r:
            return 0, defaults.default_config()
        return r["id"], json.loads(r["config_json"])

    def save_version(self, cfg, changes, reason="", source=""):
        with self.db:
            cur = self.db.execute("INSERT INTO criteria_versions (created, reason, source, config_json) VALUES (?,?,?,?)",
                                  (now(), reason, source, json.dumps(cfg, ensure_ascii=False)))
            vid = cur.lastrowid
            self.db.executemany("INSERT INTO criteria_changes (version_id, kind, no, field, before, after) VALUES (?,?,?,?,?,?)",
                                [(vid, c["kind"], c["no"], c["field"], c["before"], c["after"]) for c in changes])
        return vid

    def history(self):
        vs = [dict(r) for r in self.db.execute("SELECT id, created, reason, source FROM criteria_versions ORDER BY id")]
        for v in vs:
            v["changes"] = [dict(r) for r in self.db.execute(
                "SELECT kind, no, field, before, after FROM criteria_changes WHERE version_id=? ORDER BY id", (v["id"],))]
        return vs

    # ----- 점검 기록 -----
    def add_review(self, file, std_type, title, version, rows, is_sample=False):
        with self.db:
            cur = self.db.execute(
                "INSERT INTO reviews (created, file, std_type, title, criteria_version, is_sample, rows_json) VALUES (?,?,?,?,?,?,?)",
                (now(), str(file), std_type, title, version, 1 if is_sample else 0, json.dumps(rows, ensure_ascii=False)))
        return cur.lastrowid

    def review(self, rid):
        r = self.db.execute("SELECT * FROM reviews WHERE id=?", (rid,)).fetchone()
        return dict(r) if r else None

    # ----- 검토자 판정 -----
    def upsert_feedback(self, review_id, items):
        """items: [{seq, no, result, detail, verdict, comment}] — 같은 행을 다시 넣으면 덮어씁니다."""
        t = now()
        with self.db:
            self.db.executemany(
                """INSERT INTO feedback (review_id, seq, no, result, detail, verdict, comment, imported)
                   VALUES (?,?,?,?,?,?,?,?)
                   ON CONFLICT (review_id, seq) DO UPDATE SET
                     no=excluded.no, result=excluded.result, detail=excluded.detail,
                     verdict=excluded.verdict, comment=excluded.comment, imported=excluded.imported""",
                [(review_id, i["seq"], i["no"], i.get("result", ""), i.get("detail", ""), i.get("verdict", ""), i.get("comment", ""), t) for i in items])

    def all_feedback(self, include_samples=False):
        q = """SELECT f.*, r.file, r.is_sample FROM feedback f JOIN reviews r ON r.id=f.review_id"""
        if not include_samples:
            q += " WHERE r.is_sample=0"
        return [dict(r) for r in self.db.execute(q + " ORDER BY f.review_id, f.seq")]

    def adopted_examples(self, limit_per_no=5, include_samples=False):
        """판단 프롬프트에 넣을 과거 채택 지적 예 {no: [문구]}."""
        out = {}
        for f in self.all_feedback(include_samples):
            if f["verdict"] != "채택":
                continue
            text = f["detail"] + ((" — 검토자 의견: " + f["comment"]) if f["comment"] else "")
            lst = out.setdefault(f["no"], [])
            if len(lst) < limit_per_no and text not in lst:
                lst.append(text)
        return out
