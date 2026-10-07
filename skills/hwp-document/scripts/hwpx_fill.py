# -*- coding: utf-8 -*-
"""HWPX 양식 채우기 — ZIP/XML 문자열 수술 도구.

HWPX 파일은 ZIP 안에 든 XML이다. 이 도구는 XML을 파싱해 다시 쓰지 않고
문자열 그대로 고쳐서 원본 양식의 서식과 구조를 보존한다.
한컴 MCP의 표 셀 이동(fill_cell_after_label hops>1)이 병합 셀에서 엉뚱한 칸을
덮어쓰는 문제를 피하려고 만든 방법이다.

명령줄:
  python hwpx_fill.py inspect 양식.hwpx [--section Contents/section0.xml] [--grep 문구]
  python hwpx_fill.py fill 양식.hwpx 결과.hwpx spec.json [--dry-run] [--overwrite]

라이브러리:
  from hwpx_fill import HwpxDoc
  doc = HwpxDoc("양식.hwpx")
  doc.replace("제00회", "제18회")
  doc.fill_empty_many([(3, "홍길동"), (1, "회사명")])
  doc.save("결과.hwpx")   # linesegarray 제거, XML 검사, mimetype 무압축 재포장까지 한다

spec.json 형식:
  {
    "section": "Contents/section0.xml",
    "ops": [
      {"op": "replace", "old": "<hp:t>원</hp:t>", "new": "<hp:t>0원</hp:t>", "expect": 1},
      {"op": "replace", "old": "<hp:t>1건</hp:t>", "new": "<hp:t>진행 중</hp:t>", "nth": 3},
      {"op": "replace_text", "old": "A & B", "new": "C & D", "expect": 1},
      {"op": "fill_empty", "cells": [[38, "40,000,000원"], [5, "사업 개요"]]},
      {"op": "clone_after", "anchor": "    - 기존 항목", "lines": ["    - 새 항목"]},
      {"op": "remove_tables", "marker": "삭제 후 제출", "forbid": ["회사명"]},
      {"op": "use_section", "section": "Contents/section1.xml"}
    ]
  }
"""
import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

DEFAULT_SECTION = "Contents/section0.xml"
EMPTY_RUN = re.compile(r'<hp:run charPrIDRef="(\d+)"/>')
TBL_OPEN = re.compile(r"<hp:tbl[\s>]")
TBL_CLOSE = "</hp:tbl>"
LINESEG = re.compile(r"<hp:linesegarray>.*?</hp:linesegarray>", re.S)
TAG = re.compile(r"<[^>]+>")


def visible_text(fragment):
    """XML 조각에서 태그를 걷어내고 사람이 보는 글자만 남긴다."""
    gt, lt = fragment.find(">"), fragment.find("<")
    if gt != -1 and (lt == -1 or gt < lt):  # 잘린 태그 꼬리로 시작하면 버린다
        fragment = fragment[gt + 1:]
    lt, gt = fragment.rfind("<"), fragment.rfind(">")
    if lt > gt:  # 잘린 태그 머리로 끝나면 버린다
        fragment = fragment[:lt]
    text = TAG.sub(" ", fragment)
    text = text.replace("&lt;", "<").replace("&gt;", ">").replace("&amp;", "&")
    return re.sub(r"\s+", " ", text).strip()


class HwpxDoc:
    def __init__(self, path, section=DEFAULT_SECTION):
        self.src = path
        self.sections = {}
        self.errors = []
        with zipfile.ZipFile(path) as z:
            self.names = z.namelist()
        self.use(section)

    # ----- 섹션 -----
    def use(self, section):
        """작업할 섹션을 바꾼다. 처음 쓰는 섹션이면 읽어 온다."""
        if section not in self.names:
            raise KeyError(f"섹션 없음: {section} (있는 것: {self.section_names()})")
        if section not in self.sections:
            with zipfile.ZipFile(self.src) as z:
                self.sections[section] = z.read(section).decode("utf-8")
        self.active = section

    def section_names(self):
        return [n for n in self.names if re.match(r"Contents/section\d+\.xml$", n)]

    @property
    def xml(self):
        return self.sections[self.active]

    @xml.setter
    def xml(self, value):
        self.sections[self.active] = value

    # ----- 치환 -----
    def replace(self, old, new, expect=1):
        """XML 문자열 그대로 치환한다. 등장 횟수가 expect와 다르면 손대지 않고 오류로 남긴다."""
        count = self.xml.count(old)
        if count != expect:
            self.errors.append(f"COUNT {count}!={expect}: {old[:50]!r}")
            return False
        self.xml = self.xml.replace(old, new)
        return True

    def replace_nth(self, old, new, n):
        """n번째(0부터) 등장만 치환한다."""
        idx = -1
        for _ in range(n + 1):
            idx = self.xml.find(old, idx + 1)
            if idx == -1:
                self.errors.append(f"NTH {n}: {old[:50]!r}")
                return False
        self.xml = self.xml[:idx] + new + self.xml[idx + len(old):]
        return True

    def replace_text(self, old, new, expect=1):
        """일반 글자 치환. &, <, > 를 XML 표기로 바꿔서 찾고 넣는다."""
        return self.replace(escape(old), escape(new), expect)

    # ----- 빈 셀 -----
    def empty_runs(self):
        return list(EMPTY_RUN.finditer(self.xml))

    def fill_empty(self, index, value):
        """index번째 빈 run(<hp:run charPrIDRef="N"/>)에 글자를 넣는다.
        넣고 나면 뒤쪽 빈 run의 번호가 하나씩 당겨진다. 여러 칸은 fill_empty_many를 쓴다."""
        runs = self.empty_runs()
        if index >= len(runs):
            self.errors.append(f"EMPTY idx {index} range {len(runs)}")
            return False
        m = runs[index]
        new = f'<hp:run charPrIDRef="{m.group(1)}"><hp:t>{escape(value)}</hp:t></hp:run>'
        self.xml = self.xml[:m.start()] + new + self.xml[m.end():]
        return True

    def fill_empty_many(self, cells):
        """(번호, 값) 목록을 큰 번호부터 채운다. 번호는 inspect가 보여준 원본 기준이다."""
        indexes = [i for i, _ in cells]
        if len(set(indexes)) != len(indexes):
            self.errors.append(f"EMPTY 중복 번호: {sorted(indexes)}")
            return False
        ok = True
        for index, value in sorted(cells, key=lambda c: c[0], reverse=True):
            ok = self.fill_empty(index, value) and ok
        return ok

    # ----- 문단 복제 -----
    def _para_block(self, anchor):
        i = self.xml.find(anchor)
        if i == -1:
            return None
        s = self.xml.rfind("<hp:p ", 0, i)
        e = self.xml.find("</hp:p>", i)
        if s == -1 or e == -1:
            return None
        return s, e + len("</hp:p>")

    def clone_after(self, anchor, lines):
        """anchor 글자가 든 문단을 복제해 글자만 lines로 바꾼 문단들을 바로 뒤에 붙인다."""
        a = escape(anchor)
        pos = self._para_block(a)
        if pos is None:
            self.errors.append(f"CLONE ANCHOR: {anchor[:40]!r}")
            return False
        s, e = pos
        block = self.xml[s:e]
        clones = []
        for line in lines:
            nb = block.replace(f">{a}<", f">{escape(line)}<")
            if nb == block:
                self.errors.append(f"CLONE SWAP: {line[:40]!r}")
                return False
            clones.append(nb)
        self.xml = self.xml[:e] + "".join(clones) + self.xml[e:]
        return True

    # ----- 안내 상자(표) 제거 -----
    def _table_end(self, start):
        depth, i = 0, start
        while True:
            m = TBL_OPEN.search(self.xml, i)
            close_i = self.xml.find(TBL_CLOSE, i)
            if close_i == -1:
                return None
            if m and m.start() < close_i:
                depth += 1
                i = m.start() + 7
            else:
                depth -= 1
                i = close_i + len(TBL_CLOSE)
                if depth == 0:
                    return i

    def _enclosing_table(self, pos):
        """pos를 감싸는 가장 안쪽 표의 (시작, 끝)."""
        starts = [m.start() for m in TBL_OPEN.finditer(self.xml, 0, pos)]
        for start in reversed(starts):
            end = self._table_end(start)
            if end is not None and end > pos:
                return start, end
        return None

    def remove_tables(self, marker, forbid=()):
        """marker 글자가 든 표(예: '<삭제 후 제출>' 안내 상자)를 통째로 지운다.
        지울 표 안에 forbid 글자가 있으면 실데이터를 지우는 사고로 보고 멈춘다."""
        m = escape(marker)
        removed = 0
        while m in self.xml:
            span = self._enclosing_table(self.xml.find(m))
            if span is None:
                self.errors.append(f"TABLE 없음: {marker!r} 를 감싼 표를 못 찾음")
                break
            block = self.xml[span[0]:span[1]]
            hit = [f for f in forbid if escape(f) in block]
            if hit:
                self.errors.append(f"TABLE 가드: 지울 표 안에 {hit} 가 있음")
                break
            self.xml = self.xml[:span[0]] + self.xml[span[1]:]
            removed += 1
        return removed

    # ----- 저장 -----
    def strip_linesegs(self):
        """레이아웃 캐시(linesegarray)를 지운다. 글자 길이가 바뀐 채로 두면 한글이
        '손상·변조 가능성' 경고를 띄우며 열기를 거부한다. 한글이 열 때 다시 계산한다."""
        total = 0
        for name, text in self.sections.items():
            total += len(LINESEG.findall(text))
            self.sections[name] = LINESEG.sub("", text)
        return total

    def save(self, dst, overwrite=False, keep_linesegs=False):
        if self.errors:
            raise ValueError("오류가 있어 저장하지 않음:\n - " + "\n - ".join(self.errors))
        if os.path.abspath(dst) == os.path.abspath(self.src):
            raise ValueError("원본 양식을 덮어쓸 수 없음 — 다른 파일명으로 저장")
        if os.path.exists(dst) and not overwrite:
            raise FileExistsError(f"이미 있음: {dst} (--overwrite 로 덮어쓰기)")
        stripped = 0 if keep_linesegs else self.strip_linesegs()
        for name, text in self.sections.items():
            ET.fromstring(text)  # 문법 오류면 여기서 예외
        with zipfile.ZipFile(self.src) as zin, zipfile.ZipFile(dst, "w") as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item in self.sections:
                    data = self.sections[item].encode("utf-8")
                ctype = zipfile.ZIP_STORED if item == "mimetype" else zipfile.ZIP_DEFLATED
                zout.writestr(item, data, compress_type=ctype)
        return stripped

    # ----- 살펴보기 -----
    def inspect(self, grep=None, context=40):
        lines = [f"섹션: {', '.join(self.section_names())}  (지금: {self.active})"]
        runs = self.empty_runs()
        lines.append(f"빈 run {len(runs)}개 — 번호: 바로 앞 글자")
        for i, m in enumerate(runs):
            before = visible_text(self.xml[max(0, m.start() - 4000):m.start()])
            lines.append(f"  [{i:3d}] …{before[-context:]}")
        if grep:
            g = escape(grep)
            lines.append(f"'{grep}' 등장 {self.xml.count(g)}회")
            for hit in re.finditer(re.escape(g), self.xml):
                ctx = visible_text(self.xml[max(0, hit.start() - 600):hit.end() + 600])
                lines.append(f"  @{hit.start()}: {ctx[:context * 3]}")
        return "\n".join(lines)


def apply_spec(doc, spec):
    if spec.get("section"):
        doc.use(spec["section"])
    for n, op in enumerate(spec.get("ops", [])):
        kind = op.get("op")
        if kind == "replace":
            if "nth" in op:
                doc.replace_nth(op["old"], op["new"], op["nth"])
            else:
                doc.replace(op["old"], op["new"], op.get("expect", 1))
        elif kind == "replace_text":
            doc.replace_text(op["old"], op["new"], op.get("expect", 1))
        elif kind == "fill_empty":
            doc.fill_empty_many([tuple(c) for c in op["cells"]])
        elif kind == "clone_after":
            doc.clone_after(op["anchor"], op["lines"])
        elif kind == "remove_tables":
            removed = doc.remove_tables(op["marker"], op.get("forbid", []))
            print(f"표 제거 ({op['marker']}): {removed}개")
        elif kind == "use_section":
            doc.use(op["section"])
        else:
            doc.errors.append(f"ops[{n}] 모르는 op: {kind!r}")


def main(argv=None):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except AttributeError:
        pass
    p = argparse.ArgumentParser(description="HWPX 양식 채우기 (ZIP/XML 문자열 수술)")
    sub = p.add_subparsers(dest="cmd", required=True)
    pi = sub.add_parser("inspect", help="빈 칸 번호와 문구 등장 횟수를 보여 준다")
    pi.add_argument("src")
    pi.add_argument("--section", default=DEFAULT_SECTION)
    pi.add_argument("--grep")
    pf = sub.add_parser("fill", help="spec.json대로 채워 새 파일로 저장한다")
    pf.add_argument("src")
    pf.add_argument("dst")
    pf.add_argument("spec")
    pf.add_argument("--dry-run", action="store_true")
    pf.add_argument("--overwrite", action="store_true")
    args = p.parse_args(argv)

    if args.cmd == "inspect":
        print(HwpxDoc(args.src, args.section).inspect(args.grep))
        return 0

    with open(args.spec, encoding="utf-8-sig") as f:
        spec = json.load(f)
    doc = HwpxDoc(args.src)
    apply_spec(doc, spec)
    if doc.errors:
        print("!!! 오류 — 저장하지 않음")
        for e in doc.errors:
            print(" -", e)
        return 1
    if args.dry_run:
        for name, text in doc.sections.items():
            try:
                ET.fromstring(LINESEG.sub("", text))
            except ET.ParseError as e:
                print(f"!!! XML 문법 오류 ({name}): {e}")
                return 1
        print("dry-run 통과 (오류 0, XML 문법 OK) — 저장 안 함")
        return 0
    try:
        stripped = doc.save(args.dst, overwrite=args.overwrite)
    except (ValueError, FileExistsError, ET.ParseError) as e:
        print(f"!!! 저장 안 함 — {e}")
        return 1
    print(f"linesegarray 제거: {stripped}개 · XML 문법 OK")
    print(f"저장: {args.dst}")
    print("다음: 한글에서 열어 경고창 없이 열리는지, 쪽수·표 수가 원본과 같은지 확인")
    return 0


if __name__ == "__main__":
    sys.exit(main())
