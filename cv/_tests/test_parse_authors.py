# -*- coding: utf-8 -*-
"""
test_parse_authors.py — build_cv.py 의 `_parse_authors()` 단위테스트.

build_cv.py 는 import 만으로 CV 를 빌드해 버리므로, 소스에서 센티넬
`# --- [authors-parse] ---` ~ `# --- [/authors-parse] ---` 블록만 떼어내 exec 한다.
(_txt 는 테스트용 최소 스텁으로 주입 — build_cv 의 _txt 와 동일 동작: None -> "")

실행:  PYTHONIOENCODING=utf-8 python _tests/test_parse_authors.py
종료코드 0 = 전부 통과.
"""
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "..", "build_cv.py")

BEGIN = "# --- [authors-parse] ---"
END = "# --- [/authors-parse] ---"


def load_block():
    txt = open(SRC, encoding="utf-8").read()
    i, j = txt.find(BEGIN), txt.find(END)
    if i < 0 or j < 0:
        raise SystemExit("센티넬 블록을 찾지 못했습니다 — build_cv.py 의 주석을 확인하세요.")
    ns = {"re": re, "_txt": lambda v: "" if v is None else str(v).strip()}
    exec(compile(txt[i:j], "<authors-parse>", "exec"), ns)
    return ns


NS = load_block()
_parse_authors = NS["_parse_authors"]
_is_initials = NS["_is_initials"]

CASES = [
    # (설명, 입력, 기대)
    ("APA 3저자 & 있음", "Jang, G., Kim, S., & Lee, J.",
     ["Jang, G.", "Kim, S.", "Lee, J."]),
    ("APA 3저자 & 없음", "Jang, G., Kim, S., Lee, J.",
     ["Jang, G.", "Kim, S.", "Lee, J."]),
    ("APA 2저자 & 있음", "Jang, G., & Kim, S.",
     ["Jang, G.", "Kim, S."]),
    ("APA 2저자 and 구분자", "Jang, G., and Kim, S.",
     ["Jang, G.", "Kim, S."]),
    ("APA 2저자 쉼표 없는 and", "Jang, G. and Kim, S.",
     ["Jang, G.", "Kim, S."]),
    ("APA 1저자", "Jang, G.", ["Jang, G."]),
    ("APA 4저자 다중 이니셜·하이픈", "Jang, G., Kim, J. S., Park, S.-H., & Cho, H. R.",
     ["Jang, G.", "Kim, J. S.", "Park, S.-H.", "Cho, H. R."]),
    ("APA 마침표 없는 이니셜", "Jang, G, Kim, S", ["Jang, G", "Kim, S"]),
    ("APA 이니셜 붙은 형태", "Jang, G.H., & Kim, S.", ["Jang, G.H.", "Kim, S."]),
    ("구형식 풀네임 2저자", "Gukhwa Jang, Saehoon Kim",
     ["Gukhwa Jang", "Saehoon Kim"]),
    ("구형식 풀네임 3저자", "Gukhwa Jang, Saehoon Kim, Jae Seung Lee",
     ["Gukhwa Jang", "Saehoon Kim", "Jae Seung Lee"]),
    ("구형식 풀네임 1저자", "Gukhwa Jang", ["Gukhwa Jang"]),
    ("구형식 and 구분자", "Gukhwa Jang and Saehoon Kim",
     ["Gukhwa Jang", "Saehoon Kim"]),
    ("구형식 선두 이니셜", "G. Jang, S. Kim", ["G. Jang", "S. Kim"]),
    ("빈값(빈 문자열)", "", []),
    ("빈값(None)", None, []),
    ("빈값(공백만)", "   ", []),
    ("말단 공백·말단 쉼표", "  Jang, G.,  Kim, S. ,  ",
     ["Jang, G.", "Kim, S."]),
    ("중복 구분자", "Jang, G.,, & Kim, S.", ["Jang, G.", "Kim, S."]),
    ("대문자 성은 이니셜로 보지 않는다", "JANG, GUKHWA", ["JANG", "GUKHWA"]),
]

INIT_CASES = [
    ("G.", True), ("G", True), ("J. S.", True), ("S.-H.", True), ("G.H.", True),
    ("Lee", False), ("LEE", False), ("Gukhwa", False), ("", False),
    ("Jang, G.", False),
]


def main():
    fails = []
    for name, src, want in CASES:
        got = _parse_authors(src)
        ok = got == want
        print(("PASS  " if ok else "FAIL  ") + name + (f"\n      got={got!r} want={want!r}" if not ok else ""))
        if not ok:
            fails.append(name)
    for tok, want in INIT_CASES:
        got = _is_initials(tok)
        ok = got == want
        print(("PASS  " if ok else "FAIL  ") + f"_is_initials({tok!r}) -> {got}")
        if not ok:
            fails.append(f"_is_initials({tok!r})")
    print(f"\n{len(CASES) + len(INIT_CASES) - len(fails)}/{len(CASES) + len(INIT_CASES)} passed")
    if fails:
        print("FAILED: " + ", ".join(fails))
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
