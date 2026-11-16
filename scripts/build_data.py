"""把书目 Excel 转成网页用的数据文件。

用法：python scripts/build_data.py
读取仓库根目录下的 .xlsx，生成 public/books.json（列表用）和 public/details.json（简介，稍后加载）。

每本书的主题标签按这个顺序决定：
  1. Excel「标签」列里手动填的（用逗号、顿号或分号隔开），填了就以它为准
  2. scripts/tags.json 里按 ISBN 记下的人工分类
  3. 都没有时，按下面 KEYWORDS 里的关键词从书名、作者、简介里猜
"""

import glob
import json
import os
import re
import sys
from datetime import datetime, timedelta, timezone

from openpyxl import load_workbook

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PREFERRED_XLSX = "私家书藏-全部书籍.xlsx"
TAGS_PATH = os.path.join(ROOT, "scripts", "tags.json")
OUT_DIR = os.path.join(ROOT, "public")

# 网页上主题的固定顺序；Excel 里手填的新标签会排在这些后面
TAG_ORDER = [
    "酷儿", "女性主义与性别", "性与亲密", "社会学与人类学", "哲学与思想",
    "历史与政治", "劳动与经济", "公益与行动", "文学", "漫画与艺术",
    "心理与灵性", "身体与健康", "生活与工具书",
]

# 新书没有人工分类时的兜底规则。只放指向性强的词，宁缺毋滥。
KEYWORDS = {
    "酷儿": ["酷儿", "酷兒", "同志", "同性恋", "同性戀", "彩虹", "出柜", "出櫃", "跨性别", "跨性別",
           "queer", "gay", "lesbian", "lgbt", "拉拉", "断袖", "邱妙津", "赛菊克", "賽菊克", "埃里蓬"],
    "女性主义与性别": ["女性主义", "女性主義", "女权", "女權", "厌女", "厭女", "父权", "父權", "性别", "性別",
                "女性", "妇女", "婦女", "上野千鹤子", "波伏瓦", "巴特勒", "胡克斯", "费代里奇"],
    "性与亲密": ["性社会学", "性教育", "性史", "性工作", "亲密关系", "親密關係", "艾滋病时代", "避孕",
             "情欲", "情慾", "潘绥铭", "黄盈盈"],
    "社会学与人类学": ["社会学", "社會學", "人类学", "人類學", "民族志", "田野", "社会理论", "社會理論",
                "韦伯", "布尔迪厄", "布迪厄", "涂尔干", "戈夫曼", "鲍曼", "吉登斯", "项飙", "费孝通", "斯科特"],
    "哲学与思想": ["哲学", "哲學", "伦理学", "倫理學", "思想史", "福柯", "黑格尔", "康德", "尼采", "德里达",
              "拉康", "德勒兹", "阿伦特", "哈贝马斯", "朗西埃", "存在主义", "现象学"],
    "历史与政治": ["历史", "歷史", "史纲", "全球史", "帝国", "帝國", "革命", "政治", "战争", "戰爭",
              "史学", "史學", "王朝", "清代", "民国", "民國", "冷战"],
    "劳动与经济": ["劳动", "勞動", "工人", "打工", "工厂", "工廠", "职场", "職場", "经济学", "經濟學",
              "资本主义", "資本主義", "不平等", "马克思", "阶级", "階級"],
    "公益与行动": ["公益", "慈善", "社会运动", "社會運動", "引导技", "引导工具", "社区组织", "非营利",
              "志愿", "倡导", "行动者"],
    "文学": ["小说", "小說", "散文", "随笔", "隨筆", "诗集", "詩集", "诗歌", "文学", "文學", "短篇", "长篇"],
    "漫画与艺术": ["漫画", "漫畫", "绘本", "繪本", "画集", "設定集", "设定集", "美术", "艺术", "藝術",
              "音乐", "音樂", "乐队", "电影", "電影", "comics", "插画"],
    "心理与灵性": ["心理学", "心理學", "心理咨询", "精神分析", "抑郁", "创伤", "佛教", "佛法", "灵性", "冥想", "疗愈"],
    "身体与健康": ["医学", "醫學", "医师", "临床", "健身", "肌力", "力量训练", "训练指南", "营养", "營養",
              "马拉松", "拉伸", "瑜伽", "解剖"],
    "生活与工具书": ["词汇", "詞彙", "雅思", "gre", "vocabulary", "食谱", "便当", "便當", "养猫", "猫",
               "旅行", "地理", "百科", "手册"],
}

BRACKETS = r"[\[\(（【〔［]\s*([一-鿿]{1,6})\s*[\]\)）】〕］]"


def text(v):
    return "" if v is None else str(v).strip()


def clean_title(s):
    s = re.sub(r"\s+", " ", text(s))
    m = re.fullmatch(r"《(.+)》", s)
    return m.group(1) if m and "《" not in m.group(1) else s


def clean_author(s):
    """'[美] 甲;[英]乙;' -> '[美]甲、[英]乙'，并去掉重复的人名。"""
    names, seen = [], set()
    for part in re.split(r"[;；]+", text(s)):
        part = re.sub(r"\s+", " ", part).strip(" ,，")
        if not part:
            continue
        part = re.sub(r"^" + BRACKETS + r"\s*", lambda m: "[" + m.group(1) + "]", part)
        bare = re.sub(r"^\[[^\]]+\]", "", part)
        bare = re.sub(r"[\(（][^\)）]*[\)）]", "", bare).strip()
        key = bare or part
        if key in seen:
            continue
        seen.add(key)
        names.append(part)
    return "、".join(names)


def clean_long(s):
    s = text(s).replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"[ \t　]+\n", "\n", s)
    s = re.sub(r"\n{3,}", "\n\n", s)
    return s.strip()


def split_tags(s):
    out = []
    for t in re.split(r"[,，;；/、|]+", text(s)):
        t = t.strip()
        if t and t not in out:
            out.append(t)
    return out


def guess_tags(title, author, desc):
    head = f"{title} {author}".lower()
    body = desc[:300].lower()
    scores = {}
    for tag, words in KEYWORDS.items():
        score = 0
        for w in words:
            if w in head:
                score += 3
            elif w in body:
                score += 1
        if score:
            scores[tag] = score
    ranked = sorted(scores, key=lambda t: (-scores[t], TAG_ORDER.index(t)))
    strong = [t for t in ranked if scores[t] >= 2]
    return (strong or ranked)[:2]


def find_xlsx():
    preferred = os.path.join(ROOT, PREFERRED_XLSX)
    if os.path.exists(preferred):
        return preferred
    found = sorted(p for p in glob.glob(os.path.join(ROOT, "*.xlsx")) if not os.path.basename(p).startswith("~$"))
    if not found:
        sys.exit("仓库根目录下没有找到 .xlsx 书目表格")
    return found[0]


def main():
    path = find_xlsx()
    ws = load_workbook(path, read_only=True, data_only=True).worksheets[0]
    rows = ws.iter_rows(values_only=True)
    header = [text(h) for h in next(rows)]
    if "书名" not in header:
        sys.exit("表格第一行需要有「书名」这一列")

    curated = {}
    if os.path.exists(TAGS_PATH):
        with open(TAGS_PATH, encoding="utf-8") as f:
            curated = json.load(f)

    books, details, guessed, seen_ids = [], {}, [], set()
    for n, values in enumerate(rows, start=2):
        row = {h: text(v) for h, v in zip(header, values) if h}
        title = clean_title(row.get("书名"))
        if not title:
            continue
        isbn = re.sub(r"\.0$", "", row.get("ISBN", ""))
        book_id = isbn if isbn and isbn not in seen_ids else f"r{n}"
        seen_ids.add(book_id)

        author = clean_author(row.get("作者"))
        book_desc = clean_long(row.get("图书简介"))
        author_desc = clean_long(row.get("作者简介"))

        tags = split_tags(row.get("标签"))
        if not tags and isbn in curated:
            tags = list(curated[isbn].get("标签", []))
        elif not tags:
            tags = guess_tags(title, author, book_desc)
            guessed.append(f"{title} -> {'、'.join(tags) or '（没猜出来）'}")

        book = {"id": book_id, "t": title, "g": tags, "l": row.get("书柜") or "还没上架"}
        if author:
            book["a"] = author
        if row.get("出版社"):
            book["p"] = row["出版社"]
        for key, col in (("lent", "借给"), ("status", "阅读状态"), ("note", "备注")):
            if row.get(col):
                book[key] = row[col]
        books.append(book)
        if book_desc or author_desc:
            details[book_id] = [book_desc, author_desc]

    extra = []
    for b in books:
        for t in b["g"]:
            if t not in TAG_ORDER and t not in extra:
                extra.append(t)

    os.makedirs(OUT_DIR, exist_ok=True)
    with open(os.path.join(OUT_DIR, "books.json"), "w", encoding="utf-8") as f:
        updated = datetime.now(timezone(timedelta(hours=8))).strftime("%Y年%-m月%-d日")
        json.dump({"updated": updated, "tags": TAG_ORDER + extra, "books": books}, f, ensure_ascii=False, separators=(",", ":"))
    with open(os.path.join(OUT_DIR, "details.json"), "w", encoding="utf-8") as f:
        json.dump(details, f, ensure_ascii=False, separators=(",", ":"))

    print(f"读取 {os.path.basename(path)}：共 {len(books)} 本，{len(details)} 本有简介")
    if guessed:
        print(f"以下 {len(guessed)} 本没有人工分类，按关键词猜的（可在 Excel「标签」列或 scripts/tags.json 里改）：")
        for line in guessed:
            print("  " + line)


if __name__ == "__main__":
    main()
