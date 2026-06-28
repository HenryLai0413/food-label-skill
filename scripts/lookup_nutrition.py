#!/usr/bin/env python3
"""
TFND 食品營養成分查詢工具
資料來源：台灣食品營養成分資料庫 2025版（衛生福利部食藥署）

使用方式：
  python lookup_nutrition.py <食材名稱> [--top N] [--category 食品分類]

  python lookup_nutrition.py 雞肉
  python lookup_nutrition.py 猴頭菇 --top 5
  python lookup_nutrition.py 醬油 --category 調味料及香辛料類

輸出：每100公克的營養成分（JSON格式）
  calories, protein, fat, saturated_fat, trans_fat_g,
  carbs, fiber, sugar, sodium（mg）
"""

import sys
import json
import csv
import os
import io
import re
import argparse
from pathlib import Path

TFND_CSV = Path(__file__).parent.parent / "references" / "tfnd_2025.csv"

# 常見食材別名對應表（台灣廚房用語 → TFND 正式品名搜尋詞）
SYNONYMS = {
    "麻油": "芝麻油",       # 麻油=芝麻油
    "香油": "芝麻油",
    "胡麻油": "芝麻油",
    "黑麻油": "黑芝麻油",
    "豬油": "豬板油",
    "薑": "老薑",
    "雞高湯": "雞湯",
    "柴魚": "鰹魚",
    "昆布": "海帶",
    "干貝": "扇貝",
    "豬五花": "豬腹脅肉",
    "五花肉": "豬腹脅肉",
    "梅花肉": "豬梅花肉",
    "絞肉": "豬絞肉",
    "豬絞肉": "豬後腿絞肉",
    "雞絞肉": "雞胸肉",
    "太白粉": "樹薯澱粉",
    "地瓜粉": "甘藷澱粉",
    "玉米澱粉": "玉米澱粉",
    "木薯粉": "樹薯澱粉",
    "味醂": "味霖",
    "清酒": "米酒",
    "料酒": "米酒",
    "蠔油": "蠔油",
    "辣豆瓣醬": "豆瓣醬",
    "甜麵醬": "甜麵醬",
    "豆腐乳": "紅豆腐乳",
    "板豆腐": "傳統豆腐",
    "嫩豆腐": "嫩豆腐",
    "油豆腐": "油豆腐",
    "豆皮": "豆包",
    "凍豆腐": "凍豆腐",
    "花生油": "花生油",
    "葵花油": "葵花油",
    "橄欖油": "橄欖油",
    "苦茶油": "苦茶油",
    "紅蘿蔔": "紅蘿蔔",
    "紅薯": "甘藷",
    "地瓜": "甘藷",
    "馬鈴薯": "馬鈴薯",
    "洋蔥": "洋蔥",
    "大蒜": "蒜頭",
    "蒜頭": "蒜頭",
    "蔥": "青蔥",
    "韭菜": "韭菜",
    "芹菜": "芹菜",
    "香菜": "香菜",
    "木耳": "黑木耳",
    "白木耳": "白木耳",
    "金針菇": "金針菇",
    "香菇": "香菇",
    "鴻喜菇": "鴻喜菇",
    "杏鮑菇": "杏鮑菇",
    "巴西蘑菇": "巴西蘑菇",
}


def load_tfnd() -> list[dict]:
    if not TFND_CSV.exists():
        raise FileNotFoundError(f"TFND 資料庫不存在：{TFND_CSV}")
    rows = []
    with open(TFND_CSV, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            rows.append(row)
    return rows


def normalize(text: str) -> str:
    """移除空白、全形轉半形，小寫化，方便比對。"""
    text = text.strip().lower()
    text = text.replace("（", "(").replace("）", ")")
    text = re.sub(r"\s+", "", text)
    return text


def score_match(query: str, row: dict) -> tuple[int, int]:
    """
    回傳 (主分數, 名稱長度的負值)。
    主分數越高越好；同分時名稱越短（越精確）優先。
    0 = 不匹配。
    """
    q = normalize(query)
    name = normalize(row.get("name", ""))
    aliases = normalize(row.get("aliases", "") or "")
    n_len = len(name) if name else 1
    q_len = len(q)

    # 覆蓋率加成：查詢佔品名比例越高，分數加成越多（避免長名稱搶排名）
    coverage_bonus = int((q_len / n_len) * 30)  # 最多 +30（查詢=品名時 coverage=1）

    # 完全符合品名
    if q == name:
        return (1000, -n_len)
    # 品名開頭（加覆蓋率加成）
    if name.startswith(q):
        return (800 + coverage_bonus, -n_len)
    # 品名結尾
    if name.endswith(q):
        return (700 + coverage_bonus, -n_len)
    # 品名包含整個查詢字串
    if q in name:
        return (600 + coverage_bonus, -n_len)
    # 俗名完全符合
    for alias in aliases.split(","):
        alias = alias.strip()
        if not alias:
            continue
        if q == alias:
            return (750, -n_len)
        if q in alias:
            return (500, -n_len)

    # 字元重疊分數（查詢字元出現在品名中的比例）
    if q_len >= 2:
        overlap = sum(1 for c in q if c in name)
        ratio = overlap / q_len
        if ratio >= 0.5:
            return (int(10 + ratio * 20), -n_len)

    return (0, 0)


def to_float(val) -> float | None:
    try:
        f = float(val)
        return None if f != f else f  # NaN check
    except (ValueError, TypeError):
        return None


def format_result(row: dict) -> dict:
    return {
        "id":            row.get("id", ""),
        "category":      row.get("category", ""),
        "name":          row.get("name", ""),
        "aliases":       row.get("aliases", ""),
        "per_100g": {
            "calories":       to_float(row.get("calories")),
            "protein":        to_float(row.get("protein")),
            "fat":            to_float(row.get("fat")),
            "saturated_fat":  to_float(row.get("saturated_fat")),
            "trans_fat_g":    to_float(row.get("trans_fat_g")),
            "carbs":          to_float(row.get("carbs")),
            "fiber":          to_float(row.get("fiber")),
            "sugar":          to_float(row.get("sugar")),
            "sodium_mg":      to_float(row.get("sodium")),
        },
        "data_source": "TFND 2025（衛生福利部食藥署）",
    }


def search(query: str, top_n: int = 5, category: str = None) -> list[dict]:
    rows = load_tfnd()
    if category:
        rows = [r for r in rows if category in r.get("category", "")]

    # 套用同義詞對照表
    resolved = SYNONYMS.get(query, query)

    scored = [(score_match(resolved, r), r) for r in rows]
    # 過濾主分數 = 0 的結果
    scored = [((main, tie), r) for (main, tie), r in scored if main > 0]
    # 主分數降序，同分時 tie（=-名稱長度）降序（即名稱越短越前）
    scored.sort(key=lambda x: x[0], reverse=True)

    return [format_result(r) for _, r in scored[:top_n]]


def calculate_weighted(ingredient_weights: list[dict]) -> dict:
    """
    加權計算整道產品的每100g營養成分。

    ingredient_weights 格式：
    [
      {"name": "雞腿肉", "weight_g": 200},
      {"name": "薑",     "weight_g": 10},
      ...
    ]

    回傳：每100g加權平均值（未找到的成分會標記 missing）
    """
    total_weight = sum(i["weight_g"] for i in ingredient_weights)
    if total_weight == 0:
        return {}

    result = {k: 0.0 for k in
              ["calories", "protein", "fat", "saturated_fat",
               "trans_fat_g", "carbs", "fiber", "sugar", "sodium_mg"]}
    missing = []
    found_weight = 0.0

    for item in ingredient_weights:
        name = item["name"]
        weight = item["weight_g"]
        hits = search(name, top_n=1)
        if not hits:
            missing.append(name)
            continue

        n = hits[0]["per_100g"]
        ratio = weight / 100.0
        for key in result:
            val = n.get(key)
            if val is not None:
                result[key] += val * ratio
        found_weight += weight

    # 換算回每100g
    if found_weight > 0:
        scale = 100.0 / total_weight
        result = {k: round(v * scale, 2) for k, v in result.items()}
    else:
        result = {k: None for k in result}

    return {
        "per_100g":       result,
        "total_weight_g": total_weight,
        "found_weight_g": found_weight,
        "missing":        missing,
        "coverage_pct":   round(found_weight / total_weight * 100, 1),
        "data_source":    "TFND 2025（衛生福利部食藥署）",
    }


if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

    parser = argparse.ArgumentParser(description="TFND 食品營養成分查詢")
    parser.add_argument("query", help="食材名稱")
    parser.add_argument("--top", type=int, default=5, help="顯示前N筆（預設5）")
    parser.add_argument("--category", default=None, help="限定食品分類")
    parser.add_argument("--json", action="store_true", help="輸出純JSON")
    args = parser.parse_args()

    results = search(args.query, top_n=args.top, category=args.category)

    if not results:
        print(f"找不到符合「{args.query}」的食材，請嘗試：")
        print("  - 更短的關鍵字（如「雞腿」而非「去骨雞腿肉」）")
        print("  - 中文俗名（如「猴頭菇」）")
        sys.exit(0)

    if args.json:
        print(json.dumps(results, ensure_ascii=False, indent=2))
    else:
        print(f"\n查詢「{args.query}」的 TFND 結果（每100公克）：\n")
        for i, r in enumerate(results, 1):
            n = r["per_100g"]
            print(f"  {i}. [{r['category']}] {r['name']}")
            if r["aliases"]:
                print(f"     俗名：{r['aliases']}")
            print(f"     熱量：{n['calories']} 大卡 | 蛋白質：{n['protein']} g | 脂肪：{n['fat']} g")
            print(f"     碳水：{n['carbs']} g（糖 {n['sugar']} g）| 鈉：{n['sodium_mg']} mg")
            print()
