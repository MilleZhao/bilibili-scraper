"""整理评论：去无效 + 按点赞排序"""
import json, csv, re
from pathlib import Path

src = Path(__file__).parent / "outputs" / "comments.json"
data = json.loads(src.read_text(encoding="utf-8"))

# ---- 去无效规则 ----
def is_valid(c):
    text = (c.get("content") or "").strip()
    if not text:
        return False
    if len(text) <= 2:                        # 过短
        return False
    if re.match(r"^[👍🔥😂😭💪❤️🎉🤔👀🙏]+$", text):   # 纯 emoji
        return False
    if text in ("好", "加油", "支持", "6", "厉害", "牛", "顶", "打卡", "来了", "第一",
                "棒", "不错", "nb", "NB", "哈哈", "哈哈哈", "可以的", "好家伙"):
        return False
    if re.fullmatch(r"[\d\s,，]+楼", text):   # "X楼"
        return False
    if len(set(text)) <= 3 and len(text) >= 3:   # 重复单字如"啊啊啊啊"
        return False
    # 纯数字/符号
    if re.fullmatch(r"[\d\s\.\,\!\?\#\@\$\%\^\&\*\(\)\+\=]+", text):
        return False
    return True

valid = [c for c in data if is_valid(c)]
invalid = len(data) - len(valid)

# ---- 按点赞降序 ----
valid.sort(key=lambda c: c.get("likes", 0), reverse=True)

# ---- 保存 ----
out_json = src.parent / "comments_sorted.json"
out_csv  = src.parent / "comments_sorted.csv"

out_json.write_text(json.dumps(valid, ensure_ascii=False, indent=2), encoding="utf-8")

fields = ["rpid","parent_rpid","is_sub","username","user_mid","content","ctime","likes","reply_count"]
with out_csv.open("w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
    w.writeheader()
    w.writerows(valid)

# ---- 统计 ----
print(f"原始:  {len(data):,} 条")
print(f"去除:  {invalid:,} 条")
print(f"有效:  {len(valid):,} 条")
print()
print(f"Top 10 高赞评论:")
for i, c in enumerate(valid[:10], 1):
    tag = "[子]" if c["is_sub"] else "[主]"
    print(f"  {i:>2}. {tag} 👍{c['likes']:>6}  {c['username']}: {c['content'][:80]}")
print()
print(f"文件: {out_json}")
print(f"文件: {out_csv}")
