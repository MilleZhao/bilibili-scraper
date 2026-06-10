import json, csv
from pathlib import Path

src = Path(__file__).parent / "outputs" / "comments_all.json"
dst = src.with_suffix(".csv")

data = json.loads(src.read_text(encoding="utf-8"))

fields = ["rpid","parent_rpid","is_sub","username","user_mid","content","ctime","likes","reply_count"]

with dst.open("w", encoding="utf-8-sig", newline="") as f:
    w = csv.DictWriter(f, fieldnames=fields)
    w.writeheader()
    w.writerows(data)

main_count = sum(1 for r in data if not r["is_sub"])
sub_count  = sum(1 for r in data if r["is_sub"])

print(f"✅ 导出完成！")
print(f"   主评论: {main_count:,} 条")
print(f"   子回复: {sub_count:,} 条")
print(f"   合计:   {len(data):,} 条")
print(f"   文件:   {dst}")
