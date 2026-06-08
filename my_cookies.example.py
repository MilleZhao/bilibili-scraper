# Bilibili Cookie 配置文件
# 1. 打开 bilibili.com 并登录
# 2. 按 F12 → Application → Cookies → bilibili.com
# 3. 找到下面四个值，填入等号右边
# 4. 将本文件重命名为 my_cookies.py

SESSDATA = ""     # 必填，登录态核心 cookie
BILI_JCT = ""     # 必填，csrf token
DEDEUSERID = ""   # 选填，用户 ID（不填也能用）
BUVID3 = ""       # 选填，设备标识（不填也能用）
