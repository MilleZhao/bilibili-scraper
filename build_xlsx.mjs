import fs from "node:fs/promises";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "file:///C:/Users/LENOVO/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs";

const cwd = process.cwd();
const dataDir = path.join(cwd, "data");
const outputDir = path.join(cwd, "outputs");
const outputPath = path.join(outputDir, "B站AI_AIGC原始数据集.xlsx");

const samplePosts = [
  {
    bvid: "BV11mFLziEyP",
    aid: 116006697701745,
    title: "〖牌子〗当世界过分“诚实”，我们要如何保持好奇与勇气〖B站AI创作大赛-开放赛道〗",
    author: "DiDi_OK",
    pubdate: "2026-02-03 20:31:04",
    url: "https://www.bilibili.com/video/BV11mFLziEyP/",
    desc: "这次的故事讲的是生活中突然出现了许多人们习以为常却非常“离谱”的路牌，它们似乎是对人类好奇心的警告与惩罚。",
    duration: 0,
    view_count: 19611000,
    danmaku_count: 0,
    reply_count: 22000,
    favorite_count: 140200,
    coin_count: 0,
    like_count: 350000,
    share_count: 0,
    tag: "B站AI创作大赛;原创;科幻;脑洞;AI;短片;电影;ai短片;原创短片;AIGC",
    source_keyword: "seed",
    collect_time: "template",
  },
  {
    bvid: "BV1XrZXB7EMc",
    aid: 0,
    title: "当AI开始自暴自弃 〖B站AI创作大赛-开放赛道〗",
    author: "YuFILMMAKER",
    pubdate: "2026-02-18 18:29:00",
    url: "https://www.bilibili.com/video/BV1XrZXB7EMc/",
    desc: "世界碎片 FF0A01 <<最初的回音>>",
    duration: 906,
    view_count: 121100,
    danmaku_count: 0,
    reply_count: 460,
    favorite_count: 15000,
    coin_count: 0,
    like_count: 2822,
    share_count: 0,
    tag: "B站AI创作大赛;科幻;自制短片;ALAYA",
    source_keyword: "B站AI创作大赛",
    collect_time: "template",
  },
  {
    bvid: "BV16XZHBzEJH",
    aid: 0,
    title: "血溅红牌楼 〖B站AI创作大赛-开放赛道〗",
    author: "飞顺八",
    pubdate: "2026-02-15 09:50:08",
    url: "https://www.bilibili.com/video/BV16XZHBzEJH/",
    desc: "兄弟三人创立动物园，竟因为一位不速之客反目成仇，五年后化解恩怨开始复仇……",
    duration: 228,
    view_count: 92000,
    danmaku_count: 0,
    reply_count: 95,
    favorite_count: 6142,
    coin_count: 0,
    like_count: 1184,
    share_count: 0,
    tag: "B站AI创作大赛;犯罪;动物园;电棍;抽象;山泥若",
    source_keyword: "B站AI创作大赛",
    collect_time: "template",
  },
  {
    bvid: "BV1RyAHzkEGq",
    aid: 0,
    title: "不 成 方 圆 〖B站AI创作大赛-开放赛道〗",
    author: "xyang的AI日常",
    pubdate: "2026-03-20 23:43:17",
    url: "https://www.bilibili.com/video/BV1RyAHzkEGq/",
    desc: "赛博故事会之精卫填海",
    duration: 560,
    view_count: 2179000,
    danmaku_count: 0,
    reply_count: 1444,
    favorite_count: 68000,
    coin_count: 0,
    like_count: 9282,
    share_count: 0,
    tag: "B站AI创作大赛;AI;人工智能;不成方圆;赛博故事会",
    source_keyword: "B站AI创作大赛",
    collect_time: "template",
  },
  {
    bvid: "BV1Pr6dBPEGL",
    aid: 0,
    title: "（完整版）科比参加《荒野求生》 〖B站AI创作大赛-开放赛道〗",
    author: "仁菜的苦茶",
    pubdate: "2026-01-11 16:22:40",
    url: "https://www.bilibili.com/video/BV1Pr6dBPEGL/",
    desc: "AI整活短片，围绕荒野求生与科比展开。",
    duration: 0,
    view_count: 48000,
    danmaku_count: 0,
    reply_count: 193,
    favorite_count: 2694,
    coin_count: 0,
    like_count: 458,
    share_count: 0,
    tag: "B站AI创作大赛;AI;荒野求生;搞笑;科比",
    source_keyword: "B站AI创作大赛",
    collect_time: "template",
  },
  {
    bvid: "BV1keZBBdEP5",
    aid: 0,
    title: "自费搓礼花套餐｜〖B站AI创作大赛-开放赛道〗by2026",
    author: "寒墨Brack",
    pubdate: "2026-02-20 01:46:50",
    url: "https://www.bilibili.com/video/BV1keZBBdEP5/",
    desc: "嘘！这是 AI 科幻放烟花的原创短片。",
    duration: 236,
    view_count: 149000,
    danmaku_count: 0,
    reply_count: 68,
    favorite_count: 1364,
    coin_count: 0,
    like_count: 120,
    share_count: 0,
    tag: "B站AI创作大赛;日本;大赛;人工智能;AI;机器人;高市早苗;赛道;AI创作",
    source_keyword: "B站AI创作大赛",
    collect_time: "template",
  },
  {
    bvid: "BV1n6AUzcENe",
    aid: 0,
    title: "B站首发，AI制作短片《困境》—〖B站AI创作大赛-开放赛道〗",
    author: "不必知所云",
    pubdate: "2026-02-27 17:13:20",
    url: "https://www.bilibili.com/video/BV1n6AUzcENe/",
    desc: "在可见的未来，工作可被替代，身份可被重写，思想可被复制。",
    duration: 0,
    view_count: 213000,
    danmaku_count: 0,
    reply_count: 555,
    favorite_count: 16000,
    coin_count: 0,
    like_count: 8616,
    share_count: 0,
    tag: "B站AI创作大赛;短片;剧情向;微电影;AI;原创;AI制作;AI创作",
    source_keyword: "B站AI创作大赛",
    collect_time: "template",
  },
  {
    bvid: "BV1NMwuzdETP",
    aid: 0,
    title: "暗区突围《理想国：异变》 〖B站AI创作大赛-开放赛道〗",
    author: "凌终饿鸟",
    pubdate: "2026-03-15 20:05:38",
    url: "https://www.bilibili.com/video/BV1NMwuzdETP/",
    desc: "一张模糊的照片，一具扭曲的身影。",
    duration: 0,
    view_count: 306000,
    danmaku_count: 0,
    reply_count: 930,
    favorite_count: 39000,
    coin_count: 0,
    like_count: 13000,
    share_count: 0,
    tag: "B站AI创作大赛;卡莫纳;暗区突围;理想国;AI创作;暗区突围无限",
    source_keyword: "B站AI创作大赛",
    collect_time: "template",
  },
  {
    bvid: "BV1a4ZHB4ELu",
    aid: 0,
    title: "疑似春晚机器人表演后，集体暴走了！ 〖B站AI创作大赛-开放赛道〗",
    author: "",
    pubdate: "",
    url: "https://www.bilibili.com/video/BV1a4ZHB4ELu/",
    desc: "",
    duration: 0,
    view_count: 0,
    danmaku_count: 0,
    reply_count: 0,
    favorite_count: 0,
    coin_count: 0,
    like_count: 0,
    share_count: 0,
    tag: "B站AI创作大赛",
    source_keyword: "B站AI创作大赛",
    collect_time: "template",
  },
];

const sampleComments = [
  {
    bvid: "BV11mFLziEyP",
    aid: 116006697701745,
    root_id: 0,
    parent_id: 0,
    rpid: 0,
    user_name: "",
    user_mid: 0,
    content: "待在网络可用环境中运行采集脚本后自动补全。",
    ctime: "",
    like_count: 0,
    reply_count: 0,
    url: "https://www.bilibili.com/video/BV11mFLziEyP/",
    collect_time: "template",
  },
];

async function loadJsonMaybe(filePath, fallback) {
  try {
    const text = await fs.readFile(filePath, "utf8");
    const value = JSON.parse(text);
    return Array.isArray(value) ? value : fallback;
  } catch {
    return fallback;
  }
}

function autoWidth(text, min = 10, max = 36) {
  const len = String(text ?? "").length;
  return Math.max(min, Math.min(max, Math.ceil(len * 1.3)));
}

const posts = await loadJsonMaybe(path.join(dataDir, "posts.json"), samplePosts);
const comments = await loadJsonMaybe(path.join(dataDir, "comments.json"), sampleComments);

const wb = Workbook.create();
const postsSheet = wb.worksheets.add("posts");
const commentsSheet = wb.worksheets.add("comments");
const readmeSheet = wb.worksheets.add("readme");

const postHeaders = Object.keys(posts[0] ?? samplePosts[0]);
const commentHeaders = Object.keys(comments[0] ?? sampleComments[0]);

postsSheet.getRange(`A1:${String.fromCharCode(64 + postHeaders.length)}1`).values = [postHeaders];
postsSheet.getRange(`A2:${String.fromCharCode(64 + postHeaders.length)}${posts.length + 1}`).values = posts.map((row) => postHeaders.map((key) => row[key] ?? ""));

commentsSheet.getRange(`A1:${String.fromCharCode(64 + commentHeaders.length)}1`).values = [commentHeaders];
commentsSheet.getRange(`A2:${String.fromCharCode(64 + commentHeaders.length)}${comments.length + 1}`).values = comments.map((row) => commentHeaders.map((key) => row[key] ?? ""));

const readmeValues = [
  ["字段说明", "说明"],
  ["项目主题", "B站 AI / AIGC 公开视频采集"],
  ["种子链接", "https://www.bilibili.com/video/BV11mFLziEyP/"],
  ["数据说明", "本文件当前为结构化模板与示例数据；运行 collect_bilibili_ai_aigc.py 后可补全真实采集结果。"],
  ["posts", "一行一帖，包含原帖全文/标题、点赞、评论、转发、收藏等指标及链接。"],
  ["comments", "一行一条评论，包含评论内容、楼中楼关系、点赞数和对应帖子链接。"],
  ["采集口径", "围绕种子视频扩展到同主题 AI/AIGC 视频，优先按评论量排序补足总量。"],
  ["交付建议", "先运行采集脚本，再重新生成此工作簿，以替换示例数据。"],
];
readmeSheet.getRange(`A1:B${readmeValues.length}`).values = readmeValues;

// Light touch formatting
for (const sheet of [postsSheet, commentsSheet, readmeSheet]) {
  const used = sheet.getUsedRange();
  if (used) {
    // No-op calculation to force workbook awareness of the filled range.
    used.calculate?.();
  }
}

await fs.mkdir(outputDir, { recursive: true });
const xlsx = await SpreadsheetFile.exportXlsx(wb);
await xlsx.save(outputPath);

console.log(`saved ${outputPath}`);
