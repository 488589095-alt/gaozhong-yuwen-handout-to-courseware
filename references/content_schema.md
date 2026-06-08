# content.json schema（内容 swap 点）

换讲次只改这个文件。完整可运行示例见同目录 `content_example_第2讲古诗.json`（70页成稿的实际输入）。

```jsonc
{
  "lecture_no": "第02讲",
  "title": "【古诗】细品炼字之妙，探掘诗歌深意",
  "grade": "高一",
  "teacher": "主讲老师：",          // 留空给老师填名
  "term": "2026寒春",
  "system_name": "学习成长与规划系统",
  "modules": ["学习目标","考情分析","教材链接","典型例题","针对练习"],  // 目录项，5个

  "objectives": { "header": ["序号","内容","难度星级"],
                  "rows": [["①","…（重点）","★★★✩✩"], ...] },   // 讲义学习目标表
  "star_legend": ["★★★★★  120+分必会", ...],                      // 固定4行
  "exam": { "header": ["年份","卷别","篇名","题干","设题角度"], "rows": [...] },
  "textbook_link": { "title": "教材链接", "paragraphs": ["…", "…"] },

  "example": {                       // 典型例题（1首诗，必做+选做）
    "source": "【2024·全国甲卷】阅读下面这首宋诗，完成下面的小题。",
    "poem": "诗题\n作者\n句1\n句2\n…\n【注】①…②…",   // \n 分行；【注】行自动仿宋
    "questions": [
      { "tag": "【必做题】", "type": "essay",
        "text": "…请简要赏析。（6分）",
        "answer": "【答案】①…②…",          // 只放【答案】不放【解析】
        "analysis": "…" },                   // 解析自动提取保存但默认不渲染
      { "tag": "【选做题】", "type": "choice",
        "text": "下列…不正确的一项是（      ）\nA. …\nB. …\nC. …\nD. …",
        "answer_letter": "C" }
    ]
  },

  "practices": [                     // 针对练习 N 首
    { "index": "01",                 // 两位序号，进左上圆圈
      "source": "【2025 • 广西月考 】阅读下面这首唐诗，完成下面的小题。",
      "tag": "【课前测】",            // 可选；只标在该练习第1小问
      "poem": "…",
      "questions": [ {type/text/answer_letter|answer/analysis 同上} ] }
  ],

  "knowledge": {                     // extract_handout.py --merge 自动写入
    "concept": "…", "angle_intro": "…", "ask_ways": [...],
    "word_types": {header, rows},    // 7类常考词作用表
    "steps": ["第一步  释字义，破表层", ...]   // 渲染到结尾「内容总结」页
  },
  "end": { "big": "下节课我们\n再见啦～", "small": "本期课结束" }
}
```

## 信息/现代文类（example 与 practices 用 material 代替 poem）

古诗用 `poem`；信息/论述/现代文类阅读对象是**长材料**，用 `material`（渲染时自动分页）：

```jsonc
"example": {
  "source": "【2024·全国甲卷】阅读下面的文字，完成下面的小题。",
  "material": "偷梁换柱…（长文段，\n 分自然段；编者批注【xxx】会被自动剔除）",
  "questions": [
    { "tag": "【必做题】", "type": "essay",      // 补写题：长题干带横线文段 + 短答案
      "text": "（1）请…在横线处补写恰当的词语。（3分）\n工程实例：…被安装在 _____ 附近…",
      "answer": "【答案】原柱　新柱　“假柱”" },
    { "tag": "【选做题】", "type": "choice", "text": "（1）下列…不正确的一项是（  ）\nA…\nB…\nC…\nD…", "answer_letter": "C" },
    { "tag": "【选做题】", "type": "essay", "text": "（2）…请简要分析。", "answer": "【答案】①…②…③…" }
  ]
},
"practices": [ { "index":"01", "source":"…", "material":"…", "questions":[…] } ],
"exam": { "header": ["年份","卷别","材料出处","考题再现"], "rows": [...] }   // 信息类常 4 列
```

- 用 `extract_handout_xinxi.py --docx 讲义.docx content.json` 自动抽取（材料/题/答/表/教材链接，
  例题与 4 练习的题型 choice/essay 与答案字母均自动对齐；产出后人工核对例题）。
- 材料里的编者批注（`【解释成语…】`等）渲染前会被剔除；生僻字换常用形。

## 填写要点

- 诗/题干/选项/答案文字**逐字取自讲义**，不要改写润色；讲义原文笔误也保留（忠实转写）。
- 出处标签、【必做/选做】、【课前/中/后测】、（X分）以**学生版讲义 PDF** 为准（它是官方
  curated 版）；标杆课件可交叉确认。
- 选择题 `type:"choice"` + `answer_letter`；简答 `type:"essay"` + `answer`（以【答案】开头）。
- `knowledge` 和每题 `analysis` 由 `extract_handout.py --merge` 从讲义 docx 自动抽取
  （含答案交叉校验），不要手抄长解析。
- 生僻字（如"斥𮭨"）换常用形（"斥鷃"），防字体回退乱码。
