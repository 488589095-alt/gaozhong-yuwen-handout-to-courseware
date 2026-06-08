# gaozhong-yuwen-handout-to-courseware

高中语文「讲义 docx → 课件 PPT」生成 skill（Claude Code skill）。
输入讲义（+学生版 PDF 核范围、可选标杆课件对照），用 `高一语文.pptx` 模板版式，
一键产出整套课件：封面/目录/学习目标/考情/教材链接/典型例题/针对练习/内容总结/结束页。

已在两类讲次验证：**第2讲古诗（70 页）** + **第10讲信息·情境填空（79 页）**。

## 目录
- `SKILL.md` —— 流程、范围分类法(A/B/C)、字体规范、防出框三道防线、踩坑表
- `scripts/`
  - `build_pptx.py` —— 渲染（克隆模板+关系重映射、字体定稿、选择题标红、_fit_size 防出框、材料分页）
  - `extract_handout.py` —— 古诗类讲义提取器
  - `extract_handout_xinxi.py` —— 信息/现代文类讲义提取器（长材料+补写题）
  - `check_overflow.py` —— 出稿自检（文字溢出+越界，双0才交付）
- `references/` —— 字体颜色规范 / 模板页映射 / content schema / 两讲示例 content.json
- `evals/` —— 测试用例

## 用法
```bash
# ① 提取讲义内容 → content.json（古诗用 extract_handout，信息类用 extract_handout_xinxi）
python3 scripts/extract_handout_xinxi.py 讲义.docx content.json
# ② 按 references/content_schema.md 核对/补全 content.json
# ③ 渲染
python3 scripts/build_pptx.py --content content.json --template "高一语文.pptx" -o 课件.pptx
# ④ 出稿自检（必须双0）
python3 scripts/check_overflow.py 课件.pptx
```

## 双题型
| 类型 | 阅读对象字段 | 渲染 | 题型 |
|---|---|---|---|
| 古诗 | `poem`（短，1页） | 楷体居中、注释①②③上标、【注】仿宋 | 选择 / 简答赏析 |
| 信息·现代文 | `material`（长，自动分页） | 宋体左对齐、右上"材料 n/m" | 选择 / 补写填空 / 简答 |

> 字体名用 黑体/微软雅黑/宋体/楷体/仿宋（WPS/PowerPoint 通用）。
