# Official Document Formatter (公文格式化工具)

> 你是否还在为文章格式不符合公文写作要求而感到头疼？字体不对、页边距不对、行距不对、标题层级混乱……现在，只需几步，即可将任意 Word 文章一键转换为符合国家标准《党政机关公文格式》（GB/T 9704-2012）的规范公文。

专为公文写作、文秘与行政办公打造的 AI Skill，基于 A4 页面版式与国标字体规范，自动完成页面设置、字体字号、段落样式、标题层级的一整套格式化工作。

---

## 中文说明

### 一、核心功能

- **标准公文版式**：A4 纸张（210×297mm），页边距上 37mm / 下 35mm / 左 28mm / 右 26mm，每页 22 行、每行 28 字
- **公文字体自动应用**：
  - 标题：二号 方正小标宋简体
  - 正文：三号 仿宋_GB2312
  - 一级标题：三号 黑体
  - 二级标题：三号 楷体_GB2312
  - 三、四级标题：三号 仿宋_GB2312
- **段落格式规范**：首行缩进 2 字符、固定行距 28 磅、标题居中
- **标题自动识别**：支持“一、”“（一）”“1.”“（1）”编号体系及无格式标题识别
- **字体自动安装**：内置公文字体文件，检测缺失时自动安装，无需手动下载
- **安全不覆盖**：默认生成 `原文件名_formatted.docx`，原文件不受影响

### 二、使用环境

- **适用人群**：公文写作者、文秘、行政办公及需要规范格式文章的任何人
- **运行平台**：Trae Work / WorkBuddy 等 AI 办公软件（Windows 系统）
- **适用场景**：通知、报告、讲话稿、工作总结、会议纪要等 Word 文稿
- **输入格式**：`.docx` 格式 Word 文档

### 三、使用教程（三步完成）

**第 1 步 · 安装 Skill**：将下面这段话直接复制发给 AI 软件：

> 请访问 GitHub 仓库 https://github.com/starxxy/official-document-formatter ，将其中名称为 official-document-formatter 的 skill 安装到当前环境，并应用到全局。

**第 2 步 · 调用 Skill**：在输入框输入 `/转换为公文格式`，或复制以下提示词发送：

> 请使用 official-document-formatter skill，将我的文章转换为规范公文格式。

**第 3 步 · 输出文档**：把文章内容粘贴给软件（或上传 `.docx` 文件）。软件会自动完成格式化，并输出规范的 Word 文档（默认生成 `*_formatted.docx`）。

### 四、使用效果

| 维度 | 转换前 | 转换后 |
| --- | --- | --- |
| 纸张 | 任意 | A4（210×297mm） |
| 页边距 | 任意 | 上37 / 下35 / 左28 / 右26 mm |
| 标题 | 样式不一 | 二号方正小标宋简体，居中 |
| 正文 | 样式不一 | 三号仿宋_GB2312，首行缩进2字符 |
| 行距 | 任意 | 固定 28 磅 |
| 标题层级 | 混乱 | 按 一、/（一）/ 1. /（1）规范识别 |

### 五、目录结构

```
official-document-formatter/
├── LICENSE           开源许可证（Apache-2.0）
├── SKILL.md          Skill 定义与使用说明
├── config/           公文格式参数配置
├── scripts/          核心脚本（字体管理 / 标题识别 / 文档处理 / 格式化入口）
└── fonts/            内置公文字体
```

### 六、许可说明

本项目基于 **Apache-2.0** 开源协议发布，可自由使用、修改与分发。内置方正字体版权归方正字库所有，商业使用请购买正版授权。

---

## English

### Overview

> Are you still troubled by articles that don't meet official document (公文) formatting standards? Wrong fonts, wrong margins, wrong line spacing, messy heading levels... With this Skill, you can convert any Word article into a standards-compliant official document in just a few steps.

**Official Document Formatter** is an AI Skill for official writing, secretarial work, and administrative office tasks. Based on the A4 layout and the national standard GB/T 9704-2012, it automatically handles page setup, fonts, paragraph styles, and heading hierarchy.

### 1. Key Features

- **Standard document layout**: A4 (210×297mm), margins top 37mm / bottom 35mm / left 28mm / right 26mm, 22 lines per page, 28 characters per line
- **Automatic official fonts**:
  - Title: FZXiaoBiaoSong (方正小标宋简体), No.2 size
  - Body: FangSong_GB2312 (仿宋_GB2312), No.3 size
  - Heading 1: SimHei (黑体), No.3 size
  - Heading 2: KaiTi_GB2312 (楷体_GB2312), No.3 size
  - Heading 3/4: FangSong_GB2312, No.3 size
- **Standard paragraph format**: 2-character first-line indent, fixed 28pt line spacing, centered title
- **Automatic title detection**: supports "一、" "（一）" "1." "（1）" numbering and unformatted titles
- **Auto font installation**: bundled fonts, auto-installed when missing
- **Non-destructive**: outputs `originalname_formatted.docx` by default

### 2. Environment

- **Target users**: official document writers, secretaries, and admin staff
- **Platform**: AI tools such as Trae Work / WorkBuddy (Windows)
- **Scenarios**: notices, reports, speeches, work summaries, meeting minutes
- **Input**: `.docx` Word documents

### 3. Usage (3 Steps)

**Step 1 · Install the Skill**: copy and send the following message to the AI software:

> Please visit the GitHub repository https://github.com/starxxy/official-document-formatter , install the skill named official-document-formatter into the current environment, and apply it globally.

**Step 2 · Invoke the Skill**: type `/转换为公文格式` in the input box, or send the following prompt:

> Please use the official-document-formatter skill to convert my article into a standardized official document.

**Step 3 · Get the Output**: paste your article text (or upload a `.docx` file) to the software. It will format the document automatically and output a standards-compliant Word file (default: `*_formatted.docx`).

### 4. Effects (Before vs After)

| Dimension | Before | After |
| --- | --- | --- |
| Paper | Any | A4 (210×297mm) |
| Margins | Any | Top 37 / Bottom 35 / Left 28 / Right 26 mm |
| Title | Inconsistent | FZXiaoBiaoSong No.2, centered |
| Body | Inconsistent | FangSong_GB2312 No.3, 2-char indent |
| Line spacing | Any | Fixed 28pt |
| Heading levels | Messy | Standardized (一、/（一）/ 1. /（1）) |

### 5. Directory Structure

```
official-document-formatter/
├── LICENSE           Open source license (Apache-2.0)
├── SKILL.md          Skill definition & guide
├── config/           Document format parameters
├── scripts/          Core scripts (font manager / title detector / doc processor / formatter)
└── fonts/            Bundled official fonts
```

### 6. License

This project is released under the **Apache-2.0** License, free to use, modify, and distribute. The bundled Founder fonts are copyrighted by Founder Type; please purchase a license for commercial use.