<div align="center">

# 🎯 PPT Master

**AI-powered PowerPoint slide generator** — 说需求，出 PPTX，双击即用。

AI Skill for [Claude Code](https://docs.anthropic.com/en/docs/claude-code) / [Codex](https://openai.com/index/codex/) / OpenClaw · 12 Slide Types · 8 Color Schemes · Real `.pptx` Output

<br>

[![python](https://img.shields.io/badge/python-3.7+-blue?style=flat-square&logo=python)](https://www.python.org)
[![python-pptx](https://img.shields.io/badge/python--pptx-1.0.0-blue?style=flat-square)](https://python-pptx.readthedocs.io)
[![License](https://img.shields.io/badge/license-MIT-green?style=flat-square)](LICENSE)

<br>

[Quick Start](#-quick-start) · [Gallery](#-gallery) · [Layouts](#-layouts) · [Customization](#-customization)

</div>

---

> [!TIP]
> 把需求说出来，就能拿到一份可以直接演示的 PPTX 文件。
> **不需要 PowerPoint**，不需要排版，支持任何能跑 Python 的平台。

<br>

## ✨ 这是什么

PPT Master 是一个 AI Skill（插件），用于 Claude Code / Codex / OpenClaw，能根据你的需求自动生成专业的 PowerPoint 文件：

```
需求 → 大纲确认 → 逐页生成 → .pptx 导出 → 交互修改
```

你只需要在对话中确认大纲，剩下的全部自动完成。

### 支持的 PPT 场景

| 场景 | 建议页数 | 风格 |
|------|---------|------|
| 💼 商业路演 / 融资计划 | 8-12 | 简洁有力，视觉冲击 |
| 🚀 产品发布 / 新品介绍 | 10-15 | 产品亮点突出 |
| 📊 工作汇报 / 年度总结 | 10-20 | 数据驱动，专业稳重 |
| 📈 项目复盘 / 数据分析 | 10-15 | 结果导向，图表丰富 |
| 🏢 企业介绍 / 公司宣传 | 10-15 | 品牌调性，视觉统一 |

### 12 种专业版式

Cover · Section Divider · Text Only · Text + Image · Two Column · Three Column · Data Chart · Data Table · Timeline · Quote · List · Contact

### 8 种配色方案

🔵 Blue · 🔴 Red · 🟢 Green · 🟣 Purple · 🟠 Orange · ⚫ Dark · 🔷 Teal · 🎨 Custom

---

## 🚀 Quick Start

### 安装

```bash
# Claude Code
git clone https://github.com/mianmian5/ppt-master.git ~/.claude/skills/ppt-master

# Codex
git clone https://github.com/mianmian5/ppt-master.git ~/.codex/skills/ppt-master

# 安装依赖
pip install python-pptx
```

### 使用

1. 创建一个新目录，进入
2. 说：**"帮我做一个产品介绍的 PPT"**
3. 根据提示回答几个问题（场景、页数、配色）
4. 确认大纲 → 自动生成 `.pptx` 文件

### 命令行预览（不依赖 AI）

```bash
# 用示例数据生成
python3 scripts/generate.py examples/product-launch/outline.json test.pptx
```

---

## 📐 工作原理

```
用户的文字需求
      │
      ▼
 ┌──────────────────┐
 │ 需求收集         │ ← AI 在对话中询问：场景、受众、页数、配色
 └──────────────────┘
      │
      ▼
 ┌──────────────────┐
 │ 大纲交互确认     │ ← AI 提出结构 → 用户逐页确认 → 保存 outline.json
 └──────────────────┘
      │
      ▼
 ┌──────────────────┐
 │ generate.py      │ ← 读取 outline.json，调用 python-pptx 生成
 └──────────────────┘
      │
      ▼
     .pptx 文件 ← 双击即可打开编辑
```

---

## 📁 项目结构

```
ppt-master/
├── SKILL.md                   # 核心 Skill 指令文件（AI 读取并执行）
├── references/
│   ├── layouts.md             # 12 种版式的视觉描述和槽位定义
│   └── layout-registry.yaml   # 版式选型规则和场景映射
├── assets/
│   ├── themes/                # 8 种配色方案定义
│   │   ├── blue.yaml
│   │   ├── red.yaml
│   │   ├── green.yaml
│   │   ├── purple.yaml
│   │   ├── orange.yaml
│   │   ├── dark.yaml
│   │   └── teal.yaml
│   └── config.yaml            # 用户配置模板
├── scripts/
│   ├── generate.py            # 核心 PPTX 生成引擎
│   └── requirements.txt       # Python 依赖
├── examples/
│   └── product-launch/
│       └── outline.json       # 产品发布示例大纲
├── docs/                      # 文档截图
├── CHANGELOG.md
├── README.md
└── LICENSE
```

---

## 🛠️ 自定义

### 配色方案

在 `assets/themes/` 中添加新的 yaml 文件：

```yaml
# my-theme.yaml
primary: [0x33, 0x33, 0x66]
accent: [0xFF, 0x66, 0x00]
bg: [0xF5, 0xF5, 0xFA]
text: [0x22, 0x22, 0x44]
text_light: [0x88, 0x88, 0xAA]
dark_bg: [0x22, 0x22, 0x44]
```

### 添加新版式

1. 在 `scripts/generate.py` 中添加 `make_xxx()` 函数
2. 注册到 `LAYOUTS` 字典
3. 更新 `references/layouts.md` 中的描述
4. 更新 `references/layout-registry.yaml` 中的注册

---

## 📝 许可证

MIT
