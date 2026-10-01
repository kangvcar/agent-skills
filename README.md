# agent-skills

给 Claude Code 用的 agent skills 合集。

## 安装

不需要先装什么，`npx` 会自己拉 CLI：

```bash
# 列出仓库里有哪些 skill
npx skills add kangvcar/agent-skills --list

# 装某一个（-g 装到用户级 ~/.claude/skills/，不加则装进当前项目）
npx skills add kangvcar/agent-skills --skill linux-lab-report -g -a claude-code

# 全装
npx skills add kangvcar/agent-skills --all
```

装完重启 Claude Code，当你的请求命中 skill 的 description 时它会自动加载。也可以用
`/linux-lab-report` 手动触发。

更新已装的 skill：

```bash
npx skills check     # 看看有没有新版
npx skills update
```

## 仓库结构

```
skills/
└── <skill-name>/
    ├── SKILL.md          # 必需。YAML frontmatter 带 name 和 description
    ├── assets/           # 模板等资源
    ├── scripts/          # 脚本
    ├── references/       # 供 agent 深入阅读的参考文档
    └── evals/            # 评测用例与打分器
```

## 包含的 skill

### linux-lab-report

为《Linux操作系统配置与管理》课程编写实验报告（实验实训 docx）。

面向高职大数据专业大二上学期、刚接触命令行的学生，三条设计主线：**每步给 Windows 对照**、
**每步自检**、**每站一张截图**。在 `assets/template.docx` 上原位改写，保留外壳表格、
页眉与样式表，生成的文件直接可用。

依赖 `python-docx`：

```bash
pip install python-docx
```

自带 19 条程序化断言（`evals/grade.py`），生成后可逐项核对，详见
[`skills/linux-lab-report/evals/README.md`](skills/linux-lab-report/evals/README.md)。

## 说明

- 仓库里的 skill 会被 `npx skills add` 原样复制到目标目录，**脚本不会被执行、依赖不会被安装**。
  需要额外依赖的 skill 会在自己的 `SKILL.md` 里写清怎么装。
- 用 `skills` CLI 安装时默认开启匿名遥测（采集 skill 名、文件内容、时间戳）。
  介意的话设 `DISABLE_TELEMETRY=1`。
