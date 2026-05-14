<!-- prettier-ignore -->
[English](README.md) · **简体中文**

# docx-master

一个让 AI agent 真正能处理 Word 文档的 skill。不绕道 Markdown 或 HTML，直接读写 `.docx` 的 OOXML，所以样式、编号、字段、修订这些东西都不会在中间丢失。

> 快速上手：从 [Releases](https://github.com/hawa130/docx-master/releases) 下载对应 harness 的 zip，丢进 skills 目录即可。

## 同类 Word skill

其他 Word skill 大多是某个底层库（python-docx / docx-js / OpenXML SDK 等）的 LLM 使用文档：交给 agent 一组操作原语，「合格的 Word 文档该长什么样」是 agent 自己想清楚的事。

docx-master 自带一套 Word 文档的写法约定：

- 每个段落都绑命名样式（标题、正文、列表等），格式跟着样式走
- 章节、图、表、公式都走 Word 内置的自动编号
- 「见图 1.2」这类引用是活的，章节顺序变了所有编号一起更新
- 中英混排时中文字号和英文字号互不干扰

每条约定都有配套工具：`pattern_rules` / `bulk_rules` 按规则一次改整篇，`migrate_captions` 探测旧稿里手打的图表编号，`audit` 提前扫违规，12 个 inspect / find 工具让 agent 改之前先看清楚。

下面对比的几个同类 skill：

- [**anthropics/docx**](https://skills.sh/anthropics/skills/docx) — Anthropic 官方。unpack + 改 XML + docx-js 创建，通用 Word 编辑。
- [**qodex-ai/word-document-processor**](https://skills.sh/qodex-ai/ai-agent-skills/word-document-processor) — pandoc + docx-js + python-docx + 直接 XML 的组合工具箱，侧重 redlining 工作流。
- [**minimax-ai/minimax-docx**](https://skills.sh/minimax-ai/skills/minimax-docx) — C# + OpenXML SDK（.NET），侧重结构化编辑。
- [**claude-office-skills/docx-manipulation**](https://skills.sh/claude-office-skills/skills/docx-manipulation) — python-docx 封装，侧重模板占位符替换。

具体能力对比：

| 场景 | anthropics | qodex-ai | minimax-ai | claude-office | docx-master |
|---|:---:|:---:|:---:|:---:|:---:|
| 把内容填进已有的空模板 | ~ | ~ | ~ | ✓ | ✓ |
| 章节标题自动编号（不用手打 1.1.1） | — | — | — | — | ✓ |
| 中英文混排时字体互不覆盖 | — | — | — | — | ✓ |
| 图、表、公式按「章号.内号」编号 | — | — | — | — | ✓ |
| 正文「见图 1.1」是活字段，重排自动更新 | — | — | — | — | ✓ |
| LaTeX 公式渲染为编号居中独立块 | ~ pandoc | ~ pandoc | — | — | ✓ |
| 审阅意见用 tracked changes 来回传 | ✓ | ✓ | ~ | ~ | ✓ |
| 旧稿手打的「图 2.1」转成活字段 | — | — | — | — | ✓ |
| 表单空白填值，不破坏底线 | — | — | — | — | ✓ |
| 改表格里某个单元格的内容 | — | — | ✓ | ✓ | ✓ |
| 格式合规检查 | — | — | — | — | ✓ |
| 借其他文档的样式套到当前稿子 | — | — | — | — | ✓ |

> ✓ 内置 helper，声明即用 · ~ 能做，但 agent 需要自己拼 · — 没有专门支持

日常创建和简单编辑，anthropics/docx 就够了。redlining 流程密集的场景 qodex-ai 更合手，.NET 项目走 minimax-ai 自然，单纯做模板占位符填充 claude-office-skills 最短平快。docx-master 真正的位置在长文档（特别是中文）的结构化重排。

docx-master 完整的子命令、Block 类型、reference 文档清单见下面的「包含什么」。

## 包含什么

### Skill：docx-master

专注 Word 自动化的一个 skill，附带 10 份按需加载的参考文档（[查看 skill](skill/SKILL.md)）：

| 参考文档 | 包含 |
|---|---|
| [standardize.md](skill/references/standardize.md) | 整篇重排：`styles[]`、`numbering`、`pattern_rules`、`bulk_rules`、`assignments`、`exclude` |
| [edit.md](skill/references/edit.md) | 精修编辑：定位器、操作类型（replace / insert / delete / set-run）、MDF、tracked changes |
| [config-schema.md](skill/references/config-schema.md) | `apply` 配置完整字段参考 |
| [captions.md](skill/references/captions.md) | 基于 SEQ 的图表标题、章号前缀编号、REF 交叉引用 |
| [cross-references.md](skill/references/cross-references.md) | `InlineRef` schema：图 / 表 / 公式 / 章节引用 |
| [numbering-formats.md](skill/references/numbering-formats.md) | 多级编号形式：十进制、带括号、中文序号、项目符号 |
| [tables.md](skill/references/tables.md) | `edits[]` 插入表格的 Block schema |
| [equations.md](skill/references/equations.md) | LaTeX → OMML 行内 / 块级公式，编号公式布局 |
| [chinese-font-sizes.md](skill/references/chinese-font-sizes.md) | 小四 / 五号 / 三号 / … 到 half-points 的换算 |
| [audit.md](skill/references/audit.md) | 只读合规扫描流程 + 扫描维度 |

### 15 个工具

所有工具都用 `node scripts/<name>.js <args>` 调用，输出到 stdout。Skill 的 prompt 会引导 agent 选合适的工具。

| 工具 | 何时使用 |
|------|----------|
| `overview` | 任何任务的第一次调用。元数据、页面设置、主题、样式定义、编号方案、指纹统计、文档骨架 |
| `inspect_range` | 指定范围段落的完整文本 + 计算后样式 |
| `inspect_runs` | 单段逐 run 的 rPr dump，用在混合格式段落或表单空白段 |
| `inspect_neighbors` | 看某一段周围有什么，用来判断它的角色（图表标题 / 紧跟标题后 / 等等） |
| `inspect_style` | 某个指纹在整篇里扮演什么角色 |
| `inspect_style_def` | `styles.xml` 里预定义的样式 + `basedOn` 链路 |
| `inspect_section` | 各分节之间页面设置的差异 |
| `inspect_table` | 顶层表格 + 单元格文本 + 段落索引区间 |
| `inspect_blockers` | edit 阶段会拒绝的段落（tracked changes、复杂域、SDT） |
| `inspect_caption` | 文档里基于 SEQ 的图表标题；列出每个 identifier、逐项 dump |
| `migrate_captions` | 只读探测「手打编号的图表段」 |
| `find_paragraphs` | 跨文档正则搜索，用来验证 `pattern_rules` 的覆盖面 |
| `find_text` | 字符级定位：段落索引、run 索引、字符偏移、上下文预览 |
| `validate` | 对任意 `.docx` 做 schema-aware 的 OOXML 校验 |
| `apply` | 统一写入器。`--dry-run` 用来迭代，去掉 flag 才会真正写盘 |

### `apply` 流水线

```
装 styles + numbering + theme + template
  → 执行 edits（可以引用刚装好的 styleId）
  → 重新计算指纹
  → 执行 rules（pattern_rules / bulk_rules / assignments / exclude ——
    既覆盖原有 chrome，也覆盖 agent 插入的新内容）
  → 校验、写盘
```

只有被声明的块才会生效，其它部分原封不动。

## 设计原则

工具只暴露可见事实，分类和判断交给 agent。默认输出不做预分类，agent 看到的是人类读者也能看到的属性，做完初步判断后再按需读隐藏元数据。

机械层面的正确性由脚本兜底。段落遍历、XML 命名空间、跨 run 格式保留、`numId` 冲突避免、blocker 检测、校验，全部不交给 agent，也不会因为代码重构而松动。

要校验的是意图，不是系统自己对输入的解读。拿系统的解读去判定系统的输出，等于没校验。真正的校验要么是人能读的对照（原文对解析后字段），要么是把输出用独立的不变量重新解析。

原文件不动。每次写入产出新文件，校验通过才保留；失败就丢弃、上报，不做静默重试。

## 安装

从 [Releases](https://github.com/hawa130/docx-master/releases) 下载对应 harness 的 zip，解压到该 harness 加载 skills 的目录即可。每个 release 都会发布以下文件：

| 文件 | 解压后路径 | 适用 |
|---|---|---|
| `docx-master.zip` | `docx-master/` | 通用，适用任何加载 Markdown skill 的 harness |
| `docx-master-claude-code.zip` | `.claude/skills/docx-master/` | Claude Code |
| `docx-master-cursor.zip` | `.cursor/skills/docx-master/` | Cursor |
| `docx-master-codex.zip` | `.agents/skills/docx-master/` | Codex CLI |
| `docx-master-gemini.zip` | `.gemini/skills/docx-master/` | Gemini CLI |
| `docx-master-opencode.zip` | `.opencode/skills/docx-master/` | OpenCode |
| `docx-master-github.zip` | `.github/skills/docx-master/` | GitHub Copilot |

选对应 harness 的 zip，在合适的范围里解压：

```bash
# Claude Code，用户级
cd ~ && unzip ~/Downloads/docx-master-claude-code.zip

# Claude Code，项目级
cd your-project && unzip ~/Downloads/docx-master-claude-code.zip

# 或用通用 bundle，丢到任意 harness 的 skills 目录
unzip ~/Downloads/docx-master.zip -d ~/.claude/skills/
```

Harness 特定注意事项：

- Cursor：切换到 Nightly 通道（Settings → Beta），并启用 Agent Skills（Settings → Rules）。[文档](https://cursor.com/docs/context/skills)。
- Gemini CLI：`npm i -g @google/gemini-cli@preview`，`/settings` 启用 "Skills"，用 `/skills list` 验证。[文档](https://geminicli.com/docs/cli/skills/)。

### 运行时依赖

Skill 内带的脚本是打包好的 Node CJS。只要 harness 能在用户机器上跑 `node`（Node 18+），docx-master 就能跑。不需要 Python、不需要 pip install、没有系统级依赖，OOXML schema 和 xmllint-wasm 校验器都已经打包进 bundle。

## 使用

prompt 里只要涉及 Word 文档任务，skill 会自动触发。一些日常场景：

> 把这份学校发的开题报告补全。中文宋体、英文 Times New Roman，正文小四，
> 首行缩进 2 字符，图表编号用「章号.章内序号」，所有引用走活字段。

> 审计一下这份学位论文。哪些标题没绑样式、哪些图表编号是手打的、章节
> 重排后哪些引用会失效。

> 把手稿里所有手打的「图 2.1」「表 1.3」转成 SEQ 字段，正文里的「如图
> 2.1 所示」也改成活的 REF 引用。

> 在合同第 42 段后面插入这三条补充条款，开 track changes 让法务在 Word
> 里审。

> 这份表单有「项目名称：____」这种带下划线占位的空格，把「项目名称」后
> 面的空白填上「Q3 营销计划」，别动标签、也别动下划线本身。

> 把这些 LaTeX 公式渲染成带编号的居中独立公式，正文里用活字段引用对应
> 的公式号。

> 借用同事 reference.docx 里的标题样式和图表标题样式，应用到我这份稿子
> 上，但别动我现有的编号方案。

每次写入都会生成一份新副本，通过 schema 校验才保留，原文件不动。编号、图表标题、交叉引用都以活字段形式落地，章节顺序改了，也只需要在 Word 里「更新域」就能保持一致。

## 仓库结构

| 路径 | 作用 |
|---|---|
| `skill/SKILL.md` | 暴露给 agent 的入口约定（顶层路由） |
| `skill/references/` | 按需细节；只在相关时才加载 |
| `skill/tools/` | 15 个 CLI 的 TS 源码 |
| `lib/` | 非 tool 类的 TS 模块（xml / parse / config / apply / edit / shared） |
| `test/fixtures/` | 用来人工验证的 `.docx` 样例 |
| `build-skill.ts` | 打包 `dist/docx-master/` + zip + 多 harness fan-out |
| `CLAUDE.md` | 给贡献者 / 后续 agent 看的项目工作指南 |

## 从源码构建

```bash
bun install
bun run build:skill   # → 产物全部在 dist/ 下
bun run typecheck
bun run lint
bun run fmt:check
```

`bun run build:skill` 在 `dist/` 下产出：

- `docx-master/`：打包好的 skill bundle（`SKILL.md` + `references/` + 编译后的 `scripts/`）
- `docx-master.zip`：通用发布产物
- `<provider>/<configDir>/skills/docx-master/`：每个支持的 harness 一份（`claude-code/.claude/`、`cursor/.cursor/`、`codex/.agents/`、`gemini/.gemini/`、`opencode/.opencode/`、`github/.github/`）
- `plugin/skills/docx-master/`：给 `.claude-plugin/marketplace.json` 引用，用于本地测试 Claude Code marketplace 安装

`dist/` 整体在 gitignore 里，每次构建会重新生成。release 用的各 harness zip 由 [release workflow](.github/workflows/release.yml) 打包。

目前没有自动化测试。改动后用 `test/fixtures/*.docx` 人工跑一遍，并检查产出的 bundle。

## 支持的 harness

- [Claude Code](https://claude.com/claude-code)：主要目标，支持 plugin 形式安装
- [Cursor](https://cursor.com)
- [Codex CLI](https://github.com/openai/codex)
- [Gemini CLI](https://github.com/google-gemini/gemini-cli)
- [OpenCode](https://opencode.ai)
- [GitHub Copilot](https://github.com/features/copilot)

其他能加载「带 YAML frontmatter 的 Markdown skill」的 harness，把通用 bundle 直接放进 skills 目录就能用，不需要任何转换。

## 贡献

仓库约定、skill 内容遵循的设计原则、引擎维持的跨命令不变量，都写在 [CLAUDE.md](CLAUDE.md)。提 PR 时请注意 agent 直接面对的部分（`SKILL.md` + `references/`）有 token 预算，这部分每次调用都会被完整加载，不要无谓地往里加内容。

## 许可证

Apache 2.0。见 [LICENSE](LICENSE)。
