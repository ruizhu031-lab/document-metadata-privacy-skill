# Document Metadata Privacy

用于 Codex 的文档交付前元数据隐私清理技能，覆盖 **Word、PowerPoint、Excel、WPS 和 PDF**。

A local metadata privacy skill for Codex. Clean unnecessary properties, apply user-specified author and time defaults, and preserve the delivery filename after backup and verification.

## 用途

在分享文档前，减少文件属性中不必要的作者、最后修改者、公司、制作软件、版本和编辑时间等信息。适用于对外发送报告、演示文稿、工作簿，以及检查由文档生成库默认写入的属性。

技能默认在**最后一次保存或导出之后、交付链接之前**执行，并与 Word、演示文稿、表格和 PDF 制作技能配合。工具在本地运行，不上传被处理的文档。

清除元数据不等于证明人工创作，也不保证匿名、无法识别 AI 或完全无痕。正文、版式、批注、附件、图片和文件结构仍可能包含来源信息。

## 本仓库的个人默认值

新生成或导出的文档默认作者为 `19436`，最后修改者为 `朱文好`；已有文档按当次请求限定范围处理，正文姓名、学号和邮箱空栏不受影响。

- 按北京时间生成时间：修改日期为操作当日，修改时刻在16:00–22:00随机生成；创建时刻在14:00–20:00随机生成。分钟和秒均取1–59，避开整点及窗口端点。
- 忽略日期，仅比较当日时分秒，两者绝对差至少3小时；完整时间戳同时须满足创建早于修改且至少相隔48小时。创建日期根据内容规模选择，课程报告通常提前3个日历日；不满足条件则重新生成。
- 这些是用户指定的属性时间，不表示真实编辑历史；不据此修改文件系统时间、日志或云端记录。
- 默认保留原文件名，不添加清理后缀。先在独立目录备份和暂存，检查内容及属性，并确认原件未被新的编辑改动后，再更新原交付路径。

`sanitize.py clean` 本身仍会清空目标属性，并且拒绝覆盖已有输出。技能负责在清理后按格式原生方法写入上述默认值并验证；直接运行命令行脚本不会自动应用个人默认值。具体流程以 [SKILL.md](SKILL.md) 为准。

## 支持范围

| 文件 | 自动处理 | 说明 |
| --- | --- | --- |
| `.docx` | 核心属性、自定义属性、指定扩展属性、ZIP 注释和额外字段 | Word / WPS 文字保存的现代格式 |
| `.pptx` | 同上 | PowerPoint / WPS 演示；备注和批注会报告，保留内容 |
| `.xlsx` | 同上 | Excel / WPS 表格；公式和工作表内容保留 |
| `.pdf` | Info 属性、可达对象上的 Metadata/XMP、被丢弃的孤立对象 | 完整重写普通 PDF，避免增量更新保留旧属性 |
| `.wps/.et/.dps`、`.doc/.xls/.ppt` | 不直接处理 | 用原应用实际另存为现代格式后检查，不可直接改扩展名 |
| 宏文件、ODF、RTF、HTML 等 | 格式分流指引 | 见 [formats.md](references/formats.md) |

OOXML 的核心和自定义属性默认全部清空；扩展属性删除 `Application`、`AppVersion`、`Company`、`Manager`、`Template`、`TotalTime`。ZIP 内部条目时间归一化，但不修改源文件或操作系统的文件时间。

工具保留原件，只写入不存在的新输出；对签名、加密、PDF/A、PDF/UA 等情况拒绝直接清理，避免破坏签名或符合性要求。WPS 兼容性目前由文件结构测试验证，尚未覆盖各版本 WPS 客户端实测。

## 安装到 Codex

需要 Python 3.10+ 和 Git。克隆技能集合仓库后，将本技能子目录放入用户级技能目录；新版本 Codex 官方文档使用 `~/.agents/skills`，已有配置使用其他技能目录时可按实际环境调整。若目标技能目录已存在，请先备份并检查差异。

```bash
git clone https://github.com/ruizhu031-lab/zwh_skill.git
mkdir -p ~/.agents/skills
cp -R zwh_skill/document-metadata-privacy ~/.agents/skills/document-metadata-privacy
python -m pip install -r ~/.agents/skills/document-metadata-privacy/requirements.txt
```

Windows PowerShell 示例：

```powershell
git clone https://github.com/ruizhu031-lab/zwh_skill.git
$skillDir = Join-Path $HOME '.agents/skills'
New-Item -ItemType Directory -Force -Path $skillDir | Out-Null
Copy-Item -Recurse -LiteralPath './zwh_skill/document-metadata-privacy' -Destination $skillDir
python -m pip install -r "$HOME/.agents/skills/document-metadata-privacy/requirements.txt"
```

Codex 桌面环境若已有依赖运行时，优先使用其提供的 Python，无需重复安装。如果技能未显示，重新打开 Codex 或开始新聊天。

手动调用：

```text
使用 $document-metadata-privacy 检查并清理这份文档的属性，保留内容并报告残留。
```

`agents/openai.yaml` 已允许隐式调用。若希望作为个人默认交付规则，可将以下内容加入用户级 `~/.codex/AGENTS.md`，保留原有规则：

```markdown
生成、导出或修改并交付 Word、PPT、Excel、WPS、PDF 等文档时，
在最后一次保存/导出后、给出交付链接前，默认调用 document-metadata-privacy 技能。
清理最终副本并复查，报告未处理的残留；用户要求保留的属性优先。
```

隐式调用由代理依据技能描述与指令选择，并非后台文件拦截器。

## 命令行使用

在仓库的 `document-metadata-privacy` 子目录运行：

```bash
# 只读检查，不改变文件
python scripts/sanitize.py inspect /path/to/source.docx

# 输出到一个不存在的新文件，扩展名必须一致
python scripts/sanitize.py clean /path/to/source.docx --output /path/to/final.docx
python scripts/sanitize.py clean /path/to/source.pdf --output /path/to/final.pdf
```

路径包含空格时请加引号。脚本返回 JSON：

- `metadata_fields`：发现的目标属性名称，不输出原作者等属性值。
- `review_flags`：需要进一步复查的批注、修订、备注、附件、媒体等类别。
- `marker_parts`：含常见生成器关键词的 OOXML 部件名，关键词可能是合法正文，不自动删除。
- `remaining`：清理后再次检查的结果。

退出码 `0` 表示命令完成，`2` 表示未完成。`0` 不表示所有隐藏内容都已移除，仍须阅读复查项。PDF 的检查范围主要是 Info 与 Metadata，不扫描全部正文或任意字符串。

若同时交付可编辑文件和 PDF，应分别清理。清理后再次用 Office/WPS 保存或重新导出，可能重新写入属性，需要再跑一次流程。

## 内容与边界

工具保留非属性 OOXML 部件的字节；不会全局替换正文中的 Codex、OpenAI 或 WPS 字样。批注、修订、演讲者备注、嵌入附件及图片元数据不会自动删除。自定义属性可能被动态字段或业务系统引用，处理已有业务文档前应检查依赖。

不会删除云端版本历史、本机日志或外部副本，也不会移除来源凭证或伪造作者、软件名称。更多格式限制与 WPS 使用说明见 [格式分流](references/formats.md)，机制出处见 [官方来源](references/sources.md)。

## 开发与验证

```bash
python -m pip install -r requirements-dev.txt
python -m unittest discover -s tests -v
```

从有写权限的目录运行测试；测试在当前目录创建并清理临时文件。7 项测试覆盖现代 Office/WPS 属性清理、非属性部件保持不变、公式与页面内容保留、只读检查、覆盖保护、签名/加密/符合性保护和旧格式拒绝。PDF 路径已在 pypdf 6.10.0 上验证。

欢迎提交可复现问题和改进。提供样本前请先自行移除敏感内容；优先使用人工构造的最小文件。

## 许可证

[MIT License](LICENSE)。允许使用、修改和分发，需保留许可证与版权声明。
