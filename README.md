# document-metadata-privacy-skill
some good skill in my studying

个人学习与使用中的 Codex 技能集合。

## 技能列表

| 技能 | 用途 | 支持范围 |
| --- | --- | --- |
| [document-metadata-privacy](document-metadata-privacy/README.md) | 交付前清理非必要属性，应用个人作者和时间默认值，备份验证后保留原文件名 | Word、PowerPoint、Excel、WPS 保存的 DOCX/PPTX/XLSX，以及普通 PDF |

### 文档元数据隐私清理

适合对外发送报告、演示文稿或工作簿之前使用，也可与 Codex 的文档生成技能配合，在最后一次保存或导出后执行。工具在本地运行，保留原件；不上传被处理的文档。清理元数据不保证匿名、无法识别 AI 或完全无痕。

- [用途、安装步骤和命令行示例](document-metadata-privacy/README.md)
- [Codex 技能入口](document-metadata-privacy/SKILL.md)
- [Office / WPS / PDF 的格式说明](document-metadata-privacy/references/formats.md)
- [官方参考来源](document-metadata-privacy/references/sources.md)

安装后可在 Codex 中输入：

```text
使用 $document-metadata-privacy 检查并清理这份文档的属性，保留内容并报告残留。
```

## 开源许可证

本仓库采用 [MIT License](LICENSE)，版权归 `ruizhu031-lab`。允许使用、复制、修改、分发和商用，需保留版权与许可证声明。软件按原样提供。
