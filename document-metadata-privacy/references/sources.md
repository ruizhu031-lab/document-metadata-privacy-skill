# 查证来源

查阅日期：2026-10-21。优先采用厂商和项目官方文档。

1. [Microsoft：检查并移除文档隐藏数据与个人信息](https://support.microsoft.com/en-gb/office/collab-files/remove-hidden-data-and-personal-information-by-inspecting-documents-presentations-or-workbooks?nochrome=true)
   文档属性以外还包括批注、修订、隐藏内容等；建议在副本上操作，检查器也有检测/移除限制。
2. [Microsoft Open XML：扩展文件属性定义](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.extendedproperties.properties?view=openxml-3.0.1)
   Application、AppVersion、Company、Manager、Template、TotalTime 等是独立属性，清空作者并不能清除全部软件与身份属性。
3. [WPS：Document.RemovePersonalInformation](https://open.wps.cn/documents/app-integration-dev/wps365/client/wpsoffice/jsapi/wps/Document/member/RemovePersonalInformation)
   WPS 文字的文档对象可设置保存时删除用户信息。适用范围是文档对象接口，不能据此推定所有 WPS 产品具有相同接口。
4. [pypdf：Metadata](https://pypdf.readthedocs.io/en/stable/user/metadata.html)
   PDF 可同时拥有普通属性与 XMP。`writer.metadata = None` 用于移除 Info，XMP 需另外处理。脚本使用已安装版本验证 API 与结果。
5. [OpenAI：Build skills](https://learn.chatgpt.com/docs/build-skills)
   SKILL.md 的名称和描述决定匹配；隐式调用可由 agents/openai.yaml 控制。安装路径以当前运行环境可发现的用户技能目录为准。
6. [OpenAI：AGENTS.md](https://learn.chatgpt.com/docs/agent-configuration/agents-md)
   用户级指令可为未来任务设定默认行为。自动选择技能是代理遵循规则，不是对所有写文件操作的后台拦截。

工程结论：先完成保存/导出，再清理最终副本并复查，能减少文件中不必要的身份与软件信息；不能据此证明人工创作，也不能保证无法被识别、关联或推断。
