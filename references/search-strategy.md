# Search Strategy

Use online retrieval only after the user confirms that public information search is allowed.

## User-Facing Rule

Do not expose tool names, MCP names, CLI commands, scripts, API keys, or installation details to the user.

Say:

“我可以在你确认后检索最新公开资料，作为调研假设和待现场核实问题的参考。”

Do not say:

- “我将使用 AnySearch”
- “请授权执行脚本”
- “请安装搜索工具”
- “请提供 API Key”
- “我将运行 CLI 命令”

## Internal Tool Priority

Use the best available online retrieval capability in the current environment:

1. Prefer an already-available AnySearch MCP/tool interface if the environment exposes one.
2. If no AnySearch MCP/tool interface is available, use the environment's built-in web/search capability.
3. Do not ask the user to install AnySearch or any extra local tool.
4. Do not require the user to provide an API key.
5. Avoid local CLI/script-based search flows when they would trigger user-visible execution approvals for technical commands.
6. If no online retrieval capability is available, tell the user in business terms: “当前环境暂时无法联网检索，请提供本地资料或稍后重试。”

## Search Query Groups

For a region and industry, search in parallel where possible:

- `{区域} {产业} 产业规划`
- `{区域} {产业} 政策`
- `{区域} {产业} 重点企业`
- `{区域} {产业} 园区`
- `{区域} {产业} 产业集群`
- `{区域} {产业} 招商`
- `{区域} {产业} 统计公报`
- `{区域} {产业} 产业链`
- `{区域} {产业} 数字化转型`
- `{区域} {产业} 产业大脑`

## Freshness Rules

- Prefer the newest official source.
- For policies, projects, enterprise lists, and statistics, prioritize the latest 1-3 years.
- Do not rely on old data when newer sources are available.
- If a source has no clear date, mark it as weak evidence.

## Source Priority

1. Local government portals, department websites, official planning documents, statistical bulletins.
2. Park, association, and industry alliance websites.
3. Enterprise official websites and announcements.
4. Authoritative media and industry platforms.
5. General web pages and uncited summaries.

## Extraction Rules

Read full page content for important pages instead of relying only on snippets:

- policy documents
- planning documents
- statistical bulletins
- park introductions
- government news about key projects
- official lists of enterprises or platforms

## Output Handling

When public search is used, separate:

- 已掌握公开信息
- 初步判断
- 待现场核实事项
- 建议补充材料

Include source dates or publication years for key numbers.
