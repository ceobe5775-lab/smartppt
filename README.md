# SmartPPT - 讲师式知识点分页引擎（V2）

当前版本目标：从“字数分页”升级为“知识点分页”。

## 现在支持什么

- 上传后立即解析 `.docx` 文本并分页（`.doc` 标记 unsupported）。
- 单次批量上传最多 50 份（适合你 20~50 份验收）。
- 输出结构化 JSON：`page_type/topic/bullets/quotes/evidence/quality_score`。

## V2 分页规则（更细腻）

1. **标题强切页**
   - `一、二、三...`、`1.`、短标题冒号（如 `建安风骨：...`）会强制新页。
2. **人物切换切页**
   - 检测 `X作为... / X是... / X则是...`（X 为 2~3 字中文名），切成独立人物页。
3. **诗句/引文单独页**
   - 命中引号诗句或短句诗行模式，优先拆成 `quote` 页。
4. **字数/要点兜底**
   - 超过阈值会自动续页（不再作为主策略，只兜底）。

## 输出结构（示例）

```json
{
  "page_no": 2,
  "page_type": "person_profile",
  "title": "曹操：核心知识点",
  "topic": "曹操",
  "bullets": ["现实关怀", "风格沉郁雄健"],
  "quotes": ["老骥伏枥，志在千里"],
  "evidence": {"signals": ["person_switch", "quote_block"], "source_chunks": [6, 7]},
  "quality_score": 92
}
```

## 如何上传你“已分段/带注释”的 20~50 份文档

1. 打开页面 `http://localhost:8000`
2. 文件框一次多选（Ctrl/Shift 或 macOS Command）
3. 直接提交（上限 50）

建议：
- 你的“分段标题”尽量用 Word 标题样式（Heading1/2/3）或规范标题文本（如“一、...”），会显著提升分页质量。
- 批注/修订本版不专门抽取，主要用于正文知识点分页。

## 调参（一次把引擎设计对）

`word_upload_demo.py` 的 `EngineConfig` 可调：

- `max_chars_per_page`：单页最大字数
- `max_bullets_per_page`：单页 bullet 上限
- `short_title_char_limit`：短标题冒号识别阈值

你可以按课程类型出不同 preset（历史课更细，技术课更密）。

## 运行

### 浏览器优先（Codespaces / 本地都可）

```bash
python word_upload_demo.py --open-browser
```

### Windows 双击

双击 `start_demo.bat`。

## 测试

```bash
python3 -m unittest -v
```
