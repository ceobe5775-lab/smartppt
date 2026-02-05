# SmartPPT - Word 上传与分页解析（V1 无训练版）

这个版本已经支持：

1. 上传 Word 文档后立即解析文本（V1）
2. 输出分页结果 JSON（每页 `title` / `content` / `char_count` / `page_no`）
3. 单次上传窗口支持多文件（最多 50 份，可用于 20~50 份验收）

## 训练是否必须？

- **V1 不需要训练**：使用规则提取+分页即可跑通验收。
- 当你后续要做“语义分页、风格统一、行业模板强约束”时，再考虑模型微调。

## 方案一（推荐）：GitHub Codespaces（不碰本地 CMD）

1. 打开仓库 → `Code` → `Codespaces` → `Create codespace`。
2. 在浏览器里的终端运行：

```bash
python word_upload_demo.py --open-browser
```

3. 打开转发端口 `8000`，即可看到上传页。

## 方案二（Windows 双击启动）

双击 `start_demo.bat`，会自动启动服务并打开浏览器。

## 本地手动运行（开发者）

### Windows

```cmd
python word_upload_demo.py --open-browser
```

或：

```cmd
py -3 word_upload_demo.py --open-browser
```

### macOS / Linux

```bash
python3 word_upload_demo.py --open-browser
```

## 使用说明（20~50 份验收）

1. 打开页面 `http://localhost:8000`
2. 点击文件框，按住 `Ctrl`/`Shift`（macOS 用 `Command`）多选文档
3. 一次上传 20~50 份（上限 50）
4. 页面会展示最近一次解析 JSON
5. 同时结果会落盘到 `outputs/latest_result.json`

## 输出 JSON 结构（示例）

```json
{
  "total_files": 2,
  "results": [
    {
      "file": "sample.docx",
      "status": "ok",
      "total_pages": 2,
      "total_chars": 1234,
      "pages": [
        {
          "page_no": 1,
          "title": "第一章",
          "content": "...",
          "char_count": 580
        }
      ]
    }
  ]
}
```

## 当前 V1 边界

- `.docx`：支持解析与分页。
- `.doc`：会接收，但 V1 仅标记为 `unsupported`（建议先转 `.docx`）。

## 常见问题

### 1) 为什么之前会报 `UnicodeEncodeError`？

原因是上传后 303 跳转的 `Location` 头里包含了中文提示文本。`http.server` 在发送 header 时按 latin-1 编码，中文会触发编码错误。

本版本已修复：对跳转查询参数做 URL 编码，保证 `Location` 头只包含 ASCII。

### 2) 我这 20~50 份文档是“做过分段和注释”的，怎么上传？

直接在上传窗口多选后一次提交即可（支持最多 50 份）。

- 你做过的“分段”如果体现在 **标题样式（Heading1/2/3）**，V1 会把它识别成分页标题。
- 普通正文按长度自动分页。
- 批注/修订这类注释信息，V1 目前不做专门提取（先以正文分页验收为主）。

## 测试

```bash
python3 -m unittest -v
```
