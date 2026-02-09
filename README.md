# SmartPPT - 讲师式知识点分页引擎（V2）

当前版本目标：从“字数分页”升级为“知识点分页”，并让你**不用反复手动调试**。

## 现在支持什么

- 上传后立即解析 `.docx` 文本并分页（`.doc` 标记 unsupported）。
- 单次批量上传最多 50 份（适合你 20~50 份验收）。
- 输出结构化 JSON：`page_type/topic/bullets/quotes/evidence/quality_score`。
- 页面和结果会显示运行元信息：`engine_version/git_sha/build_time`，方便确认是否是最新版本。
- 上传后可直接下载：`latest_result.json` + `latest_report.txt`（不依赖页面刷新状态）。

## 最省心的使用方式（避免不断调试）

1. 启动服务：`python word_upload_demo.py --open-browser`
2. 上传文档后，直接下载 `latest_result.json` 和 `latest_report.txt`
3. 看 JSON 顶部 `metadata.git_sha`，确认是不是你期望的提交

> 这样就不用猜“是不是旧进程/旧页面”。

## V2 分页规则（更细腻）

1. **标题强切页**
   - `一、二、三...`、`1.`、短标题冒号（如 `建安风骨：...`）会强制新页。
2. **人物切换切页**
   - 检测 `X作为... / X是... / X则是...`（X 为 2~3 字中文名），切成独立人物页。
3. **诗句/引文单独页**
   - 命中引号诗句或短句诗行模式，优先拆成 `quote` 页。
4. **字数/要点兜底**
   - 超过阈值会自动续页（不再作为主策略，只兜底）。
5. **长句自动拆 bullet**
   - 先按句号/问号/分号切，再按逗号细分，减少过长单条 bullet。

## 输出结构（示例）

```json
{
  "metadata": {
    "engine_version": "v2-knowledge-point",
    "git_sha": "abc1234",
    "build_time": "2026-02-05T12:00:00+00:00"
  },
  "total_files": 1,
  "results": [
    {
      "file": "sample.docx",
      "status": "ok",
      "pages": [
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
      ]
    }
  ]
}
```

## 如何上传你“已分段/带注释”的 20~50 份文档

1. 打开页面 `http://localhost:8000`
2. 文件框一次多选（Ctrl/Shift 或 macOS Command）
3. 直接提交（上限 50）

建议：
- 你的“分段标题”尽量用 Word 标题样式（Heading1/2/3）或规范标题文本（如“一、...”），会显著提升分页质量。
- 批注/修订本版不专门抽取，主要用于正文知识点分页。

## 调参（只改配置思路）

`word_upload_demo.py` 的 `EngineConfig` 可调：

- `max_chars_per_page`：单页最大字数
- `max_bullets_per_page`：单页 bullet 上限
- `short_title_char_limit`：短标题冒号识别阈值
- `max_bullet_chars`：单条 bullet 最长字符数

## 自动化验收（Golden Test）

已在测试中加入固定样本约束，检查：

- 页数下限
- 每页字数上限
- 结构字段完整性
- 平均质量分数下限

并提供 GitHub Actions：每次 PR 自动执行 `python3 -m unittest -v`。

## 运行

### 浏览器优先（Codespaces / 本地都可）

```bash
python word_upload_demo.py --open-browser
```

### Windows 双击

双击 `start_demo.bat`。

## 测试

# SmartPPT - Word 上传最小演示单元

你现在的核心诉求是：**不要本地 CMD/终端，点开就能用**。下面给你 3 个路径（从最省事到正式产品化）。

## 方案一（最推荐）：GitHub Codespaces（完全不碰本地 CMD）

一句话：**在浏览器里启动 Python 服务并打开上传页**。

### 你会得到什么
- 不需要本地安装 Python
- 不需要打开你电脑上的 CMD / PowerShell
- 全程在 GitHub 网页里完成

### 操作步骤
1. 打开本仓库 GitHub 页面。
2. 点击 `Code` → `Codespaces` → `Create codespace on work`（或你的工作分支）。
3. 等待 1～2 分钟进入网页版 VS Code。
4. 在 **浏览器里的终端**（不是你本地 CMD）执行：

```bash
python word_upload_demo.py --open-browser
```

5. Codespaces 会提示端口转发（8000），点击 `Open in Browser` 即可看到上传页。

> 本项目 demo 基于 Python 标准库，无需额外 `pip install`。

---

## 方案二（本机“点开就用”）：双击 `start_demo.bat`（Windows）

如果你希望在本机上也尽量“零命令”：

1. 确保电脑安装了 Python 3。
2. 在项目目录双击 `start_demo.bat`。
3. 它会自动：
   - 启动上传服务
   - 自动打开浏览器到上传页面

> 这仍然是本机运行，但你不需要手动输入命令。

---

## 方案三（正式产品化）：部署在线地址（Render/Railway/Fly.io）

一句话：把 demo 部署成公网网页，任何人直接打开链接上传。

适合场景：
- 给非技术同事使用
- 对外演示
- 不希望每次都启动本地服务

建议下一步：我可以直接帮你补一份最小部署配置（例如 Render），让你一键发布。

---

## 本地手动方式（保留给开发者）

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

默认地址为 `http://localhost:8000`。

---

## 上传与测试

- 页面选择 `.doc` / `.docx` 上传，成功后文件会写入 `uploads/`。
- 单元测试（可选）：

```bash
python3 -m unittest -v
```
