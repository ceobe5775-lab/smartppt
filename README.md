
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
