# SmartPPT - Word 上传最小演示单元

这是一个最小可运行 demo，用于先验证「能否上传 Word 文档」。

## 先解释你在 GitHub 页面看到的问题

你看到仓库页面还是：

- `Add a README`
- `No releases published`

通常是因为你当前查看的是 **main 分支**，而本次改动在工作分支（例如 `work`）上，还没有合并到 main。

### 怎么确认

1. 在 GitHub 左上角分支下拉框，切换到 `work`（或你的 PR 分支）看是否能看到 `README.md`。
2. 打开 Pull Request，确认是否已经 **Merge** 到 main。
3. 合并后再切回 main，README 就会显示在仓库首页。

> `Releases` 是“版本发布”功能，和是否能打开上传网页不是一回事；不创建 release 也完全可以本地跑起来测试。

## 1) 运行服务

### Windows（推荐先看）

你刚才输入的是：

```cmd
Python 3 word_upload_demo.py
```

这里 `Python` 和 `3` 被当成了两个参数，所以 Python 会把 `3` 当作文件名，才会出现：

- `can't open file 'C:\\Users\\...\\3'`

请改为下面任意一种（**中间不要有空格 `Python 3`**）：

```cmd
python word_upload_demo.py
```

或：

```cmd
py -3 word_upload_demo.py
```

想要自动弹出浏览器窗口（更省事）：

```cmd
python word_upload_demo.py --open-browser
```

如果你当前目录不是项目目录，请先切到脚本所在目录再执行：

```cmd
cd /d D:\your\project\smartppt
python word_upload_demo.py
```

### macOS / Linux

```bash
python3 word_upload_demo.py
```

自动打开浏览器：

```bash
python3 word_upload_demo.py --open-browser
```

默认监听 `http://localhost:8000`。

## 2) 浏览器测试

打开：

- `http://localhost:8000`

选择一个 `.doc` 或 `.docx` 文件并上传。上传成功后，文件会写入本地 `uploads/` 目录。

## 3) 命令行快速测试（可选）

Windows PowerShell：

```powershell
curl -F "file=@C:/path/to/your/test.docx" http://localhost:8000
```

macOS / Linux：

```bash
curl -F "file=@/path/to/your/test.docx" http://localhost:8000
```

## 4) 单元测试

Windows：

```cmd
python -m unittest -v
```

macOS / Linux：

```bash
python3 -m unittest -v
```

当前单元测试覆盖：文件扩展名校验函数 `is_allowed_word_file`。
