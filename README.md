# SmartPPT - Word 上传最小演示单元

这是一个最小可运行 demo，用于先验证「能否上传 Word 文档」。

## 1) 运行服务

```bash
python3 word_upload_demo.py
```

默认监听 `http://localhost:8000`。

## 2) 浏览器测试

打开：

- `http://localhost:8000`

选择一个 `.doc` 或 `.docx` 文件并上传。上传成功后，文件会写入本地 `uploads/` 目录。

## 3) 命令行快速测试（可选）

```bash
curl -F "file=@/path/to/your/test.docx" http://localhost:8000
```

## 4) 单元测试

```bash
python3 -m unittest -v
```

当前单元测试覆盖：文件扩展名校验函数 `is_allowed_word_file`。
