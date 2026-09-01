# Excel 文件扫描工具

## 功能

扫描 `input` 文件夹中的 `.xlsx` 文件，并输出：

- Excel 文件总数；
- 每个文件的文件名和大小；
- 每个文件包含的工作表（Sheet）数量。

如果某个文件损坏、无法读取或格式异常，程序会单独提示该文件读取失败，并继续处理其他文件。如果 `input` 文件夹中没有 `.xlsx` 文件，程序也会给出明确提示。

## 当前版本

v1.1

## 环境要求

- Python 3.8 或更高版本
- `openpyxl`

## 安装依赖

在 `excel-tool` 目录中运行：

```bash
python -m pip install -r requirements.txt
```

## 使用方法

1. 将需要扫描的 `.xlsx` 文件放入 `input` 文件夹；如果该文件夹不存在，请在 `excel-tool` 目录中创建它。
2. 在 `excel-tool` 目录中运行：

   ```bash
   python app.py
   ```

3. 在终端中查看文件数量、文件名、文件大小和 Sheet 数量。

输出示例：

```text
Excel 文件扫描工具
------------------
共发现 1 个 Excel 文件

文件：测试数据.xlsx
大小：12.35 KB
Sheet数量：3
```
## 自动化测试

本项目使用 GitHub Actions 自动执行 Python 检查和单元测试。

每次向 `main` 分支创建 Pull Request 时，GitHub 会自动执行以下检查：

1. 安装 Python 3.11；
2. 安装 `requirements.txt` 中的依赖；
3. 使用 `py_compile` 检查 `app.py` 是否存在 Python 语法错误；
4. 使用 `unittest` 运行 `excel-tool/tests` 目录中的自动化测试。

当前自动化测试主要验证：

- 是否能够正确识别 `.xlsx` 文件；
- 是否能够正确读取 Excel 工作表（Sheet）数量；
- 损坏或异常的 Excel 文件是否不会影响其他文件继续处理。

只有 Required Checks 通过后，Pull Request 才允许合并到 `main`。
