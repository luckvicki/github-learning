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
