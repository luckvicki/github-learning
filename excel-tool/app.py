from pathlib import Path

from openpyxl import load_workbook


input_folder = Path(__file__).resolve().parent / "input"
excel_files = sorted(input_folder.glob("*.abc"))

print("Excel 文件扫描工具")
print("------------------")
print(f"共发现 {len(excel_files)} 个 Excel 文件")

if not excel_files:
    print("提示：input 文件夹中没有 .xlsx 文件。")

for file in excel_files:
    size_kb = file.stat().st_size / 1024
    print()
    print(f"文件：{file.name}")
    print(f"大小：{size_kb:.2f} KB")

    try:
        workbook = load_workbook(file, read_only=True, data_only=True)
        try:
            sheet_count = len(workbook.sheetnames)
        finally:
            workbook.close()
        print(f"Sheet数量：{sheet_count}")
    except Exception as error:
        print(f"读取失败：{error}")
