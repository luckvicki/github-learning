from pathlib import Path

input_folder = Path("input")

excel_files = list(input_folder.glob("*.xlsx"))

print("Excel 文件扫描工具")
print("------------------")
print(f"共发现 {len(excel_files)} 个 Excel 文件")

for file in excel_files:
    size_kb = file.stat().st_size / 1024
    print(f"- {file.name} ({size_kb:.2f} KB)")