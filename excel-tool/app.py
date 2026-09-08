from pathlib import Path

from openpyxl import load_workbook
import xlrd


SUPPORTED_EXTENSIONS = {".xlsx", ".xls"}


def scan_excel_files(input_folder):
    """返回 input 目录中的所有 Excel 文件。"""
    input_folder = Path(input_folder)
    if not input_folder.exists():
        return []

    return sorted(
        file
        for file in input_folder.iterdir()
        if file.is_file() and file.suffix.lower() in SUPPORTED_EXTENSIONS
    )


def get_sheet_count(file_path):
    """读取 Excel 文件并返回工作表数量。"""
    file_path = Path(file_path)
    suffix = file_path.suffix.lower()

    if suffix == ".xlsx":
        workbook = load_workbook(file_path, read_only=True, data_only=True)
        try:
            return len(workbook.sheetnames)
        finally:
            workbook.close()

    if suffix == ".xls":
        workbook = xlrd.open_workbook(str(file_path))
        return len(workbook.sheet_names())

    raise ValueError("不支持的 Excel 文件格式：仅支持 .xlsx 和 .xls 文件。")


def main(input_folder=None):
    """扫描并输出 Excel 文件信息。"""
    if input_folder is None:
        input_folder = Path(__file__).resolve().parent / "input"

    excel_files = scan_excel_files(input_folder)

    print("Excel 文件扫描工具")
    print("------------------")
    print(f"共发现 {len(excel_files)} 个 Excel 文件")

    if not excel_files:
        print("提示：input 文件夹中没有 .xlsx 或 .xls 文件。")

    for file in excel_files:
        size_kb = file.stat().st_size / 1024
        print()
        print(f"文件：{file.name}")
        print(f"大小：{size_kb:.2f} KB")

        try:
            sheet_count = get_sheet_count(file)
            print(f"Sheet数量：{sheet_count}")
        except Exception:
            print("读取失败：该文件可能已损坏，或不是有效的 Excel 文件。")


if __name__ == "__main__":
    main()
