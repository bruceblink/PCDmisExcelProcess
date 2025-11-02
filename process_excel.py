import sys
import os
from openpyxl import load_workbook
from datetime import datetime

def log(msg):
    """输出日志并写入文件"""
    print(msg)
    with open("process.log", "a", encoding="utf-8") as f:
        f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")

def get_values(ws, col, max_rows=20):
    """读取一列数据"""
    vals = []
    for i in range(1, max_rows + 1):
        vals.append(ws[f"{col}{i}"].value)
    return vals

def start(backup_path: str, report_path: str):
    log(f"=== 开始执行 ===")
    log(f"备份文件: {backup_path}")
    log(f"报告文件: {report_path}")

    if not os.path.exists(backup_path):
        log(f"❌ 找不到备份文件: {backup_path}")
        return
    if not os.path.exists(report_path):
        log(f"❌ 找不到报告文件: {report_path}")
        return

    wb_backup = load_workbook(backup_path)
    wb_report = load_workbook(report_path, data_only=True)

    # === 1. 启动条件检查 ===
    if "Sheet1" not in wb_backup.sheetnames:
        log("❌ 备份文件中没有 Sheet1")
        return

    ws_sheet1 = wb_backup["Sheet1"]
    check_value = ws_sheet1["F29"].value

    if check_value not in ["李春宁", "刘文"]:
        log("⚠️ 启动条件不满足：Sheet1!F29 不是 '李春宁' 或 '刘文'")
        return
    else:
        log(f"✅ 启动条件通过: {check_value}")

    # === 2. 获取报告表 ===
    if "PCDmisExcel1" not in wb_report.sheetnames:
        log("❌ 报告文件中找不到 PCDmisExcel1 工作表")
        return

    ws_report = wb_report["PCDmisExcel1"]
    log("✅ 找到 PCDmisExcel1")

    # === 3. 读取基础数据 C/F/G/D/A ===
    dataC = get_values(ws_report, "C")
    dataF = get_values(ws_report, "F")
    dataG = get_values(ws_report, "G")
    dataD = get_values(ws_report, "D")
    dataA = get_values(ws_report, "A")

    arr_backup = [["" for _ in range(5)] for _ in range(20)]
    for i in range(20):
        c, f, g, d, a = dataC[i], dataF[i], dataG[i], dataD[i], dataA[i]
        if c is None and f is None and g is None and d is None and a is None:
            continue
        arr_backup[i][0] = c
        arr_backup[i][2] = f
        arr_backup[i][3] = g
        arr_backup[i][1] = a if (d == 0 or d is None) else d
        if g not in (None, ""):
            arr_backup[i][4] = "CMM"

    # === 4. 写入 A8:E27 ===
    for ws in wb_backup.worksheets:
        for r in range(8, 28):
            for c in range(1, 6):
                ws.cell(r, c, None)
        for r in range(20):
            for c in range(5):
                ws.cell(r + 8, c + 1, arr_backup[r][c])

    log("✅ 写入 A8:E27 完成")

    # === 5. 收集报告文件中的 PCDmisExcel 工作表 ===
    pcd_sheets = [s for s in wb_report.sheetnames if s.startswith("PCDmisExcel")]
    pcd_sheets = pcd_sheets[:200]
    log(f"共找到 {len(pcd_sheets)} 个 PCDmisExcel 工作表")

    pcd_data = {}

    for idx, sheet_name in enumerate(pcd_sheets):
        ws = wb_report[sheet_name]

        def get_last_row(col):
            for row in range(20, 0, -1):
                if ws[f"{col}{row}"].value not in (None, ""):
                    return row
            return 0

        row_h = get_last_row("H")
        row_i = get_last_row("I")
        row_count = max(row_h, row_i, 0)
        if row_count == 0:
            continue

        # 判断使用 H 还是 I 列
        sumH = 0.0
        for r in range(1, 21):
            val = ws[f"H{r}"].value
            try:
                if val not in (None, ""):
                    sumH += float(val)
            except Exception:
                pass

        data_col = "I" if sumH == 0 else "H"
        data_vals = [ws[f"{data_col}{r}"].value for r in range(1, row_count + 1)]

        pcd_data[sheet_name] = data_vals
        log(f"读取 {sheet_name}: {len(data_vals)} 行, 使用列 {data_col}")

    # === 6. 写入备份文件 F8:Y27 ===
    for ws in wb_backup.worksheets:
        for r in range(8, 28):
            for c in range(6, 26):
                ws.cell(r, c, None)

    for i, (sheet_name, values) in enumerate(pcd_data.items()):
        backup_index = i // 20
        backup_col_offset = (i % 20) + 6
        if backup_index < len(wb_backup.worksheets):
            ws_target = wb_backup.worksheets[backup_index]
            for r, val in enumerate(values[:20]):
                ws_target.cell(r + 8, backup_col_offset, val)

    log("✅ 写入 F8:Y27 完成")

    # === 7. 删除空白工作表 ===
    sheets_to_delete = []
    for ws in wb_backup.worksheets:
        if ws["F8"].value in (None, "", "###EMPTY###"):
            sheets_to_delete.append(ws.title)

    if len(sheets_to_delete) < len(wb_backup.worksheets):
        for name in sheets_to_delete:
            del wb_backup[name]
        log(f"🗑️ 删除空白工作表: {sheets_to_delete}")
    else:
        log("⚠️ 所有工作表的F8都为空，至少保留一个工作表！")

    # === 8. 保存结果 ===
    wb_backup.save(backup_path)
    log(f"✅ 数据处理完成！共处理了 {len(pcd_sheets)} 个 PCDmisExcel 工作表。")
    log("=== 执行结束 ===\n")


if __name__ == "__main__":
    # 命令行支持：python process_excel.py [backup.xlsx] [1.xlsx]
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()

    origin_file = filedialog.askopenfilename(title="请选择源文件", filetypes=[("Excel 文件", "*.xlsx")])
    if not origin_file:
        messagebox.showwarning("提示", "未选择源文件，已取消。")
        sys.exit()

    template_file = filedialog.askopenfilename(title="请选择模板文件 template .xlsx", filetypes=[("Excel 文件", "*.xlsx")])
    if not template_file:
        messagebox.showwarning("提示", "未选择模板文件，已取消。")
        sys.exit()

    start(template_file, origin_file)
    messagebox.showinfo("完成", "Excel 数据处理完成！\n详细信息见 process.log。")
