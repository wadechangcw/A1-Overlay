# ============================================================
# PART 1 - Imports / Utility / VBA Macro
# ============================================================

import os
import re
import time
import threading
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

import win32com.client as win32
from win32com.client import constants

# ============================================================
# Utility
# ============================================================

def sanitize(name: str) -> str:
    """Make sheet names safe for Excel."""
    invalid = [":", "\\", "/", "?", "*", "[", "]"]
    for c in invalid:
        name = name.replace(c, "_")
    return name[:31]


def make_short_name(sheet_name):
    """Generate short sheet name for split."""
    parts = sheet_name.split()
    if len(parts) >= 2:
        short = parts[0] + "_" + "".join(w[0] for w in parts[1:])
    else:
        short = sheet_name[:10]
    return sanitize(short)


# ============================================================
# VBA Macro (Your Draw_MultiCharts_Final)
# ============================================================

VBA_MACRO = r'''
Sub Draw_MultiCharts_Final()

    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastCol As Long, lastRow As Long
    Dim c As Long
    Dim chartCount As Long
    Dim seriesCount As Long
    Dim chartObj As ChartObject
    Dim ch As Chart
    Dim S As Series
    Dim chartTop As Double, chartLeft As Double
    Dim chartHeight As Double

    lastCol = ws.Cells(6, ws.Columns.Count).End(xlToLeft).Column

    '======================================
    ' 刪除所有舊圖
    '======================================
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj

    ' 放置圖的起點位置
    chartTop = ws.Range("H2").Top
    chartLeft = ws.Range("H2").Left
    chartHeight = 500

    chartCount = 0
    seriesCount = 0

    '======================================
    ' 主迴圈：每 2 欄為一組 X,Y
    '======================================
    For c = 1 To lastCol Step 2

        If seriesCount = 0 Then
            chartCount = chartCount + 1

            ' 建立新圖表
            Set chartObj = ws.ChartObjects.Add( _
                chartLeft, _
                chartTop + (chartCount - 1) * (chartHeight + 50), _
                900, _
                chartHeight)

            Set ch = chartObj.Chart
            ch.ChartType = xlXYScatterSmoothNoMarkers

            ' 清理預設 series
            Do While ch.SeriesCollection.Count > 0
                ch.SeriesCollection(1).Delete
            Loop
        End If

        If c + 1 > lastCol Then Exit For

        lastRow = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If lastRow < 6 Then GoTo ContinueLoop

        Set S = ch.SeriesCollection.NewSeries
        S.Name = ws.Cells(1, c + 1).Value
        S.XValues = ws.Range(ws.Cells(6, c), ws.Cells(lastRow, c))
        S.Values = ws.Range(ws.Cells(6, c + 1), ws.Cells(lastRow, c + 1))

        seriesCount = seriesCount + 1
        If seriesCount >= 200 Then seriesCount = 0

ContinueLoop:
    Next c


    '======================================
    ' 套用格式
    '======================================
    For Each chartObj In ws.ChartObjects
        Set ch = chartObj.Chart

        With ch.Axes(xlCategory)
            .ScaleType = xlLogarithmic
            .MinimumScale = 100
            .MaximumScaleIsAuto = True
            .TickLabelPosition = xlTickLabelPositionLow
            .HasMajorGridlines = True
            .HasMinorGridlines = True
            .CrossesAt = 100
        End With

        With ch.Axes(xlValue)
            .HasMajorGridlines = True
            .HasMinorGridlines = True
        End With

        ch.HasLegend = True
        ch.Legend.Position = xlLegendPositionBottom
        ch.Legend.Height = 35
        ch.Legend.Left = 100
        ch.Legend.Top = ch.ChartArea.Height - 50

        With ch.PlotArea
            .Left = 40
            .Width = ch.ChartArea.Width - 60
            .Top = 40
            .Height = ch.ChartArea.Height - 120
        End With

        ch.HasTitle = True
        ch.ChartTitle.Text = "Chart " & chartObj.Index

    Next chartObj

End Sub
'''
# ============================================================
# PART 2 - Split Excel Sheets (Your Latest Split Logic)
# ============================================================

def split_excel_file(input_path, output_dir):
    """
    Split the given Excel file into many smaller sheets based on your rules:
    - Sheets named Summary are copied as-is.
    - If a column label matches xxx(number), split by that.
    - Otherwise fallback: split every two columns as a block.
    """

    xls = pd.ExcelFile(input_path)
    base_name = os.path.basename(input_path).replace(".xlsx", "").replace(".xlsm", "")
    out_path = os.path.join(output_dir, f"{base_name}_SPLIT.xlsx")

    writer = pd.ExcelWriter(out_path, engine="xlsxwriter")

    for sh in xls.sheet_names:

        df = pd.read_excel(input_path, sheet_name=sh, header=None)

        # ---------- Summary sheet: copy directly ----------
        if sh.lower().startswith("summary"):
            df.to_excel(writer, sheet_name=sanitize(sh), index=False, header=False)
            continue

        short = make_short_name(sh)
        row1 = df.iloc[1].tolist()
        total_cols = df.shape[1]

        # ----------------------------------------------------------
        # Step 1: Detect columns matching label format xxx(123)
        # ----------------------------------------------------------
        regex_cols = []
        for col in range(0, total_cols, 2):
            if col < len(row1):
                label = row1[col]
                if isinstance(label, str) and re.match(r".+\(\d+\)", label.strip()):
                    regex_cols.append(col)

        # ----------------------------------------------------------
        # Step 2: If regex columns exist → Use label split
        # ----------------------------------------------------------
        if len(regex_cols) > 0:
            for col in regex_cols:
                label = row1[col].strip()
                block = df.iloc[:, col:col+2]

                new_sheet = sanitize(f"{short}_{label}")
                base = new_sheet
                cnt = 1
                while new_sheet in writer.sheets:
                    new_sheet = sanitize(f"{base}_{cnt}")
                    cnt += 1

                block.to_excel(writer, sheet_name=new_sheet, index=False, header=False)

        else:
            # ------------------------------------------------------
            # Step 3: Fallback — every two columns
            # ------------------------------------------------------
            for col in range(0, total_cols, 2):

                # Skip if this block is entirely empty
                if df.iloc[:, col:col+2].dropna(how="all").empty:
                    continue

                label = row1[col] if col < len(row1) else f"Col{col}"
                label = str(label).strip()

                if label == "" or label.lower() == "nan":
                    label = f"Block_{col//2 + 1}"

                block = df.iloc[:, col:col+2]

                new_sheet = sanitize(f"{short}_{label}")
                base = new_sheet
                cnt = 1
                while new_sheet in writer.sheets:
                    new_sheet = sanitize(f"{base}_{cnt}")
                    cnt += 1

                block.to_excel(writer, sheet_name=new_sheet, index=False, header=False)

    writer.close()
    return out_path

# ============================================================
# PART 3 - Batch Merge Logic (Stable, Keep Sheet Order)
# ============================================================

def batch_merge_split_files(split_files, output_dir, batch_size=25,
                            progress_callback=None, status_callback=None):

    batch_results = []
    total_batches = (len(split_files) + batch_size - 1) // batch_size
    global_step = 0
    global_total_steps = len(split_files)  # For UI progress

    for b in range(total_batches):

        batch_files = split_files[b*batch_size:(b+1)*batch_size]

        if status_callback:
            status_callback(f"批次 {b+1}/{total_batches}：讀取 {len(batch_files)} 檔案中…")

        # Load this batch into memory (sheet_name=None → read all sheets)
        cache = {f: pd.read_excel(f, sheet_name=None, header=None) for f in batch_files}

        # Determine sheet order using the first file of the batch
        base_order = list(cache[batch_files[0]].keys())

        # Find common sheets across all files in this batch
        common = set(base_order)
        for f in batch_files:
            common &= set(cache[f].keys())

        # Output for this batch
        batch_output = os.path.join(output_dir, f"MERGE_BATCH_{b+1}.xlsx")
        writer = pd.ExcelWriter(batch_output, engine="xlsxwriter")

        for sh in base_order:
            if sh not in common:
                continue

            if status_callback:
                status_callback(f"批次 {b+1} → 合併 Sheet：{sh}")

            # Merge all sheets from cache
            merged_list = [cache[f][sh] for f in batch_files]
            merged = pd.concat(merged_list, axis=1, ignore_index=True)

            # Create a header row containing filenames (per two columns)
            header_row = []
            per_file_width = merged.shape[1] // len(batch_files)
            for f in batch_files:
                header_row += [os.path.basename(f).replace("_SPLIT.xlsx", "")] * per_file_width

            final_df = pd.DataFrame([header_row])
            final_df = pd.concat([final_df, merged], axis=0, ignore_index=True)

            # Write into sheet
            final_df.to_excel(writer, sheet_name=sanitize(sh), index=False, header=False)

        writer.close()
        batch_results.append(batch_output)

        # Progress update
        global_step += len(batch_files)
        if progress_callback:
            progress_callback(global_step, global_total_steps)

        del cache  # free memory

    # ---------------------------------------------------------
    # FINAL MERGE of all MERGE_BATCH_xxx → ALL_MERGED.xlsx
    # ---------------------------------------------------------
    if status_callback:
        status_callback("開始最終合併所有批次結果…")

    ok, result = merge_final_batches(
        batch_results,
        output_dir,
        progress_callback=progress_callback,
        status_callback=status_callback
    )

    return ok, result


# ============================================================
# Final Merge (MERGE_BATCH → ALL_MERGED)
# ============================================================

def merge_final_batches(batch_results, output_dir,
                        progress_callback=None, status_callback=None):

    first = pd.ExcelFile(batch_results[0])
    base_order = first.sheet_names  # final sheet order is determined here

    sets = [set(pd.ExcelFile(f).sheet_names) for f in batch_results]
    common = set.intersection(*sets)

    out_path = os.path.join(output_dir, "ALL_MERGED.xlsx")
    writer = pd.ExcelWriter(out_path, engine="xlsxwriter")

    total_steps = len(common)
    cur = 0

    for sh in base_order:
        if sh not in common:
            continue

        if status_callback:
            status_callback(f"最終合併 → {sh}")

        merged = None

        for f in batch_results:
            df = pd.read_excel(f, sheet_name=sh, header=None)

            if merged is None:
                merged = df
            else:
                merged = pd.concat([merged, df], axis=1, ignore_index=True)

        merged.to_excel(writer, sheet_name=sanitize(sh), index=False, header=False)

        cur += 1
        if progress_callback:
            progress_callback(cur, total_steps)

    writer.close()
    return True, out_path

# ============================================================
# PART 4 - Excel COM: Inject VBA, Run Macro on Each Sheet (Hidden Mode)
# ============================================================

def run_vba_on_merged_excel(excel_path, vba_code,
                            progress_callback=None,
                            status_callback=None):
    """
    Open Excel in hidden mode, inject VBA macro, run it on each non-Summary sheet,
    remove the module, save back as pure .xlsx, and close Excel safely.
    """

    if status_callback:
        status_callback("啟動 Excel（隱藏模式）...")

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 1  # msoAutomationSecurityForceDisable

    wb = excel.Workbooks.Open(excel_path)

    # -----------------------------------
    # Insert VBA module
    # -----------------------------------
    if status_callback:
        status_callback("匯入 VBA 程式碼（暫時）...")

    vbcomp = wb.VBProject.VBComponents.Add(1)  # 1 = vbext_ct_StdModule
    vbcomp.CodeModule.AddFromString(vba_code)

    # Count sheets that will run the macro
    sheets = [ws for ws in wb.Worksheets if not ws.Name.lower().startswith("summary")]
    total = len(sheets)
    cur = 0

    # -----------------------------------
    # Run macro on each sheet individually
    # -----------------------------------
    for ws in sheets:

        cur += 1
        if status_callback:
            status_callback(f"執行 VBA → {ws.Name}  ({cur}/{total})")

        if progress_callback:
            progress_callback(cur, total)

        ws.Activate()
        excel.Run("Draw_MultiCharts_Final")

    # -----------------------------------
    # Remove VBA module (to keep file clean)
    # -----------------------------------
    if status_callback:
        status_callback("清除暫存 VBA 模組...")

    wb.VBProject.VBComponents.Remove(vbcomp)

    # -----------------------------------
    # Save as pure .xlsx
    # -----------------------------------
    if status_callback:
        status_callback("儲存結果檔案（純 .xlsx）...")

    wb.Save()  # Save back to ALL_MERGED.xlsx (no macro)

    wb.Close(SaveChanges=True)
    excel.Quit()

    # Make sure Excel COM is fully released
    del wb
    del excel

    if status_callback:
        status_callback("所有分頁圖表已完成！Excel 已關閉。")

    return True
# ============================================================
# PART 5 - GUI (Folder Selection, File List, Progress, ETA)
# ============================================================

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel TX/MI 拆分＋批次合併＋自動畫圖工具（穩定版）")

        self.file_list = []          # all excel files in selected folder
        self.selected_files = []     # user selected files
        self.output_dir = ""         # default = same as folder
        self.drag_start_index = None # for sliding multi-select

        # =====================================================
        # UI Layout
        # =====================================================
        frm = tk.Frame(root)
        frm.pack(padx=10, pady=10)

        # -------- Folder Select --------
        tk.Button(frm, text="選擇來源資料夾", command=self.pick_folder)\
            .grid(row=0, column=0, sticky="w")

        self.folder_label = tk.Label(frm, text="未選擇來源資料夾")
        self.folder_label.grid(row=0, column=1, padx=10, sticky="w")

        # -------- Output Folder --------
        tk.Button(frm, text="變更輸出資料夾", command=self.change_output)\
            .grid(row=1, column=0, sticky="w")

        self.output_label = tk.Label(frm, text="輸出資料夾：未設定")
        self.output_label.grid(row=1, column=1, padx=10, sticky="w")

        # -------- File Listbox --------
        tk.Label(frm, text="請選擇要處理的 Excel：").grid(row=2, column=0, sticky="w")

        self.listbox = tk.Listbox(frm, width=80, height=12,
                                  selectmode=tk.MULTIPLE)
        self.listbox.grid(row=3, column=0, columnspan=3, pady=5)

        # -- Standard selection update --
        self.listbox.bind("<<ListboxSelect>>", self.update_selected_count)

        # -- Enable sliding multi-select (拖曳反白選取功能) --
        self.listbox.bind("<Button-1>", self.drag_start)
        self.listbox.bind("<B1-Motion>", self.drag_motion)

        # -------- Select All / None --------
        tk.Button(frm, text="全選", command=self.select_all)\
            .grid(row=4, column=0, sticky="w")

        tk.Button(frm, text="全不選", command=self.select_none)\
            .grid(row=4, column=1, sticky="w")

        self.count_label = tk.Label(frm, text="已選擇：0 個檔案")
        self.count_label.grid(row=4, column=2, sticky="e")

        # -------- Progress Bar --------
        self.progress = ttk.Progressbar(frm, length=450, mode="determinate")
        self.progress.grid(row=5, column=0, columnspan=3, pady=5)

        # ETA
        self.eta_label = tk.Label(frm, text="", fg="green")
        self.eta_label.grid(row=6, column=0, columnspan=3)

        # Status messages
        self.status = tk.Label(frm, text="", fg="blue")
        self.status.grid(row=7, column=0, columnspan=3)

        # -------- Action Buttons --------
        tk.Button(frm, text="開始拆分", command=self.start_split)\
            .grid(row=8, column=0, pady=10)

        tk.Button(frm, text="合併 SPLIT → ALL_MERGED.xlsx",
                  command=self.start_merge)\
            .grid(row=8, column=1, pady=10)

        tk.Button(frm, text="執行 VBA：產生所有圖表",
                  command=self.start_vba)\
            .grid(row=8, column=2, pady=10)


    # ============================================================
    # Folder Selection
    # ============================================================
    def pick_folder(self):
        folder = filedialog.askdirectory(title="選擇來源資料夾")
        if not folder:
            return

        self.folder_label.config(text=folder)
        self.output_dir = folder
        self.output_label.config(text=f"輸出資料夾：{folder}")

        self.file_list = []
        self.selected_files = []
        self.listbox.delete(0, tk.END)

        # Load all Excel files
        for fname in os.listdir(folder):
            if fname.lower().endswith((".xlsx", ".xlsm")):
                fullpath = os.path.join(folder, fname)
                self.file_list.append(fullpath)
                self.listbox.insert(tk.END, fullpath)

        self.update_selected_count()


    # ============================================================
    # Output Folder
    # ============================================================
    def change_output(self):
        folder = filedialog.askdirectory(title="選擇輸出資料夾")
        if folder:
            self.output_dir = folder
            self.output_label.config(text=f"輸出資料夾：{folder}")


    # ============================================================
    # Selection tools
    # ============================================================
    def select_all(self):
        self.listbox.select_set(0, tk.END)
        self.update_selected_count()

    def select_none(self):
        self.listbox.select_clear(0, tk.END)
        self.update_selected_count()

    # Update selected count
    def update_selected_count(self, event=None):
        idxs = self.listbox.curselection()
        self.selected_files = [self.listbox.get(i) for i in idxs]
        self.count_label.config(text=f"已選擇：{len(self.selected_files)} 個檔案")


    # ============================================================
    # Drag-Select (拖曳連續選取功能)
    # ============================================================
    def drag_start(self, event):
        widget = event.widget
        self.drag_start_index = widget.nearest(event.y)

        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(self.drag_start_index)
        self.update_selected_count()

    def drag_motion(self, event):
        widget = event.widget
        current_index = widget.nearest(event.y)

        start = min(self.drag_start_index, current_index)
        end = max(self.drag_start_index, current_index)

        self.listbox.selection_clear(0, tk.END)
        self.listbox.selection_set(start, end)
        self.update_selected_count()


    # ============================================================
    # Split
    # ============================================================
    def start_split(self):
        if not self.selected_files:
            messagebox.showwarning("提醒", "請先選擇要拆分的 Excel 檔案！")
            return

        self.status.config(text="開始拆分...")
        self.progress["value"] = 0
        self.progress["maximum"] = len(self.selected_files)

        threading.Thread(target=self.process_split_thread, daemon=True).start()

    def process_split_thread(self):
        start_time = time.time()

        for idx, f in enumerate(self.selected_files, start=1):
            try:
                split_excel_file(f, self.output_dir)
                self.status.config(text=f"完成 {idx}/{len(self.selected_files)} → {os.path.basename(f)}")
            except Exception as e:
                self.status.config(text=f"錯誤：{e}")

            self.progress["value"] = idx

            # ETA
            elapsed = time.time() - start_time
            avg = elapsed / idx
            remain = avg * (len(self.selected_files) - idx)
            self.eta_label.config(text=f"預估剩餘時間：約 {remain:.1f} 秒")

        self.status.config(text="拆分完成！")
        messagebox.showinfo("完成", "全部拆分完成！")


    # ============================================================
    # MERGE
    # ============================================================
    def start_merge(self):
        split_files = [
            os.path.join(self.output_dir, f)
            for f in os.listdir(self.output_dir)
            if f.endswith("_SPLIT.xlsx")
        ]

        if not split_files:
            messagebox.showwarning("提醒", "找不到任何 _SPLIT.xlsx！")
            return

        self.status.config(text="開始合併...")
        self.progress["value"] = 0

        threading.Thread(target=self.process_merge_thread,
                         args=(split_files,), daemon=True).start()

    def process_merge_thread(self, split_files):

        def update_progress(cur, total):
            self.progress["maximum"] = total
            self.progress["value"] = cur

        def update_status(msg):
            self.status.config(text=msg)

        ok, result = batch_merge_split_files(
            split_files,
            self.output_dir,
            batch_size=25,
            progress_callback=update_progress,
            status_callback=update_status
        )

        if ok:
            messagebox.showinfo("完成", f"合併完成！輸出檔案：{result}")
            self.status.config(text=f"合併完成 → {result}")
        else:
            messagebox.showerror("錯誤", result)


    # ============================================================
    # VBA Execute
    # ============================================================
    def start_vba(self):
        excel_path = os.path.join(self.output_dir, "ALL_MERGED.xlsx")
        if not os.path.exists(excel_path):
            messagebox.showwarning("提醒", "ALL_MERGED.xlsx 不存在，請先執行合併！")
            return

        self.status.config(text="啟動 Excel，準備執行 VBA...")
        self.progress["value"] = 0

        threading.Thread(target=self.process_vba_thread,
                         args=(excel_path,), daemon=True).start()

    def process_vba_thread(self, excel_path):

        def update_progress(cur, total):
            self.progress["maximum"] = total
            self.progress["value"] = cur

        def update_status(msg):
            self.status.config(text=msg)

        run_vba_on_merged_excel(
            excel_path,
            VBA_MACRO,
            progress_callback=update_progress,
            status_callback=update_status
        )

        messagebox.showinfo("完成", "所有圖表已成功產生！")
        self.status.config(text="🎉 圖表製作完成")

# ============================================================
# PART 6 - Main Entry Point
# ============================================================

def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()

